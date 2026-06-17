// [수정사항: 2026-05-19] 기록을 위해 상단에 수정 이력을 남깁니다.
// 1. POST /api/products 추가 (품목 추가 기능)
// 2. 입고 기능 관련 로직 정리 (실사 업데이트 위주로 변경)
// 3. 재고 업데이트 시 Last_Updated(최종 수정일) 기록 기능 추가

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const https = require('https');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

const serviceAccountAuth = new JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    // [수정사항 2026-05-19] Render 환경변수 입력 시 포함된 양끝 쌍따옴표 제거
    key: process.env.GOOGLE_PRIVATE_KEY ? process.env.GOOGLE_PRIVATE_KEY.replace(/^"|"$/g, '').replace(/\\n/g, '\n') : '',
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

let doc;
if (process.env.GOOGLE_SHEET_ID) {
    doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID, serviceAccountAuth);
}

// 텔레그램 푸시 알림 함수
function sendTelegramMessage(message) {
    const token = process.env.TELEGRAM_BOT_TOKEN;
    const chatId = process.env.TELEGRAM_CHAT_ID;
    
    if (!token || !chatId) return;

    const data = JSON.stringify({
        chat_id: chatId,
        text: message,
        parse_mode: 'HTML'
    });

    const options = {
        hostname: 'api.telegram.org',
        path: `/bot${token}/sendMessage`,
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Content-Length': Buffer.byteLength(data)
        }
    };

    const req = https.request(options, (res) => {});
    req.on('error', (e) => {
        console.error('Telegram notification error:', e);
    });
    req.write(data);
    req.end();
}

// [수정사항 2026-05-22] 텔레그램 연동 디버깅용 API
app.get('/api/test-telegram', (req, res) => {
    const token = process.env.TELEGRAM_BOT_TOKEN;
    const chatId = process.env.TELEGRAM_CHAT_ID;
    
    if (!token || !chatId) {
        return res.json({ success: false, error: "Render에 TELEGRAM_BOT_TOKEN 또는 TELEGRAM_CHAT_ID 환경변수가 설정되지 않았습니다." });
    }

    const data = JSON.stringify({ chat_id: chatId, text: "🔔 <b>테스트 메시지</b>\n정상적으로 연결되었습니다!", parse_mode: 'HTML' });
    const options = {
        hostname: 'api.telegram.org', path: `/bot${token}/sendMessage`, method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) }
    };

    const request = https.request(options, (response) => {
        let responseData = '';
        response.on('data', chunk => responseData += chunk);
        response.on('end', () => {
            res.json({ success: response.statusCode === 200, status: response.statusCode, response: JSON.parse(responseData) });
        });
    });
    request.on('error', (e) => res.json({ success: false, error: e.message }));
    request.write(data);
    request.end();
});

async function loadSheets() {
    if (!doc) throw new Error("구글 시트 ID가 설정되지 않았습니다.");
    await doc.loadInfo(); 
    const productsSheet = doc.sheetsByTitle['품목 및 재고'];
    const movementsSheet = doc.sheetsByTitle['재고 변동 히스토리'];
    
    if (!productsSheet || !movementsSheet) {
        throw new Error("필요한 시트('품목 및 재고', '재고 변동 히스토리')를 찾을 수 없습니다.");
    }

    // [수정사항 2026-06-08] 해주세요(TODO) 탭 로드 및 자동 생성
    // [수정사항 2026-06-17] Confirm, Progress 컬럼 추가 및 기존 시트 헤더 마이그레이션
    let todosSheet = doc.sheetsByTitle['해주세요'];
    if (!todosSheet) {
        todosSheet = await doc.addSheet({
            title: '해주세요',
            headerValues: ['ID', 'Text', 'Priority', 'Creator', 'CreatedAt', 'Completed', 'Status', 'DeletedAt', 'Confirm', 'Progress']
        });
    } else {
        await todosSheet.loadHeaderRow();
        const headers = todosSheet.headerValues || [];
        if (!headers.includes('Confirm') || !headers.includes('Progress')) {
            const newHeaders = [...headers];
            if (!newHeaders.includes('Confirm')) newHeaders.push('Confirm');
            if (!newHeaders.includes('Progress')) newHeaders.push('Progress');
            await todosSheet.setHeaderRow(newHeaders);
        }
    }
    
    return { productsSheet, movementsSheet, todosSheet };
}

// [수정사항 2026-05-19] 품목에 Last_Updated 가져오기 추가
app.get('/api/products', async (req, res) => {
    try {
        const { productsSheet } = await loadSheets();
        const rows = await productsSheet.getRows();
        
        // [수정사항 2026-05-19] row_index 추가하여 완벽한 고유 식별자로 활용
        const products = rows.map((row, index) => ({
            row_index: index,
            sku: row.get('SKU') || '',
            item_name: row.get('Item_Name') || '',
            category: row.get('Category') || '',
            unit: row.get('Unit') || '',
            min_stock_level: parseFloat(row.get('Min_Stock_Level')) || 0,
            current_stock: parseFloat(row.get('Current_Stock')) || 0,
            supplier: row.get('Supplier') || '',
            location: row.get('Location') || '',
            last_updated: row.get('Last_Updated') || '',
            last_worker: row.get('마지막수정자') || row.get('수정자') || row.get('Worker') || ''
        }));

        res.json(products);
    } catch (error) {
        console.error('Error fetching products:', error);
        res.status(500).json({ error: '데이터를 불러오는데 실패했습니다.', details: error.message });
    }
});

// [수정사항 2026-05-19] 1. 품목 추가 API 생성 (SKU 백엔드 생성 제외 처리)
app.post('/api/products', async (req, res) => {
    try {
        const { item_name, category, unit, min_stock_level, location, supplier } = req.body;
        
        if (!item_name) {
            return res.status(400).json({ error: '품목명은 필수입니다.' });
        }

        const { productsSheet } = await loadSheets();
        
        await productsSheet.addRow({
            SKU: '', // 구글 시트에서 수식이나 앱스로 연동되도록 비워둠
            Item_Name: item_name,
            Category: category || '',
            Unit: unit || '',
            Min_Stock_Level: min_stock_level || 0,
            Location: location || '',
            Supplier: supplier || '',
            Status: 'TRUE',
            Current_Stock: 0,
            Last_Updated: new Date().toISOString()
        });

        res.json({ success: true, message: '품목이 추가되었습니다.' });
    } catch (error) {
        console.error('Error posting product:', error);
        res.status(500).json({ error: '품목 추가에 실패했습니다.', details: error.message });
    }
});

// [수정사항 2026-05-19] 2, 3. 실사 재고 업데이트 및 업데이트 시간 기록
app.post('/api/movements', async (req, res) => {
    try {
        const { date, sku, item_name, type, quantity, worker, location, row_index } = req.body;
        
        // [수정사항 2026-05-19] SKU가 없을 수도 있으므로 필수 체크에서 제외 (품목명으로 대체)
        if (!date || !type || quantity === undefined || !item_name) {
            return res.status(400).json({ error: '필수 데이터(날짜, 품목명, 수량)가 누락되었습니다.' });
        }

        const { productsSheet, movementsSheet } = await loadSheets();
        
        await movementsSheet.addRow({
            date: date,
            sku: sku,
            item_name: item_name,
            type: type,
            quantity: quantity,
            worker: worker,
            created_at: new Date().toISOString()
        });

        const rows = await productsSheet.getRows();
        
        // [수정사항 2026-05-19] row_index가 있으면 무조건 해당 행 사용, 없으면 기존 폴백 로직
        let productRow;
        if (row_index !== undefined && row_index !== null && rows[row_index]) {
            productRow = rows[row_index];
        } else {
            productRow = rows.find(r => {
                if (sku) return r.get('SKU') === sku;
                if (location) return r.get('Item_Name') === item_name && r.get('Location') === location;
                return r.get('Item_Name') === item_name;
            });
        }

        if (productRow) {
            // [수정사항] 입고 메뉴가 삭제되었으므로, 넘어오는 type은 '실사/조정' 기반의 덮어쓰기입니다.
            productRow.set('Current_Stock', parseFloat(quantity)); 
            
            // [수정사항] 업데이트된 날짜/시간 기록 및 마지막 수정자 기록
            productRow.set('Last_Updated', date); 
            try { productRow.set('마지막수정자', worker); } catch(e) {}
            
            await productRow.save(); 
            
            // [수정사항 2026-05-22] 텔레그램 푸시 알림 발송
            const locText = location ? `[${location}] ` : '';
            const msg = `🔔 <b>재고 업데이트</b>\n\n👤 <b>${worker || '담당자'}</b>님이 <b>${locText}${item_name}</b>의 재고를 <b>${quantity}</b>개로 변경했습니다.\n🕒 ${date}`;
            sendTelegramMessage(msg);
        }

        res.json({ success: true, message: '성공적으로 일괄 기록되었습니다.' });
    } catch (error) {
        console.error('Error posting movement:', error);
        res.status(500).json({ error: '데이터 저장에 실패했습니다.', details: error.message });
    }
});

// [수정사항 2026-06-15] 최근 재고 변동 내역 조회 API 추가 (오늘, 3일, 1주일 필터)
app.get('/api/movements', async (req, res) => {
    try {
        const days = req.query.days !== undefined ? parseInt(req.query.days) : 7;
        const { movementsSheet, productsSheet } = await loadSheets();
        
        // 1. 제품명 -> 거래처(Supplier) 매핑 생성
        const productRows = await productsSheet.getRows();
        const supplierMap = {};
        productRows.forEach(r => {
            const name = r.get('Item_Name');
            const supplier = r.get('Supplier');
            if (name) {
                supplierMap[name] = supplier || '-';
            }
        });
        
        // 2. 이동 기록 가져오기
        const movementRows = await movementsSheet.getRows();
        
        // 3. KST (UTC+9) 기준 날짜 필터링 기준일 계산
        const nowKst = new Date(new Date().getTime() + 9 * 60 * 60 * 1000);
        const cutoffKst = new Date(nowKst);
        if (days === 0) {
            cutoffKst.setUTCHours(0, 0, 0, 0);
        } else {
            cutoffKst.setUTCDate(cutoffKst.getUTCDate() - days);
            cutoffKst.setUTCHours(0, 0, 0, 0);
        }
        const cutoff = new Date(cutoffKst.getTime() - 9 * 60 * 60 * 1000);
        
        const movements = [];
        // 최신 업데이트가 아래에 쌓이므로 역순(최신순)으로 탐색
        for (let i = movementRows.length - 1; i >= 0; i--) {
            const row = movementRows[i];
            const dateStr = row.get('date') || '';
            const createdAtStr = row.get('created_at') || dateStr;
            
            let itemDate = null;
            if (createdAtStr) {
                if (createdAtStr.includes('T') && createdAtStr.endsWith('Z')) {
                    itemDate = new Date(createdAtStr);
                } else {
                    const cleanDateStr = createdAtStr.replace(' ', 'T');
                    itemDate = new Date(cleanDateStr.includes('+') ? cleanDateStr : cleanDateStr + '+09:00');
                }
            }
            
            if (!itemDate || isNaN(itemDate.getTime())) continue;
            
            // 기준일 이전 데이터면 건너뜀
            if (itemDate < cutoff) continue;
            
            const itemName = row.get('item_name') || '';
            movements.push({
                date: dateStr,
                sku: row.get('sku') || '',
                item_name: itemName,
                type: row.get('type') || '',
                quantity: parseFloat(row.get('quantity')) || 0,
                worker: row.get('worker') || '',
                supplier: supplierMap[itemName] || '-'
            });
        }
        
        res.json(movements);
    } catch (error) {
        console.error('Error fetching movements:', error);
        res.status(500).json({ error: '데이터를 불러오는데 실패했습니다.', details: error.message });
    }
});

// [수정사항 2026-05-19] 5. 사용자 로그인 API 추가
app.post('/api/login', async (req, res) => {
    try {
        const { id, password } = req.body;
        if (!id || !password) return res.status(400).json({ error: '아이디와 비밀번호를 입력해주세요.' });

        if (!doc) throw new Error("구글 시트 ID가 설정되지 않았습니다.");
        await doc.loadInfo();
        
        const usersSheet = doc.sheetsByTitle['사용자'];
        if (!usersSheet) {
            return res.status(500).json({ error: '구글 시트에 "사용자" 탭이 존재하지 않습니다. 먼저 탭을 생성해주세요.' });
        }

        const rows = await usersSheet.getRows();
        const userRow = rows.find(r => r.get('ID') === id && r.get('Password') === password);

        if (userRow) {
            res.json({ success: true, name: userRow.get('Name') });
        } else {
            res.status(401).json({ success: false, error: '아이디 또는 비밀번호가 일치하지 않습니다.' });
        }
    } catch (error) {
        console.error('Login error:', error);
        res.status(500).json({ error: '로그인 처리 중 서버 오류가 발생했습니다.' });
    }
});

// [수정사항 2026-06-08] 6. "해주세요" (TODO) 공용 시트 연동 API 추가
// [수정사항 2026-06-17] 30일 경과된 휴지통 항목 물리 삭제 및 Confirm/Progress 조회 추가
app.get('/api/todos', async (req, res) => {
    try {
        const { todosSheet } = await loadSheets();
        const rows = await todosSheet.getRows();
        
        const now = new Date();
        const todos = [];
        const rowsToDelete = [];
        
        for (const row of rows) {
            const status = row.get('Status') || 'Active';
            const deletedAtStr = row.get('DeletedAt');
            
            // 30일 경과된 휴지통 항목 수집
            if (status === 'Trash' && deletedAtStr) {
                const deletedAt = new Date(deletedAtStr);
                const diffTime = Math.abs(now - deletedAt);
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                if (diffDays > 30) {
                    rowsToDelete.push(row);
                    continue;
                }
            }
            
            todos.push({
                id: row.get('ID'),
                text: row.get('Text'),
                priority: row.get('Priority'),
                creator: row.get('Creator'),
                createdAt: parseInt(row.get('CreatedAt')),
                completed: row.get('Completed') === 'TRUE',
                status: status,
                deletedAt: deletedAtStr || '',
                confirm: row.get('Confirm') || '확인',
                progress: row.get('Progress') || '진행중'
            });
        }

        // 30일 경과된 휴지통 항목 영구 삭제 (역순으로 안전하게 진행)
        for (let i = rowsToDelete.length - 1; i >= 0; i--) {
            try {
                await rowsToDelete[i].delete();
            } catch (err) {
                console.error('Failed to permanently delete old trash row:', err);
            }
        }

        res.json(todos);
    } catch (error) {
        console.error('Error fetching todos:', error);
        res.status(500).json({ error: '데이터를 불러오는데 실패했습니다.' });
    }
});

// [수정사항 2026-06-17] Confirm, Progress 컬럼 기본값 ('확인', '진행중') 세팅 추가
app.post('/api/todos', async (req, res) => {
    try {
        const { id, text, priority, creator, createdAt, date } = req.body;
        
        const { todosSheet } = await loadSheets();
        await todosSheet.addRow({
            ID: id,
            Text: text,
            Priority: priority,
            Creator: creator,
            CreatedAt: createdAt,
            Completed: 'FALSE',
            Status: 'Active',
            DeletedAt: '',
            Confirm: '확인',
            Progress: '진행중'
        });
        
        // 텔레그램 알림
        let priorityText = '';
        if (priority === 'high') priorityText = '[🚨긴급] ';
        else if (priority === 'normal') priorityText = '[📦주문] ';
        else priorityText = '[💬기타] ';

        const msg = `🔔 <b>새로운 "해주세요" 등록</b>\n\n👤 <b>${creator || '담당자'}</b>님이 새로운 요청을 남겼습니다.\n\n${priorityText}<b>${text}</b>\n\n🕒 ${date}`;
        sendTelegramMessage(msg);
        
        res.json({ success: true, message: '요청사항이 추가되었습니다.' });
    } catch (error) {
        console.error('Error posting todo:', error);
        res.status(500).json({ error: '요청사항 추가에 실패했습니다.' });
    }
});

// [수정사항 2026-06-17] confirm, progress 개별 업데이트 및 상호 동기화 추가
app.put('/api/todos/:id', async (req, res) => {
    try {
        const todoId = req.params.id;
        const { completed, status, confirm, progress } = req.body;
        
        const { todosSheet } = await loadSheets();
        const rows = await todosSheet.getRows();
        
        const row = rows.find(r => r.get('ID') === String(todoId));
        if (row) {
            if (completed !== undefined) {
                row.set('Completed', completed ? 'TRUE' : 'FALSE');
                // Completed 여부에 맞춰 Progress 상태 자동 동기화
                row.set('Progress', completed ? '완료' : '진행중');
            }
            if (status !== undefined) {
                row.set('Status', status);
                if (status === 'Trash') {
                    row.set('DeletedAt', new Date().toISOString());
                } else if (status === 'Active') {
                    row.set('DeletedAt', '');
                }
            }
            if (confirm !== undefined) {
                row.set('Confirm', confirm);
            }
            if (progress !== undefined) {
                row.set('Progress', progress);
                // Progress 상태에 맞춰 Completed 여부 자동 동기화
                if (progress === '완료') {
                    row.set('Completed', 'TRUE');
                } else if (progress === '진행중') {
                    row.set('Completed', 'FALSE');
                }
            }
            await row.save();
            res.json({ success: true, message: '상태가 업데이트되었습니다.' });
        } else {
            res.status(404).json({ error: '항목을 찾을 수 없습니다.' });
        }
    } catch (error) {
        console.error('Error updating todo:', error);
        res.status(500).json({ error: '상태 업데이트에 실패했습니다.' });
    }
});

app.get('/', (req, res) => {
    res.send('TMStock API Server is running. 🚀');
});

app.listen(PORT, () => {
    console.log(`🚀 Server is running on http://localhost:${PORT}`);
});
