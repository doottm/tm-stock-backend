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
    return { productsSheet, movementsSheet };
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

app.get('/', (req, res) => {
    res.send('TMStock API Server is running. 🚀');
});

app.listen(PORT, () => {
    console.log(`🚀 Server is running on http://localhost:${PORT}`);
});
