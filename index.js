// [수정사항: 2026-05-19] 기록을 위해 상단에 수정 이력을 남깁니다.
// 1. POST /api/products 추가 (품목 추가 기능)
// 2. 입고 기능 관련 로직 정리 (실사 업데이트 위주로 변경)
// 3. 재고 업데이트 시 Last_Updated(최종 수정일) 기록 기능 추가

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const https = require('https');
const http = require('http');
const ws = require('ws');
const { google } = require('googleapis');
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

// ==========================================================================
// Real-Time Face-to-Face Translation Integrations (Ported from AG_Translation)
// ==========================================================================

// Google Sheets Logging Route for Translation Logs
// [수정사항 2026-06-20] Sheet1 하드코딩 → '통역기록' 시트 자동 생성 방식으로 변경
app.post('/api/log', async (req, res) => {
  const { staffScript, customerScript, targetLanguage } = req.body;

  // Gracefully handle missing credentials without crashing
  if (!doc) {
    console.warn('[Google Sheets] Missing GOOGLE_SHEET_ID. Log saved only in server console.');
    return res.status(200).json({
      success: false,
      warning: 'Google Sheets not configured. Saved to server logs instead.',
      data: { staffScript, customerScript, targetLanguage }
    });
  }

  try {
    await doc.loadInfo();

    // '통역기록' 시트가 없으면 자동 생성
    let logSheet = doc.sheetsByTitle['통역기록'];
    if (!logSheet) {
      logSheet = await doc.addSheet({
        title: '통역기록',
        headerValues: ['Timestamp', 'Staff_Script', 'Customer_Script', 'Language']
      });
      console.log('[Google Sheets] Created new sheet: 통역기록');
    }

    // 로그 행 추가
    const timestamp = new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });
    await logSheet.addRow({
      Timestamp: timestamp,
      Staff_Script: staffScript || '',
      Customer_Script: customerScript || '',
      Language: targetLanguage || ''
    });

    console.log('[Google Sheets] Translation log appended to 통역기록 sheet.');
    return res.json({ success: true, message: '상담 기록이 구글 시트 [통역기록] 탭에 저장되었습니다.' });
  } catch (error) {
    console.error('[Google Sheets] Error logging to sheets:', error);
    return res.status(500).json({ 
      success: false, 
      error: 'Failed to save log to Google Sheets.',
      details: error.message 
    });
  }
});

// REST Text Translation Endpoint for Quick Phrases (canned phrases)
app.post('/api/translate', async (req, res) => {
  const { text, targetLanguage } = req.body;
  const geminiKey = process.env.GEMINI_API_KEY;
  if (!geminiKey) {
    console.error('[Translation API] GEMINI_API_KEY is not defined');
    return res.status(500).json({ error: 'GEMINI_API_KEY is not configured on server' });
  }

  try {
    const model = 'gemini-2.5-flash';
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${geminiKey}`;
    
    console.log(`[Translation API] Request: Translate "${text}" to "${targetLanguage}"`);
    
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        contents: [{
          parts: [{
            text: `Translate the following Korean text into language code "${targetLanguage}". Return ONLY the translation, without any introduction, quotes, explanations, or extra text.\n\nText: ${text}`
          }]
        }]
      })
    });

    const data = await response.json();
    
    if (!response.ok) {
      console.error('[Translation API] Gemini API responded with error:', response.status, JSON.stringify(data));
      return res.status(response.status).json({ 
        error: 'Gemini API call failed', 
        status: response.status,
        details: data 
      });
    }

    if (!data.candidates || data.candidates.length === 0) {
      console.warn('[Translation API] No translation candidates returned. Full Response:', JSON.stringify(data));
      return res.json({ translatedText: text, warning: 'No candidates returned from Gemini' });
    }

    let translatedText = data.candidates[0].content?.parts?.[0]?.text?.trim();
    if (!translatedText) {
      console.warn('[Translation API] Candidate text is empty. Full Response:', JSON.stringify(data));
      translatedText = text;
    }
    
    // Clean up any double quotes Gemini might wrap the text with
    if (translatedText.startsWith('"') && translatedText.endsWith('"')) {
      translatedText = translatedText.slice(1, -1);
    }
    
    console.log(`[Translation API] Successfully translated "${text}" to "${translatedText}" (${targetLanguage})`);
    res.json({ translatedText });
  } catch (err) {
    console.error('[Translation API] Exception occurred:', err);
    res.status(500).json({ error: 'Translation failed due to exception', details: err.message });
  }
});

// Google Sheets API: Get Shared Quick Phrases
app.get('/api/phrases', async (req, res) => {
  if (!doc) {
    // Fallback if sheets not configured
    return res.json([
      { id: '1', text: '안녕하세요. 무엇을 도와드릴까요?' },
      { id: '2', text: '이쪽에 서명해 주세요.' },
      { id: '3', text: '여권을 보여주시겠어요?' },
      { id: '4', text: '결제가 완료되었습니다. 감사합니다.' },
      { id: '5', text: '잠시만 기다려 주세요.' }
    ]);
  }

  try {
    await doc.loadInfo();
    let phraseSheet = doc.sheetsByTitle['자주사용하는문구'];
    if (!phraseSheet) {
      phraseSheet = await doc.addSheet({
        title: '자주사용하는문구',
        headerValues: ['ID', 'Text', 'CreatedAt']
      });
      // Add default phrases
      await phraseSheet.addRows([
        { ID: '1', Text: '안녕하세요. 무엇을 도와드릴까요?', CreatedAt: new Date().toISOString() },
        { ID: '2', Text: '이쪽에 서명해 주세요.', CreatedAt: new Date().toISOString() },
        { ID: '3', Text: '여권을 보여주시겠어요?', CreatedAt: new Date().toISOString() },
        { ID: '4', Text: '결제가 완료되었습니다. 감사합니다.', CreatedAt: new Date().toISOString() },
        { ID: '5', Text: '잠시만 기다려 주세요.', CreatedAt: new Date().toISOString() }
      ]);
      console.log('[Google Sheets] Created new sheet: 자주사용하는문구 and added default phrases.');
    }

    const rows = await phraseSheet.getRows();
    const phrases = rows.map(row => ({
      id: row.get('ID'),
      text: row.get('Text'),
      createdAt: row.get('CreatedAt')
    }));

    return res.json(phrases);
  } catch (error) {
    console.error('[Google Sheets] Error getting phrases:', error);
    return res.status(500).json({ error: 'Failed to retrieve phrases from Google Sheets.' });
  }
});

// Google Sheets API: Add Shared Quick Phrase
app.post('/api/phrases', async (req, res) => {
  const { text } = req.body;
  if (!text || !text.trim()) {
    return res.status(400).json({ error: 'Text is required.' });
  }

  if (!doc) {
    return res.status(503).json({ error: 'Google Sheets not configured.' });
  }

  try {
    await doc.loadInfo();
    let phraseSheet = doc.sheetsByTitle['자주사용하는문구'];
    if (!phraseSheet) {
      phraseSheet = await doc.addSheet({
        title: '자주사용하는문구',
        headerValues: ['ID', 'Text', 'CreatedAt']
      });
    }

    const rows = await phraseSheet.getRows();
    const exists = rows.some(row => row.get('Text').trim() === text.trim());
    if (exists) {
      return res.status(409).json({ error: 'Phrase already exists.' });
    }

    const id = Date.now().toString() + Math.random().toString(36).substring(2, 6);
    const createdAt = new Date().toISOString();
    await phraseSheet.addRow({
      ID: id,
      Text: text.trim(),
      CreatedAt: createdAt
    });

    console.log(`[Google Sheets] Added new phrase: ${text}`);
    return res.json({ success: true, phrase: { id, text: text.trim(), createdAt } });
  } catch (error) {
    console.error('[Google Sheets] Error adding phrase:', error);
    return res.status(500).json({ error: 'Failed to add phrase to Google Sheets.' });
  }
});

// Google Sheets API: Delete Shared Quick Phrase
app.post('/api/phrases/delete', async (req, res) => {
  const { id } = req.body;
  if (!id) {
    return res.status(400).json({ error: 'ID is required.' });
  }

  if (!doc) {
    return res.status(503).json({ error: 'Google Sheets not configured.' });
  }

  try {
    await doc.loadInfo();
    const phraseSheet = doc.sheetsByTitle['자주사용하는문구'];
    if (!phraseSheet) {
      return res.status(404).json({ error: 'Sheet not found.' });
    }

    const rows = await phraseSheet.getRows();
    const rowToDelete = rows.find(row => row.get('ID') === id.toString());
    if (rowToDelete) {
      await rowToDelete.delete();
      console.log(`[Google Sheets] Deleted phrase ID: ${id}`);
      return res.json({ success: true });
    } else {
      return res.status(404).json({ error: 'Phrase not found.' });
    }
  } catch (error) {
    console.error('[Google Sheets] Error deleting phrase:', error);
    return res.status(500).json({ error: 'Failed to delete phrase from Google Sheets.' });
  }
});

// Create HTTP Server wrapped with Express
const server = http.createServer(app);

// Attach WebSocket Server to HTTP server
const wss = new ws.WebSocketServer({ noServer: true });

// Handle WebSocket upgrade on /ws
server.on('upgrade', (request, socket, head) => {
  const url = new URL(request.url, `http://${request.headers.host || 'localhost'}`);
  if (url.pathname === '/ws') {
    wss.handleUpgrade(request, socket, head, (wsConnection) => {
      wss.emit('connection', wsConnection, request);
    });
  } else {
    socket.destroy();
  }
});

// WebSocket Server Connection Handler (Gemini Live Proxy)
wss.on('connection', (clientWs, req) => {
  console.log('[WebSocket] Client connected.');

  const geminiKey = process.env.GEMINI_API_KEY;
  if (!geminiKey) {
    console.error('[WebSocket] GEMINI_API_KEY is not defined in environment variables.');
    clientWs.close(1008, 'GEMINI_API_KEY is missing on server');
    return;
  }

  // Connect to Google Gemini Live API
  const geminiUrl = `wss://generativelanguage.googleapis.com/ws/google.ai.generativelanguage.v1beta.GenerativeService.BidiGenerateContent?key=${geminiKey}`;
  const geminiWs = new ws.WebSocket(geminiUrl);

  let isGeminiOpen = false;
  const messageQueue = [];

  geminiWs.on('open', () => {
    console.log('[WebSocket] Connected to Gemini Live API upstream.');
    isGeminiOpen = true;

    // Send any messages queued while connection was establishing
    while (messageQueue.length > 0) {
      const msg = messageQueue.shift();
      geminiWs.send(msg);
    }
  });

  geminiWs.on('message', (data, isBinary) => {
    // Forward message from Gemini to Client
    if (clientWs.readyState === ws.WebSocket.OPEN) {
      try {
        const response = JSON.parse(data);
        if (response.serverContent) {
          const keys = Object.keys(response.serverContent);
          console.log(`[WebSocket] Received from Gemini: serverContent with keys: [${keys.join(', ')}]`);
          
          if (response.serverContent.inputTranscription && response.serverContent.inputTranscription.text !== undefined) {
            console.log(`[WebSocket] Gemini Input Transcript: "${response.serverContent.inputTranscription.text}"`);
          }
          if (response.serverContent.outputTranscription && response.serverContent.outputTranscription.text !== undefined) {
            console.log(`[WebSocket] Gemini Output Transcript: "${response.serverContent.outputTranscription.text}"`);
          }
          
          if (response.serverContent.modelTurn) {
            const parts = response.serverContent.modelTurn.parts;
            if (parts) {
              parts.forEach((part, index) => {
                const partKeys = Object.keys(part);
                console.log(`  -> Part ${index} keys: [${partKeys.join(', ')}]`);
                if (part.text) {
                  console.log(`  -> Part ${index} text: "${part.text}"`);
                }
                if (part.inlineData) {
                  console.log(`  -> Part ${index} inlineData: mimeType="${part.inlineData.mimeType}", length=${part.inlineData.data.length} chars`);
                }
              });
            }
          }
        } else {
          console.log('[WebSocket] Received from Gemini (non-serverContent):', Object.keys(response));
        }
      } catch (e) {
        console.log('[WebSocket] Error parsing Gemini message or binary message');
      }
      
      // Ensure we forward text messages as strings so the client browser receives them as strings
      clientWs.send(isBinary ? data : data.toString('utf-8'));
    }
  });

  geminiWs.on('close', (code, reason) => {
    console.log(`[WebSocket] Gemini upstream closed. Code: ${code}, Reason: ${reason}`);
    
    // WS specification restricts sending certain local close codes (e.g. 1005, 1006, 1015)
    const safeCode = (code >= 1000 && code <= 1011 && code !== 1004 && code !== 1005 && code !== 1006) ? code : 1000;
    const safeReason = reason ? reason.toString().substring(0, 100) : '';
    
    if (clientWs.readyState === ws.WebSocket.OPEN || clientWs.readyState === ws.WebSocket.CONNECTING) {
      clientWs.close(safeCode, safeReason);
    }
  });

  geminiWs.on('error', (err) => {
    console.error('[WebSocket] Gemini upstream error:', err);
    clientWs.close(1011, 'Internal connection error to upstream');
  });

  // Client messages
  clientWs.on('message', (message) => {
    try {
      const parsed = JSON.parse(message);
      if (parsed && parsed.type === 'ping') {
        clientWs.send(JSON.stringify({ type: 'pong' }));
        return;
      }
    } catch (e) {
      // Ignore parsing errors for binary / non-JSON data
    }

    if (isGeminiOpen) {
      if (geminiWs.readyState === ws.WebSocket.OPEN) {
        try {
          const parsed = JSON.parse(message);
          if (parsed.realtimeInput && parsed.realtimeInput.mediaChunks) {
            const dataLength = parsed.realtimeInput.mediaChunks[0].data.length;
            console.log(`[WebSocket] Received audio chunk: ${dataLength} base64 chars`);
          } else {
            console.log('[WebSocket] Received non-audio payload:', Object.keys(parsed));
          }
        } catch (e) {
          console.log('[WebSocket] Received binary/non-JSON message');
        }
        geminiWs.send(message);
      }
    } else {
      // Queue message until Gemini connection is open
      messageQueue.push(message);
    }
  });

  clientWs.on('close', () => {
    console.log('[WebSocket] Client disconnected.');
    if (geminiWs.readyState === ws.WebSocket.OPEN || geminiWs.readyState === ws.WebSocket.CONNECTING) {
      geminiWs.close();
    }
  });

  clientWs.on('error', (err) => {
    console.error('[WebSocket] Client WebSocket error:', err);
    if (geminiWs.readyState === ws.WebSocket.OPEN || geminiWs.readyState === ws.WebSocket.CONNECTING) {
      geminiWs.close();
    }
  });
});

// Start Server
server.listen(PORT, () => {
    console.log(`🚀 Server is running on http://localhost:${PORT}`);
    console.log(`🔌 WebSocket server is active on ws://localhost:${PORT}/ws`);
});
