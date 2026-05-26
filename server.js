require('dotenv').config();
const express = require('express');
const { middleware, Client } = require('@line/bot-sdk');
const cors    = require('cors');
const dayjs   = require('dayjs');
const path    = require('path');
const { google } = require('googleapis');

const app  = express();
const PORT = process.env.PORT || 3000;

// ─── LINE CONFIG ───────────────────────────────────────────
const lineConfig = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret:      process.env.LINE_CHANNEL_SECRET,
};
const lineClient = new Client(lineConfig);

// ─── GOOGLE SHEETS AUTH ────────────────────────────────────
const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    private_key:  (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
  },
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const SHEET_ID  = process.env.GOOGLE_SHEET_ID;
const BB_PRICE  = 250;   // ฿250/room/night flat (covers 2 pax)
const MAX_ROOMS = 5;     // max rooms per booking via Line OA

async function getSheets() {
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

// ─── MIDDLEWARE ────────────────────────────────────────────
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─── HELPERS ───────────────────────────────────────────────
function genId() {
  return `NR-${dayjs().format('YYYYMMDD')}-${Math.random().toString(36).substr(2,4).toUpperCase()}`;
}
function nightsBetween(ci, co) {
  return Math.round((new Date(co) - new Date(ci)) / 86400000);
}
function thaiDate(str) {
  const m = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  const d = dayjs(str);
  return `${d.date()} ${m[d.month()]} ${d.year()+543}`;
}

// ─── API: GET ROOMS ────────────────────────────────────────
app.get('/api/rooms', async (req, res) => {
  try {
    const { checkin } = req.query;
    const sheets = await getSheets();

    const avRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Availability_Control!A4:ZZ11',
    });
    const avRows  = avRes.data.values || [];
    const dateRow = avRows[0] || [];
    const colIdx  = checkin ? dateRow.indexOf(checkin) : -1;

    const riRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Room_Inventory!A4:G9',
    });
    const riRows = riRes.data.values || [];

    const rooms = riRows.map((row, i) => {
      const roomRow  = avRows[i + 1] || [];
      const openToday = colIdx >= 0 ? parseInt(roomRow[colIdx]) || 0 : parseInt(row[5]) || 0;
      const status = openToday === 0 ? 'CLOSED' : openToday <= 2 ? 'LOW' : 'OPEN';
      return {
        code:       row[0],
        name:       row[1],
        bedType:    row[2],
        baseRate:   parseInt(row[3]) || 0,
        totalRooms: parseInt(row[4]) || 0,
        openToday,
        status,
        canBook:    openToday > 0,
      };
    });

    res.json({ success: true, rooms, bbPrice: BB_PRICE, maxRooms: MAX_ROOMS });
  } catch (err) {
    console.error('GET /api/rooms:', err.message);
    res.status(500).json({ error: 'ไม่สามารถดึงข้อมูลห้องพักได้' });
  }
});

// ─── API: CREATE BOOKING ───────────────────────────────────
app.post('/api/bookings', async (req, res) => {
  try {
    const { lineUserId, guestName, roomCode, checkin, checkout,
            numRooms, breakfastOption, paymentType } = req.body;

    if (!lineUserId || !guestName || !roomCode || !checkin || !checkout)
      return res.status(400).json({ error: 'กรอกข้อมูลไม่ครบ' });

    if (numRooms > MAX_ROOMS)
      return res.status(400).json({
        error: `จองได้สูงสุด ${MAX_ROOMS} ห้อง/ครั้ง\nหากต้องการมากกว่านี้ กรุณาติดต่อฝ่ายขาย ${process.env.HOTEL_PHONE}`,
      });

    const sheets = await getSheets();

    // ตรวจสอบห้องว่างจาก Availability_Control
    const avRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: 'Availability_Control!A4:ZZ11',
    });
    const avRows  = avRes.data.values || [];
    const dateRow = avRows[0] || [];
    const colIdx  = dateRow.indexOf(checkin);
    const roomCodes = avRows.slice(1).map(r => r[0]);
    const rIdx   = roomCodes.indexOf(roomCode);
    const openCount = (colIdx >= 0 && rIdx >= 0) ? parseInt(avRows[rIdx+1][colIdx]) || 0 : 0;

    if (openCount < numRooms)
      return res.status(409).json({ error: 'ห้องว่างไม่เพียงพอสำหรับวันที่เลือก' });

    // ดึงราคาห้องจาก Room_Inventory
    const riRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: 'Room_Inventory!A4:G9',
    });
    const riRow = (riRes.data.values || []).find(r => r[0] === roomCode);
    if (!riRow) return res.status(404).json({ error: 'ไม่พบประเภทห้องพัก' });

    const baseRate = parseInt(riRow[3]) || 0;
    const roomName = riRow[1];
    const n        = nightsBetween(checkin, checkout);
    const roomCost = baseRate * numRooms * n;
    const bbCost   = breakfastOption === 'BB' ? BB_PRICE * numRooms * n : 0;
    const total    = roomCost + bbCost;
    const deposit  = Math.round(total * 0.5);
    const payAmt   = paymentType === 'full' ? total : deposit;
    const bookingId = genId();
    const now      = dayjs().format('YYYY-MM-DD HH:mm:ss');

    // บันทึกลง Bookings Sheet
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: 'Bookings!A:Y',
      valueInputOption: 'USER_ENTERED',
      resource: { values: [[
        bookingId, lineUserId, guestName, roomCode, roomName,
        checkin, checkout, n, numRooms, baseRate,
        roomCost, breakfastOption, bbCost, total,
        deposit, 0, total,
        'pending_payment', paymentType, '', '', now, 0, 'FALSE', '',
      ]] },
    });

    // ส่ง QR ให้ลูกค้า
    await lineClient.pushMessage(lineUserId,
      buildQRFlex(bookingId, roomName, checkin, checkout, numRooms, breakfastOption, payAmt, n));

    // แจ้ง Staff Group
    await lineClient.pushMessage(process.env.HOTEL_GROUP_LINE_ID,
      buildStaffFlex(bookingId, guestName, roomName, numRooms, checkin, checkout, payAmt, paymentType));

    res.json({ success: true, bookingId, total, payAmount: payAmt });
  } catch (err) {
    console.error('POST /api/bookings:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── API: GET USER BOOKINGS ────────────────────────────────
app.get('/api/bookings/user/:uid', async (req, res) => {
  try {
    const sheets = await getSheets();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: 'Bookings!A2:Y5000',
    });
    const rows = (r.data.values || []).filter(row => row[1] === req.params.uid);
    const bookings = rows.map(row => ({
      bookingId:    row[0],
      guestName:    row[2],
      roomCode:     row[3],
      roomName:     row[4] || row[3],
      checkin:      row[5],
      checkout:     row[6],
      numRooms:     parseInt(row[8])  || 1,
      bbOption:     row[11],
      total:        parseInt(row[13]) || 0,
      depositPaid:  parseInt(row[15]) || 0,
      remaining:    (parseInt(row[13])||0) - (parseInt(row[15])||0),
      status:       row[17],
      modifyCount:  parseInt(row[22]) || 0,
    }));
    res.json({ success: true, bookings });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── LINE WEBHOOK ──────────────────────────────────────────
app.post('/webhook', middleware(lineConfig), async (req, res) => {
  res.sendStatus(200);
  for (const event of (req.body.events || [])) {
    try {
      if (event.type === 'follow')   await handleFollow(event);
      if (event.type === 'postback') await handlePostback(event);
    } catch (err) { console.error('webhook:', err.message); }
  }
});

async function handleFollow(event) {
  await lineClient.replyMessage(event.replyToken, {
    type: 'text',
    text: '🏨 ยินดีต้อนรับสู่ Hotel Nara!\n\nกด "สำรองห้องพัก" ในเมนูด้านล่างเพื่อดูห้องว่างและจองได้เลยครับ 🛏️',
  });
}

async function handlePostback(event) {
  const { data } = event.postback;
  if (data.startsWith('confirm_')) await confirmBooking(data.replace('confirm_',''), event.replyToken);
  if (data.startsWith('reject_'))  await rejectBooking(data.replace('reject_',''),  event.replyToken);
}

async function confirmBooking(bookingId, replyToken) {
  const sheets = await getSheets();
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Bookings!A2:Y5000',
  });
  const rows = r.data.values || [];
  const idx  = rows.findIndex(row => row[0] === bookingId);
  if (idx < 0) return;
  const row = rows[idx]; const sheetR = idx + 2;
  const deposit = parseInt(row[14]) || 0;
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    resource: { valueInputOption: 'USER_ENTERED', data: [
      { range: `Bookings!R${sheetR}`, values: [['confirmed']] },
      { range: `Bookings!P${sheetR}`, values: [[deposit]] },
    ]},
  });
  await lineClient.pushMessage(row[1], buildConfirmFlex(row));
  await lineClient.replyMessage(replyToken, {
    type: 'text', text: `✅ ยืนยัน #${bookingId} แล้ว ส่งใบจองให้ลูกค้าเรียบร้อย`,
  });
}

async function rejectBooking(bookingId, replyToken) {
  const sheets = await getSheets();
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: 'Bookings!A2:B5000',
  });
  const rows = r.data.values || [];
  const idx  = rows.findIndex(row => row[0] === bookingId);
  if (idx < 0) return;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: `Bookings!R${idx+2}`,
    valueInputOption: 'USER_ENTERED', resource: { values: [['rejected']] },
  });
  await lineClient.pushMessage(rows[idx][1], {
    type: 'text',
    text: `⚠️ ไม่สามารถยืนยัน #${bookingId} ได้\nกรุณาตรวจสอบสลิปและแนบใหม่ หรือโทร ${process.env.HOTEL_PHONE}`,
  });
  await lineClient.replyMessage(replyToken, { type: 'text', text: `❌ ปฏิเสธ #${bookingId}` });
}

// ─── FLEX MESSAGE BUILDERS ─────────────────────────────────
const GREEN = '#1D9E75';

function rowFlex(label, value, color='#1a1a1a') {
  return {
    type:'box', layout:'horizontal', margin:'xs',
    contents:[
      { type:'text', text:label, size:'xs', color:'#888888', flex:3 },
      { type:'text', text:String(value), size:'xs', color, flex:4, align:'end', wrap:true },
    ],
  };
}

function buildQRFlex(id, room, ci, co, rooms, bb, amount, n) {
  const bbLabel = bb === 'BB'
    ? `รวมอาหารเช้า 2 ท่าน (+฿${BB_PRICE*rooms*n})`
    : 'ไม่รวมอาหารเช้า';
  return {
    type:'flex', altText:`💳 ชำระ #${id} | ฿${amount.toLocaleString()}`,
    contents:{
      type:'bubble',
      header:{ type:'box', layout:'vertical', backgroundColor:'#1a3a5c', contents:[
        { type:'text', text:'💳 ชำระเงิน', color:'#fff', weight:'bold', size:'md' },
        { type:'text', text:`เลขจอง: ${id}`, color:'#AACCFF', size:'xs' },
      ]},
      body:{ type:'box', layout:'vertical', spacing:'sm', alignItems:'center', contents:[
        { type:'text', text:`฿${amount.toLocaleString()}`, size:'xxl', weight:'bold', color:GREEN, align:'center' },
        { type:'text', text:`PromptPay: ${process.env.PROMPTPAY_NUMBER}`, size:'sm', align:'center', margin:'sm' },
        { type:'separator', margin:'md' },
        rowFlex('ห้องพัก', room),
        rowFlex('อาหารเช้า', bbLabel),
        rowFlex('วันที่', `${thaiDate(ci)} – ${thaiDate(co)} (${n} คืน)`),
        rowFlex('จำนวน', `${rooms} ห้อง`),
        { type:'text', text:'⏱ QR หมดอายุใน 15 นาที', size:'xs', color:'#D97706', align:'center', margin:'md' },
      ]},
      footer:{ type:'box', layout:'vertical', contents:[{
        type:'button', style:'primary', color:GREEN,
        action:{ type:'uri', label:'📎 แนบสลิปหลังโอน',
          uri:`${process.env.BASE_URL}/upload?id=${id}` },
      }]},
    },
  };
}

function buildStaffFlex(id, guest, room, rooms, ci, co, amount, payType) {
  const ptLabel = payType === 'full' ? '✅ เต็มจำนวน' : '💰 มัดจำ 50%';
  return {
    type:'flex', altText:`🔔 จองใหม่ #${id}`,
    contents:{
      type:'bubble',
      header:{ type:'box', layout:'vertical', backgroundColor:'#1a1a2e', contents:[
        { type:'text', text:'🔔 จองใหม่', color:'#fff', weight:'bold' },
        { type:'text', text:dayjs().format('DD/MM/YY HH:mm'), color:'#888', size:'xs' },
      ]},
      body:{ type:'box', layout:'vertical', spacing:'sm', contents:[
        rowFlex('เลขจอง', id, '#FFD700'),
        rowFlex('ลูกค้า', guest),
        rowFlex('ห้อง', `${room} × ${rooms}`),
        rowFlex('วันที่', `${thaiDate(ci)} – ${thaiDate(co)}`),
        { type:'separator' },
        rowFlex('ยอด', `฿${amount.toLocaleString()} (${ptLabel})`, GREEN),
        { type:'text', text:'สถานะ: ⏳ รอตรวจสลิป', size:'xs', color:'#D97706', margin:'sm' },
      ]},
      footer:{ type:'box', layout:'horizontal', spacing:'sm', contents:[
        { type:'button', style:'primary',   color:GREEN, height:'sm',
          action:{ type:'postback', label:'✅ ยืนยัน', data:`confirm_${id}` }},
        { type:'button', style:'secondary', height:'sm',
          action:{ type:'postback', label:'❌ ปฏิเสธ', data:`reject_${id}` }},
      ]},
    },
  };
}

function buildConfirmFlex(row) {
  const total = parseInt(row[13]) || 0;
  const paid  = parseInt(row[15]) || 0;
  const rem   = total - paid;
  return {
    type:'flex', altText:`✅ ยืนยัน #${row[0]}`,
    contents:{
      type:'bubble',
      header:{ type:'box', layout:'vertical', backgroundColor:'#059142', contents:[
        { type:'text', text:'✅ ยืนยันการจองสำเร็จ', color:'#fff', weight:'bold', size:'md' },
        { type:'text', text:`เลขจอง: ${row[0]}`, color:'#CCFFCC', size:'xs' },
      ]},
      body:{ type:'box', layout:'vertical', spacing:'sm', contents:[
        rowFlex('ห้องพัก', row[4]||row[3]),
        rowFlex('อาหารเช้า', row[11]==='BB' ? 'รวมอาหารเช้า 2 ท่าน' : 'ไม่รวมอาหารเช้า'),
        rowFlex('เช็คอิน',   `${thaiDate(row[5])} | 14:00 น.`),
        rowFlex('เช็คเอาต์', `${thaiDate(row[6])} | 12:00 น.`),
        { type:'separator' },
        rowFlex('ราคารวม',       `฿${total.toLocaleString()}`,    GREEN),
        rowFlex('มัดจำชำระแล้ว', `฿${paid.toLocaleString()}`,     GREEN),
        rem > 0
          ? rowFlex('ค้างชำระ (วันเข้าพัก)', `฿${rem.toLocaleString()}`, '#D97706')
          : { type:'text', text:'✅ ชำระครบแล้ว', color:GREEN, size:'xs', margin:'sm' },
        { type:'text',
          text:'💡 ชำระส่วนที่เหลือได้ที่โรงแรม: เงินสด / โอน / บัตรเครดิต',
          size:'xxs', color:'#888', wrap:true, margin:'sm' },
      ]},
      footer:{ type:'box', layout:'vertical', contents:[{
        type:'button', style:'primary', color:GREEN,
        action:{ type:'uri', label:'📋 จัดการการจอง',
          uri:`${process.env.BASE_URL}/manage` },
      }]},
    },
  };
}

// ─── SERVE LIFF SPA ────────────────────────────────────────
app.get('*', (req, res) =>
  res.sendFile(path.join(__dirname, 'public', 'index.html')));

app.listen(PORT, () =>
  console.log(`🏨 Hotel Nara running on port ${PORT}`));
