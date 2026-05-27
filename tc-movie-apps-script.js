// ============================================================
// TC MOVIE RATING SYSTEM — Google Apps Script
// วางโค้ดนี้ใน Extensions > Apps Script แล้ว Deploy
// ============================================================

const SHEET_NAME = 'ratings';

// ============================================================
// GET — ดึงข้อมูลสำหรับแสดงผลหน้าเว็บ
// ============================================================
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'get') {
    return buildResponse(getSummaryData());
  }

  // default: ส่งหน้าเว็บ (ไม่ได้ใช้ในกรณีนี้)
  return ContentService.createTextOutput('TC Movie API OK');
}

// ============================================================
// POST — รับคะแนนจากหน้าเว็บ
// ============================================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    saveRating(payload);
    return buildResponse({ status: 'ok' });
  } catch (err) {
    return buildResponse({ status: 'error', message: err.message });
  }
}

// ============================================================
// SAVE — บันทึกแถวใหม่ใน Sheet
// ============================================================
function saveRating(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  sheet.appendRow([
    new Date(),           // timestamp
    data.movie_id    || '',
    data.movie_title || '',
    data.movie_genre || '',
    data.movie_year  || '',
    Number(data.overall) || 0,
    Number(data.story)   || 0,
    Number(data.acting)  || 0,
    Number(data.camera)  || 0,
    Number(data.sound)   || 0,
    Number(data.edit)    || 0,
    data.comment   || '',
    data.reviewer  || 'Anonymous',
  ]);
}

// ============================================================
// GET SUMMARY — คำนวณสถิติ, leaderboard, recent reviews
// ============================================================
function getSummaryData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const rows  = sheet.getDataRange().getValues();

  if (rows.length <= 1) {
    return {
      stats:       { total: 0, avg: 0 },
      leaderboard: [],
      recent:      [],
    };
  }

  // ข้ามแถวหัว (index 0)
  const data = rows.slice(1);

  // --- สถิติรวม ---
  const total   = data.length;
  const sumAll  = data.reduce((s, r) => s + (Number(r[5]) || 0), 0);
  const avg     = total > 0 ? sumAll / total : 0;

  // --- Leaderboard: จัดกลุ่มตาม movie_id ---
  const movieMap = {};
  data.forEach(r => {
    const id    = r[1];
    const title = r[2];
    const genre = r[3];
    const score = Number(r[5]) || 0;

    if (!movieMap[id]) {
      movieMap[id] = { id, title, genre, sum: 0, count: 0 };
    }
    movieMap[id].sum   += score;
    movieMap[id].count += 1;
  });

  const leaderboard = Object.values(movieMap)
    .map(m => ({
      id:    m.id,
      title: m.title,
      genre: m.genre,
      avg:   m.count > 0 ? m.sum / m.count : 0,
      votes: m.count,
    }))
    .sort((a, b) => b.avg - a.avg || b.votes - a.votes);

  // --- Recent reviews: 5 รายการล่าสุด ---
  const recent = data.slice(-5).reverse().map(r => ({
    title:    r[2],
    overall:  Number(r[5]) || 0,
    comment:  r[11] || '',
    reviewer: r[12] || 'Anonymous',
  }));

  return {
    stats: { total, avg },
    leaderboard,
    recent,
  };
}

// ============================================================
// HELPER — สร้าง JSON Response พร้อม CORS
// ============================================================
function buildResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
