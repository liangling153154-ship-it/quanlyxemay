/**
 * SỔ THUÊ XE — Google Apps Script Web App
 * Database tạm trên Google Sheets, đồng bộ từ app index.html
 *
 * ============ CÁCH CÀI ĐẶT ============
 * 1. Mở Google Sheets > tạo 1 file mới (hoặc dùng file có sẵn).
 * 2. Trên thanh menu: Tiện ích mở rộng (Extensions) > Apps Script.
 * 3. Xóa hết code mẫu, dán TOÀN BỘ file này vào.
 * 4. Bấm Lưu (biểu tượng đĩa).
 * 5. Bấm "Triển khai" (Deploy) > "Tùy chọn triển khai mới" (New deployment).
 *    - Loại (Select type, biểu tượng bánh răng): chọn "Ứng dụng web" (Web app).
 *    - Mô tả: tùy ý (vd "So thue xe").
 *    - Thực thi với tư cách (Execute as): Tôi (Me / chính email của bạn).
 *    - Ai có quyền truy cập (Who has access): "Bất kỳ ai" (Anyone).
 *      >>> BẮT BUỘC chọn "Anyone" thì app mới gọi được, không thì sẽ lỗi.
 * 6. Bấm Triển khai > cấp quyền (Authorize) > chọn tài khoản > Nâng cao
 *    (Advanced) > "Đi tới ... (không an toàn)" > Cho phép (Allow).
 * 7. Copy "URL ứng dụng web" (kết thúc bằng /exec).
 * 8. Mở app Sổ thuê xe > Cài đặt > dán URL vào ô "Google Sheets URL".
 *
 * Khi bạn sửa code này, phải Deploy lại:
 *   Triển khai > Quản lý triển khai > (bút chì) > Phiên bản: Mới > Triển khai.
 * (Giữ nguyên URL cũ.)
 */

// Tên sheet (tab) lưu dữ liệu. GAS tự tạo nếu chưa có.
var SHEET_NAME = 'Rentals';

// Thứ tự cột. Cột đầu (id) là khóa để upsert.
var HEADERS = [
  'id',           // A - khóa định danh, do app sinh
  'ngay_chot',    // B - ngày đẩy/cập nhật lên sheet
  'loai',         // C - "Thuê xe" | "Ứng trước"
  'trang_thai',   // D - Dự kiến | Đang thuê | Cần TT | Đã hủy
  'xe',           // E - tên xe
  'ghi_chu_xe',   // F
  'khach',        // G
  'tu_ngay',      // H
  'den_ngay',     // I
  'so_ngay',      // J
  'gia_ngay',     // K
  'thanh_tien',   // L - số tiền (thuê dương, ứng trước âm)
  'mo_ta_chi_phi',// M - cho khoản ứng trước
  'json'          // N - dữ liệu gốc của app (đừng sửa tay cột này)
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // tránh ghi đè khi nhiều request tới cùng lúc
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'upsert';

    if (action === 'upsert') {
      return json(upsertRow(body.row));
    } else if (action === 'delete') {
      return json(deleteRow(body.id));
    } else if (action === 'list') {
      return json(listRows());
    } else if (action === 'ping') {
      return json({ ok: true, msg: 'pong' });
    }
    return json({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Cho phép mở URL bằng trình duyệt để kiểm tra nhanh (GET trả về trạng thái).
function doGet() {
  return json({ ok: true, msg: 'So thue xe web app dang chay.' });
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  // Đảm bảo có hàng tiêu đề (và nâng cấp khi thêm cột mới)
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    var head = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    if (head[HEADERS.length - 1] !== HEADERS[HEADERS.length - 1]) {
      sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
  }
  return sh;
}

// Trả về toàn bộ lượt (đọc từ cột json) để app nạp khi mở.
function listRows() {
  var sh = getSheet();
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, rows: [] };
  var jsonCol = HEADERS.indexOf('json') + 1;
  var vals = sh.getRange(2, jsonCol, last - 1, 1).getValues();
  var rows = [];
  for (var i = 0; i < vals.length; i++) {
    var cell = vals[i][0];
    if (!cell) continue;
    try { rows.push(JSON.parse(cell)); } catch (e) {}
  }
  return { ok: true, rows: rows };
}

// Tìm số dòng (1-based) theo id, trả 0 nếu không có.
function findRowById(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues(); // cột A từ dòng 2
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

// row: object có các khóa trùng HEADERS. Tạo mới nếu chưa có id, ngược lại cập nhật.
function upsertRow(row) {
  if (!row || !row.id) return { ok: false, error: 'Thiếu id' };
  var sh = getSheet();
  var values = HEADERS.map(function (h) {
    return row[h] !== undefined && row[h] !== null ? row[h] : '';
  });
  var rowNum = findRowById(sh, row.id);
  if (rowNum === 0) {
    sh.appendRow(values);
    return { ok: true, mode: 'insert', id: row.id };
  } else {
    sh.getRange(rowNum, 1, 1, HEADERS.length).setValues([values]);
    return { ok: true, mode: 'update', id: row.id, rowNum: rowNum };
  }
}

function deleteRow(id) {
  if (!id) return { ok: false, error: 'Thiếu id' };
  var sh = getSheet();
  var rowNum = findRowById(sh, id);
  if (rowNum === 0) return { ok: true, mode: 'noop', id: id }; // đã không còn
  sh.deleteRow(rowNum);
  return { ok: true, mode: 'delete', id: id };
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
