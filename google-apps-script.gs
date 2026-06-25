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
  'thuoc_luot',   // N - chi phí gắn vào lượt thuê nào
  'json'          // O - dữ liệu gốc của app (đừng sửa tay cột này)
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // tránh ghi đè khi nhiều request tới cùng lúc
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'upsert';
    var base = body.base; // mã cơ sở: chọn tab dữ liệu (vd 'hoaan', 'sens')

    if (action === 'upsert') {
      return json(upsertRow(body.row, base));
    } else if (action === 'delete') {
      return json(deleteRow(body.id, base));
    } else if (action === 'list') {
      return json(listRows(base));
    } else if (action === 'ping') {
      return json({ ok: true, msg: 'pong', sheet: sheetNameForBase(base) });
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

// Mỗi cơ sở dùng 1 tab riêng trong cùng file Sheet.
// hoaan / mặc định → tab 'Rentals' (giữ nguyên dữ liệu cũ).
function sheetNameForBase(base) {
  base = String(base || '').toLowerCase();
  if (base === 'sens') return 'Rentals_Sens';
  return SHEET_NAME; // 'Rentals' cho Hòa An và mọi request cũ chưa có base
}

function getSheet(base) {
  var name = sheetNameForBase(base);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  // Sheet trống → tạo hàng tiêu đề
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // Nâng cấp: thêm cột 'thuoc_luot' trước 'json' nếu sheet cũ chưa có.
  // Chèn cột thật để dữ liệu 'json' không bị lệch ô.
  var curW = sh.getLastColumn();
  var head = sh.getRange(1, 1, 1, curW).getValues()[0];
  if (head.indexOf('thuoc_luot') === -1) {
    var jsonIdx = head.indexOf('json'); // 0-based
    if (jsonIdx !== -1) {
      sh.insertColumnBefore(jsonIdx + 1);              // chèn cột trống trước json
      sh.getRange(1, jsonIdx + 1).setValue('thuoc_luot');
    } else {
      sh.getRange(1, curW + 1).setValue('thuoc_luot'); // không có json: thêm cuối
    }
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

// Trả về toàn bộ lượt (đọc từ cột json) để app nạp khi mở.
function listRows(base) {
  var sh = getSheet(base);
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
function upsertRow(row, base) {
  if (!row || !row.id) return { ok: false, error: 'Thiếu id' };
  var sh = getSheet(base);
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

function deleteRow(id, base) {
  if (!id) return { ok: false, error: 'Thiếu id' };
  var sh = getSheet(base);
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
