/**
 * GHI GIỜ THU ĐẦU TIÊN / GIỜ HOÀN THÀNH 100% CỦA TỪNG ĐƠN VỊ – phục vụ bảng xếp hạng của website.
 * Dán vào dự án Apps Script của xã (file mới, ví dụ ThuDauTien.gs), rồi chạy hàm ftSetup() một lần (cấp quyền).
 * Cứ mỗi phút ftSnapshot() đọc sheet tháng hiện tại và ghi vào tab "Thu đầu tiên":
 * Tháng | Đơn vị | Bắt đầu thu | Hoàn thành 100%   (giờ Việt Nam, yyyy-MM-dd HH:mm:ss).
 * Mỗi ô chỉ ghi MỘT LẦN (lúc đơn vị lần đầu có đảng viên nộp / lần đầu đủ 100%), không ghi đè về sau.
 * Lần chạy đầu: đơn vị đã có số liệu từ trước thì lấy mốc từ tab "Lịch sử ngày" (ngày đầu tiên, tính 23:00:00), nếu không có thì lấy giờ chạy.
 */
const FT = { tab: 'Thu đầu tiên', hist: 'Lịch sử ngày', tz: 'Asia/Ho_Chi_Minh', y0: 2026 };
const FT_COLS = { name: 'Tên đơn vị', members: 'Tổng sl đảng viên (1)', paidCount: 'Đã duyệt, đã nộp (8)',
  exempt: 'Số lượng được miễn giảm ĐP (5)', moved: 'Số lượng ĐV đã chuyển sinh hoạt đảng tạm thời (4)' };
function ftNum_(x) {
  if (typeof x === 'number') return isFinite(x) ? x : 0;
  const t = String(x == null ? '' : x).trim().replace(/\s/g, '').replace('%', '');
  if (!t) return 0;
  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) return Number(t.replace(/\./g, '').replace(',', '.')) || 0;
  if (/^-?\d+,\d+$/.test(t)) return Number(t.replace(',', '.')) || 0;
  return Number(t) || 0;
}
function ftKey_(s) { return String(s || '').toLowerCase().replace(/[^\p{L}\p{N}]+/gu, ''); }
function ftSnapshot() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return 'Đang bận';
  try {
    const ss = SpreadsheetApp.getActive(), now = new Date();
    const y = Number(Utilities.formatDate(now, FT.tz, 'yyyy')), mo = Number(Utilities.formatDate(now, FT.tz, 'M'));
    const src = ss.getSheetByName('Thang' + mo + (y === FT.y0 ? '' : '-' + y));
    if (!src) return 'Chưa có sheet tháng ' + mo;
    const vals = src.getDataRange().getValues();
    const hi = vals.findIndex(r => r.some(c => String(c).trim() === FT_COLS.name));
    if (hi < 0) return 'Không thấy dòng tiêu đề';
    const H = vals[hi].map(c => String(c).trim()), ix = k => H.indexOf(FT_COLS[k]);
    const ym = Utilities.formatDate(now, FT.tz, 'yyyy-MM'), stamp = Utilities.formatDate(now, FT.tz, 'yyyy-MM-dd HH:mm:ss');

    /* mốc sớm nhất trong "Lịch sử ngày" của tháng này, dùng để bù cho lần chạy đầu */
    const hs = {};
    const hsh = ss.getSheetByName(FT.hist);
    if (hsh && hsh.getLastRow() > 1) {
      hsh.getRange(2, 1, hsh.getLastRow() - 1, 4).getValues().forEach(r => {
        const d = String(r[0] instanceof Date ? Utilities.formatDate(r[0], FT.tz, 'yyyy-MM-dd') : r[0]).trim();
        if (d.slice(0, 7) !== ym || !r[1]) return;
        const k = ftKey_(r[1]), due = ftNum_(r[2]), paid = ftNum_(r[3]), e = hs[k] || (hs[k] = {});
        if (paid > 0 && (!e.s || d < e.s)) e.s = d;
        if (due > 0 && paid >= due && (!e.d || d < e.d)) e.d = d;
      });
    }

    let sh = ss.getSheetByName(FT.tab);
    if (!sh) { sh = ss.insertSheet(FT.tab); sh.appendRow(['Tháng', 'Đơn vị', 'Bắt đầu thu', 'Hoàn thành 100%']); sh.getRange('A:D').setNumberFormat('@'); }
    const last = sh.getLastRow();
    const cur = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
    const at = {};
    cur.forEach((r, i) => { if (String(r[0]) === ym) at[ftKey_(r[1])] = i; });

    let changed = 0;
    vals.slice(hi + 1).forEach(r => {
      const name = String(r[ix('name')] || '').trim(); if (!name) return;
      const g = k => ix(k) < 0 ? 0 : ftNum_(r[ix(k)]);
      const due = Math.max(0, g('members') - g('exempt') - g('moved')), paid = g('paidCount');
      if (due <= 0) return;
      const k = ftKey_(name), startNow = paid > 0, doneNow = paid >= due;
      if (!startNow && !doneNow) return;
      const h = hs[k] || {};
      const t = d => d ? d + ' 23:00:00' : stamp;
      if (k in at) {
        const row = cur[at[k]];
        let ch = false;
        if (startNow && !row[2]) { row[2] = t(h.s); ch = true; }
        if (doneNow && !row[3]) { row[3] = t(h.d); if (!row[2]) row[2] = row[3]; ch = true; }
        if (ch) { sh.getRange(at[k] + 2, 3, 1, 2).setValues([[row[2], row[3]]]); changed++; }
      } else {
        const s = startNow ? t(h.s) : '', d = doneNow ? t(h.d) : '';
        sh.appendRow([ym, name, s || d, d]);
        changed++;
      }
    });
    return 'Đã cập nhật ' + changed + ' đơn vị lúc ' + stamp;
  } finally { lock.releaseLock(); }
}
function ftSetup() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'ftSnapshot').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('ftSnapshot').timeBased().everyMinutes(1).create();
  return ftSnapshot();
}
