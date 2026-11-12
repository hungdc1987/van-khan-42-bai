// Văn Khấn Việt – nhận đơn từ trang 42bai.vankhanviet.store vào Google Sheet "Van khan"
// Giá lấy từ phía máy chủ, khách không sửa được từ trình duyệt.
const CONFIG = {
  spreadsheetId: '1-SITDWuQAFeVCt4ZSzzm2ZZMdtqwHkEx1WllrrTNvZw',
  sheetName: 'Trang tính1',
  recipient: 'hungdc1987@gmail.com',
  products: {
    '42-bai': { name: 'Văn khấn toàn tập (42 bài)', price: 179000 }
  },
  maxOrdersPerPhonePerHour: 3
};

function authorize() {
  SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.sheetName);
  MailApp.getRemainingDailyQuota();
}

function jsonResponse(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function plain(value, max) {
  const s = String(value || '').trim();
  if (!s || s.length > max) throw new Error('Thông tin không hợp lệ');
  return /^[=+@-]/.test(s) ? "'" + s : s;
}

function doGet() {
  return jsonResponse({ service: 'Van Khan orders', status: 'ready' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const raw = (e && e.postData && e.postData.contents) || '';
    if (raw.length > 5000) throw new Error('Yêu cầu quá dài');
    const data = JSON.parse(raw);

    // Ô ẩn chống bot: người thật không điền ô này
    if (data.website) return jsonResponse({ ok: true, orderId: 'ignored' });

    const product = CONFIG.products[String(data.product || '')];
    if (!product) throw new Error('Sản phẩm không hợp lệ');

    const name = plain(data.name, 100);
    const address = plain(data.address, 400);
    const phone = String(data.phone || '').replace(/\s/g, '');
    if (!/^(0\d{9}|\+84\d{9})$/.test(phone)) throw new Error('Số điện thoại không hợp lệ');
    const quantity = Number(data.quantity);
    if (!Number.isInteger(quantity) || quantity < 1 || quantity > 20) throw new Error('Số lượng không hợp lệ');
    const requestId = String(data.requestId || '');
    if (!/^[a-zA-Z0-9-]{16,80}$/.test(requestId)) throw new Error('Mã yêu cầu không hợp lệ');

    lock.waitLock(20000);
    const sheet = SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.sheetName);
    if (!sheet) throw new Error('Không tìm thấy trang nhận đơn');

    const last = sheet.getLastRow();
    const rows = last > 1 ? sheet.getRange(2, 1, last - 1, 5).getValues() : [];

    // Gửi lại cùng một yêu cầu thì trả về đơn cũ, không ghi trùng
    const dup = rows.find(r => String(r[0]) === requestId);
    if (dup) return jsonResponse({ ok: true, orderId: dup[1], duplicate: true });

    // Giới hạn số đơn mỗi số điện thoại trong 1 giờ
    const since = Date.now() - 3600 * 1000;
    const recent = rows.filter(r => String(r[4]).replace(/^'/, '') === phone && new Date(r[2]).getTime() > since);
    if (recent.length >= CONFIG.maxOrdersPerPhonePerHour) throw new Error('Quá nhiều yêu cầu');

    const orderId = 'VKV-' + Utilities.getUuid();
    const time = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
    const SHIPPING = 25000; // phí ship đồng giá toàn quốc
    const total = quantity * product.price + SHIPPING;
    sheet.appendRow([requestId, orderId, time, name, "'" + phone, address, quantity,
      product.price, total, 'Mới', 'Chờ gửi', 'Website 42bai.vankhanviet.store', product.name]);
    const row = sheet.getLastRow();
    SpreadsheetApp.flush();

    let emailSent = false;
    try {
      MailApp.sendEmail({
        to: CONFIG.recipient,
        subject: '[Văn Khấn Việt] Đơn mới ' + orderId,
        body: ['Đơn: ' + orderId, 'Sản phẩm: ' + product.name, 'Thời gian: ' + time,
          'Khách: ' + name, 'SĐT: ' + phone, 'Địa chỉ: ' + address, 'Số lượng: ' + quantity,
          'Tiền sách: ' + (quantity * product.price).toLocaleString('vi-VN') + 'đ', 'Phí giao hàng: 25.000đ',
          'Tổng thanh toán: ' + total.toLocaleString('vi-VN') + 'đ',
          'Quà tặng: app Văn Khấn Việt',
          'https://docs.google.com/spreadsheets/d/' + CONFIG.spreadsheetId + '/edit'].join('\n')
      });
      emailSent = true;
      sheet.getRange(row, 11).setValue('Đã gửi');
    } catch (mailError) {
      sheet.getRange(row, 11).setValue('Chưa gửi được — kiểm tra hạn mức/quyền MailApp');
    }
    return jsonResponse({ ok: true, orderId: orderId, emailSent: emailSent });
  } catch (error) {
    return jsonResponse({ ok: false, error: 'Không thể tiếp nhận yêu cầu. Vui lòng thử lại hoặc gọi 0989 830 947.' });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// ===== Bản đồ đơn hàng theo tỉnh (chỉ trả số đếm, không có tên/SĐT) =====
const PROVINCES = [
  {n:"Hà Nội",a:["ha noi", "gia lam", "minh khai", "tao quang buu", "ngoc liep", "quoc oai", "hai ba trung"],lat:21.03,lon:105.85},
  {n:"Hà Giang",a:["ha giang"],lat:22.83,lon:104.98},
  {n:"Cao Bằng",a:["cao bang"],lat:22.67,lon:106.26},
  {n:"Lai Châu",a:["lai chau"],lat:22.39,lon:103.46},
  {n:"Điện Biên",a:["dien bien"],lat:21.39,lon:103.02},
  {n:"Lào Cai",a:["lao cai"],lat:22.48,lon:103.97},
  {n:"Tuyên Quang",a:["tuyen quang"],lat:21.82,lon:105.21},
  {n:"Lạng Sơn",a:["lang son"],lat:21.85,lon:106.76},
  {n:"Quảng Ninh",a:["quang ninh", "mong cai", "ha long"],lat:21.01,lon:107.3},
  {n:"Bắc Giang",a:["bac giang"],lat:21.27,lon:106.19},
  {n:"Phú Thọ",a:["phu tho", "viet tri"],lat:21.32,lon:105.4},
  {n:"Thái Nguyên",a:["thai nguyen"],lat:21.59,lon:105.84},
  {n:"Bắc Ninh",a:["bac ninh"],lat:21.19,lon:106.07},
  {n:"Hải Phòng",a:["hai phong"],lat:20.86,lon:106.68},
  {n:"Hải Dương",a:["hai duong"],lat:20.94,lon:106.33},
  {n:"Hưng Yên",a:["hung yen"],lat:20.65,lon:106.05},
  {n:"Thái Bình",a:["thai binh"],lat:20.45,lon:106.34},
  {n:"Hà Nam",a:["ha nam"],lat:20.54,lon:105.91},
  {n:"Nam Định",a:["nam dinh"],lat:20.42,lon:106.17},
  {n:"Ninh Bình",a:["ninh binh"],lat:20.25,lon:105.97},
  {n:"Thanh Hóa",a:["thanh hoa", "thanb hoa"],lat:19.81,lon:105.78},
  {n:"Nghệ An",a:["nghe an", "nghe tinh"],lat:19.23,lon:104.92},
  {n:"Hà Tĩnh",a:["ha tinh"],lat:18.34,lon:105.9},
  {n:"Quảng Bình",a:["quang binh"],lat:17.47,lon:106.62},
  {n:"Quảng Trị",a:["quang tri"],lat:16.8,lon:107.1},
  {n:"Thừa Thiên Huế",a:["thua thien hue", "hue"],lat:16.46,lon:107.59},
  {n:"Đà Nẵng",a:["da nang"],lat:16.05,lon:108.21},
  {n:"Quảng Nam",a:["quang nam"],lat:15.57,lon:108.47},
  {n:"Quảng Ngãi",a:["quang ngai"],lat:15.12,lon:108.8},
  {n:"Bình Định",a:["binh dinh"],lat:14.17,lon:108.9},
  {n:"Phú Yên",a:["phu yen"],lat:13.09,lon:109.32},
  {n:"Khánh Hòa",a:["khanh hoa", "nha trang"],lat:12.24,lon:109.19},
  {n:"Ninh Thuận",a:["ninh thuan"],lat:11.67,lon:108.86},
  {n:"Bình Thuận",a:["binh thuan"],lat:10.93,lon:108.1},
  {n:"Kon Tum",a:["kon tum"],lat:14.35,lon:108.0},
  {n:"Gia Lai",a:["gia lai"],lat:13.98,lon:108.0},
  {n:"Đắk Lắk",a:["dak lak", "buon ma thuot"],lat:12.67,lon:108.04},
  {n:"Đắk Nông",a:["dak nong"],lat:12.0,lon:107.69},
  {n:"Lâm Đồng",a:["lam dong", "da lat", "thanh nhat"],lat:11.94,lon:108.44},
  {n:"Bình Phước",a:["binh phuoc"],lat:11.75,lon:106.72},
  {n:"Tây Ninh",a:["tay ninh"],lat:11.31,lon:106.1},
  {n:"Bình Dương",a:["binh duong"],lat:11.17,lon:106.67},
  {n:"Đồng Nai",a:["dong nai", "bien hoa"],lat:11.0,lon:107.0},
  {n:"Bà Rịa - Vũng Tàu",a:["ba ria", "vung tau"],lat:10.54,lon:107.24},
  {n:"TP Hồ Chí Minh",a:["ho chi minh", "tp hcm", "tphcm", "sai gon", "hcm", "tan phu", "tang nhon phu", "quan 1", "quan 3", "tan binh", "binh thanh"],lat:10.78,lon:106.7},
  {n:"Long An",a:["long an"],lat:10.54,lon:106.41},
  {n:"Tiền Giang",a:["tien giang"],lat:10.36,lon:106.36},
  {n:"Bến Tre",a:["ben tre"],lat:10.24,lon:106.38},
  {n:"Trà Vinh",a:["tra vinh"],lat:9.94,lon:106.34},
  {n:"Vĩnh Long",a:["vinh long"],lat:10.25,lon:105.96},
  {n:"Đồng Tháp",a:["dong thap"],lat:10.46,lon:105.63},
  {n:"An Giang",a:["an giang"],lat:10.52,lon:105.13},
  {n:"Kiên Giang",a:["kien giang"],lat:10.01,lon:105.08},
  {n:"Cần Thơ",a:["can tho"],lat:10.05,lon:105.79},
  {n:"Hậu Giang",a:["hau giang"],lat:9.78,lon:105.47},
  {n:"Sóc Trăng",a:["soc trang"],lat:9.6,lon:105.97},
  {n:"Bạc Liêu",a:["bac lieu"],lat:9.29,lon:105.73},
  {n:"Cà Mau",a:["ca mau"],lat:9.18,lon:105.15},
  {n:"Sơn La",a:["son la"],lat:21.33,lon:103.91},
  {n:"Yên Bái",a:["yen bai"],lat:21.72,lon:104.9},
  {n:"Hòa Bình",a:["hoa binh"],lat:20.81,lon:105.34},
  {n:"Bắc Kạn",a:["bac kan"],lat:22.14,lon:105.83},
  {n:"Vĩnh Phúc",a:["vinh phuc"],lat:21.31,lon:105.6}
];
function normVN_(s) {
  return ' ' + String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/[^a-z0-9]+/g, ' ').trim() + ' ';
}
function provinceOf_(address) {
  const a = normVN_(address);
  const pairs = [];
  PROVINCES.forEach(p => p.a.forEach(k => pairs.push([p.n, ' ' + k + ' '])));
  pairs.sort((x, y) => y[1].length - x[1].length); // khớp cụm dài trước
  for (const [n, k] of pairs) if (a.indexOf(k) !== -1) return n;
  return null;
}
function getProvinceMap_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('provmap');
  if (hit) return JSON.parse(hit);
  const sheet = SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.sheetName);
  const last = sheet.getLastRow();
  const counts = {};
  let total = 0, unknown = 0;
  const since = Date.now() - 30 * 24 * 3600 * 1000;
  if (last > 1) {
    const rows = sheet.getRange(2, 1, last - 1, 6).getValues(); // C = thời gian, F = địa chỉ
    rows.forEach(r => {
      if (!(new Date(r[2]).getTime() >= since)) return;
      total++;
      const p = provinceOf_(r[5]);
      if (p) counts[p] = (counts[p] || 0) + 1; else unknown++;
    });
  }
  const out = { updated: new Date().toISOString(), days: 30, total: total, unknown: unknown, counts: counts };
  cache.put('provmap', JSON.stringify(out), 300);
  return out;
}
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'map') {
    return ContentService.createTextOutput(JSON.stringify(getProvinceMap_())).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
