/************ 1. CẤU HÌNH ************/
const SH_FORM = 'Form';
const SH_DB   = 'Đơn hàng';
const WITH_DAYS_LEFT = true;        // false nếu KHÔNG dùng cột K

/***** E-MAIL cảnh báo hết hạn *****/
const ADMIN_EMAIL = 'samsung2015.vp@gmail.com';  // địa chỉ nhận báo cáo
const ALERT_THRESHOLD = 1;   // 0 hoặc 1 ngày

/***** BACKUP SHEET *****/
const BACKUP_SHEET_ID = '1ZP7_ySuZUETS92OYvvAP9X_sl5pobfZSMIYFcL6-G88'; // Thay bằng ID của sheet backup
const BACKUP_SHEET_NAME = 'Backup'; // Tên sheet trong file backup

/***** CẤU HÌNH LỌC DỮ LIỆU - MỚI *****/
const FILTER_CONFIG_SHEET = 'Filter_Config'; // Sheet lưu cấu hình lọc

/* vị trí cột (1‑based) – ĐÃ CẬP NHẬT THÊM MÃ ĐƠN HÀNG */
const COL = {                       
  stt     : 1,   // A - STT
  orderId : 2,   // B - MÃ ĐƠN HÀNG (MỚI)
  name    : 3,   // C - TÊN KHÁCH  
  contact : 4,   // D - HÌNH THỨC LIÊN HỆ
  pack    : 5,   // E - LOẠI TÀI KHOẢN
  price   : 6,   // F - GIÁ
  cost    : 7,   // G - VỐN
  paid    : 8,   // H - Trạng thái thanh toán
  buy     : 9,   // I - NGÀY MUA
  exp     : 10,  // J - NGÀY HẾT HẠN
  left    : 11,  // K - DaysLeft (nếu có)
  note    : WITH_DAYS_LEFT ? 12 : 11   // L hoặc K - GHI CHÚ
};

/****************************************/

// Hàm tiện ích
function SH(n){ return SpreadsheetApp.getActive().getSheetByName(n); }

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const row = payload.row;  // row number trong sheet
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_DB);
  // Tạo orderId & ghi vào sheet
  const orderId = createUniqueOrderId();
  sh.getRange(row, COL.orderId).setValue(orderId);
  // Build lineData từ payload nếu cần, hoặc đọc lại toàn row
  const rowData = sh.getRange(row, 1, 1, Object.keys(COL).length).getValues()[0];
  backupData(rowData, false);
  
  // CẬP NHẬT TẤT CẢ FILTER SHEETS SAU KHI THÊM DỮ LIỆU MỚI
  updateAllFilterSheets();
  
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', orderId }))
    .setMimeType(ContentService.MimeType.JSON);
}

function readForm(){ 
  const values = SH(SH_FORM).getRange('B1:B9').getValues().flat();
  // Convert date values to Date objects if they exist
  values[4] = values[4] ? new Date(values[4]) : null; // buy date
  values[5] = values[5] ? new Date(values[5]) : null; // exp date
  return values;
}

function writeForm(a){ 
  SH(SH_FORM).getRange('B1:B9').setValues(a.map(v=>[v])); 
}

function clearForm(){
  const f = SH(SH_FORM);
  f.getRange('B1:B9').clearContent();
  f.getRange('E2:F2').clearContent();
  f.getRange('Z1').clearContent();     // hàng cache cho update/xoá
}

// Tìm hàng trống đầu tiên - CẢI TIẾN
function findEmptyRow(){
  const sh = SH(SH_DB);
  const lastRow = sh.getLastRow();
  
  // Nếu sheet còn trống (chỉ có header)
  if(lastRow <= 1) return 2;
  
  // Lấy tất cả dữ liệu cột C (TÊN KHÁCH) để tìm hàng trống
  const range = sh.getRange(2, COL.name, lastRow - 1, 1);
  const values = range.getValues();
  
  // Tìm hàng trống đầu tiên
  for(let i = 0; i < values.length; i++){
    if(!values[i][0] || values[i][0] === ''){
      return i + 2; // +2 vì bắt đầu từ hàng 2 và index từ 0
    }
  }
  
  // Nếu không có hàng trống, thêm vào cuối
  return lastRow + 1;
}

// Tạo mã đơn hàng duy nhất - MỚI
function generateOrderId(){
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMdd');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HHmmss');
  const randomStr = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  
  return `DH${dateStr}${timeStr}${randomStr}`;
}

// Kiểm tra mã đơn hàng có trùng không - MỚI
function isOrderIdExists(orderId){
  const sh = SH(SH_DB);
  const lastRow = sh.getLastRow();
  
  if(lastRow <= 1) return false;
  
  const orderIds = sh.getRange(2, COL.orderId, lastRow - 1, 1).getValues().flat();
  return orderIds.includes(orderId);
}

// Tạo mã đơn hàng duy nhất không trùng - MỚI
function createUniqueOrderId(){
  let orderId;
  let attempts = 0;
  
  do {
    orderId = generateOrderId();
    attempts++;
    if(attempts > 100) {
      // Nếu thử quá 100 lần, thêm timestamp để đảm bảo unique
      orderId = `DH${Date.now()}`;
      break;
    }
  } while(isOrderIdExists(orderId));
  
  return orderId;
}

/************ 2. THÊM - CẢI TIẾN ************/
function addCustomer(){
  const ui = SpreadsheetApp.getUi();
  
  try {
    const [name, contact, price, cost, buy, exp, pack, paid, note] = readForm();
    
    // Kiểm tra dữ liệu bắt buộc
    if(!name){ 
      ui.alert('⚠️ Lỗi', 'Vui lòng nhập TÊN KHÁCH HÀNG!', ui.ButtonSet.OK); 
      return; 
    }
    if(!buy){ 
      ui.alert('⚠️ Lỗi', 'Vui lòng nhập NGÀY MUA!', ui.ButtonSet.OK); 
      return; 
    }

    const row = findEmptyRow();
    const sh = SH(SH_DB);
    const stt = row - 1; // STT = dòng - 1
    const orderId = createUniqueOrderId(); // Tạo mã đơn hàng duy nhất
    
    // Chuẩn bị dữ liệu
    const line = [
      stt,                            // A - STT
      orderId,                        // B - MÃ ĐƠN HÀNG
      name,                           // C - TÊN KHÁCH
      contact || '',                  // D - HÌNH THỨC LIÊN HỆ
      pack || '',                     // E - LOẠI TÀI KHOẢN
      price || 0,                     // F - GIÁ
      cost || 0,                      // G - VỐN
      paid ? 'đã thanh toán' : 'chưa thanh toán',  // H - Trạng thái thanh toán
      buy,                           // I - NGÀY MUA
      exp || ''                      // J - NGÀY HẾT HẠN
    ];
    
    // Thêm cột DaysLeft và Ghi chú
    if(WITH_DAYS_LEFT){ 
      // Thêm công thức tính Days Left cho cột K
      line.push(''); // Sẽ thêm công thức sau
      line.push(note || '');         // L - GHI CHÚ
    } else { 
      line.push(note || '');         // K - GHI CHÚ
    }

    // Ghi vào sheet chính
    sh.getRange(row, 1, 1, line.length).setValues([line]);
    
    // Thêm công thức Days Left nếu có ngày hết hạn
    if(WITH_DAYS_LEFT && exp){
      const formula = `=IF(J${row}>0;J${row}-today();"")`;
      sh.getRange(row, COL.left).setFormula(formula);
    }
    
    // Backup dữ liệu
    // Tính giá trị Days Left để backup (thay vì công thức)
    let backupLine = [...line];
    if(WITH_DAYS_LEFT && exp){
      const today = new Date();
      today.setHours(0,0,0,0);
      const daysLeft = Math.ceil((exp - today) / 86400000);
      backupLine[10] = daysLeft; // Thay thế giá trị trống bằng số ngày
    }
    backupData(backupLine);
    
    // CẬP NHẬT TẤT CẢ FILTER SHEETS SAU KHI THÊM DỮ LIỆU MỚI
    updateAllFilterSheets();
    
    // Clear form và thông báo
    clearForm();
    ui.alert('✅ Thành công', `Đã thêm khách hàng "${name}" vào dòng ${row}.\nMã đơn hàng: ${orderId}\nDữ liệu đã được backup và cập nhật filter sheets.`, ui.ButtonSet.OK);
    
  } catch(error) {
    ui.alert('❌ Lỗi', 'Đã xảy ra lỗi: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/************ 3. TÌM KIẾM - CẢI TIẾN ************/
function searchCustomer(){
  const ui = SpreadsheetApp.getUi();
  const [nameIn, contactIn] = readForm();
  
  if(!nameIn && !contactIn){ 
    ui.alert('📌 Hướng dẫn', 'Nhập TÊN hoặc LIÊN HỆ để tìm kiếm khách hàng.', ui.ButtonSet.OK); 
    return; 
  }

  const norm = s => s.toString().toLowerCase()
                    .normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();
  const nKey = norm(nameIn || '');
  const cKey = norm(contactIn || '');
  
  const sh = SH(SH_DB);
  const data = sh.getDataRange().getValues();
  let found = false;
  let foundCount = 0;
  let foundRows = [];

  for(let r = 1; r < data.length; r++){
    const row = data[r];
    if( (!nKey || norm(row[COL.name-1]).includes(nKey)) &&
        (!cKey || norm(row[COL.contact-1]).includes(cKey)) ){
      
      foundCount++;
      foundRows.push({
        row: r+1, 
        name: row[COL.name-1], 
        contact: row[COL.contact-1], 
        orderId: row[COL.orderId-1]
      });
      
      if(foundCount === 1){
        // Load dữ liệu của kết quả đầu tiên
        const back = [
          row[COL.name-1],    // B1 - Tên
          row[COL.contact-1], // B2 - Liên hệ
          row[COL.price-1],   // B3 - Giá
          row[COL.cost-1],    // B4 - Vốn
          row[COL.buy-1],     // B5 - Ngày mua
          row[COL.exp-1],     // B6 - Ngày hết hạn
          row[COL.pack-1],    // B7 - Gói
          row[COL.paid-1] === 'đã thanh toán', // B8 - Đã thanh toán (boolean)
          row[COL.note-1]     // B9 - Ghi chú
        ];
        writeForm(back);
        SH(SH_FORM).getRange('Z1').setValue(r+1);
        found = true;
      }
    }
  }
  
  if(!found){
    ui.alert('❌ Không tìm thấy', 'Không tìm thấy khách hàng phù hợp với từ khóa tìm kiếm.', ui.ButtonSet.OK);
  } else if(foundCount === 1){
    ui.alert('✅ Tìm thấy', `Đã tìm thấy 1 khách hàng.\nMã đơn hàng: ${foundRows[0].orderId}\nBạn có thể Cập nhật hoặc Xóa thông tin.`, ui.ButtonSet.OK);
  } else {
    let message = `Tìm thấy ${foundCount} khách hàng:\n\n`;
    foundRows.forEach((item, idx) => {
      if(idx < 5){ // Chỉ hiển thị 5 kết quả đầu
        message += `${idx+1}. ${item.name} - ${item.contact} (${item.orderId})\n`;
      }
    });
    if(foundCount > 5) message += `\n... và ${foundCount - 5} kết quả khác.`;
    message += '\n\nĐã load thông tin khách hàng đầu tiên.';
    ui.alert('📋 Kết quả tìm kiếm', message, ui.ButtonSet.OK);
  }
}

/************ 4. CẬP NHẬT - CẢI TIẾN ************/
function updateCustomer(){
  const ui = SpreadsheetApp.getUi();
  const row = Number(SH(SH_FORM).getRange('Z1').getValue());
  
  if(!row){ 
    ui.alert('⚠️ Lỗi', 'Vui lòng TÌM KIẾM khách hàng trước khi cập nhật!', ui.ButtonSet.OK); 
    return; 
  }

  const sh = SH(SH_DB);
  if(row > sh.getLastRow()){ 
    ui.alert('⚠️ Lỗi', 'Dòng này đã bị xóa hoặc không tồn tại!', ui.ButtonSet.OK); 
    return; 
  }

  try {
    const [name, contact, price, cost, buy, exp, pack, paid, note] = readForm();
    
    // Lấy STT và mã đơn hàng cũ (không thay đổi)
    const oldSTT = sh.getRange(row, COL.stt).getValue();
    const oldOrderId = sh.getRange(row, COL.orderId).getValue();
    
    const line = [
      oldSTT,                         // Giữ nguyên STT
      oldOrderId,                     // Giữ nguyên mã đơn hàng
      name,
      contact || '',
      pack || '',
      price || 0,
      cost || 0,
      paid ? 'đã thanh toán' : 'chưa thanh toán',
      buy,
      exp || ''
    ];
    
    if(WITH_DAYS_LEFT){ 
      line.push(''); // Sẽ thêm công thức sau
      line.push(note || '');
    } else { 
      line.push(note || '');
    }

    sh.getRange(row, 1, 1, line.length).setValues([line]);
    
    // Thêm công thức Days Left nếu có ngày hết hạn
    if(WITH_DAYS_LEFT && exp){
      const formula = `=IF(J${row}>0;J${row}-today();"")`;
      sh.getRange(row, COL.left).setFormula(formula);
    }
    
    // Backup dữ liệu cập nhật
    // Tính giá trị Days Left để backup
    let backupLine = [...line];
    if(WITH_DAYS_LEFT && exp){
      const today = new Date();
      today.setHours(0,0,0,0);
      const daysLeft = Math.ceil((exp - today) / 86400000);
      backupLine[10] = daysLeft;
    }
    backupData(backupLine, true); // true = update mode
    
    // CẬP NHẬT TẤT CẢ FILTER SHEETS SAU KHI CẬP NHẬT DỮ LIỆU
    updateAllFilterSheets();
    
    ui.alert('✏️ Cập nhật thành công', `Đã cập nhật thông tin khách hàng "${name}".\nMã đơn hàng: ${oldOrderId}\nFilter sheets đã được cập nhật.`, ui.ButtonSet.OK);
    
  } catch(error) {
    ui.alert('❌ Lỗi', 'Đã xảy ra lỗi khi cập nhật: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/************ 5. XOÁ - CẢI TIẾN ************/
function deleteCustomer(){
  const ui = SpreadsheetApp.getUi();
  const row = Number(SH(SH_FORM).getRange('Z1').getValue());
  
  if(!row){ 
    ui.alert('⚠️ Lỗi', 'Vui lòng TÌM KIẾM khách hàng trước khi xóa!', ui.ButtonSet.OK); 
    return; 
  }
  
  const sh = SH(SH_DB);
  const customerName = sh.getRange(row, COL.name).getValue();
  const orderId = sh.getRange(row, COL.orderId).getValue();
  
  const result = ui.alert(
    '⚠️ Xác nhận xóa', 
    `Bạn có chắc muốn xóa khách hàng "${customerName}" (${orderId})?\n\nLưu ý: Thao tác này không thể hoàn tác!`, 
    ui.ButtonSet.YES_NO
  );
  
  if(result == ui.Button.NO) return;
  
  try {
    // Xóa chỉ xóa nội dung, không xóa hàng để giữ cấu trúc
    sh.getRange(row, 1, 1, WITH_DAYS_LEFT ? 12 : 11).clearContent();
    
    // CẬP NHẬT TẤT CẢ FILTER SHEETS SAU KHI XÓA DỮ LIỆU
    updateAllFilterSheets();
    
    clearForm();
    ui.alert('🗑️ Xóa thành công', `Đã xóa khách hàng "${customerName}" (${orderId}).\nFilter sheets đã được cập nhật.`, ui.ButtonSet.OK);
  } catch(error) {
    ui.alert('❌ Lỗi', 'Đã xảy ra lỗi khi xóa: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/************ 6. DỌN DẸP FORM ************/
function clearFormButton(){ 
  clearForm(); 
  SpreadsheetApp.getUi().alert('🧹 Đã dọn dẹp', 'Form nhập liệu đã được xóa sạch.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/************ 7. DOANH THU - CẢI TIẾN ************/
function calcRevenue(){
  const today = new Date(); 
  today.setHours(0,0,0,0);
  const ym = [today.getFullYear(), today.getMonth()];
  
  let dayRevenue = 0;
  let monthRevenue = 0;
  let dayCount = 0;
  let monthCount = 0;
  
  const sh = SH(SH_DB);
  const data = sh.getDataRange().getValues();
  
  data.slice(1).forEach(r => {
    if(!r[COL.name-1]) return; // Bỏ qua hàng trống
    
    const price = Number(r[COL.price-1]) || 0;
    const buyDate = r[COL.buy-1];
    
    if(buyDate){
      const d = new Date(buyDate); 
      d.setHours(0,0,0,0);
      
      // Doanh thu ngày
      if(d.getTime() == today.getTime()){
        dayRevenue += price;
        dayCount++;
      }
      
      // Doanh thu tháng
      if(d.getFullYear() == ym[0] && d.getMonth() == ym[1]){
        monthRevenue += price;
        monthCount++;
      }
    }
  });
  
  const f = SH(SH_FORM);
  f.getRange('E2').setValue(dayRevenue);
  f.getRange('F2').setValue(monthRevenue);
  
  const monthName = today.toLocaleDateString('vi-VN', { month: 'long', year: 'numeric' });
  SpreadsheetApp.getUi().alert(
    '📈 Báo cáo doanh thu',
    `Doanh thu hôm nay: ${dayRevenue.toLocaleString('vi-VN')}đ (${dayCount} đơn)\n` +
    `Doanh thu ${monthName}: ${monthRevenue.toLocaleString('vi-VN')}đ (${monthCount} đơn)`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/************ 8. BACKUP DATA - CẢI TIẾN ************/
function backupData(rowData, isUpdate = false) {
  try {
    // Kiểm tra cấu hình backup
    if(!BACKUP_SHEET_ID || BACKUP_SHEET_ID === 'YOUR_BACKUP_SHEET_ID_HERE'){
      console.log('Chưa cấu hình Backup Sheet ID');
      return;
    }
    
    // Mở file backup
    const backupFile = SpreadsheetApp.openById(BACKUP_SHEET_ID);
    let backupSheet = backupFile.getSheetByName(BACKUP_SHEET_NAME);
    
    // Nếu chưa có sheet backup, tạo mới
    if(!backupSheet){
      backupSheet = backupFile.insertSheet(BACKUP_SHEET_NAME);
      // Tạo header cho sheet backup
      const headers = [
        'MÃ ĐƠN HÀNG', 'TÊN KHÁCH', 'HÌNH THỨC LIÊN HỆ', 'LOẠI TÀI KHOẢN', 
        'GIÁ', 'VỐN', 'Trạng thái thanh toán', 'NGÀY MUA', 'NGÀY HẾT HẠN'
      ];
      if(WITH_DAYS_LEFT) headers.push('Days Left');
      headers.push('GHI CHÚ', 'Thời gian backup', 'Loại thao tác');
      
      backupSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      backupSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    // Thêm timestamp và loại thao tác
    const timestamp = new Date();
    const dataWithTimestamp = [...rowData, timestamp, isUpdate ? 'UPDATE' : 'NEW'];
    
    // Lấy index của cột Order ID (cột B) trong sheet backup
    const ORDER_ID_COL = 2;

    // Xác định số dòng hiện có (có header ở dòng 1)
    const backupLastRow = backupSheet.getLastRow();

    // Lấy mảng giá trị cột Order ID từ dòng 2 đến backupLastRow
    const idValues = backupSheet
      .getRange(2, ORDER_ID_COL, backupLastRow - 1, 1)
      .getValues()
      .flat();

    // Tìm hàng trống đầu tiên (nếu cell Order ID là chuỗi rỗng hoặc null)
    let targetRow = idValues.findIndex(v => !v) + 2;

    // Nếu không tìm thấy (tức findIndex = -1), thì thêm xuống cuối
    if (targetRow < 2) {
      targetRow = backupLastRow + 1;
    }

    // Ghi dữ liệu vào sheet backup tại hàng tìm được
    // Bỏ phần tử đầu (STT) trong mảng, rồi ghi ra bắt đầu từ cột 2 (B)
    const valuesWithoutStt = dataWithTimestamp.slice(1);
    backupSheet
      .getRange(targetRow, 2, 1, valuesWithoutStt.length)
      .setValues([valuesWithoutStt]);

    // Format ngày tháng cho dễ đọc
    const dateColumns = [9, 10]; // Cột NGÀY MUA và NGÀY HẾT HẠN
    if(WITH_DAYS_LEFT) dateColumns.push(13); // Cột timestamp
    else dateColumns.push(12);
    
    dateColumns.forEach(col => {
      if(dataWithTimestamp[col-1]){
        backupSheet.getRange(targetRow, col).setNumberFormat('dd/mm/yyyy');
      }
    });
    
    console.log(`Backup thành công: dòng ${targetRow} trong sheet ${BACKUP_SHEET_NAME}`);
    
  } catch(error) {
    console.error('Lỗi backup:', error);
    // Không thông báo lỗi backup cho user để không làm gián đoạn workflow
  }
}

/************ 9. TRIGGER ô D2 - CẢI TIẾN ************/
function onEdit(e){
  if(!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  
  // Xử lý form điều khiển
  if(sheet.getName() === SH_FORM && e.range.getA1Notation() === 'D2') {
    const value = e.value;
    if(!value) return;
    
    // Clear cell ngay lập tức
    e.range.clearContent();
    
    // Thực hiện hành động dựa trên dropdown
    switch(value){
      case 'Thêm':      addCustomer();      break;
      case 'Tìm kiếm':  searchCustomer();   break;
      case 'Cập nhật':  updateCustomer();   break;
      case 'Xóa':       deleteCustomer();   break;
      case 'Dọn dẹp':   clearFormButton();  break;
      case 'Doanh thu': calcRevenue();      break;
    }
    return;
  }
  
  // XỬ LÝ CHỈNH SỬA TRÊN FILTER SHEETS - MỚI
  handleFilterSheetEdit(e);
}

/************ 10. CẢNH BÁO HẾT HẠN - CẢI TIẾN ************/
function sendExpiryAlert(){
  const sh = SH(SH_DB);
  const data = sh.getDataRange().getValues();
  const result = [];
  const today = new Date();
  today.setHours(0,0,0,0);

  data.slice(1).forEach((r, idx) => {
    if(!r[COL.name-1]) return; // Bỏ qua hàng trống
    
    let days;
    if (WITH_DAYS_LEFT){
      days = r[COL.left-1];
    } else {
      const exp = r[COL.exp-1];
      if(!exp) return;
      days = Math.ceil((new Date(exp) - today) / 86400000);
    }
    
    if(days !== '' && days <= ALERT_THRESHOLD){
      result.push({
        row: idx + 2,
        orderId: r[COL.orderId-1],
        name: r[COL.name-1],
        contact: r[COL.contact-1],
        pack: r[COL.pack-1],
        exp: r[COL.exp-1] ? Utilities.formatDate(new Date(r[COL.exp-1]), Session.getScriptTimeZone(),'dd/MM/yyyy') : '',
        days: days,
        status: days < 0 ? 'Đã hết hạn' : (days === 0 ? 'Hết hạn hôm nay' : `Còn ${days} ngày`)
      });
    }
  });

  if(result.length === 0) return;

  // Sắp xếp theo số ngày còn lại
  result.sort((a, b) => a.days - b.days);

  // Tạo bảng HTML đẹp hơn
  const rows = result.map(o => `
    <tr>
      <td style="text-align:center">${o.row}</td>
      <td style="text-align:center"><b>${o.orderId}</b></td>
      <td><b>${o.name}</b></td>
      <td>${o.contact || '-'}</td>
      <td>${o.pack || '-'}</td>
      <td style="text-align:center">${o.exp}</td>
      <td style="text-align:center; color: ${o.days < 0 ? 'red' : (o.days === 0 ? 'orange' : '#333')}">
        <b>${o.status}</b>
      </td>
    </tr>
  `).join('');

  const html = `
    <div style="font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto;">
      <h2 style="color: #e74c3c;">🔔 Cảnh báo đơn hàng sắp/đã hết hạn</h2>
      <p>Có <b>${result.length}</b> đơn hàng cần chú ý (≤ ${ALERT_THRESHOLD} ngày):</p>
      
      <table border="0" cellpadding="8" cellspacing="0" style="width:100%; border-collapse:collapse; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
        <thead>
          <tr style="background-color: #3498db; color: white;">
            <th style="border: 1px solid #2980b9;">Dòng</th>
            <th style="border: 1px solid #2980b9;">Mã đơn hàng</th>
            <th style="border: 1px solid #2980b9;">Khách hàng</th>
            <th style="border: 1px solid #2980b9;">Liên hệ</th>
            <th style="border: 1px solid #2980b9;">Gói dịch vụ</th>
            <th style="border: 1px solid #2980b9;">Ngày hết hạn</th>
            <th style="border: 1px solid #2980b9;">Trạng thái</th>
          </tr>
        </thead>
        <tbody style="background-color: #fff;">
          ${rows}
        </tbody>
      </table>
      
      <hr style="margin: 20px 0; border: none; border-top: 1px solid #ecf0f1;">
      <p style="font-size: 12px; color: #7f8c8d;">
        📅 Báo cáo tự động - ${Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}<br>
        📧 Email tự động từ hệ thống quản lý bán hàng
      </p>
    </div>
  `;

  GmailApp.sendEmail(
    ADMIN_EMAIL,
    `🔔 [CẢNH BÁO] ${result.length} đơn hàng sắp/đã hết hạn`,
    `Có ${result.length} đơn hàng cần xử lý. Vui lòng xem chi tiết trong email.`,
    {
      htmlBody: html,
      name: 'Hệ thống Quản lý Bán hàng'
    }
  );
}

/************ 11. MENU TÙY CHỈNH - CẢI TIẾN ************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ Công cụ')
    .addItem('📊 Báo cáo doanh thu', 'calcRevenue')
    .addItem('📧 Gửi cảnh báo hết hạn', 'sendExpiryAlert')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔍 Bộ lọc dữ liệu')
      .addItem('➕ Thêm bộ lọc mới', 'addNewFilter')
      .addItem('🔄 Cập nhật tất cả filter', 'updateAllFilterSheets')
      .addItem('📋 Quản lý filter', 'manageFilters')
      .addItem('🗑️ Xóa filter', 'deleteFilter'))
    .addSeparator()
    .addItem('⚙️ Cài đặt Trigger tự động', 'setupTriggers')
    .addItem('🔧 Xóa tất cả Trigger', 'deleteTriggers')
    .addItem('🔄 Kiểm tra Sheet Backup', 'checkBackupSheet')
    .addItem('🔢 Cập nhật công thức Days Left', 'updateDaysLeftFormulas')
    .addItem('🆔 Tạo lại header cho sheet mới', 'setupNewSheetHeaders')
    .addToUi();
}

/************ 12. QUẢN LÝ TRIGGERS - MỚI ************/
function setupTriggers() {
  const ui = SpreadsheetApp.getUi();
  
  // Xóa triggers cũ trước
  deleteTriggers();
  
  // Tạo trigger cho cảnh báo hàng ngày
  ScriptApp.newTrigger('sendExpiryAlert')
    .timeBased()
    .everyDays(1)
    .atHour(9) // 9h sáng
    .create();
    
  ui.alert(
    '✅ Đã cài đặt', 
    'Hệ thống sẽ tự động gửi email cảnh báo lúc 9h sáng mỗi ngày.', 
    ui.ButtonSet.OK
  );
}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/************ 13. KIỂM TRA BACKUP SHEET - MỚI ************/
function checkBackupSheet() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    if(!BACKUP_SHEET_ID || BACKUP_SHEET_ID === 'YOUR_BACKUP_SHEET_ID_HERE'){
      ui.alert('⚠️ Chưa cấu hình', 'Vui lòng cấu hình BACKUP_SHEET_ID trong code!', ui.ButtonSet.OK);
      return;
    }
    
    const backupFile = SpreadsheetApp.openById(BACKUP_SHEET_ID);
    let backupSheet = backupFile.getSheetByName(BACKUP_SHEET_NAME);
    
    if(!backupSheet){
      // Tạo sheet backup mới
      backupSheet = backupFile.insertSheet(BACKUP_SHEET_NAME);
      const headers = [
        'STT', 'MÃ ĐƠN HÀNG', 'TÊN KHÁCH', 'HÌNH THỨC LIÊN HỆ', 'LOẠI TÀI KHOẢN', 
        'GIÁ', 'VỐN', 'Trạng thái thanh toán', 'NGÀY MUA', 'NGÀY HẾT HẠN'
      ];
      if(WITH_DAYS_LEFT) headers.push('Days Left');
      headers.push('GHI CHÚ', 'Thời gian backup', 'Loại thao tác');
      
      backupSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      backupSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      
      ui.alert('✅ Thành công', 'Đã tạo sheet backup mới!', ui.ButtonSet.OK);
    } else {
      // Lấy số dòng cuối cùng
      const backupLastRow = backupSheet.getLastRow();
    
      // Lấy mảng Order ID (cột B) từ dòng 2 đến hết
      const idValues = backupSheet
        .getRange(2, 2, backupLastRow - 1, 1)
        .getValues()
        .flat();
    
      // Đếm những ô thực sự có giá trị Order ID
      const recordCount = idValues.filter(v => v !== '' && v !== null).length;
    
      ui.alert(
        '✅ Sheet backup hoạt động tốt',
        `Sheet backup đang có ${recordCount} bản ghi.`,
        ui.ButtonSet.OK
      );
    }
    
  } catch(error) {
    ui.alert('❌ Lỗi', 'Không thể truy cập sheet backup: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/************ 14. CẬP NHẬT CÔNG THỨC DAYS LEFT - MỚI ************/
function updateDaysLeftFormulas() {
  if(!WITH_DAYS_LEFT) {
    SpreadsheetApp.getUi().alert('ℹ️ Thông báo', 'Tính năng Days Left đã được tắt trong cấu hình.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const sh = SH(SH_DB);
  const lastRow = sh.getLastRow();
  
  if(lastRow <= 1) {
    SpreadsheetApp.getUi().alert('ℹ️ Thông báo', 'Không có dữ liệu để cập nhật.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Cập nhật công thức cho toàn bộ cột K (Days Left)
  for(let row = 2; row <= lastRow; row++) {
    const hasExpDate = sh.getRange(row, COL.exp).getValue();
    if(hasExpDate) {
      const formula = `=IF(J${row}>0;J${row}-today();"")`;
      sh.getRange(row, COL.left).setFormula(formula);
    } else {
      sh.getRange(row, COL.left).clearContent();
    }
  }
  
  SpreadsheetApp.getUi().alert('✅ Thành công', `Đã cập nhật công thức Days Left cho ${lastRow - 1} dòng.`, ui.ButtonSet.OK);
}

/************ 15. TẠO HEADER MỚI CHO SHEET - MỚI ************/
function setupNewSheetHeaders() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '⚠️ Xác nhận', 
    'Bạn có muốn tạo lại header cho sheet "Đơn hàng" với cột MÃ ĐƠN HÀNG mới không?\n\nLưu ý: Thao tác này sẽ ghi đè header hiện tại!', 
    ui.ButtonSet.YES_NO
  );
  
  if(result == ui.Button.NO) return;
  
  try {
    const sh = SH(SH_DB);
    
    // Tạo header mới
    const headers = [
      'STT',                    // A
      'MÃ ĐƠN HÀNG',           // B (MỚI)
      'TÊN KHÁCH',             // C
      'HÌNH THỨC LIÊN HỆ',     // D
      'LOẠI TÀI KHOẢN',        // E
      'GIÁ',                   // F
      'VỐN',                   // G
      'Trạng thái thanh toán', // H
      'NGÀY MUA',              // I
      'NGÀY HẾT HẠN'           // J
    ];
    
    if(WITH_DAYS_LEFT) {
      headers.push('Days Left'); // K
      headers.push('GHI CHÚ');   // L
    } else {
      headers.push('GHI CHÚ');   // K
    }
    
    // Ghi header vào dòng 1
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.getRange(1, 1, 1, headers.length).setBackground('#4285f4');
    sh.getRange(1, 1, 1, headers.length).setFontColor('white');
    
    ui.alert('✅ Thành công', `Đã tạo header mới với ${headers.length} cột.\n\nLưu ý: Bạn cần thêm cột "24/07/2025" vào ô K1 để công thức Days Left hoạt động.`, ui.ButtonSet.OK);
    
  } catch(error) {
    ui.alert('❌ Lỗi', 'Đã xảy ra lỗi khi tạo header: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/************ 16. QUẢN LÝ BỘ LỌC DỮ LIỆU - MỚI ************/

// Tạo sheet cấu hình filter nếu chưa có
function createFilterConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(FILTER_CONFIG_SHEET);
  
  if (!configSheet) {
    configSheet = ss.insertSheet(FILTER_CONFIG_SHEET);
    
    // Tạo header cho sheet cấu hình
    const headers = ['Tên Sheet', 'Trường lọc', 'Giá trị lọc', 'Ngày tạo', 'Trạng thái'];
    configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    configSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    configSheet.getRange(1, 1, 1, headers.length).setBackground('#34495e');
    configSheet.getRange(1, 1, 1, headers.length).setFontColor('white');
    
    // Ẩn sheet cấu hình
    configSheet.hideSheet();
  }
  
  return configSheet;
}

// Thêm bộ lọc mới
function addNewFilter() {
  const ui = SpreadsheetApp.getUi();
  
  // Nhập tên sheet
  const sheetNameResult = ui.prompt(
    '📝 Tên Sheet Filter',
    'Nhập tên cho sheet filter mới (ví dụ: ChatGPT, Netflix, Spotify):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sheetNameResult.getSelectedButton() != ui.Button.OK || !sheetNameResult.getResponseText().trim()) {
    return;
  }
  
  const sheetName = sheetNameResult.getResponseText().trim();
  
  // Kiểm tra tên sheet đã tồn tại
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(sheetName)) {
    ui.alert('⚠️ Lỗi', `Sheet "${sheetName}" đã tồn tại!`, ui.ButtonSet.OK);
    return;
  }
  
  // Lấy danh sách loại tài khoản có sẵn
  const mainSheet = SH(SH_DB);
  const data = mainSheet.getDataRange().getValues();
  const accountTypes = new Set();
  
  data.slice(1).forEach(row => {
    if (row[COL.pack-1] && row[COL.pack-1].toString().trim()) {
      accountTypes.add(row[COL.pack-1].toString().trim());
    }
  });
  
  const accountTypesArray = Array.from(accountTypes).sort();
  
  if (accountTypesArray.length === 0) {
    ui.alert('⚠️ Lỗi', 'Không tìm thấy loại tài khoản nào trong dữ liệu!', ui.ButtonSet.OK);
    return;
  }
  
  // Hiển thị danh sách loại tài khoản
  let accountTypesList = 'Các loại tài khoản có sẵn:\n\n';
  accountTypesArray.forEach((type, index) => {
    accountTypesList += `${index + 1}. ${type}\n`;
  });
  
  // Nhập loại tài khoản cần lọc
  const filterValueResult = ui.prompt(
    '🔍 Giá trị lọc',
    accountTypesList + '\nNhập chính xác loại tài khoản cần lọc:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (filterValueResult.getSelectedButton() != ui.Button.OK || !filterValueResult.getResponseText().trim()) {
    return;
  }
  
  const filterValue = filterValueResult.getResponseText().trim();
  
  // Kiểm tra giá trị lọc có tồn tại không
  if (!accountTypesArray.includes(filterValue)) {
    ui.alert('⚠️ Lỗi', `Loại tài khoản "${filterValue}" không tồn tại!\n\nVui lòng nhập chính xác theo danh sách.`, ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Tạo sheet filter mới
    const filterSheet = ss.insertSheet(sheetName);
    
    // Tạo header cho sheet filter (giống sheet chính)
    const headers = [
      'STT', 'MÃ ĐƠN HÀNG', 'TÊN KHÁCH', 'HÌNH THỨC LIÊN HỆ', 'LOẠI TÀI KHOẢN', 
      'GIÁ', 'VỐN', 'Trạng thái thanh toán', 'NGÀY MUA', 'NGÀY HẾT HẠN'
    ];
    
    if(WITH_DAYS_LEFT) {
      headers.push('Days Left');
      headers.push('GHI CHÚ');
    } else {
      headers.push('GHI CHÚ');
    }
    
    filterSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    filterSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    filterSheet.getRange(1, 1, 1, headers.length).setBackground('#2ecc71');
    filterSheet.getRange(1, 1, 1, headers.length).setFontColor('white');
    
    // Lưu cấu hình filter
    const configSheet = createFilterConfigSheet();
    const lastRow = configSheet.getLastRow();
    const newRow = lastRow + 1;
    
    configSheet.getRange(newRow, 1, 1, 5).setValues([[
      sheetName,
      'LOẠI TÀI KHOẢN',
      filterValue,
      new Date(),
      'Hoạt động'
    ]]);
    
    // Cập nhật dữ liệu cho sheet filter mới
    updateFilterSheet(sheetName, 'LOẠI TÀI KHOẢN', filterValue);
    
    ui.alert(
      '✅ Thành công', 
      `Đã tạo sheet filter "${sheetName}" cho loại tài khoản "${filterValue}".\n\nSheet sẽ tự động cập nhật khi có thay đổi dữ liệu.`, 
      ui.ButtonSet.OK
    );
    
  } catch(error) {
    ui.alert('❌ Lỗi', 'Đã xảy ra lỗi khi tạo filter: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// Cập nhật dữ liệu cho một sheet filter
function updateFilterSheet(sheetName, filterField, filterValue) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const filterSheet = ss.getSheetByName(sheetName);
    if (!filterSheet) return;
    
    const mainSheet = SH(SH_DB);
    const data = mainSheet.getDataRange().getValues();
    
    // Lọc dữ liệu theo điều kiện
    const filteredData = [];
    let sttCounter = 1;
    
    data.slice(1).forEach(row => {
      if (!row[COL.name-1]) return; // Bỏ qua hàng trống
      
      let shouldInclude = false;
      
      switch(filterField) {
        case 'LOẠI TÀI KHOẢN':
          shouldInclude = row[COL.pack-1] && row[COL.pack-1].toString().trim() === filterValue;
          break;
        case 'HÌNH THỨC LIÊN HỆ':
          shouldInclude = row[COL.contact-1] && row[COL.contact-1].toString().trim() === filterValue;
          break;
        case 'Trạng thái thanh toán':
          shouldInclude = row[COL.paid-1] && row[COL.paid-1].toString().trim() === filterValue;
          break;
      }
      
      if (shouldInclude) {
        // Tạo dòng dữ liệu mới với STT được đánh số lại
        const newRow = [...row];
        newRow[COL.stt-1] = sttCounter++;
        filteredData.push(newRow);
      }
    });
    
    // Xóa dữ liệu cũ (giữ lại header)
    const lastRow = filterSheet.getLastRow();
    if (lastRow > 1) {
      filterSheet.getRange(2, 1, lastRow - 1, WITH_DAYS_LEFT ? 12 : 11).clearContent();
    }
    
    // Ghi dữ liệu mới
    if (filteredData.length > 0) {
      filterSheet.getRange(2, 1, filteredData.length, WITH_DAYS_LEFT ? 12 : 11).setValues(filteredData);
      
      // Thêm công thức Days Left nếu cần
      if (WITH_DAYS_LEFT) {
        for (let i = 0; i < filteredData.length; i++) {
          const row = i + 2;
          const expDate = filteredData[i][COL.exp-1];
          if (expDate) {
            const formula = `=IF(J${row}>0;J${row}-today();"")`;
            filterSheet.getRange(row, COL.left).setFormula(formula);
          }
        }
      }
    }
    
    console.log(`Đã cập nhật ${filteredData.length} bản ghi cho sheet "${sheetName}"`);
    
  } catch(error) {
    console.error(`Lỗi cập nhật filter sheet "${sheetName}":`, error);
  }
}

// Cập nhật tất cả filter sheets
function updateAllFilterSheets() {
  try {
    const configSheet = createFilterConfigSheet();
    const lastRow = configSheet.getLastRow();
    
    if (lastRow <= 1) return; // Không có filter nào
    
    const configs = configSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    configs.forEach(config => {
      const [sheetName, filterField, filterValue, , status] = config;
      
      if (status === 'Hoạt động') {
        updateFilterSheet(sheetName, filterField, filterValue);
      }
    });
    
    console.log(`Đã cập nhật ${configs.length} filter sheets`);
    
  } catch(error) {
    console.error('Lỗi cập nhật tất cả filter sheets:', error);
  }
}

// Quản lý filters
function manageFilters() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const configSheet = createFilterConfigSheet();
    const lastRow = configSheet.getLastRow();
    
    if (lastRow <= 1) {
      ui.alert('ℹ️ Thông báo', 'Chưa có filter nào được tạo.', ui.ButtonSet.OK);
      return;
    }
    
    const configs = configSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    let message = 'Danh sách Filter hiện có:\n\n';
    configs.forEach((config, index) => {
      const [sheetName, filterField, filterValue, createDate, status] = config;
      message += `${index + 1}. Sheet: "${sheetName}"\n`;
      message += `   Lọc: ${filterField} = "${filterValue}"\n`;
      message += `   Trạng thái: ${status}\n`;
      message += `   Tạo: ${Utilities.formatDate(createDate, Session.getScriptTimeZone(), 'dd/MM/yyyy')}\n\n`;
    });
    
    ui.alert('📋 Quản lý Filter', message, ui.ButtonSet.OK);
    
  } catch(error) {
    ui.alert('❌ Lỗi', 'Đã xảy ra lỗi khi quản lý filter: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// Xóa filter
function deleteFilter() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const configSheet = createFilterConfigSheet();
    const lastRow = configSheet.getLastRow();
    
    if (lastRow <= 1) {
      ui.alert('ℹ️ Thông báo', 'Chưa có filter nào để xóa.', ui.ButtonSet.OK);
      return;
    }
    
    const configs = configSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    let filterList = 'Chọn filter cần xóa:\n\n';
    configs.forEach((config, index) => {
      const [sheetName, filterField, filterValue] = config;
      filterList += `${index + 1}. ${sheetName} (${filterField} = "${filterValue}")\n`;
    });
    
    const result = ui.prompt(
      '🗑️ Xóa Filter',
      filterList + '\nNhập số thứ tự filter cần xóa:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (result.getSelectedButton() != ui.Button.OK) return;
    
    const index = parseInt(result.getResponseText()) - 1;
    
    if (isNaN(index) || index < 0 || index >= configs.length) {
      ui.alert('⚠️ Lỗi', 'Số thứ tự không hợp lệ!', ui.ButtonSet.OK);
      return;
    }
    
    const sheetName = configs[index][0];
    
    // Xác nhận xóa
    const confirmResult = ui.alert(
      '⚠️ Xác nhận xóa',
      `Bạn có chắc muốn xóa filter "${sheetName}"?\n\nSheet sẽ bị xóa vĩnh viễn!`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResult != ui.Button.YES) return;
    
    // Xóa sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const filterSheet = ss.getSheetByName(sheetName);
    if (filterSheet) {
      ss.deleteSheet(filterSheet);
    }
    
    // Xóa config
    configSheet.deleteRow(index + 2);
    
    ui.alert('✅ Thành công', `Đã xóa filter "${sheetName}".`, ui.ButtonSet.OK);
    
  } catch(error) {
    ui.alert('❌ Lỗi', 'Đã xảy ra lỗi khi xóa filter: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// Xử lý chỉnh sửa trên filter sheets
function handleFilterSheetEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  // Bỏ qua sheet chính và form
  if (sheetName === SH_DB || sheetName === SH_FORM || sheetName === FILTER_CONFIG_SHEET) {
    return;
  }
  
  // Kiểm tra xem có phải filter sheet không
  const configSheet = SH(FILTER_CONFIG_SHEET);
  if (!configSheet) return;
  
  const lastRow = configSheet.getLastRow();
  if (lastRow <= 1) return;
  
  const configs = configSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const isFilterSheet = configs.some(config => config[0] === sheetName);
  
  if (!isFilterSheet) return;
  
  // Lấy dữ liệu từ filter sheet
  const editedRow = e.range.getRow();
  if (editedRow <= 1) return; // Không cho sửa header
  
  try {
    const rowData = sheet.getRange(editedRow, 1, 1, WITH_DAYS_LEFT ? 12 : 11).getValues()[0];
    const orderId = rowData[COL.orderId-1];
    
    if (!orderId) return;
    
    // Tìm và cập nhật dòng tương ứng trong sheet chính
    const mainSheet = SH(SH_DB);
    const mainData = mainSheet.getDataRange().getValues();
    
    for (let i = 1; i < mainData.length; i++) {
      if (mainData[i][COL.orderId-1] === orderId) {
        // Cập nhật dòng trong sheet chính
        mainSheet.getRange(i + 1, 1, 1, WITH_DAYS_LEFT ? 12 : 11).setValues([rowData]);
        
        // Thêm lại công thức Days Left nếu cần
        if (WITH_DAYS_LEFT && rowData[COL.exp-1]) {
          const formula = `=IF(J${i + 1}>0;J${i + 1}-today();"")`;
          mainSheet.getRange(i + 1, COL.left).setFormula(formula);
        }
        
        console.log(`Đã đồng bộ dữ liệu từ sheet "${sheetName}" về sheet chính`);
        break;
      }
    }
    
    // Cập nhật tất cả filter sheets khác
    setTimeout(() => {
      updateAllFilterSheets();
    }, 1000);
    
  } catch(error) {
    console.error('Lỗi xử lý chỉnh sửa filter sheet:', error);
  }
}