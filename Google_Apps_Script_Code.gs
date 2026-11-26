function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['訂單編號', '建立時間', '取餐日', '取餐時段', '稱呼', '品項明細', '備註', '總金額', '狀態', 'itemsJSON']);
  }
  var data = JSON.parse(e.postData.contents);
  sheet.appendRow([
    data.id,
    new Date(data.createdAt).toLocaleString('zh-TW'),
    data.day,
    data.slot,
    data.customer,
    data.itemsText,
    data.note || '',
    data.total,
    '待製作',
    JSON.stringify(data.items)
  ]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var itemsObj = {};
    try { itemsObj = JSON.parse(r[9] || '{}'); } catch (err) {}
    list.push({
      id: String(r[0]),
      day: String(r[2]),
      slot: String(r[3]),
      customer: String(r[4]),
      note: String(r[6] || ''),
      total: Number(r[7] || 0),
      status: String(r[8]) === '已取餐' ? 'done' : 'pending',
      items: itemsObj,
      createdAt: new Date(r[1]).getTime() || Date.now()
    });
  }
  return ContentService.createTextOutput(JSON.stringify({ orders: list.reverse() }))
    .setMimeType(ContentService.MimeType.JSON);
}
