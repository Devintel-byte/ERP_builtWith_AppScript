function recordActivity_(actorType, actorName, actorId, action, entity, entityId, details) {
  const sheet = ensureActivityLogSheet_();
  sheet.appendRow([
    new Date(),
    String(actorType || 'System'),
    String(actorName || ''),
    String(actorId || ''),
    String(action || ''),
    String(entity || ''),
    String(entityId || ''),
    String(details || '').slice(0,2000)
  ]);
}

function ensureActivityLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = activityLogSheetName_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const headers = activityLogHeaders_();
  const existing = sheet.getRange(1,1,1,headers.length).getValues()[0];
  if (headers.some((header,index) => String(existing[index] || '') !== header)) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getActivityLogPage_(filters) {
  filters = filters || {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(activityLogSheetName_());
  if (!sheet || sheet.getLastRow() < 2) return {rows:[],page:1,pageSize:25,total:0,totalPages:0};
  const pageSize = 25;
  const page = Math.max(1,Math.floor(Number(filters.page) || 1));
  const query = String(filters.query || '').trim().toLowerCase();
  const action = String(filters.action || '').trim().toLowerCase();
  const from = String(filters.from || '');
  const to = String(filters.to || '');
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,8).getValues();
  const rows = values.map(function(row,index) {
    const timestamp = row[0] instanceof Date ? row[0] : new Date(row[0]);
    return {
      timestamp:!isNaN(timestamp.getTime()) ? Utilities.formatDate(timestamp,Session.getScriptTimeZone() || 'Africa/Lagos','yyyy-MM-dd HH:mm:ss') : String(row[0] || ''),
      dateKey:!isNaN(timestamp.getTime()) ? Utilities.formatDate(timestamp,Session.getScriptTimeZone() || 'Africa/Lagos','yyyy-MM-dd') : '',
      actorType:String(row[1] || ''),actorName:String(row[2] || ''),actorId:String(row[3] || ''),
      action:String(row[4] || ''),entity:String(row[5] || ''),entityId:String(row[6] || ''),
      details:String(row[7] || ''),sortTime:!isNaN(timestamp.getTime()) ? timestamp.getTime() : 0,
      sheetRow:index+2
    };
  }).filter(function(row) {
    return (!from || row.dateKey >= from) &&
      (!to || row.dateKey <= to) &&
      (!action || row.action.toLowerCase() === action) &&
      (!query || [row.actorType,row.actorName,row.actorId,row.action,row.entity,row.entityId,row.details].join(' ').toLowerCase().includes(query));
  });
  const direction = filters.sort === 'oldest' ? 1 : -1;
  rows.sort((a,b) => direction*(a.sortTime-b.sortTime || a.sheetRow-b.sheetRow));
  const total = rows.length;
  const totalPages = total ? Math.ceil(total/pageSize) : 0;
  const start = (page-1)*pageSize;
  return {rows:rows.slice(start,start+pageSize),page:page,pageSize:pageSize,total:total,totalPages:totalPages};
}

function activityLogSheetName_() {
  return 'Activity Log';
}

function activityLogHeaders_() {
  return ['Timestamp','Actor Type','Actor Name','Actor / Staff ID','Action','Entity','Entity ID','Details'];
}
