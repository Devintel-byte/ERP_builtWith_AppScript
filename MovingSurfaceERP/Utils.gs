function toMinutes_(hhmm) {
  const p = String(hhmm || '00:00').split(':');
  return Number(p[0] || 0) * 60 + Number(p[1] || 0);
}

function getDistanceMetres_(lat1,lon1,lat2,lon2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2-lat1), dLon = toRad(lon2-lon1);
  const a = Math.sin(dLat/2)*Math.sin(dLat/2)+Math.cos(toRad(lat1))*Math.cos(toRad(lat2))*Math.sin(dLon/2)*Math.sin(dLon/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function formatMoney_(n) {
  return Number(n || 0).toLocaleString('en-NG', {maximumFractionDigits:0});
}


function clientSafeDate_(value) {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d.getTime())) return String(value);
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Africa/Lagos', 'yyyy-MM-dd HH:mm:ss');
}

function escapeHtml_(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}