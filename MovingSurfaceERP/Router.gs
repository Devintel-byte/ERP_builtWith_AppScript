const WEB_APP_ROUTES_ = {
  attendance:{view:'AttendanceView',title:'Moving Surface Attendance'},
  staff:{parameter:'staff',view:'StaffPortalView',title:'Moving Surface Staff Portal'},
  recovery:{parameter:'recover',view:'RecoveryView',title:'Staff account recovery'},
  req:{parameter:'req',view:'RequisitionView',title:'Moving Surface Requisition'},
  admin:{parameter:'admin',view:'AdminPortalView',title:'Moving Surface ERP'}
};
const WEB_APP_ROUTE_PRECEDENCE_ = ['admin','req','recovery','staff'];

function doGet(e) {
  try {
    const params = normalizeQueryParameters_(e && e.parameter ? e.parameter : {});
    if (params.approvereq) return handleEmailApprovalPage_(params);
    const route = resolveRoute_(params);
    if (route.key === 'recovery') {
      const template = HtmlService.createTemplateFromFile('Index');
      template.viewName = route.view;
      template.recoveryToken = String(params.token || '');
      template.canonicalUrl = buildCanonicalWebAppUrl_(ScriptApp.getService().getUrl(),params);
      return template.evaluate().setTitle(route.title).addMetaTag('viewport','width=device-width, initial-scale=1');
    }
    return renderPage_(route.view, route.title,params);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<h2>System Error</h2><pre>' + escapeHtml_(err.stack || err.message) + '</pre>'
    );
  }
}

function resolveRoute_(params) {
  for (let i = 0; i < WEB_APP_ROUTE_PRECEDENCE_.length; i++) {
    const route = WEB_APP_ROUTE_PRECEDENCE_[i];
    const definition = WEB_APP_ROUTES_[route];
    if (definition.parameter && String(getQueryParameterIgnoreCase_(params,definition.parameter) || '') === '1') {
      return {key:route,view:definition.view,title:definition.title};
    }
  }
  return {key:'attendance',view:WEB_APP_ROUTES_.attendance.view,title:WEB_APP_ROUTES_.attendance.title};
}

function getQueryParameterIgnoreCase_(params,name) {
  const expected = String(name || '').toLowerCase();
  const keys = Object.keys(params || {});
  for (let i=0;i<keys.length;i++) {
    if (keys[i].toLowerCase() === expected) return params[keys[i]];
  }
  return '';
}

function normalizeQueryParameters_(params) {
  const normalized = {};
  Object.keys(params || {}).forEach(function(key) {
    const normalizedKey = key.toLowerCase();
    if (!Object.prototype.hasOwnProperty.call(normalized,normalizedKey)) {
      normalized[normalizedKey] = params[key];
    }
  });
  return normalized;
}

function getWebAppRouteUrl_(route, queryParams) {
  const routeName = Object.prototype.hasOwnProperty.call(WEB_APP_ROUTES_, route) ? route : 'attendance';
  const definition = WEB_APP_ROUTES_[routeName];
  const baseUrl = ScriptApp.getService().getUrl();
  if (!baseUrl) return '';
  const params = [];
  if (definition.parameter) params.push(encodeURIComponent(definition.parameter) + '=1');
  Object.keys(queryParams || {}).sort().forEach(function(key) {
    const value = queryParams[key];
    if (value !== undefined && value !== null) {
      params.push(encodeURIComponent(key) + '=' + encodeURIComponent(String(value)));
    }
  });
  return baseUrl + (params.length ? '?' + params.join('&') : '');
}

/** Folder of each page on our own domain, matching the output of web/build.js. */
const PUBLIC_ROUTE_PATHS_ = {attendance:'/', staff:'/staff/', req:'/requisition/', recovery:'/recover/', admin:'/admin/'};

/** The PUBLIC_BASE_URL setting, or '' when it is blank or not a plain https address. */
function normalizePublicBaseUrl_(value) {
  const base = String(value == null ? '' : value).trim().replace(/\/+$/,'');
  return /^https:\/\/[^\s?#"'<>]+$/i.test(base) ? base : '';
}

/**
 * The address of a page on our own domain. A recovery code goes after '#', which
 * browsers never send to a web server, so it cannot end up in the host's logs.
 */
function buildPublicRouteUrl_(base,route,queryParams) {
  const key = Object.prototype.hasOwnProperty.call(PUBLIC_ROUTE_PATHS_,route) ? route : 'attendance';
  const params = Object.keys(queryParams || {}).sort()
    .filter(name => queryParams[name] !== undefined && queryParams[name] !== null)
    .map(name => encodeURIComponent(name)+'='+encodeURIComponent(String(queryParams[name])));
  return base + PUBLIC_ROUTE_PATHS_[key] + (params.length ? (key === 'recovery' ? '#' : '?') + params.join('&') : '');
}

/**
 * Where people should be sent for a page: our own domain when PUBLIC_BASE_URL is
 * set, otherwise the Google web app. Email approval links do NOT use this, because
 * only the Google web app serves the approval page.
 */
function getPublicRouteUrl_(route,queryParams) {
  const base = normalizePublicBaseUrl_(getKeyValueSheet_(SETTINGS_SHEET).PUBLIC_BASE_URL);
  return base ? buildPublicRouteUrl_(base,route,queryParams) : getWebAppRouteUrl_(route,queryParams);
}

function getStaffRecoveryUrl() {
  return getPublicRouteUrl_('recovery');
}

function getPublicPortalLinks() {
  return [
    {label:'Staff portal',url:getPublicRouteUrl_('staff')},
    {label:'Requisitions',url:getPublicRouteUrl_('req')},
    {label:'Admin / ERP',url:getPublicRouteUrl_('admin')}
  ].filter(link=>Boolean(link.url));
}

function getAttendanceUrl() {
  return getPublicRouteUrl_('attendance');
}

function buildCanonicalWebAppUrl_(baseUrl,params) {
  const query = Object.keys(params || {}).sort().map(function(key) {
    return encodeURIComponent(key.toLowerCase())+'='+encodeURIComponent(String(params[key]));
  });
  return String(baseUrl || '')+(query.length ? '?'+query.join('&') : '');
}

function renderPage_(viewName, title,params) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.viewName = viewName;
  template.recoveryToken = '';
  template.canonicalUrl = buildCanonicalWebAppUrl_(ScriptApp.getService().getUrl(),params || {});
  return template.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
