/*
 * Stands in for google.script.run on pages served from our own domain.
 *
 * The page code stays exactly as it is for the Google-hosted version: it chains
 * withSuccessHandler / withFailureHandler / withUserObject and calls a server
 * function by name. Here that call becomes a POST to the Apps Script web app.
 *
 * Nothing is retried. A request that saves something may have reached the server
 * even when the reply is lost, and sending it again would save it twice.
 */
(function () {
  if (window.google && window.google.script && window.google.script.run) return;

  var NOT_SET_UP = 'This site is not connected to the server yet. Set the web app address in config.js.';
  var NO_REPLY = 'Could not reach the server. Check your connection. If you were saving something, check whether it was saved before trying again.';
  var BAD_REPLY = 'The server could not be reached properly. Admin: check the web app address in config.js and that the deployment is open to anyone.';

  function request(name, args) {
    var url = String(window.MS_API_URL || '');
    if (!/^https?:\/\/\S+$/.test(url)) return Promise.reject(new Error(NOT_SET_UP));
    return window.fetch(url, {
      method: 'POST',
      // text/plain keeps this a "simple" request: Apps Script cannot answer a CORS pre-check.
      headers: {'Content-Type': 'text/plain;charset=utf-8'},
      body: JSON.stringify({fn: name, args: args}),
      credentials: 'omit',
      redirect: 'follow',
      cache: 'no-store'
    }).then(function (response) {
      return response.text();
    }, function () {
      throw new Error(NO_REPLY);
    }).then(function (text) {
      var reply;
      try { reply = JSON.parse(text); } catch (error) { throw new Error(BAD_REPLY); }
      if (!reply || typeof reply.ok !== 'boolean') throw new Error(BAD_REPLY);
      if (!reply.ok) throw new Error(reply.error || 'The request failed.');
      return reply.value;
    });
  }

  function runner(handlers) {
    return new Proxy({}, {
      get: function (target, name) {
        if (typeof name !== 'string') return undefined;
        var with_ = function (key) {
          return function (value) {
            var next = {success: handlers.success, failure: handlers.failure, userObject: handlers.userObject};
            next[key] = value;
            return runner(next);
          };
        };
        if (name === 'withSuccessHandler') return with_('success');
        if (name === 'withFailureHandler') return with_('failure');
        if (name === 'withUserObject') return with_('userObject');
        return function () {
          // Two separate callbacks: an error thrown by the success handler must not
          // be reported as a failed request.
          request(name, Array.prototype.slice.call(arguments)).then(function (value) {
            if (handlers.success) handlers.success(value, handlers.userObject);
          }, function (error) {
            if (handlers.failure) handlers.failure(error, handlers.userObject);
            else if (window.console) window.console.error(name + ' failed: ' + error.message);
          });
        };
      }
    });
  }

  window.google = window.google || {};
  window.google.script = window.google.script || {};
  Object.defineProperty(window.google.script, 'run', {configurable: true, get: function () { return runner({}); }});
})();
