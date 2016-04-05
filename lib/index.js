'use strict';

var Promise = require('bluebird');
var rp = require('request-promise');
var dnsResolve = Promise.promisify(require('dns').resolve);
var parseString = Promise.promisify(require('xml2js').parseString);
const EWS_URL_SETTING = 'ExternalEwsUrl';

/**
 * Removes the potential prefix of a string and makes the first character
 * lower case to make it easier to work with.
 *
 * @param {String} string
 * @returns {String}
 */
function removePrefix(string) {
  var splitString = string.split(":");
  var withoutPrefix = splitString[1] || splitString[0];
  return withoutPrefix.charAt(0).toLowerCase() + withoutPrefix.slice(1);
}

/**
 * Takes an XML string and transforms it into JSON and strips the tags
 * and attributes of any prefixes.
 *
 * @param  {String} xmlString
 * @return {Object} The XML string transformed into a JavaScript object.
 */
function xmlToJson(xmlString) {
  return parseString(xmlString, {
    tagNameProcessors: [removePrefix],
    attrNameProcessors: [removePrefix],
    explicitArray: false,
    mergeAttrs: true
  });
}

/**
 * Does a query on the DNS of the provided domain.
 * If it should fail, and empty array is returned.
 *
 * @param {String} domain
 * @returns {Promise} - Resolves with an array of other potential autodiscover domains
 */
function queryDns(domain) {
  return dnsResolve('_autodiscover._tcp.' + domain, 'SRV')
    .then(response => response.map(e => e.name))
    .catch(() => []);
}

function tryEndpoint(url, username, password, requestBody) {
  return rp({
    uri: url,
    method: 'POST',
    headers: {
      'Content-Type': 'text/xml; charset=utf-8'
    },
    auth: {
      user: username,
      pass: password
    },
    body: requestBody,
    followRedirect: false
  });
}

/**
 * Formats requested settings for SOAP (i.e. <a:Setting>SETTING</a:Setting>).
 *
 * https://msdn.microsoft.com/en-us/library/office/dd877068(v=exchg.150).aspx
 *
 * @param  {?Array|String} settings List of AD settings to request.
 * @return {String} Setting nodes to embed in SOAP XML request.
 */
function wrapSettingsRequest(settings) {
  // Allows settings to be both arrays and single strings
  settings = [].concat(settings);
  // Only add EWS Url if it's missing
  if (settings.indexOf(EWS_URL_SETTING) === -1)
    settings.push(EWS_URL_SETTING)
  return settings.map(setting => `<a:Setting>${setting}</a:Setting>`).join('');
};

function createAutodiscoverSoap(emailAddress, settings) {
  return '' +
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:a="http://schemas.microsoft.com/exchange/2010/Autodiscover" ' +
    'xmlns:wsa="http://www.w3.org/2005/08/addressing" ' +
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
    'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    '  <soap:Header>' +
    '    <a:RequestedServerVersion>Exchange2010</a:RequestedServerVersion>' +
    '    <wsa:Action>http://schemas.microsoft.com/exchange/2010/Autodiscover/Autodiscover/GetUserSettings</wsa:Action>' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <a:GetUserSettingsRequestMessage xmlns:a="http://schemas.microsoft.com/exchange/2010/Autodiscover">' +
    '      <a:Request>' +
    '        <a:Users>' +
    '          <a:User>' +
    '            <a:Mailbox>' + emailAddress + '</a:Mailbox>' +
    '          </a:User>' +
    '        </a:Users>' +
    '        <a:RequestedSettings>' + wrapSettingsRequest(settings) + '</a:RequestedSettings>' +
    '      </a:Request>' +
    '    </a:GetUserSettingsRequestMessage>' +
    '  </soap:Body>' +
    '</soap:Envelope>';
}

/**
 * Tries every possible autodiscover url in parallel.
 *
 * https://msdn.microsoft.com/en-us/library/office/jj900169(v=exchg.150).aspx
 *
 * @param {Array} domains
 * @param {String} emailAddress
 * @param {String} password
 * @param {String} username
 * @param {?Array|String} Requested settings
 * @returns {Promise}
 */
function autodiscoverDomains(domains, emailAddress, password, username, settings) {
  var promises = [];
  var requestBody = createAutodiscoverSoap(emailAddress, settings);
  domains.forEach(domain => {
    promises.push(tryEndpoint('https://' + domain + '/autodiscover/autodiscover.svc',
      username, password, requestBody));

    promises.push(tryEndpoint('https://autodiscover.' + domain + '/autodiscover/autodiscover.svc',
      username, password, requestBody));

    promises.push(rp({
      uri: 'http://autodiscover.' + domain + '/autodiscover/autodiscover.svc',
      method: 'GET',
      followRedirect: false,
      simple: false,
      resolveWithFullResponse: true
    }).then(response => {
      // Just take redirects into consideration.
      if (response.statusCode !== 302) {
        throw new Error();
      }
      return tryEndpoint(response.headers.location, username, password, requestBody);
    }));
  });

  return Promise
    .any(promises)
    .then(xmlToJson)
    .then(result => {
      var userSettings = result.envelope.body.getUserSettingsResponseMessage
        .response.userResponses.userResponse.userSettings.userSetting;
      // Make sure we're working with an array
      userSettings = [].concat(userSettings);
      return userSettings.reduce((userSettings, setting) => {
        userSettings[setting.name] = setting.value;
        return userSettings;
      }, {});
    });
}

/**
 * Tries to find the url of the EWS.
 *
 * @param {Object} params
 * @param {String} params.emailAddress
 * @param {String} params.password
 * @param {String} [params.username]
 * @param {Boolean} [params.queryDns]
 * @param {Array} [params.settings]
 * @param {Function} [cb]
 * @returns {Promise} Resolves with the EWS url
 */
module.exports = function (params, cb) {
  var emailAddress = params.emailAddress;
  var password = params.password;
  var username = params.username || emailAddress;
  var query = params.queryDns || true;
  var requestedSettings = params.settings;

  var smtpDomain = emailAddress.substr(emailAddress.indexOf("@") + 1);
  var domains = [smtpDomain];
  var promise;
  if (query) {
    promise = queryDns(smtpDomain);
  } else {
    promise = Promise.resolve([]);
  }

  return promise
    .then(otherDomains => {
      domains = domains.concat(otherDomains);
      return autodiscoverDomains(domains, emailAddress, password, username, requestedSettings);
    })
    .then(settings => {
      // If no extra settings were requested, just return the EWS URL as string
      var result = requestedSettings ? settings : settings[EWS_URL_SETTING];
      if (cb) {
        cb(null, result);
      }
      return result;
    })
    .catch(errors => {
      if (cb) {
        cb(errors);
      }
      throw errors;
    });
};
