const Promise = require('bluebird');
const rp = require('request-promise');
const dnsResolve = Promise.promisify(require('dns').resolve);
const parseString = Promise.promisify(require('xml2js').parseString);
const EWS_URL_SETTING = 'ExternalEwsUrl';

/**
 * Removes the potential prefix of a string and makes the first character
 * lower case to make it easier to work with.
 * @param   {string} string - Potentially prefixed string.
 * @returns {string} String without prefix.
 */
function removePrefix(string) {
  const splitString = string.split(":");
  const withoutPrefix = splitString[1] || splitString[0];
  return withoutPrefix.charAt(0).toLowerCase() + withoutPrefix.slice(1);
}

/**
 * Takes an XML string and transforms it into an object and strips the tags
 * and attributes of any prefixes.
 * @param  {string} xmlString
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
 * @param   {string} domain
 * @returns {Promise} Resolves with an array of other potential autodiscover domains.
 */
function queryDns(domain) {
  return dnsResolve('_autodiscover._tcp.' + domain, 'SRV')
    .then(response => response.map(e => e.name))
    .catch(() => []);
}

/**
 * @param {string} url
 * @param {string} username
 * @param {string} password
 * @param {string} requestBody
 * @returns {Promise}
 */
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
 * https://msdn.microsoft.com/en-us/library/office/dd877068(v=exchg.150).aspx
 * @param  {(string|Array)} settings - List of AD settings to request.
 * @return {string} Setting nodes to embed in SOAP XML request.
 */
function wrapSettingsRequest(settings) {
  // Allows settings to be both arrays and single strings.
  settings = [].concat(settings);
  // Only add EWS Url if it's missing.
  if (settings.indexOf(EWS_URL_SETTING) === -1) {
    settings.push(EWS_URL_SETTING);
  }

  return settings.map(setting => `<a:Setting>${setting}</a:Setting>`).join('');
};

function createAutodiscoverSoap(emailAddress, settings) {
  return `<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope
      xmlns:a="http://schemas.microsoft.com/exchange/2010/Autodiscover"
      xmlns:wsa="http://www.w3.org/2005/08/addressing"
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
      xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
    >
      <soap:Header>
        <a:RequestedServerVersion>Exchange2010</a:RequestedServerVersion>
        <wsa:Action>http://schemas.microsoft.com/exchange/2010/Autodiscover/Autodiscover/GetUserSettings</wsa:Action>
      </soap:Header>
      <soap:Body>
        <a:GetUserSettingsRequestMessage xmlns:a="http://schemas.microsoft.com/exchange/2010/Autodiscover">
          <a:Request>
            <a:Users>
              <a:User>
                <a:Mailbox>${emailAddress}</a:Mailbox>
              </a:User>
            </a:Users>
            <a:RequestedSettings>${wrapSettingsRequest(settings)}</a:RequestedSettings>
          </a:Request>
        </a:GetUserSettingsRequestMessage>
      </soap:Body>
    </soap:Envelope>
  `;
}

/**
 * Tries every possible autodiscover url in parallel.
 * https://msdn.microsoft.com/en-us/library/office/jj900169(v=exchg.150).aspx
 * @param {Array} domains
 * @param {string} emailAddress
 * @param {string} password
 * @param {string} username
 * @param {(string|Array)} settings
 * @returns {Promise}
 */
function autodiscoverDomains(domains, emailAddress, password, username, settings) {
  const promises = [];
  const requestBody = createAutodiscoverSoap(emailAddress, settings);

  domains.forEach(domain => {
    promises.push(tryEndpoint(`https://${domain}/autodiscover/autodiscover.svc`, username, password, requestBody));
    promises.push(tryEndpoint(`https://autodiscover.${domain}/autodiscover/autodiscover.svc`, username, password, requestBody));
    promises.push(rp({
      method: 'GET',
      uri: `http://autodiscover.${domain}/autodiscover/autodiscover.svc`,
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
      let userSettings = result.envelope.body.getUserSettingsResponseMessage
        .response.userResponses.userResponse.userSettings.userSetting;

      // Cancel all other network requests.
      promises.forEach(promise => promise.cancel());
      // Make sure we're working with an array.
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
 * @param {string} params.emailAddress
 * @param {string} params.password
 * @param {string} [params.username]
 * @param {boolean} [params.queryDns]
 * @param {Array} [params.settings]
 * @param {Function} [cb]
 * @returns {Promise} Resolves with the EWS url
 */
module.exports = function (params, cb) {
  const emailAddress = params.emailAddress;
  const password = params.password;
  const username = params.username || emailAddress;
  const query = 'queryDns' in params ? params.queryDns : true;
  const requestedSettings = params.settings;
  const smtpDomain = emailAddress.substr(emailAddress.indexOf("@") + 1);
  const promise = query ? queryDns(smtpDomain) : Promise.resolve([]);
  let domains = [smtpDomain];

  return promise.then(otherDomains => {
    domains = domains.concat(otherDomains);
    return autodiscoverDomains(domains, emailAddress, password, username, requestedSettings);
  }).then(settings => {
    // If no extra settings were requested, just return the EWS URL as string
    const result = requestedSettings ? settings : settings[EWS_URL_SETTING];
    if (cb) {
      cb(null, result);
    }
    return result;
  }).catch(errors => {
    if (cb) {
      cb(errors);
    }
    throw errors;
  });
};
