# Node Exchange Autodiscover

Retrieve the URL of your EWS by accessing Microsoft's SOAP Autodiscover Service.

Differs from other similar packages in that it queries the DNS. It also tries out all the potential
autodiscover urls in parallel, sacrificing a bit more bandwidth for speed.

## Usage

```javascript
const autodiscover = require('exchange-autodiscover');

// Promise
autodiscover({ emailAddress: "foo@bar.onmicrosoft.com", password: "pass" })
  .then(console.log.bind(console))  // 'https://outlook.microsoft.com/ews/exchange.asmx'
  .catch(console.error.bind(console));

// Callback
autodiscover({
  emailAddress: "foobar@yourdomain.com",
  password: "pass",
  username: "ad\\foobar77"
}, function(err, ewsUrl) {
  if (err) {
    console.error(err);
  } else {
    console.log(ewsUrl); // 'https://mail.yourdomain.com/ews/exchange.asmx'
  }
});
```

You can optionally request any of the [extra settings found here](https://msdn.microsoft.com/en-us/library/office/dd877068(v=exchg.150).aspx).
This will return an object with matched settings (EWS address will always be included by default).

```javascript
// With extra settings
autodiscover({
  emailAddress: "foo@bar.onmicrosoft.com",
  password: "pass",
  settings: [
    'EwsSupportedSchemas',
    'ExternalEwsVersion'
  ]
}).then(function (settings) {
  console.log(settings);
  // {
  //   EwsSupportedSchemas: 'Exchange2007, Exchange2007_SP1, Exchange2010, Exchange2010_SP1, Exchange2010_SP2, Exchange2013, Exchange2013_SP1, Exchange2015',
  //   ExternalEwsUrl: 'https://outlook.office365.com/EWS/Exchange.asmx',
  //   ExternalEwsVersion: '15.00.0000.000'
  // }
});
```

## API

```javascript
/**
 * Tries to find the url of the EWS.
 * @param   {Object}   params
 * @param   {string}   params.emailAddress
 * @param   {string}   params.password
 * @param   {string}   [params.username] - Defaults to emailAddress
 * @param   {boolean}  [params.queryDns] - Defaults to true
 * @param   {string[]} [params.settings]
 * @param   {Function} [cb]
 * @returns {Promise}  Resolves with the EWS url
 */
autodiscover(params, callback);
```
