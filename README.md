# Node Exchange Autodiscover
Retrieve the URL of your EWS by accessing Microsoft's SOAP Autodiscover Service.

Differs from other similar packages in that it queries the DNS. It also tries out all the potential
autodiscover urls in parallel, sacrificing a bit more bandwidth for speed.

## Usage

```javascript
var autodiscover = require('exchange-autodiscover');

// Promise
autodiscover({ emailAddress: "foo@bar.onmicrosoft.com", password: "pass"  })
  .then(console.log.bind(console))  // "https://outlook.microsoft.com/ews/exchange.asmx"
  .catch(console.error.bind(console))
  
// Callback
autodiscover({ emailAddress: "foobar@yourdomain.com", password: "pass", username: "ad\\foobar77" },
  function(err, ewsUrl) {
    if(err) {
      console.error(err);
    } else {
      console.log(ewsUrl); // "https://mail.yourdomain.com/ews/exchange.asmx"
    }
  }
```

## API

```javascript
/**
 * Tries to find the url of the EWS.
 *
 * @param   {Object}   params
 * @param   {String}   params.emailAddress
 * @param   {String}   params.password
 * @param   {String}   [params.username]    - Defaults to emailAddress
 * @param   {Boolean}  [params.queryDns]    - Defaults to true
 * @param   {Function} [cb]
 * @returns {Promise}  Resolves with the EWS url
 */
autodiscover(params, callback);
```

## License

See [license](LICENSE) (MIT License).
