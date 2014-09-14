var express = require('express');
var app = express();

app.use(function(req, res, next){
  console.log('%s %s', req.method, req.url);
  next();
});

app.use(plain());
app.use(express.compress());
app.use(express.json());
app.use(express.urlencoded());
app.use(express.cookieParser('cookie-secret'));

// Standard
app.get('/get', standardResponse);
app.post('/post', standardResponse);
app.post('/text', standardResponse);
app.put('/put', standardResponse);
app.patch('/patch', standardResponse);
app.delete('/delete', standardResponse);

// Statuses
app.get('/status/:code', function(req, res) {
  res.send(parseInt(req.params.code));
});

// Timeout
app.get('/timeout', function(req, res) {
  var ms = req.query.ms || 10000;
  setTimeout(function() {
    res.send(200, 'Finally resolves (took ' + ms/1000 + ' seconds)');
  }, ms);
});

// JSON
app.get('/json', function(req, res) {
  res.json({a: '1', b: 2, c: 3.14, d: false, e: [4, 5], f: {a: '1', b: 2}});
});

// form-urlencoded
app.get('/formurlencoded', function(req, res) {
  res.send(200, 'a=1&b=2&c=3.14');
});

// xml
app.get('/xml', function(req, res) {
  res.send(200, '<Point><X>1.23</X><Y>4.56</Y></Point>')
});

// Cookies
app.get('/cookie', function(req, res) {
  res.cookie('unsigned-cookie', 'simple-cookie');
  res.cookie('signed-cookie', 'special-cookie', {signed: true});
  res.cookie('tricky;cookie', 'includes; semi-colon and space at end ');
  res.cookie('duplicate-cookie', 'A');
  res.cookie('duplicate-cookie', 'B');
  res.send(200);
});

// Simple text in body
app.get('/howdy', function(req, res) {
  res.send(200, 'Howdy!');
});

function standardResponse(req, res) {
  res.send(200, {
    method: req.route.method.toUpperCase(),
    query: req.query,
    headers: {
      'content-type': req.get('content-type'),
      'accept': req.get('accept'),
      'custom': req.get('custom'),
      'custom-a': req.get('custom-a'),
      'custom-b': req.get('custom-b')
    },
    body: req.text || req.body,
    cookies: req.cookies,
    signed_cookies: req.signedCookies
  });
}

function plain() {
  return function(req, res, next){
    if (req.is('text/*')) {
      req.text = '';
      req.setEncoding('utf8');
      req.on('data', function(chunk) { req.text += chunk });
      req.on('end', next);
    } else {
      next();
    }
  }
}

app.listen(3000);
console.log('Listening on port 3000');
