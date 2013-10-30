var express = require('express');
var app = express();

app.use(plain());
app.use(express.json());
app.use(express.urlencoded());

app.use(function(req, res, next){
  console.log('%s %s', req.method, req.url);
  next();
});

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

function standardResponse(req, res) {
  res.send(200, {
    method: req.route.method.toUpperCase(),
    query: req.query,
    headers: {
      'content-type': req.get('content-type'),
      'custom': req.get('custom')
    },
    body: req.text || req.body
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
