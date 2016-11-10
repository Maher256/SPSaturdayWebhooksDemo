var express = require('express'),
    bodyParser = require('body-parser'),
    app = express(),
    server = require('http').Server(app),
    SPO = require('./GetSPOChanges.js'),
    io = require('socket.io')(server);

app.use(bodyParser.json());

app.post('/Webhook', function (req, res) {
    var message = processRequest(req);
    res.send(message);
});

function processRequest(req){
    console.log('Incoming request for:' + req.protocol + '://' + req.get('host') + req.originalUrl);
    
    if(req.query.validationtoken != null){ 
        console.log('Subscription confirmed with Token:' + req.query.validationtoken);
        return req.query.validationtoken; 
    }
    else{
        console.log('Notification received for id: '+ req.body.value[0].subscriptionId);
        SPO.getChanges(io, req.body.value[0].siteUrl, req.body.value[0].resource);
        return 'Request processed';
    }

}


server.listen(80);

app.get('/', function (req, res) {
  res.sendFile(__dirname + '/index.html');
});

app.use(express.static(__dirname + '/public'));


//nodemon index.js