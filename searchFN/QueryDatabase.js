var http = require('http');
var fs = require('fs');
var url = require('url');
var db = require('/QOpenSys/QIBM/ProdData/OPS/Node6/os400/db2i/lib/db2a');

var DBname = "*LOCAL";
var userId = "USER";
var passwd = "PASSWORD";
var ip = "URL";
var port = ****;

var webserver = http.createServer((req,res) => {
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Request-Method', '*');
        res.setHeader('Access-Control-Allow-Methods', 'OPTIONS, GET');
        res.setHeader('Access-Control-Allow-Headers', '*');
	var sql = url.parse(req.url, true).query.sql;
	var dbconn = new db.dbconn();
        dbconn.conn(DBname, userId, passwd);  // Connect to the DB
        var stmt = new db.dbstmt(dbconn);
        stmt.exec(sql, (rs) => { // Query the statement
          res.end(JSON.stringify(rs));
          stmt.close();
          dbconn.disconn();
          dbconn.close();
        });
});

webserver.listen(port, ip);

console.log('Server running at http://' + ip + ':' + port);
