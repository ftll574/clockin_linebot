var Connection = require('tedious').Connection;
var Request = require('tedious').Request;

var config = 
   {
     userName: 'Zhenyu', // update me
     password: 'Wheal60908', // update me
     server: 'mindwave-user-data-by-zhenyu.database.windows.net', // update me
     options: 
        {
           database: 'MindWave_User_Data_by_Zhenyu' //update me
           , encrypt: true
        }
   };
   
var connection = new Connection(config);

connection.on('connect', function(err) 
   {
     if (err) 
       {
          console.log(err)
       }
    else
       {
           queryDatabase()
       }
   }
 );
 
 function queryDatabase()
   { console.log('Reading rows from the Table...');

       // Read all rows from table
     request = new Request(
          "SELECT [User_ID],[MindWave_Max],[MindWave_Mean],[MindWave_Ske],[MindWave_Kur],[MindWave_FFT_Max],[MindWave_FFT_Mean],[MindWave_PSD_Max],[MindWave_PSD_Mean]FROM [dbo].[User_data]",
             function(err, rowCount, rows) 
                {
                    console.log(rowCount + ' row(s) returned');
                    process.exit();
                }
            );

     request.on('row', function(columns) {
        columns.forEach(function(column) {
            console.log("%s\t%s", column.metadata.colName, column.value);
         });
             });
     connection.execSql(request);
   }

var server,
    ip   = "127.0.0.1",
    port = 1111,
    http = require('http');

    server = http.createServer(function (req, res) {
    res.writeHead(200, {'Content-Type': 'text/plain'});
    res.end('Hello World\n');
});

server.listen(port, ip);

console.log("Server running at http://" + ip + ":" + port);