# database-js-adodb
[![Build Status](https://ci.appveyor.com/api/projects/status/2acip2g9mv50lhc2?svg=true)](https://ci.appveyor.com/project/mlaanderson/database-js-adodb)

Database-js Wrapper for ADODB

## About
Database-js-adodb is a wrapper around the [node-adodb](https://github.com/nuintun/node-adodb) package by nuintun. It is intended to be used with the [database-js](https://github.com/mlaanderson/database-js) package. However it can also be used in stand alone mode. The only reason to do that would be to use [Promises](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise).

Because it depends on node-adodb, database-js-adodb will only work on Windows, or more correctly on operating systems which can do ADO access from the Windows Script cscript.
## Usage
### Stand Alone
~~~~
var adodb = require('database-js-adodb');

(async () => {
    let connection, rows;
    connection = mysql.open({
        Database: 'C:\\Users\\me\\Desktop\\access_db.mdb'
    });
    
    try {
        rows = await connection.query("SELECT * FROM tablea WHERE user_name = 'not_so_secret_user'");
        console.log(rows);
    } catch (error) {
        console.log(error);
    } finally {
        await connection.close();
    }
})();
~~~~
### With Database-js
~~~~
var Connection = require('database-js').Connection;

(async () => {
    let connection, statement, rows;
    connection = new Connection('database-js-adodb:///C:\\Users\\me\\Desktop\\access_db.mdb');
    
    try {
        statement = await connection.prepareStatement("SELECT * FROM tablea WHERE user_name = ?");
        rows = await statement.query('not_so_secret_user');
        console.log(rows);
    } catch (error) {
        console.log(error);
    } finally {
        await connection.close();
    }
})();
~~~~
### For Excel version 8
~~~~
var Connection = require('database-js').Connection;

(async () => {
    let connection, statement, rows;
    connection = new Connection('database-js-adodb:///C:\\Users\\me\\Desktop\\excel_file.xls?Extended Properties='Excel 8.0;HDR=Yes;IMEX=1';');
    
    try {
        statement = await connection.prepareStatement("SELECT * FROM [Sheet1$A1:C52] WHERE user_name = ?");
        rows = await statement.query('not_so_secret_user');
        console.log(rows);
    } catch (error) {
        console.log(error);
    } finally {
        await connection.close();
    }
})();
~~~~
