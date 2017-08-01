# database-js-xlsx
Database-js Wrapper for XLSX files
## About
Database-js-firebase is a [database-js](https://github.com/mlaanderson/database-js) wrapper around the [xlsx-populate](https://github.com/dtjohnson/xlsx-populate) library by Dave Johnson. 

Tables are emulated by using a defined address `Sheet1$A1:C52`, or just a sheet name `Sheet1`. The first row must be the column headings.

xlsx-populate works with an in-memory copy of the spreadsheet, so the underlying spreadsheet can be changed after it is loaded by database-js-xlsx. When the connection is closed, the in-memory copy is written back to disk and any changes outside of the database-js-xlsx will be overwritten.

xlsx-populate works cross platform. This means that database-js-xlsx also works cross platform, unlike the database-js-adodb driver which is Windows only.

database-js-xlsx uses the [node-sqlparser](https://github.com/alibaba/nquery) library by fish. SQL commands are limited to SELECT, UPDATE, INSERT and DELETE. WHERE works well. JOINs are not allowed. GROUP BY is not yet supported. LIMIT and OFFSET are combined into a single LIMIT syntax: `LIMIT [offset,]number`

## Usage
~~~~
var Database = require('database-js2');

(async () => {
    let connection, statement, rows;
    connection = new Database('database-js-xlsx:///test.xlsx');
    
    try {
        statement = await connection.prepareStatement("SELECT * FROM Sheet1 WHERE State = ?");
        rows = await statement.query('South Dakota');
        console.log(rows);
    } catch (error) {
        console.log(error);
    } finally {
        await connection.close();
    }
})();
~~~~