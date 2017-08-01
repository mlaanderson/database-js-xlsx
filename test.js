var database = require('.');

(async function() {
    var conn = database.open({ Database: 'test.xlsx' });

    await conn.execute("DELETE FROM Sheet1 WHERE State LIKE 'North%'");
    let results = await conn.query("SELECT * FROM Sheet1");
    console.log(results);
    //await conn.close();
})();