var database = require('.');

(async function() {
    var conn = database.open({ Database: 'test.xlsx' });

    let results = await conn.query("SELECT * FROM Sheet1 WHERE Ranking = 10");
    console.log(results);
})();