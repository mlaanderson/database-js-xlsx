var database = require('.');

(async function() {
    var conn = database.open({ Database: 'test.xlsx' });

    await conn.execute("INSERT INTO Sheet1(State, Ranking, Population) VALUES('Andersonia', 52, 5)")

    let results = await conn.query("SELECT * FROM Sheet1");
    console.log(results);
})();