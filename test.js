var database = require('.');

var connection = database.open({ Database: 'test.xlsx' });

function handleError(error) {
    console.log("ERROR:", error);
    process.exit(1);
}


connection.query("SELECT * FROM Sheet1$A1:C52 Where State = 'South Dakota'").then((data) => {
    if (data.length != 1) {
        handleError(new Error("Invalid data returned"));
    }
    connection.close().then(() => {
        process.exit(0);
    }).catch(handleError);
}).catch(handleError);