var ADODB = require(".");

var connection = ADODB.open({
    Database: 'test.mdb'
});

function handleError(error) {
    console.log("ERROR:", error);
    process.exit(1);
}

connection.query("SELECT * FROM states WHERE State = 'South Dakota'").then((data) => {
    if (data.length != 1) {
        handleError(new Error("Invalid data returned"));
    }
    console.log(data);
    connection.close().then(() => {
        process.exit(0);
    }).catch(handleError);
}).catch(handleError);
