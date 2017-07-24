var ADODB = require('node-adodb');

class ADODB_Promise_Connection {
    constructor(connection) {
        this.__connection = connection;
    }

    query(sql) {
        var self = this;
        return new Promise((resolve, reject) => {
            self.__connection.query(sql)
            .on('done', (data, message) => {
                resolve(data, message);
            })
            .on('fail', (error) => {
                reject(error);
            });
        });
    }

    execute(sql, scalar) {
        var self = this;
        return new Promise((resolve, reject) => {
            self.__connection.execute(sql, scalar)
            .on('done', (data, message) => {
                resolve(data, message);
            })
            .on('fail', (error) => {
                reject(error);
            });
        });
    }

    close() {

    }
}

module.exports = {
    open: function(connection_parameters) {
        var connection_string = `Provider=Microsoft.Jet.OLEDB.4.0;Data Source=${connection_parameters.Database};${ connection_parameters.Parameters ? connection_parameters.Parameters : '' }`
        var base = ADODB.open(connection_string);
        return new ADODB_Promise_Connection(base);
    }
};