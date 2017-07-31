var ADODB = require('node-adodb');

class ADODB_Promise_Connection {
    constructor(connection) {
        this.__connection = connection;
    }

    /**
     * Runs the passed SQL as a query.
     * 
     * @param {string} sql 
     * @returns {Promise<array>}
     * @memberof ADODB_Promise_Connection
     */
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

    /**
     * Executes the passed SQL
     * 
     * @param {string} sql 
     * @param {any} scalar 
     * @returns {Promise<any>}
     * @memberof ADODB_Promise_Connection
     */
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

    /**
     * Does nothing, required for database-js
     * 
     * @returns {Promise<boolean>}
     * @memberof ADODB_Promise_Connection
     */
    close() {
        return Promise.resolve(true);
    }
}

module.exports = {
    /**
     * @param {ConnectionObject}
     * @returns {ADODB_Promise_Connection}
     */
    open: function(connection_parameters) {
        var connection_string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + connection_parameters.Database + ";" + (connection_parameters.Parameters ? connection_parameters.Parameters : '' );
        var base = ADODB.open(connection_string);
        return new ADODB_Promise_Connection(base);
    }
};