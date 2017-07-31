var XlsxPopulate = require('xlsx-populate');
var parse = require('node-sqlparser').parse;

var addressRegex = /^(?:((?:[^!$]+)|(?:\'[^\']+'))\$)?([A-Z]{1,3}\d+)(?::([A-Z]{1,3}\d+))?$/;

class AddressParser {
    constructor(address) {
        let match;
        if (addressRegex.test(address)) {
            [match, this.Sheet, this.Start, this.End] = addressRegex.exec(address);
            if (this.Sheet === undefined) {
                this.Sheet = 0;
            }
            this.Matched = true;
        } else {
            this.Matched = false;
        }
    }
}

class XlsxTable {
    constructor(database, address) {
        let range;
        if (address) {
            let parsedAddress = new AddressParser(address);
            if (parsedAddress.Matched) {
                this.worksheet = database.workbook.sheet(parsedAddress.Sheet);
                if (parsedAddress.End) {
                    range = this.worksheet.range(parsedAddress.Start + ":" + parsedAddress.End);
                } else {
                    let startCell = this.worksheet.Cell(parsedAddress.Start);
                    let startRowNumber = startCell.rowNumber();
                    let startColumnNumber = startCell.columnNumber();
                    let endRowNumber = this.worksheet.usedRange().endCell().rowNumber();
                    let endColumnNumber = this.worksheet.usedRange().endCell().columnNumber();
                    range = this.worksheet.range(startRowNumber, startColumnNumber, endRowNumber, endColumnNumber);
                }
            } else {
                // try to see if it's a valid sheet
                this.worksheet = database.workbook.sheet(address);
                if (this.worksheet === null) {
                    throw new Error("Cannot find address " + address);
                }
                range = this.worksheet.usedRange();
            }
        } else {
            this.worksheet = database.workbook.sheet(0);
            range = this.worksheet.usedRange();
        }

        this.headers = this.worksheet.range(range.startCell().rowNumber(), range.startCell().columnNumber(), 
            range.startCell().rowNumber(), range.endCell().columnNumber());

        this.body = this.worksheet.range(range.startCell().rowNumber() + 1, range.startCell().columnNumber(), 
            range.endCell().rowNumber(), range.endCell().columnNumber());
    }

    get headerText() {
        return this.headers.map(cell => cell.value())[0];
    }

    get width() {
        return this.body.endCell().columnNumber() - this.body.startCell().columnNumber() + 1;
    }

    get height() {
        return this.body.endCell().rowNumber() - this.body.startCell().rowNumber() + 1;
    }

    header(column) {
        return this.headers.cell(0, column).value();
    }

    value(rowNumber, columnNumber) {
        return this.body.cell(rowNumber, columnNumber).value();
    }
}

class XlsxDatabase {
    constructor(options) {
        var self = this;
        let opts, defaults = {
            filename: null,
            data: null
        };

        options = options || {};

        if (typeof options === 'string') {
            if (/\.xlsx/i.test(options)) {
                options = {
                    filename: options
                }
            } else {
                options = {
                    data: options
                }
            }
        }

        if (
            (options.__proto__ === Uint8Array.prototype) ||
            (options.__proto__ === ArrayBuffer.prototype) || 
            ((options.__proto__ === Array.prototype) && (options.every(v => typeof v === "number"))) ||
            (options.__proto__ === Promise.prototype)
        ) {
            options = {
                data: options
            }
        }

        opts = Object.assign(defaults, options);

        if (opts.filename) {
            this.loader = new Promise((resolve, reject) => {
                XlsxPopulate.fromFileAsync(opts.filename).then(workbook => {
                    this.workbook = workbook;
                    resolve();
                });
            });
        } else if (opts.data) {
            this.loader = new Promise((resolve, reject) => {
                XlsxPopulate.fromDataAsync(data).then(workbook => {
                    this.workbook = workbook;
                    resolve();
                });
            });
        } else {
            this.loader = new Promise((resolve, reject) => {
                XlsxPopulate.fromBlankAsync().then(workbook => {
                    this.workbook = workbook;
                    resolve();
                });
            });
        }
    }

    ready() {
        return this.loader;
    }


    /**
     * Tests the passed row based on the where array.
     * This could be faster on the server if the data has been
     * indexed, and if the user only wants a single WHERE.
     * 
     * @param {object} where Where object from the SQL Parser
     * @param {object} row Row to compare
     * @returns {boolean} if the row matches the where object
     * @memberof XlsxDatabase
     */
    doWhere(where, row) {
        if (where === null) return true;
        var self = this;

        function getVal(obj) {
            if (obj.type === "column_ref") return row[obj.column];
            if (obj.type === "binary_expr") return self.doWhere(obj, row);
            return obj.value;
        }

        switch (where.type) {
            case "binary_expr":
                switch(where.operator) {
                    case "=":
                        return getVal(where.left) == getVal(where.right);
                    case "!=":
                    case "<>":
                        return getVal(where.left) != getVal(where.right);
                    case "AND":
                        return getVal(where.left) && getVal(where.right);
                    case "OR":
                        return getVal(where.left) && getVal(where.right);
                    case "IS":
                        return getVal(where.left) === getVal(where.right)
                    default:
                        return false;
                }
                break;
            default:
                return false;
        }
    }

    /**
     * Used to push a row into the data object. If the fields are limited
     * in the query, only places the requested fields.
     * 
     * @param {object} sqlObj 
     * @param {Array} data 
     * @param {object} row 
     * @returns 
     * @memberof XlsxDatabase
     */
    chooseFields(sqlObj, data, row) {
        if (sqlObj.columns === "*") {
            data.push(row);
            return;
        }

        let isAggregate = sqlObj.columns.some((col) => { return col.expr.type === 'aggr_func'; });

        if (isAggregate === true) {
            if (data.length === 0) {
                data.push({});
            }

            for (let col of sqlObj.columns) {
                let name, data_row;
                switch(col.expr.type) {
                    case 'column_ref':
                        name = col.as || col.expr.column;
                        data[0][name] = row[col.expr.column];
                        break;
                    case 'aggr_func': // TODO implement group by
                        name = col.as || col.expr.name.toUpperCase() + "(" + col.expr.args.expr.column + ")";
                        
                        switch(col.expr.name.toUpperCase()) {
                            case 'SUM':
                                if (data[0][name] === undefined) {
                                    data[0][name] = 0;
                                }
                                data[0][name] += row[col.expr.args.expr.column];
                                break;
                            case 'COUNT':
                                if (data[0][name] === undefined) {
                                    data[0][name] = 0;
                                }
                                data[0][name]++;
                                break;
                        }
                        break;
                }
            }
        } else {
            let result = {};
            for (let col of sqlObj.columns) {
                let name = col.as || col.expr.column;
                result[name] = row[col.expr.column];
            }
            data.push(result);
        }
    }

    /**
     * Performs an SQL SELECT. This is called from a Promise.
     * 
     * @param {function} resolve 
     * @param {function} reject 
     * @param {any} sqlObj 
     * @returns 
     * @memberof XlsxDatabase
     */
    doSelect(resolve, reject, sqlObj) {
        if (sqlObj.from.length !== 1) {
            return reject("Selects from more than one table are not supported");
        }
        
        if (sqlObj.groupby !== null) {
            console.warn("GROUP BY is unsupported");
        }

        let xlTable = new XlsxTable(this, sqlObj.from[0].table);
        let raw = xlTable.body.map(cell => cell.value());
        let headers = xlTable.headerText;
        let rows = [];
        for (let row of raw) {
            let oRow = {};
            for (let n = 0; n < headers.length; n++) {
                oRow[headers[n]] = row[n];
            }
            if (this.doWhere(sqlObj.where, oRow) === true) {
                this.chooseFields(sqlObj, rows, oRow);
            }
        }

        if (sqlObj.orderby) {
            rows.sort((a, b) => {
                for (let orderer of sqlObj.orderby) {
                    if (orderer.expr.type !== 'column_ref') {
                        throw new Error("ORDER BY only supported for columns, aggregates are not supported");
                    }

                    if (a[orderer.expr.column] > b[orderer.expr.column]) {
                        return orderer.type = 'ASC' ? 1 : -1;
                    }
                    if (a[orderer.expr.column] < b[orderer.expr.column]) {
                        return orderer.type = 'ASC' ? -1 : 1;
                    }
                }
                return 0;
            });
        }

        if (sqlObj.limit) {
            if (sqlObj.limit.length !== 2) {
                throw new Error("Invalid LIMIT expression: Use LIMIT [offset,] number");
            }
            let offs = parseInt(sqlObj.limit[0].value);
            let len = parseInt(sqlObj.limit[1].value);
            rows = rows.slice(offs, offs + len);
        }
        resolve(rows);
    }

    /**
     * Runs the SQL statement
     * 
     * @param {string} sql 
     * @returns {Promise<array>} Promise of array of selected rows, updated rows, inserted rows, or deleted row XlsxDatabase keys
     * @memberof XlsxDatabase
     */
    runSQL(sql) {
        var self = this;
        return new Promise((resolve, reject) => {
            this.ready().then(() => {
                // we are now loaded
                let sqlObj;
                try {
                    sqlObj = parse(sql);
                } catch (err) {
                    // deletes aren't yet supported by the node-sqlparser
                    // so fake a SELECT and then change the type after the parse
                    if (/^delete/i.test(sql) === true) {
                        sql = sql.replace(/^delete/i, 'SELECT * ');
                        sqlObj = parse(sql);
                        sqlObj.type = 'delete';
                        delete sqlObj.columns;
                    } else {
                        reject(err);
                    }
                }

                switch(sqlObj.type) {
                    case 'select':
                        this.doSelect(resolve, reject, sqlObj);
                        break;
                    // case 'update':
                    //     this.doUpdate(resolve, reject, sqlObj);
                    //     break;
                    // case 'insert':
                    //     this.doInsert(resolve, reject, sqlObj);
                    //     break;
                    // case 'delete':
                    //     this.doDelete(resolve, reject, sqlObj);
                    //     break;
                    default:
                        resolve(sqlObj);
                        break;
                }
            });
        });
    }

    /**
     * Executes the passed SQL
     * 
     * @param {string} sql 
     * @returns {Promise<array>} Promise of array of selected rows, updated rows, inserted rows, or deleted row XlsxDatabase keys
     * @memberof XlsxDatabase
     */
    execute(sql) {
        return this.runSQL(sql);
    }

    /**
     * Executes the passed SQL
     * 
     * @param {string} sql 
     * @returns {Promise<array>} Promise of array of selected rows, updated rows, inserted rows, or deleted row XlsxDatabase keys
     * @memberof XlsxDatabase
     */
    query(sql) {
        return this.runSQL(sql);
    }

    /**
     * Closes the connection, sets XlsxDatabase to offline mode.
     * 
     * @returns {Promise<boolean>}
     * @memberof XlsxDatabase
     */
    close() {
        return Promise.resolve(true);
    }

}

module.exports = {
    open: function(connection) {
        return new XlsxDatabase(connection.Database);
    }
};
