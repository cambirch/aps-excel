///<reference path='node.d.ts'/>
var aps;
(function (aps) {
    var edge = require('edge');

    var Excel = (function () {
        function Excel() {
        }
        Excel.prototype.load = function (path) {
            var workbook = functions({ fn: 'excel-load', path: path }, true);
            return new Workbook(workbook);
        };
        return Excel;
    })();
    aps.Excel = Excel;

    var Workbook = (function () {
        function Workbook(workbook) {
            this.workbook = workbook;
        }
        Workbook.prototype.save = function (path) {
            functions({ fn: 'workbook-save', path: path, workbook: this.workbook }, true);
        };

        Workbook.prototype.getSheetAt = function (index) {
            var sheet = functions({ fn: 'workbook-getSheetAt', index: index, workbook: this.workbook }, true);
            return new Sheet(sheet);
        };
        Workbook.prototype.getSheet = function (name) {
            var sheet = functions({ fn: 'workbook-getSheet', name: name, workbook: this.workbook }, true);
            return new Sheet(sheet);
        };
        return Workbook;
    })();
    aps.Workbook = Workbook;

    var Sheet = (function () {
        function Sheet(sheet) {
            this.sheet = sheet;
        }
        /**
        * getRow returns a row object that you can read or edit
        * note there is a flag for whether the row previously existed on not
        * so it is safe to call getRow without first calling getRowExists
        * if getRow returns you a row that doesn't exist you will need to call either:
        *     Row.create() or
        *     Sheet.createRow(####)
        */
        Sheet.prototype.getRow = function (row) {
            if (row === 0)
                throw new Error('row 0 is not valid');

            var rowData = functions({ fn: 'sheet-getRow', sheet: this.sheet, row: row - 1 }, true);
            if (rowData == null)
                return null;
            return new Row(this.sheet, rowData, row);
        };

        /**
        * getRowExists lets you know whether a specific row exists.
        * if you call it with a row number greater than getRowCount()
        * you should expect it to return false.
        * if you call it with a row number less or equal to getRowCount()
        * it will let you know whether that row exists.
        */
        Sheet.prototype.getRowExists = function (row) {
            if (row === 0)
                throw new Error('row 0 is not valid');

            return functions({ fn: 'sheet-getRowExists', sheet: this.sheet, row: row - 1 }, true);
        };

        /**
        * getRowCount() tells you NOT how 'many' rows exist, but rather,
        * what is the last row that MIGHT exist. So if getRowCount() returns 99
        * it is possible that the only row existing is row 99, the first 98 may
        * be blank and non-existing. This makes it excellent when reading through
        * a file to know when to stop looking for more rows.
        * so, you might say for row 12 to getRowCount() - read the row and see if
        * is any data in that row that I care about.
        */
        Sheet.prototype.getRowCount = function () {
            return functions({ fn: 'sheet-getRowCount', sheet: this.sheet }, true);
        };

        /**
        * cloneRow() takes the source row and copies it to the destination row
        * it does NOT insert it, it deletes whatever is currently in the destination
        * row including formatting.
        * So if you want to do an insert, with the current API you would have to
        * - close last row to last plus 1, clone last-1 to initial last etc.., In other
        * words, avoid doing an insert!
        *
        * LETS ASSUME YOU ARE CREATING A SPREADSHEET AS AN EXPORT ROUTINE,
        * and lets assume you have the formatting in the last row
        * and lets assume you want to have the same formatting for every row you add:
        * One option is this:
        * I am about to write to row n, SO...
        * cloneRow(n,n+1) // n is blank and has the formatting I want
        * put my vaues in row n
        * repeat until all rows added.
        */
        Sheet.prototype.cloneRow = function (source, destination) {
            if (source === 0)
                throw new Error('row 0 is not valid');
            if (destination === 0)
                throw new Error('row 0 is not valid');

            var rowData = functions({ fn: 'sheet-cloneRow', sheet: this.sheet, sourceRow: source - 1, destRow: destination - 1 }, true);
            return new Row(this.sheet, rowData, destination);
        };

        /**
        * protect() makes the cells on this sheet that have cell.setLock(true)
        * not reachable, without the password. In some cases you might just use a
        * known password because you are just trying to make the user's life easy,
        * but you don't want them to lock it out. So for example, with our language
        * exports, we make it public knowledge that the password is 'locked', that
        * way they can easily remember it if they need it.
        */
        Sheet.prototype.protect = function (password) {
            functions({ fn: 'sheet-protect', sheet: this.sheet, password: password }, true);
        };

        /**
        * unprotect() note that with xls/xlsx, you don't need a password to unlock/
        * unprotect a sheet. This is because the file is not encrypted, it is just
        * flagged for user convenience as locked.
        */
        Sheet.prototype.unprotect = function () {
            functions({ fn: 'sheet-unprotect', sheet: this.sheet }, true);
        };

        /**
        * createRow() lets you create a row that you can then edit cells on on this sheet.
        */
        Sheet.prototype.createRow = function (row) {
            if (row === 0)
                throw new Error('row 0 is not valid');

            var rowData = functions({ fn: 'sheet-createRow', sheet: this.sheet, row: row - 1 }, true);
            return new Row(this.sheet, rowData, row);
        };

        /**
        * A shortcut function for getting the value of a cell.
        */
        Sheet.prototype.getCellValue = function (row, column) {
            if (row === 0)
                throw new Error('row 0 is not valid');
            if (column === 0)
                throw new Error('column 0 is not valid');

            var rowObj = this.getRow(row);
            if (!rowObj.isExists())
                return null;

            var cellObj = rowObj.getCell(column);
            if (!cellObj.isExists())
                return null;

            return cellObj.getValue();
        };

        /**
        * setCellValue is a js wrapper we wrote for the excel access.
        * A shortcut function for setting the value of a cell.
        * It will create a row and column as needed.
        */
        Sheet.prototype.setCellValue = function (row, column, value) {
            if (row === 0)
                throw new Error('row 0 is not valid');
            if (column === 0)
                throw new Error('column 0 is not valid');

            var rowObj = this.getRow(row);
            if (!rowObj.isExists())
                rowObj.create();

            var cellObj = rowObj.getCell(column);
            if (!cellObj.isExists())
                cellObj.create();

            cellObj.setValue(value);
        };
        return Sheet;
    })();
    aps.Sheet = Sheet;

    var Row = (function () {
        function Row(sheet, row, index) {
            this.sheet = sheet;
            this.row = row;
            this.index = index;
        }
        Row.prototype.getCell = function (cell) {
            if (cell === 0)
                throw new Error('cell 0 is not valid');

            var cellData = functions({ fn: 'row-getCell', row: this.row, cell: cell - 1 }, true);
            if (cellData == null)
                return null;
            return new Cell(this.row, cellData, cell);
        };
        Row.prototype.getCellExists = function (cell) {
            if (cell === 0)
                throw new Error('cell 0 is not valid');

            return functions({ fn: 'row-getCellExists', row: this.row, cell: cell - 1 }, true);
        };

        /**
        * lets you know if a row exists that you used getRow() to get for example
        * If you now want to EDIT it, call either
        *     Row.create() or
        *     Sheet.createRow(####)
        */
        Row.prototype.isExists = function () {
            return !!this.row;
        };

        /**
        * Creates the row that is specified by this Row instance
        */
        Row.prototype.create = function () {
            if (this.isExists())
                return;
            this.row = functions({ fn: 'sheet-createRow', row: this.index, sheet: this.sheet }, true);
        };

        Row.prototype.createCell = function (cell) {
            if (cell === 0)
                throw new Error('cell 0 is not valid');

            var cellData = functions({ fn: 'row-createCell', row: this.row, cell: cell - 1 }, true);
            return new Cell(this.row, cellData, cell);
        };
        return Row;
    })();
    aps.Row = Row;

    var Cell = (function () {
        function Cell(row, cell, index) {
            this.row = row;
            this.cell = cell;
            this.index = index;
        }
        Cell.prototype.isExists = function () {
            return !!this.cell;
        };

        /**
        * Creates the cell that is specified by this Cell instance
        * note that sheet.setCellValue automatically creates the cell if needed
        */
        Cell.prototype.create = function () {
            if (this.isExists())
                return;
            this.cell = functions({ fn: 'row-createCell', row: this.row, cell: this.index - 1 }, true);
        };

        Cell.prototype.getValue = function () {
            return functions({ fn: 'cell-getValue', cell: this.cell }, true);
        };

        /**
        * The cell must exist before you setValue (Unlike the Sheet.setCellValue which will
        * create the cell if it doesn't exist.)
        */
        Cell.prototype.setValue = function (value) {
            functions({ fn: 'cell-setValue', cell: this.cell, value: value }, true);
        };

        /**
        * set the lock flag this cell to locked or unlocked.
        * Note you have to also call protect before it is locked from the UI perspective.
        */
        Cell.prototype.setLock = function (lock) {
            if (typeof lock === "undefined") { lock = true; }
            functions({ fn: 'cell-setLock', cell: this.cell, lock: lock }, true);
        };
        Cell.prototype.getLock = function () {
            return functions({ fn: 'cell-getLock', cell: this.cell }, true);
        };
        return Cell;
    })();
    aps.Cell = Cell;

    //#region Edge Function Inits
    var functions = edge.func({
        source: __dirname + "/.cs/Excel.cs",
        typeName: 'aps_excel_cs.Excel',
        methodName: 'Invoke',
        references: ['System.Data.dll', __dirname + '/.lib/NPOI.dll']
    });
})(aps || (aps = {}));

exports.Excel = new aps.Excel();
//# sourceMappingURL=excel.js.map
