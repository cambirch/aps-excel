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
            functions({ fn: 'workbook-save', path: path }, true);
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
        Sheet.prototype.getRow = function (row) {
            var rowData = functions({ fn: 'sheet-getRow', sheet: this.sheet, row: row }, true);
            if (rowData == null)
                return null;
            return new Row(this.sheet, rowData, row);
        };
        Sheet.prototype.getRowExists = function (row) {
            return functions({ fn: 'sheet-getRowExists', sheet: this.sheet, row: row }, true);
        };

        Sheet.prototype.getRowCount = function () {
            return functions({ fn: 'sheet-getRowCount', sheet: this.sheet }, true);
        };
        Sheet.prototype.cloneRow = function (source, destination) {
            var rowData = functions({ fn: 'sheet-cloneRow', sheet: this.sheet, sourceRow: source, destRow: destination }, true);
            return new Row(this.sheet, rowData, destination);
        };

        Sheet.prototype.protect = function (password) {
            functions({ fn: 'sheet-protect', sheet: this.sheet, password: password }, true);
        };
        Sheet.prototype.unprotect = function () {
            functions({ fn: 'sheet-unprotect', sheet: this.sheet }, true);
        };

        Sheet.prototype.createRow = function (row) {
            var rowData = functions({ fn: 'sheet-createRow', sheet: this.sheet, row: row }, true);
            return new Row(this.sheet, rowData, row);
        };

        /**
        * A shortcut function for getting the value of a cell.
        */
        Sheet.prototype.getCellValue = function (row, column) {
            var rowObj = this.getRow(row);
            if (!rowObj.isExists())
                return null;

            var cellObj = rowObj.getCell(column);
            if (!cellObj.isExists())
                return null;

            return cellObj.getValue();
        };

        /**
        * A shortcut function for setting the value of a cell.  Will create row/column as needed.
        */
        Sheet.prototype.setCellValue = function (row, column, value) {
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
            var cellData = functions({ fn: 'row-getCell', row: this.row, cell: cell }, true);
            if (cellData == null)
                return null;
            return new Cell(this.row, cellData, cell);
        };
        Row.prototype.getCellExists = function (cell) {
            return functions({ fn: 'row-getCellExists', row: this.row, cell: cell }, true);
        };

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
            var cellData = functions({ fn: 'row-createCell', row: this.row, cell: cell }, true);
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
        */
        Cell.prototype.create = function () {
            if (this.isExists())
                return;
            this.cell = functions({ fn: 'row-createCell', row: this.row, cell: this.index }, true);
        };

        Cell.prototype.getValue = function () {
            return functions({ fn: 'cell-getValue', cell: this.cell }, true);
        };
        Cell.prototype.setValue = function (value) {
            functions({ fn: 'cell-setValue', cell: this.cell, value: value }, true);
        };

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
