///<reference path='node.d.ts'/>

module aps {
    var edge = require('edge');


    export class Excel {
        constructor() {
        }

        public load(path: string): Workbook {
            var workbook = functions({ fn: 'excel-load', path: path }, true);
            return new Workbook(workbook);
        }
    }

    export class Workbook {
        constructor(private workbook: any) {
        }

        public save(path: string): void {
            functions({ fn: 'workbook-save', path: path }, true);
        }

        public getSheetAt(index: number): Sheet {
            var sheet = functions({ fn: 'workbook-getSheetAt', index: index, workbook: this.workbook }, true);
            return new Sheet(sheet);
        }
        public getSheet(name: string): Sheet {
            var sheet = functions({ fn: 'workbook-getSheet', name: name, workbook: this.workbook }, true);
            return new Sheet(sheet);
        }
        
    }

    export class Sheet {
        constructor(private sheet: any) {
        }

        public getRow(row: number): Row {
            var rowData = functions({ fn: 'sheet-getRow', sheet: this.sheet, row: row }, true);
            if (rowData == null) return null;
            return new Row(this.sheet, rowData, row);
        }
        public getRowExists(row: number): boolean {
            return <boolean>functions({ fn: 'sheet-getRowExists', sheet: this.sheet, row: row }, true);
        }

        public getRowCount(): number {
            return functions({ fn: 'sheet-getRowCount', sheet: this.sheet }, true);
        }
        public cloneRow(source: number, destination: number): Row {
            var rowData = functions({ fn: 'sheet-cloneRow', sheet: this.sheet, sourceRow: source, destRow: destination }, true);
            return new Row(this.sheet, rowData, destination);
        }

        public protect(password: string): void {
            functions({ fn: 'sheet-protect', sheet: this.sheet, password: password }, true);
        }
        public unprotect(): void {
            functions({ fn: 'sheet-unprotect', sheet: this.sheet }, true);
        }

        public createRow(row: number): Row {
            var rowData = functions({ fn: 'sheet-createRow', sheet: this.sheet, row: row }, true);
            return new Row(this.sheet, rowData, row);
        }

        /**
         * A shortcut function for getting the value of a cell.
         */
        public getCellValue(row: number, column: number): any {
            var rowObj = this.getRow(row);
            if (!rowObj.isExists()) return null;
            
            var cellObj = rowObj.getCell(column);
            if (!cellObj.isExists()) return null;

            return cellObj.getValue();
        }
        /**
         * A shortcut function for setting the value of a cell.  Will create row/column as needed.
         */
        public setCellValue(row: number, column: number, value: any): void {
            var rowObj = this.getRow(row);
            if (!rowObj.isExists()) rowObj.create();

            var cellObj = rowObj.getCell(column);
            if (!cellObj.isExists()) cellObj.create();

            cellObj.setValue(value);
        }
    }

    export class Row {
        constructor(private sheet: any, private row: any, private index: number) {
        }
        
        public getCell(cell: number): Cell {
            var cellData = functions({ fn: 'row-getCell', row: this.row, cell: cell }, true);
            if (cellData == null) return null;
            return new Cell(this.row, cellData, cell);
        }
        public getCellExists(cell: number): boolean {
            return functions({ fn: 'row-getCellExists', row: this.row, cell: cell }, true);
        }

        public isExists(): bool {
            return !!this.row;
        }

        /**
         * Creates the row that is specified by this Row instance
         */
        public create(): void {
            if (this.isExists()) return;
            this.row = functions({ fn: 'sheet-createRow', row: this.index, sheet: this.sheet }, true);
        }

        public createCell(cell: number): Cell {
            var cellData = functions({ fn: 'row-createCell', row: this.row, cell: cell }, true);
            return new Cell(this.row, cellData, cell);
        }
    }

    export class Cell {
        constructor(private row: any, private cell: any, private index: number) {
        }

        public isExists(): bool {
            return !!this.cell;
        }

        /**
         * Creates the cell that is specified by this Cell instance
         */
        public create(): void {
            if (this.isExists()) return;
            this.cell = functions({ fn: 'row-createCell', row: this.row, cell: this.index }, true);
        }

        public getValue(): any {
            return functions({ fn: 'cell-getValue', cell: this.cell }, true);
        }
        public setValue(value: any): void {
            functions({ fn: 'cell-setValue', cell: this.cell, value: value }, true);
        }
        
        public setLock(lock: boolean = true): void {
            functions({ fn: 'cell-setLock', cell: this.cell, lock: lock }, true);
        }
        public getLock(): boolean {
            return <boolean>functions({ fn: 'cell-getLock', cell: this.cell }, true);
        }
    }

    //#region Edge Function Inits

    var functions = edge.func({
        source: __dirname + "/.cs/Excel.cs",
        typeName: 'aps_excel_cs.Excel',
        methodName: 'Invoke',
        references: ['System.Data.dll', __dirname + '/.lib/NPOI.dll']
    });

    //#endregion

}

exports.Excel = new aps.Excel();