aps-excel
=========

A node.js library that uses Edge.js and NPOI to manipulate Excel files


### Installing for use ###

This module is **not** published to NPM nor will it be anytime soon.  This is important at this point!

Using this module requires a few simple steps in order to install it as a natural NPM module.

- Clone the repository into a local directory
- Open an Administrator command prompt
- Browse to the directory that the repository is cloned into
- Use the cmd: 
    > `npm link`
- Browse to the directory that you are building a node project that requires the aps-excel module
- Use the cmd: 
	> `npm link aps-excel`
- Use the module as you would any other 
	> `require('aps-excel')`

## Using ##

    var excel = require('aps-excel').Excel;
    var workbook = excel.load('filename');
    var sheet = workbook.getSheetAt(0);
    console.log(sheet.getCellValue(0, 0);


## API ##
API description provided in Typescript syntax for ease.

    Excel
        .load(path: string): Workbook

    Workbook
        .save(path: string): void
        .getSheetAt(index: number): Sheet
        .getSheet(name: string): Sheet

    Sheet
        getRow(row: number): Row
        getRowExists(row: number): boolean
        getRowCount(): number
        cloneRow(source: number, destination: number): Row
        protect(password: string): void
        unprotect(): void
        createRow(row: number): Row
        /**
         * A shortcut function for getting the value of a cell.
         */
        getCellValue(row: number, column: number): any
        /**
         * A shortcut function for setting the value of a cell.  Will create row/column as needed.
         */
        setCellValue(row: number, column: number, value: any): void

    Row
        getCell(cell: number): Cell
        getCellExists(cell: number): boolean
        isExists(): boolean
        /**
         * Creates the row that is specified by this Row instance
         */
        create(): void
        createCell(cell: number): Cell

    Cell
        isExists(): boolean
        /**
         * Creates the cell that is specified by this Cell instance
         */
        create(): void
        getValue(): any
        setValue(value: any): void
        setLock(lock: boolean = true): void
        getLock(): boolean


