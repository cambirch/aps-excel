aps-excel v0.2.0
=======================================

A node.js library that uses Edge.js and NPOI to manipulate Excel files


### Installing for use ###

Please note that this module is **not** published to NPM nor will it be anytime soon.

Installation has been simplified in this version.

    npm install camMCC/aps-excel --save

**NOTE:** The above command is CASE SENSITIVE.


## Using ##

    var excel = require('aps-excel').Excel;
    
    // Load a workbook
    var workbook = excel.load('filename.xls');
    
    // Get the first sheet
    var sheet = workbook.getSheetAt(0);

    // Get a sheet by name
    var sheet = workbook.getSheet('Sheet1');
    
    // Print the value of the first cell (A1)
    console.log(sheet.getCellValue(1, 1);

    // Set the value of the first cell (A1)
    sheet.setCellValue(1, 1, 'Awesome!');


## API ##
API description provided in Typescript syntax for ease.

    Excel
        .load(path: string): Workbook

    Workbook
        .save(path: string): void
        .getSheetAt(index: number): Sheet
        .getSheet(name: string): Sheet

    Sheet
        /**
        * getRow returns a row object that you can read or edit
        * note there is a flag for whether the row previously existed on not
        * so it is safe to call getRow without first calling getRowExists
        * if getRow returns you a row that doesn't exist you will need to call either:
        *     Row.create() or 
        *     Sheet.createRow(####)
        **/
	getRow(row: number): Row
        /**
        * getRowExists lets you know whether a specific row exists.
        * if you call it with a row number greater than getRowCount()
        * you should expect it to return false.
        * if you call it with a row number less or equal to getRowCount()
        * it will let you know whether that row exists.
        **/
        getRowExists(row: number): boolean
        /**
        * getRowCount() tells you NOT how 'many' rows exist, but rather,
        * what is the last row that MIGHT exist. So if getRowCount() returns 99
        * it is possible that the only row existing is row 99, the first 98 may
        * be blank and non-existing. This makes it excellent when reading through
        * a file to know when to stop looking for more rows.
        * so, you might say for row 12 to getRowCount() - read the row and see if
        * is any data in that row that I care about.
        **/
        getRowCount(): number
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
        **/
        cloneRow(source: number, destination: number): Row
        /**
        * protect() makes the cells on this sheet that have cell.setLock(true)
        * not reachable, without the password. In some cases you might just use a
        * known password because you are just trying to make the user's life easy,
        * but you don't want them to lock it out. So for example, with our language
        * exports, we make it public knowledge that the password is 'locked', that 
        * way they can easily remember it if they need it.
        **/
        protect(password: string): void
        /**
        * unprotect() note that with xls/xlsx, you don't need a password to unlock/
        * unprotect a sheet. This is because the file is not encrypted, it is just
        * flagged for user convenience as locked.
        **/
        unprotect(): void
        /**
        * createRow() lets you create a row that you can then edit cells on on this sheet.
        **/
        createRow(row: number): Row
        /**
         * A shortcut function for getting the value of a cell.
         */
        getCellValue(row: number, column: number): any
        /**
         * setCellValue is a js wrapper we wrote for the excel access.
         * A shortcut function for setting the value of a cell.  
         * It will create a row and column as needed.
         */
        setCellValue(row: number, column: number, value: any): void

    Row
        /**
        * 
        **/
        getCell(cell: number): Cell
        /**
        * 
        **/
        getCellExists(cell: number): boolean
        /**
        * lets you know if a row exists that you used getRow() to get for example
        * If you now want to EDIT it, call either
        *     Row.create() or 
        *     Sheet.createRow(####)
        **/
        isExists(): boolean
        /**
         * Creates the row that is specified by this Row instance
         */
        create(): void
        createCell(cell: number): Cell

    Cell
        /**
        * 
        **/
        isExists(): boolean
        /**
         * Creates the cell that is specified by this Cell instance
         * note that sheet.setCellValue automatically creates the cell if needed
        **/
        create(): void
        /**
        *
        **/
        getValue(): any
        /**
        * The cell must exist before you setValue (Unlike the Sheet.setCellValue which will
        * create the cell if it doesn't exist.)
        **/
        setValue(value: any): void
        /**
        * set the lock flag this cell to locked or unlocked.
        * Note you have to also call protect before it is locked from the UI perspective.
        **/
        setLock(lock: boolean = true): void
        /**
        * 
        **/
        getLock(): boolean


