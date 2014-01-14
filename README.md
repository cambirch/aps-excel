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

> var excel = require('aps-excel').Excel;
> 
> var workbook = excel.load('filename');
> 
> var sheet = workbook.getSheetAt(0);
> 
> console.log(sheet.getCellValue(0, 0);



