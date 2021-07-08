# node-red-contrib-excelsheets

## Overview

Features the ability to save data to multiple sheets. I found it wasn't supported in some of the other nodes already available.

To see an example, install the package, then in your Node-RED environment go to Import > Examples > node-red-contrib-excelsheets

There are **two methods** of structuring your incoming **msg**,

Standard format is this:

```
payload: [
    {
        header: {
            col1: "Column 1",
            col2: "Column 2"
        } ,
        items: [
            {
                col1: "data",
                col2: "data"
            }
        ] ,
        sheetName: "Sheet 1"
    } ,
    {header, items, sheetName} ,
    {header, items, sheetName} ,
    {header, items, sheetName} 
] ,
filepath: "output.xlsx"
```

Flexible format is this:

```
payload: {
    sheet1: [
        header1: "data",
        header2: "data",
        header3:{
            sub1: "data",
            sub2: "data"
        }
    ],
    sheet2: [
        header1: "data",
        header2: "data"
    ],
    sheet3: [
        header1: "data",
        header2: "data"
    ]
} ,
filepath: "output.xlsx"
```

When using the "Flexible" format, the headers are automatically determined from the key names. If the data in the sheet is nested in multiple objects, the path to that data is concatenated with hyphens recursively.

For the example above: sheet1, will have four headers: "header1", "header2", "header3-sub1", and "header3-sub2". sheet2 and sheet3 will just have "header1" and "header2". Furthermore, the headers do **not** need to be consistent, even in the same sheet. Unique headers are added to the header list when they are encountered.

The output of the node will always transform the data into the "Standard" format.

## Change Log
* 1.1.0 - Created "Flexible" input option that parses object keys as sheet names and recursively parses multi-level objects in the row data
* 1.0.5 - Updated docs to refer to correct msg property: filepath (not filename, )
* 1.0.4 - Additional Example "basic w server". Cleanup and documentation.
* 1.0.1 ~ 3 - Cleanup and documentation.
* 1.0.0 - Original version, forked from "node-red-contrib-excel" and swapped out a different dependency.

## TODO
- [ ] Confirm where files are saved if no path is given
- RE: above, make it default to $HOME dir if it doesn't already

Digging deeper into the dependencies, there's a really powerful package called "xlsx" [<npmjs.com>](https://www.npmjs.com/package/xlsx) that has a very rich base of features. It would be nice to create a node that uses everything available, maybe someday.
