# node-red-contrib-excelsheets

## Overview

Features the ability to save data to multiple sheets. I found it wasn't supported in some of the other nodes already available.

To see an example, install the package, then in your Node-RED environment go to Import > Examples > node-red-contrib-excelsheets

Basic **msg** format is,

```
payload: [
    {
        header: {
            col1: "Column 1",
            col2: "Column 2",
            ...
        } ,
        items: [
            {
                col1: "data",
                col2: "data",
                ...
            }
        ] ,
        sheetName: "Sheet 1"
    } ,
    {header, items, sheetName} ,
    {header, items, sheetName} ,
    {header, items, sheetName} ,
    ...
] ,
filepath: "output.xlsx"
```

## Change Log
* 1.0.5 - Updated docs to refer to correct msg property: filepath (not filename, )
* 1.0.4 - Additional Example "basic w server". Cleanup and documentation.
* 1.0.1 ~ 3 - Cleanup and documentation.
* 1.0.0 - Original version, forked from "node-red-contrib-excel" and swapped out a different dependency.

## TODO
- [ ] Give richer status feedback / error feedback
- [ ] Digging deeper into the dependencies, there's a really powerful package called "xlsx" [<npmjs.com>](https://www.npmjs.com/package/xlsx) that has a rich base of features, I would like to create a node that uses everything available.
- [ ] Come up with a few examples that do something more interesting
