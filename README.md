# js-xlsx-exportxlsservice
Angular service that through js-xlsx export data to excel. It can merge more than one xls file.

> This service require the insallation of js-xlsx package.

## Import
```
// Import the service
import { ExportXlsService, DataArrayModel, DataType } from './services/export.xls.service';
```

## Example
```
// Simple data 
let xls1 = new DataArrayModel('First table title', [['A1', 'B1'], ['A2', 'B2']], DataType.Data);
let xls2 = new DataArrayModel('Second table title', [['A3', 'B3'], ['A4', 'B4']], DataType.Data);

// Set the column width for the output file
let colWidth: number = 15;

// Merge more data into one xlsx file
this.exportXlsService.exportDataArrayToXls('Example.xlsx', colWidth, xls1, xls2).then(result => 
{
  console.log('Merge completed');
});
```

The output will be:
```
First table title	
A1	B1
A2	B2
	
Second table title	
A3	B3
A4	B4
```


If we have one or more [BLOB](https://en.wikipedia.org/wiki/Binary_large_object) that represents the xls files, we can also merge them into a single file.

```
// Blob data
let xlsBlob1 = new DataArrayModel('First file title', e.data, DataType.Blob);
let xlsBlob2 = new DataArrayModel('Second file title', e.data, DataType.Blob);

// Set the column width for the output file
let colWidth: number = 15;

// Merge more data into one xlsx file
this.exportXlsService.exportDataArrayToXls('Example.xlsx', colWidth, xls1, xls2).then(result => 
{
  console.log('Merge completed');
});
```

Supposing that the first file have the content:

```
A1	B1
A2	B2
```
And the second file have the content:

```
A3	B3
A4	B4
```

The output will be:

```
First file title	
A1	B1
A2	B2
	
Second file title	
A3	B3
A4	B4
```

> If the BLOBs contains merge informations for the cell, this will be updated to make sure that they are valid even after the union of the files.
