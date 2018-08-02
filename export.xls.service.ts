import { Injectable } from '@angular/core';
import { Subject } from 'rxjs/Subject';

import { utils, write, writeFile, read, readFile, WorkBook, WorkSheet, ColInfo } from 'xlsx';

/**
 * Export array data type.
 */
export enum DataType
{
    Blob = 1,
    Data
}

/**
 * Export data array model.
 */
export class DataArrayModel
{
    constructor(title: string, data: any, type: DataType)
    {
        this.title = title;
        this.data = data;
        this.type = type;
    }

    public title: string;
    public data: any;
    public type: DataType;
}

/**
 * Export class service.
 * */
@Injectable()
export class ExportXlsService
{
    // Observable resource
    private exportObservable = new Subject<boolean>();

    // Observable streams for subscriber
    public exportSubscriber = this.exportObservable.asObservable();

    /**
     * Export content changed notification for subscriber.
     * @param doExport Flag indicating whether it is possible to export.
     */
    public exportChangedValue(doExport: boolean)
    {
        this.exportObservable.next(doExport);
    }

    /**
     * Export an array of JSON objects to an xls file.
     * @param fileName File name.
     * @param wsName WorkSheet name.
     * @param data JSON data to export.
     */
    public jsonToSheetExport(fileName: string, wsName: string, data: any)
    {
        // WorkBook creation
        let wb: WorkBook = { SheetNames: [], Sheets: {} };

        // Convert JSON dato to WorkSheet
        let ws: WorkSheet = utils.json_to_sheet(data);

        // Add WorkSheet to the list
        wb.SheetNames.push(wsName);

        // Set the WorkSheet
        wb.Sheets[wsName] = ws;
                
        // Write file
        writeFile(wb, fileName, { bookSST: true });
    }

    //public sheetAddJsonExport(fileName: string, wsName: string, ...dataArray: any[])
    //{
    //    // WorkBook creation
    //    let wb: WorkBook = { SheetNames: [], Sheets: {} };

    //    // Convert JSON dato to WorkSheet
    //    let ws: WorkSheet; // = utils.json_to_sheet(data);
    //    dataArray.forEach(data =>
    //    {
    //        ws = utils.sheet_add_json(ws, data)
    //    });

    //    // Add WorkSheet to the list
    //    wb.SheetNames.push(wsName);

    //    // Set the WorkSheet
    //    wb.Sheets[wsName] = ws;

    //    //// Write data into WorkBook
    //    //let wbout = write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });
    //    //// Save WorkBook on file
    //    //saveAs(new Blob([this.str2ab(wbout)], { type: 'application/octet-stream' }), fileName);

    //    // Write file
    //    writeFile(wb, fileName, { bookSST: true });
    //}

    /**
     * Convert a stream into an array buffer.
     * @param s Stream object to convert.
     */
    private str2ab(s)
    {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i !== s.length; ++i)
        {
            view[i] = s.charCodeAt(i) & 0xFF;
        };
        return buf;
    }

    /**
     * Convert a BLOB xls data into a JSON Object.
     * @param data BLOB data to convert.
     */
    public xlsBlobToJSON(data: any): Promise<any>
    {
        return new Promise((resolve, reject) =>
        {
            // Set the BLOB data reader
            let reader = new FileReader();

            // Processing readed data
            reader.onload = () =>
            {
                // Create the array buffer
                let u8 = new Uint8Array(reader.result);

                // Create the WorkBook reading the buffer
                let wb: WorkBook = read(u8, { type: 'array' });

                // Get the WorkSheet name
                let wsname: string = wb.SheetNames[0];

                // Get the WorkSheet n
                let ws: WorkSheet = wb.Sheets[wsname];

                // Get object properties
                let propArray = this.getCellsReference(ws);

                // Get last range element
                let lastElem = propArray[propArray.length - 1];

                // Set the range to read all the cell
                let range: string = "A1:" + lastElem;

                // Convert the sheet into a JSON Object
                let xlsData = utils.sheet_to_json(ws, { range: range, header: 1 });

                // Resolve the promise
                resolve(xlsData);
            };

            // Read blob array buffer
            reader.readAsArrayBuffer(data);
        });
    }    

    /**
     * Exrpot one or more data (xls BLOB data or simple string data) into one xls file (merge the data if there are more than one).
     * @param fileName Xls file name.
     * @param colWidth Column width to set for the result sheet.
     * @param dataArray Data array to export.
     */
    public exportDataArrayToXls(fileName: string, colWidth: number, ...dataArray: DataArrayModel[]): Promise<any>
    {
        return new Promise((resolve, reject) =>
        {
            // Output WorkBook
            let wbout: WorkBook;

            // Ouput WorkSheet
            let wsout: WorkSheet;
            
            let arrayLength = dataArray.length;

            // Processing the data array
            for (let i = 0; i < arrayLength; i++)
            {
                // Processing the BLOB data
                if (dataArray[i].type == DataType.Blob)
                {
                    // Set the BLOB data reader
                    let reader = new FileReader();

                    // Processing readed data
                    reader.onload = () =>
                    {
                        // Create the array buffer
                        let u8 = new Uint8Array(reader.result);

                        // Create the WorkBook
                        let wb: WorkBook = read(u8, { type: 'array', cellDates: true, cellStyles: true });

                        // Create the output WorkBook
                        if (this.testIsUndefined(wbout)) wbout = wb;

                        // Get the WorkSheet name
                        let wsname: string = wb.SheetNames[0];

                        // Create a WorkSheet with the title
                        let ws: WorkSheet = utils.aoa_to_sheet([[dataArray[i].title]]);

                        // Refresh range before sheet merge (it's needed)
                        ws["!ref"] = this.getRange(ws);

                        if (this.testIsUndefined(wsout)) wsout = ws; // Create the output WorkSheet
                        else wsout = this.mergeWorkSheet(wsout, ws, 2); // Merge WorkSheets

                        // Get the WorkSheet with data
                        ws = wb.Sheets[wsname];

                        // Refresh range before sheet merge
                        ws["!ref"] = this.getRange(ws);

                        // Merge WorkSheets
                        wsout = this.mergeWorkSheet(wsout, ws, 1);

                        if (i == arrayLength - 1)
                        {
                            // Set the column widtt
                            wsout['!cols'] = this.setColWidth(wsout, colWidth);

                            // Set the output sheet
                            wbout.Sheets[wsname] = wsout;

                            // Write the xls file
                            writeFile(wbout, fileName, { bookSST: true, cellDates: true, cellStyles: true });

                            // Resolve the promise
                            resolve();
                        }                        
                    };

                    // Read blob array buffer
                    reader.readAsArrayBuffer(dataArray[i].data);
                }

                // Processing the other data
                if (dataArray[i].type == DataType.Data)
                {
                    if (this.testIsUndefined(wbout))
                    {
                        // Set the output WorkBook
                        wbout = { SheetNames: <string[]>[], Sheets: {} };
                        wbout = utils.book_new();

                        // Add data to the output sheet
                        wsout = utils.aoa_to_sheet(dataArray[i].data);

                        // Add the sheet to the output WorkBook
                        let ws_name = "Sheet";
                        wbout.SheetNames.push(ws_name);
                        wbout.Sheets[ws_name] = wsout;
                    }
                    else
                    {
                        // Create a new WorkSheet to be merged with the output
                        let ws: WorkSheet = utils.aoa_to_sheet(dataArray[i].data);

                        // Merge WorkSheets
                        wsout = this.mergeWorkSheet(wsout, ws, 2);
                    }                    
                }
            }
        });
    }

    /**
     * Set the column width for all the column with data.
     * @param ws WorkSheet.
     * @param width Width to set.
     */
    private setColWidth(ws: WorkSheet, width: number): ColInfo[]
    {
        // Get the WorkSheet properties
        let propArray = this.getCellsReference(ws);

        // Get the column array with data
        let colArray = propArray.map(elem => elem.substring(0, 1));
        
        // Get unique column array
        let uniqueColArray = colArray.filter((obj, pos, arr) =>
        {
            return arr.map(mapObj => mapObj[0]).indexOf(obj[0]) === pos;
        });

        // Set the width structure
        let result: ColInfo[] = [];
        uniqueColArray.forEach(col =>
        {
            let ci: ColInfo = { wch: width };
            result.push(ci);
        });

        // Return the column width structure
        return result;
    }

    /**
     * Returns a string array that represents the coordinates of all the cell with data.
     * @param ws WorkSheet.
     */
    private getCellsReference(ws: WorkSheet): string[]
    {
        // Get object properties
        let propArray = Object.getOwnPropertyNames(ws);

        // Exlude the "!" properties
        propArray = propArray.filter(f => !f.match("!"));

        return propArray;
    }

    /**
     * Get the correct range of the WorkSheet.
     * @param ws WorkSheet.
     */
    private getRange(ws: WorkSheet): string
    {
        // Get object properties
        let propArray = this.getCellsReference(ws);

        // Get last range element
        let lastElem = propArray[propArray.length - 1];

        let progArray = propArray.map(elem => Number(elem.substring(1)));
        let maxProg = progArray.sort(this.sortNumber)[progArray.length - 1];
        // Set the range to write all the cell
        let range: string = "A1:" + lastElem.substring(0, 1) + maxProg;

        return range;
    }

    /**
     * Returns the merged range of two WorkSheet.
     * @param ws1 First WorkSheet.
     * @param ws2 Second WorkSheet.
     */
    private mergeRange(ws1: WorkSheet, ws2: WorkSheet): string
    {
        // Get cell array of first WorkSheet
        let cellArray1 = this.getCellsReference(ws1);

        // Get cell array of second WorkSheet
        let cellArray2 = this.getCellsReference(ws2);

        // Select the max row of first WorkSheet
        let rowArray1 = cellArray1.map(elem => Number(elem.substring(1)));        
        let maxRow1 = rowArray1.sort(this.sortNumber)[rowArray1.length - 1];

        // Select the max row of second WorkSheet
        let rowArray2 = cellArray2.map(elem => Number(elem.substring(1)));        
        let maxRow2 = rowArray2.sort(this.sortNumber)[rowArray2.length - 1];

        // Select the max column of first WorkSheet
        let colArray1 = cellArray1.map(elem => elem.substring(0, 1));
        let maxCol1 = colArray1.sort(this.sortString)[colArray1.length - 1];

        // Select the max column of second WorkSheet
        let colArray2 = cellArray2.map(elem => elem.substring(0, 1));        
        let maxCol2 = colArray2.sort(this.sortString)[colArray2.length - 1];

        // Get last range element
        let lastCol = (maxCol1 > maxCol2) ? maxCol1 : maxCol2;
        let lastRow = ((maxRow1 > maxRow2) ? maxRow1 : maxRow2).toString();

        // Set the merged range
        let range: string = "A1:" + lastCol + lastRow.toString();

        // Return the merged range
        return range;
    }

    /**
     * Sort number funtion. Returns a number that indicates if the first value in greater, equal or lower than the second.
     * @param a First number.
     * @param b Second number.
     */
    private sortNumber(a: number, b: number): number
    {
        return a - b;
    }

    /**
     * Sort string funtion. Returns a number that indicates if the first string in greater, equal or lower than the second.
     * @param a First string.
     * @param b Second string.
     */
    private sortString(a, b)
    {
        if (a < b) return -1;
        if (a > b) return 1;
        return 0;
    }

    /**
     * Merges two WorkSheet.
     * @param ws1 First WorkSheet.
     * @param ws2 Second WorkSheet.
     * @param offset Offset to be applied between the two merged WorkSheet.
     */
    private mergeWorkSheet(ws1: WorkSheet, ws2: WorkSheet, offset:number): WorkSheet
    {
        // Set the output WorkSheet with the frist one
        let ws: WorkSheet = ws1;

        // Get the first row count
        let ws1RowsCount = this.getRowsCount(ws1);

        // Set the offset
        let rowOffset: number = ws1RowsCount + offset;

        // Get object properties
        let cellsArrayWs2 = this.getCellsReference(ws2);

        cellsArrayWs2.forEach(elem =>
        {
            let newProg: number = rowOffset + Number(elem.substring(1));
            let newProp: string = elem.substring(0, 1) + newProg.toString();;
            ws[newProp] = ws2[elem];
        });

        // Update range info
        ws["!ref"] = this.mergeRange(ws, ws2);

        // Clone objects
        let merges1 = Object.assign([], ws1["!merges"]);
        let merges2 = Object.assign([], ws2["!merges"]);

        // Update merge info
        ws["!merges"] = this.updateMerge(merges1, merges2, rowOffset);

        // Return the merged result
        return ws;
    }

    /**
     * Returns the WorkSheet row count decoding range info..
     * @param ws WorkSheet.
     */
    private getRowsCount(ws: WorkSheet)
    {
        var range = utils.decode_range(ws['!ref']);
        let count = range.e.r - range.s.r;
        return count;
    }

    /**
     * Updates the merge info of the two merge structure.
     * @param merges1 Merge info of the first WorkSheet.
     * @param merges2 Merge info of the second WorkSheet.
     * @param offset Offset.
     */
    private updateMerge(mergesWs1: any[], mergesWs2: any[], offset: number): any
    {
        let mergeResult: any[] = [];

        mergesWs1.forEach(merge =>
        {
            mergeResult.push(merge);
        });

        mergesWs2.forEach(merge =>
        {
            merge.s.r = Number(merge.s.r) + offset;
            merge.e.r = Number(merge.e.r) + offset;
            mergeResult.push(merge);
        });

        return mergeResult;
    }

    private testIsUndefined(value: any)
    {
        return typeof value === 'undefined';
    }

    /*
    private sheet2arr(sheet: WorkSheet)
    {
        var result = [];
        var row;
        var rowNum;
        var colNum;
        var range = utils.decode_range(sheet['!ref']);

        for (rowNum = range.s.r; rowNum <= range.e.r; rowNum++)
        {
            row = [];
            for (colNum = range.s.c; colNum <= range.e.c; colNum++)
            {
                var nextCell = sheet[
                    utils.encode_cell({ r: rowNum, c: colNum })
                ];
                if (typeof nextCell === 'undefined')
                {
                    row.push(void 0);
                } else row.push(nextCell.w);
            }
            result.push(row);
        }

        return result;
    };*/
}