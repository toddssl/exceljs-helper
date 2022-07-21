
const ExcelJS = require('exceljs');

var exportExcel = function (columns = [], dataJson = [], headerArr = [], config = {}) {

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(config.sheetName);

    console.log('helper input dataJson = ', dataJson);
    console.log('helper input columns = ', columns);
    console.log('helper input headerArr = ', headerArr);
    console.log('helper input config = ', config);

    let commonConfig = {
        style: {
            font: { name: 'Century Gothic', size: 12 },
            // alignment: {
            //     vertical: 'middle',
            //     horizontal: 'center',
            // },
            // // border: {
            //     bottom: {style:'double', color: {argb:'000000'}},
            // }
            numFmt: '#,##',
        },
        width: 20,
    }

    columns.forEach(function(data){
        // Object.assign(data, commonConfig);
        // if (data.width == undefined) {
        //     data.width = 20;
        // } else {
        //     data.width = parseInt(data.width / 5);
        // }
        if (data.width == undefined) {
            data.width = 20;
        }

        if (data.style == undefined) {
            Object.assign(data, commonConfig);
        }
        
        if (data.style.font == undefined) {
            data.style.font = commonConfig.style.font;
        }

        if (data.dataType == 'date') {
            let cloneStyle = JSON.parse(JSON.stringify(data.style))
            data.style = Object.assign(cloneStyle, { numFmt: 'yyyy/mm/dd' });
        }
        if (data.dataType == 'percent') {
            let cloneStyle = JSON.parse(JSON.stringify(data.style))
            data.style = Object.assign(cloneStyle, { numFmt: '#,#0.00%' });
        }
        if (data.dataType == 'float') {
            let cloneStyle = JSON.parse(JSON.stringify(data.style))
            data.style = Object.assign(cloneStyle, { numFmt: data.formatExcel });
        }
    });

    // console.log('columns = ', columns);

    worksheet.columns = columns;

    var addRow = function (dataJson, startOutlineIndex = 0, startOutlineLevel = -1, outlineResult = []) {

        for (var key in dataJson) {
            var item = dataJson[key];
            // console.log('item = ', item);
            startOutlineIndex = _addRow(worksheet, item, columns, startOutlineIndex, startOutlineLevel + 1, outlineResult);
        }

        return outlineResult;
    }

    var _addRow = function (worksheet, item, columns, outlineIndex = 0, outlineLevel = 0, outlineResult = []) {

        if (item.children) {

            item.children.forEach(function(itemChildren){

                outlineIndex = _addRow(worksheet, itemChildren, columns, outlineIndex, outlineLevel + 1, outlineResult);

                // console.log('set outlineIndex = ', outlineIndex, ' outlineLevel = ', outlineLevel, ' itemChildren = ', itemChildren);

                if (outlineResult[outlineLevel] == undefined) {
                    outlineResult[outlineLevel] = [];
                }
                outlineResult[outlineLevel].push(outlineIndex);
            });
        }

        var rowData = {};
        for (var key in columns) {

            var columnObj = columns[key];
            let displayValue = item[columnObj.key];

            // console.log('--columnObj = ', columnObj, ' displayValue = ', displayValue, ' type = ', typeof(displayValue));

            if (typeof (displayValue) == 'number' && !Number.isInteger(displayValue)) {
                let displayValueNum = Number(displayValue);
                rowData[columnObj.key] = parseFloat(displayValueNum.toFixed(2));
            } else {
                rowData[columnObj.key] = displayValue;
            }

            if (columnObj.dataType == 'date' && displayValue != undefined) {
                let tmpDate = new Date(displayValue);
                rowData[columnObj.key] = new Date(Date.UTC(tmpDate.getFullYear(), tmpDate.getMonth(), tmpDate.getDate()));
            }

            if (columnObj.dataType == 'percent' && displayValue != undefined) {
                
                // let displayValueNum = Number(displayValue) / 100;
                let displayValueNum = Number(displayValue);
                // console.log('!!! displayValueNum = ', displayValueNum, ' displayValue = ', displayValue);
                if (displayValueNum == 'Infinity') {
                    displayValueNum = 0;
                }
                // rowData[columnObj.key] = parseFloat(displayValueNum.toFixed(2));
                rowData[columnObj.key] = parseFloat(displayValueNum);
            }

            if (columnObj.formatExcel && typeof columnObj.formatExcel === 'function') {
                rowData[columnObj.key] = columnObj.formatExcel(displayValue);
            }
        }

        outlineIndex += 1;
        // console.log('!! add one outlineIndex = ', outlineIndex, ' outlineLevel = ', outlineLevel, ' row = ', rowData['grp'], ' rowData = ', rowData);
        worksheet.addRow(rowData);

        return outlineIndex;
    }

    // insert row and get row outline position
    var outlineRow = addRow(dataJson, headerArr.length);

    // get column outline position
    var outlineColumn = [];
    headerArr.slice().forEach(function(data, index){

        // console.log('outline header arr data = ', data, ' index = ', index);

        if (index < headerArr.length - 1) {

            data.slice().forEach(function (data, indexChild) {

                if (outlineColumn[index] == undefined) {
                    outlineColumn[index] = [];
                }
                
                // console.log('outline header arr child data = ', data, ' indexChild = ', indexChild);
                if (data == '' && indexChild > 0) {
                    outlineColumn[index].push(indexChild + 1);
                }
            });
        }
    });

    // clone a fill up header array
    var headeFillupArr = [];
    headerArr.slice().forEach(function(data, index){

        // console.log('header index = ', index, ' data = ', data);

        headeFillupArr[index] = [];

        let keepValue = '';
        for (let indexChild = data.length - 1; indexChild >= 0; indexChild--) {
            
            const dataChild = data[indexChild];
            // console.log('header item indexChild = ', indexChild, ' dataChild = ', dataChild, ' keepValue = ', keepValue);
            
            if (indexChild == 0) {
                headeFillupArr[index].unshift(dataChild);                
                return;
            }
            if (keepValue == '') {
                keepValue = dataChild;
            }
            if (dataChild != '' && keepValue != dataChild) {
                keepValue = dataChild;
            }
            headeFillupArr[index].unshift(keepValue);
        }
    });
    console.log('headerArr = ', headerArr);
    console.log('headeFillupArr = ', headeFillupArr);
    
    // print header
    let spliceFirstRowAmount = 1; // remove origin header
    headeFillupArr.slice().reverse().forEach(function(data, index){
        worksheet.spliceRows(1, spliceFirstRowAmount, data);
        spliceFirstRowAmount = 0;
    });
    
    console.log('outlineRow = ', outlineRow);
    console.log('outlineColumn = ', outlineColumn);

    // render outline
    outlineRow.slice().forEach(function (data, index) {
        data.forEach(function (child, indexChild) {
            worksheet.getRow(child).outlineLevel = index + 1;
            worksheet.getRow(child).hidden = true;
        })
    });
    outlineColumn.slice().forEach(function (data, index) {
        data.forEach(function (child, indexChild) {
            worksheet.getColumn(child).outlineLevel = index + 1;
            worksheet.getColumn(child).hidden = true;
        })
    });

    // header background color
    headerArr.slice().forEach(function(data, index){
        var headerRow = worksheet.getRow(index + 1);
        headerRow.eachCell(function(cell, rowNumber) {
            cell.style.font = {
                size: '12',
                bold: true,
                name: 'Century Gothic',
            }
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'DCE6F1' },
            }
        });    
    });

    var buff = workbook.xlsx.writeBuffer().then(function (data) {
        // console.log('data = ', data);
        var blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        saveAs(blob, config.excelName + new Date().getTime() + ".xlsx");
    });

};

export {
    exportExcel,
}