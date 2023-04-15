

    // ****** Global variable ********
    console.log("Excel Framework version 0.2 loaded")
    var viz, sheet, xmlTABName, filterTABName, _downloadDetail, _progressBARInterval;
    var dynamicURL = document.referrer;

    // ****** Disable Tableau Toolbar On tableau UI ************ 
    // ****** Toolbar Should be disabled on tableau UI on initial load for Main Dashboard      
    let _disableUIToolbar = parent.document.getElementById('toolbar-container')
    _disableUIToolbar.style.display ="none" ;

    //************** Tableau API Export  **************/
    // var uri = 'data:application/vnd.ms-excel;base64,'
    var uri = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,'
        , tmplWorkbookXML = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
            + '<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"></DocumentProperties>'
            + '<Styles><Style ss:ID="colstyle"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Calibri" ss:Bold="1" ss:Size="11"/></Style></Styles>'
            + '{worksheets}</Workbook>'
        , tmplWorksheetXML = '<Worksheet ss:Name="{nameWS}"><Table>{colwidth}{rows}</Table></Worksheet>'
        , tmplCellXML = '<Cell><Data ss:Type="{nameType}">{data}</Data></Cell>'
        , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
        , format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
    var ctx = "";
    var workbookXML = "";
    var worksheetsXML = "";
    var rowsXML = "";
    var dynamicColumnSize = "";

    const getURL = (extensionURL) => {
        let regexMethod = extensionURL.split("/");
        let regexURL = "";
        for (let i = 0; i < regexMethod.length; i++) {
            regexURL += regexMethod[i] + "/";
            if (i === regexMethod.length - 2) {
                regexURL += "MDContainer"
                break;
            }
        }
        return regexURL;
    }

    //*****  URL config : -> Dynamic call window location****
    function _urlConfig() {
        console.log("Download Dialog -> _getURL => (tableauReportUrl) :  " + dynamicURL)
        console.log("Download Dialog -> document.referrer " + document.referrer)
        console.log("Download Dialog -> Windows location path : " + window.location)
        console.log("Download Dialog -> Window location parent : " + window.parent.location)
    }

    //******** Appending dynamic tableau scripts based on environment
    const _injectTableauScripts = () => {
        try {
            let _headerContainer = document.getElementById("headerContainer")
            let _scriptElement = document.createElement('script')
            let injectScript = dynamicURL.split(".net")[0] + ".net/javascripts/api/tableau-2.2.1.min.js"
            _scriptElement.setAttribute('type', 'text/javascript');
            _scriptElement.setAttribute('src', injectScript)
            _headerContainer.insertBefore(_scriptElement, _headerContainer.firstChild);
            console.log("Injecting tableau api call src : " + _headerContainer)
        } catch (error) {
            alert("Dynamic call on tableau Scripts not loaded, please check with tableau administrator")
        }
    }


    // **************** Onload init & Event Listener *************
    const initViz = () => {
        var containerDiv = document.getElementById("vizContainer"),
            url = getURL(dynamicURL);

        options = {
            hideTabs: false,
            hideToolbar: false,
            toolbarPosition: tableau.ToolbarPosition.TOP,
            onFirstInteractive: function () {
                sheet = viz.getWorkbook().getPublishedSheetsInfo();
                document.getElementById("excelBTN").style.visibility = "visible";
                document.getElementById("excelBTN").disabled = false;

                // --> Onclick Listener Function 
                document.getElementById("excelBTN").addEventListener("click", _excelExport);
                document.getElementById("closeWindow").addEventListener("click", closeWindow);

                // Sheet name 
                console.log(viz.getWorkbook().getActiveSheet()._impl.$a.$y)
            }
        };
        if (viz == undefined)
            viz = new tableau.Viz(containerDiv, url, options);
    }



    const donwloadStatus = () => {
        document.getElementById("downloadDetails").value += "Starting download....." + '\r\n';
        document.getElementById("downloadStatus").innerHTML = "Status : Downloading....";
        let counter = 0;
        _progressBARInterval = setInterval(() => {
            counter += 1;
            document.getElementById("_progressBar").value = counter;
            if (counter === 100) {
                document.getElementById("_progressBar").value = 0;
                counter = 0;
            }
        }, 50);
        for (let i = 0; i < 100; i++) {
            document.getElementById("_progressBar").value = i;
        }
    }

    /******************* Export Excel Function ***********/
    const saveExcel = (filename) => {
        // Blob Declaraion method
        ctx = { created: "DOWNLOAD REPORT", worksheets: worksheetsXML };
        workbookXML = format(tmplWorkbookXML, ctx);
        var excel_blob = new Blob([workbookXML], { type: 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        var link = document.createElement("A");
        link.href = window.webkitURL.createObjectURL(excel_blob)
        link.download = filename + ".xls";
        link.target = '_blank';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        setTimeout(() => {
            document.getElementById("downloadDetails").value += "Exporting " + filename + ".xls exported ....." + '\r\n';
            document.getElementById("downloadDetails").value += "Download Completed " + '\r\n';
            document.getElementById("downloadDetails").value += '\r\n' + '\r\n' + "****NOTE****" + "Please wait till downloaded icon is visible at the bottom of browser" + '\r\n';
            document.getElementById("_progressBar").value = 100;
            document.getElementById("downloadStatus").innerHTML = "Status : complete";
            clearInterval(_progressBARInterval)
            document.getElementById("closeWindow").disabled = false;
        }, 2000);
    }

    /******** Load All WorkSheet / Tableau API CALL / Handle Dashboard API *******/
    const exportListener = async () => {
        let valFilter;
        let val, valExport, tempName;
        let _ignoreDashoard = ["Main Dashboard", "MD Container"]


        for (let i = 0; i < sheet.length; i++) {
            if (sheet[i].$0.name != _ignoreDashoard[0] && sheet[i].$0.name != _ignoreDashoard[1]) {

                val = await viz.getWorkbook().activateSheetAsync(sheet[i])

                // ** Data Extraction Multitab / Current release**
                document.getElementById("downloadDetails").value += sheet[i].$0.name.trim() + " : Download in progress ....." + '\r\n';
                await generateExport(val, sheet[i].$0.name.trim())

                const _changeFilterSheet = modifyTabName(val, sheet[i].$0.name.trim())
                for (let a = 0; a < val.getWorksheets().length; a++) {
                    let sheetName = val.getWorksheets()[a].getName();

                    if (sheetName != "User" && sheetName != "Data Refresh" && sheetName != "User2" && sheetName != "Reset" && sheetName != "Row Count" && sheetName != "Rowcount (2)" && sheetName != "User ") {
                        document.getElementById("downloadDetails").value += sheet[i].$0.name.trim() + " : filter's download in progress" + '\r\n';
                        await viz.getWorkbook().getActiveSheet().getWorksheets().get(sheetName).getFiltersAsync().then(function (filters) {
                            appendFilter(filters)
                            document.getElementById("downloadDetails").value += sheet[i].$0.name.trim() + " : Filter's download Completed" + '\r\n';
                            (!_changeFilterSheet) ? filterTABName = sheetName.substring(0, 27) : filterTABName = _changeFilterSheet[a].substring(0, 27);
                        });

                        //  Handle Multiple Sheetname
                        if (rowsXML != '') {
                            ctx = { rows: rowsXML, nameWS: "F - " + filterTABName, colwidth: dynamicColumnSize };
                            worksheetsXML += format(tmplWorksheetXML, ctx);
                            rowsXML = "";
                        }
                    }
                }
            }
        }
        saveExcel(viz.getWorkbook().getActiveSheet()._impl.$a.$y);
    };


    //***** Load data inside worksheet & dashboard / Tableau API call *****
    const generateExport = async (_sheetDetails, sheetName) => {
        let _worksheet;

        // ****** Decalration ******
        const _changeSheetName = modifyTabName(_sheetDetails, sheetName)
        const _CURRENT_REPORT = viz.getWorkbook().getActiveSheet()._impl.$a.$y

        // ****** Option for Certain Condition 
        if (_sheetDetails.getSheetType() === 'dashboard') {
            for (let i = 0; i < _sheetDetails.getWorksheets().length; i++) {
                try {
                    _worksheet = _sheetDetails.getWorksheets()[i].getName();
                    if (_worksheet != "User" && _worksheet != "Data Refresh" && _worksheet != "Rowcount" && _worksheet != "Rowcount (2)") {
                        let _save;
                        options = {
                            maxRows: 150000, // Max rows to return. Use 0 to return all rows
                            ignoreAliases: false,
                            ignoreSelection: true
                        };

                        await _sheetDetails.getWorksheets()[i].getSummaryDataAsync(options).then(dataTable => {
                            _save = dataTable.getData().length >= options.maxRows ? true : false

                            //***** Get if report is proxy server 
                            const _DISABLE_PROXY = () => {
                                const array_condition = new Array()
                                for (let row = 0; row < dataTable.getColumns().length; row++) {
                                    array_condition.push(dataTable.getColumns()[row].getFieldName());
                                }
                                return array_condition.includes('Measure Names') && array_condition.includes('Measure Values')
                                    ? true : false
                            }


                            // Warning Condtion for attributes declared in measures in report
                            if (_save) {
                                const _RecordCount = () => {
                                    let _tmp = new Array()
                                    for (let i = 0; i < dataTable.getData().length; i++) {
                                        for (let a = 0; a < dataTable.getColumns().length; a++) {
                                            let data = dataTable.getData()[i][a]
                                            if (data.value.includes('sql')) {
                                                _tmp.push(data.formattedValue)
                                            }
                                        }
                                    }
                                    const _proxyelement = _tmp.filter((item, index) => _tmp.indexOf(item) === index);
                                    return (_DISABLE_PROXY()) ? (dataTable.getData().length / 3) : dataTable.getData().length;
                                }

                                document.getElementById('alert_box').style.visibility = 'visible'
                                document.getElementById('error_msg').innerHTML += `<p><strong> Warning : </strong> Records exceeded Max Count for  ${sheetName} </p>`;
                                document.getElementById('error_msg').innerHTML += `<p> Record count downloaded : ${_RecordCount()} </p>`;

                                document.getElementById("closebutton").onclick = async () => {
                                    document.getElementById('alert_box').style.visibility = 'hidden'
                                };

                                const _MaxCount = _DISABLE_PROXY() ? _RecordCount() : options.maxRows;
                                document.getElementById("downloadDetails").value += `${sheetName} : Records Exceeded Max Count :` + '\r\n';
                                document.getElementById("downloadDetails").value += `${sheetName} : Records Downloaded :  ${_MaxCount}` + '\r\n';
                                appendDataSummary(dataTable, _DISABLE_PROXY())
                                document.getElementById("downloadDetails").value += `${sheetName} : Download Completed` + '\r\n';
                                (!_changeSheetName) ? xmlTABName = _sheetDetails.getName() : xmlTABName = _changeSheetName[i].substring(0, 31);
                            } else {
                                appendDataSummary(dataTable, _DISABLE_PROXY())
                                document.getElementById("downloadDetails").value += sheetName + " : Download Completed" + '\r\n';
                                (!_changeSheetName) ? xmlTABName = _sheetDetails.getName() : xmlTABName = _changeSheetName[i].substring(0, 31);
                            }
                        });
                    }
                } catch (e) {
                    console.log('error' + error);
                }
                if (rowsXML != '') {
                    ctx = { rows: rowsXML, nameWS: xmlTABName, colwidth: dynamicColumnSize };
                    worksheetsXML += format(tmplWorksheetXML, ctx);
                    rowsXML = "";
                }
            }
        }
    }
     /*
    *   ModifTabName function will dynamically change tab name in excel if it has duplicated name
    *   For e.g if tab name is appended asMTD report and MTD Report in excel [ As certain dashboard has > 1 sheet ]
    *   this function will modify it to [DashboardName] + [ReportName] 
    */
    const modifyTabName = (_sheetDetails, sheetName) => {
        let _noOfSheet = new Array()
        const counter = new Map();

        // insert only sheet if its dashboard
        if (_sheetDetails.getSheetType() === 'dashboard') {
            for (let i = 0; i < _sheetDetails.getWorksheets().length; i++) {
                _noOfSheet.push(sheetName)
            }
        }

        try {
            return result =
                _noOfSheet.map(element => {
                    let tempNo = 1;
                    if (!counter.has(element)) {
                        counter.set(element, _sheetDetails.getWorksheets()[tempNo].getName());
                        return element;
                    }
                    const count = counter.get(element);
                    counter.set(element, _sheetDetails.getWorksheets()[tempNo + 1].getName());
                    return element + "-" + count;
                }
                )
        } catch (error) {
            return false;
        }
    }

    //***** Load Data filter into excel
    const appendFilter = (filters) => {
        let transposeArr = new Array()

        // Data Columns
        rowsXML += '<Row>'
        for (let i = 0; i < filters.length; i++) {
            let colHeader = filters[i].$1
            colHeader = colHeader.includes('SUM') ? colHeader.replace('SUM(', '').replace(')', '') : colHeader
            colHeader = colHeader.includes('AGG') ? colHeader.replace('AGG(', '').replace(')', '') : colHeader
            ctx = { nameType: 'String', data: colHeader };
            rowsXML += format(tmplCellXML, ctx);

            //  Append values in 2D array for filters that calculated : Sum
            if (filters[i].$4 === "SUM") filters[i].$9 = [filters[i].$9];

            for (let a = 0; a < filters[i].$9.length; a++) {
                if (!transposeArr[a]) transposeArr[a] = [''];
                transposeArr[a][i] = filters[i].$9[a].formattedValue;
            }
        }
        rowsXML += '</Row>'

        // // append data to excel
        for (let i = 0; i < transposeArr.length; i++) {
            rowsXML += '<Row>'
            for (let a = 0; a < transposeArr[i].length; a++) {
                let dataValue = transposeArr[i][a];
                if (dataValue === undefined) dataValue = " ";
                ctx = { nameType: 'String', data: dataValue };
                rowsXML += format(tmplCellXML, ctx);
            }
            rowsXML += '</Row>'
        }
    }

    //********* API Append data *******
    //********* Data Optimisation Algorithm  ***************
    const appendDataSummary = (dataTable, _DISABLE_PROXY) => {
        // let _proxy_element = new Array()
        let _setProxyColumn = new Array()
        let _requireOutput = new Array();
        let _multipleProxy = new Array();
        let _counter = 0;
        let rmRow = ''
        let _replaceAttr = ['SUM', 'ARRG', 'AGG', 'ATTR', "AVG"]

        if (_DISABLE_PROXY) {
            // Pattern for sql occurance in loop
            const export_proxy = () => {
                _findProxy = dataTable.getData()[0]
                return (_findProxy[_findProxy.length - 3].value.includes('sql'))
                    ? true : false
            }

            let _setColumnsArr = dataTable.getColumns().map(col => {
                if (col.getFieldName() != 'Measure Names' && col.getFieldName() != 'Measure Values') { return col.getFieldName() }
            }).filter(element => { return element !== undefined })
            let _setColumnsArr_length = _setColumnsArr.length


            for (let i = 0; i < dataTable.getData().length; i++) {
                for (let a = 0; a < dataTable.getColumns().length; a++) {
                    let data = dataTable.getData()[i][a]
                    if (data.value.includes('sql')) {
                        _setColumnsArr_length++
                        _setProxyColumn.push(data.formattedValue)
                    }
                }
            }

            // _sqlProxy is for getting the variable from sql object dynamically 
            const _getSqlProxy = _setProxyColumn.filter((item, index) => _setProxyColumn.indexOf(item) === index);
            const _getColumnsArr = _setColumnsArr.concat(_getSqlProxy)

            rowsXML += '<Row ss:StyleID="colstyle">'
            for (let i = 0; i < _getColumnsArr.length; i++) {
                // Dynamicall append column size with column header
                // dynamicColumnSize += '<Column ss:AutoFitWidth="0" ss:Width="' + _getColumnsArr[i].length * 7.5 + '"/>'
                dynamicColumnSize += '<Column ss:AutoFitWidth="0" ss:Width="' + 150 + '"/>'
                ctx = {
                    nameType: 'String',
                    data: _getColumnsArr[i]
                };
                rowsXML += format(tmplCellXML, ctx);
            }
            rowsXML += '</Row>'


            //************************************* ROW *************************************//
            // Merge row if before column is same = date | reduces xml bytes 
            for (let i = 0; i < dataTable.getData().length; i++) {
                for (let a = 0; a < dataTable.getColumns().length; a++) {

                    let outputdata = (_counter == 0)
                        ? (!dataTable.getData()[i][a].value.includes('sql') && a != (dataTable.getColumns().length - 1))
                            ? _requireOutput.push(dataTable.getData()[i][a].formattedValue)
                            : null : null

                    if (dataTable.getData()[i][a].value.includes('sql')) {
                        if (export_proxy()) _multipleProxy.push(dataTable.getData()[i][a + 2].formattedValue)
                        if (!export_proxy()) _multipleProxy.push(dataTable.getData()[i][a + 1].formattedValue)
                    }
                }

                _counter++
                if ((i + 1) % _getSqlProxy.length == 0) {
                    rowsXML += '<Row>'

                    for (const element of _multipleProxy) {
                        _requireOutput.push(element)
                    }

                    for (const element of _requireOutput) {
                        ctx = { nameType: 'String', data: element };
                        rowsXML += format(tmplCellXML, ctx);
                    }
                    _requireOutput = []
                    _multipleProxy = []
                    _counter = 0
                    rowsXML += '</Row>'

                }

            }
        }


        //***** Report name exist in array loop ******
        //***** Report does not consist of SQLPRoxy array ****
        if (!_DISABLE_PROXY) {
            rowsXML += '<Row ss:StyleID="colstyle">'
            for (let i = 0; i < dataTable.getColumns().length; i++) {
                let tempData = dataTable.getColumns()[i].getFieldName();
                for (let a = 0; a < _replaceAttr.length; a++) {
                    tempData = tempData.includes(_replaceAttr[a]) ? tempData.replace(_replaceAttr[a] + '(', '').replace(')', '') : tempData
                }
                // Dynamicall append column size with column header
                // dynamicColumnSize += '<Column ss:AutoFitWidth="0" ss:Width="' + _getColumnsArr[i].length * 7.5 + '"/>'
                dynamicColumnSize += '<Column ss:AutoFitWidth="0" ss:Width="' + 150+ '"/>'
                ctx = {
                    nameType: 'String',
                    data: tempData
                };
                rowsXML += format(tmplCellXML, ctx);
            }
            rowsXML += '</Row>'


            //  *** Row data ***
            for (let i = 0; i < dataTable.getData().length; i++) {
                rowsXML += '<Row>'
                for (let a = 0; a < dataTable.getColumns().length; a++) {
                    var dataValue = dataTable.getData()[i][a].value;
                    dataValue = dataTable.getData()[i][a].formattedValue
                    if(dataTable.getData()[i][a].formattedValue === 'Null') dataValue = ' ';
                    dataValue = (dataValue.includes("sqlproxy")) ? dataTable.getData()[i][a].formattedValue : dataValue;
                    if (!isNaN(dataValue) && dataValue.includes('.')) dataValue = parseFloat(dataValue).toFixed(4);
                    dataValue = (dataValue.includes("%null%")) ? console.log('null values') : dataValue
                    if (rmRow != '' && dataValue.toLowerCase().includes(rmRow)) continue;
                    ctx = {
                        nameType: 'String',
                        data: dataValue
                    };
                    rowsXML += format(tmplCellXML, ctx);
                }
                rowsXML += '</Row>'
            }
        }
    }


    // ******* Class Run Framework *********
    class _run {
        constructor(exportListener, donwloadStatus) {
            this.exportListener = exportListener;
            this.downloadStaus = donwloadStatus
        }
        _runExport() {
            this.exportListener
            this.donwloadStatus
            if (Object.keys(ctx).length === 0) {
                ctx = { created: "DOWNLOAD REPORT", worksheets: worksheetsXML };
                workbookXML = format(tmplWorkbookXML, ctx);
                tmplCellXML = '<Cell><Data ss:Type="{nameType}">{data}</Data></Cell>'
            }
        }
        _clearExcel() {
            ctx = {};
            workbookXML = "";
            worksheetsXML = "";
            tmplCellXML = ""
        }
        _addTableauScript() {
            _injectTableauScripts()
        }
    }

    //********** Actual Trigger ***********
    const run = new _run() // => Run Program
    run._addTableauScript()
    _urlConfig(); // => Check URL for dyanamic call

    // Declaring a new class to handle run and stop export
    const _excelExport = () => {
        let container = document.getElementById("popUpWindow").style.visibility = "visible";
        if (document.getElementById("vizContainer_disabled") != null) {
            var dynamicURL = document.getElementById("vizContainer_disabled").style.visibility = "visible";
        }
        document.getElementById("closeWindow").disabled = true;
        document.getElementById("downloadDetails").value += "Initialising download....." + '\r\n';
        document.getElementById("downloadDetails").value += "Please wait....." + '\r\n';
        run._runExport(exportListener(), donwloadStatus())
    }

    const closeWindow = () => {
        document.getElementById("downloadDetails").value = "";
        document.getElementById("popUpWindow").style.visibility = "hidden";
        document.getElementById("vizContainer_disabled").style.visibility = "hidden";
        run._clearExcel()
    }
