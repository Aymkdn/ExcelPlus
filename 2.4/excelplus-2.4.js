/*!
 * ExcelPlus JavaScript Library v2.3
 *
 * Copyright 2016, aymeric@kodono.info
 * Licensed under GPL Version 3 licenses.
 *
 * http://aymkdn.github.io/ExcelPlus/
 */
// event handler polyfill
function addEventListener(el, eventName, handler) {
  if (el.addEventListener) {
    el.addEventListener(eventName, handler);
  } else if (el.attachEvent) {
    el.attachEvent('on' + eventName, function(){
      handler.call(el);
    });
  }
}

// Global variables
var __ExcelPlus_flashPath="";

/**
  * @name ExcelPlus()
  * @constructor
  * @description This is the constructor
  *
  * @param {Object} params
  *   @param {Boolean} [params.showErrors=true] The errors will be shown and will throw an error
  *
  */
function ExcelPlus(params) {
  this.aLetters = [];
  this.aLetters["A"]=1, this.aLetters["B"]=2, this.aLetters["C"]=3, this.aLetters["D"]=4, this.aLetters["E"]=5, this.aLetters["F"]=6, this.aLetters["G"]=7, this.aLetters["H"]=8, this.aLetters["I"]=9, this.aLetters["J"]=10, this.aLetters["K"]=11, this.aLetters["L"]=12, this.aLetters["M"]=13, this.aLetters["N"]=14, this.aLetters["O"]=15, this.aLetters["P"]=16, this.aLetters["Q"]=17;
  this.aLetters["R"]=18, this.aLetters["S"]=19, this.aLetters["T"]=20, this.aLetters["U"]=21, this.aLetters["V"]=22, this.aLetters["W"]=23, this.aLetters["X"]=24, this.aLetters["Y"]=25, this.aLetters["Z"]=26;
  
  this.oSheet = null; // the current sheet that we are working on
  this.selectedSheet = ""; // name of the selected sheet
  this.nbSheets = 0; // number of sheets
  this.oFile = null; // the File object
  this.nbRows = 0, this.nbColumns = 0; // the number of rows and columns used in the opened file
  this.filename = ""; // name of the file we read
  this.usedRange = ""; // e.g. A1-C250
  this.flashUsed = false;
  
  params = params || {};
  this.error = "";

  // to show the errors if any, otherwise you can handle the errors using YourObject.errors
  this.showErrors = (params.showErrors === undefined ? true : params.showErrors);
}

ExcelPlus.prototype = {
  /**
    @name ExcelPlus().reset
    @function
    @description This function permits to reset the current ExcelPlus object
  */
  reset:function() {
    this.oFile=null;
    this.nbSheets=0;
    this.nbRows=0;
    this.nbColumns=0;
    this.usedRange="";
    this.oSheet=null;
    this.selectedSheet="";
    
    return this;
  },
  /**
   * @name ExcelPlus().openLocal
   * @function
   * @description This function permits to read a local file (on the user's computer)
   * @description To do so, we'll use a input[file] -- integration of the code from http://aymkdn.github.io/FileToDataURI/
   * @description The HTML code for the input[file] will be &lt;object id="file-object">&lt;/object>
   *
   * @param {Object} params
   *   @param {String} [params.idButton="file-object"] This is the ID of the &lt;object> tag in your code
   *   @param {String} [params.flashPath="./"] You must provide the absolute or relative path to the Flash files (swfobject.js, FileToDataURI.swf and expressInstall.swf) 
   *   @param {String} [params.labelButton="Load a file"] This is the label for the button
   * @param {Function} params.onReady=function(){} The function that will be called when the file has been loaded
   */
  openLocal:function(params, onReady) {
    if (typeof params === "function") {
      onReady=params;
      params={};
    }
    onReady = onReady || function() {};
    
    params.idButton = params.idButton || "file-object";
    if (document.getElementById(params.idButton) === null) {
      this.error = "[openLocal] You must have a <object> tag in your code.";
      this._showErrors();
      return this;
    }
    
    params.labelButton = params.labelButton || "Load a file";

    __ExcelPlus_flashPath = params.flashPath || "";   // related to Flash features
  
    var _this=this;
    // if FileReader is available
    if (typeof FileReader === "function") {
      // So we use a regular input file
      document.getElementById(params.idButton).outerHTML = '<input id="'+params.idButton+'" type="file" value="'+params.labelButton+'" />';
      // add an event when loading the file
      addEventListener(document.getElementById(params.idButton), "change", function(e) {
        var files = e.target.files,file;
        if (!files || files.length == 0) return;
        file = files[0];
        var fileReader = new FileReader();
        fileReader.onload = function (e) {
          // ATTENTION: to have the same result than the Flash object we need to split
          // our result to keep only the Base64 part
          _this.filename = file.name;
          // then call 'xlsx' to read the file
          _this.oFile = XLSX.read(e.target.result, {type: 'binary', cellDates:true, cellStyles:true});
          _this.nbSheets = _this.oFile.SheetNames.length;
          if (_this.nbSheets == 1) {
            _this.selectSheet(1)
          }
          onReady.call(_this, true);
        };
        fileReader.readAsArrayBuffer(file);
      });
    } else {
      _this.flashUsed=true;
      Flash.getButtonLabel = function() { return params.labelButton };
      // we call Flash to be able to read the file
      Flash.createButton(params.idButton, function(base64) {
        // then call our function
        _this.filename = "Unknown Name";
        _this.oFile = XLSX.read(base64, {type: 'base64', cellDates:true, cellStyles:true});
        _this.nbSheets = _this.oFile.SheetNames.length;
        if (_this.nbSheets == 1) {
          _this.selectSheet(1);
        }
        onReady.call(_this, true);
      })
    }
    
    return _this;
  },
  /**
   * @name ExcelPlus().openRemote
   * @function
   * @description This function read a remote Excel file
   *
   * @param {String} url The url of the Excel file (make sure it follows the cross domain security)
   * @param {Function} onReady=function(passed){} This function is called when the Excel file has been read, and 'passed' will return true (success) or false (failed)
   */
  openRemote:function(url, onReady) {
    var _this=this;
    // get the remote file binary content
    var xhr = new XMLHttpRequest();
    var convertResponseBodyToText = function(e) { return e };
    xhr.open("GET", url, true);
    if (xhr.overrideMimeType) xhr.overrideMimeType("text/plain; charset=x-user-defined")
    else {
      // for old browsers like IE8
      xhr.setRequestHeader("Accept-Charset", "x-user-defined");
      // see http://stackoverflow.com/questions/1919972/how-do-i-access-xhr-responsebody-for-binary-data-from-javascript-in-ie
      var VB_Fix_IE = '<script language="VBScript">'+"\r\n"
                    + "Function IEBinaryToArray_ByteStr(Binary)\r\n"
                    + "  IEBinaryToArray_ByteStr = CStr(Binary)\r\n"
                    + "End Function\r\n"
                    + "Function IEBinaryToArray_ByteStr_Last(Binary)\r\n"
                    + "  Dim lastIndex\r\n"
                    + "  lastIndex = LenB(Binary)\r\n"
                    + "  if lastIndex mod 2 Then\r\n"
                    + "    IEBinaryToArray_ByteStr_Last = Chr( AscB( MidB( Binary, lastIndex, 1 ) ) )\r\n"
                    + "  Else\r\n"
                    + "    IEBinaryToArray_ByteStr_Last = "+'""'+"\r\n"
                    + "  End If\r\n"
                    + "End Function\r\n"
                    + "\<\/script>\r\n";
      document.write(VB_Fix_IE);
      
      convertResponseBodyToText = function(binary) {
        var byteMapping = {};
        for ( var i = 0; i < 256; i++ ) {
          for ( var j = 0; j < 256; j++ ) {
            byteMapping[ String.fromCharCode( i + j * 256 ) ] = String.fromCharCode(i) + String.fromCharCode(j);
          }
        }
        var rawBytes = IEBinaryToArray_ByteStr(binary);
        var lastChr = IEBinaryToArray_ByteStr_Last(binary);
        return rawBytes.replace(/[\s\S]/g, function( match ) { return byteMapping[match]; }) + lastChr;
      };
    }
    xhr.onreadystatechange = function(event) {
      if (xhr.readyState == 4) {
        if(xhr.status == 200) {
          _this.filename = url.split("/").slice(-1);
          var data = (xhr.overrideMimeType ? xhr.responseText : convertResponseBodyToText(xhr.responseBody));
          _this.oFile = XLSX.read(data, {type: 'binary', cellDates:true, cellStyles:true});
          _this.nbSheets = _this.oFile.SheetNames.length;
          if (_this.nbSheets == 1) {
            _this.selectSheet(1);
          }
          onReady.call(_this, true);
        } else {
          _this.error = "[openRemote] Cannot find the remote file";
          _this._showErrors();
          onReady.call(_this, false)
        }
      }
    };
    xhr.send(null);
    
    return this;
  },
  /**
   * @name ExcelPlus().getSheetNames
   * @function
   * @description This function returns an array of the sheet names (the key is the sheet #)
   *
   * @return {Array} The different sheets related to the Excel file
   */
  getSheetNames:function() {
    if (this.oFile == null) {
      this.error = "[getSheetNames] No Excel file opened.";
      this._showErrors();
      return this;
    }
    
    return this.oFile.SheetNames;
  },
  /**
   * @name ExcelPlus().selectSheet
   * @function
   * @description This function selects the specific sheet
   *
   * @param {String|Number} sheet If it's a number (starting from 1) that reprents the sheet position in the Excel file, otherwise it must be the name of the sheet (case sensitive)
   */
  selectSheet:function(sheet) {
    if (this.oFile == null) {
      this.error = "[selectSheet] No Excel file opened.";
      this._showErrors();
      return this;
    }
    
    // if we have a number, then we search the related sheetname
    if (!isNaN(sheet)) {
      // find the sheet name based on the number
      var arr = this.getSheetNames();
      sheet-=1;
      var len = arr.length;
      if (sheet >= len) {
        this.error = "[selectSheet] Impossible to select sheet #"+sheet+" because only "+len+" sheet"+(len>1?"s":"")+" found in the file.";
        this._showErrors();
        return this;
      }
      
      for (var i=0; i < len; i++) {
        if (sheet == i) {
          sheet = arr[i];
          break;
        }
      }
    }
    if (!isNaN(sheet)) {
      this.error = "[selectSheet] Impossible to select sheet #"+sheet+" because I cannot find it.";
      this._showErrors();
      return this;
    }

    // make sure the provided name exists
    if (this.oFile.SheetNames.indexOf(sheet) === -1) {
      this.error = "[selectSheet] Impossible to select sheet "+sheet+" because there is no sheet with that name.";
      this._showErrors();
      return this;
    }
    
    this.oSheet      = this.oFile.Sheets[sheet];
    this.selectedSheet = sheet;
    if (this.oSheet["!ref"]) {
      this.usedRange = this.oSheet["!ref"]; // e.g. A1-C250
      var lastCell   = this._getCellCoord(this.usedRange.split(":")[1]);
      this.nbRows    = lastCell.row * 1;
      this.nbColumns = lastCell.column * 1;
      // if flash is used and the A1 is null, then do nbColumns-1
      //if (this.flashUsed && this.oSheet["A1"] == null) this.nbColumns--;
    }
    
    return this;
  },
  /**
   * @name ExcelPlus().readAll
   * @function
   * @description This function will read the all of the used range (we assume that the content starts at A1)
   *
   * @param {Object} options
   *   @param {Boolean} [options.parseDate=false] If a Date is found then it will try to parse the date to return a JS Date object
   *   @param {Boolean} [options.properties=false] When it's true it will return an object with all properties has defined by sheetjs/XLSX-JS instead of return the cell value
   * @return {Array} it returns a 2D array; the first one is the rows, and the second one is the columns
   */
  readAll:function(options) {
    if (this._checkSheet()==false) {
      this.error = "[readAll] You have to select a sheet before using `readAll`";
      this._showErrors();
      return this;
    }
       
    var table = [],line;
    for (var row=1; row <= this.nbRows; row++) {
      line = [];
      for (var col=1; col <= this.nbColumns; col++) {
        line.push(this.read(this._getColumnName(col) + "" + row, options));
      }
      table.push(line);
    }
    return table;
  },
  /**
   * @name ExcelPlus().read
   * @function
   * @description This function permits to read a cell or a range
   *
   * @param {String} range can be a range (e.g. "A1:D1"), or a single cell (e.g. "A1")
   * @param {Object} options
   *   @param {Boolean} [options.parseDate=false] If a Date is found then it will try to parse the date to return a JS Date object
   *   @param {Boolean} [options.properties=false] When it's true it will return an object with all properties has defined by sheetjs/XLSX-JS instead of return the cell value
   * @return {Date|String|Array|Object} if it's a range then it returns a 2D array (the first one is the rows, and the second one is the columns), if it's a cell then it returns the value of the cell, or the object if optoins.properties is true
   */
  read:function(range, options) {
    if (this._checkSheet()==false) return this
    options = options || {};
    options.parseDate = (options.parseDate === true ? true : false);
    options.properties = (options.properties === true ? true : false);
    var val=null;
    
    if (this._isRange(range)) {
      // if we use Flash, then the first column is probably NULL
      range = XLSX.utils.decode_range(range);
      var table=[], line=[], R, C;
      for (R = range.s.r; R <= range.e.r; ++R) {
        line = [];
        for (C = range.s.c; C <= range.e.c; ++C) {
          line.push(this.read(XLSX.utils.encode_cell({c:C, r:R}), options));
        }
        table.push(line);
      }
      return table;
    }
    else if (this._isCell(range)) {
      val = this.oSheet[range];
      if (val) {
        // sheetjs doesn't provide the correct .t="d" when it's date
        // so I've changed it, and now it includes .fmt
        // In my tests, .fmt is like "[$-409]d\-mmm\-yy;@"
        if (val.fmt && val.fmt.charAt(0) === "[") {
          val.t = "d"; 
        }
        
        if (options.properties) return val;
        
        // if it's a Date val.t == "d"
        // then we parse it
        if (val.t === "d" && options.parseDate) {
          val = XLSX.SSF.parse_date_code(val.v);
          val = new Date(val.y,val.m-1,val.d,val.H,val.M,val.S);
        }
        else {
          val = val.w || val.v;
        }
      } else val=null;

      return val;
    }
    
    return null;
  },
  /**
   * @name ExcelPlus().createFile
   * @function
   * @description Permits to create an empty Excel file
   *
   * @param {String|Array} [sheetnames="Sheet1"] You can provide the names of the sheets in an array
   * @example
   *    var myExcel = new ExcelPlus();
   *    myExcel.createFile(["Sheet 1", "Sheet 2"]);
   */
  createFile:function(sheetnames) {
    sheetnames = sheetnames || ["Sheet1"];
    if (typeof sheetnames === "string") sheetnames=[sheetnames];
    this.oFile = {
      "SheetNames": [],
      "Sheets": {}
    };
    for (var s=0; s<sheetnames.length; s++) {
      if (this.createSheet(sheetnames[s]) === false) return this
    }
    
    return this;
  },
  /**
   * @name ExcelPlus().write
   * @function
   * @description Write values into the workbook
   *
   * @param {Object} params
   *   @param {String|Boolean|Number|Date|Array} params.content If it's an array then we ignore the "cell" param (e.g. [ [ "Content Cell A1", "Content Cell B1" ] [ "Content Cell A2", "Content Cell B2" ] ], otherwise the "cell" needs to be provided
   *   @param {String} [params.sheet] You can give the name of the sheet where to write, or the current selected sheet will be use, or the first sheet will be used
   *   @param {String} [params.cell] You can specify the cell reference (e.g. "A1")
   *
   */
  write:function(options) {
    options.sheet = options.sheet || this.selectedSheet || this.oFile.SheetNames[0];
    // select the sheet requested here
    this.selectSheet(options.sheet);
    
    if (!options.content) {
      this.error = "[write] Please provide the 'content'";
      this._showErrors();
      return this;
    }
    // check if "content" is an array, then ignore "cell"
    if (Array.isArray(options.content)) options.cell=null;
    else {
      if (!options.cell) {
        this.error = "[write] Please provide the 'cell' reference";
        this._showErrors();
        return this;
      }
    }
    
    if (!options.cell) {
      var data=options.content, R, C, cell_ref, cell;
      for (R = 0; R != data.length; ++R) {
        for (C = 0; C != data[R].length; ++C) {
          cell = this._createCell(data[R][C]);
          if (cell.v == null) continue;
          cell_ref = XLSX.utils.encode_cell({c:C,r:R});
          this.oFile.Sheets[this.selectedSheet][cell_ref] = cell;
          this._updateUsedRange(cell_ref);
        }
      }
    } else {
      this.writeCell(options.cell, options.content);
    }
    
    this.oSheet = this.oFile.Sheets[this.selectedSheet];
    
    return this;
  },
  /**
   * @name ExcelPlus().writeCell
   * @function
   * @description Write a value into a cell for the current selected sheet
   *
   * @param {String} cell The cell reference (e.g. "A1")
   * @param {String|Number|Date|Boolean} content This is the value for the cell (could be `null`)
   */
  writeCell:function(cell, content) {
    if (!this.selectedSheet) {
      if (this.oFile.SheetNames.length>0) {
        this.error = "[writeCell] No sheet selected";
        this._showErrors();
        return this;
      } else {
        this.error = "[writeCell] You have to create a sheet first";
        this._showErrors();
        return this;
      }
    }
    
    if (typeof content === "undefined") {
      this.error = "[writeCell] You have to provide the content";
      this._showErrors();
      return this;
    }
    if (!this._isCell(cell)) {
      this.error = "[writeCell] You have to provide a cell reference, e.g. 'A1'";
      this._showErrors();
      return this;
    }
    
    // try to select the sheet to make sure there is no error
    if (content == null) delete this.oFile.Sheets[this.selectedSheet][cell];
    else this.oFile.Sheets[this.selectedSheet][cell] = this._createCell(content);
    this._updateUsedRange(cell);
    // update the info about the selected sheet
    this.selectSheet(this.selectedSheet);
    
    return this;
  },
  /**
   * @name ExcelPlus().writeRow
   * @function
   * @description Write a full row for the current selected sheet
   *
   * @param {String} row The row reference (e.g. "1" for the first row in the sheet)
   * @param {Array} content This is the values for each cell of the row (could be `null`)
   * @example
   *   ep.writeRow(1, ["City", "Postal Code", "Country"]);
   *   ep.writeRow(2, ["Montpellier", "34000", "France"]);
   */
  writeRow:function(row, content) {
    if (!this.selectedSheet) {
      if (this.oFile.SheetNames.length>0) {
        this.error = "[writeRow] No sheet selected";
        this._showErrors();
        return this;
      } else {
        this.error = "[writeRow] You have to create a sheet first";
        this._showErrors();
        return this;
      }
    }
    
    if (typeof content === "undefined") {
      this.error = "[writeRow] You have to provide the content";
      this._showErrors();
      return this;
    }
    if (isNaN(row)) {
      this.error = "[writeRow] You have to provide a row reference, e.g. '1'";
      this._showErrors();
      return this;
    }
    if (Object.prototype.toString.call(content) !== '[object Array]') {
      this.error = "[writeRow] You have to provide an array for the content";
      this._showErrors();
      return this; 
    }
    for (var i=0; i<content.length; i++) {
      this.writeCell(this._getColumnName(i+1)+""+row, content[i]);
    }
        
    return this;
  },
  /**
   * @name ExcelPlus().writeNextRow
   * @function
   * @description Write on the next available row for the current selected sheet
   *
   * @param {Array} content This is the values for each cell of the row (could be `null`)
   * @example
   *   ep.writeRow(["Paris 1", "75001", "France"]);
   */
  writeNextRow:function(content) {
    if (Object.prototype.toString.call(content) !== '[object Array]') {
      this.error = "[writeRow] You have to provide an array for the content";
      this._showErrors();
      return this; 
    }
    this.writeRow(this.nbRows+1, content);

    return this;
  },
  /**
   * @name ExcelPlus().deleteSheet
   * @function
   * @description This function permits to delete a sheet
   *
   * @param {String} sheet The name of the sheet to delete (by default the selected one)
   */
  deleteSheet:function(sheet) {
    sheet = sheet || this.selectedSheet;
    if (sheet) {
      delete this.oFile.Sheets[sheet];
      if (sheet == this.selectedSheet) this.selectedSheet="";
      var idx = this.oFile.SheetNames.indexOf(sheet);
      if (idx > -1) this.oFile.SheetNames.splice(idx,1)
    }
    
    return this
  },
  /**
   * @name ExcelPlus().createSheet
   * @function
   * @description This function permits to create a new empty sheet
   *
   * @param {String} sheet The name of the sheet to create
   * @param {Boolean} [select=true] If you want to automatically select this sheet
   */
  createSheet:function(sheet, select) {
    select = (select === undefined ? true : select);
    
    if (!sheet) {
      this.error = "[createSheet] You have to provide a name for the sheet";
      this._showErrors();
      return this;
    }
    if (this.oFile.SheetNames.indexOf(sheet) !== -1) {
      this.error = "[createSheet] There is already a sheet with the name '"+sheet+"'";
      this._showErrors();
      return this;   
    }
    
    this.oFile.SheetNames.push(sheet);
    this.oFile.Sheets[sheet] = { "!ref":"A1:A1" }
    
    if (select) this.selectSheet(sheet)
    
    return this;
  },
  /**
   * @name ExcelPlus().saveAs
   * @function
   * @description This function permits de save the current book to a real Excel file (only for Modern Browsers and IE10+)
   *
   * @param {String} filename The name of the file, .e.g "book1.xlsx"
   */
  saveAs:function(filename) {
    filename = filename || "book1.xlsx";
    var ext = filename.split(".").pop();
    // check the type/extension
    if (['xlsx', 'xlsm', 'xlsb'].indexOf(ext) === -1) {
      ext = "xlsx";
      filename += "."+ext;
    }
    
    if (typeof Uint8Array === "undefined") {
      this.error = "[saveAs] Sorry but this function is only supported by modern browsers";
      this._showErrors();
      return this;
    }
    
    var wbout = XLSX.write(this.oFile, {bookType:ext, bookSST:true, type: 'binary'});
    var s2ab = function(s) {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    }
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), filename);
    
    return this;
  },
  /**
   * @name ExcelPlus().saveTo
   * @function
   * @description This function permits to return the Excel formatted content that we can then send back to the server
   * @description It actually uses the `XLSX.write()` function
   *
   * @param {String} [ext="xlsx"] It's the extension type ('xlsx', 'xlsm' or 'xlsb')
   * @param {String} [type="base64"] It can be "base64" or "binary"
   * @return {String|Binary} The encoded content of the Excel file
   */
  saveTo:function(ext, type) {
    return XLSX.write(this.oFile, {bookType:ext||"xlsx", bookSST:true, type: type||"base64"});
  },
  /**
   * @ignore
   * This function takes a value and converts it to a `cell` format defined by SheetJS
   *
   * @param {String|Boolean|Date|Number} value The value for a cell
   * @return {Object} It will return a cell object with 't' for the format, 'v' for the formatted value, etc (see https://github.com/SheetJS/js-xlsx#cell-object)
   */
  _createCell:function(value) {
    var cell = {"v":value, "t":this._getValueType(value)};
    if (value instanceof Date) {
      cell.z = XLSX.SSF._table[14];
      cell.v = this._dateToNum(value);
    }

    return cell;
  },
  /**
   * @ignore
   * This function permits to update the used range for a sheet
   * It's based on the SheetJS example code
   *
   * @param {Object|String} coord It's the coordonate of a cell e.g. {column:1, row:1} or "A1"
   */
  _updateUsedRange:function(coord) {
    var range = XLSX.utils.decode_range(this.oSheet["!ref"]);
    if (this._isCell(coord)) coord=this._getCellCoord(coord);
    coord.row-=1;
    coord.column-=1;
    var anyChange = false
    if(range.s.r > coord.row) { anyChange=true; range.s.r = coord.row; }
    if(range.s.c > coord.column) { anyChange=true; range.s.c = coord.column; }
    if(range.e.r < coord.row) { anyChange=true; range.e.r = coord.row; }
    if(range.e.c < coord.column) { anyChange=true; range.e.c = coord.column; }
    
    if (anyChange) {
      range = XLSX.utils.encode_range(range);
      this.oFile.Sheets[this.selectedSheet]["!ref"] = range;
      this.usedRange = range;
      var lastCell   = this._getCellCoord(this.usedRange.split(":")[1]);
      this.nbRows    = lastCell.row * 1;
      this.nbColumns = lastCell.column * 1;
    }
    
    return this;
  },
  /**
   * @ignore
   * This function permits to return the type of a value
   * It's based on SheetJS code example
   *
   * @param {String|Number|Date|Boolean} value The value of a cell
   * @return {String} 'n' for number and date, 'b' for boolean, and 's' for string 
   */
  _getValueType:function(value) {
    if(typeof value === 'number') return 'n';
    if(typeof value === 'boolean') return 'b';
    if(value instanceof Date) return 'n';
    return 's';
  },
  /**
   * @ignore
   * This function permits to change a Date to a number that is readable by Excel
   * It's based on SheetJS code example
   *
   * @param {Date} v It's the Date you want to convert to
   * @param {Boolean} [date1904=false] For Macintosh the calendar starts in 1904 instead of 1900, so if you build something for Mac this param should be TRUE
   * @return {Number} A number version of the date passed
   */
  _dateToNum:function(v, date1904) {
    if(date1904) v+=1462;
    var epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
  }, 
  // @param name is the column name (one or more letters)
  // return the number related to a column name
  _getColumn:function(name) {
    var ret = 0;
    if (name.length == 1) return this.aLetters[name];
    if (name.length > 1) {
      var letters=name.split("");
      var len = letters.length;
      for (var i=0; i < len; i++) {
        if (i+1<len) ret += this.aLetters[letters[i]]*26;
        else ret += this.aLetters[letters[i]];
      }
      return ret;
    } else return 0;
  },
  /**
   * @ignore
   * This function returns the column letter based on a number (e.g. 1 is A)
   *
   * @param {Number} index Index of the column
   * @return [String} the column letter
   */
  _getColumnName:function(index) {
    if (index <= 26) {
      var i = 1;
      for (var letter in this.aLetters) {
        if (letter.length==1) { // to avoid some issues with PrototypeJS
          if (i==index) return letter;
          i++;
        }
      }
    } else {
      var incr = 0;
      do {
        index -= 26;
        incr++;
      } while (index > 26);
      // find the first letter
      var ret = "";
      var i = 1;
      for (var letter in this.aLetters) {
        if (letter.length==1) { // to avoid some issues with PrototypeJS
          if (i==incr) { ret = letter; break; }
          i++;
        }
      }
      // find the second letter
      i = 1;
      for (var letter in this.aLetters) {
        if (letter.length==1) { // to avoid some issues with PrototypeJS
          if (i==index) return ret + "" + letter;
          i++;
        }
      }
    }
  },
  /**
   * @ignore
   * This function transforms a cell reference into its coordinates
   *
   * @param {String} cell The cell reference (e.g. "A1")
   * @return {Object} {column: Y, row: X} (starting from 1 from both column and row)
   */
  _getCellCoord:function(cell) {
    var arr = cell.match(/([A-Z]+)([0-9]+)/);
    if (arr == null || arr.length != 3) {
      this.error = "[_getCellCoord] The format of the cell is wrong.";
      this._showErrors();
      return this;
    }
    return {"column": this._getColumn(arr[1]), "row": arr[2]*1};
  },
  /**
   * @ignore
   * see the index color at http://www.mvps.org/dmcritchie/excel/colors.htm
   */
  _getColorIndex:function(color) {
    if (!isNaN(color)) return color;
    switch (color) {
      case "black":     return 1;
      case "white":     return 2;
      case "red":       return 3;
      case "green":     return 4;
      case "darkblue":  return 5;
      case "blue":      return 23;
      case "yellow":    return 6;
      case "magenta":   return 7;
      case "cyan":      return 8;
      case "purple":    return 13;
      case "lightgray": return 15;
      case "gray":      return 16;
      case "orange":    return 45;
      default: this.error = "[_getColorIndex] '"+color+"' is an unknown color. Known colors are: black, white, red, green, blue, yellow, magenta and cyan";
               this._showErrors();
               return this;
    }
  },
  /**
   * @ignore
   * This function check if it's a range  (e.g. "A1:D3")
   * 
   * @param {String} range
   * @eturn {Boolean} TRUE if it's a range
   */
  _isRange:function(range) { return (typeof range === "string" && range.indexOf(":") != -1 && range.match(/[1-9]/) != null); },
  /**
   * @ignore
   * This function check if it's a cell (e.g. "A1")
   * 
   * @param {String} range
   * @eturn {Boolean} TRUE if it's a cell
   */
  _isCell:function(range) { return (typeof range === "string" && range.indexOf(":") == -1 && range.match(/[1-9]/) != null); },
  /**
   * @ignore
   * This function check if it's a column (e.g. "AA")
   * 
   * @param {String} range
   * @eturn {Boolean} TRUE if it's a column
   */
  _isColumn:function(range) { return (typeof range === "string" && range.match(/[1-9]/) == null); },
  /**
   * @ignore
   * This function checks if a sheet is currently selected
   *
   * @return {Boolean} TRUE if it's the case
   */
  _checkSheet:function() {
    if (this.oSheet==null) {
      this.error = "[_checkSheet] Please select a sheet first";
      this._showErrors();
      return this;
    }
  },
  /**
   * @ignore
   * This function is used to display the errors
   */
  _showErrors:function() {
    if (this.showErrors) {
      if (this.error !== "") {
        throw "[ExcelPlus] "+this.error
      }
      else console.log("[ExcelPlus] No Error");
    }
  },
  /* we use the CSS codes to change the style of the cells (http://www.mvps.org/dmcritchie/excel/font.htm)
  // @param range is a range, a cell or a column (example: "A1", "A1:A3", "A", "A:B")
  // @param hashstyle is the CSS style in the hash format (example: {"width":30, "text-align":"right"})
  setStyle:function(range, hashstyle) {},
  // create a grid in the range
  grid:function(range, color) {},
  // merge a range of cells
  merge:function(range) {},
  // return the first cell after the used rows
  getNewLine:function() {},*/
}

// Create the upload button then read the file using either FileAPI or Flash (see http://aymkdn.github.com/FileToDataURI/)
var swfobject;
var __Flash_getFileData_callback;
var Flash = {
  callback:function() {},
  /**
    * This function permits to create a button to get the file
    *
    * @param {String} id It's the ID of the HTML object
    * @param {Function} callback The function that will be called when the file is read with "base64" version of the file
    */
  createButton:function(id, callback) {
    if (typeof callback==="function") this.callback=callback;
    __Flash_getFileData_callback=this.callback;
    
    // we call swfobject.js
    if (swfobject == undefined) {
      var fileref=document.createElement('script');
      fileref.setAttribute("type","text/javascript")
      fileref.setAttribute("src", __ExcelPlus_flashPath+"/swfobject.js");
      document.getElementsByTagName("head")[0].appendChild(fileref);
      // wait for the file to be loaded
      return this._waitForObjectToBeLoaded(id,callback);
    } else {
      // find the widht based on the button label
      var label = Flash.getButtonLabel();
      var wdth = 8 * label.length; // 8px by letter
      swfobject.embedSWF(__ExcelPlus_flashPath+"/FileToDataURI.swf", id, wdth+"px" /* width of the Flash zone */, "22px" /* height of the Flash zone */, "10", __ExcelPlus_flashPath+"/expressInstall.swf", {}, {}, {});
    }
  },
  /**
    * This function is called when we have read the file
    *
    * @param {String} name It's the filename
    * @param {String} base64 It's the base64 version of the file
    */
  getFileData: function(base64) {
    __Flash_getFileData_callback(base64);
  },
  /** 
    * This function is used to wait for the swfobject to be loaded by the old browser
    * The parameters are the same as createButton
    * This is for internal use only
    */
  _waitForObjectToBeLoaded:function(id,callback) {
    if (swfobject == undefined) setTimeout(function() { Flash._waitForObjectToBeLoaded(id,callback) }, 50)
    else this.createButton(id,callback)
  }
};


/*! FileSaver.js
 *  A saveAs() FileSaver implementation.
 *  2014-01-24
 *
 *  By Eli Grey, http://eligrey.com
 *  License: X11/MIT
 *    See LICENSE.md
 */
/*! @source http://purl.eligrey.com/github/FileSaver.js/blob/master/FileSaver.js */
;var saveAs=saveAs||(typeof navigator!=="undefined"&&navigator.msSaveOrOpenBlob&&navigator.msSaveOrOpenBlob.bind(navigator))||(function(h){if(typeof navigator!=="undefined"&&/MSIE [1-9]\./.test(navigator.userAgent)){return}var r=h.document,l=function(){return h.URL||h.webkitURL||h},e=h.URL||h.webkitURL||h,n=r.createElementNS("http://www.w3.org/1999/xhtml","a"),g=!h.externalHost&&"download" in n,j=function(t){var s=r.createEvent("MouseEvents");s.initMouseEvent("click",true,false,h,0,0,0,0,0,false,false,false,false,0,null);t.dispatchEvent(s)},o=h.webkitRequestFileSystem,p=h.requestFileSystem||o||h.mozRequestFileSystem,m=function(s){(h.setImmediate||h.setTimeout)(function(){throw s},0)},c="application/octet-stream",k=0,b=[],i=function(){var t=b.length;while(t--){var s=b[t];if(typeof s==="string"){e.revokeObjectURL(s)}else{s.remove()}}b.length=0},q=function(t,s,w){s=[].concat(s);var v=s.length;while(v--){var x=t["on"+s[v]];if(typeof x==="function"){try{x.call(t,w||t)}catch(u){m(u)}}}},f=function(t,v){var w=this,C=t.type,F=false,y,x,s=function(){var G=l().createObjectURL(t);b.push(G);return G},B=function(){q(w,"writestart progress write writeend".split(" "))},E=function(){if(F||!y){y=s(t)}if(x){x.location.href=y}else{if(navigator.userAgent.match(/7\.[\d\s\.]+Safari/)&&typeof window.FileReader!=="undefined"&&t.size<=1024*1024*150){var G=new window.FileReader();G.readAsDataURL(t);G.onloadend=function(){var H=r.createElement("iframe");H.src=G.result;H.style.display="none";r.body.appendChild(H);B();return};w.readyState=w.DONE;w.savedAs=w.SAVEDASUNKNOWN;return}else{window.open(y,"_blank");w.readyState=w.DONE;w.savedAs=w.SAVEDASBLOB;B();return}}},A=function(G){return function(){if(w.readyState!==w.DONE){return G.apply(this,arguments)}}},z={create:true,exclusive:false},D;w.readyState=w.INIT;if(!v){v="download"}if(g){y=s(t);r=h.document;n=r.createElementNS("http://www.w3.org/1999/xhtml","a");n.href=y;n.download=v;var u=r.createEvent("MouseEvents");u.initMouseEvent("click",true,false,h,0,0,0,0,0,false,false,false,false,0,null);n.dispatchEvent(u);w.readyState=w.DONE;w.savedAs=w.SAVEDASBLOB;B();return}if(h.chrome&&C&&C!==c){D=t.slice||t.webkitSlice;t=D.call(t,0,t.size,c);F=true}if(o&&v!=="download"){v+=".download"}if(C===c||o){x=h}if(!p){E();return}k+=t.size;p(h.TEMPORARY,k,A(function(G){G.root.getDirectory("saved",z,A(function(H){var I=function(){H.getFile(v,z,A(function(J){J.createWriter(A(function(K){K.onwriteend=function(L){x.location.href=J.toURL();b.push(J);w.readyState=w.DONE;w.savedAs=w.SAVEDASBLOB;q(w,"writeend",L)};K.onerror=function(){var L=K.error;if(L.code!==L.ABORT_ERR){E()}};"writestart progress write abort".split(" ").forEach(function(L){K["on"+L]=w["on"+L]});K.write(t);w.abort=function(){K.abort();w.readyState=w.DONE;w.savedAs=w.FAILED};w.readyState=w.WRITING}),E)}),E)};H.getFile(v,{create:false},A(function(J){J.remove();I()}),A(function(J){if(J.code===J.NOT_FOUND_ERR){I()}else{E()}}))}),E)}),E)},d=f.prototype,a=function(s,t){return new f(s,t)};d.abort=function(){var s=this;s.readyState=s.DONE;s.savedAs=s.FAILED;q(s,"abort")};d.readyState=d.INIT=0;d.WRITING=1;d.DONE=2;d.FAILED=-1;d.SAVEDASBLOB=1;d.SAVEDASURI=2;d.SAVEDASUNKNOWN=3;d.error=d.onwritestart=d.onprogress=d.onwrite=d.onabort=d.onerror=d.onwriteend=null;h.addEventListener("unload",i,false);a.unload=function(){i();h.removeEventListener("unload",i,false)};return a}(typeof self!=="undefined"&&self||typeof window!=="undefined"&&window||this.content));
