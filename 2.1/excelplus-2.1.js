/*!
 * ExcelPlus JavaScript Library v2.1
 *
 * Copyright 2014, aymeric@kodono.info
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
  * This is the master function for playing with a JS file
  *
  * @param {Object} params
  *   @param {Boolean} [params.showErrors=true] The errors will be shown in the console
  *
  */
function ExcelPlus(params) {
  this.aLetters = [];
  this.aLetters["A"]=1, this.aLetters["B"]=2, this.aLetters["C"]=3, this.aLetters["D"]=4, this.aLetters["E"]=5, this.aLetters["F"]=6, this.aLetters["G"]=7, this.aLetters["H"]=8, this.aLetters["I"]=9, this.aLetters["J"]=10, this.aLetters["K"]=11, this.aLetters["L"]=12, this.aLetters["M"]=13, this.aLetters["N"]=14, this.aLetters["O"]=15, this.aLetters["P"]=16, this.aLetters["Q"]=17;
  this.aLetters["R"]=18, this.aLetters["S"]=19, this.aLetters["T"]=20, this.aLetters["U"]=21, this.aLetters["V"]=22, this.aLetters["W"]=23, this.aLetters["X"]=24, this.aLetters["Y"]=25, this.aLetters["Z"]=26;
  
  this.oExcel = null; // Excel object
  this.oSheet = null; // the current sheet that we are working on
  this.nbSheets = 0; // number of sheets
  this.oFile = null; // the File object
  this.nbRows = 0, this.nbColumns = 0; // the number of rows and columns used in the opened file
  this.filename = ""; // name of the file we read
  this.usedRange = ""; // e.g. A1-C250

  params = params || {};
  this.error = "";

  // to show the errors if any, otherwise you can handle the errors using YourObject.errors
  this.showErrors = (params.showErrors === undefined ? true : params.showErrors);
}

ExcelPlus.prototype = {
  reset:function() {
    this.oFile=null;
    this.oWorkbooks=null;
    this.nbSheets=0;
    this.nbRows=0;
    this.nbColumns=0;
    this.usedRange="";
    this.oSheet=null;
    this.oExcel=null;
  },
  // create a new Excel file empty
  create:function(visible) {
  },
  /**
   * This function permits to read a local file (on the user's computer)
   * To do so, we'll use a input[file] -- integration of the code from http://aymkdn.github.io/FileToDataURI/
   * The HTML code for the input[file] will be <object id="file-object"></object>
   *
   * @param {Object} params
   *   @param {String} [params.idButton="file-object"] This is the ID of the <object> tag in your code
   *   @param {String} [params.labelButton="Load a file"] This is the label for the button
   *   @param {String} [params.flashPath="./"] You must provide the path to the Flash files (swfobject.js, FileToDataURI.swf and expressInstall.swf) 
   * @param {Function} params.onReady=function(){} The function that will be called when the file has been loaded
   */
  openLocal:function(params, onReady) {
    params = params || {};
    params.idButton = params.idButton || "file-object";
    if (document.getElementById(params.idButton) === null) {
      this.error = "You must have a <object> tag in your code.";
      this._showErrors();
      return this;
    }

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
          _this.oFile = XLSX.read(e.target.result, {type: 'binary'});
          _this.nbSheets = _this.oFile.SheetNames.length;
          if (_this.nbSheets == 1) _this.selectSheet(1);
          onReady.call(_this, true);
        };
        fileReader.readAsBinaryString(file);
      });
    } else {
      // we call Flash to be able to read the file
      Flash.createButton(params.idButton, function(name,base64) {
        _this.filename = name;
        // then call our function
        _this.oFile = XLSX.read(base64, {type: 'base64'});
        _this.nbSheets = _this.oFile.SheetNames.length;
        if (_this.nbSheets == 1) _this.selectSheet(1);
        onReady.call(_this, true);
      })
    }
  },
  /**
   * This function read a remote Excel file
   *
   * @param {String} url The url of the Excel file
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
          _this.oFile = XLSX.read(data, {type: 'binary'});
          _this.nbSheets = _this.oFile.SheetNames.length;
          if (_this.nbSheets == 1) _this.selectSheet(1);
          onReady.call(_this, true);
        } else {
          _this.error = "Cannot find the remote file";
          _this._showErrors();
          onReady.call(_this, false)
        }
      }
    };
    xhr.send(null);
  },
  // return an array of the sheet names (the key is the sheet #)
  getSheetNames:function() {
    if (this.oFile == null) {
      this.error = "No Excel file opened.";
      this._showErrors();
      return false;
    }
    
    return this.oFile.SheetNames;
  },
  // select the specific sheet
  // @param sheet is a number (from 1 to X) that reprents the sheet in the Excel file
  //        ==> if sheet is not a number, then it must be the name of the sheet
  selectSheet:function(sheet) {
    if (this.oFile == null) {
      this.error = "No Excel file opened.";
      this._showErrors();
      return false;
    }
    
    if (!isNaN(sheet)) {
      // find the sheet name based on the number
      var arr = this.getSheetNames();
      for (var i=0; i < arr.length; i++) {
        if (sheet == i) {
          sheet = arr[i];
          break;
        }
      }
    }

    this.oSheet    = this.oFile.Sheets[sheet];
    this.usedRange = this.oSheet["!ref"]; // e.g. A1-C250
    var lastCell   = this._getCellCoord(this.usedRange.split(":")[1]);
    this.nbRows    = lastCell.row * 1;
    this.nbColumns = lastCell.column * 1;
  },
  // read the all usedRange
  // we assume that the content starts at A1
  // return if it's a range then it returns a 2D array; the first one is the rows, and the second one is the columns
  //        for example: array[0][2] <=> first row, third column
  readAll:function() {
    if (this._checkSheet()==false) {
      this.error = "You have to select a sheet before using `readAll`";
      this._showErrors();
      return false;
    }
    
    var table = [],line,val,cell;
    for (var row=1; row <= this.nbRows; row++) {
      line = [];
      for (var col=1; col <= this.nbColumns; col++) {
        val="";
        cell=this._getColumnName(col) + "" + row;
        if (this.oSheet[cell]) val = this.oSheet[cell].v;
        line.push(val);
      }
      table.push(line);
    }
    return table;
  },
  // @param range can be a range, or a single cell
  // return
  //      ==> if it's a range then it returns a 2D array; the first one is the rows, and the second one is the columns
  //      for example: array[0][2] <=> first row, third column
  //      ==> if it's a cell then it returns the value of the cell
  read:function(range) {
    if (this._checkSheet()==false) return false

    if (this._isRange(range)) {
      var arr = range.split(":");
      // find the first and the last cell
      var firstCell = this._getCellCoord(arr[0]);
      var lastCell  = this._getCellCoord(arr[1]);
      // TODO with no ActiveX      
      var table = [], line, val;
      for (var row=firstCell.row; row <= lastCell.row; row++) {
        line = [];
        for (var col=firstCell.column; col <= lastCell.column; col++) {
          val="";
          if (this.useActiveX) val=this.oSheet.Cells(row,col).value;
          else if (this.oSheet[row-1] && this.oSheet[row-1][col-1]) val = this.oSheet[row-1][col-1].value;
          if (val == undefined) val = "";
          line.push(val);
        }
        table.push(line);
      }
      return table;
    }
    else if (this._isCell(range)) {
      var cell = this._getCellCoord(range);
      var val  = "";
      return this.oSheet[range].v;
    }
    return false;
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
  // we won't count after ZZ
  // @param index is a number
  // return the column name based on its number
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
  // return the coord of a cell based on its name
  // the returned hash is : {column: Y, row: X}
  _getCellCoord:function(cell) {
    var arr = cell.match(/([A-Z]+)([0-9]+)/);
    if (arr == null || arr.length != 3) {
      this.error = "The format of the cell is wrong.";
      this._showErrors();
      return false;
    }
    return {"column": this._getColumn(arr[1]), "row": arr[2]};
  },
  // see the index color at http://www.mvps.org/dmcritchie/excel/colors.htm
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
      default: this.error = "'"+color+"' is an unknown color. Known colors are: black, white, red, green, blue, yellow, magenta and cyan";
               this._showErrors();
               return false;
    }
  },
  _isRange:function(range) { return (range.indexOf(":") != -1 && range.match(/[1-9]/) != null); },
  _isCell:function(range) { return (range.indexOf(":") == -1 && range.match(/[1-9]/) != null); },
  _isColumn:function(range) { return (range.match(/[1-9]/) == null); },
  _checkSheet:function() {
    if (this.oSheet==null) {
      this.error = "The sheet is not specified";
      this._showErrors();
      return false;
    }
  },
  _showErrors:function() {
    if (this.showErrors) {
      if (this.error !== "") console.error("[ExcelPlus] "+this.error);
      else console.log("[ExcelPlus] No Error");
    }
  },
  // we use the CSS codes to change the style of the cells (http://www.mvps.org/dmcritchie/excel/font.htm)
  // @param range is a range, a cell or a column (example: "A1", "A1:A3", "A", "A:B")
  // @param hashstyle is the CSS style in the hash format (example: {"width":30, "text-align":"right"})
  setStyle:function(range, hashstyle) {},
  // create a grid in the range
  grid:function(range, color) {},
  // merge a range of cells
  merge:function(range) {},
  // write a value into a cell
  // the value can also be a formula, e.g. "=SUM(A1:B1)"
  writeCell:function(cell, value) {},
  // return the first cell after the used rows
  getNewLine:function() {},
  // write something in the Excel status bar
  statusBar:function(text) {}
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
    * @param {Function} callback The function that will be called when the file is read with two parameters "the file name" and the "base64" version of the file
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
      swfobject.embedSWF(__ExcelPlus_flashPath+"/FileToDataURI.swf", id, "80px" /* width of the Flash zone */, "22px" /* height of the Flash zone */, "10", __ExcelPlus_flashPath+"/expressInstall.swf", {}, {}, {});
    }
  },
  /**
    * This function is called when we have read the file
    *
    * @param {String} name It's the filename
    * @param {String} base64 It's the base64 version of the file
    */
  getFileData: function(name, base64) {
    if (base64===undefined) { base64=name; name="Unknown Name" }
    __Flash_getFileData_callback(name,base64);
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