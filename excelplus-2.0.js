/*!
 * ExcelPlus JavaScript Library v2.0
 *
 * Copyright 2013, aymeric@kodono.info
 * Licensed under GPL Version 3 licenses.
 *
 * <script type="text/javascript" src="../Global%20Documents/js/xlsx.js"></script>
 * <script type="text/javascript" src="../Global%20Documents/js/jszip/jszip-all.min.js"></script>
 * <script type="text/javascript" src="../Global%20Documents/js/excelplus-2.0.js"></script>
 * <object id="file-excel" />
 * <script>
 * var oExcel = new ExcelPlus();
 * // create a button to open a remote file
 * oExcel.openRemote({
 *   idButton:'file-excel',
 *   flashPath:'/some/where/swfobject/',
 *   readCallback:function(name,base64) {
 *     console.log(this.getSheetNames()); // list of the worksheets
 *     // if there are several sheets you need to define which one to read:
 *     if (this.nbSheets > 1) this.activeSheet(1); // the first sheet is the one to read
 *     // you can show the sheet names with : `this.getSheetNames()`
 *     console.log("Value in the cell A2 - "+this.read("A2"));
 *     console.log("Values in the cells A1:B3 - "+this.read("A1:B3"));
 *     // we can also get the full sheet content into a 2D-array with `this.readAll()`
 *   }
 * })
 * </script>
 */

// Global variables
var __ExcelPlus_flashPath="";

/**
  * This is the master function for playing with a JS file
  *
  * @param {Object} params
  *   @param {Boolean} [params.useActiveX=true] We can force the file to use ActiveX when it's available
  *
  * @example
  * // note: the .error contains some error messages
  */
function ExcelPlus(params) {
  // variables used to design when writing into an Excel file
  this.xlStyles = [];
  this.xlStyles["xlAutomatic"] = -4105;
   
  this.xlStyles["center"] = -4108;
  this.xlStyles["middle"] = -4108;
  this.xlStyles["left"]   = -4131;
  this.xlStyles["right"]  = -4152;
  this.xlStyles["bottom"] = -4107;
  this.xlStyles["none"]   = -4142;
   
  this.xlSides = [];
  this.xlSides["DiagonalDown"] = 5;
  this.xlSides["DiagonalUp"] = 6;
  this.xlSides["left"] = 7;
  this.xlSides["top"] = 8;
  this.xlSides["bottom"] = 9;
  this.xlSides["right"] = 10;
  this.xlSides["InsideVertical"] = 11;
  this.xlSides["InsideHorizontal"] = 12;
   
  this.xlStyles["thin"] = 2;
  this.xlStyles["thick"] = 4;
  this.xlStyles["solid"] = 1; // xlContinous
  
  this.aLetters = [];
  this.aLetters["A"]=1, this.aLetters["B"]=2, this.aLetters["C"]=3, this.aLetters["D"]=4, this.aLetters["E"]=5, this.aLetters["F"]=6, this.aLetters["G"]=7, this.aLetters["H"]=8, this.aLetters["I"]=9, this.aLetters["J"]=10, this.aLetters["K"]=11, this.aLetters["L"]=12, this.aLetters["M"]=13, this.aLetters["N"]=14, this.aLetters["O"]=15, this.aLetters["P"]=16, this.aLetters["Q"]=17;
  this.aLetters["R"]=18, this.aLetters["S"]=19, this.aLetters["T"]=20, this.aLetters["U"]=21, this.aLetters["V"]=22, this.aLetters["W"]=23, this.aLetters["X"]=24, this.aLetters["Y"]=25, this.aLetters["Z"]=26;
  
  this.oExcel = null; // Excel object
  this.oSheet = null; // the current sheet that we are working on
  this.nbSheets = 0; // number of sheets
  this.oFile = null; // the File object
  this.nbRows = 0, this.nbColumns = 0; // the number of rows and columns used in the opened file
  this.filename = ""; // name of the file we read
  
  // check if we can use ActiveX here
  this.canActiveX = (params.useActiveX !== false ? this._init() : false);

  params = params || {};
  this.useActiveX = (params.useActiveX == undefined ? true : params.useActiveX); // we can force the use of ActiveX
  if (!this.canActiveX) this.useActiveX = false; // except if it's not available
  this.error = "";
}

var ExcelPlus_savedObjects = []; // this variable is used to close the opened Excel when we use ActiveX -- so no need to call .close()
ExcelPlus.prototype = {
  _init:function() {
    this._end();
    try {
      this.oExcel = new ActiveXObject("Excel.Application");
      ExcelPlus_savedObjects.push(this);
      return true;
    } catch(e) {
      this.error = "Unable to active the XObjet; the ActiveX must be desactivated, or you're not using IE";
      return false;
    }
  },
  _end:function() {
    this.oFile=null;
    this.oWorkbooks=null;
    this.nbSheets=0;
    this.nbRows=0;
    this.nbColumns=0;
    this.oSheet=null;
    this.oExcel=null;
  },
  // @param closeFile (default: false) will close the Excel window if it's true
  close:function(closeFile) {
    if (this.useActiveX) {
      if (typeof closeFile == "undefined") closeFile = false;
      if (this.oFile != null  && closeFile)  this.oFile.Close(false);
      if (this.oExcel != null && closeFile) this.oExcel.Quit();
      this._end();
    }
  },
  // create a new Excel file empty
  create:function(visible) {
    visible = (visible == undefined ? true : visible );
    if (this.oExcel==null) {
      if (this._init() == false) return false;
    }
    this.oExcel.Visible=visible;
    var newWorkbook=this.oExcel.Workbooks.Add;
    newWorkbook.WorkSheets(1).Activate;
    this.oSheet=newWorkbook.WorkSheets(1);
  },
  /**
    * This function permits to read a remote file in creating a upload button that will work with different browser versions
    *
    * @param {Object} params
    *   @param {String} [params.flashPath="./"] You must provide the path to the Flash files (swfobject.js, FileToDataURI.swf and expressInstall.swf) 
    *   @param [String} params.idButton The ID of the HTML button that we'll use to get the file
    *   @param {Function} [params.readCallback=function(name,base64){}] The function that will be called when the file is read with two parameters (one: the file name, two: the base64 string of the file)
    *   @param {Boolean} [visible=false] If we use ActiveX then it's possible to make the Excel file we currently read visible on the screen
    */
  openRemote:function(params) {
    params = params || {};
    __ExcelPlus_flashPath = params.flashPath || "";   // related to Flash features
    if (!params.idButton) {
      this.error = "You need to define the 'idButton' parameter!"
      return this
    }
  
    params.readCallback = params.readCallback || function() {};
    var _this=this;
    if (this.useActiveX) {
      params.visible = (params.visible == undefined ? false : params.visible);
      // So we use a regular input file
      $("#"+params.idButton).replaceWith('<input id="'+params.idButton+'" type="file" />');
      $("#"+params.idButton).on('change',function(e) {
        _this.filename = this.value.replace(/\\/g,"/").split("/");
        _this.filename = _this.filename[_this.filename.length-1];
        _this.open(this.value, params.visible);
        params.readCallback.call(_this,_this.filename,"");
      });
    } else {
      // we want to create a button to select the file
      // then the file will be read with Flash or FileAPI
      Flash.createButton(params.idButton, function(name,base64) {
        _this.filename = name;
        // then call our function
        _this.oFile = xlsx(base64);
        _this.nbSheets = _this.oFile.worksheets.length;
        if (_this.nbSheets == 1) _this.selectSheet(1);
        // call the function of the user
        params.readCallback.call(_this,name,base64);
      })
    }
  },
  // open an excepting Excel file
  // @param path is the path to the file
  // @param visible (default: false) show the opened file on the screen
  open:function(path, visible) {
    var check = true;
    if (this.oExcel==null) check=this._init();
    if (check == false) return false;
    if (arguments.length != 2) visible = false;
    
    this.oExcel.Visible = visible;
    var oWorkbooks      = this.oExcel.Workbooks;
    this.oFile          = oWorkbooks.open(path);
    // number of sheets
    this.nbSheets       = this.oFile.WorkSheets.Count;
    if (this.nbSheets == 1) this.selectSheet(1);
  },
  // return an array of the sheet names (the key is the sheet #)
  getSheetNames:function() {
    if (this.oFile == null) {
      this.error = "No Excel file opened.";
      return false;
    }
    
    var arr = [];
    for (var i=1; i <= this.nbSheets; i++) arr.push(this.useActiveX ? this.oFile.WorkSheets(i).Name : this.oFile.worksheets[i-1].name);
    return arr;
  },
  // select the specific sheet
  // @param sheet is a number (from 1 to X) that reprents the sheet in the Excel file
  //        ==> if sheet is not a number, then it must be the name of the sheet
  selectSheet:function(sheet) {
    if (this.oFile == null) {
      this.error = "No Excel file opened.";
      return false;
    }
    
    // it's not a number so search for the corresping #
    if (isNaN(sheet)) {
      var arr = this.getSheetNames();
      for (var i=1; i <= this.nbSheets; i++) {
        if (sheet.toLowerCase() == arr[i].toLowerCase()) {
          sheet = i;
          break;
        }
      }
    }
    
    if (sheet == 0 || sheet > this.nbSheets || isNaN(sheet)) {
      this.error = "The sheet # is uncorrect.";
      return false;
    }

    if (!this.useActiveX) sheet--;

    // select the sheet
    this.oSheet    = (this.useActiveX ? this.oFile.WorkSheets(sheet) : this.oFile.worksheets[sheet].data);
    var usedRange  = (this.useActiveX ? this.oSheet.UsedRange : this.oSheet); // TODO
    this.nbRows    = (this.useActiveX ? usedRange.Rows.Count : this.oSheet.length);
    this.nbColumns = (this.useActiveX ? usedRange.Columns.Count : this.oSheet[0].length);
  },
  // read the all usedRange
  // we assume that the content starts at A1
  // return if it's a range then it returns a 2D array; the first one is the rows, and the second one is the columns
  //        for example: array[0][2] <=> first row, third column
  readAll:function() {
    if (this._checkSheet()==false) return false;
    
    var table = [],line,val;
    for (var row=1; row <= this.nbRows; row++) {
      line = [];
      for (var col=1; col <= this.nbColumns; col++) {
        val="";
        if (this.useActiveX) val = this.oSheet.Cells(row,col).value;
        else if (this.oSheet[row-1] && this.oSheet[row-1][col-1]) val = this.oSheet[row-1][col-1].value;
        if (val == undefined) val = "";
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
      if (this.useActiveX) val=this.oSheet.Cells(cell.row,cell.column).value;
      else if (this.oSheet[cell.row-1] && this.oSheet[cell.row-1][cell.column-1]) val = this.oSheet[cell.row-1][cell.column-1].value;
      if (val == undefined) val = "";
      return val;
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
      return false;
    }
    return {"column": this._getColumn(arr[1]), "row": arr[2]};
  },
  // see the index color at http://www.mvps.org/dmcritchie/excel/colors.htm
  _getColorIndex:function(color) {
    if (!isNaN(color)) return color;
    switch (color) {
      case "black":     return 1;
      case "white":	    return 2;
      case "red":       return 3;
      case "green":     return 4;
      case "darkblue":  return 5;
      case "blue":	    return 23;
      case "yellow":    return 6;
      case "magenta":   return 7;
      case "cyan":      return 8;
      case "purple":    return 13;
      case "lightgray": return 15;
      case "gray":      return 16;
      case "orange":    return 45;
      default: this.error = "'"+color+"' is an unknown color. Known colors are: black, white, red, green, blue, yellow, magenta and cyan";
               return false;
    }
  },
  _isRange:function(range) { return (range.indexOf(":") != -1 && range.match(/[1-9]/) != null); },
  _isCell:function(range) { return (range.indexOf(":") == -1 && range.match(/[1-9]/) != null); },
  _isColumn:function(range) { return (range.match(/[1-9]/) == null); },
  _checkSheet:function() {
    if (this.oSheet==null) {
      this.error = "The sheet is not specified";
      return false;
    }
  },
  // we use the CSS codes to change the style of the cells (http://www.mvps.org/dmcritchie/excel/font.htm)
  // @param range is a range, a cell or a column (example: "A1", "A1:A3", "A", "A:B")
  // @param hashstyle is the CSS style in the hash format (example: {"width":30, "text-align":"right"})
  setStyle:function(range, hashstyle) {
    if (this._checkSheet()==false) return false;
    
    // range can be a range, a single cell or a column
    var isRange  = this._isRange(range);
    var isCell   = this._isCell(range);
    var isColumn = this._isColumn(range);
    var column   = "";
  	var row      = "";
  	var cell     = "";
	
    if (!isCell && !isRange && !isColumn) {
      this.error = "The range provided is not a range, neither a cell, neither a column!";
      return false;
    }
    
    // if it's a cell, we find the coord
    if (isCell) {
      cell   = this._getCellCoord(range);
      if (cell === false) return false;
	  column = this._getColumnName(cell.column)+":"+this._getColumnName(cell.column);
	  row    = cell.row+":"+cell.row;
    }
	
    if (isRange) {
      var tmp = range.match(/([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)/);
	  if (tmp == null || tmp.length != 5) return false;
	  column  = tmp[1]+":"+tmp[3];
	  row     = tmp[2]+":"+tmp[4];
    }
    
    // build an array based on the {}
    var style = [];
    for (var key in hashstyle) {
      if (key == "border-style") {
        style["border-style-top"]    = hashstyle[key]
        style["border-style-right"]  = hashstyle[key]
        style["border-style-bottom"] = hashstyle[key]
        style["border-style-left"]   = hashstyle[key]
        style.length += 4;
      } else if (key == "border-color") {
        style["border-color-top"]    = hashstyle[key]
        style["border-color-right"]  = hashstyle[key]
        style["border-color-bottom"] = hashstyle[key]
        style["border-color-left"]   = hashstyle[key]      
      }else {
        style[key] = hashstyle[key];
        style.length++;
      }
    }
    
    // apply the style
    if (style["width"] != undefined) {
      if (style["width"] == -1)
        this.oSheet.Columns(column).AutoFit();
      else
        this.oSheet.Columns(column).ColumnWidth = style["width"];
    }
    if (style["max-width"] != undefined) {
      if (this.oSheet.Columns(column).Width > style["max-width"])
        this.oSheet.Columns(column).ColumnWidth = style["max-width"];
    }
    if (style["height"] != undefined) {
      if (!isColumn) this.oSheet.Rows(row).RowHeight = style["height"];
    }
    if (style["text-align"] != undefined) {
      if (this.xlStyles[style["text-align"]] == undefined) this.error = "Property '"+style["text-align"]+"' is not available for 'text-align'";
      else {
        if (isRange)       this.oSheet.Range(range).HorizontalAlignment   = this.xlStyles[style["text-align"]];
        else if (isCell)   this.oSheet.Cells(cell.row, cell.column).HorizontalAlignment = this.xlStyles[style["text-align"]];
        else if (isColumn) this.oSheet.Columns(column).HorizontalAlignment = this.xlStyles[style["text-align"]];
      }
    }
    if (style["vertical-align"] != undefined) {
      if (this.xlStyles[style["vertical-align"]] == undefined) this.error = "Property '"+style["vertical-align"]+"' is not available for 'vertical-align'";
      else {
        if (isRange)       this.oSheet.Range(range).VerticalAlignment   = this.xlStyles[style["vertical-align"]];
        else if (isCell)   this.oSheet.Cells(cell.row, cell.column).VerticalAlignment = this.xlStyles[style["vertical-align"]];
        else if (isColumn) this.oSheet.Columns(column).VerticalAlignment = this.xlStyles[style["vertical-align"]];
      }
    }
    if (style["font-weight"] != undefined) {
      var isBold = (style["font-weight"] == "bold");
      if (isRange)       this.oSheet.Range(range).Font.Bold   = isBold;
      else if (isCell)   this.oSheet.Cells(cell.row, cell.column).Font.Bold = isBold;
      else if (isColumn) this.oSheet.Columns(column).Font.Bold = isBold;
    }
    if (style["font-style"] != undefined) {
      var isItalic = (style["font-style"] == "italic");
      if (isRange)       this.oSheet.Range(range).Font.Italic   = isItalic;
      else if (isCell)   this.oSheet.Cells(cell.row, cell.column).Font.Italic = isItalic;
      else if (isColumn) this.oSheet.Columns(column).Font.Italic = isItalic;
    }
    if (style["text-decoration"] != undefined) {
      var isUnderline = (style["text-decoration"] == "underline");
      if (isRange)       this.oSheet.Range(range).Font.Underline   = isUnderline;
      else if (isCell)   this.oSheet.Cells(cell.row, cell.column).Font.Underline = isUnderline;
      else if (isColumn) this.oSheet.Columns(column).Font.Underline = isUnderline;
    }
    if (style["font-family"] != undefined) {
      if (isRange)       this.oSheet.Range(range).Font.Name   = style["font-family"];
      else if (isCell)   this.oSheet.Cells(cell.row, cell.column).Font.Name = style["font-family"];
      else if (isColumn) this.oSheet.Columns(column).Font.Name = style["font-family"];
    }
    if (style["color"] != undefined) {
      var indexColor = this._getColorIndex(style["color"]);
      if (indexColor===false) {
        alert("The color '"+style["color"]+"' is not available");
        return false;
      }
      if (isRange)       this.oSheet.Range(range).Font.ColorIndex   = indexColor;
      else if (isCell)   this.oSheet.Cells(cell.row, cell.column).Font.ColorIndex = indexColor;
      else if (isColumn) this.oSheet.Columns(column).Font.ColorIndex = indexColor;
    }
    if (style["background-color"] != undefined) {
      var indexColor = this._getColorIndex(style["background-color"]);
      if (indexColor===false) {
        alert("The color '"+style["background-color"]+"' is not available");
        return false;
      }
      if (isRange)       this.oSheet.Range(range).Interior.ColorIndex   = indexColor;
      else if (isCell)   this.oSheet.Cells(cell.row, cell.column).Interior.ColorIndex = indexColor;
      else if (isColumn) this.oSheet.Columns(column).Interior.ColorIndex = indexColor;
    }
    if (style["font-size"] != undefined) {
      if (isRange)       this.oSheet.Range(range).Font.Size   = style["font-size"];
      else if (isCell)   this.oSheet.Cells(cell.row, cell.column).Font.Size = style["font-size"];
      else if (isColumn) this.oSheet.Columns(column).Font.Size = style["font-size"];
    }
    if (style["white-space"] != undefined) {
      var doWrap = (style["white-space"]!="nowrap");
      if (isRange)       this.oSheet.Range(range).WrapText   = doWrap;
      else if (isCell)   this.oSheet.Cells(cell.row, cell.column).WrapText = doWrap;
      else if (isColumn) this.oSheet.Columns(column).WrapText = doWrap;
    }
    // the number format
    // for example: "dd/mm/yyyy hh:mm:ss", "0", "0.00"
    if (style["number-format"] != undefined) {
      if (isRange)       this.oSheet.Range(range).NumberFormat   = style["number-format"];
      else if (isCell)   this.oSheet.Cells(cell.row, cell.column).NumberFormat = style["number-format"];
      else if (isColumn) this.oSheet.Columns(column).NumberFormat = style["number-format"];
    }
    
    var aBorders = ["top", "right", "bottom", "left"];
    for (var i=0; i < aBorders.length; i++) {
      if (style["border-style-"+aBorders[i]] != null) {
        if (this.xlStyles[style["border-style-"+aBorders[i]]] == undefined) this.error = "Property '"+style["border-style-"+aBorders[i]]+"' is not available for 'border-style-"+aBorders[i]+"'";
        else {
          var side  = this.xlSides[aBorders[i]];
          if (isRange)       this.oSheet.Range(range).Borders(side).LineStyle   = this.xlStyles[style["border-style-"+aBorders[i]]];
          else if (isCell)   this.oSheet.Cells(cell.row, cell.column).Borders(side).LineStyle = this.xlStyles[style["border-style-"+aBorders[i]]];
          else if (isColumn) this.oSheet.Columns(column).Borders(side).LineStyle = this.xlStyles[style["border-style-"+aBorders[i]]];
        }
      }
      if (style["border-color-"+aBorders[i]] != null) {
        if (this.xlStyles[style["border-color-"+aBorders[i]]] == undefined) this.error = "Property '"+style["border-color-"+aBorders[i]]+"' is not available for 'border-color-"+aBorders[i]+"'";
        else {
          var side  = this.xlSides[aBorders[i]];
          var indexColor = this._getColorIndex(style["border-color-"+aBorders[i]]);
          if (isRange)       this.oSheet.Range(range).Borders(side).ColorIndex   = style["border-color-"+aBorders[i]];
          else if (isCell)   this.oSheet.Cells(cell.row, cell.column).Borders(side).ColorIndex = style["border-color-"+aBorders[i]];
          else if (isColumn) this.oSheet.Columns(column).Borders(side).ColorIndex = style["border-color-"+aBorders[i]];
        }
      }
    }
  },
  // create a grid in the range
  grid:function(range, color) {
    if (this._checkSheet()==false) return false;
    if (this._isCell(range)===true) this.setStyle(range, {"border-style":"solid"});
    if (this._isRange(range)===false) return false;
    
    this.oSheet.Range(range).Borders.LineStyle = this.xlStyles["solid"]
    if (color != undefined)
      this.oSheet.Range(range).Borders.ColorIndex = this._getColorIndex(color);
  },
  // merge a range of cells
  merge:function(range) {
    if (this._checkSheet()==false) return false;
    if (this._isRange(range)===false) return false;
    this.oSheet.Range(range).Merge();
  },
  // write a value into a cell
  // the value can also be a formula, e.g. "=SUM(A1:B1)"
  writeCell:function(cell, value) {
    if (this._checkSheet()==false) return false;
    // find the corresponding cell
    var coord = this._getCellCoord(cell);
    if (coord === false) return false;
    
    // write the value
    this.oSheet.Cells(coord.row, coord.column).value = value;
  },
  // return the first cell after the used rows
  getNewLine:function() {
    if (this._checkSheet()==false) return false;
    this.nbColumns++;
    return "A"+this.nbColumns;
  },
  // write something in the Excel status bar
  statusBar:function(text) {
    this.oExcel.StatusBar = text;
  },
  show:function() {
    this.oExcel.Visible = true;
  },
  hide:function() {
    this.oExcel.Visible = false;
  }
}

// we want to make sure to close the opened file if we use ActiveX
window.onbeforeunload = function (e) {
  for (var i=0; i < ExcelPlus_savedObjects.length; i++) {
    try { ExcelPlus_savedObjects[i].close() }
    catch(e) { continue }
  }
};

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
    
    // check if the FileReader API exists... if not then load the Flash object
    if (typeof FileReader !== "function") {
      // we call swfobject.js
      if (swfobject == undefined) {
        var fileref=document.createElement('script');
        fileref.setAttribute("type","text/javascript")
        fileref.setAttribute("src", __ExcelPlus_flashPath+"/swfobject.js");
        document.getElementsByTagName("head")[0].appendChild(fileref);
        // wait for the file to be loaded
        return this._waitForObjectToBeLoaded(id,callback);
      } else
        swfobject.embedSWF(__ExcelPlus_flashPath+"/FileToDataURI.swf", id, "100px", "40px", "10", __ExcelPlus_flashPath+"/expressInstall.swf", {}, {}, {});
    } else {
      // replace the <object> by an input file
      $("#"+id).replaceWith('<input id="'+id+'" type="file" />');
      $("#"+id).on('change',function(e) {
        // we use the normal way to read a file with FileReader
        var files = e.target.files,file;
    
        if (!files || files.length == 0) return;
        file = files[0];

        var fileName=$(this).val(); 
        var fileReader = new FileReader();
        fileReader.onload = function (e) {
          Flash.getFileData(fileName,e.target.result.split(",")[1]);
        };
        fileReader.readAsDataURL(file);
      });
    }
  },
  /**
    * This function is called when we have read the file
    *
    * @param {String} name It's the filename
    * @param {String} base64 It's the base64 version of the file
    */
  getFileData: function(name, base64) {
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
