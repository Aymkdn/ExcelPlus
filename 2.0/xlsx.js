if (!Date.prototype.toISOString) {

    (function() {

        function pad(number) {
            var r = String(number);
            if (r.length === 1) {
                r = '0' + r;
            }
            return r;
        }

        Date.prototype.toISOString = function() {
            return this.getUTCFullYear()
                    + '-'
                    + pad(this.getUTCMonth() + 1)
                    + '-'
                    + pad(this.getUTCDate())
                    + 'T'
                    + pad(this.getUTCHours())
                    + ':'
                    + pad(this.getUTCMinutes())
                    + ':'
                    + pad(this.getUTCSeconds())
                    + '.'
                    + String((this.getUTCMilliseconds() / 1000).toFixed(3))
                            .slice(2, 5) + 'Z';
        };

    }());
}
if (!Date.now) {
    Date.now = function() {
        return new Date().valueOf();
    }
}

if (!Array.prototype.indexOf) {
    Array.prototype.indexOf = function(searchElement /* , fromIndex */) {
        'use strict';
        if (this == null) {
            throw new TypeError();
        }
        var n, k, t = Object(this), len = t.length >>> 0;

        if (len === 0) {
            return -1;
        }
        n = 0;
        if (arguments.length > 1) {
            n = Number(arguments[1]);
            if (n != n) { // shortcut for verifying if it's NaN
                n = 0;
            } else if (n != 0 && n != Infinity && n != -Infinity) {
                n = (n > 0 || -1) * Math.floor(Math.abs(n));
            }
        }
        if (n >= len) {
            return -1;
        }
        for (k = n >= 0 ? n : Math.max(len - Math.abs(n), 0); k < len; k++) {
            if (k in t && t[k] === searchElement) {
                return k;
            }
        }
        return -1;
    };
}
if (!window.JSON){
    /*! JSON v3.2.6 | http://bestiejs.github.io/json3 | Copyright 2012-2013, Kit Cambridge | http://kit.mit-license.org */
    ;(function(){var n=null;
    (function(G){function m(a){if(m[a]!==s)return m[a];var c;if("bug-string-char-index"==a)c="a"!="a"[0];else if("json"==a)c=m("json-stringify")&&m("json-parse");else{var e;if("json-stringify"==a){c=o.stringify;var b="function"==typeof c&&l;if(b){(e=function(){return 1}).toJSON=e;try{b="0"===c(0)&&"0"===c(new Number)&&'""'==c(new String)&&c(p)===s&&c(s)===s&&c()===s&&"1"===c(e)&&"[1]"==c([e])&&"[null]"==c([s])&&"null"==c(n)&&"[null,null,null]"==c([s,p,n])&&'{"a":[1,true,false,null,"\\u0000\\b\\n\\f\\r\\t"]}'==c({a:[e,
    !0,!1,n,"\x00\u0008\n\u000c\r\t"]})&&"1"===c(n,e)&&"[\n 1,\n 2\n]"==c([1,2],n,1)&&'"-271821-04-20T00:00:00.000Z"'==c(new Date(-864E13))&&'"+275760-09-13T00:00:00.000Z"'==c(new Date(864E13))&&'"-000001-01-01T00:00:00.000Z"'==c(new Date(-621987552E5))&&'"1969-12-31T23:59:59.999Z"'==c(new Date(-1))}catch(f){b=!1}}c=b}if("json-parse"==a){c=o.parse;if("function"==typeof c)try{if(0===c("0")&&!c(!1)){e=c('{"a":[1,true,false,null,"\\u0000\\b\\n\\f\\r\\t"]}');var j=5==e.a.length&&1===e.a[0];if(j){try{j=!c('"\t"')}catch(d){}if(j)try{j=
    1!==c("01")}catch(h){}if(j)try{j=1!==c("1.")}catch(k){}}}}catch(N){j=!1}c=j}}return m[a]=!!c}var p={}.toString,q,x,s,H=typeof define==="function"&&define.amd,y="object"==typeof JSON&&JSON,o="object"==typeof exports&&exports&&!exports.nodeType&&exports;o&&y?(o.stringify=y.stringify,o.parse=y.parse):o=G.JSON=y||{};var l=new Date(-3509827334573292);try{l=-109252==l.getUTCFullYear()&&0===l.getUTCMonth()&&1===l.getUTCDate()&&10==l.getUTCHours()&&37==l.getUTCMinutes()&&6==l.getUTCSeconds()&&708==l.getUTCMilliseconds()}catch(O){}if(!m("json")){var t=
    m("bug-string-char-index");if(!l)var u=Math.floor,I=[0,31,59,90,120,151,181,212,243,273,304,334],z=function(a,c){return I[c]+365*(a-1970)+u((a-1969+(c=+(c>1)))/4)-u((a-1901+c)/100)+u((a-1601+c)/400)};if(!(q={}.hasOwnProperty))q=function(a){var c={},e;if((c.__proto__=n,c.__proto__={toString:1},c).toString!=p)q=function(a){var c=this.__proto__,a=a in(this.__proto__=n,this);this.__proto__=c;return a};else{e=c.constructor;q=function(a){var c=(this.constructor||e).prototype;return a in this&&!(a in c&&
    this[a]===c[a])}}c=n;return q.call(this,a)};var J={"boolean":1,number:1,string:1,undefined:1};x=function(a,c){var e=0,b,f,j;(b=function(){this.valueOf=0}).prototype.valueOf=0;f=new b;for(j in f)q.call(f,j)&&e++;b=f=n;if(e)x=e==2?function(a,c){var e={},b=p.call(a)=="[object Function]",f;for(f in a)!(b&&f=="prototype")&&!q.call(e,f)&&(e[f]=1)&&q.call(a,f)&&c(f)}:function(a,c){var e=p.call(a)=="[object Function]",b,f;for(b in a)!(e&&b=="prototype")&&q.call(a,b)&&!(f=b==="constructor")&&c(b);(f||q.call(a,
    b="constructor"))&&c(b)};else{f=["valueOf","toString","toLocaleString","propertyIsEnumerable","isPrototypeOf","hasOwnProperty","constructor"];x=function(a,c){var e=p.call(a)=="[object Function]",b,g;if(g=!e)if(g=typeof a.constructor!="function"){g=typeof a.hasOwnProperty;g=g=="object"?!!a.hasOwnProperty:!J[g]}g=g?a.hasOwnProperty:q;for(b in a)!(e&&b=="prototype")&&g.call(a,b)&&c(b);for(e=f.length;b=f[--e];g.call(a,b)&&c(b));}}return x(a,c)};if(!m("json-stringify")){var K={92:"\\\\",34:'\\"',8:"\\b",
    12:"\\f",10:"\\n",13:"\\r",9:"\\t"},v=function(a,c){return("000000"+(c||0)).slice(-a)},D=function(a){var c='"',b=0,g=a.length,f=g>10&&t,j;for(f&&(j=a.split(""));b<g;b++){var d=a.charCodeAt(b);switch(d){case 8:case 9:case 10:case 12:case 13:case 34:case 92:c=c+K[d];break;default:if(d<32){c=c+("\\u00"+v(2,d.toString(16)));break}c=c+(f?j[b]:t?a.charAt(b):a[b])}}return c+'"'},B=function(a,c,b,g,f,j,d){var h,k,i,l,m,o,r,t,w;try{h=c[a]}catch(y){}if(typeof h=="object"&&h){k=p.call(h);if(k=="[object Date]"&&
    !q.call(h,"toJSON"))if(h>-1/0&&h<1/0){if(z){l=u(h/864E5);for(k=u(l/365.2425)+1970-1;z(k+1,0)<=l;k++);for(i=u((l-z(k,0))/30.42);z(k,i+1)<=l;i++);l=1+l-z(k,i);m=(h%864E5+864E5)%864E5;o=u(m/36E5)%24;r=u(m/6E4)%60;t=u(m/1E3)%60;m=m%1E3}else{k=h.getUTCFullYear();i=h.getUTCMonth();l=h.getUTCDate();o=h.getUTCHours();r=h.getUTCMinutes();t=h.getUTCSeconds();m=h.getUTCMilliseconds()}h=(k<=0||k>=1E4?(k<0?"-":"+")+v(6,k<0?-k:k):v(4,k))+"-"+v(2,i+1)+"-"+v(2,l)+"T"+v(2,o)+":"+v(2,r)+":"+v(2,t)+"."+v(3,m)+"Z"}else h=
    n;else if(typeof h.toJSON=="function"&&(k!="[object Number]"&&k!="[object String]"&&k!="[object Array]"||q.call(h,"toJSON")))h=h.toJSON(a)}b&&(h=b.call(c,a,h));if(h===n)return"null";k=p.call(h);if(k=="[object Boolean]")return""+h;if(k=="[object Number]")return h>-1/0&&h<1/0?""+h:"null";if(k=="[object String]")return D(""+h);if(typeof h=="object"){for(a=d.length;a--;)if(d[a]===h)throw TypeError();d.push(h);w=[];c=j;j=j+f;if(k=="[object Array]"){i=0;for(a=h.length;i<a;i++){k=B(i,h,b,g,f,j,d);w.push(k===
    s?"null":k)}a=w.length?f?"[\n"+j+w.join(",\n"+j)+"\n"+c+"]":"["+w.join(",")+"]":"[]"}else{x(g||h,function(a){var c=B(a,h,b,g,f,j,d);c!==s&&w.push(D(a)+":"+(f?" ":"")+c)});a=w.length?f?"{\n"+j+w.join(",\n"+j)+"\n"+c+"}":"{"+w.join(",")+"}":"{}"}d.pop();return a}};o.stringify=function(a,c,b){var g,f,j,d;if(typeof c=="function"||typeof c=="object"&&c)if((d=p.call(c))=="[object Function]")f=c;else if(d=="[object Array]"){j={};for(var h=0,k=c.length,i;h<k;i=c[h++],(d=p.call(i),d=="[object String]"||d==
    "[object Number]")&&(j[i]=1));}if(b)if((d=p.call(b))=="[object Number]"){if((b=b-b%1)>0){g="";for(b>10&&(b=10);g.length<b;g=g+" ");}}else d=="[object String]"&&(g=b.length<=10?b:b.slice(0,10));return B("",(i={},i[""]=a,i),f,j,g,"",[])}}if(!m("json-parse")){var L=String.fromCharCode,M={92:"\\",34:'"',47:"/",98:"\u0008",116:"\t",110:"\n",102:"\u000c",114:"\r"},b,A,i=function(){b=A=n;throw SyntaxError();},r=function(){for(var a=A,c=a.length,e,g,f,j,d;b<c;){d=a.charCodeAt(b);switch(d){case 9:case 10:case 13:case 32:b++;
    break;case 123:case 125:case 91:case 93:case 58:case 44:e=t?a.charAt(b):a[b];b++;return e;case 34:e="@";for(b++;b<c;){d=a.charCodeAt(b);if(d<32)i();else if(d==92){d=a.charCodeAt(++b);switch(d){case 92:case 34:case 47:case 98:case 116:case 110:case 102:case 114:e=e+M[d];b++;break;case 117:g=++b;for(f=b+4;b<f;b++){d=a.charCodeAt(b);d>=48&&d<=57||d>=97&&d<=102||d>=65&&d<=70||i()}e=e+L("0x"+a.slice(g,b));break;default:i()}}else{if(d==34)break;d=a.charCodeAt(b);for(g=b;d>=32&&d!=92&&d!=34;)d=a.charCodeAt(++b);
    e=e+a.slice(g,b)}}if(a.charCodeAt(b)==34){b++;return e}i();default:g=b;if(d==45){j=true;d=a.charCodeAt(++b)}if(d>=48&&d<=57){for(d==48&&(d=a.charCodeAt(b+1),d>=48&&d<=57)&&i();b<c&&(d=a.charCodeAt(b),d>=48&&d<=57);b++);if(a.charCodeAt(b)==46){for(f=++b;f<c&&(d=a.charCodeAt(f),d>=48&&d<=57);f++);f==b&&i();b=f}d=a.charCodeAt(b);if(d==101||d==69){d=a.charCodeAt(++b);(d==43||d==45)&&b++;for(f=b;f<c&&(d=a.charCodeAt(f),d>=48&&d<=57);f++);f==b&&i();b=f}return+a.slice(g,b)}j&&i();if(a.slice(b,b+4)=="true"){b=
    b+4;return true}if(a.slice(b,b+5)=="false"){b=b+5;return false}if(a.slice(b,b+4)=="null"){b=b+4;return n}i()}}return"$"},C=function(a){var c,b;a=="$"&&i();if(typeof a=="string"){if((t?a.charAt(0):a[0])=="@")return a.slice(1);if(a=="["){for(c=[];;b||(b=true)){a=r();if(a=="]")break;if(b)if(a==","){a=r();a=="]"&&i()}else i();a==","&&i();c.push(C(a))}return c}if(a=="{"){for(c={};;b||(b=true)){a=r();if(a=="}")break;if(b)if(a==","){a=r();a=="}"&&i()}else i();(a==","||typeof a!="string"||(t?a.charAt(0):
    a[0])!="@"||r()!=":")&&i();c[a.slice(1)]=C(r())}return c}i()}return a},F=function(a,b,e){e=E(a,b,e);e===s?delete a[b]:a[b]=e},E=function(a,b,e){var g=a[b],f;if(typeof g=="object"&&g)if(p.call(g)=="[object Array]")for(f=g.length;f--;)F(g,f,e);else x(g,function(a){F(g,a,e)});return e.call(a,b,g)};o.parse=function(a,c){var e,g;b=0;A=""+a;e=C(r());r()!="$"&&i();b=A=n;return c&&p.call(c)=="[object Function]"?E((g={},g[""]=e,g),"",c):e}}}H&&define(function(){return o})})(this);
    }());
}

var JSZip = null
if (typeof require === 'function') {
  JSZip = require('node-zip');
}

//----------------------------------------------------------
// Copyright (C) Microsoft Corporation. All rights reserved.
// Released under the Microsoft Office Extensible File License
// https://raw.github.com/stephen-hardy/xlsx.js/master/LICENSE.txt
//----------------------------------------------------------
function xlsx(file) { 
  'use strict'; // v2.3.0

  var defaultFontName = 'Calibri';
  var defaultFontSize = 11;

  var result, zip = new JSZip(), zipTime, processTime, s, f, i, j, k, l, t, w, sharedStrings, styles, index, data, val, style, borders, border, borderIndex, fonts, font, fontIndex,
    docProps, xl, xlWorksheets, worksheet, contentTypes = [[], []], props = [], xlRels = [], worksheets = [], id, columns, cols, colWidth, cell, row, merges, merged,
    numFmts = ['General', '0', '0.00', '#,##0', '#,##0.00',,,,, '0%', '0.00%', '0.00E+00', '# ?/?', '# ??/??', 'mm-dd-yy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM',
      'h:mm', 'h:mm:ss', 'm/d/yy h:mm',,,,,,,,,,,,,,, '#,##0 ;(#,##0)', '#,##0 ;[Red](#,##0)', '#,##0.00;(#,##0.00)', '#,##0.00;[Red](#,##0.00)',,,,, 'mm:ss', '[h]:mm:ss', 'mmss.0', '##0.0E+0', '@'],
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

  function numAlpha(i) { 
    var t = Math.floor(i / 26) - 1; return (t > -1 ? numAlpha(t) : '') + alphabet.charAt(i % 26); 
  }

  function alphaNum(s) { 
    var t = 0; if (s.length === 2) { t = alphaNum(s.charAt(0)) + 1; } return t * 26 + alphabet.indexOf(s.substr(-1)); 
  }

  function convertDate(input) { 
    return typeof input === 'object' ? ((input - new Date(1900, 0, 0)) / 86400000) + 1 : new Date(+new Date(1900, 0, 0) + (input - 1) * 86400000); 
  }

  function typeOf(obj) {
    return ({}).toString.call(obj).match(/\s([a-zA-Z]+)/)[1].toLowerCase();
  }
  
  function getAttr(s, n) { 
    s = s.substr(s.indexOf(n + '="') + n.length + 2); return s.substring(0, s.indexOf('"')); 
  }
  
  function escapeXML(s) { 
    return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#x27;'); 
  } // see http://www.w3.org/TR/xml/#syntax
  
  function unescapeXML(s) { 
    return (s || '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#x27;/g, '\''); 
  }

   if (typeof file === 'string') { 
    // Load
    zipTime = Date.now();
    zip = zip.load(file, { base64: true });
    result = { worksheets: [], zipTime: Date.now() - zipTime };
    processTime = Date.now();

    // Process sharedStrings
    sharedStrings = []; s = zip.file('xl/sharedStrings.xml');
    if (s) {
      s = s.asText().split(/<t.*?>/g); i = s.length;
      while(--i) { sharedStrings[i - 1] = unescapeXML(s[i].substring(0, s[i].indexOf('</t>'))); } // Do not process i === 0, because s[0] is the text before first t element
    }

    // Get file info from "docProps/core.xml"
    s = zip.file('docProps/core.xml').asText();
    s = s.substr(s.indexOf('<dc:creator>') + 12);
    result.creator = s.substring(0, s.indexOf('</dc:creator>'));
    s = s.substr(s.indexOf('<cp:lastModifiedBy>') + 19);
    result.lastModifiedBy = s.substring(0, s.indexOf('</cp:lastModifiedBy>'));
    s = s.substr(s.indexOf('<dcterms:created xsi:type="dcterms:W3CDTF">') + 43);
    result.created = new Date(s.substring(0, s.indexOf('</dcterms:created>')));
    s = s.substr(s.indexOf('<dcterms:modified xsi:type="dcterms:W3CDTF">') + 44);
    result.modified = new Date(s.substring(0, s.indexOf('</dcterms:modified>')));

    // Get workbook info from "xl/workbook.xml" - Worksheet names exist in other places, but "activeTab" attribute must be gathered from this file anyway
    s = zip.file('xl/workbook.xml').asText(); index = s.indexOf('activeTab="');
    if (index > 0) {
      s = s.substr(index + 11); 
      // Must eliminate first 11 characters before finding the index of " on the next line. Otherwise, it finds the " before the value.
      result.activeWorksheet = +s.substring(0, s.indexOf('"'));
    } else { 
      result.activeWorksheet = 0; 
    }
    s = s.split('<sheet '); i = s.length;
    while (--i) { // Do not process i === 0, because s[0] is the text before the first sheet element
      id = s[i].substr(s[i].indexOf('name="') + 6);
      result.worksheets.unshift({ name: id.substring(0, id.indexOf('"')), data: [] });
    }

    // Get style info from "xl/styles.xml"
    styles = [];
    s = zip.file('xl/styles.xml').asText().split('<numFmt '); i = s.length;
    while (--i) { 
      t = s[i]; numFmts[+getAttr(t, 'numFmtId')] = getAttr(t, 'formatCode'); 
    }
    s = s[s.length - 1]; s = s.substr(s.indexOf('cellXfs')).split('<xf '); i = s.length;
    while (--i) {
      id = getAttr(s[i], 'numFmtId'); f = numFmts[id];
      if (f.indexOf('m') > -1) {
        t = 'date'; 
      } else if (f.indexOf('0') > -1) { 
        t = 'number'; 
      } else if (f === '@') { 
        t = 'string'; 
      } else { 
        t = 'unknown'; 
      }
      styles.unshift({ formatCode: f, type: t });
    }

    // Get worksheet info from "xl/worksheets/sheetX.xml"
    i = result.worksheets.length;
    while (i--) {
      s = zip.file('xl/worksheets/sheet' + (i + 1) + '.xml' ).asText().split('<row ');
      w = result.worksheets[i];
      w.table = s[0].indexOf('<tableParts ') > 0;
      t = getAttr(s[0].substr(s[0].indexOf('<dimension')), 'ref');
      t = t.substr(t.indexOf(':') + 1);
      w.maxCol = alphaNum(t.match(/[a-zA-Z]*/g)[0]) + 1;
      w.maxRow = +t.match(/\d*/g).join('');
      w = w.data;
      j = s.length;
      while (--j) { // Don't process j === 0, because s[0] is the text before the first row element
        row = w[+getAttr(s[j], 'r') - 1] = [];
        columns = s[j].split('<c ');
        k = columns.length;
        while (--k) { // Don't process l === 0, because k[0] is the text before the first c (cell) element
          cell = columns[k];
          f = styles[+getAttr(cell, 's')] || { type: 'General', formatCode: 'General' };
          t = getAttr(cell, 't') || f.type;
          val = cell.substring(cell.indexOf('<v>') + 3, cell.indexOf('</v>'));
          val = val ? +val : ''; // turn non-zero into number
          switch (t) {
            case 's': val = sharedStrings[val]; break;
            case 'b': val = val === 1; break;
            case 'date': val = convertDate(val); break;
          }
          row[alphaNum(getAttr(cell, 'r').match(/[a-zA-Z]*/g)[0])] = { value: val, formatCode: f.formatCode };
        }
      }
    }

    result.processTime = Date.now() - processTime;

  } else {
    // Save
    processTime = Date.now();
    sharedStrings = [[], 0];
    // Fully static
    zip.folder('_rels').file('.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    docProps = zip.folder('docProps');

    xl = zip.folder('xl');
    xl.folder('theme').file('theme1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>');
    xlWorksheets = xl.folder('worksheets');

    // Not content dependent
    docProps.file('core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>'
      + (file.creator || 'XLSX.js') + '</dc:creator><cp:lastModifiedBy>' + (file.lastModifiedBy || 'XLSX.js') + '</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">'
      + (file.created || new Date()).toISOString() + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + (file.modified || new Date()).toISOString() + '</dcterms:modified></cp:coreProperties>');

    // Content dependent
    styles = new Array(1);
    borders = new Array(1);
    fonts = new Array(1);
    w = file.worksheets.length;
    while (w--) { 
      // Generate worksheet (gather sharedStrings), and possibly table files, then generate entries for constant files below
      id = w + 1;
      // Generate sheetX.xml in var s
      worksheet = file.worksheets[w]; data = worksheet.data;
      s = '';
      columns = [];
      merges = [];
      i = -1; l = data.length;
      while (++i < l) {
        j = -1; k = data[i].length;
        s += '<row r="' + (i + 1) + '" x14ac:dyDescent="0.25">';
        while (++j < k) {
          cell = data[i][j]; val = cell.hasOwnProperty('value') ? cell.value : cell; t = ''; 
          // supported styles: borders, hAlign, formatCode and font style
          style = {
            borders: cell.borders, 
            hAlign: cell.hAlign,
            vAlign: cell.vAlign,
            bold: cell.bold,
            italic: cell.italic,
            fontName: cell.fontName,
            fontSize: cell.fontSize,
            formatCode: cell.formatCode || 'General'
          };
          colWidth = 0;
          if (val && typeof val === 'string' && !isFinite(val)) { 
            // If value is string, and not string of just a number, place a sharedString reference instead of the value
            val = escapeXML(val);
            sharedStrings[1]++; // Increment total count, unique count derived from sharedStrings[0].length
            index = sharedStrings[0].indexOf(val);
            colWidth = val.length;
            if (index < 0) {
              index = sharedStrings[0].push(val) - 1; 
            }
            val = index;
            t = 's';
          } else if (typeof val === 'boolean') { 
            val = (val ? 1 : 0); t = 'b'; 
            colWidth = 1;
          } else if (typeOf(val) === 'date') { 
            val = convertDate(val); 
            style.formatCode = cell.formatCode || 'mm-dd-yy'; 
            colWidth = val.length;
          } else if (typeof val === 'object') {
            // unsupported value
            val = null
          } else {
            // number, or string which is a number 
            colWidth = (''+val).length;
          }
          
          // use stringified version as unic and reproductible style signature
          style = JSON.stringify(style);
          index = styles.indexOf(style);
          if (index < 0) { 
            style = styles.push(style) - 1; 
          } else { 
            style = index; 
          }
          // keeps largest cell in column, and autoWidth flag that may be set on any cell
          if (columns[j] == null) {
            columns[j] = {autoWidth: false, max:0};
          }
          if (cell.autoWidth) {
            columns[j].autoWidth = true
          }
          if (colWidth > columns[j].max) {
            columns[j].max = colWidth;
          }
          // store merges if needed and add missing cells. Cannot have rowSpan AND colSpan
          if (cell.colSpan > 1) {
            // horizontal merge. ex: B12:E12. Add missing cells (with same attribute but value) to current row
            merges.push([numAlpha(j) + (i + 1), numAlpha(j+cell.colSpan-1) + (i + 1)]);
            merged = [j, 0]
            for (var m = 0; m < cell.colSpan-1; m++) {
              merged.push(cell);
            }
            data[i].splice.apply(data[i], merged);
            k += cell.colSpan-1;
          } else if (cell.rowSpan > 1) {
            // vertical merge. ex: B12:B15. Add missing cells (with same attribute but value) to next columns
            for (var m = 1; m < cell.rowSpan; m++) {
              if (data[i+m]) {
                data[i+m].splice(j, 0, cell)
              } else {
                // readh the end of data
                cell.rowSpan = m;
                break;
              }
            }
            merges.push([numAlpha(j) + (i + 1), numAlpha(j) + (i + cell.rowSpan)]);
          }
          if (cell.rowSpan > 1 ||cell.colSpan > 1) {
            // deletes value, rowSpan and colSpan from cell to avoid refering it from copied cells
            delete cell.value;
            delete cell.rowSpan;
            delete cell.colSpan;
          }
          s += '<c r="' + numAlpha(j) + (i + 1) + '"' + (style ? ' s="' + style + '"' : '') + (t ? ' t="' + t + '"' : '');
          if (val != null) {
            s += '>' + (cell.formula ? '<f>' + cell.formula + '</f>' : '') + '<v>' + val + '</v></c>';
          } else {
            s += '/>';
          }
        }
        s += '</row>';
      }

      cols = []
      for (i = 0; i < columns.length; i++) {
        if (columns[i].autoWidth) {
          cols.push('<col min="', i+1, '" max="', i+1, '" width="', columns[i].max, '" bestFit="1"/>');
        }
      }
      // only add cols definition if not empty
      if (cols.length > 0) {
        cols = ['<cols>'].concat(cols, ['</cols>']).join('');
      }

      s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
        + '<dimension ref="A1:' + numAlpha(data[0].length - 1) + data.length + '"/><sheetViews><sheetView ' + (w === file.activeWorksheet ? 'tabSelected="1" ' : '')
        + ' workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
        + cols
        + '<sheetData>'
        + s 
        + '</sheetData>';
      if (merges.length > 0) {
        s += '<mergeCells count="' + merges.length + '">';
        for (i = 0; i < merges.length; i++) {
          s += '<mergeCell ref="' + merges[i].join(':') + '"/>';
        }
        s += '</mergeCells>';
      }
      s += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
      if (worksheet.table) { 
        s += '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'; 
      }
      xlWorksheets.file('sheet' + id + '.xml', s + '</worksheet>');

      if (worksheet.table) {
        i = -1; l = data[0].length; t = numAlpha(data[0].length - 1) + data.length;
        s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="' + id
          + '" name="Table' + id + '" displayName="Table' + id + '" ref="A1:' + t + '" totalsRowShown="0"><autoFilter ref="A1:' + t + '"/><tableColumns count="' + data[0].length + '">';
        while (++i < l) { 
          s += '<tableColumn id="' + (i + 1) + '" name="' + (data[0][i].hasOwnProperty('value') ? data[0][i].value : data[0][i]) + '"/>'; 
        }
        s += '</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';

        xl.folder('tables').file('table' + id + '.xml', s);
        xlWorksheets.folder('_rels').file('sheet' + id + '.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table' + id + '.xml"/></Relationships>');
        contentTypes[1].unshift('<Override PartName="/xl/tables/table' + id + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>');
      }

      contentTypes[0].unshift('<Override PartName="/xl/worksheets/sheet' + id + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
      props.unshift(escapeXML(worksheet.name) || 'Sheet' + id);
      xlRels.unshift('<Relationship Id="rId' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + id + '.xml"/>');
      worksheets.unshift('<sheet name="' + (escapeXML(worksheet.name) || 'Sheet' + id) + '" sheetId="' + id + '" r:id="rId' + id + '"/>');
    }

    // xl/styles.xml
    i = styles.length; t = [];
    while (--i) { 
      // Don't process index 0, already added
      style = JSON.parse(styles[i]);

      // cell formating, refer to it if necessary
      if (style.formatCode !== 'General') {
        index = numFmts.indexOf(style.formatCode);
        if (index < 0) { 
          index = 164 + t.length; 
          t.push('<numFmt formatCode="' + style.formatCode + '" numFmtId="' + index + '"/>'); 
        }
        style.formatCode = index
      } else {
        style.formatCode = 0
      }

      // border declaration: add a new declaration and refer to it in style
      borderIndex = 0
      if (style.borders) {
        border = ['<border>']
        // order is significative
        for (var edge in {left:0, right:0, top:0, bottom:0, diagonal:0}) {
          if (style.borders[edge]) {
            var color = style.borders[edge];
            // add transparency if missing
            if (color.length === 6) {
              color = 'FF'+color;
            }
            border.push('<', edge, ' style="thin">', '<color rgb="', style.borders[edge], '"/></', edge, '>');
          } else {
            border.push('<', edge, '/>');
          }
        }
        border.push('</border>');
        border = border.join('');
        // try to reuse existing border
        borderIndex = borders.indexOf(border);
        if (borderIndex < 0) {
          borderIndex = borders.push(border) - 1;
        }
      }

      // font declaration: add a new declaration and refer to it in style
      fontIndex = 0
      if (style.bold || style.italic || style.fontSize || style.fontName) {
        font = ['<font>']
        if (style.bold) {
          font.push('<b/>');
        }
        if (style.italic) {
          font.push('<i/>');
        }
        font.push('<sz val="', style.fontSize || defaultFontSize, '"/>');
        font.push('<color theme="1"/>');
        font.push('<name val="', style.fontName || defaultFontName, '"/>');
        font.push('<family val="2"/>', '</font>');
        font = font.join('');
        // try to reuse existing font
        fontIndex = fonts.indexOf(font);
        if (fontIndex < 0) {
          fontIndex = fonts.push(font) - 1;
        }
      }

      // declares style, and refer to optionnal formatCode, font and borders
      styles[i] = ['<xf xfId="0" fillId="0" borderId="', 
        borderIndex, 
        '" fontId="',
        fontIndex,
        '" numFmtId="',
        style.formatCode,
        '" ',
        (style.hAlign || style.vAlign? 'applyAlignment="1" ' : ' '),
        (style.formatCode > 0 ? 'applyNumberFormat="1" ' : ' '),
        (borderIndex > 0 ? 'applyBorder="1" ' : ' '),
        (fontIndex > 0 ? 'applyFont="1" ' : ' '),
        '>'
      ];
      if (style.hAlign || style.vAlign) {
        styles[i].push('<alignment');
        if (style.hAlign) {
          styles[i].push(' horizontal="', style.hAlign, '"');
        }
        if (style.vAlign) {
          styles[i].push(' vertical="', style.vAlign, '"');
        }
        styles[i].push('/>');
      }
      styles[i].push('</xf>');
      styles[i] = styles[i].join('');
    }
    t = t.length ? '<numFmts count="' + t.length + '">' + t.join('') + '</numFmts>' : '';

    xl.file('styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
      + t + '<fonts count="'+ fonts.length + '" x14ac:knownFonts="1"><font><sz val="' + defaultFontSize + '"/><color theme="1"/><name val="' + defaultFontName + '"/><family val="2"/>'
      + '<scheme val="minor"/></font>' + fonts.join('') + '</fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
      + '<borders count="' + borders.length + '"><border><left/><right/><top/><bottom/><diagonal/></border>'
      + borders.join('') + '</borders><cellStyleXfs count="1">'
      + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="' + styles.length + '"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
      + styles.join('') + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>'
      + '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
      + '<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
      + '<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>');

    // [Content_Types].xml
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + contentTypes[0].join('') + '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
      + contentTypes[1].join('') + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');

    // docProps/app.xml
    docProps.file('app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>XLSX.js</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>'
      + file.worksheets.length + '</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="' + props.length + '" baseType="lpstr"><vt:lpstr>' + props.join('</vt:lpstr><vt:lpstr>')
      + '</vt:lpstr></vt:vector></TitlesOfParts><Manager></Manager><Company>Microsoft Corporation</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>1.0</AppVersion></Properties>');

    // xl/_rels/workbook.xml.rels
    xl.folder('_rels').file('workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + xlRels.join('') + '<Relationship Id="rId' + (xlRels.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
      + '<Relationship Id="rId' + (xlRels.length + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      + '<Relationship Id="rId' + (xlRels.length + 3) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>');

    // xl/sharedStrings.xml
    xl.file('sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'
      + sharedStrings[1] + '" uniqueCount="' + sharedStrings[0].length + '"><si><t>' + sharedStrings[0].join('</t></si><si><t>') + '</t></si></sst>');

    // xl/workbook.xml
    xl.file('workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + '<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView '
      + (file.activeWorksheet ? 'activeTab="' + file.activeWorksheet + '" ' : '') + 'xWindow="480" yWindow="60" windowWidth="18195" windowHeight="8505"/></bookViews><sheets>'
      + worksheets.join('') + '</sheets><calcPr calcId="145621"/></workbook>');

    processTime = Date.now() - processTime;
    zipTime = Date.now();
    result = {
      base64: zip.generate({ compression: 'DEFLATE' }), zipTime: Date.now() - zipTime, processTime: processTime,
      href: function() { return 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + this.base64; }
    };
  }
  return result;
}

// NodeJs export
if (typeof exports === 'object' && typeof module === 'object') {
  module.exports = xlsx;
}