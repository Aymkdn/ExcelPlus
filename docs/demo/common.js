/**
 * The MIT License (MIT) - https://github.com/yanatan16/nanoajax/blob/master/LICENSE
 * Source : https://github.com/yanatan16/nanoajax
 */
!function(t,e){function n(t){return t&&e.XDomainRequest&&!/MSIE 1/.test(navigator.userAgent)?new XDomainRequest:e.XMLHttpRequest?new XMLHttpRequest:void 0}function o(t,e,n){t[e]=t[e]||n}var r=["responseType","withCredentials","timeout","onprogress"];t.ajax=function(t,a){function s(t,e){return function(){c||(a(void 0===f.status?t:f.status,0===f.status?"Error":f.response||f.responseText||e,f),c=!0)}}var u=t.headers||{},i=t.body,d=t.method||(i?"POST":"GET"),c=!1,f=n(t.cors);f.open(d,t.url,!0);var l=f.onload=s(200);f.onreadystatechange=function(){4===f.readyState&&l()},f.onerror=s(null,"Error"),f.ontimeout=s(null,"Timeout"),f.onabort=s(null,"Abort"),i&&(o(u,"X-Requested-With","XMLHttpRequest"),e.FormData&&i instanceof e.FormData||o(u,"Content-Type","application/x-www-form-urlencoded"));for(var p,m=0,v=r.length;v>m;m++)p=r[m],void 0!==t[p]&&(f[p]=t[p]);for(var p in u)f.setRequestHeader(p,u[p]);return f.send(i),f},e.nanoajax=t}({},function(){return this}());

// add http://stackoverflow.com/a/31374433/1134119
// example: loadJS('yourcode.js', document.body, yourCodeToBeCalled);
var loadJS=function(a,n,e){var t=document.createElement("script");t.src=a,t.onload=e||function(){},t.onreadystatechange=e,n.appendChild(t)};

function demo_addEvent(t,n,e){t.attachEvent?t.attachEvent("on"+n,e):t.addEventListener(n,e,!0)}

// get version of a lib
function getVersion(repo, callback) {
  nanoajax.ajax({url:'https://api.github.com/repos/'+repo+'/tags'}, function (code, responseText) {
    var version = "0";
    try {
      var res = JSON.parse(responseText);
      if (Object.prototype.toString.call(res) === '[object Array]' && res.length > 0) {
        version = res[0].name.replace(/^v/,"");
      }
    } catch(e) {}
    callback(version)
  })
}
var lib_versions={"js-xlsx":0, "ExcelPlus":0};
// Call our libraries
getVersion('sheetjs/js-xlsx', function(version) {
  lib_versions["js-xlsx"]=version;
  // load the latest release
  loadJS('https://cdnjs.cloudflare.com/ajax/libs/xlsx/'+version+'/xlsx.core.min.js', document.head, function() {
    // get ExcelPlus
    getVersion('aymkdn/excelplus', function(version) {
      lib_versions.ExcelPlus=version;
      if (document.location.host) {
        loadJS('http://github-proxy.kodono.info/?q=https://raw.githubusercontent.com/Aymkdn/ExcelPlus/master/'+version+'/excelplus-'+version+'.js', document.head, onReady)
      } else {
        loadJS('../../2.5/excelplus-2.5.js', document.head, onReady)
      }

      // show demo code
      document.body.insertAdjacentHTML('beforeend', '<div><h1>Code for this demo (js-xlsx v'+lib_versions["js-xlsx"]+' / ExcelPlus v'+lib_versions.ExcelPlus+'):</h1><pre id="code" class="prettyprint lang-js"></pre></div>');
      document.querySelector('#code').innerHTML = document.querySelector('#demo').innerHTML.replace(/</g,"&lt;");
      loadJS('https://cdn.rawgit.com/google/code-prettify/master/loader/run_prettify.js', document.body);
    })
  })
})

