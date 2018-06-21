/*!
 * 
 * Copyright (c) 2018 fenghao <howiefh@gmail.com>
 * 
 * Dependencies  : https://github.com/SheetJS/js-xlsx
 * 
 * Licensed under the MIT license.
 */
(function (global, f) {
  if (typeof exports === "object" && typeof module !== "undefined") {
    module.exports = f();
  } else if (typeof define === "function" && define.amd) {
    define([], f);
  } else {
    global.ExcelImporter = f();
  }
})(this, function () {
  var defaults = {
    onLoaded: null,
    headerMap: null,
    headerRow: null,
    headerIndex: null,
    bindClick: true,
    onlyFirstSheet: true,
    includeUnknowHeader: false,
    includeEmptyHeader: false,
    dateNF: 'yyyy-MM-dd'
  };

  function ExcelImporter(element, options) {
    var self = this;
    self.opts = Object.assign({}, defaults, options);
    self.setHeaderMap(self.opts.headerMap);
    self.setHeaderRow(self.opts.headerRow);

    self.el = getByID(element);
    if (self.opts.bindClick) {
      self.input = create('input', {
        type: 'file',
        accept: '.xlsx, .xls',
        style: 'display: none; z-index: -9999;'
      });
      self.el.parentElement.appendChild(self.input);
      bind(self.input, 'change', function(e) {
        var files = e.target.files;
        var itemFile = files[0]; // only use files[0]
        if (!itemFile) return;
        readerData(itemFile, self.opts);
        self.input.value = null; // fix can't select the same excel
      });
      bind(self.el, 'click', function(e) {
        self.input.click()
      })
    }
  }
  ExcelImporter.prototype = {
    setHeaderMap: function(headerMap) {
      var self = this;
      if (isArray(headerMap)) {
        var newMap = {};
        for (var i = 0; i < headerMap.length; i++) {
          var item = headerMap[i];
          newMap[item.label] = item.name;
        }
        self.opts.headerMap = newMap;
      } else if (typeof headerMap === 'object') {
        self.opts.headerMap = headerMap;
      }
    },
    setHeaderRow: function(row) {
      this.opts.headerRow = row;
      this.opts.headerIndex = headerIndex(this.opts.headerRow);
    },
    loadExcel: function (file, opts) {
        if (!file) return;
        readerData(file, opts);
    }
  }

  function isArray(item) {
    return (item && Array.isArray(item))
  }

  function getByID(id) {
    return document.getElementById(id);
  }

  function create(element, attrs) {
    var el = document.createElement(element);
    for (var key in attrs) {
      el[key] = attrs[key];
    }
    return el;
  }

  function bind(element, event, callback, options) {
    if (element.addEventListener) {
      element.addEventListener(event, callback, options);
    } else {
      // IE8 fallback
      element.attachEvent('on' + event, function (event) {
        // `event` and `event.target` are not provided in IE8
        event = event || window.event;
        event.target = event.target || event.srcElement;
        callback(event);
      });
    }
  }

  function readerData(itemFile, opts) {
    var reader = new FileReader();
    reader.onload = function (e) {
      var data = e.target.result;
      var fixedData = fixData(data);
      var workbook = XLSX.read(btoa(fixedData), {
        type: 'base64',
        cellText: false,
        cellDates: true
      });
      var excelData = [];
      for (var i = 0; i < workbook.SheetNames.length; i++) {
        var sheetName = workbook.SheetNames[i];
        var worksheet = workbook.Sheets[sheetName];
        if (!worksheet['!ref']) {
          continue;
        }
        var range = XLSX.utils.decode_range(worksheet['!ref']);
        if (opts.headerIndex && range.s.r < opts.headerIndex && range.e.r > opts.headerIndex) {
            range.s.r = opts.headerIndex;
        }
        var header = getHeaderRow(worksheet, range, opts);
        if (header && header.length && range.e.c !== header.length - 1) {
            range.e.c = header.length - 1;
        }
        var results = XLSX.utils.sheet_to_json(worksheet, {
          range: range,
          dateNF: opts.dateNF
        });
        var item = generateData(sheetName, header, results, opts);
        excelData.push(item);
        if (opts.onlyFirstSheet) {
          break;
        }
      }
      if (typeof opts.onLoaded === 'function') {
        opts.onLoaded(excelData)
      }
    };
    reader.readAsArrayBuffer(itemFile)
  }
  function fixData(data) {
    var o = '';
    var l = 0;
    var w = 10240;
    for (; l < data.byteLength / w; ++l) {
      o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)))
    }

    o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)))
    return o
  }
  
  function getHeaderRow(sheet, range, opts) {
    var headers = [];
    var C;
    var R = range.s.r; /* start in the first row */
    for (C = range.s.c; C <= range.e.c; ++C) { /* walk every column in the range */
      var cell = sheet[XLSX.utils.encode_cell({
        c: C,
        r: R
      })]; /* find the cell in the first row */
      var hdr = 'UNKNOWN ' + C; // <-- replace with your desired default
      if (cell && cell.t) {
        hdr = XLSX.utils.format_cell(cell);
        headers.push(hdr);
      } else if (opts.includeEmptyHeader) {
          headers.push(hdr);
      }
    }
    return headers
  }

  function headerIndex(headerRow) {
    var isValidHeaderRow = typeof headerRow === 'number' && headerRow > 0;
    return isValidHeaderRow ? headerRow - 1 : undefined;
  }

  function generateData(sheetName, header, results, opts) {
    if (opts.headerMap) {
      header = header.map(function (v) {
        return newVal ? newVal : v
      });

      var newHeader = [];
      for (var i = 0; i < header.length; i++) {
        var newVal = opts.headerMap[header[i]];
        if (newVal) {
          newHeader.push(newVal);
        } else if (opts.includeUnknowHeader) {
          newHeader.push(header[i]);
        }
      }
      header = newHeader;

      results = results.map(function (v) {
        var newObj = {};
        for (var key in v) {
          var newKey = opts.headerMap[key];
          if (newKey) {
            newObj[newKey] = v[key];
          } else if (opts.includeUnknowHeader) {
            newObj[key] = v[key];
          }
        }
        return newObj
      });
    }
    return {
      sheet: sheetName,
      header: header,
      results: results
    }
  }
  return ExcelImporter;
});