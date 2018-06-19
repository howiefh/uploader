/*!
 * 
 * Copyright (c) 2018 fenghao <howiefh@gmail.com>
 * 
 * Dependencies  : https://github.com/SheetJS/js-xlsx, https://github.com/eligrey/FileSaver.js,
 * https://github.com/eligrey/Blob.js, https://github.com/axios/axios
 * 
 * Licensed under the MIT license.
 */
(function (global, f) {
  if (typeof exports === "object" && typeof module !== "undefined") {
    module.exports = f();
  } else if (typeof define === "function" && define.amd) {
    define([], f);
  } else {
    global.ExcelExporter = f();
  }
})(this, function () {
  var defaults = {
    tableId: null,
    url: null,
    data: null,
    ajaxData: null,
    filename: 'excel',
    autoWidth: true,
    header: null,
    bindClick: true,
    nullToEmpty: true,
    headerMap: null,
    formatterMap: null,
    fields: null
  };

  function ExcelExporter(element, options) {
    var self = this;
    self.opts = Object.assign({}, defaults, options);
    if (isArray(self.opts.headerMap)) {
      var newMap = {};
      for (var i = 0; i < self.opts.headerMap.length; i++) {
        var item = self.opts.headerMap[i];
        newMap[item.name] = item.label
      }
      self.opts.headerMap = newMap
    }

    if (!self.opts.fields && self.opts.headerMap) {
      self.opts.fields = Object.keys(self.opts.headerMap)
    }
    if (!self.opts.header && self.opts.headerMap) {
      self.opts.header = Object.values(self.opts.headerMap)
    }
    if (!self.opts.header && self.opts.fields) {
      self.opts.header = self.opts.fields
    }

    self.el = getByID(element);
    if (self.opts.tableId) {
      self.table = getByID(self.opts.tableId);
    }
    self.downloadStatus = false;
    self.opts.bindClick && bind(self.el, 'click', function () {
      if (self.opts.url) {
        var data;
        if (typeof self.opts.ajaxData === 'function') {
          data = self.opts.ajaxData()
        } else if (typeof self.opts.ajaxData === 'object') {
          data = self.opts.ajaxData
        }
        if (data === false) {
          return false
        }

        if (self.downloadStatus) {
          alert('Downloading...');
          return false;
        }
        self.downloadStatus = true;
        self.el['disabled'] = true;
        axios({
          url: self.opts.url,
          method: 'get',
          data: data
        }).then(function (resp) {
          var data = resp.data;
          if (data.code === 200) {
            self.exportJsonToExcel(data.rows, self.opts.header)
          } else {
            alert('System Error' + data.msg);
          }
          self.downloadStatus = false;
          self.el['disabled'] = false;
        }).catch(function (err) {
          alert(err);
          self.downloadStatus = false;
          self.el['disabled'] = false;
        });
      } else if (self.opts.tableId) {
        self.exportTableToExcel(self.opts.tableId);
      } else if (self.opts.data) {
        if (typeof self.opts.data === 'function') {
          data = self.opts.data()
        } else if (typeof self.opts.ajaxData === 'object') {
          data = self.opts.data
        }
        self.exportJsonToExcel(data, self.opts.header)
      }
    })
  };

  ExcelExporter.prototype.exportTableToExcel = function (id) {
    var theTable;
    if (!this.table && id) {
      this.table = getByID(id);
    }

    theTable = this.table;
    var filename = this.opts.filename || Date.now();

    var oo = generateArray(theTable);
    var ranges = oo[1];

    /* original data */
    var data = oo[0];
    var ws_name = "SheetJS";

    var wb = new Workbook(), ws = sheetFromArrayOfArrays(data);

    /* add ranges to worksheet */
    // ws['!cols'] = ['apple', 'banan'];
    ws['!merges'] = ranges;

    /* add worksheet to workbook */
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;

    var wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: false, type: 'binary' });

    saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), filename + ".xlsx");
  };

  ExcelExporter.prototype.exportJsonToExcel = function (data, header, filename, autoWidth) {
    if (!data) {
      alert('Data empty');
      return false;
    }
    header = header || this.opts.header;
    if (!header && data.length > 0) {
      header = Object.keys(data[0]);
    }
    var fields = this.opts.fields || header || []
    filename = filename || this.opts.filename || Date.now();
    if (typeof autoWidth !== 'boolean') {
      autoWidth = this.opts.autoWidth;
    }

    data = filterData(fields, data, this.opts);

    data.unshift(header);
    var ws_name = "SheetJS";
    var wb = new Workbook(), ws = sheetFromArrayOfArrays(data);

    if (autoWidth) {
      /*设置worksheet每列的最大宽度*/
      const colWidth = data.map(function (row) {
        return row.map(function (val) {
          /*先判断是否为null/undefined*/
          if (val == null) {
            return { 'wch': 10 };
          }
          /*再判断是否为中文*/
          else if (val.toString().charCodeAt(0) > 255) {
            return { 'wch': val.toString().length * 2 };
          } else {
            return { 'wch': val.toString().length };
          }
        })
      });
      /*以第一行为初始值*/
      var result = colWidth[0];
      for (var i = 1; i < colWidth.length; i++) {
        for (var j = 0; j < colWidth[i].length; j++) {
          if (result[j]['wch'] < colWidth[i][j]['wch']) {
            result[j]['wch'] = colWidth[i][j]['wch'];
          }
        }
      }
      ws['!cols'] = result;
    }

    /* add worksheet to workbook */
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;

    var wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: false, type: 'binary' });
    saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), filename + ".xlsx");
  }

  function getByID(id) {
    return document.getElementById(id);
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

  function generateArray(table) {
    var out = [];
    var rows = table.querySelectorAll('tr');
    var ranges = [];
    for (var R = 0; R < rows.length; ++R) {
      var outRow = [];
      var row = rows[R];
      var columns = row.querySelectorAll('th,td');
      for (var C = 0; C < columns.length; ++C) {
        var cell = columns[C];
        var colspan = cell.getAttribute('colspan');
        var rowspan = cell.getAttribute('rowspan');
        var cellValue = cell.innerText;
        if (cellValue !== "" && cellValue == +cellValue) cellValue = +cellValue;

        //Skip ranges
        ranges.forEach(function (range) {
          if (R >= range.s.r && R <= range.e.r && outRow.length >= range.s.c && outRow.length <= range.e.c) {
            for (var i = 0; i <= range.e.c - range.s.c; ++i) outRow.push(null);
          }
        });

        //Handle Row Span
        if (rowspan || colspan) {
          rowspan = rowspan || 1;
          colspan = colspan || 1;
          ranges.push({ s: { r: R, c: outRow.length }, e: { r: R + rowspan - 1, c: outRow.length + colspan - 1 } });
        }
        ;

        //Handle Value
        outRow.push(cellValue !== "" ? cellValue : null);

        //Handle Colspan
        if (colspan) for (var k = 0; k < colspan - 1; ++k) outRow.push(null);
      }
      out.push(outRow);
    }
    return [out, ranges];
  };

  function datenum(v, date1904) {
    if (date1904) v += 1462;
    var epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
  }

  function sheetFromArrayOfArrays(data, opts) {
    var ws = {};
    var range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
    for (var R = 0; R != data.length; ++R) {
      for (var C = 0; C != data[R].length; ++C) {
        if (range.s.r > R) range.s.r = R;
        if (range.s.c > C) range.s.c = C;
        if (range.e.r < R) range.e.r = R;
        if (range.e.c < C) range.e.c = C;
        var cell = { v: data[R][C] };
        if (cell.v == null) continue;
        var cell_ref = XLSX.utils.encode_cell({ c: C, r: R });

        if (typeof cell.v === 'number') cell.t = 'n';
        else if (typeof cell.v === 'boolean') cell.t = 'b';
        else if (cell.v instanceof Date) {
          cell.t = 'n';
          cell.z = XLSX.SSF._table[14];
          cell.v = datenum(cell.v);
        }
        else cell.t = 's';

        ws[cell_ref] = cell;
      }
    }
    if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
    return ws;
  }

  function Workbook() {
    if (!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
  }

  function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }

  function isArray(item) {
    return (item && Array.isArray(item))
  }

  function filterData(filterVal, jsonData, opts) {
    return jsonData.map(function (v) {
      return filterVal.map(function (j) {
        var val = v[j];
        if (opts.formatterMap && opts.formatterMap[j]) {
          var tmp = opts.formatterMap[j][val];
          val = !tmp ? val : tmp;
        }
        if ((val === null || val === undefined) && opts.nullToEmpty) {
          return ''
        } else if (typeof val === 'object') {
          return JSON.stringify(val)
        } else {
          return val
        }
      })
    })
  }

  return ExcelExporter;
});