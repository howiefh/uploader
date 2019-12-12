/*!
 * 
 * Copyright (c) 2018 fenghao <howiefh@gmail.com>
 * 
 * Dependencies  : https://github.com/SheetJS/js-xlsx, https://github.com/eligrey/FileSaver.js,
 * https://github.com/eligrey/Blob.js
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
    /**
     * 表格元素 id，需要导出页面表格时配置
     */
    tableId: null,
    /**
     * 远程地址，需要导出远程数据时配置
     */
    url: null,
    /**
     * 导出数据，需要导出本地数据时配置
     */
    data: null,
    /**
     * 远程请求，响应中数据属性名
     */
    dataKey: 'data',
    /**
     * 导出远程数据时，ajax 请求所携带的数据，可以是方法或对象，方法适用于每次导出时参数可能会改变的情况
     * 例如
     * ajaxData: function() {
     *   return {t: Date.now()}
     * }
     * 或
     * ajaxData: {t: Date.now()}
     */
    ajaxData: null,
    /**
     * 导出的文件名，不需要带后缀
     */
    filename: 'excel',
    /**
     * 是否自动适应宽度
     */
    autoWidth: true,
    /**
     * 表头样式
     */
    headerStyle: null,
    /**
     * 表头行数
     */
    headerRows: 1,
    /**
     * 多级表头
     */
    columns: null,
    /**
     * 表头数组，不传的话从headerMap中解析，如果headerMap也没设置则使用fields配置
     */
    header: null,
    /**
     * Excel表格头名和返回数据字段名的映射，接受对象和数组两种数据, 如果使用jqGrid直接将colModel传入即可
     * 如：
     * {'姓名': 'name', '性别': 'sex'}
     * 或
     * [{label:'姓名', name:'name'}, {label:'性别', name: 'sex'}]
     */
    headerMap: null,
    /**
     * 字段名数组，如果没设置从 headerMap 中读取，如果设置，导出时只导出该数组中包含的字段
     */
    fields: null,
    /**
     * 是否为元素绑定点击事件，默认true
     */
    bindClick: true,
    /**
     * 是否将null转为空字符串 默认true
     */
    nullToEmpty: true,
    /**
     * 字段格式化映射，key 为字段名，value为方法或者对象
     * 例如
     * formatterMap: {
     *   name: function(val) {
     *     return 'name:' + val;
     *   },
     *   sex: {'f':'女','m':'男'}
     * }
     */
    formatterMap: null,
  };

  let tableLeafColumns = []
  /**
   * Excel 导出器
   * @param {*} element 元素 id
   * @param {*} options 配置参数
   */
  function ExcelExporter(element, options) {
    var self = this;
    self.opts = Object.assign({}, defaults, options);
    const hasColumnsArray = isArray(self.opts.columns);
    if (self.opts.headerRows === 1 && hasColumnsArray) {
      self.opts.headerRows = treeDepth(self.opts.columns)
    }

    if (self.opts.headerRows > 1 && hasColumnsArray) {
      tableLeafColumns = []
      parseColumns(self.opts.columns, 0, 0, self.opts.headerRows)
      self.opts.headerMap = tableLeafColumns
    } else if (!self.opts.headerMap && hasColumnsArray) {
      self.opts.headerMap = self.opts.columns
    }

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
      self.opts.header = self.opts.fields.map(function (t) {
        return self.opts.headerMap[t] || t;
      })
    }
    if (!self.opts.header && self.opts.fields) {
      self.opts.header = self.opts.fields
    }

    if (self.opts.headerStyle) {
      self.opts.headerStyle = parseStyle(self.opts.headerStyle)
    }

    self.el = getByID(element);
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
        const url = encodeQueryString(self.opts.url, data);
        fetch(url, {
            method: 'GET'
          }).then(res => res.json())
          .then((data) => {
            if (data.code === 200) {
              var rows;
              if (typeof self.opts.dataKey === 'function') {
                rows = self.opts.dataKey(data)
              } else if (typeof self.opts.dataKey === 'string') {
                rows = data[self.opts.dataKey]
              }
              if (self.opts.headerRows > 1) {
                self.exportJsonToExcelWithGroupHeader(rows)
              } else {
                self.exportJsonToExcel(rows, self.opts.header)
              }
            } else {
              alert('System Error: ' + data.msg);
            }
            self.downloadStatus = false;
            self.el['disabled'] = false;
          })
          .catch((err) => {
            alert(err);
            self.downloadStatus = false;
            self.el['disabled'] = false;
          });
      } else if (self.opts.tableId) {
        self.exportTableToExcel(self.opts.tableId);
      } else if (self.opts.data) {
        if (typeof self.opts.data === 'function') {
          data = self.opts.data()
        } else if (typeof self.opts.data === 'object') {
          data = self.opts.data
        }
        if (self.opts.headerRows > 1) {
          self.exportJsonToExcelWithGroupHeader(data)
        } else {
          self.exportJsonToExcel(data, self.opts.header)
        }
      }
    })
  };

  /**
   * 导出页面表格
   * @param {*} id 表格id
   */
  ExcelExporter.prototype.exportTableToExcel = function (id) {
    var theTable = getByID(id);

    var filename = this.opts.filename || Date.now();

    var dataAndRanges = generateArray(theTable);
    var ranges = dataAndRanges[1];

    /* original data */
    var data = dataAndRanges[0];
    var wsName = "SheetJS";

    var wb = new Workbook(),
      ws = sheetFromArrayOfArrays(data);

    /* add ranges to worksheet */
    // ws['!cols'] = ['apple', 'banan'];
    ws['!merges'] = ranges;

    /* add worksheet to workbook */
    wb.SheetNames.push(wsName);
    wb.Sheets[wsName] = ws;

    var wbout = XLSX.write(wb, {
      bookType: 'xlsx',
      bookSST: false,
      type: 'binary'
    });

    saveAs(new Blob([s2ab(wbout)], {
      type: "application/octet-stream"
    }), filename + ".xlsx");
  };

  /**
   * 导出json数据
   * @param {*} data json 数组
   * @param {*} header 表头 可以不传 默认使用配置参数
   * @param {*} fields 表头对应的字段名 可以不传 默认使用配置参数
   * @param {*} filename 文件名 可以不传 默认使用配置参数
   * @param {*} autoWidth 是否自动调整宽度 可以不传 默认使用配置参数
   */
  ExcelExporter.prototype.exportJsonToExcel = function (data, header, fields, filename, autoWidth) {
    if (!data || !data.length) {
      console.log('没有数据');
      return false;
    }
    fields = fields || this.opts.fields || Object.keys(data[0])
    header = header || this.opts.header || fields
    filename = filename || this.opts.filename || Date.now();
    if (typeof autoWidth !== 'boolean') {
      autoWidth = this.opts.autoWidth;
    }

    data = filterData(fields, data, this.opts);

    const wsName = Date.now() + '-' + data.length
    data=[...data]
    data.unshift(header);
    const ws = sheetFromArrayOfArrays(data, { headerRows: this.opts.headerRows, headerStyle: this.opts.headerStyle })

    internalExportJsonToExcel({ ws, wsName, data, filename, autoWidth })
  }

  /**
   * 导出多级表头的json数据
   * @param {*} data json 数组
   * @param {*} header 表头 可以不传 默认使用配置参数
   * @param {*} fields 表头对应的字段名 可以不传 默认使用配置参数
   * @param {*} filename 文件名 可以不传 默认使用配置参数
   * @param {*} autoWidth 是否自动调整宽度 可以不传 默认使用配置参数
   */
  ExcelExporter.prototype.exportJsonToExcelWithGroupHeader = function (data, header, fields, filename, autoWidth) {
    if (!data || !data.length) {
      console.log('没有数据');
      return false;
    }
    fields = fields || this.opts.fields || Object.keys(data[0])
    let headerRows
    if (isArray(header)) {
      const oldTableLeafColumns = tableLeafColumns
      headerRows = treeDepth(header)
      parseColumns(header, 0, 0, headerRows)
      tableLeafColumns = oldTableLeafColumns
    } else {
      header = this.opts.columns
      headerRows = this.opts.headerRows
    }
    filename = filename || this.opts.filename || Date.now();
    if (typeof autoWidth !== 'boolean') {
      autoWidth = this.opts.autoWidth;
    }

    const result = generateHeaderRows(header)
    const headers = result.rows
    const merges = result.merges

    data = filterData(fields, data, this.opts);
    const wsName = Date.now() + '-' + data.length
    data = [...data]
    data.unshift(...headers)
    const ws = sheetFromArrayOfArrays(data, { headerRows: headerRows, headerStyle: this.opts.headerStyle })
    ws['!merges'] = merges

    internalExportJsonToExcel({ ws, wsName, data, filename, autoWidth })
  }

  function treeDepth(columns) {
    const depth = [0]
    for (let i = 0; i < columns.length; i++) {
      depth.push(treeNodeDepth(columns[i]))
    }
    return Math.max(...depth)
  }

  function treeNodeDepth(column) {
    // 叶子节点
    if ((!column.children || !column.children.length)) {
      return 1
    }

    return treeDepth(column.children) + 1
  }

  /**
   * @param startRow 从 0 起始
   * @param startCol 从 0 起始
   */
  function parseColumns(columns, startRow, startCol, headerRows) {
    let colSpan = 0
    let prevColSpan = 0
    for (let i = 0; i < columns.length; i++) {
      const column = columns[i]
      // 没有设置 label 时 使用 name 属性
      column.label = column.label || column.name

      column.startRow = startRow
      startCol = startCol + prevColSpan
      column.startCol = startCol
      // 叶子节点
      if ((!column.children || !column.children.length)) {
        column.colSpan = 1
        column.rowSpan = headerRows - column.startRow
        column.isLeaf = true
        tableLeafColumns.push(column)
      } else {
        column.isLeaf = false
        column.colSpan = parseColumns(column.children, startRow + 1, startCol, headerRows)
        column.rowSpan = 1
      }
      colSpan += column.colSpan
      prevColSpan = column.colSpan
    }
    return colSpan
  }

  function internalExportJsonToExcel({ws, wsName, data, filename='excel-list', autoWidth=true}={}) {
    var wb = new Workbook()

    if(autoWidth){
      /*设置worksheet每列的最大宽度*/
      const colWidth = data.map(row => row.map(val => {
        /*先判断是否为null/undefined*/
        if (val == null) {
          return {'wch': 10};
        }
        /*再判断是否为中文*/
        else if (val.toString().charCodeAt(0) > 255) {
          return {'wch': val.toString().length * 2};
        } else {
          return {'wch': val.toString().length};
        }
      }))
      /*以第一行为初始值*/
      let result = colWidth[0];
      for (let i = 1; i < colWidth.length; i++) {
        for (let j = 0; j < colWidth[i].length; j++) {
          if (result[j]['wch'] < colWidth[i][j]['wch']) {
            result[j]['wch'] = colWidth[i][j]['wch'];
          }
        }
      }
      ws['!cols'] = result;
    }

    /* add worksheet to workbook */
    wb.SheetNames.push(wsName);
    wb.Sheets[wsName] = ws;

    var wbout = XLSX.write(wb, {bookType: 'xlsx', bookSST: false, type: 'binary'});
    saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}), filename + ".xlsx");
  }

  function parseStyle(style) {
    if (!style) {
      return style
    }
    const newStyle = {}
    Object.keys(style).forEach(key => {
      const value = style[key]
      switch (key) {
        case 'fgColor': {
          if (!newStyle.fill) {
            newStyle.fill = {}
          }
          newStyle.fill.fgColor = { rgb: value }
          break
        }
        case 'bgColor': {
          if (!newStyle.fill) {
            newStyle.fill = {}
          }
          newStyle.fill.bgColor = { rgb: value }
          break
        }
        case 'fontColor': {
          if (!newStyle.font) {
            newStyle.font = {}
          }
          newStyle.font.color = { rgb: value }
          break
        }
        case 'fontSize': {
          if (!newStyle.font) {
            newStyle.font = {}
          }
          newStyle.font.sz = value
          break
        }
        case 'fontBold': {
          if (!newStyle.font) {
            newStyle.font = {}
          }
          newStyle.font.bold = value
          break
        }
        case 'align': {
          if (!newStyle.alignment) {
            newStyle.alignment = {}
          }
          newStyle.alignment.vertical = value
          newStyle.alignment.horizontal = value
          break
        }
        case 'horizontalAlign': {
          if (!newStyle.alignment) {
            newStyle.alignment = {}
          }
          newStyle.alignment.horizontal = value
          break
        }
        case 'verticalAlign': {
          if (!newStyle.alignment) {
            newStyle.alignment = {}
          }
          newStyle.alignment.vertical = value
          break
        }
        default:
          newStyle[key] = value
      }
    })
    return newStyle
  }

  function encodeQueryString(url, params) {
    if (!url) {
      url = '';
    }
    if (!params) {
      return url;
    } 
    const keys = Object.keys(params);
    let prefix = '';
    if (url.lastIndexOf('?') === -1) {
      prefix = '?';
    }
    return keys.length ? url + prefix + keys.map(key => encodeURIComponent(key) + "=" + encodeURIComponent(params[key])).join("&") : url;
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

  function generateHeaderRows(headerTree) {
    const queue = [...headerTree]
    const merges = []
    const headerRows = []
    const result = {}
    while (queue.length > 0) {
      const headerItem = queue.shift()
      if (headerItem.children && headerItem.children.length) {
        queue.push(...headerItem.children)
      }
  
      const startRow = headerItem.startRow
      const startCol = headerItem.startCol
      const rowSpan = headerItem.rowSpan
      const colSpan = headerItem.colSpan
  
      for (let i = 0; i < rowSpan; i++) {
        const row = startRow + i
        let headerRow = headerRows[row]
        if (!headerRow) {
          headerRow = []
          headerRows[row] = headerRow
        }
  
        for (let j = 0; j < colSpan; j++) {
          const col = startCol + j
          if (row === startRow && col === startCol) {
            headerRow[col] = headerItem.label
          } else {
            headerRow[col] = null
          }
        }
      }
      if (rowSpan > 1 || colSpan > 1) {
        merges.push({s: {r: startRow, c: startCol}, e: {r: startRow + rowSpan - 1, c: startCol + colSpan - 1}})
      }
    }
    result['merges'] = merges
    result['rows'] = headerRows
    return result
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
          ranges.push({
            s: {
              r: R,
              c: outRow.length
            },
            e: {
              r: R + rowspan - 1,
              c: outRow.length + colspan - 1
            }
          });
        };

        //Handle Value
        outRow.push(cellValue !== "" ? cellValue : null);

        //Handle Colspan
        if (colspan)
          for (var k = 0; k < colspan - 1; ++k) outRow.push(null);
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
    opts.headerRows = typeof opts.headerRows !== 'number' ? 1 : opts.headerRows
    var range = {
      s: {
        c: 10000000,
        r: 10000000
      },
      e: {
        c: 0,
        r: 0
      }
    };
    for (var R = 0; R != data.length; ++R) {
      for (var C = 0; C != data[R].length; ++C) {
        if (range.s.r > R) range.s.r = R;
        if (range.s.c > C) range.s.c = C;
        if (range.e.r < R) range.e.r = R;
        if (range.e.c < C) range.e.c = C;
        var cell = {
          v: data[R][C]
        };
        if (cell.v == null) continue;
        var cell_ref = XLSX.utils.encode_cell({
          c: C,
          r: R
        });

        if (typeof cell.v === 'number') cell.t = 'n';
        else if (typeof cell.v === 'boolean') cell.t = 'b';
        else if (cell.v instanceof Date) {
          cell.t = 'n';
          cell.z = XLSX.SSF._table[14];
          cell.v = datenum(cell.v);
        } else cell.t = 's';

        if (opts.headerStyle && R < opts.headerRows) {
          cell.s = opts.headerStyle
        }

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
        var val = v[j],
          tmp;
        if (opts.formatterMap && opts.formatterMap[j]) {
          if (typeof opts.formatterMap[j] === 'function') {
            tmp = opts.formatterMap[j](val);
            val = !tmp ? val : tmp;
          } else {
            tmp = opts.formatterMap[j][val];
            val = !tmp ? val : tmp;
          }

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
