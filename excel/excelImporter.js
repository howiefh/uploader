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
    /**
     * 解析完Excel后的回调方法, 回调方法参数为数组类型，数组的元素为一个 Sheet 的数据，包括sheet名、表头数组及数据数组
     * 例如
     * [
     *   {
     *     "sheet": "SheetJS",
     *     "header": [
     *       "Name",
     *       "Sex",
     *       "Score",
     *       "Date"
     *     ],
     *     "results": [
     *       {
     *         "Name": "Tom",
     *         "Sex": "Male",
     *         "Score": 100,
     *         "Date": 1.9
     *       },
     *       {
     *         "Name": "Lucy",
     *         "Sex": "Female",
     *         "Score": 99,
     *         "Date": 1.9
     *       }
     *     ]
     *   }
     * ]
     *
     *
     *
     */
    onLoaded: null,
    /**
     * 解析Excel错误时的回调方法, 参数为错误信息
     */
    onError: null,
    /**
     * Excel表格头名和返回数据字段名的映射，接受对象和数组两种数据, 如果使用jqGrid直接将colModel传入即可
     * 如：
     * {'姓名': 'name', '性别': 'sex'}
     * 或
     * [{label:'姓名', name:'name'}, {label:'性别', name: 'sex'}]
     */
    headerMap: null,
    /**
     * 表头所在行，从1起始，默认1
     */
    headerRow: 1,
    headerIndex: 0,
    /**
     * 是否为元素绑定点击事件，默认true
     */
    bindClick: true,
    /**
     * 是否只解析第一个sheet，默认true
     */
    onlyFirstSheet: true,
    /**
     * 导出的数据中是否要包含没有在 headerMap 配置的, 默认false
     */
    includeUnknowHeader: false,
    /**
     * 是否包含为空的表头, 默认false
     */
    includeEmptyHeader: false,
    /**
     * 是否按原生类型解析，如果为false按string解析，默认true
     */
    raw: true,
    /**
     * 日期格式 默认 yyyy-MM-dd
     */
    dateNF: 'yyyy-MM-dd',
    /**
     * 必填字段名数组 校验参数
     */
    requiredFields: [],
    /**
     * 数值字段名数组 校验参数
     */
    numberFields: [],
    /**
     * 日期字段名数组 校验参数
     */
    dateFields: [],
    /**
     * 重复字段名数组 校验参数
     */
    duplicateFields: [],
    /**
     * 字段默认值映射 校验参数
     */
    defaultFields: {},
    /**
     * 如果defaultFields中字段映射的值是以 $ 开头，则从这里取值 校验参数
     *
     * 例如
     * defaultFields: {name: '$username'}
     * defaultValues: {'$username': 'howie'}
     * 则name默认值是 howie
     *
     */
    defaultValues: {}
  };

  /**
   * Excel 导入器
   * @param {*} element 元素 id
   * @param {*} options 配置参数
   */
  function ExcelImporter(element, options) {
    var self = this;
    if (typeof element !== 'string') {
      options = element || {};
      options.bindClick = false;
    }
    self.opts = Object.assign({}, defaults, options);
    self.setHeaderMap(self.opts.headerMap);
    self.setHeaderRow(self.opts.headerRow);

    if (self.opts.bindClick) {
      self.el = getByID(element);
      self.input = create('input', {
        type: 'file',
        accept: '.xlsx, .xls',
        style: 'display: none; z-index: -9999;'
      });
      self.el.parentElement.appendChild(self.input);
      bind(self.input, 'change', function (e) {
        var files = e.target.files;
        var itemFile = files[0]; // only use files[0]
        if (!itemFile) return;
        readerData(itemFile, self.opts);
        self.input.value = null; // fix can't select the same excel
      });
      bind(self.el, 'click', function (e) {
        self.input.click()
      });
    }
  }

  function checkUploadData(uploadData, opts) {
    var requiredFields = opts.requiredFields || [],
      duplicateFields = opts.duplicateFields || [],
      numberFields = opts.numberFields || [],
      dateFields = opts.dateFields || [],
      defaultFields = opts.defaultFields || {},
      defaultValues = opts.defaultValues || {};
    var error = [], field, val;
    if (!uploadData || uploadData.length === 0) {
      error.push('没有数据上传');
      return error;
    }

    var duplicateKeys = [];
    var duplicateData = [];
    for (var i = 0, row = 1; i < uploadData.length;) {
      var item = uploadData[i];

      for (var r = 0; r < requiredFields.length; r++) {
        field = requiredFields[r];
        val = item[field];
        var type = typeof val;
        var valStr = '' + val;
        if (type === 'null' || type === 'undefined' || valStr.trim() === '') {
          if (typeof defaultFields[field] === 'undefined') {
            error.push('第' + row + '条数据, ' + field + ' 为空');
          } else {
            var dVal = defaultFields[field];
            item[field] = (dVal + '').indexOf('$') === 0 ? defaultValues[dVal] : dVal;
          }
        }
      }
      for (var n = 0; n < numberFields.length; n++) {
        field = numberFields[n];
        val = item[field];
        if (val && !/^(-?\d+)(\.\d+)?$/.test(val)) {
          error.push('第' + row + '条数据, ' + field + ' 不是数字');
        }
      }

      for (var d = 0; d < dateFields.length; d++) {
        field = dateFields[d];
        val = item[field];
        if (val && !/^[1-9]\d{3}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])\s+(20|21|22|23|[0-1]\d):[0-5]\d:[0-5]\d$/.test(val)) {
          var newVal = val.replace('年', '/').replace('月', '/').replace('日', '').trim();
          var date = Date.parse(newVal);
          if (isNaN(date)) {
            error.push('第' + row + '条数据, ' + val + ' 不是正确的日期格式');
          } else {
            date = new Date(date);
            item[field] = date.format('yyyy-MM-dd hh:mm:ss');
            error.push('第' + row + '条数据, ' + val + ' 不是正确的日期格式');
          }
        }
      }

      var duplicateKey = '';
      for (var du = 0; du < duplicateFields.length; du++) {
        field = duplicateFields[du];
        val = item[field];
        duplicateKey += val;
      }

      if (duplicateKeys.indexOf(duplicateKey) >= 0) {
        duplicateData.push(item);
        uploadData.splice(i, 1);
        error.push('第' + row + '条数据, ' + duplicateKey + ' 重复');
        row++;
      } else {
        duplicateKeys.push(duplicateKey);
        row++;
        i++;
      }
    }

    return {
      error: error.length > 0,
      duplicate: duplicateData.length > 0,
      errors: error
    };
  }

  ExcelImporter.prototype = {
    /**
     * 设置 Excel表格头名和返回数据字段名的映射，接受对象和数组两种数据
     * @param {*} headerMap Excel表格头名和返回数据字段名的映射
     */
    setHeaderMap: function (headerMap) {
      this.opts.headerMap = checkAndConvertHeaderMap(headerMap);
    },
    /**
     * 设置表头所在行，从1起始
     * @param {*} row 行号，从1起始
     */
    setHeaderRow: function (row) {
      this.opts.headerRow = row;
      this.opts.headerIndex = headerIndex(this.opts.headerRow);
    },
    /**
     * 加载 Excel 文件
     * @param {} file 文件对象
     * @param {*} opts 参数，见参数说明
     */
    loadExcel: function (file, opts) {
      if (!file) return;
      opts.headerMap = checkAndConvertHeaderMap(opts.headerMap);
      opts = Object.assign({}, this.opts, opts);
      readerData(file, opts);
    },
    /**
     * 检查数据, 可以做非空校验、数值类型校验、日期类型校验、重复值校验
     *
     * @param {*} data 待校验的数组 对应 onLoaded 回调方法参数中的 results
     * @param {*} opts 校验参数 见参数说明
     * @returns {error:true, duplicate:true, errors:[]}
     */
    checkData: function (data, opts) {
      opts = Object.assign({}, this.opts, opts);
      return checkUploadData(data, opts)
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

  function checkAndConvertHeaderMap(headerMap) {
    var newMap = {};
    if (isArray(headerMap)) {
      for (var i = 0; i < headerMap.length; i++) {
        var item = headerMap[i];
        newMap[item.label] = item.name;
      }
    } else if (typeof headerMap === 'object') {
      newMap = headerMap;
    }
    return newMap;
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
          raw: opts.raw,
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
    reader.onerror = function (ev) {
      if (typeof opts.onError === 'function') {
        opts.onError(ev)
      }
    };
    reader.readAsArrayBuffer(itemFile)
  }

  function fixData(data) {
    var o = '';
    var l = 0;
    var w = 10240;
    for (; l < data.byteLength / w; ++l) {
      o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
    }

    o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
    return o;
  }

  function getHeaderRow(sheet, range, opts) {
    var headers = [];
    var C;
    var R = range.s.r;
    /* start in the first row */
    for (C = range.s.c; C <= range.e.c; ++C) {
      /* walk every column in the range */
      var cell = sheet[XLSX.utils.encode_cell({
        c: C,
        r: R
      })];
      /* find the cell in the first row */
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
