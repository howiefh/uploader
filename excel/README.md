# 示例

[demo](https://howiefh.github.io/uploader/excel/demo/index.html)

# 导入
## 快速上手

```
<!-- 引入js -->
<script type="text/javascript" src="//unpkg.com/xlsx/dist/shim.min.js"></script>
<script type="text/javascript" src="//unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
<script src="../excelImporter.js"></script>

<button id="import">Import</button>

<script>
    
    var excelImporter = new ExcelImporter('import', {
      onLoaded: function(res) {
        header = res[0].header;
        rows = res[0].results;
        document.getElementById('container').innerHTML = table(header, rows);
      }
    });
</script>
```

## 参数说明

#### onLoaded

解析完Excel后的回调方法, 回调方法参数为数组类型，数组的元素为一个 Sheet 的数据，包括sheet名、表头数组及数据数组

例如

```
[
  {
    "sheet": "SheetJS",
    "header": [
    "Name",
    "Sex",
    "Score",
    "Date"
    ],
    "results": [
    {
        "Name": "Tom",
        "Sex": "Male",
        "Score": 100,
        "Date": 1.9
    },
    {
        "Name": "Lucy",
        "Sex": "Female",
        "Score": 99,
        "Date": 1.9
    }
    ]
  }
]
```

#### onError

解析Excel错误时的回调方法, 参数为错误信息

#### headerMap

Excel表格头名和返回数据字段名的映射，接受对象和数组两种数据, 如果使用jqGrid直接将colModel传入即可
如：
```
{'姓名': 'name', '性别': 'sex'}
```
或
```
[{label:'姓名', name:'name'}, {label:'性别', name: 'sex'}]
```

#### headerRow

表头所在行，从1起始，默认1

#### bindClick

是否为元素绑定点击事件，默认true

#### onlyFirstSheet

是否只解析第一个sheet，默认true

#### includeUnknowHeader

导出的数据中是否要包含没有在 headerMap 配置的, 默认false

#### includeEmptyHeader

是否包含为空的表头, 默认false

#### dateNF

日期格式 默认 yyyy-MM-dd


# 导出
## 快速上手

```
<!-- 引入js -->
<script type="text/javascript" src="//unpkg.com/xlsx/dist/shim.min.js"></script>
<script type="text/javascript" src="//unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
<script type="text/javascript" src="//unpkg.com/blob.js@1.0.1/Blob.js"></script>
<script type="text/javascript" src="//unpkg.com/file-saver@1.3.3/FileSaver.js"></script>
<script src="../excelExporter.js"></script>

<button id="exportRemote">Export Remote</button>

<script>
    
    var excelRemoteExporter = new ExcelExporter('exportRemote', {
      url: './data.json',
      ajaxData: {
        t: Date.now()
      }
    });

</script>
```

## 参数说明

#### tableId

表格元素 id，需要导出页面表格时配置

#### url

远程地址，需要导出远程数据时配置

#### data

导出数据，需要导出本地数据时配置

#### ajaxData
导出远程数据时，ajax 请求所携带的数据，可以是方法或对象，方法适用于每次导出时参数可能会改变的情况
例如
```
ajaxData: function() {
  return {t: Date.now()}
}
```
或
```
ajaxData: {t: Date.now()}
```

#### filename

导出的文件名，不需要带后缀

#### autoWidth

是否自动适应宽度 默认true

#### header

表头数组，不传的话从headerMap中解析，如果headerMap也没设置则使用fields配置

#### headerMap

Excel表格头名和返回数据字段名的映射，接受对象和数组两种数据, 如果使用jqGrid直接将colModel传入即可

如：
```
{'姓名': 'name', '性别': 'sex'}
```
或
```
[{label:'姓名', name:'name'}, {label:'性别', name: 'sex'}]
```

#### fields

字段名数组，如果没设置从 headerMap 中读取，如果设置，导出时只导出该数组中包含的字段

#### bindClick

是否为元素绑定点击事件，默认true

#### nullToEmpty

是否将null转为空字符串 默认true

#### formatterMap

字段格式化映射，key 为字段名，value为方法或者对象

例如

```
formatterMap: {
  name: function(val) {
    return 'name:' + val;
  },
  sex: {'f':'女','m':'男'}
}
```
