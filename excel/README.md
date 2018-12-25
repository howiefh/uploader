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
        console.log(rows);
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

#### raw

是否按原生类型解析，如果为false按string解析，默认true

#### dateNF

日期格式 默认 yyyy-MM-dd

#### requiredFields

必填字段名数组 校验参数

#### numberFields

数值字段名数组 校验参数

#### dateFields

日期字段名数组 校验参数

#### duplicateFields

重复字段名数组 校验参数

#### defaultFields

字段默认值映射 校验参数

#### defaultValues

如果defaultFields中字段映射的值是以 $ 开头，则从这里取值 校验参数

例如
```
defaultFields: {name: '$username'},
defaultValues: {'$username': 'howie'}
```
则name默认值是 howie

## 方法

#### setHeaderMap(headerMap)

设置 Excel表格头名和返回数据字段名的映射，接受对象和数组两种数据

* headerMap Excel表格头名和返回数据字段名的映射

#### setHeaderRow(row)

设置表头所在行，从1起始

* row 行号，从1起始

#### loadExcel(file, opts)

加载 Excel 文件

* file 文件对象
* opts 参数，见参数说明

#### checkData(data, opts)

检查数据, 可以做非空校验、数值类型校验、日期类型校验、重复值校验

* data 待校验的数组 对应 onLoaded 回调方法参数中的 results
* opts 校验参数 见参数说明

* 返回 { error:true, duplicate:true, errors:[] } error: 是否错误， duplicate: 是否有重复值，errors: 错误信息数组

## 示例

```
var excelImporter = new ExcelImporter();
excelImporter.loadExcel(file, {
    headerMap: {'姓名':'name', '性别':'sex', '手机号':'mobileNo'}
    onError: function (e) {
        alert('解析excel文件失败');
    },
    onLoaded: function (json) {
        var excel = json[0]['results'];
        var opts = {
            requiredFields: ['name','mobileNo'],
            duplicateFields: ['mobileNo']
        }

        var error = excelImporter.checkData(excel, opts);
        if (error.error) {
            var message = '';
            if (error.duplicate) {
                message = 'Excel 中有重复数据，请核实\n';
            }
            alert(message + error.errors.join('\n'));
            return false;
        }

        $.ajax({
          url: 'upload',
          type: "POST",
          data: {
            'content': JSON.stringify(excel)
          },
          dataType: 'json',
          success: function (res) {
            // 处理结果
          },
          error: function (XMLHttpRequest, textStatus, errorThrown) {
            // 处理异常
          }
        });
    }
});
```

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

tableId， url, data 三个设置一个即可，都设置时，优先级由高到低为 url, tableId, data

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

## 方法
#### exportJsonToExcel(data, header, fields, filename, autoWidth) {

导出json数据

* data json 数组
* header 表头 可以不传 默认使用配置参数
* fields 表头对应的字段名 可以不传 默认使用配置参数
* filename 文件名 可以不传 默认使用配置参数
* autoWidth 是否自动调整宽度 可以不传 默认使用配置参数

#### exportTableToExcel(id) {

导出页面表格

* id 表格id

## 示例

```
new ExcelExporter('exportBtn', {
    url: '/users/export',
    filename: '用户-' + Date.now(),
    headerMap: colModel,
    formatterMap: {
        sex: {'f': '女', 'm':'男'},
        modifiedDate: formatDate
    },
    ajaxData: function() {
        return {
          id: $("#id").val()
        }
    },
    fields : ['name', 'sex', 'mobileNo', 'modifiedDate']
});

new ExcelExporter('exportSelect', {
    filename: '用户-' + Date.now(),
    headerMap: colModel,
    data: function() {
        return selectedRowObjArray($grid);
    },
    formatterMap: {
        sex: {'f': '女', 'm':'男'},
        modifiedDate: formatDate
    },
    fields : ['name', 'sex', 'mobileNo', 'modifiedDate']
});
```
