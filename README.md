phpexcel-api
============

一个 php 处理 excel 的 http api，供其他程序语言调用。

Python SDK: <https://github.com/EaseCloud/python-xlscript>

基本架构：
---------

通过 HTTP 提交一个 Post请求，然后返回一个下载 Excel 的响应。

可以通过一个如下的表单来模拟这个请求：

```http
<form action="xlsapi/index.php" method="post" enctype="multipart/form-data">
    <input type="file" name="template" />
    <textarea name="xlscript"></textarea>
    <input type="submit" />
</form>
```

请求方式是 multipart/form-data，同时提交表单数据以及附件。

因此，请求头应当符合如下格式，这一点请自行参见 multipart/form-data 请求的实现方法：

```
[Request Headers]
Content-Type: Content-Type:multipart/form-data; boundary=BOUNDER_STRING
```

因此，POST 的请求体里面包括两个主要的部分：

1. $_FILE['template']: 用于作为模板的 excel 文件，然后通过。（缺省为空 xls 文件）
2. $_POST['data']: 用于传递填充模板的“脚本”，具体语法后面会进行叙述。

另外，配置参数通过 $_POST['config'] 提供 json 的配置字典，用于覆盖 config 的默认值。

具体配置参数后面会提到。

请求发出之后，本接口会打开模板 xls 文件，然后顺着脚本内容填充数据，最后返回一个响应下载 xls。

具体的响应头形如：

```
[Response Headers]
Content-Disposition: attachment; filename="excel.xls"
Content-Type: application/vnd.ms-excel
```

注意两点：

1. 如果上传的模板文件是 xlsx 的 2007 格式，响应回来的格式会是 xlsx，否则为 xls；
2. 返回的文件名是配置的文件名，缺省为模板的文件名，如模板缺省，则为 excel.xls。

# 填充脚本

本 api 的运作步骤如下：

1. 打开一个模板文件，转到第一个 WorkSheet；
2. 根据配置项*row_delimiter*分割脚本行，每行一个脚本命令；
3. 根据配置项*col_delimiter*分割每个脚本命令，得到命令参数 args；
4. 执行命令，args 第一个参数为命令名称，后面的是命令参数；

## 脚本命令 


### 选择 WorkSheet **SELECT_WORKSHEET** *[OK]*
+ index: 选择 worksheet 的序号（从 0 开始）

例如下面的例子会当前活动的 worksheet 选择为第二个（下标为 1）

```
SELECT_WORKSHEET|1
```

### 填充单元格 **FILL** *[OK]*

+ cell: 单元格位置，例如 A1
+ content: 填充到该单元格的内容

下面的例子会将`今天天气很好，我们去割草。`这句话填充到 A1 单元格中：

```text
FILL|A1|今天天气很好，我们去割草
```

### 填充单元格（根据坐标） **FILL2** *[OK]*

+ col: 列坐标，从 0 起算
+ row: 行坐标，从 1 起算
+ content: 填充到该单元格的内容

例如，3-4 坐标对应的是 D4 单元格，下面例子将`$5.25`填写进 D4 单元格中：

```text
FILL2|3|4|$5.25
```

### 合并单元格 **MERGE** *[OK]*

+ begin: 选取区域的开始单元格
+ end: 选取区域的结束单元格

例如，下面的例子会将 A1-D3 这 12 个单元格合并：

```
MERGE|A1|D3
```

### 调整水平对齐 **ALIGN** *[OK]*

+ begin: 选取区域的开始单元格
+ end: 选取区域的结束单元格
+ pValue: 对齐方式：
  * general: 无
  * left: 左对齐
  * right: 右对齐
  * center: 居中对齐
  * centerContinuous: 连续居中对齐
  * justify:
  * fill:
  * distributed: (Excel2007+ Only)

例如，下面的例子会将 A1-D3 这 12 个单元格设置为右对齐：

```
ALIGN|A1|D3|right
```

### 调整垂直对齐 **VALIGN** *[OK]*

+ begin: 选取区域的开始单元格
+ end: 选取区域的结束单元格
+ pValue: 对齐方式：
  * bottom: 靠下
  * top: 靠上
  * center: 居中
  * justify:
  * distributed: (Excel2007+ Only)

例如，下面的例子会将 A1-D3 这 12 个单元格设置为垂直居中：

```
VALIGN|A1|D3|center
```

### 设置边框样式 **SET_BORDER** *[OK]*

+ begin: 选取区域的开始单元格
+ end: 选取区域的结束单元格
+ border_position: 边框位置
  * all
  * outline
  * inside
  * horizontal
  * vertical
  * top
  * right
  * bottom
  * left
  * diagonal
+ border_style: 边框的样式
  * none
  * dashDot
  * dashDotDot
  * dashed
  * dotted
  * double
  * hair
  * medium
  * mediumDashDot
  * mediumDashDotDot
  * mediumDashed
  * slantDashDot
  * thick
  * thin
+ border_color: ARGB 表示的边框颜色，例如半透明红色为 7FFF0000

例如下面例子将 A3-C5 单元格全部框线设置为普通红色框线

```
SET_BORDER|A3|C5|thin|FFFF0000
```

### 调整区域样式 **STYLE** *[OK]*

+ begin: 选取区域的开始单元格
+ end: 选取区域的结束单元格
+ style: 加粗 BOLD / 斜体 ITALIC / 下划线 UNDERLINE，如果是取消对应的样式，在前面加`~`符号

例如，下面的例子会将 A2 单元格加粗显示：

```
STYLE|A2|A2|BOLD
```

下面的例子取消 A1-F100 范围内所有的加粗效果：

```
STYLE|A1|F100|~BOLD
```

### 设置列宽度 **SET_WIDTH** *[OK]*

+ col: 列名，例如 C 或者 AZ
+ width: 像素宽度

例如将 D 列宽度设置为 22.75 像素

```
SET_WIDTH|D|22.75
```

### 设置列高度 **SET_HEIGHT** *[OK]*

+ row: 行号，例如 1
+ height: 像素高度

例如将第一行高度设置为 18.5 像素

```
SET_HEIGHT|1|18.5
```

### 调整字号 **FONT_SIZE** *[OK]*

+ begin: 选取区域的开始单元格
+ end: 选取区域的结束单元格
+ size: 字号，用数字表示

例如下面的例子会将 A1 单元格字号调成 16

```
FONT_SIZE|A1|A1|16
```

### 调整单元格是否自动换行 **WRAP_TEXT** *[OK]*

+ begin: 选取区域的开始单元格
+ end: 选取区域的结束单元格
+ wrap: 0 不换行，1 换行

例如下面的例子会将 A1 单元格设置成自动换行

```
FILL|A1|this cell\nhas multline values
WRAP_TEXT|A1|A1|1
```

### 设置单元格超链接 **SET_URL** *[OK]*
+ cell: 单元格位置，例如 A1
+ url: 设置单元格链接到的 url

例如下面的例子会给 A1 单元格添加到百度的链接

```
SET_URL|A1|http://www.baidu.com/
```

### 插入整行 **INSERT_ROW** *[OK]*
+ cell: 单元格位置，例如 A1
+ number: 插入行数量，插入number行

例如下面的例子会给 A1 插入10行

```
INSERT_ROW|A1|10
```