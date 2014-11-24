phpexcel-api
============

一个 php 处理 excel 的 http api，供其他程序语言调用。

基本架构：
---------

通过 HTTP 提交一个 Post请求，然后返回一个下载 Excel 的响应。

可以通过一个如下的表单来模拟这个请求：

```http
<form action="xlsapi/index.php" method="post" enctype="multipart/form-data">
    <input type="file" name="template" />
    <textarea name="data"></textarea>
    <input type="submit" />
</form>
```

请求方式是 multipart/form-data，同时提交表单数据以及附件。

因此，请求头应当符合如下格式，这一点请自行参见 multipart/form-data 请求的实现方法：
> [Request Headers]
> **Content-Type:** Content-Type:multipart/form-data; boundary=BOUNDER_STRING

因此，POST 的请求体里面包括两个主要的部分：

1. $_FILE['template']: 用于作为模板的 excel 文件，然后通过。（缺省为空 xls 文件）
2. $_POST['data']: 用于传递填充模板的“脚本”，具体语法后面会进行叙述。

另外，配置参数通过 $_POST['config'] 提供 json 的配置字典，用于覆盖 config 的默认值。

具体配置参数后面会提到。

请求发出之后，本接口会打开模板 xls 文件，然后顺着脚本内容填充数据，最后返回一个响应下载 xls。

具体的响应头如下：

> [Response Headers]
> **Content-Disposition:** attachment; filename="excel.xls"
> **Content-Type:** application/vnd.ms-excel

注意两点：

1. 如果上传的模板文件是 xlsx 的 2007 格式，响应回来的格式会是 xlsx，否则为 xls；
2. 返回的文件名是配置的文件名，缺省为模板的文件名，如模板缺省，则为 excel.xls。

# 填充脚本

本 api 的运作步骤如下：

1. 打开一个模板文件，转到第一个 WorkSheet；
2. 根据配置项*row_delimeter*分割脚本行，每行一个脚本命令；
3. 根据配置项*col_delimeter*分割每个脚本命令，得到命令参数 args；
4. 执行命令，args 第一个参数为命令名称，后面的是命令参数；

## 脚本命令 

### 1. 填充单元格 **F**：

+ cell: 单元格位置，例如 A1
+ content: 填充到该单元格的内容

下面的例子会将`今天天气很好，我们去割草。`这句话填充到 A1 单元格中：

```text
F|A1|今天天气很好，我们去割草
```

### 2. 填充单元格（根据坐标） **FC**：

+ col: 列坐标，从 0 起算
+ row: 行坐标，从 1 起算
+ content: 填充到该单元格的内容

例如，3-4 坐标对应的是 D4 单元格，下面例子将`$5.25`填写进 D4 单元格中：

```text
FC|3|4|$5.25
```
