HTTP 请求格式：
--------------

本 http 接口采用 multipart/form-data 的请求方式提交数据，下面来具体讲述一下这种提交方式：

# 请求样例（省略了非关键的内容）


**[Request Header]**
```
Content-Type: multipart/form-data; boundary=FORM-BOUNDARY
```

**[Request Payload]**
```
--FORM-BOUNDARY
Content-Disposition: form-data; name="template"; filename="人员名录.xls"
Content-Type: application/vnd.ms-excel

### BINARY DATA ###
--FORM-BOUNDARY
Content-Disposition: form-data; name="data"

FILL|A1|姓名
FILL|B1|年龄
FILL|A2|张三
FILL|B2|18
FILL|A3|李四
FILL|B4|22
STYLE|A1|A2|BOLD
--FORM-BOUNDARY
Content-Disposition: form-data; name="data"

--FORM-BOUNDARY
Content-Disposition: form-data; name="config"

### CONFIG_JSON ###
--FORM-BOUNDARY--
```

# 样例的解释

## Header 部分

首先，我们看 `Request Header`，这里关键的一个头信息是 `Content-Type`，我们先是指定 `multipart/form-data`，来说明提交请求的方式（这会告诉 HTTP 服务我们后面跟着的请求体是按照这种方式——顾名思义是分成多部分的表单数据——来提交请求的）。

然后 Content-Type 后面还指定了一个参数 `boundary=FORM-BOUNDARY` 参数，这个参数是用来区分请求的的分块的，我们可以随意指定这个分割符，只要足够复杂，不会误解即可。这个分隔符必须与请求体的分隔符相匹配。

## Body 结构

后面就是 Request Payload 的内容，也就是请求的 Body 部分：

我们可以看到请求体以 `--FORM-BOUNDARY` 开始每一个块，然后最后以 `--FORM-BOUNDARY--` 结束。

也就是说，块的起始是 `--` 两个减号紧跟分隔符，最后的结束是分割符前后都加上一个 `--`。


然后是块内部的结构：

## 块的头部

每个块开头紧接着的是块的 header 信息（类似请求头），注意没有空行，下面分两类讨论：

### 第一种是文件上传

这种情况，块头信息主要有两行:

1. `Content-Disposition`: 先指定 form-data；然后指定 name 对应于 `<input type="file" name="filename" />` 里面填写的 `name` 属性；然后 `filename` 属性为上传文件的原始文件名；
2. `Content-Type`: 这是上传的文件的 MIME 类型，在我们这个问题域中，如果是 xls (Excel5) 格式，对应于 `application/vnd.ms-excel`，如果是 xlsx (Excel2007) 格式，对应的 MIME 类型是 `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`；

于是，这样的提交在 PHP 终究可以通过 `$_FILES['filename']` 来获取。

### 第二种是普通的表单字段

1. 同样需要指定 `Content-Disposition` 信息，里面的 name 属性同样对应于 `<input>` 标签里面的 name 属性；

## 块的内容

然后就是块的内容，注意在块 header 后面要加一个空行（重要！），然后后面再来写请求体的内容。

如果是文件上传，那么这里应该放置待上传文件的二进制流；如果是表单字段，那就应当是表单字段的文本内容直接写在里面。

OK，只要是用这种方式进行提交，就可以跟我们的 api 对接上了，因为 api 里面是通过 `$_FILES['template']` 来获取文件，然后通过 `$_POST['data']` 获取脚本，并通过 `$_POST['config']` 来获取配置的。

最后提醒一下，所有的请求体 Payload 里面的空行，都是以 `\r\n` 进行换行的。