HTTP 请求格式：
--------------

本 http 接口采用 multipart/form-data 的请求方式提交数据，下面来具体讲述一下这种提交方式：

# 请求样例（省略了非关键的内容）

```
[Request Header]
Content-Type: multipart/form-data; boundary=--FORM-BOUNDARY

[Request Payload]
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


