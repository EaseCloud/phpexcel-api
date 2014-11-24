phpexcel-api
============

用于创建一个 php 处理 excel 的 http api，供其他程序语言调用。

简单的想法：

Post 数据包括上传（Post 请求体）一个模板的 excel 文件，然后输入一个坐标映射到数据的键值对。
然后将数据填入到上传的 excel 模板，并且启动 http 下载。
