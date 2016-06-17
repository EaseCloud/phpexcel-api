import os
import re
import json
from urllib.request import urlopen, Request
from urllib.parse import quote, unquote, urlencode
import mimetypes
import requests

   
def get_content_type(file_path):
    return mimetypes.guess_type(file_path)[0] or 'application/octet-stream'


def encode_multipart_formdata(fields, files=()):
    """
    fields is a sequence of (name, value) elements for regular form fields.
    files is a sequence of (name, filepath) elements for data to be uploaded as files
    Return (content_type, body) ready for httplib.HTTP instance
    :param fields: The fields to encode.
    :param files: The file-type field set.
    """
    boundary = '----xlsapi-boundary'
    lines = []
    for (key, value) in fields:
        lines.append('--' + boundary)
        lines.append('Content-Disposition: form-data; name="%s"' % key)
        lines.append('')
        lines.append(value)
    for (key, file_path) in files:
        lines.append('--' + boundary)
        lines.append(
            'Content-Disposition: form-data; name="%s"; filename=%s' %
            (key, os.path.basename(file_path))
        )
        lines.append('Content-Type: %s' % get_content_type(file_path))
        lines.append('')
        lines.append(open(file_path, 'rb').read())
    lines.append('--' + boundary + '--')
    body = b'\r\n'.join([l.encode() if type(l) == str else l for l in lines])
    content_type = 'multipart/form-data; boundary=%s' % boundary
    return content_type, body


def render_excel_response(xlscript='',
                          template=(),
                          api_url='http://excelapi',
                          config=None):
    """发送请求处理一个 xlscript 渲染
    返回一个 http.client.HTTPResponse 对象，该响应对应于一个 excel 文件的下载响应。
    :param xlscript: 提交的 xlscript 脚本文本
    :param template: (可选) excel 模板的路径
    :param config: (可选) 配置项参数
    :param api_url: phpexcel-api 的服务 url
    :returns: :class:`Response <Response>` object.
    """
    content_type, body = encode_multipart_formdata({
        'xlscript': xlscript,
        'config': json.dumps(config or {}),
    }.items(), template and [('template', template)]
    )

    req = Request(api_url, data=body, headers={'Content-Type': content_type})
    return urlopen(req)


def render_excel(xlscript='', template=(), api_url='http://excelapi', config=None):
    """请求一个 xlscript 渲染，返回文件名，文件二进制内容，以及 mime 类型
    :param xlscript: 渲染的 xlscript 脚本
    :param template:
    :return: (file_name, bytes, mime_type)
    """
    response = render_excel_response(xlscript, template, api_url, config)
    assert response.status == 200, '接口调用失败，返回状态码：'+response.status
    headers = dict(response.getheaders())
    # print(headers)
    # print(response.read().decode())
    return (
        unquote(re.findall('filename="(.+)"', headers['Content-Disposition'])[0]),
        response.read(),
        headers.get('Content-Type'),
    )


def render_xlscript(xlscript):

    fname, data, mime = render_excel(
        xlscript,
        config={'row_delimiter': '$$$', 'col_delimiter': '|'},
    )
    resp = HttpResponse(data, content_type=mime)
    resp['Content-Disposition'] = 'attachment; ' + urlencode({'filename': fname})
    return resp
