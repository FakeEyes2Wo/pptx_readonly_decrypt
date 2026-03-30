# pptx_readonly_decrypt

## 项目简介
pptx_readonly_decrypt 是一个 Python 脚本，用于解除 PowerPoint (.pptx) 文件的只读保护。这通过删除 .pptx 文件中的 <p:modifyVerifier> 标签来实现，从而允许用户编辑文件。

## 使用方法
### 原位置解密：
```bash
python pptx_readclear.py /path/to/files --in-place
# 也可以使用简写
python pptx_readclear.py /path/to/files -i
```
### 指定路径保存：
```bash
python pptx_readclear.py /path/to/files -t /path/to/save
```