# pptx_readonly_decrypt

##项目简介
pptx_readonly_decrypt 是一个 Python 脚本，用于解除 PowerPoint (.pptx) 文件的只读保护。这通过删除 .pptx 文件中的 <p:modifyVerifier> 标签来实现，从而允许用户编辑文件。

##使用方法
###解密单个文件
要解密单个 .pptx 文件，可以使用以下命令：
python pptx_readclear.py --source_path source_flie --target_path target_file
这里，--source_path 参数指定要解密的 .pptx 文件的路径，--target_path 参数指定解密后的文件保存的位置。

###解密整个目录
要递归地解密目录中的所有 .pptx 文件并复制目录结构，可以使用以下命令：
python pptx_readclear.py --source_path source_dir --target_path target_dir
这里，--source_path 参数指定包含 .pptx 文件的目录路径，--target_path 参数指定解密后的文件和目录结构保存的位置。
