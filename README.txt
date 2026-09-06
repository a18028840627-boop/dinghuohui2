Windows 绿色免安装版构建包

这不是 Windows EXE 成品；它需要在一台 Windows 10 / 11（64 位）电脑上构建。

最简单的使用方法：

1. 将整个“AI商品资料整理助手_Windows绿色版构建包”文件夹复制到 Windows 电脑。
2. 安装 64 位 Python 3.11。
   安装时勾选“Add Python to PATH”。
3. 双击 build_portable_windows.bat。
4. 等待依赖安装和打包完成。首次通常需要较长时间。
5. 成品会生成在：
   dist_portable\AI_Product_Assistant\
6. 将整个 AI_Product_Assistant 文件夹复制给其他 Windows 用户。
   用户双击里面的 AI_Product_Assistant.exe 即可运行，无需安装 Python。

重要：

- 这是绿色免安装文件夹版，不能只复制 EXE，必须保留整个文件夹。
- 第一次识别时，PaddleOCR 可能需要联网下载模型；之后会使用电脑本机缓存。
- 目前无法在 Mac 上直接生成可运行的 Windows EXE，因此必须在 Windows 上执行构建脚本。
- 如果之前打包失败，请使用当前文件夹内更新后的 build_portable_windows.bat 重新执行；build_cache 内的文件只是临时缓存，不是成品。
