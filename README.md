# tableExtractor

### version
V1

### Note
对pdf文件的处理可能会出现表格错位的情况，建议用office先转成word再进行处理
具体步骤是:
  1. 右键pdf文件 然后选择MS word 打开
  2. 碰见警告点击确认
  3. 另存为 docx 文件

对全是图片的pdf文件需要先转为png 或者 jpg 才能进行识别

### Usage
安装需要的库
```bash
pip install -r requirements.txt
```

直接在根目录运行
```bash
python main.py
```

在命令行输入:
```bash
请输入SecretId: #从腾讯云获取
请输入SecretKey: #从腾讯云获取
请输入文件夹路径(相对or绝对): # 直接回车表示默认处理在 "rawFile/" 路径下文件
```

输出表格在 table/ 中可以找到


