### start
安装依赖： pip install -r requirements.txt

### step1
将文件放在 files 文件夹下，暂不支持嵌套文件夹，默认doc格式

### step2
运行
```
python transfer.py
```
可以看到 filesT 文件夹下，出现了对应的docx文件

### step3
修改 index.py 里面的
```
...
targetPath = "D:\\program\\wordparse\\filesT\\"

words = "财务报表审计"  // 修改该字段

newDoc = Document()
...

```

### step4 
运行
```
python index.py
```
根目录下会出现 综合.docx 文件

### step5
可以在 index 的循环存储部分调整样式
```
def saveFile(document, startNum):
    for i in range(8):
        if i == 0:
            newDoc.add_paragraph(document.paragraphs[startNum + i].text, style='List Bullet')
        else :
            newDoc.add_paragraph(document.paragraphs[startNum + i].text, style='Normal')
    newDoc.add_paragraph()
```