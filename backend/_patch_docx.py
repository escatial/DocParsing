"""\u4e34\u65f6\u811a\u672c\uff1a\u4fee\u6539 download_format \u4ee5\u4f7f\u7528\u811a\u6ce8\u751f\u6210\u7684 docx"""
import io
import os

filepath = r'd:\code\個人開發項目\202607\文檔解析\backend\main.py' if False else (
    r'd:\code\個人開發項目\202607\文檔解析\backend\main.py'
)
# \u4f7f\u7528\u7eaf\u5b57\u7b26\u4e32\u8def\u5f84\u4ee5\u514d\u4e2d\u6587\u8def\u5f84\u95ee\u9898
filepath = r'D:\code\個人開發項目\202607\文檔解析\backend\main.py'
filepath = filepath.replace('個人開發項目', '\u500b\u4eba\u958b\u767c\u9805\u76ee')
filepath = filepath.replace('文檔解析', '\u6587\u6a64\u89e3\u6790')
# \u4e0d\u8981\u8fd9\u4e48\u590d\u6742\uff0c\u76f4\u63a5\u4f7f\u7528\u539f\u59cb\u5b57\u7b26\u4e32
filepath = 'd:\\code\\\u500b\u4eba\u958b\u767c\u9805\u76ee\\202607\\\u6587\u6a64\u89e3\u6790\\backend\\main.py'

print('Path:', filepath)
print('Exists:', os.path.exists(filepath))