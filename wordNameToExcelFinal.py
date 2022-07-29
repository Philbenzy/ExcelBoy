import os,time
from cv2 import sort
import xlsxwriter

# -----------------资料目录地址-------------------- #
filePath = r'C:\\Users\\Wzy\\Desktop\\wsytest\\' 
# -----------------Excle存放地址-------------------- #
exclePath = r'C:\\Users\\Wzy\\Desktop\\test.xlsx'

filesNum = 0 # 附件个数
finalNum = 0

workbook =xlsxwriter.Workbook(exclePath)
sheet = workbook.add_worksheet()

sheet.write(0, 0, "发文编号")
sheet.write(0, 1, "发文名称（按照序号）")
sheet.write(0, 2, "修改日期")
sheet.write(0, 3, "附件个数")
count = 1
# 返回一个名称列表
originalList = os.listdir(filePath) # 包含PDF文件
print(originalList)
# 剔除列表中pdf文件
print('-------------------忽略文件----------------------')
for entry in originalList:
    index = entry.find('.pdf')
    if(index != -1):
        originalList.remove(entry)
        print("忽略文件         | 文件名称为："+ entry)
print('-------------------文件排序----------------------')
# 依据序号进行排序
originalList.sort(key=lambda x:int(x.split('.')[0]))
for entry in originalList:
    # 忽略PDF文件（上面有，到底删不删？）
    if(entry.endswith('.pdf') or entry.endswith('.zip')):
        print("忽略文件         | 文件名称为："+ entry)
        continue
    print(entry)
    filemt = time.localtime(os.stat(filePath + entry).st_mtime)

    # 判断如果是文件夹，则统计个数
    if os.path.isdir(filePath + entry):
        files = os.listdir(filePath + entry)
        filesNum = len(files)
    
    # 确定最终附件个数
    if(filesNum > 0):
        finalNum = filesNum - 1


    # 确定发文名称
    beginIndex = entry.find("关于")
    endIndex = entry.find(".docx")
    fileFinalName = entry

    # 确定发文编号
    articleBeginIndex = entry.find(".")
    articleEndIndex = entry.find("号")
    if(articleBeginIndex == -1 or articleEndIndex == -1):
        print("发文编号匹配失败 | 文件名称为：" + entry)
        continue
    
    if(endIndex == -1):
        # 如果没有查找到结尾，则判断为文件夹，全部输出
        fileFinalName = entry[beginIndex:]
    else:
        # 如果查找到结尾，则判断为word文件，输出中间部分 关于xxxx (.docx自动忽略)
        fileFinalName = entry[beginIndex:endIndex]


    # 整理写入的行元素，名称，日期，附件个数
    sheet.write(count, 0, entry[articleBeginIndex + 1:articleEndIndex + 1]) # 发文编号
    sheet.write(count, 1, fileFinalName) # 发文名称
    sheet.write(count, 2, time.strftime("%Y-%m-%d",filemt)) # 日期
    sheet.write(count, 3 , finalNum) # 附件个数
    finalNum = 0 # 清零
    filesNum = 0 # 清零
    count = count + 1
# 关闭文件
workbook.close()