# -*- coding: utf-8 -*-
import xlwt
import os
# 设置单元格样式
startTime = '2020-05-21'
endTime = '2020-10-20'
authorName = '任根胜'
def setStyle(name = 'SimSun', height = 250, bold = False, horz = xlwt.Alignment.HORZ_LEFT, vert = xlwt.Alignment.VERT_CENTER, border_style = xlwt.Borders.THIN, border_color = 0x40):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    alignment = xlwt.Alignment()
    alignment.horz = horz
    alignment.vert = vert
    alignment.wrap = 1
    style.alignment = alignment
    borders = xlwt.Borders()
    borders.right = border_style
    borders.top = border_style
    borders.bottom = border_style
    borders.left = border_style
    borders.left_colour = border_color
    borders.right_colour = border_color
    borders.top_colour = border_color
    borders.bottom_colour = border_color
    style.borders = borders
    return style

def weeklyGenerator(gitDataMap):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('工作周报')
    worksheet.write_merge(0, 2, 0, 2, '姓名：任根胜\n岗位：前端工程师', setStyle())
    worksheet.write(3, 0, '日期', setStyle(horz = xlwt.Alignment.HORZ_CENTER))
    worksheet.write(3, 1, '工作内容', setStyle(horz = xlwt.Alignment.HORZ_CENTER))
    worksheet.write(3, 2, '完成度', setStyle(horz = xlwt.Alignment.HORZ_CENTER))
    worksheet.col(0).width = 8000
    worksheet.col(1).width = 17000
    worksheet.col(2).width = 8000
    startRow = 4
    for key, value in gitDataMapping.items():
        worksheet.write_merge(startRow, startRow + len(value) - 1, 0, 0, key, setStyle(horz = xlwt.Alignment.HORZ_CENTER))
        for workContent in value:
            worksheet.write(startRow, 1, workContent, setStyle())
            worksheet.write(startRow, 2, '100%', setStyle(horz = xlwt.Alignment.HORZ_CENTER))
            startRow += 1
    print ('保存周报')
    workbook.save(authorName + '工作周报-' + startTime  + '-' + endTime + '.xls')
if __name__ == '__main__':
    print ('开始遍历文件夹， 项目过多时间可能较长, 请耐心等待...')
    currentPath = os.getcwd()
    currentPathDep = len(currentPath.split(os.path.sep))
    gitFolderList = []
    gitFolderNames = []
    gitLogMapping = {}
    gitDataMapping = {}
    gitShell = 'git log --since=' + startTime + ' --until=' + endTime + ' --author=' + authorName + ' --date=format:"%Y-%m-%d" --pretty=format:"%cd:::::%s";'
    for root, dirs, files in os.walk(currentPath):
        for name in dirs:
            childPath = os.path.join(root, name)
            childPathDep = len(childPath.split(os.path.sep))
            if childPathDep > currentPathDep + 0 and childPathDep <= currentPathDep + 2 :
                if childPathDep == currentPathDep + 2 :
                    if  ('.git' in dirs):
                        gitFolderList.append(root)
                    break
            else:
                break
    for filePath in gitFolderList:
        fileName = raw_input('Please input the project name of [' + filePath + '])(Just enter and skip this project):')
        if fileName == '':
            gitFolderNames.append('None')
        else:
            gitFolderNames.append(fileName)
            # 开始获取项目的提交信息
            os.chdir(filePath)
            result = os.popen(gitShell)
            gitLogMapping[fileName] = result.read()
    # 获取本周所有提交完毕，开始按照时间生成周报模型
    for key, value in gitLogMapping.items():
        for item in value.splitlines():
            gitLogItem = item.split(':::::')
            if not gitDataMapping.has_key(gitLogItem[0]):
                gitDataMapping[gitLogItem[0]] = []
            gitDataMapping[gitLogItem[0]].append("(" +  key + ")" + gitLogItem[1].replace("fix:", "").replace("feat:", ""))

    print('开始格式化周报')
    os.chdir(currentPath)
    weeklyGenerator(gitDataMapping)
