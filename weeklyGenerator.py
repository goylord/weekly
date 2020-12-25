# -*- coding: utf-8 -*-
import xlwt
import os
import datetime
# 设置单元格样式
startTime = '2020-01-01'
endTime = '2020-12-31'
authorName = 'RenGS'
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
def getCurrentWeek():
    monday, today = datetime.date.today(), datetime.date.today()
    one_day = datetime.timedelta(days=1)
    while monday.weekday() != 0:
        monday -= one_day
    monday -= one_day
    return monday, today
def weeklyGenerator(gitDataMap):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('工作周报')
    worksheet.write_merge(0, 2, 0, 2, '姓名：' + authorName + '\n岗位：前端工程师', setStyle())
    worksheet.write(3, 0, '日期', setStyle(horz = xlwt.Alignment.HORZ_CENTER))
    worksheet.write(3, 1, '工作内容', setStyle(horz = xlwt.Alignment.HORZ_CENTER))
    worksheet.write(3, 2, '完成度', setStyle(horz = xlwt.Alignment.HORZ_CENTER))
    worksheet.col(0).width = 8000
    worksheet.col(1).width = 17000
    worksheet.col(2).width = 8000
    startRow = 4
    for key in sorted(gitDataMap.keys()):
        worksheet.write_merge(startRow, startRow + len(gitDataMap[key]) - 1, 0, 0, key, setStyle(horz = xlwt.Alignment.HORZ_CENTER))
        for workContent in gitDataMap[key]:
            worksheet.write(startRow, 1, workContent, setStyle())
            worksheet.write(startRow, 2, '100%', setStyle(horz = xlwt.Alignment.HORZ_CENTER))
            startRow += 1
    print ('保存周报')
    workbook.save(authorName + '工作周报-' + startTime  + '-' + endTime + '.xls')
if __name__ == '__main__':
    currentPath = os.getcwd()
    tmpDataMapping = {}
    gitFolderList = []
    gitFolderNames = []
    gitLogMapping = {}
    gitDataMapping = {}
    menuIndex = '1'
    if os.path.exists(currentPath + os.path.sep + '.tmp.txt'):
        # 如果存在缓存文件
        print("1. 使用缓存中的项目名称\n")
        print("2. 重新获取项目名称\n")
        print("3. 扫描添加新项目\n")
        menuIndex = raw_input("请输入选项（默认为1）：")
        if (menuIndex == '1' or menuIndex == '3'):
            tmpFileOpen = open(currentPath + os.path.sep + '.tmp.txt')
            tmpFile = tmpFileOpen.read()
            tmpFileOpen.close()
            tmpFileArrays = tmpFile.split(':::::::')
            for tmpFileLine in tmpFileArrays:
                keyValue = tmpFileLine.split('&&&&&&&&&&&&&&&&&&')
                if len(keyValue) is 2:
                    tmpDataMapping[keyValue[0]] = keyValue[1]
                    gitFolderList.append(keyValue[0])
                    gitFolderNames.append(keyValue[1])
    else:
        menuIndex = '2'
    startTime = raw_input("生成周报提交开始时间（默认周一）：")
    endTime = raw_input("生成周报提交结束时间（默认当天）：")
    currentMonday, currentFriday = getCurrentWeek()
    if startTime == '':
        startTime = currentMonday.strftime('%Y-%m-%d')
    if endTime == '':
        endTime = currentFriday.strftime('%Y-%m-%d')
    authorName = raw_input("人员姓名（默认为空）：")
    gitShell = 'git log --since=' + startTime + ' --until=' + endTime + ' --author=' + authorName + ' --date=format:"%Y-%m-%d" --pretty=format:"%cd:::::%s";'
    if menuIndex == '2' or menuIndex == '3':
        for dir in os.listdir(currentPath):
            if os.path.exists(dir + os.path.sep + '.git'):
                childProjectPath = currentPath + os.path.sep + dir
                if menuIndex == '2' or (menuIndex == '3' and (childProjectPath not in gitFolderList)):
                    gitFolderList.append(childProjectPath)
    for folderIdx, filePath in enumerate(gitFolderList):
        fileName = ''
        if menuIndex == '2':
            fileName = raw_input('Please input the project name of [' + filePath + '])(Just enter and skip this project):')
        elif menuIndex == '3' and not tmpDataMapping.has_key(filePath):
            fileName = raw_input('Please input the project name of [' + filePath + '])(Just enter and skip this project):')
        if fileName == '' and (menuIndex == '2' or menuIndex == '3' and not tmpDataMapping.has_key(filePath)):
            gitFolderNames.append('None')
        else:
            if menuIndex == '2' or (menuIndex == '3' and not tmpDataMapping.has_key(filePath)):
                gitFolderNames.append(fileName)
            else:
                fileName = gitFolderNames[folderIdx]
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

    print('开始生成周报...')
    os.chdir(currentPath)
    weeklyGenerator(gitDataMapping)
    if menuIndex != '1':
        print('保存成功, 开始生成缓存文件...')
        tmpTxt = ''
        for inx, projectName in enumerate(gitFolderNames):
            if projectName != 'None':
                tmpTxt += gitFolderList[inx] + '&&&&&&&&&&&&&&&&&&' + projectName
                if inx != len(gitFolderNames) - 1:
                    tmpTxt += ':::::::'
        file = open(currentPath + os.path.sep + '.tmp.txt', mode='w')
        file.write(tmpTxt)
        file.close()
        print('缓存成功, 程序退出\n')
