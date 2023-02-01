import requests
import json
import xlwings

 
def analysisData(sht, schoolCode, i):
    geturl = 'https://rxyj.hzedu.gov.cn/hzjyAppServer/api/AppSchoolInfo/getSchoolInfo?year=2023&schoolName=' + schoolCode
    schoolres = requests.get(geturl)
    schooldata = json.loads(schoolres.text).get('result')
    sht.range((i, 1), (i, 10)).api.Borders.LineStyle = 1
    sht.range((i, 1), (i, 10)).api.HorizontalAlignment = -4131
    sht.range((i, 1), (i, 10)).api.VerticalAlignment = -4130
    sht.range('C'+str(i)).column_width = 10
    sht.range('F'+str(i)).column_width = 50
    sht.range('G'+str(i)).column_width = 25
    sht.range('I'+str(i)).column_width = 20

    sht.range('A'+str(i)).value = schooldata['appSchoolInfoEntity']['schoolName']
    sht.range('B'+str(i)).value = '公办' if schooldata['appSchoolInfoEntity']['gmblx'] == '非民办' else '民办'
    """sht.range('C'+str(i)).value=schooldata['appSchoolInfoEntity']['schoolDetail']"""
    sht.range('D'+str(i)).value = schooldata['appSchoolInfoEntity']['schoolTel']
    sht.range('E'+str(i)).value = schooldata['appSchoolInfoEntity']['address']
    areaData = schooldata.get('appSchoolDistrictInfoEntityList')
    areaStr = ""
    for areaEntity in areaData:
        areaStr += areaEntity['xqmc']
        areaStr += " "
        areaStr += areaEntity['buildingNumber'] if areaEntity['buildingNumber'] is not None else ""
        areaStr += " "
        areaStr += areaEntity['dw'] if areaEntity['dw'] is not None else ""
        areaStr += " "
        areaStr += areaEntity['bak2'] if areaEntity['bak2'] is not None else ""
        areaStr += ";"
    newHZRData = schooldata.get('appSchoolDistrictInfoEntityListNewHZR')
    for HZREntity in newHZRData:
        areaStr += HZREntity['xqmc']
        areaStr += " "
        areaStr += HZREntity['buildingNumber'] if HZREntity['buildingNumber'] is not None else ""
        areaStr += " "
        areaStr += HZREntity['dw'] if HZREntity['dw'] is not None else ""
        areaStr += " "
        areaStr += HZREntity['bak2'] if HZREntity['bak2'] is not None else ""
        areaStr += ";"

    areaValue = schooldata['appSchoolInfoEntity']['schoolScope'] if schooldata[
        'appSchoolInfoEntity']['schoolScope'] is not None else "。"
    areaValue += areaStr
    sht.range('F'+str(i)).value = areaValue
    lastYear = ""
    lastYear2 = ""
    lastYear3 = ""
    if (schooldata['appSchoolInfoEntity']['localShunt1'] == "1"):
        lastYear += "户籍生有分流；"
    if (schooldata['appSchoolInfoEntity']['localShunt1'] == "2"):
            lastYear += "户籍生无分流；"
    if (schooldata['appSchoolInfoEntity']['oneSuper1'] == "1"):
        lastYear += "一表生有分流 "
        if (schooldata['appSchoolInfoEntity']['oneSuperTime1'] is not None and len(schooldata['appSchoolInfoEntity']['oneSuperTime1']) > 0):
            lastYear += "最迟落户时间：" + schooldata['appSchoolInfoEntity']['oneSuperTime1']+";"
    if (schooldata['appSchoolInfoEntity']['oneSuper1'] == "2"):
        lastYear += "一表生无分流 "
    if (schooldata['appSchoolInfoEntity']['otherProvShunt1'] == "1"):
        lastYear += "随迁子女有分流；"
    if(schooldata['appSchoolInfoEntity']['localShunt1'] == "2" and schooldata['appSchoolInfoEntity']['oneSuper1'] =="2" and schooldata['appSchoolInfoEntity']['otherProvShunt1']=="2"):
        lastYear = "无分流；"

    if (schooldata['appSchoolInfoEntity']['localShunt2'] == "1"):
        lastYear2 += "户籍生有分流；"
    if (schooldata['appSchoolInfoEntity']['localShunt2'] == "2"):
        lastYear2 += "户籍生无分流；"
    if (schooldata['appSchoolInfoEntity']['oneSuper2'] == "1"):
        lastYear2 += "一表生有分流 "
        if (schooldata['appSchoolInfoEntity']['oneSuperTime2'] is not None and len(schooldata['appSchoolInfoEntity']['oneSuperTime2']) > 0):
            lastYear2 += "最迟落户时间：" + schooldata['appSchoolInfoEntity']['oneSuperTime2']+";"
    if (schooldata['appSchoolInfoEntity']['oneSuper2'] == "2"):
        lastYear2 += "一表生无分流 "
    if (schooldata['appSchoolInfoEntity']['otherProvShunt2'] == "1"):
        lastYear2 += "随迁子女有分流；"
    if(schooldata['appSchoolInfoEntity']['localShunt2'] == "2" and schooldata['appSchoolInfoEntity']['oneSuper2'] =="2" and schooldata['appSchoolInfoEntity']['otherProvShunt2']=="2"):
        lastYear2 = "无分流；"

    if (schooldata['appSchoolInfoEntity']['localShunt3'] == "1"):
        lastYear3 += "户籍生有分流；"
    if (schooldata['appSchoolInfoEntity']['localShunt3'] == "2"):
        lastYear3 += "户籍生无分流；"
    if (schooldata['appSchoolInfoEntity']['oneSuper3'] == "1"):
        lastYear3 += "一表生有分流 "
        if (schooldata['appSchoolInfoEntity']['oneSuperTime3'] is not None and len(schooldata['appSchoolInfoEntity']['oneSuperTime3']) > 0):
            lastYear3 += "最迟落户时间：" + schooldata['appSchoolInfoEntity']['oneSuperTime3']+";"
    if (schooldata['appSchoolInfoEntity']['oneSuper3'] == "2"):
        lastYear3 += "一表生无分流 "
    if (schooldata['appSchoolInfoEntity']['otherProvShunt3'] == "1"):
        lastYear3 += "随迁子女有分流；"
    if(schooldata['appSchoolInfoEntity']['localShunt3'] == "2" and schooldata['appSchoolInfoEntity']['oneSuper3'] =="2" and schooldata['appSchoolInfoEntity']['otherProvShunt3']=="2"):
        lastYear3 = "无分流；"

    sht.range('G'+str(i)).value = "2022年，"+lastYear + "2021年，"+lastYear2+"2020年，"+lastYear3
    sht.range('H'+str(i)).value = '红色预警' if schooldata['appSchoolInfoEntity']['wdIsshunt1'] == '3' else '橙色预警' if schooldata['appSchoolInfoEntity']['wdIsshunt1'] == '2' else '无预警'
    sht.range('I'+str(i)).value = schooldata['appSchoolInfoEntity']['schoolWay']
    sht.range('J'+str(i)).value = schooldata['appSchoolInfoEntity']['directMiddleSchoolName']
    return sht


areaCodeDict = {"330102": "上城区", "330105":"拱墅区","330106":"西湖区","330108":"滨江区","3301A5":"钱塘区","330109":"萧山区","330111":"富阳区","330122":"桐庐县","330127":"淳安县","330182":"建德市"}

excel = xlwings.App(visible=False, add_book=False)
wb = excel.books.add()

for area in areaCodeDict:
    jsondata = {
        "expressions": {"active": { "op": "eq", "value": "1"}, "schoolType": {"op": "in", "value": [ 2, 5, 6]},"schoolName": { "op": "lk","value": ""},
            "appXcFlag": {"op": "in", "value": ["0", "1"]}, "gmblx": {"op": "eq","value": ""},"szqdm": {"op": "lk","value": "330102"},
            "hideFlag": {"column": "hideFlag", "op": "eq", "value": "true"}
        }, "start": 1, "limit": 5000, "orderByExpressions": [{"column": "gmblxSort","orderByType": "asc"},{"column": "visitTimes","orderByType": "desc"}]
    }
    jsondata['expressions']['szqdm']['value'] = area
    areaName = areaCodeDict[area]
    sht = wb.sheets.add(areaName)
    i = 1
    sht.range((i, 1), (i, 10)).api.Font.Bold = True
    sht.range((i, 1), (i, 10)).api.Borders.LineStyle = 1

    sht.range('A'+str(i)).value = '学校名称'
    sht.range('B'+str(i)).value = '性质'
    sht.range('C'+str(i)).value = '介绍'
    sht.range('D'+str(i)).value = '电话'
    sht.range('E'+str(i)).value = '地址'
    sht.range('F'+str(i)).value = '学区范围'
    sht.range('G'+str(i)).value = '往年情况'
    sht.range('H'+str(i)).value = '23年入学预测'
    sht.range('I'+str(i)).value = '交通路线'
    sht.range('J'+str(i)).value = '以往对口中学'

    res = requests.post(
        'https://rxyj.hzedu.gov.cn/hzjyAppServer/api/AppSchoolInfo/custom/paginate/child', json=jsondata)
    schoolList = json.loads(res.text).get('result').get('records')
    i=2
    for school in schoolList:
        childs = school.get('appSchoolInfoEntityList')
        if(childs is not None):
            for child in childs:
                sht=analysisData(sht,child.get('xqbsm'),i)
                i = i+1
        else :
            sht=analysisData(sht,school.get('xqbsm'),i)
            i = i+1

wb.save(f"D:\ 2023年杭州小学入学情况.xlsx")
wb.close()
excel.quit()
