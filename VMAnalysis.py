# coding=utf-8
import os

import xlrd
import xlwt
import re
import logging

# 创建list，存放xls数据
data_list1 = []
data_list2 = []
dataIDNotSame = []
dataCommonData = []
dataVMBalance = []

logging.basicConfig(filename='mylog.txt', format="%(asctime)s : %(message)s",
                    level=logging.DEBUG)


# 解析xls文件到list，用于后续数据处理数据源
def XLSRead(XLSPath):
    logging.info('begin to read file : %s', XLSPath)
    # 打开一个workbook
    workbook = xlrd.open_workbook(XLSPath)
    # 抓取所有sheet页的名称
    worksheets = workbook.sheet_names()
    # 定位到目标sheet
    worksheet = workbook.sheet_by_name(u'虚拟机列表')
    # 获取该sheet中的有效行数
    num_rows = worksheet.nrows
    # 获取列表的有效列数
    num_cols = worksheet.ncols
    # 定义list，用于存放xls数据
    dataList = []
    logging.info('begin analytical xls data.')

    # 解析第一个xls
    for rown in range(1, num_rows):
        data_dict = {}
        for coln in range(num_cols):
            if worksheet.cell_value(0, coln) == '虚拟机名称':
                data_dict['虚拟机名称'] = str(worksheet.cell_value(rown, coln)).strip()
            if worksheet.cell_value(0, coln) == '虚拟机ID':
                data_dict['虚拟机ID'] = str(worksheet.cell_value(rown, coln)).strip()
            if worksheet.cell_value(0, coln) == '状态':
                data_dict['状态'] = str(worksheet.cell_value(rown, coln)).strip()
            if worksheet.cell_value(0, coln) == '所属网元':
                data_dict['所属网元'] = str(worksheet.cell_value(rown, coln)).strip()
            if worksheet.cell_value(0, coln) == '服务器名称':
                data_dict['服务器名称'] = str(worksheet.cell_value(rown, coln)).strip()
            if worksheet.cell_value(0, coln) == '所属主机':
                data_dict['所属主机'] = str(worksheet.cell_value(rown, coln)).strip()
            if worksheet.cell_value(0, coln) == '所属主机ID':
                data_dict['所属主机ID'] = str(worksheet.cell_value(rown, coln)).strip()
            if worksheet.cell_value(0, coln) == 'CPU使用率':
                data_dict['CPU使用率'] = str(worksheet.cell_value(rown, coln)).strip()
            if worksheet.cell_value(0, coln) == '内存使用率':
                data_dict['内存使用率'] = str(worksheet.cell_value(rown, coln)).strip()
            if worksheet.cell_value(0, coln) == '磁盘使用率':
                data_dict['磁盘使用率'] = str(worksheet.cell_value(rown, coln)).strip()
        dataList.append(data_dict)
        # cell = worksheet.cell_value(rown, coln)
        # print('rown%s is %s' % (rown, data_list))
    logging.info('analytical xls data already complete')
    return dataList


# 设置单元格式 入参type (1:表头第一列样式  2:某一单元格样式)
def SetFont(type):
    style = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    borders = xlwt.Borders()

    # 设置边框
    # borders.left = 1
    # borders.right = 1
    # borders.top = 1
    # borders.bottom = 1
    # borders.bottom_colour = 0x3A
    # style.borders = borders

    if type == 1:
        # 设置字体格式
        Font = xlwt.Font()
        Font.name = "Times New Roman"
        Font.bold = True  # 加粗
        style.font = Font
    elif type == 2:
        # 设置单元格背景色
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['yellow']
        style.pattern = pattern
    return style


# 解析数据，获取对应结果
def XLSAnalysis(data_list1, data_list2):
    logging.info('get result begin ')
    # 1.判断虚拟机id，所属主机不一致情况，不一致，输出到’虚拟机所属主机不一致‘页签
    for dic1 in data_list1:
        for dic2 in data_list2:
            if 'UDM' in dic1['虚拟机名称']:
                if dic1['虚拟机ID'] == dic2['虚拟机ID'] and dic1['所属主机ID'] != dic2['所属主机ID']:
                    dictIDNotSame = {}
                    dictIDNotSame['虚拟机名称'] = dic1['虚拟机名称']
                    dictIDNotSame['虚拟机ID'] = dic1['虚拟机ID']
                    dictIDNotSame['所属主机ID1'] = dic1['所属主机ID']
                    dictIDNotSame['所属主机ID2'] = dic2['所属主机ID']
                    dictIDNotSame['所属主机1'] = dic1['所属主机']
                    dictIDNotSame['所属主机2'] = dic2['所属主机']
                    dataIDNotSame.append(dictIDNotSame)

    # 2.常规指标检测，超出对应阈值输出到’常规指标‘页签
    for dic1 in data_list1:
        if dic1['CPU使用率'] != 'N/A' and 'UDM' in dic1['虚拟机名称']:
            if dic1['状态'] != '正常' or float(dic1['CPU使用率']) > 80 or float(dic1['内存使用率']) > 80 or float(
                    dic1['磁盘使用率']) > 80:
                dictCommonData = {}
                dictCommonData['虚拟机名称'] = dic1['虚拟机名称']
                dictCommonData['虚拟机ID'] = dic1['虚拟机ID']
                dictCommonData['状态'] = dic1['状态']
                dictCommonData['CPU使用率'] = dic1['CPU使用率']
                dictCommonData['内存使用率'] = dic1['内存使用率']
                dictCommonData['磁盘使用率'] = dic1['磁盘使用率']
                dataCommonData.append(dictCommonData)

    # 3. 虚拟机均衡，相同类型虚拟机不能在同一机柜，如果存在则输出
    dataVMBalance.extend(VMBalance(data_list1))
    logging.info('get result end ')
    # for i in dataVMBalance:
    #     print(i)


'''
虚拟机均衡，相同类型虚拟机不能在同一机柜，如果存在则输出
规则：
1.VDU1_1和 VDU2_1属于一类，他两不能在同一机柜   机柜判断方法：判断截取的R03C04不能相同，如果相同，输出结果结果
2.末尾结构是2_*的，判断截取的'R02C03U24'不能相同
3如果末尾只是一种数字的，判断截取的R03C04不能相同，如果相同，输出结果结果
'''


def VMBalance(data_list1):
    logging.info('vmbalance begin')
    res = []
    for i in range(len(data_list1)):
        for j in range(i + 1, len(data_list1)):
            if 'UDM' in data_list1[i]['虚拟机名称'] and 'UDM' in data_list1[j]['虚拟机名称']:
                lvmf = data_list1[i]['虚拟机名称']
                lvms = data_list1[j]['虚拟机名称']
                lhf = data_list1[i]['所属主机']
                lhs = data_list1[j]['所属主机']
                ltmp = VMCmp(lvmf, lvms, lhf, lhs)
                if len(res):
                    res.extend(ltmp)
                else:
                    res = ltmp
    # for i in distinct2(res):
    #     print(i)
    logging.info('vmbalance end')
    return distinct2(res)


# 去重
def distinct2(items):
    exist_questions = set()
    result = []
    for item in items:
        question = item['虚拟机名称']
        if question not in exist_questions:
            exist_questions.add(question)
            result.append(item)
    return result


# 是否以数字结尾
def end_num(string):
    # 以一个数字结尾字符串
    text = re.compile(r".*[0-9]$")
    if text.match(string):
        return True
    else:
        return False


def VMCmp(lvmf, lvms, lhf, lhs):
    # 按'_'分割
    lvm1 = lvmf.split('_')
    lvm2 = lvms.split('_')
    lh1 = lhf.split('-')
    lh2 = lhs.split('-')
    # 存放结果
    lres = []
    # 截取所属主机对应位置
    lh1 = lh1[4].split('U')[0]
    lh2 = lh2[4].split('U')[0]

    # 规则一：按‘_’分割，如果取到的最后两个‘_’前面的数据一样，认为同类型
    lvmfront1 = ''.join(lvm1[:len(lvm1) - 2])
    lvmfront2 = ''.join(lvm2[:len(lvm2) - 2])
    if lvmfront1 == lvmfront2 and lh1 == lh2:
        dict1 = {}
        dict2 = {}
        dict1['虚拟机名称'] = lvmf
        dict1['所属主机'] = lhf
        lres.append(dict1)
        dict2['虚拟机名称'] = lvms
        dict2['所属主机'] = lhs
        lres.append(dict2)

    # 规则一：按’_'分割，如果倒数第三个有中划线，取除中划线前面数据，拼接上倒数两个'_'的数据比较，如果一样，认为同类型
    l1 = lvm1[len(lvm1) - 3]
    l2 = lvm2[len(lvm2) - 3]
    if '-' in l1 and '-' in l2:
        lvmfront1 = ''.join(lvm1[:len(lvm1) - 3] + l1.split('-')[0:1] + lvm1[len(lvm1) - 2:])
        lvmfront2 = ''.join(lvm2[:len(lvm1) - 3] + l2.split('-')[0:1] + lvm2[len(lvm2) - 2:])
        # lvmfront2 = re.split('\d+$', ''.join(lvm1[:len(lvm2) - 2]))
        if lvmfront1 == lvmfront2 and lh1 == lh2:
            dict1 = {}
            dict2 = {}
            dict1['虚拟机名称'] = lvmf
            dict1['所属主机'] = lhf
            lres.append(dict1)
            dict2['虚拟机名称'] = lvms
            dict2['所属主机'] = lhs
            lres.append(dict2)
    return lres


def XLSWrite(XLSPath):
    # 写数据时，行计数器
    logging.info('xls write begin')
    shtNum1 = 1
    shtNum2 = 1
    shtNum3 = 1
    # 实例化一个execl对象xls=工作薄
    xls = xlwt.Workbook()
    # 实例化一个工作表，名叫Sheet1
    sht1 = xls.add_sheet('虚拟机迁移检测')
    sht2 = xls.add_sheet('常规指标检测')
    sht3 = xls.add_sheet('虚拟机反亲和性检测')
    # 第一个参数是行，第二个参数是列，第三个参数是值,第四个参数是格式
    sht1.write(0, 0, '虚拟机名称', SetFont(1))
    sht1.write(0, 1, '虚拟机ID', SetFont(1))
    sht1.write(0, 2, '所属主机ID-new', SetFont(1))
    sht1.write(0, 3, '所属主机ID-old', SetFont(1))
    sht1.write(0, 4, '所属主机-new', SetFont(1))
    sht1.write(0, 5, '所属主机-old', SetFont(1))

    sht2.write(0, 0, '虚拟机名称', SetFont(1))
    sht2.write(0, 1, '虚拟机ID', SetFont(1))
    sht2.write(0, 2, '状态', SetFont(1))
    sht2.write(0, 3, 'CPU使用率', SetFont(1))
    sht2.write(0, 4, '内存使用率', SetFont(1))
    sht2.write(0, 5, '磁盘使用率', SetFont(1))

    sht3.write(0, 0, '虚拟机名称', SetFont(1))
    sht3.write(0, 1, '所属主机', SetFont(1))

    # 数据写入
    # sheet1
    if len(dataIDNotSame):
        for dic1 in dataIDNotSame:
            # print(dic1)
            sht1.write(shtNum1, 0, dic1['虚拟机名称'])
            sht1.write(shtNum1, 1, dic1['虚拟机ID'])
            sht1.write(shtNum1, 2, dic1['所属主机ID1'])
            sht1.write(shtNum1, 3, dic1['所属主机ID2'])
            sht1.write(shtNum1, 4, dic1['所属主机1'])
            sht1.write(shtNum1, 5, dic1['所属主机2'])
            shtNum1 = shtNum1 + 1
    else:
        sht1.write_merge(shtNum1, shtNum1, 0, 5, '检测后，当前无虚拟机迁移相关数据', SetFont(2))

    # sheet2
    if len(dataCommonData):
        for dic2 in dataCommonData:
            # print(dic2)
            sht2.write(shtNum2, 0, dic2['虚拟机名称'])
            sht2.write(shtNum2, 1, dic2['虚拟机ID'])
            if dic2['状态'] != '正常':
                sht2.write(shtNum2, 2, dic2['状态'], SetFont(2))
            else:
                sht2.write(shtNum2, 2, dic2['状态'])
            if float(dic2['CPU使用率']) > 80:
                sht2.write(shtNum2, 3, dic2['CPU使用率'], SetFont(2))
            else:
                sht2.write(shtNum2, 3, dic2['CPU使用率'])
            if float(dic2['内存使用率']) > 80:
                sht2.write(shtNum2, 4, dic2['内存使用率'], SetFont(2))
            else:
                sht2.write(shtNum2, 4, dic2['内存使用率'])
            if float(dic2['磁盘使用率']) > 80:
                sht2.write(shtNum2, 5, dic2['磁盘使用率'], SetFont(2))
            else:
                sht2.write(shtNum2, 5, dic2['磁盘使用率'])
            shtNum2 = shtNum2 + 1
    else:
        sht2.write_merge(shtNum2, shtNum2, 0, 5, '检测后，当前无常规指标相关数据', SetFont(2))

    # sheet3
    if len(dataVMBalance):
        for dic3 in dataVMBalance:
            sht3.write(shtNum3, 0, dic3['虚拟机名称'])
            sht3.write(shtNum3, 1, dic3['所属主机'])
            shtNum3 = shtNum3 + 1
    else:
        sht3.write_merge(shtNum3, shtNum3, 0, 1, '检测后，当前无虚拟机反亲和性相关数据', SetFont(2))
    xls.save(XLSPath)
    logging.info('xls write end')


def main():
    # 解析xls文件到list，用于后续数据处理数据源
    logging.info('welcome to xls world.')
    try:
        data_list1 = XLSRead(os.getcwd() + '\\虚拟机列表new.xls')
        data_list2 = XLSRead(os.getcwd() + '\\虚拟机列表old.xls')
        # 解析数据，获取对应结果
        XLSAnalysis(data_list1, data_list2)
        # 结果输出
        XLSWrite(os.getcwd() + '\\result.xls')
    except Exception as err:
        logging.error(err)

    logging.info("end xls world")


if __name__ == '__main__':
    main()
