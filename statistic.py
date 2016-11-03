# coding:utf

import docx2txt
import os
import pandas as pd
import re
import string
import pdb
import csv
import xlwt
import sys
import argparse

#pdb.set_trace()
def cvstran2xls(filename):
    with open(filename,'rb') as f:
        reader = csv.reader(f)
        cvscontent = []
        for row in reader:
            try:
              cvscontent.append([x.decode('utf-8') for x in row])
            except:
              cvscontent.append([x.decode('gbk') for x in row])
    file = xlwt.Workbook()
    table = file.add_sheet('sheet1',cell_overwrite_ok = True)
    #pdb.set_trace()
    for row in xrange(len(cvscontent)):
        for col in xrange(len(cvscontent[row])):
            table.write(row,col,cvscontent[row][col])
    file.save(os.path.splitext(filename)[0]+'.xls')
    print os.path.splitext(filename)[0]+'.xls','OK'



def printPath(level, path):
    global allFileNum
    '''
    打印一个目录下的所有文件夹和文件
    '''
    # 所有文件夹，第一个字段是次目录的级别
    dirList = []
    # 所有文件
    fileList = []
    # 返回一个列表，其中包含在目录条目的名称(google翻译)
    files = os.listdir(path)
    # 先添加目录级别
    dirList.append(str(level))
    for f in files:
        if(os.path.isdir(path + '/' + f)):
            # 排除隐藏文件夹。因为隐藏文件夹过多
            if(f[0] == '.'):
                continue
            else:
                # 添加非隐藏文件夹
                dirList.append(f)
        if(os.path.isfile(path + '/' + f)):
            # 添加文件
            if(f[0] == '.'):
                continue
            else:
                fileList.append(f)
    # 当一个标志使用，文件夹列表第一个级别不打印
    i_dl = 0
    for dl in dirList:
        if(i_dl == 0):
            i_dl = i_dl + 1
        else:
            # 打印至控制台，不是第一个的目录
            print '-' * (int(dirList[0])), dl
            # 打印目录下的所有文件夹和文件，目录级别+1
            printPath((int(dirList[0]) + 1), path + '/' + dl)
    return fileList


def cut2sentence(text):
    all_sentence = []
    for key in string.split(text, '\n'):
        if key != u'':
            all_sentence.append(key)
    return all_sentence


def analysis(path, group):

    file_list = printPath(1, path)
    result = []


    tc_pattern = [re.compile(u' [\u4e00-\u9fa5]+'), re.compile(r' \w+'), re.compile(r':\w+'), re.compile(u'：\w+'), re.compile(u'：[\u4e00-\u9fa5]+'), re.compile(u':[\u4e00-\u9fa5]+')]
    chinese_pattern = re.compile(u'[\u4e00-\u9fa5]{5}')

    for filename in file_list:
        if 'docx' in filename and '~$' not in filename:
            #print filename
            text = docx2txt.process(path+'/'+filename)
            all_paragraph = cut2sentence(text)

            # english subject
            english_subject = all_paragraph[0]

            # chinese subject
            chinese_subject = all_paragraph[1]

            # name
            translator_name = ''
            checker_name = ''

            for i in xrange(1, 10):
                paragraph = all_paragraph[i]
                first_word = paragraph[0]

                if first_word == u'翻' or first_word  == u'译':
                    c = ''
                    for k in xrange(0, len(tc_pattern)):
                        pattern = tc_pattern[k]
                        a = pattern.search(paragraph)
                        if a:
                            c = a.group()
                            break

                    for j in xrange(1, len(c)):
                        if c[j] != ' ':
                            translator_name += c[j]

                elif first_word == u'审' or first_word == u'校':
                    c = ''
                    for k in xrange(0, len(tc_pattern)):
                        pattern = tc_pattern[k]
                        a = pattern.search(paragraph)
                        if a:
                            c = a.group()
                            break

                    for j in xrange(1, len(c)):
                        if c[j] != ' ':
                            checker_name += c[j]

            # the number of words:

            num_chinese = 0
            num_english = 0
            for paragraph in all_paragraph:
                if chinese_pattern.search(paragraph):
                    #print paragraph
                    for words in paragraph:
                        if re.compile(u'[\u4e00-\u9fa5]').search(words) :
                            num_chinese += 1

                else:
                    '''for word in paragraph:
                        if word in string.punctuation:
                            num_english += 1'''
                    num_english += string.split(paragraph).__len__()
            #if group == 'ST':
            #    pdb.set_trace()
            if translator_name == u'M':
                translator_name = u'Meatle'
            if translator_name == u'bing':
                translator_name = u'Bing'
            if checker_name == u'M':
                checker_name = u'Meatle'
            if checker_name == u'Arc' or checker_name == u'Mahon' or checker_name==u'马小崩':
                checker_name = u'马宏'

            result.append([english_subject, chinese_subject, translator_name, checker_name, num_english, num_chinese, group])
    
    length = result.__len__()
    result = pd.DataFrame(result)
    result.set_axis(1, [u'英文题目', u'中文题目', u'翻译', u'审校', u'英文字数', u'中文字数', u'组别'])
    result.set_axis(0, range(1, length+1))

    group1 = result[[u'英文字数',u'中文字数']].groupby(result[u'翻译'])
    group2 = result[[u'英文字数',u'中文字数']].groupby(result[u'审校'])
    #pdb.set_trace()
    group1.sum().to_csv(group+'翻译量汇总.csv',encoding='utf')
    group2.sum().to_csv(group+'审校量汇总.csv',encoding='utf')
    result.to_csv(group+'统计结果.csv', encoding='utf')

    cvstran2xls(group+'翻译量汇总.csv')
    cvstran2xls(group+'审校量汇总.csv')
    cvstran2xls(group+'统计结果.csv')
'''
def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--SS",
            required=True, type=str,help="sixty-seconds path")
    parser.add_argument("--ST",
            required=True, type=str,help="science-talk path")
    return parser.parse_args()
'''

def main():
    #args = parse_args()
    #ss_path = args.SS
    #st_path = args.ST
    ss_path = sys.argv[1]
    analysis(ss_path, 'SS')
    #analysis(st_path, 'ST')


if __name__ == "__main__":
    main()




