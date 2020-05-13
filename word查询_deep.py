# coding=utf-8
from win32com import client as wc
import os
import fnmatch
import re
import configparser
import threading
from collections import namedtuple
import shutil
import tkinter
import pythoncom
from tkinter import END
import math
from concurrent.futures import ThreadPoolExecutor, wait, ALL_COMPLETED



# 递归获取文件夹根目录下的所有文件
def search_file(file_root,all_files = []):
    files_base = os.listdir(file_root)
    for file in files_base:
        file_child = os.path.join(file_root,file)
        if os.path.isfile(file_child):
            all_files.append(file_child)
        elif os.path.isdir(file_child):
            # 对系统文件夹进行剔除
            if file_child.__contains__('System Volume') or file_child.__contains__('$RECYCLE.BIN'):
                continue
            all_files.append(search_file(file_child,all_files))
        else:
            continue
    return all_files

#去掉配置文件开头的BOM字节
def remove_BOM(config_path):
    with open(config_path, 'r',encoding = 'utf-8') as f:
        
        content = f.read()
        content = re.sub(r"\xfe\xff","", content)
        content = re.sub(r"\xff\xfe","", content)
        content = re.sub(r"\xef\xbb\xbf","", content)
        content = re.sub(r"\ufeff","", content)
    with open(config_path, 'w',encoding = 'utf-8') as f:
        f.write(content)


class Searcher(object):
 
    def __init__(self,win_cls):
        # self.config = configparser.ConfigParser()
        self.win_cls = win_cls
        project_path = os.path.dirname(os.path.abspath(__file__)) # 获取当前文件路径的上一级目录

        # file_path = project_path + config_file_path # 拼接路径字符串
        # remove_BOM(file_path)
        # self.config.read(file_path, encoding='utf-8')


        self.old_path = self.win_cls.information_from_form.searchpath
        self.key_word = self.win_cls.information_from_form.keyword
        # 该目下创建一个新目录tmp_dir，用来放转化后的txt文本, 日志缓存目录
        self.tmp_path = os.path.abspath(os.path.join(self.old_path, 'tmp_dir'))


        # --------------------------------------------------------------------
        # 定义namedtuple类，存储对真实文件地址、tmp中存储的txt文件名称的对应关系
        self.file_tuple = namedtuple('file_tuple', ['file_name', 'file_tmp'])
        # ---------------------------------------------------------------------
        self.file_tuple_list = []
  
        # 可以使用多线程加速对查找速度进行优化, 此处由于时间限制不再进行
        self.threads = []       # 用来处理python多线程的容器
        self.process_list = []  # 此目录用于多线程处理对搜索引擎进行加速
 
        # 读取配置文件
        # 存在缓存时是否每次都重新加载目录文件内容
        # self.win_cls.log_data_Text.insert(tkinter.END,"[开始读取配置文件]")
        if (self.win_cls.information_from_form.alwaysload == 1):
            self.win_cls.log_data_Text.insert(tkinter.END,"[持续加载开关]【开启】"+ '\n')
            self.win_cls.log_data_Text.see(END)
            self.win_cls.log_data_Text.update()
            self.always_load = True
        else:
            self.win_cls.log_data_Text.insert(tkinter.END,"[持续加载开关]【关闭】"+ '\n')
            self.win_cls.log_data_Text.see(END)
            self.win_cls.log_data_Text.update()
            self.always_load = False
        # 是否输出详细开发日志
        if (self.win_cls.information_from_form.debugmode == 1):
            self.win_cls.log_data_Text.insert(tkinter.END,"[调试模式开关]【开启】"+ '\n')
            self.win_cls.log_data_Text.see(END)
            self.win_cls.log_data_Text.update()
            self.debug_mode = True
        else:
            self.win_cls.log_data_Text.insert(tkinter.END,"[调试模式开关]【关闭】"+ '\n')
            self.win_cls.log_data_Text.see(END)
            self.win_cls.log_data_Text.update()
            self.debug_mode = False
        # 是否开启联想查询模式
        if (self.win_cls.information_from_form.smartmode == 1):
            self.win_cls.log_data_Text.insert(tkinter.END,"[智能模式开关]【开启】"+ '\n')
            self.win_cls.log_data_Text.see(END)
            self.win_cls.log_data_Text.update()
            self.smart_mode = True
        else:
            self.win_cls.log_data_Text.insert(tkinter.END,"[智能模式开关]【关闭】"+ '\n')
            self.win_cls.log_data_Text.see(END)
            self.win_cls.log_data_Text.update()
            self.smart_mode = False
        # 若为智能模式, 则获取最短智能子串长度
        if(self.win_cls.information_from_form.smartmode == 1):
            try:
                if(self.win_cls.information_from_window.childlength is not int):
                
                    self.win_cls.log_data_Text.insert('未设置智能子串长度，采用默认值3'+ '\n')
                    self.win_cls.log_data_Text.see(END)
                    self.win_cls.log_data_Text.update()
            except:
                self.child_length = 3   # 若未设置, 则最短子串长度为3
                self.win_cls.log_data_Text.insert(tkinter.END,"[智能子串长度]【3】"+ '\n')
                self.win_cls.log_data_Text.see(END)
                self.win_cls.log_data_Text.update()
 
    def Translate(self):
        '''
        将一个目录下所有doc和docx文件转成txt
        该目录下创建一个新目录newdir
        新目录下fileNames.txt创建一个文本存入所有的word文件名
        本版本具有一定的容错性，即允许对同一文件夹多次操作而不发生冲突
        '''
        # 该目录下所有文件的名字
        # files_base = os.listdir(self.old_path)
        
        # files_base = os.listdir(self.old_path)
        # for file in files_base:
        #     if 

        # 遍历文件夹下所有文件
        filea_all = []
        files = search_file(self.old_path,filea_all)

        if not os.path.exists(self.tmp_path):
            os.mkdir(self.tmp_path)
        else:   # 目录存在, 对tmp目录中的文件数量进行判断, 若与目标不同则仍需要重新加载
            if(not self.always_load):
                self.win_cls.log_data_Text.insert(tkinter.END,"[基础极速加载完毕 ...]" + '\n')
                self.win_cls.log_data_Text.see(END)
                self.win_cls.log_data_Text.update()
                return
        self.win_cls.log_data_Text.insert(tkinter.END,"[基础数据解析准备中 ...]"+ '\n')
        self.win_cls.log_data_Text.see(END)
        self.win_cls.log_data_Text.update()

 
        # for filename in files:
        #     try:
        #         # 如果不是word文件：继续
        #         if not fnmatch.fnmatch(filename, '*.doc') and not fnmatch.fnmatch(filename, '*.docx'):
        #             continue
        #         # 如果是word临时文件：继续
        #         if fnmatch.fnmatch(filename, '~$*'):
        #             continue
        #         self.process_list.append(filename)
        #         self.Process()      # 执行doc到txt转换过程, 但是由于其调用了微软office api接口, 因此无法用多线程进行加速
        #     except Exception as e:
        #         print(e)
        #         pass

        for filename in files:
            if not isinstance(filename,str):
                continue
            # 如果不是word文件：继续
            if not fnmatch.fnmatch(filename, '*.doc') and not fnmatch.fnmatch(filename, '*.docx'):
                continue
            # 如果是word临时文件：继续
            if fnmatch.fnmatch(filename, '~$*'):
                continue
            self.process_list.append(filename)
            self.Process()      # 执行doc到txt转换过程, 但是由于其调用了微软office api接口, 因此无法用多线程进行加速


        # 启用线程池，实现对多数据的并行处理

        # length = len(files)
        # n = 5
        # # with ThreadPoolExecutor(max_workers=20) as pool:
        # #     for i in range(length):
        # #         files_list = files[math.floor(i / n * length):math.floor((i + 1) / n * length)]    
             
        # #         pool.submit(self.thread_process, files_list)
        # files_list_save = []
        # executor = ThreadPoolExecutor(max_workers=5)
        # for i in range(length):
        #         files_list = files[math.floor(i / n * length):math.floor((i + 1) / n * length)]    
        #         files_list_save.append(files_list)

        # all_task = [executor.submit(self.thread_process, (files_list_one)) for files_list_one in files_list_save]

        


        # with ThreadPoolExecutor(max_workers=20) as pool:
        #     for i in range(length):
        #         files_list = files[math.floor(i / n * length):math.floor((i + 1) / n * length)]    
             
        #         pool.submit(self.thread_process, files_list)

        

            
             

        # self.win_cls.log_data_Text.insert(tkinter.END,"[基础解析加载完毕 ...]"+ '\n')
        # self.win_cls.log_data_Text.see(END)
        # self.win_cls.log_data_Text.update()     


        # wait(all_task)
        shutil.rmtree(self.tmp_path)
        self.win_cls.result_data_Text.insert(tkinter.END, '【------------------------------------------------搜索过程结束-------------------------------------------------】' + '\n')




    # 开启多线程，实现对文件的数据的并行处理
    def thread_process(self,files):
        for filename in files:
            if not isinstance(filename,str):
                continue
            # 如果不是word文件：继续
            if not fnmatch.fnmatch(filename, '*.doc') and not fnmatch.fnmatch(filename, '*.docx'):
                continue
            # 如果是word临时文件：继续
            if fnmatch.fnmatch(filename, '~$*'):
                continue
            self.process_list.append(filename)
            self.Process()      # 执行doc到txt转换过程, 但是由于其调用了微软office api接口, 因此无法用多线程进行加速




    def Process(self):
        '''
        子进程处理程序, 多进程齐开，对合约文件进行快速处理
        :return: 
        '''

 
        if(len(self.process_list) != 0):
            file_name = self.process_list[0]
            self.process_list.remove(file_name)
        else:
            return
        docpath = os.path.abspath(os.path.join(self.old_path, file_name))
        if (self.debug_mode):
            self.win_cls.log_data_Text.insert(tkinter.END,"Dealing office file: " + docpath + '\n')
            self.win_cls.log_data_Text.see(END)
            self.win_cls.log_data_Text.update()
 
        # 得到一个新的文件名,把原文件名的后缀改成txt
        if fnmatch.fnmatch(file_name, '*.doc'):
            new_txt_name = file_name.split('\\')[-1][:-4] + '.txt'
        else:
            new_txt_name = file_name.split('\\')[-1][:-5] + '.txt'
        word_to_txt = os.path.join(os.path.join(self.old_path, 'tmp_dir'), new_txt_name)
        
        # self.file_tuple_list.append(file_tuple_new)
        try:
            # Windows系统虽然未装Microsoft Word, python中依然可以使用其模块
            pythoncom.CoInitialize()

            wordapp = wc.DispatchEx('Word.Application')
            wordapp.Visible = 0
            wordapp.DisplayAlerts = 0
            doc = wordapp.Documents.Open(docpath)
            # 为了让python可以在后续操作中r方式读取txt和不产生乱码，参数为4
            doc.SaveAs(word_to_txt, 4)
            doc.Close()

            # pythoncom.CoUninitialize()
            
            file_tuple_new = self.file_tuple(file_name,word_to_txt)

            self.Search_one_by_one(file_tuple_new)


            # self.file_tuple_list.append(file_tuple_new)
            if (self.debug_mode):
                self.win_cls.log_data_Text.insert(tkinter.END,"Finish Dealing file: " + docpath + '\n')
                self.win_cls.log_data_Text.see(END)
                self.win_cls.log_data_Text.update()

        except Exception as e:
            print(e)
            self.win_cls.log_data_Text.insert(tkinter.END,"Warmming: Rollback file " + file_name + '\n')
            self.win_cls.log_data_Text.see(END)
            self.win_cls.log_data_Text.update()
            self.process_list.append(file_name)
        finally:
            wordapp.Quit()
            pythoncom.CoUninitialize()
        


 
    # def Search(self):
    #     '''
    #     对记录的文本进行查询处理, 使用多线程进行加速快查, 看是否存在所需要的key值m
    #     :param key: 查询键值
    #     :return: 
    #     '''
    #     key = self.key_word
    #     sum = 0
    #     # files = []
    #     # files = os.listdir(self.tmp_path)
    #     # for i in range(len(self.file_tuple_list)):
    #     #     files.append(self.file_tuple_list[i].file_tmp)
    #     content = ''    # 用来存放整个文本的文字内容
    #     for file_tuple in self.file_tuple_list:
    #         # filePath = os.path.join(self.tmp_path, file)
    #         try:
    #             with open(file_tuple.file_tmp, 'r') as fr:
    #                 lines = fr.readlines()
    #                 for line in lines:
    #                     line = line.strip('\n').strip().strip('\t')
    #                     if(len(line) == 0):
    #                         continue
    #                     content = content + str(line)
    #                 if(True == self.RegexProA(line=content, key=key, file_tuple=file_tuple)):
    #                     sum = sum + 1
    #                 content = ''
    #         except Exception as e:
    #             print(e)
    #     self.win_cls.log_data_Text.insert(tkinter.END,"[获得搜索结果总数]【" + str(sum) + "】" + '\n')
    #     shutil.rmtree(self.tmp_path)
 

    def Search_one_by_one(self,file_tuple):
        '''
        对记录的文本进行查询处理, 使用多线程进行加速快查, 看是否存在所需要的key值m
        :param key: 查询键值
        :return: 
        '''
        key = self.key_word
        sum = 0
        # files = []
        # files = os.listdir(self.tmp_path)
        # for i in range(len(self.file_tuple_list)):
        #     files.append(self.file_tuple_list[i].file_tmp)
        content = ''    # 用来存放整个文本的文字内容
        # for file_tuple in self.file_tuple_list:
            # filePath = os.path.join(self.tmp_path, file)
        try:
            with open(file_tuple.file_tmp, 'r') as fr:
                lines = fr.readlines()
                for line in lines:
                    line = line.strip('\n').strip().strip('\t')
                    if(len(line) == 0):
                        continue
                    content = content + str(line)
                if(True == self.RegexProA(line=content, key=key, file_tuple=file_tuple)):
                    sum = sum + 1
                content = ''
        except Exception as e:
            print(e)
    # self.win_cls.log_data_Text.insert(tkinter.END,"[获得搜索结果总数]【" + str(sum) + "】" + '\n')
    
 
    def RegexProA(self, line, key, file_tuple):
        '''
        对line值进行正则处理, 判断其是否为有效行,此步逻辑优化可以使功能更加强大
        :param line: 正则处理行, 对所搜索的内容进行简单判断
        :return: 
        '''
        rate = 0
        if(not self.smart_mode):
            links = re.findall(str(key),line, re.IGNORECASE)  # 搜索时忽视大小写
            rate = len(links)
            if(rate != 0):
                # self.win_cls.result_data_Text.insert(tkinter.END,"[" + str(file_tuple.file_name) + "] Find [" + str(key) + "] for [" + str(rate) + "] times ..."+ '\n')
                self.win_cls.result_data_Text.insert(tkinter.END, str(file_tuple.file_name) +'\n')


                return True
            else:
                return False
        else:
            head = 1
            list = self.child(key)
            list.reverse()
            links = re.findall(str(key),line, re.IGNORECASE)  # 搜索时忽视大小写
            rate = len(links)
            if(rate != 0):
                # self.win_cls.result_data_Text.insert(tkinter.END,"[" + str(file_tuple.file_name) + "] Find [" + str(key) + "] for [" + str(rate) + "] times ..."+ '\n')
                self.win_cls.result_data_Text.insert(tkinter.END, str(file_tuple.file_name) + '\n')

                return  True # 当整串都能被搜索到时, 搜索子串是没有必要的
            for element in list:
                head = head + 1
                if(len(element) >= int(self.child_length) and len(element) < len(key)):
                    links = re.findall(str(element), line, re.IGNORECASE)  # 搜索时忽视大小写
                    rate = len(links)
                    if (rate != 0):
                        # self.win_cls.result_data_Text.insert(tkinter.END,"[" + str(file_tuple.file_name) + "] Find [" + str(element) + "] for [" + str(rate) + "] times ..."+ '\n')
                        self.win_cls.result_data_Text.insert(tkinter.END, str(file_tuple.file_name) + '\n')

                        return  True # 当搜索到长串时, 不再搜索短串
            return False
 
 
    def child(self, s=''):
        '''
        输出字符串s的所有排列组合
        :param s: 待处理字符串
        :return: 所有子串容器
        '''
        results = []
        # x + 1 表示子字符串长度
        for x in range(len(s)):
            # i 表示偏移量
            for i in range(len(s) - x):
                results.append(s[i:i + x + 1])
        return results













