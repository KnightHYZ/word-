# 窗口界面类，实现对窗口界面的实例化
from tkinter import Label, scrolledtext, StringVar, HORIZONTAL, PanedWindow, Entry, Text, IntVar, INSERT, Tk,N, S, W, E, NSEW, Checkbutton, Button, Listbox, MULTIPLE, Event, END
from word查询_deep import *
import tkinter.messagebox
from threading import Thread
from time import sleep

class MY_GUI():
    def __init__(self,init_window_name):
        self.init_window_name = init_window_name

    # 设置界面
    def set_init_window(self):

        self.init_window_name.title('word文档查询工具_v1.0')
        # self.init_window_name.geometry('1068x681+10+10')
        # self.init_window_name.columnconfigure (0, weight=1)
        self.init_window_name.resizable(0,0)

        # -------------------搜索目录选择设置----------------------------------
        


        self.init_data_label = Label(self.init_window_name,text = '搜索目录:')
        self.init_data_label.grid(row=0, column=0,sticky=W)

        v = StringVar(self.init_window_name, value='D:\\')
        # self.init_data_entry = Entry(self.init_window_name, width = 62, textvariable = v)
        # self.init_data_entry = Entry(self.init_window_name, textvariable = v, width=self.init_window_name.winfo_screenwidth())
        self.init_data_entry = Entry(self.init_window_name, textvariable = v)
        # self.init_data_Text.grid(row=0, column=1, rowspan=10, columnspan=10)
        self.init_data_entry.grid(row=0, column=1, columnspan=9,sticky=NSEW)

        
        # -------------------搜索关键字选择设置----------------------------------



        self.init_key_label = Label(self.init_window_name,text = '搜索关键字:')
        self.init_key_label.grid(row=1, column=0,sticky=W)
               
        # self.init_key_entry = Entry(self.init_window_name, width = 62)
        self.init_key_entry = Entry(self.init_window_name)
        # self.init_data_Text.grid(row=0, column=1, rowspan=10, columnspan=10)
        self.init_key_entry.grid(row=1, column=1, columnspan=9,sticky=NSEW)
        

        # -------------------搜索结果框表设置----------------------------------
        self.result_data_label = Label(self.init_window_name, text="搜索结果:")
        self.result_data_label.grid(row=3, column=0,sticky=W)
        # self.result_data_Text = Listbox(self.init_window_name, width=75, height=20, selectmode=MULTIPLE)  #处理结果展示
        self.result_data_Text = Listbox(self.init_window_name, height=20)  #处理结果展示
        # self.result_data_Text = Listbox(self.init_window_name, width=75, height=20)  #处理结果展示
        # self.result_data_Text.grid(row=2, column=12, rowspan=15, columnspan=10)
        self.result_data_Text.grid(row=4, columnspan=10,sticky=NSEW)
        
        # 为listbox双击绑定事件
        self.result_data_Text.bind("<Button-2>",self.openword)
        # self.result_data_Text.bind("  <<ListboxSelect>>",self.openword)
      
        # -------------------搜索日志框表设置----------------------------------
        self.log_label = Label(self.init_window_name, text="搜索日志")
        self.log_label.grid(row=5, column=0,sticky=W)
        # self.log_data_Text = Text(self.init_window_name, width=75, height=9)  # 日志框
        self.log_data_Text = scrolledtext.ScrolledText(self.init_window_name, height=9)  # 日志框
        # self.log_data_Text.grid(row=4, column=0, columnspan=10)
        self.log_data_Text.grid(row=6, columnspan=10,sticky=NSEW)


        # -------------------搜索选择项框表设置----------------------------------


        self.v_reload = IntVar()
        self.reload_set = Checkbutton(self.init_window_name, text = '重加载选项',variable = self.v_reload)
        self.reload_set.grid(row=2,column=0,sticky=NSEW)
        self.reload_set.select()
        
        self.v_debug_set = IntVar()
        self.debug_set = Checkbutton(self.init_window_name, text = '调试模式',variable = self.v_debug_set)
        self.debug_set.grid(row=2,column=1, columnspan=1,sticky=NSEW)
        self.debug_set.select()
        

        self.v_smart_mode_set = IntVar()
        self.smart_mode_set = Checkbutton(self.init_window_name, text = '智能联想模式',variable = self.v_smart_mode_set)
        self.smart_mode_set.grid(row=2,column=2,sticky=NSEW)


        self.smart_child_length_set = Label(self.init_window_name,text = '智能联想最短字符串:')
        self.smart_child_length_set.grid(row=2, column=3,sticky=NSEW)             
        # self.smart_child_length_set_entry = Entry(self.init_window_name, width = 20)
        self.smart_child_length_set_entry = Entry(self.init_window_name)
        # self.init_data_Text.grid(row=0, column=1, rowspan=10, columnspan=10)
        self.smart_child_length_set_entry.grid(row=2, column=4, columnspan=5,sticky=NSEW)
        

        self.search_start_button = Button(self.init_window_name, text = '开始检索', command = self.thread_it)
        self.search_start_button.grid(row=2,column=9,sticky=NSEW)
        
        
        # self.init_window_name.columnconfigure (0, weight=1)
    
    # 单击搜寻到的结果后打开文件,在此处添加打开文件的函数
    def openword(self, event):
        
        index_select = self.result_data_Text.curselection()[-1]
        # 此处有bug，需要实时更新双击项目情况
        value = self.result_data_Text.get(index_select)
        # print(index_select)
        # print(value)
        # self.result_data_Text.curselection().clear()
        wordapp = wc.Dispatch('Word.Application')
        doc = wordapp.Documents.Open(value.strip())
    
    # 设置线程，避免界面卡死情况
    def thread_it(self,*args):
        

        t = Thread(target=self.button_Click)

        t.setDaemon(True) 
        # 启动
        t.start()
        # 阻塞--卡死界面！
        # t.join()





    def button_Click(self):
        sleep(0.1)
        self.result_data_Text.delete(0,END)
        self.log_data_Text.delete('1.0','end')


        self.information_from_form = namedtuple('information_from_form', ['alwaysload', 'debugmode', 'smartmode', 'childlength', 'searchpath', 'keyword'])
        # 获取搜索输入的目录信息
        self.information_from_form.searchpath = self.init_data_entry.get()
        # 获取搜索输入的检索关键字信息
        self.information_from_form.keyword = self.init_key_entry.get()
        # 获取搜索输入的是否重加载复选框信息
        self.information_from_form.alwaysload = self.v_reload.get()
        # 获取搜索输入的是否开启调试模式信息
        self.information_from_form.debugmode = self.v_debug_set.get()
        # 获取搜索输入的是否启动智能搜索模式信息
        self.information_from_form.smartmode = self.v_smart_mode_set.get()
        # 获取搜索输入的子串搜索长度信息
        self.information_from_form.childlength = self.smart_child_length_set_entry.get()
        
        # 获取窗口内容，并进行搜索
        instance = Searcher(self)   # 若Config.ini中未进行配置, 则采用默认搜索值

        instance.Translate()
        


        # instance.Search()





if __name__ == '__main__':
    # 建立界面窗口

    init_window = Tk()
    Word_search = MY_GUI(init_window)
    Word_search.set_init_window()
     
    init_window.mainloop()