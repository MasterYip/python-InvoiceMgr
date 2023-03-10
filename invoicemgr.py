'''
Author: MasterYip 2205929492@qq.com
Date: 2023-02-22 16:17:57
LastEditors: MasterYip 2205929492@qq.com
LastEditTime: 2023-03-10 16:05:01
FilePath: \comprehensive-coding\InvoiceMgr\invoicemgr.py
Description: InvoiceMgr Ver1.0

Copyright (c) 2023 by ${git_name_email}, All Rights Reserved. 
'''

import os
import logging
import datetime
import json
import shutil
import tkinter as tk
import tkinter.messagebox as msgbox
from tkinter import filedialog
from glob import glob
from typing import Optional
import windnd
import fitz  # fitz就是pip install PyMuPDF
from pyzbar.pyzbar import decode
from PIL import Image

# Directory Management
try:
    # Run in Terminal
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
except:
    # Run in ipykernel & interactive
    ROOT_DIR = os.getcwd()
DB_DIR = os.path.join(ROOT_DIR,"Database")
OUTPUT_DIR = os.path.join(ROOT_DIR,"Output")
CONFIG_DIR = os.path.join(ROOT_DIR,"config.json")

# Config Data
if os.path.isfile(CONFIG_DIR):
    config:dict = json.load(open(CONFIG_DIR, 'r', encoding='utf8'))
    if config.get('OUTPUT_DIR'): OUTPUT_DIR = config.get('OUTPUT_DIR')
    if config.get('DB_DIR'): DB_DIR = config.get('DB_DIR')

"""
TODO: 
1.内存管理！！！
"""
# Logger
format_str = '%(asctime)s - %(name)s - %(levelname)s - %(filename)s[:%(lineno)d] - %(funcName)s - %(message)s'
datefmt_str = "%y-%m-%d %H:%M:%S"
# Remove existing handlers for basicConfig to take effect.
# TODO: This may not be a good idea, because this will infect other modules.
root_logger = logging.getLogger()
for h in root_logger.handlers: root_logger.removeHandler(h)
logging.basicConfig(filename=os.path.join(ROOT_DIR,'log.txt'),
                    format = format_str,
                    datefmt= datefmt_str,
                    level=logging.INFO)

cil_handler = logging.StreamHandler(os.sys.stderr)#默认是sys.stderr
cil_handler.setLevel(logging.INFO) # TODO: 会被BasicConfig限制？(过滤树)
cil_handler.setFormatter(logging.Formatter(fmt=format_str, datefmt=datefmt_str))

global_logger = logging.getLogger('Global')
global_logger.addHandler(cil_handler)


def pyMuPDF_fitz(pdfPath, imagePath, prefix=""):
    '''
    description: 
    param {*} pdfPath: pdf file dir
    param {*} imagePath: output folder of imgs
    return {*}
    '''
    startTime_pdf2img = datetime.datetime.now()  # 开始时间
    pdfDoc = fitz.open(pdfPath)
    for pg in range(pdfDoc.page_count):
        page = pdfDoc[pg]
        rotate = int(0)
        # 每个尺寸的缩放系数为，这将为我们生成分辨率提高的图像。
        # 此处若是不做设置，若图片大小为：792X612, dpi=96, (1.33333333-->1056x816)   (2-->1584x1224)
        # TODO: 默认大小？
        zoom_x = 2
        zoom_y = 2
        mat = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        if not os.path.exists(imagePath):
            os.makedirs(imagePath)
        output_dir = os.path.join(imagePath, prefix +'invoice_%s.png' % pg)
        pix.save(output_dir,'png')
    endTime_pdf2img = datetime.datetime.now()  # 结束时间
    global_logger.info('file saved to '+ os.path.relpath(output_dir) + 
                     ' | time: %f' % (endTime_pdf2img - startTime_pdf2img).seconds)

def get_qrcode(file_path):
    '''提取pdf文件中左上角的二维码并识别'''
    pdfDoc = fitz.open(file_path)
    page = pdfDoc[0]    #只对第一页的二维码进行识别
    rotate = int(0)
    zoom_x = 3.0
    zoom_y = 3.0
    mat = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
    rect = page.rect
    mp = rect.tl + (rect.br - rect.tl) * 1 / 5
    clip = fitz.Rect(rect.tl, mp)
    pix = page.get_pixmap(matrix=mat, alpha=False, clip=clip)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)  
    barcodes = decode(img)
    for barcode in barcodes:
        result = barcode.data.decode("utf-8")
        return result

def read_invoice_info(invoice_path:str, item:Optional[dict]=None):
    '''
    description: return invoice info in dict or write into given item.\n
    param {*} invoice_path\n
    param {dict} item\n
    return {*}
    '''
    if not item: item=dict() # Default Param is passed by reference, we need to reconstruct a new dict when no InvoiceItem is given.
    if invoice_path.startswith('#'):# Manual import
        # format "# CODE NUMBER VALUE DATE VERI "
        # e.g. "# 033002100911 35093895 256 20220824 07016763646873251240"
        global_logger.debug("manual import:%s",invoice_path)
        info_list = invoice_path[1:].strip().split(' ')
        if len(info_list) == 5 and len(info_list[0]) == 12 and len(info_list[1]) == 8 and\
        len(info_list[3]) == 8 and len(info_list[4]) == 20:
            item[InvoiceItem.INVOICE_CODE] = info_list[0]
            item[InvoiceItem.INVOICE_NUMBER] = info_list[1]
            item[InvoiceItem.INVOICE_VALUE] = info_list[2]
            item[InvoiceItem.INVOICE_DATE] = info_list[3]
            item[InvoiceItem.INVOICE_VERI] = info_list[4]
            item[InvoiceItem.INFO_IMPORTED] = True
        else:
            global_logger.error("Info check failed.")
    elif os.path.isfile(invoice_path) and os.path.splitext(invoice_path)[1] == ".pdf":
        infostr = get_qrcode(invoice_path)
        info_list = infostr.split(',')
        if len(info_list)>6:
            global_logger.debug("invoice info:%s",infostr)
            item[InvoiceItem.INVOICE_CODE] = info_list[2]
            item[InvoiceItem.INVOICE_NUMBER] = info_list[3]
            item[InvoiceItem.INVOICE_VALUE] = info_list[4]
            item[InvoiceItem.INVOICE_DATE] = info_list[5]
            item[InvoiceItem.INVOICE_VERI] = info_list[6]
            item[InvoiceItem.INFO_IMPORTED] = True
        else:
            global_logger.error("QRCode not found or info format doesn't match.")
    else:
        global_logger.error("failed to read invoice file.")
    return item

class EntryFrame():
    def __init__(self,parent,name=""):
        self.parent =parent
        self.name = name
        self.font = ('黑体',12)
        self.frame = tk.Frame(self.parent)
        self.file_address = tk.StringVar()
        
        self.label = tk.Label(self.frame, text=self.name+": ",font=self.font)
        self.entry = tk.Entry(self.frame, textvariable=self.file_address, font=self.font, 
                              highlightcolor='Fuchsia', highlightthickness=1, width=80)
        self.button = tk.Button(self.frame, text="Open...", font=self.font, command=self.select_file)
        
        self.label.grid(row=0, column=0)
        self.entry.grid(row=0, column=1)
        self.button.grid(row=0, column=2)

        windnd.hook_dropfiles(self.entry, func=self.dragged_files)

        self.frame.pack()
        
    def select_file(self):
        file = filedialog.askopenfilename(initialdir=os.getcwd())
        self.file_address.set(file)
    
    # TODO: Win10拖放失效问题
    def dragged_files(self,files):
        # msg = '\n'.join((item.decode('gbk') for item in files))
        # msgbox.showinfo('您拖放的文件', msg)
        self.file_address.set(files[0].decode('gbk'))

class APP(object):
    def __init__(self, width=840, height=350):
        self.w = width
        self.h = height
        self.title = 'Invoice Manager Ver1.0'
        self.data_dir = os.path.join(DB_DIR,"data.txt")
        if not os.path.exists(DB_DIR): os.mkdir(DB_DIR)
        self.logger = logging.getLogger('App')
        self.logger.addHandler(cil_handler)
        # self.logger.addHandler(file_handler)
        self.root = tk.Tk(className=self.title)
        self.font = ('黑体',12)
        self.itemlist = []
        self.root.iconbitmap(default=os.path.join(ROOT_DIR,"icon.ico"))
        
        # 定义文字
        self.itemname = tk.StringVar()
        self.total = tk.Variable()
        self.itemcnt = tk.Variable()
        self.info_disp = tk.StringVar()
        # Frame空间
        frame_listbox = tk.Frame(self.root)
        frame_bottom = tk.Frame(self.root)
        
        # Menu菜单
        menu = tk.Menu(self.root)                                                            
        self.root.config(menu=menu)
        aboutmenu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label='Author: Master Yip', menu=aboutmenu)
        # 控件内容设置
        self.listbox = tk.Listbox(frame_listbox, selectmode=tk.EXTENDED, height=10, width=100, setgrid=False)
        self.listbox.bind('<<ListboxSelect>>',self.itemselected_callback)
        self.name_entry = tk.Entry(frame_bottom, textvariable=self.itemname, font=self.font, 
                        highlightcolor='Fuchsia', highlightthickness=1, width=50)
        self.info_label = tk.Entry(frame_bottom, textvariable=self.info_disp, font=self.font, 
                        highlightcolor='Fuchsia', highlightthickness=1, width=50)
        self.button_add = tk.Button(frame_bottom, text="Add/Update", font=self.font, width=15, command=self.add_item)
        self.button_batchimport = tk.Button(frame_bottom, text="Batch Import", font=self.font, width=15, command=self.batch_import)
        self.button_output = tk.Button(frame_bottom, text="Output Select", font=self.font, width=15, command=self.sel_item_output)
        self.button_delete = tk.Button(frame_bottom, text="Del Select", font=self.font, width=15, command=self.sel_item_del)
        self.button_state_done = tk.Button(frame_bottom, text="Stat Done", font=self.font, width=15, command=self.sel_item_state_done)
        self.button_state_undone = tk.Button(frame_bottom, text="Stat UnDn", font=self.font, width=15, command=self.sel_item_state_undone)
        
        self.sortlist = tk.Listbox(frame_listbox,selectmode=tk.SINGLE, height=10, width=10, setgrid=False)
        self.sortlist.insert(tk.END,'Value Up')
        self.sortlist.insert(tk.END,'Value Down')
        self.sortlist.insert(tk.END,'Date Up')
        self.sortlist.insert(tk.END,'Date Down')
        self.sortlist.insert(tk.END,'States')
        self.sortlist.insert(tk.END,'Name')
        self.sortlist.bind('<<ListboxSelect>>',self.sortmethod_callback)
        # 控件布局
        frame_listbox.pack()
        self.invoice_frame = EntryFrame(self.root,"Invoice ")
        self.order_frame = EntryFrame(self.root,"Order   ")
        self.transfer_frame = EntryFrame(self.root,"Transfer")
        frame_bottom.pack(pady=10)

        self.listbox.grid(row=0, column=0)
        self.sortlist.grid(row=0, column=1)
        self.name_entry.grid(row=0, column=0)
        self.button_add.grid(row=0, column=1)
        self.button_state_done.grid(row=0, column=2)
        self.button_state_undone.grid(row=0, column=3)
        self.info_label.grid(row=1, column=0)
        self.button_batchimport.grid(row=1, column=1)
        self.button_output.grid(row=1, column=2)
        self.button_delete.grid(row=1, column=3)
        
        self.load()
        self.database_clear()
    
    def itemselected_callback(self, event):
        items_index = self.listbox.curselection()
        if len(items_index) == 1:
            item = self.itemlist[items_index[0]]
            if item[item.INVOICE_NAME].startswith('#'): self.invoice_frame.file_address.set(item[item.INVOICE_NAME])
            elif item[item.INVOICE_NAME]:self.invoice_frame.file_address.set(item.abspath(item[item.INVOICE_NAME]))
            else: self.invoice_frame.file_address.set("")
            if item[item.ORDER_NAME]:self.order_frame.file_address.set(item.abspath(item[item.ORDER_NAME]))
            else: self.order_frame.file_address.set("")
            if item[item.TRANSFER_NAME]:self.transfer_frame.file_address.set(item.abspath(item[item.TRANSFER_NAME]))
            else: self.transfer_frame.file_address.set("")
            self.itemname.set(item['name'])
        elif len(items_index) > 1:
            self.invoice_frame.file_address.set("")
            self.order_frame.file_address.set("")
            self.transfer_frame.file_address.set("")
            self.itemname.set("")
        
        self.total.set(0)
        self.itemcnt.set(items_index.__len__())
        for it in items_index:
            self.total.set(self.total.get()+float(self.itemlist[it][InvoiceItem.INVOICE_VALUE]))
        self.info_disp.set("Total(taxfree):{:.2f}, Count:{:d}".format(self.total.get(),self.itemcnt.get()))

    def sortmethod_callback(self, event):
        # TODO: Add other sorting method.
        select = self.sortlist.curselection()
        attr = -1
        if select:
            attr = select[0]
        if attr == 0:   # Value Up
            self.itemlist.sort(key=lambda x:float(x[x.INVOICE_VALUE]))
        elif attr == 1: # Value Down
            self.itemlist.sort(key=lambda x:float(x[x.INVOICE_VALUE]),reverse=True)
        elif attr == 2: # Date Up
            self.itemlist.sort(key=lambda x:x[x.INVOICE_DATE])
        elif attr == 3: # Date Down
            self.itemlist.sort(key=lambda x:x[x.INVOICE_DATE],reverse=True)
        elif attr == 4: # State (Done/Order/Transfer)
            self.itemlist.sort(key=lambda x:10*int(x['state'])+int(bool(x[x.ORDER_NAME])))
        elif attr == 5: # Name
            self.itemlist.sort(key=lambda x:x['name'])
        self.refresh_listbox()
        self.save()
              
    def center(self):
        """
        函数说明:tkinter窗口居中
        """
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        x = int((ws / 2) - (self.w / 2))
        y = int((hs / 2) - (self.h / 2))
        self.root.geometry('{}x{}+{}+{}'.format(self.w, self.h, x, y))

    def loop(self):
        """
        函数说明:loop等待用户事件
        """
        # 禁止修改窗口大小
        self.root.resizable(True, True)
        # 窗口居中
        self.center()
        self.root.mainloop()
  
    def add_item(self):
        file = self.invoice_frame.file_address.get()
        if file:
            item_exists = False
            info = read_invoice_info(file)
            for item in self.itemlist:
                if item[item.INVOICE_CODE] == info.get(item.INVOICE_CODE) and\
                item[item.INVOICE_NUMBER] == info.get(item.INVOICE_NUMBER):
                    item_exists = True
                    self.logger.info("%s exists.", os.path.basename(file))
                    self.itemlist[self.itemlist.index(item)].edit(self.itemname.get(),
                                        self.invoice_frame.file_address.get(),
                                        self.order_frame.file_address.get(),
                                        self.transfer_frame.file_address.get())
                    break
            if not item_exists:
                self.itemlist.append(InvoiceItem(self.itemname.get(),
                                        self.invoice_frame.file_address.get(),
                                        self.order_frame.file_address.get(),
                                        self.transfer_frame.file_address.get()))
        else:
            msgbox.showerror("Error","Invoice is empty.")
        self.refresh_listbox()
        self.save()
    
    def batch_import(self):
        invoice_files = filedialog.askopenfilenames(initialdir=os.getcwd())
        for file in invoice_files:
            if os.path.isfile(file) and os.path.splitext(file)[1] == ".pdf": #pdf check
                info = read_invoice_info(file)
                item_exists = False
                if info:
                    for item in self.itemlist:
                        if item[item.INVOICE_CODE] == info.get(item.INVOICE_CODE) and\
                        item[item.INVOICE_NUMBER] == info.get(item.INVOICE_NUMBER):
                            item_exists = True
                            break
                if not item_exists and info:
                    self.itemlist.append(InvoiceItem(os.path.splitext(os.path.basename(file))[0], file))
                elif info: self.logger.info("%s exists.", os.path.basename(file))
                else: self.logger.warning("No info is found in %s", file)
                
            else:self.logger.warning("Not valid invoice file(pdf): %s", file)
        self.refresh_listbox()
        self.save()
    
    def refresh_listbox(self):
        '''
        description: Refresh listbox according to database.
        param {*} self
        return {*}
        '''    
        self.listbox.delete(0,tk.END)
        for item in self.itemlist:
            self.listbox.insert(tk.END,item.get_listbox_text())
    
    def save(self, dir=""):
        '''
        description: Save data to database.
        param {*} self
        param {*} dir
        return {*}
        '''    
        if not dir:
            dir = self.data_dir
        json.dump(self.itemlist, open(dir, 'w', encoding='utf8'),
                  ensure_ascii=False, indent=4, sort_keys=True)
    
    def load(self, dir=""):
        '''
        description: Load data from database.
        param {*} self
        param {*} dir
        return {*}
        '''     
        if not dir:
            dir = self.data_dir
        if os.path.isfile(dir):
            for item in json.load(open(dir, 'r', encoding='utf8')):
                self.itemlist.append(InvoiceItem(data=item))
        else:
            self.logger.warning("No file is found. Skipping...")
        self.refresh_listbox()
        self.items_itemfiles_check()
    
    def items_itemfiles_check(self):
        for item in self.itemlist:
            item.itemfiles_check()
        self.refresh_listbox()
        self.save()

    def sel_item_output(self):
        if os.path.exists(OUTPUT_DIR):shutil.rmtree(OUTPUT_DIR)
        if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
        self.logger.info("Output to %s", OUTPUT_DIR)
        num = 1
        for index in self.listbox.curselection():
            self.itemlist[index].invoice_files_output(prefix="{:02d}_".format(num))
            num += 1
        os.startfile(OUTPUT_DIR)
         
    def sel_item_del(self):
        for index in tuple(reversed(self.listbox.curselection())):
            shutil.rmtree(self.itemlist[index].itemfolder_abspath())
            self.itemlist.pop(index)
        self.refresh_listbox()
        self.save()
    
    def sel_item_state_done(self):
        for item in self.listbox.curselection():
            self.itemlist[item]['state'] = True
        self.refresh_listbox()
        self.save()
        
    def sel_item_state_undone(self):
        for item in self.listbox.curselection():
            self.itemlist[item]['state'] = False
        self.refresh_listbox()
        self.save()
    
    def database_clear(self):
        '''
        description: Clear database (delete useless files&items which is not recorded)
        param {*} self
        return {*}
        '''    
        folder_list = []
        for item in self.itemlist:
            folder_list.append(item.itemfolder_abspath())
        exist_folders = glob(DB_DIR+"/*",recursive=False)
        for folder in exist_folders:
            if not folder in folder_list and os.path.isdir(folder):
                self.logger.info("%s will be deleted.",folder)
                shutil.rmtree(folder)
    
class InvoiceItem(dict):
    '''
    description: 
    param {*} self
    param {*} name
    param {*} invoice
    param {*} order
    param {*} transfer
    return {*}
    '''    
    INFO_IMPORTED = "info_imported" #发票信息是否已导入
    
    INVOICE_NAME = "invoice"        #发票database相对路径
    INVOICE_CODE = "invoice_code"   #发票代码
    INVOICE_NUMBER = "invoice_number" #发票号码
    INVOICE_VALUE = "invoice_value" #发票面值 TODO: 似乎是金额（不含税额）
    INVOICE_DATE = "invoice_date"   #开票日期
    INVOICE_VERI = "invoice_veri"   #发票验证码
    
    ORDER_NAME = "order"            #订单database相对路径
    TRANSFER_NAME = "transfer"      #转账database相对路径
    

    def __init__(self,name:str=None,
                 invoice_origin_dir=None,
                 order_origin_dir=None,
                 transfer_origin_dir=None,
                 data:dict=None):
        self.logger = logging.getLogger('InvoiceItem')
        self.logger.addHandler(cil_handler)
        # self.logger.addHandler(file_handler)
        if data:# Load from file
            super().__init__(data)
        else:# Load from App
            self['name'] = ""
            self['state'] = False
            # Relpath of files
            self[self.INVOICE_NAME]=""   # self['invoice']
            self[self.ORDER_NAME]=""     # self['order']
            self[self.TRANSFER_NAME]=""  # self['transfer']
            # Invoice Info
            self[self.INVOICE_CODE] = ""
            self[self.INVOICE_NUMBER] = ""
            self[self.INVOICE_VALUE] = ""
            self[self.INVOICE_DATE] = ""
            self[self.INVOICE_VERI] = ""
            self[self.INFO_IMPORTED] = False # If not imported, error is raised.
            
            if name: self['name'] = name
            self.read_invoice_info(invoice_origin_dir)
            self.add_file(self.INVOICE_NAME,invoice_origin_dir)
            self.add_file(self.ORDER_NAME,order_origin_dir)
            self.add_file(self.TRANSFER_NAME,transfer_origin_dir)
            self.itemfiles_check()
            
    def edit(self,name:str=None,
                invoice_origin_dir=None,
                order_origin_dir=None,
                transfer_origin_dir=None):
        if name: self['name'] = name
        self.read_invoice_info(invoice_origin_dir)
        self.add_file(self.INVOICE_NAME,invoice_origin_dir)
        self.add_file(self.ORDER_NAME,order_origin_dir)
        self.add_file(self.TRANSFER_NAME,transfer_origin_dir)
        self.itemfiles_check()
    
    def itemfolder_abspath(self):
        return os.path.join(DB_DIR,'_'.join((self[self.INVOICE_DATE],self[self.INVOICE_CODE],self[self.INVOICE_NUMBER])))
    
    def abspath(self,dbpath):
        '''
        description: Generate abspath using database relative path
        param {*} self
        param {*} dbpath
        return {*} abspath
        '''    
        return os.path.join(DB_DIR,dbpath)
   
    def itemfiles_check(self):
        '''
        description: clear recorded dbpath matching no file and delete files not recorded.
        param {*} self
        return {*}
        '''    
        files = [] 
        if os.path.isfile(self.abspath(self[self.INVOICE_NAME])):files.append(self.abspath(self[self.INVOICE_NAME]))
        elif not self[self.INVOICE_NAME].startswith('#'): self[self.INVOICE_NAME]=""
        if not os.path.isfile(self.abspath(self[self.ORDER_NAME])):self[self.ORDER_NAME]=""
        else: files.append(self.abspath(self[self.ORDER_NAME]))
        if not os.path.isfile(self.abspath(self[self.TRANSFER_NAME])):self[self.TRANSFER_NAME]=""
        else: files.append(self.abspath(self[self.TRANSFER_NAME]))
        detected_files = glob(self.itemfolder_abspath()+"/*",recursive=False) #TODO: Linux Slash is suppoted?
        self.logger.debug("files detected:%s",' | '.join(files))
        for file in detected_files:
            if not file in files:
                self.logger.debug("file to be del:%s",file)
                os.remove(file)
   
    def add_file(self,filetypename,src):    
        '''
        description: Add specific filw from src to database
        param {*} self
        param {*} filetypename
        param {*} src
        return {*}
        '''    
        if not os.path.isdir(self.itemfolder_abspath()):
            os.makedirs(self.itemfolder_abspath())
        if src and os.path.isfile(src):
            abs_dst = os.path.join(self.itemfolder_abspath(),filetypename+os.path.splitext(src)[1])
            if os.path.isfile(abs_dst):
                self.logger.info(filetypename+" file will be overwritted.")
            if abs_dst == src:
                self.logger.info("The file is the same.")
            else:
                shutil.copy(src,abs_dst)
            self[filetypename] = os.path.relpath(abs_dst,DB_DIR)
        elif src and src.startswith('#') and filetypename == self.INVOICE_NAME: # Manual import, skip.
            self.logger.info("Manual import without file.")
            self[filetypename] = src
        elif src:
            self.logger.info(src+" doesn't exist.")
        else:
            self.logger.info("Failed to add "+filetypename+".")
    
    def read_invoice_info(self,invoice_path = ""):
        if not invoice_path:
            invoice_path = self.abspath(self[self.INVOICE_NAME])
        read_invoice_info(invoice_path, self)
        if not self[self.INFO_IMPORTED]:raise Exception("Import failed.")
            
    def get_listbox_text(self):
        '''
        description: Generate text which will show in listbox.
        param {*} self
        return {*}
        '''    
        # return " | ".join([self['name'],self[self.INVOICE_DATE],
        #                    self[self.INVOICE_VALUE],"Order:"+self[self.ORDER_NAME]])
        stat = '✔' if self['state'] else '✖'
        ordr = '●' if self[self.ORDER_NAME] else '○'
        tran = '●' if self[self.TRANSFER_NAME] else '○'
        
        return "{:^1}|{:^1}|{:^1}| {:^8} | {:<8} | {:^8s}".format(stat,ordr,tran,self[self.INVOICE_DATE],
                                                     self[self.INVOICE_VALUE],self['name'])

    def file_output(self,src,dstdir=OUTPUT_DIR,prefix=""):
        '''
        description: 
        param {*} self
        param {*} src: abspath of srcfile
        param {*} dst: abspath of dstfile
        return {*}
        '''    
        if not os.path.exists(dstdir): os.makedirs(dstdir)
        ext = os.path.splitext(src)[1]
        if ext == ".pdf":
            pyMuPDF_fitz(src,dstdir,prefix)
        elif ext== ".jpg" or ext==".png":
            shutil.copy(src,os.path.join(dstdir,prefix+os.path.basename(src)))
        else:
            self.logger.error("Unkown file format:%s",src)
            
    def invoice_files_output(self,prefix=""):
        '''
        description: Output files of an invoice item
        param {*} self
        return {*}
        '''    
        if self[self.INVOICE_NAME] and not self[self.INVOICE_NAME].startswith('#'):
            self.file_output(self.abspath(self[self.INVOICE_NAME]),prefix=prefix)
        if self[self.ORDER_NAME]:self.file_output(self.abspath(self[self.ORDER_NAME]),prefix=prefix)
        if self[self.TRANSFER_NAME]:self.file_output(self.abspath(self[self.TRANSFER_NAME]),prefix=prefix)
        
if __name__ == '__main__':
    app = APP()
    app.loop()
       