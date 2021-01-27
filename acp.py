#!/usr/bin/env python3

from tkinter import *
from tkinter.filedialog import askopenfilename
from process_acp import Acp
import re
import pprint


class Run:
    path: str

    def __init__(self):
        self.window = Tk()
        self.acp = Acp()
        self.path = ""
        self.window.geometry('600x250')
        self.window.title("ACP")
        self.label = Label(self.window, text="Please choose file.")
        self.label_2 = Label(self.window)
        self.label_wait = Label(self.window)
        self.label_done = Label(self.window)
        self.but = Button(self.window, text="Open", command=self.open_file)
        self.but_start = Button(self.window, text="Start process", command=self.start)
        self.entry_npp = Entry(self.window, width=3)
        self.entry_npp.grid(column=1, row=0)
        self.entry_npp.insert(0, "1")
        self.label_enter_npp = Label(self.window, text="Enter sort column:").grid(column=0, row=0)
        self.enrty_kd = Entry(self.window, width=3)
        self.enrty_kd.grid(column=3, row=0)
        self.enrty_kd.insert(0, "7")
        self.label_enter_kd = Label(self.window, text="Enter KD column:").grid(column=2, row=0)
        self.entry_bs = Entry(self.window, width=3)
        self.entry_bs.grid(column=5, row=0)
        self.entry_bs.insert(0, "8")
        self.label_enter_bs = Label(self.window, text="Enter BS column:").grid(column=4, row=0)

    def open_file(self):
        self.label_wait.configure(text="")
        self.label_done.configure(text="")
        self.path = askopenfilename(filetypes = (("Exel files","*.xlsx"),("all files","*.*")))
        self.label_2.configure(text="File - " + self.path)
        self.but_start.grid(column=2, row=1, padx=10)

    def start(self):
        if not self.path:
            return
        entry_npp = int(self.entry_npp.get()) - 1
        enrty_kd = int(self.enrty_kd.get()) - 1
        entry_bs = int(self.entry_bs.get()) - 1
        self.label_wait.configure(text="Process is started. Please wait...")
        # get table from excel file
        tb_title, tb_raw = self.acp.load_table_new(filepath=self.path, entry_npp=entry_npp)
        # group by NPP
        tb_by_npp = self.acp.group_by_npp(tb_raw)
        # count unique companies
        count_companies = self.acp.number_of_unique_companies(tabl_by_npp=tb_by_npp)
        # get specific_weight
        tb_specific_waight = self.acp.specific_weight(tabl_by_npp=tb_by_npp)
        # write file
        f_w = re.sub(r'.xlsx|.xls', '_processed.xlsx', self.path)
        self.acp.write_to_file_new(file_name=f_w, table=tb_specific_waight)
        self.label_done.configure(text="DONE")

    def run(self):
        self.label.grid(column=0, row=1)
        self.but.grid(column=1, row=1)
        self.label_2.grid(column=0, row=2, columnspan=30)
        self.label_wait.grid(column=0, row=3, columnspan=30)
        self.label_done.grid(column=0, row=4, columnspan=30)
        self.window.mainloop()


if __name__ == '__main__':
    Run().run()
