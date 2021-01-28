#!/usr/bin/env python3

from tkinter import *
from tkinter.filedialog import askopenfilename
from process_acp import Acp
from catalog import Catalog
import re
import time
import ntpath
import pprint


class Run:
    path: str
    catalog_path: str

    def __init__(self):
        self.window = Tk()
        self.acp = Acp()
        self.catalog = Catalog()
        self.path = ""
        self.catalog_path = ""
        self.catalog_content = []
        self.catalog_levels = []
        self.window.geometry('700x450')
        self.window.title("ACP")
        self.selected = StringVar()
        self.label_enter_bs = Label(self.window, text=" Enter BS column:").grid(column=4, row=0, sticky=E)
        self.entry_bs = Entry(self.window, width=3)
        self.label_enter_kd = Label(self.window, text="Enter KD column:")
        self.label_enter_kd.grid(column=2, row=0, sticky=E)
        self.enrty_kd = Entry(self.window, width=3)
        self.label_enter_npp = Label(self.window, text="     Enter sort column:")
        self.label_enter_npp.grid(column=0, row=0, sticky=E)
        self.entry_npp = Entry(self.window, width=3)
        self.but_start = Button(self.window, text="Start process", command=self.start)
        self.but = Button(self.window, text="Open", command=self.open_file)
        self.label_done = Label(self.window)
        self.label_wait = Label(self.window)
        self.label_2 = Label(self.window)
        self.label = Label(self.window, text="Please choose file.")
        self.label_catalog = Label(self.window, text="Please choose catalogue.")
        self.but_catalog = Button(self.window, text="Open", command=self.open_catalog)
        self.label_catalog_file = Label(self.window)

    def open_file(self):
        self.label_wait.configure(text="")
        self.label_done.configure(text="")
        self.path = askopenfilename(filetypes = (("Exel files","*.xlsx"),("all files","*.*")))
        self.label_2.configure(text=" " + ntpath.basename(self.path))
        self.but_start.grid(column=1, row=3, columnspan=4, sticky=W)

    def open_catalog(self):
        self.catalog_path = askopenfilename(filetypes = (("Exel files","*.xlsx"),("all files","*.*")))
        self.label_catalog_file.configure(text=" " + ntpath.basename(self.catalog_path))
        self.catalog_levels, self.catalog_content = self.catalog.get_catalog(filepath=self.catalog_path)
        # self.label_enter_npp.grid_remove()
        # self.entry_npp.grid_remove()
        i = 0
        for catalog_level in self.catalog_levels:
            Radiobutton(self.window, text=catalog_level, value=catalog_level, variable=self.selected).grid(column=6, row=i, sticky=W)
            i += 1

    def start(self):
        if not self.path:
            return
        entry_npp = int(self.entry_npp.get()) - 1
        enrty_kd = int(self.enrty_kd.get()) - 1
        entry_bs = int(self.entry_bs.get()) - 1
        self.label_wait.configure(text="Process is started. Please wait...")
        self.window.update()
        # get table from excel file
        tb_title, tb_raw = self.acp.load_table(filepath=self.path)
        # redefined tb_raw and entry_npp
        if self.catalog_path != "" and self.selected.get() != "":
            entry_npp, tb_raw = self.catalog.rebuild_tabl(
                entry_npp=entry_npp,
                selected_level=self.selected.get(),
                levels=self.catalog_levels,
                catalog_content=self.catalog_content,
                raw_tabl=tb_raw)
        # group by NPP
        tb_by_npp = self.acp.group_by_npp(tb_raw, entry_npp, entry_bs)
        # count unique companies
        count_companies = self.acp.number_of_unique_companies(tabl_by_npp=tb_by_npp, enrty_kd=enrty_kd)
        # get specific_weight
        tb_specific_waight = self.acp.specific_weight(tabl_by_npp=tb_by_npp, entry_kd=enrty_kd, entry_bs=entry_bs)
        # write file
        suffix = time.strftime("%Y_%m_%d_%H%M%S")
        f_w = re.sub(r'.xlsx|.xls', '_processed_' + suffix + '.xlsx', self.path)
        self.acp.write_to_file_new(file_name=f_w, table=tb_specific_waight, table_count=count_companies)
        self.label_done.configure(text="DONE")

    def run(self):
        self.entry_npp.grid(column=1, row=0, sticky=W)
        self.entry_npp.insert(0, "1")
        self.enrty_kd.grid(column=3, row=0, sticky=W)
        self.enrty_kd.insert(0, "7")
        self.entry_bs.grid(column=5, row=0, sticky=W)
        self.entry_bs.insert(0, "8")
        self.label_catalog.grid(column=0, row=1, sticky=E)
        self.but_catalog.grid(column=1, row=1, sticky=W)
        self.label_catalog_file.grid(column=2, row=1, sticky=W, columnspan=4)
        self.label.grid(column=0, row=2, sticky=E)
        self.but.grid(column=1, row=2, sticky=W)
        self.label_2.grid(column=2, row=2, sticky=W, columnspan=4)
        self.label_wait.grid(column=1, row=4, columnspan=4)
        self.label_done.grid(column=1, row=5, columnspan=4)
        self.window.mainloop()


if __name__ == '__main__':
    Run().run()
