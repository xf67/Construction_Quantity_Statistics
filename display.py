import os
import tkinter as tk
import tkinter.ttk as ttk
import glob
from show import ShowExcel

class DisplayDir():
    def __init__(self):
        self.root=tk.Tk()
        self.root.title('文件浏览')

        #水平和垂直滚动条
        vsb = ttk.Scrollbar(orient="vertical")
        hsb = ttk.Scrollbar(orient="horizontal")
        
        self.tree = ttk.Treeview(columns=("fullpath", "type", "size"), #指定列
        displaycolumns="size",  #指定显示的列,#0列默认显示
        yscrollcommand=lambda f, l: self._autoscroll(vsb, f, l), #水平滚动条
        xscrollcommand=lambda f, l:self._autoscroll(hsb, f, l)) #垂直滚动条
        #lambda函数的运算结果就是函数的表达式,输入的是f,l

        vsb['command'] = self.tree.yview 
        hsb['command'] = self.tree.xview

        self.tree.heading("#0", text="文件结构", anchor='w')
        self.tree.heading("size", text="文件大小", anchor='w')
        self.tree.column("size", stretch=0, width=100)

        self.populate_roots(self.tree)
        self.tree.bind('<<TreeviewOpen>>', self.update_tree) #绑定打开事件
        self.tree.bind('<Double-Button-1>', self.change_dir) #绑定双击事件

        # 将 Treeview 组件和滚动条组件放置在窗口中
        self.tree.grid(column=0, row=0, sticky='nswe')
        vsb.grid(column=1, row=0, sticky='ns')
        hsb.grid(column=0, row=1, sticky='ew')
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

    # 显示目录
    def populate_tree(self, tree, node):
        
        if tree.set(node, "type") != 'directory':
            return

        path = tree.set(node, "fullpath")
        tree.delete(*tree.get_children(node))
        # 一个星号是接受元组,两个星号是接受字典,用星号告诉程序要"解包"

        parent = tree.parent(node)
        special_dirs = [] if parent else glob.glob('.') + glob.glob('..')
        # 三目运算符,[] 或 . ..两个特殊文件夹

        for p in special_dirs + os.listdir(path): # 遍历文件目录
            ptype = None
            p = os.path.join(path, p).replace('\\', '/') # 得到绝对路径,同时解决win的问题
            if os.path.isdir(p): ptype = "directory"
            elif os.path.isfile(p): ptype = "file"

            # 将新项目插入到treeview中
            fname = os.path.split(p)[1]
            id = tree.insert(node, "end", text=fname, values=[p, ptype])

            if ptype == 'directory':
                if fname not in ('.', '..'):
                    tree.insert(id, 0, text="dummy") #  暂时将下面的节点设为dummy
                    tree.item(id, text=fname)
            elif ptype == 'file':
                size = os.stat(p).st_size
                size = size / 1024
                tree.set(id, "size", "%d KB" % size)

    # 显示当前目录
    def populate_roots(self, tree):
        dir = os.path.abspath('.').replace('\\', '/') # 得到当前目录的绝对路径
        node = tree.insert('', 'end', text=dir, values=[dir, "directory"])
        self.populate_tree(tree, node)

    # 展开一层目录
    def update_tree(self, event):
        tree = event.widget
        self.populate_tree(tree, tree.focus())# focus会输出被选中项的iid

    # 更换目录或者打开文件
    def change_dir(self, event):
        tree = event.widget
        node = tree.focus()
        if tree.parent(node):
            path = os.path.abspath(tree.set(node, "fullpath"))
            if os.path.isdir(path):
                os.chdir(path)
                tree.delete(tree.get_children(''))
                self.populate_roots(tree)
            if not os.path.isdir(path):
                ShowExcel(path)

    # 自动显示滚轮
    def _autoscroll(self, sbar, first, last):
        first, last = float(first), float(last)
        if first <= 0 and last >= 1:
            sbar.grid_remove()
        else:
            sbar.grid()
        sbar.set(first, last)
 

if __name__ == '__main__':
    ds=DisplayDir()
    ds.root.mainloop()
