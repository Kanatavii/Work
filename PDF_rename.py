from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout, QTreeView, QFileSystemModel, QPushButton, QHeaderView
from PyQt5.QtCore import QDir, Qt
import os
import sys

class CustomTreeView(QTreeView):
    def __init__(self, parent=None):
        super(CustomTreeView, self).__init__(parent)
    
    def mouseDoubleClickEvent(self, event):
        # 获取双击的项目索引
        index = self.indexAt(event.pos())
        # 检查是否是目录
        if self.model().isDir(index):
            # 取消选中状态
            self.selectionModel().clearSelection()
            # 调用默认的双击行为
            super(CustomTreeView, self).mouseDoubleClickEvent(event)

class MultiFolderDialog(QDialog):
    def __init__(self, parent=None):
        super(MultiFolderDialog, self).__init__(parent)
        
        self.setWindowTitle("选择多个文件夹")
        self.resize(800, 600)
        
        layout = QVBoxLayout(self)
        
        # 文件系统模型，用于显示文件夹
        self.model = QFileSystemModel()
        self.model.setRootPath('')
        self.model.setFilter(QDir.Dirs | QDir.NoDotAndDotDot)
        
        # 树形视图，用于展示文件夹结构
        self.tree = CustomTreeView(self)
        self.tree.setModel(self.model)
        self.tree.setSelectionMode(QTreeView.MultiSelection)
        
        # 设置固定列宽为600像素
        self.tree.header().setSectionResizeMode(0, QHeaderView.Fixed)  # 固定第0列（Name列）的宽度
        self.tree.setColumnWidth(0, 600)  # 设置Name列的固定宽度为600像素
        
        layout.addWidget(self.tree)
        
        # 确定按钮
        self.button_ok = QPushButton("确定", self)
        self.button_ok.clicked.connect(self.accept)
        layout.addWidget(self.button_ok)
        
        self.selected_folders = []
    
    def accept(self):
        indexes = self.tree.selectedIndexes()
        self.selected_folders = [self.model.filePath(index) for index in indexes if self.model.isDir(index)]
        super(MultiFolderDialog, self).accept()

def select_multiple_folders_and_get_filenames():
    app = QApplication(sys.argv)
    
    dialog = MultiFolderDialog()
    if dialog.exec_() == QDialog.Accepted:
        folders = dialog.selected_folders
        
        all_files = []
        for folder in folders:
            if os.path.isdir(folder):  # 确保处理的是文件夹
                for root_dir, dirs, files in os.walk(folder):
                    for file in files:
                        all_files.append(os.path.join(root_dir, file))
        
        return all_files
    else:
        return []

if __name__ == "__main__":
    filenames = select_multiple_folders_and_get_filenames()
    for filename in filenames:
        print(filename)
    
    sys.exit(0)  # 确保应用程序正确退出
