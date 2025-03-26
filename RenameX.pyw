import sys
import os
import re
import random
import string
import traceback
import datetime
import calendar
import math

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QListWidget, QFileDialog, QComboBox, 
    QSpinBox, QCheckBox, QGroupBox, QRadioButton, QMessageBox, 
    QGridLayout, QButtonGroup, QAbstractItemView, QSizePolicy, QFrame, 
    QToolButton, QSplitter, QListWidgetItem, QMenu, QAction, 
    QDialog, QDialogButtonBox, QFileSystemModel, QTreeView, QHeaderView, 
    QTextEdit
)
from PyQt5.QtCore import Qt, QEvent, QDir
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QIcon, QFont, QColor, QBrush


# =============================================================================
# 辅助函数：检测文件名中是否包含 Windows 不允许的字符
# =============================================================================
def has_illegal_chars(name):
    illegal_chars = ['"', '*', '<', '>', '?', '\\', '/', '|', ':']
    return any(ch in name for ch in illegal_chars)


# =============================================================================
# RenameDialog：用于手动重命名的对话框
# =============================================================================
class RenameDialog(QDialog):
    def __init__(self, current_text, parent=None):
        super().__init__(parent)
        self.setWindowTitle("重命名")
        self.setMinimumWidth(500)
      
        layout = QVBoxLayout()
        label = QLabel("输入新名称:")
        layout.addWidget(label)
      
        self.name_edit = QLineEdit(current_text)
        self.name_edit.setMinimumWidth(480)
        self.name_edit.selectAll()  # 默认全选当前名称
        layout.addWidget(self.name_edit)
      
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
      
        self.setLayout(layout)
  
    def get_new_name(self):
        return self.name_edit.text()


# =============================================================================
# HelpDialog：帮助说明对话框
# =============================================================================
class HelpDialog(QDialog):
    def __init__(self, help_text, parent=None):
        super().__init__(parent)
        self.setWindowTitle("命名规则帮助")
        self.setMinimumWidth(600)
        self.setMinimumHeight(400)
        self.resize(700, 500)
      
        layout = QVBoxLayout()
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setPlainText(help_text)
        font = self.text_edit.font()
        font.setPointSize(10)
        self.text_edit.setFont(font)
        layout.addWidget(self.text_edit)
      
        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
      
        self.setLayout(layout)


# =============================================================================
# SingleRuleDialog：单条规则编辑对话框（实时预览）
# =============================================================================
class SingleRuleDialog(QDialog):
    def __init__(self, file_path, default_rule, parent=None):
        super().__init__(parent)
        self.setWindowTitle("单条规则编辑")
        self.setMinimumWidth(600)
        self.setMinimumHeight(200)
      
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.parent_window = parent  # 主窗口
      
        layout = QVBoxLayout()
        file_info = QLabel(f"为文件设置单独规则: {self.file_name}")
        file_info.setStyleSheet("font-weight: bold;")
        layout.addWidget(file_info)
      
        rule_layout = QHBoxLayout()
        rule_layout.addWidget(QLabel("单条规则:"))
        self.rule_edit = QLineEdit(default_rule)
        self.rule_edit.setPlaceholderText("使用 * 替代原文件名，详见帮助")
        self.rule_edit.textChanged.connect(self.update_preview)
        rule_layout.addWidget(self.rule_edit)
      
        help_btn = QToolButton()
        help_btn.setText("?")
        from functools import partial
        help_btn.clicked.connect(partial(self.show_help, self.get_naming_help_text()))
        rule_layout.addWidget(help_btn)
        layout.addLayout(rule_layout)
      
        preview_layout = QHBoxLayout()
        preview_layout.addWidget(QLabel("预览结果:"))
        self.preview_label = QLabel("")
        self.preview_label.setStyleSheet("font-weight: bold; color: #0066cc;")
        preview_layout.addWidget(self.preview_label)
        layout.addLayout(preview_layout)
      
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
      
        self.setLayout(layout)
        self.update_preview()
  
    def update_preview(self):
        rule = self.rule_edit.text()
        name, ext = os.path.splitext(self.file_name)
        if not rule:
            preview = self.file_name
        else:
            try:
                preview = rule.replace('*', name) + ext
                if hasattr(self.parent_window, 'generate_new_name'):
                    file_index = -1
                    for i, file_path in enumerate(self.parent_window.files):
                        if os.path.basename(file_path) == self.file_name:
                            file_index = i
                            break
                    if file_index != -1:
                        preview = self.parent_window.generate_new_name(self.file_path, file_index, custom_rule=rule)
            except Exception as e:
                preview = f"预览错误: {str(e)}"
        if has_illegal_chars(preview):
            self.preview_label.setStyleSheet("font-weight: bold; color: red;")
            preview += "  [包含无效字符]"
        else:
            self.preview_label.setStyleSheet("font-weight: bold; color: #0066cc;")
        self.preview_label.setText(preview)
  
    def get_rule(self):
        return self.rule_edit.text()
  
    def get_naming_help_text(self):
        return (
            "命名规则说明：\n\n"
            "* - 代表原文件名\n"
            "# - 代表序号\n"
            "$ - 代表随机数字\n"
            "? - 代表原文件名中的单个字符\n"
            "<*-n> - 原文件名去掉末尾n个字符\n"
            "<-n*> - 原文件名去掉开头n个字符\n"
            "\\n - 原文件名的第n个字符\n"
            "使用\\$ \\# \\? 可直接输出这些字符\n\n"          
            "日期时间格式化：\n"
            "  yyyy/yy - 年份 (四位/两位数字)   YYYY/YY - 年份 (中文表示)\n"
            "  mm/m  - 月份 (补零/不补零)      MM/M - 月份 (中文表示)\n"
            "  dd/d  - 日期 (补零/不补零)        DD/D - 日期 (中文长/短表示)\n"
            "  hh/h  - 小时 (24/12小时制)        HH/H - 小时 (中文表示)\n"
            "  w/ddd   - 星期 (阿拉伯数字)     W/DDD - 星期 (中文数字)\n"
            "  tt/t  - 分钟 (补零/不补零)          TT/T - 分钟 (中文数字，补/不补零)\n"
            "  ss    - 秒钟 (阿拉伯数字)\n"
            "  在<>中使用上述代码时，可以混合文字，例如：<星期W>。\n\n"
            "简化日期格式：\n"
            "  <->   - 2025-3-19\n"
            "  <-->  - 2025年3月19日\n"
            "  <|>   - 二〇二五年三月十九日\n"
            "  <||>  - 农历二月二十日\n"
            "  <:>   - 0557\n"
            "  <::>  - 055712（时分秒连写）\n"
            "  <-:>  - 2025-3-19 0444\n"
            "  <.>   - 20250319\n"
            "  <>    - 空标签，使用默认规则\n\n"
            "示例：\n"
            "File_001.jpg 使用 <*-4>_新 得到 File_新.jpg\n"
            "File_001.jpg 使用 \\1\\2\\3 得到 Fil.jpg\n"
            "File_001.jpg 使用 Doc_# 得到 Doc_01.jpg\n"
            "文件名可以使用多个规则组合\n\n"
            "-- by yumumao@medu.cc"
        )
  
    def show_help(self, help_text):
        help_dialog = HelpDialog(help_text, self)
        help_dialog.exec_()


# =============================================================================
# RenameConflictDialog：用于解决文件名冲突的对话框
# =============================================================================
class RenameConflictDialog(QDialog):
    def __init__(self, conflict_items, parent=None):
        super().__init__(parent)
        self.setWindowTitle("文件名冲突")
        self.setMinimumWidth(500)
        self.conflict_items = conflict_items
        self.new_names = {}
      
        layout = QVBoxLayout()
        info_label = QLabel("以下文件名存在冲突，请修改:")
        layout.addWidget(info_label)
      
        self.name_edits = {}
        for item in conflict_items:
            item_layout = QHBoxLayout()
            item_layout.addWidget(QLabel(f"{item['index']}. {item['original_name']} -> "))
            name_edit = QLineEdit(item['new_name'])
            name_edit.setMinimumWidth(300)
            name_edit.item_index = item['index']
            self.name_edits[item['index']] = name_edit
            name_edit.textChanged.connect(self.update_name)
            item_layout.addWidget(name_edit)
            layout.addLayout(item_layout)
      
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        self.setLayout(layout)
      
        for item in conflict_items:
            self.new_names[item['index']] = item['new_name']
  
    def update_name(self):
        sender = self.sender()
        if hasattr(sender, 'item_index'):
            self.new_names[sender.item_index] = sender.text()
  
    def get_new_names(self):
        return self.new_names


# =============================================================================
# EnhancedFileDialog：支持同时选择文件和文件夹（多选）的对话框
# =============================================================================
class EnhancedFileDialog(QDialog):
    def __init__(self, parent=None, last_folder=None):
        super().__init__(parent)
        self.setWindowTitle("选择文件和文件夹")
        self.setMinimumWidth(900)
        self.setMinimumHeight(600)
      
        self.selected_paths = []
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(6, 6, 6, 6)
        main_layout.setSpacing(4)
      
        address_layout = QHBoxLayout()
        address_layout.addWidget(QLabel("地址:"))
        self.address_bar = QLineEdit()
        self.address_bar.setPlaceholderText("输入文件夹路径...")
        self.address_bar.returnPressed.connect(self.navigate_to_address)
        address_layout.addWidget(self.address_bar)
        go_btn = QPushButton("转到")
        go_btn.clicked.connect(self.navigate_to_address)
        address_layout.addWidget(go_btn)
        main_layout.addLayout(address_layout)
      
        self.splitter = QSplitter(Qt.Horizontal)
      
        folder_widget = QWidget()
        folder_layout = QVBoxLayout(folder_widget)
        folder_layout.setContentsMargins(0, 0, 0, 0)
        folder_layout.setSpacing(2)
        folder_label = QLabel("文件夹导航:")
        folder_layout.addWidget(folder_label)
        self.tree_model = QFileSystemModel()
        self.tree_model.setRootPath("")
        self.tree_model.setFilter(QDir.AllDirs | QDir.Drives | QDir.NoDotAndDotDot)
        self.tree_view = QTreeView()
        self.tree_view.setModel(self.tree_model)
        self.tree_view.setHeaderHidden(True)
        self.tree_view.hideColumn(1)
        self.tree_view.hideColumn(2)
        self.tree_view.hideColumn(3)
        folder_layout.addWidget(self.tree_view)
      
        file_widget = QWidget()
        file_layout = QVBoxLayout(file_widget)
        file_layout.setContentsMargins(0, 0, 0, 0)
        file_layout.setSpacing(2)
        file_header = QHBoxLayout()
        file_label = QLabel("文件和文件夹 (支持多选):")
        file_header.addWidget(file_label)
        self.path_label = QLabel("")
        self.path_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.path_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        file_header.addWidget(self.path_label)
        file_layout.addLayout(file_header)
        self.list_model = QFileSystemModel()
        self.list_model.setRootPath("")
        self.list_model.setFilter(QDir.AllEntries | QDir.NoDotAndDotDot)
        self.list_view = QTreeView()
        self.list_view.setModel(self.list_model)
        initial_path = last_folder if last_folder and os.path.exists(last_folder) else QDir.rootPath()
        self.list_view.setRootIndex(self.list_model.index(initial_path))
        self.path_label.setText(initial_path)
        self.address_bar.setText(initial_path)
        self.list_view.setSelectionMode(QTreeView.ExtendedSelection)
        self.list_view.setSortingEnabled(True)
        self.list_view.sortByColumn(0, Qt.AscendingOrder)
        self.list_view.header().setStretchLastSection(False)
        self.list_view.header().setSectionResizeMode(0, QHeaderView.Stretch)
        file_layout.addWidget(self.list_view)
      
        self.splitter.addWidget(folder_widget)
        self.splitter.addWidget(file_widget)
        self.splitter.setSizes([250, 650])
        main_layout.addWidget(self.splitter)
      
        button_layout = QHBoxLayout()
        tip_label = QLabel("提示: 双击文件夹进入，单击选中项，Shift/Ctrl可多选")
        tip_label.setStyleSheet("color: #666;")
        button_layout.addWidget(tip_label)
        button_layout.addStretch()
        self.add_selected_btn = QPushButton("添加选中项")
        self.add_selected_btn.clicked.connect(self.add_selected_items)
        button_layout.addWidget(self.add_selected_btn)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        main_layout.addLayout(button_layout)
      
        self.setLayout(main_layout)
      
        self.tree_view.clicked.connect(self.on_folder_clicked)
        self.list_view.doubleClicked.connect(self.on_list_double_clicked)
      
        if initial_path != QDir.rootPath():
            folder_index = self.tree_model.index(initial_path)
            if folder_index.isValid():
                self.tree_view.setCurrentIndex(folder_index)
  
    def navigate_to_address(self):
        path = self.address_bar.text().strip()
        if os.path.exists(path):
            if os.path.isdir(path):
                folder_index = self.tree_model.index(path)
                if folder_index.isValid():
                    self.tree_view.setCurrentIndex(folder_index)
                    self.on_folder_clicked(folder_index)
            else:
                dir_path = os.path.dirname(path)
                folder_index = self.tree_model.index(dir_path)
                if folder_index.isValid():
                    self.tree_view.setCurrentIndex(folder_index)
                    self.on_folder_clicked(folder_index)
        else:
            QMessageBox.warning(self, "路径错误", f"路径不存在: {path}")
  
    def on_folder_clicked(self, index):
        path = self.tree_model.filePath(index)
        self.list_view.setRootIndex(self.list_model.index(path))
        self.path_label.setText(path)
        self.address_bar.setText(path)
  
    def on_list_double_clicked(self, index):
        path = self.list_model.filePath(index)
        if os.path.isdir(path):
            folder_index = self.tree_model.index(path)
            if folder_index.isValid():
                self.tree_view.setCurrentIndex(folder_index)
                self.on_folder_clicked(folder_index)
        else:
            self.add_path_to_selection(path)
            self.accept()
  
    def add_selected_items(self):
        selected_indexes = self.list_view.selectedIndexes()
        for index in selected_indexes:
            if index.column() == 0:
                path = self.list_model.filePath(index)
                self.add_path_to_selection(path)
        if self.selected_paths:
            self.accept()
  
    def add_path_to_selection(self, path):
        if path not in self.selected_paths:
            self.selected_paths.append(path)
  
    def get_selected_paths(self):
        return self.selected_paths
  
    def get_current_directory(self):
        current_index = self.list_view.rootIndex()
        return self.list_model.filePath(current_index)


# =============================================================================
# CustomListItem：自定义列表项组件（显示序号和文件名）
# =============================================================================
class CustomListItem(QWidget):
    def __init__(self, index, filename, is_folder=False, has_single_rule=False, is_manually_renamed=False, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.index_label = QLabel(str(index))
        self.index_label.setFixedWidth(30)
        self.index_label.setAlignment(Qt.AlignVCenter | Qt.AlignRight)
        self.index_label.setStyleSheet("""
            background-color: #e6e6e6;
            padding: 2px 1px 0px 2px;
            color: #696969;
            font-size: 9pt;
            font-weight: bold;
            font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
            border-right: 1px solid #cccccc;
        """)
        self.has_single_rule = has_single_rule
        self.is_manually_renamed = is_manually_renamed
        self.is_folder = is_folder
        self.original_filename = filename
        self.current_filename = filename
        display_name = filename
        if is_folder:
            display_name = f"[文件夹] {filename}"
        if has_single_rule:
            display_name = f"({display_name})"
        self.filename_label = QLabel(display_name)
        font_style = "font-weight: bold;" if is_manually_renamed or has_single_rule else ""
        self.filename_label.setStyleSheet(f"padding: 2px; font-size: 12pt; {font_style}")
        layout.addWidget(self.index_label)
        layout.addWidget(self.filename_label)
        self.setFixedHeight(26)
  
    def get_filename(self):
        return self.filename_label.text()
  
    def get_original_filename(self):
        return self.current_filename
  
    def set_filename(self, filename):
        self.current_filename = filename
        display_name = filename
        if self.is_folder:
            display_name = f"[文件夹] {filename}"
        if self.has_single_rule:
            display_name = f"({display_name})"
        self.filename_label.setText(display_name)
  
    def set_single_rule(self, has_rule):
        self.has_single_rule = has_rule
        font_style = "font-weight: bold;" if has_rule or self.is_manually_renamed else ""
        self.filename_label.setStyleSheet(f"padding: 2px; font-size: 12pt; {font_style}")
        self.set_filename(self.current_filename)
      
    def set_manually_renamed(self, is_renamed):
        self.is_manually_renamed = is_renamed
        font_style = "font-weight: bold;" if is_renamed or self.has_single_rule else ""
        self.filename_label.setStyleSheet(f"padding: 2px; font-size: 12pt; {font_style}")
        self.set_filename(self.current_filename)


# =============================================================================
# FileListWidget：继承 QListWidget，实现文件列表显示和操作，
# 增加删除功能支持（右键和删除键）
# =============================================================================
class FileListWidget(QListWidget):
    def __init__(self, parent=None, show_empty_tip=True, is_preview_list=False):
        super().__init__(parent)
        self.setAlternatingRowColors(True)
        self.setViewMode(QListWidget.ListMode)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.is_preview_list = is_preview_list
        self.manual_renamed_items = {}  # {file_path: new_name}
        self.editing_enabled = False
      
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        self.itemDoubleClicked.connect(self.on_double_click)
        self.setStyleSheet("""
            QListWidget::item {
                height: 26px;
                padding: 0;
                margin: 0;
            }
        """)
        self.installEventFilter(self)
        self.show_empty_tip = show_empty_tip
      
        if show_empty_tip:
            self.empty_tip = QLabel("可将需改名的文件/文件夹拖动至此处", self)
            self.empty_tip.setAlignment(Qt.AlignCenter)
            self.empty_tip.setStyleSheet("""
                font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
                font-size: 12pt;
                font-weight: bold;
                color: #999999;
            """)
            self.empty_tip.hide()
        else:
            self.empty_tip = QLabel("文件/文件夹将按此处显示进行重命名", self)
            self.empty_tip.setAlignment(Qt.AlignCenter)
            self.empty_tip.setStyleSheet("""
                font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
                font-size: 12pt;
                font-weight: bold;
                color: #999999;
            """)
            self.empty_tip.hide()
      
        self.itemSelectionChanged.connect(self.on_selection_changed)
  
    def on_selection_changed(self):
        selected_items = self.selectedItems()
        if not selected_items:
            return
        main_window = self.find_main_window()
        if main_window:
            row = self.row(selected_items[0])
            main_window.sync_selection(self, row)
  
    def find_main_window(self):
        parent = self.parent()
        while parent and not isinstance(parent, FileRenamerApp):
            parent = parent.parent()
        return parent
      
    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'empty_tip'):
            self.empty_tip.resize(self.width(), self.height())
  
    def showEvent(self, event):
        super().showEvent(event)
        self.update_empty_tip()
  
    def update_empty_tip(self):
        if not hasattr(self, 'empty_tip'):
            return
        if self.count() == 0:
            self.empty_tip.show()
        else:
            self.empty_tip.hide()
  
    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress and event.key() == Qt.Key_Delete:
            self.delete_selected()
            return True
        return super().eventFilter(obj, event)
  
    def delete_selected(self):
        main_window = self.find_main_window()
        selected_items = self.selectedItems()
        if not selected_items:
            return
        rows = sorted(set(self.row(item) for item in selected_items))
        if len(rows) == 1:
            main_window.delete_file(rows[0])
        else:
            main_window.delete_files(rows)
  
    def add_file_item(self, index, filename, is_folder=False, has_single_rule=False, is_manually_renamed=False):
        item = QListWidgetItem(self)
        custom_widget = CustomListItem(index, filename, is_folder, has_single_rule, is_manually_renamed)
        item.setSizeHint(custom_widget.sizeHint())
        self.addItem(item)
        self.setItemWidget(item, custom_widget)
        self.update_empty_tip()
  
    def clear(self):
        super().clear()
        self.update_empty_tip()
  
    def get_item_filename(self, item):
        widget = self.itemWidget(item)
        if widget:
            return widget.get_filename()
        return ""
  
    def get_item_original_filename(self, item):
        widget = self.itemWidget(item)
        if widget:
            return widget.get_original_filename()
        return ""
  
    def set_item_filename(self, item, filename):
        widget = self.itemWidget(item)
        if widget:
            widget.set_filename(filename)
  
    def set_item_single_rule(self, item, has_rule):
        widget = self.itemWidget(item)
        if widget:
            widget.set_single_rule(has_rule)
  
    def set_item_manually_renamed(self, item, is_renamed):
        widget = self.itemWidget(item)
        if widget:
            widget.set_manually_renamed(is_renamed)
  
    def on_double_click(self, item):
        if self.is_preview_list and self.editing_enabled:
            row = self.row(item)
            main_window = self.find_main_window()
            main_window.preview_manual_rename(row)
        elif not self.is_preview_list:
            row = self.row(item)
            main_window = self.find_main_window()
            if main_window and row < len(main_window.files):
                file_path = main_window.files[row]
                if file_path in self.manual_renamed_items:
                    QMessageBox.information(self, "提示", 
                        "此文件已在预览结果中手动重命名，请先撤销手动重命名再设置单条规则。")
                    return
                main_window.edit_single_rule(row, file_path)
  
    def show_context_menu(self, position):
        menu = QMenu()
        selected_items = self.selectedItems()
        item_at_pos = self.itemAt(position)
        main_window = self.find_main_window()
        if not main_window:
            return
        if not self.is_preview_list:
            has_single_rules = bool(getattr(main_window, 'single_rules', {}))
            if not item_at_pos:
                if has_single_rules:
                    clear_rules_action = QAction("清除所有单条规则", self)
                    clear_rules_action.triggered.connect(main_window.clear_all_single_rules)
                    menu.addAction(clear_rules_action)
                clear_action = QAction("清空列表", self)
                clear_action.triggered.connect(lambda: main_window.clear_files())
                menu.addAction(clear_action)
            elif selected_items:
                if len(selected_items) > 1:
                    delete_action = QAction(f"删除选中的 {len(selected_items)} 项", self)
                    delete_action.triggered.connect(self.delete_selected)
                    menu.addAction(delete_action)
                else:
                    item = selected_items[0]
                    row = self.row(item)
                    single_rule_action = QAction("单条规则", self)
                    file_path = main_window.files[row]
                    if file_path in self.manual_renamed_items:
                        single_rule_action.setEnabled(False)
                        single_rule_action.setToolTip("此文件已在预览结果中手动重命名，请先撤销手动重命名")
                    else:
                        single_rule_action.triggered.connect(lambda: main_window.edit_single_rule(row, file_path))
                    menu.addAction(single_rule_action)
                    if file_path in main_window.single_rules:
                        clear_rule_action = QAction("清除单条规则", self)
                        clear_rule_action.triggered.connect(lambda: main_window.clear_single_rule(file_path))
                        menu.addAction(clear_rule_action)
                    delete_action = QAction("删除该条", self)
                    delete_action.triggered.connect(self.delete_selected)
                    menu.addAction(delete_action)
                if has_single_rules:
                    menu.addSeparator()
                    clear_rules_action = QAction("清除所有单条规则", self)
                    clear_rules_action.triggered.connect(main_window.clear_all_single_rules)
                    menu.addAction(clear_rules_action)
                menu.addSeparator()
                clear_action = QAction("清空列表", self)
                clear_action.triggered.connect(lambda: main_window.clear_files())
                menu.addAction(clear_action)
        else:
            if not item_at_pos or not selected_items:
                if self.manual_renamed_items:
                    reset_action = QAction("撤销所有手动命名", self)
                    reset_action.triggered.connect(main_window.reset_all_manual_renamed)
                    menu.addAction(reset_action)
                clear_action = QAction("清空列表", self)
                clear_action.triggered.connect(lambda: main_window.clear_files())
                menu.addAction(clear_action)
            elif selected_items:
                item = selected_items[0]
                row = self.row(item)
                if self.editing_enabled:
                    rename_action = QAction("手动重命名", self)
                    rename_action.triggered.connect(lambda: main_window.preview_manual_rename(row))
                    menu.addAction(rename_action)
                file_path = main_window.files[row]
                if file_path in self.manual_renamed_items:
                    reset_action = QAction("撤销手动命名", self)
                    reset_action.triggered.connect(lambda: main_window.reset_manual_renamed(row))
                    menu.addAction(reset_action)
                delete_action = QAction("删除该条", self)
                delete_action.triggered.connect(self.delete_selected)
                menu.addAction(delete_action)
                if self.manual_renamed_items:
                    menu.addSeparator()
                    reset_all_action = QAction("撤销所有手动命名", self)
                    reset_all_action.triggered.connect(main_window.reset_all_manual_renamed)
                    menu.addAction(reset_all_action)
                menu.addSeparator()
                clear_action = QAction("清空列表", self)
                clear_action.triggered.connect(lambda: main_window.clear_files())
                menu.addAction(clear_action)
        if not menu.isEmpty():
            menu.exec_(self.mapToGlobal(position))
  
    def set_editing_enabled(self, enabled):
        self.editing_enabled = enabled
  
    def reset_manual_renamed_items(self):
        for i in range(self.count()):
            item = self.item(i)
            if item:
                self.set_item_manually_renamed(item, False)
        self.manual_renamed_items = {}

    # 新增方法：清除所有同步选中产生的特殊背景色
    def clear_sync_highlight(self):
        for i in range(self.count()):
            item = self.item(i)
            item.setBackground(QBrush())

    # 重写鼠标与键盘事件，清除同步选中的高亮
    def mousePressEvent(self, event):
        self.clear_sync_highlight()
        super().mousePressEvent(event)

    def keyPressEvent(self, event):
        self.clear_sync_highlight()
        super().keyPressEvent(event)


# =============================================================================
# FileRenamerApp：主窗口，核心实现（选择、预览、执行、撤销、回退、冲突处理等）
# =============================================================================
class FileRenamerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("批量文件重命名工具")
        self.setGeometry(100, 100, 1000, 650)
        self.setAcceptDrops(True)
      
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "rename.ico")
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
        except:
            pass
      
        self.files = []  # 存储文件路径列表
        self.random_strings = {}
        self.is_splitter_moving = False
        self.last_folder_path = None
        self.default_empty_tag = "."
        self.has_previewed = False
        self.single_rules = {}  # {file_path: rule}
        self.last_rename_operations = []
        self.last_rename_before_files = []  # 上一次命名前的文件全路径（旧名称）
        self.last_rename_after_files = []   # 上一次命名后的文件全路径（当前名称）
      
        self.cn_num = {
            '0': '〇', '1': '一', '2': '二', '3': '三', '4': '四',
            '5': '五', '6': '六', '7': '七', '8': '八', '9': '九',
            '10': '十', '20': '廿', '30': '卅'
        }
        self.weekday_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
        self.lunar_month_names = ['正', '二', '三', '四', '五', '六', '七', '八', '九', '十', '冬', '腊']
      
        self.init_ui()
      
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
      
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setHandleWidth(6)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.splitterMoved.connect(self.on_splitter_moved)
        self.installEventFilter(self)
      
        # 左侧面板：重命名设置
        left_panel = QWidget()
        left_panel.setMinimumWidth(300)
        left_layout = QVBoxLayout(left_panel)
      
        naming_group = QGroupBox("基本命名规则")
        naming_layout = QVBoxLayout()
        pattern_layout = QHBoxLayout()
        pattern_layout.addWidget(QLabel("命名规则:"))
        self.pattern_edit = QLineEdit()
        self.pattern_edit.setPlaceholderText("使用 * 替代原文件名，详见帮助")
        self.pattern_edit.setText("<>#*")
        pattern_layout.addWidget(self.pattern_edit)
        help_btn = QToolButton()
        help_btn.setText("?")
        help_btn.setToolTip(self.get_naming_help_text())
        pattern_layout.addWidget(help_btn)
        naming_layout.addLayout(pattern_layout)
        self.help_label = QLabel("提示: * 原文件名, # 序号, $ 随机数, ? 单字符, <*-n> 去尾, <-n*> 去头, \\n 第n个字符, <-> 日期")
        smaller_font = self.help_label.font()
        smaller_font.setPointSize(smaller_font.pointSize() - 1)
        self.help_label.setFont(smaller_font)
        self.help_label.setTextFormat(Qt.PlainText)
        self.help_label.setWordWrap(False)
        self.help_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.help_label.setToolTip(self.get_naming_help_text())
        self.help_label.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        self.help_label.setMinimumWidth(100)
        naming_layout.addWidget(self.help_label)
        default_tag_layout = QHBoxLayout()
        default_tag_layout.addWidget(QLabel("空<>默认值:"))
        self.default_tag_edit = QLineEdit(self.default_empty_tag)
        self.default_tag_edit.setMinimumWidth(120)
        self.default_tag_edit.setToolTip("空<>标签的默认规则，例如设置为 'yyyy年mm月dd日' 或 '星期W' 等")
        self.default_tag_edit.textChanged.connect(self.update_default_tag_preview)
        default_tag_layout.addWidget(self.default_tag_edit)
        self.default_tag_preview = QLabel("预览: 20250319")
        self.default_tag_preview.setStyleSheet("color: #666;")
        default_tag_layout.addWidget(self.default_tag_preview)
        default_tag_layout.addStretch()
        naming_layout.addLayout(default_tag_layout)
        ext_layout = QHBoxLayout()
        ext_layout.addWidget(QLabel("扩展名:"))
        self.ext_edit = QLineEdit()
        self.ext_edit.setPlaceholderText("留空表示保持原扩展名, 也可用 *, #, $, ? 规则")
        ext_layout.addWidget(self.ext_edit)
        naming_layout.addLayout(ext_layout)
        naming_group.setLayout(naming_layout)
        left_layout.addWidget(naming_group)
      
        sequence_group = QGroupBox("序号设置")
        sequence_layout = QGridLayout()
        sequence_layout.addWidget(QLabel("起始序号:"), 0, 0)
        self.start_num = QSpinBox()
        self.start_num.setRange(0, 99999)
        self.start_num.setValue(1)
        sequence_layout.addWidget(self.start_num, 0, 1)
        sequence_layout.addWidget(QLabel("序号增量:"), 1, 0)
        self.increment_num = QSpinBox()
        self.increment_num.setRange(1, 999)
        self.increment_num.setValue(1)
        sequence_layout.addWidget(self.increment_num, 1, 1)
        sequence_layout.addWidget(QLabel("序号位数:"), 2, 0)
        self.digits_num = QSpinBox()
        self.digits_num.setRange(1, 10)
        self.digits_num.setValue(2)
        sequence_layout.addWidget(self.digits_num, 2, 1)
        sequence_layout.addWidget(QLabel("序号类型:"), 3, 0)
        seq_type_layout = QHBoxLayout()
        self.seq_type_group = QButtonGroup()
        self.number_radio = QRadioButton("数字")
        self.number_radio.setChecked(True)
        self.seq_type_group.addButton(self.number_radio)
        seq_type_layout.addWidget(self.number_radio)
        self.letter_radio = QRadioButton("字母")
        self.seq_type_group.addButton(self.letter_radio)
        seq_type_layout.addWidget(self.letter_radio)
        self.mixed_radio = QRadioButton("混合")
        self.seq_type_group.addButton(self.mixed_radio)
        seq_type_layout.addWidget(self.mixed_radio)
        self.pad_zeros = QCheckBox("补零")
        self.pad_zeros.setChecked(True)
        seq_type_layout.addWidget(self.pad_zeros)
        seq_type_widget = QWidget()
        seq_type_widget.setLayout(seq_type_layout)
        sequence_layout.addWidget(seq_type_widget, 3, 1)
        sequence_layout.addWidget(QLabel("字母大小写:"), 4, 0)
        self.letter_case = QComboBox()
        self.letter_case.addItems(["小写字母", "大写字母"])
        sequence_layout.addWidget(self.letter_case, 4, 1)
        sequence_group.setLayout(sequence_layout)
        left_layout.addWidget(sequence_group)
      
        random_group = QGroupBox("随机数设置")
        random_layout = QGridLayout()
        random_layout.addWidget(QLabel("随机数位数:"), 0, 0)
        self.random_digits = QSpinBox()
        self.random_digits.setRange(1, 20)
        self.random_digits.setValue(7)
        random_layout.addWidget(self.random_digits, 0, 1)
        random_layout.addWidget(QLabel("随机数类型:"), 1, 0)
        self.random_type = QComboBox()
        self.random_type.addItems(["纯数字", "数字+小写字母", "数字+大写字母", "数字+大小写字母"])
        random_layout.addWidget(self.random_type, 1, 1)
        random_layout.addWidget(QLabel("相同随机数:"), 2, 0)
        same_random_layout = QHBoxLayout()
        self.same_random_from = QSpinBox()
        self.same_random_from.setRange(1, 9999)
        self.same_random_from.setValue(1)
        same_random_layout.addWidget(self.same_random_from)
        same_random_layout.addWidget(QLabel("到"))
        self.same_random_to = QSpinBox()
        self.same_random_to.setRange(1, 9999)
        self.same_random_to.setValue(1)
        same_random_layout.addWidget(self.same_random_to)
        same_random_widget = QWidget()
        same_random_widget.setLayout(same_random_layout)
        random_layout.addWidget(same_random_widget, 2, 1)
        random_layout.addWidget(QLabel("(设置为相同值则不使用此功能)"), 3, 0, 1, 2)
        random_group.setLayout(random_layout)
        left_layout.addWidget(random_group)
      
        advanced_group = QGroupBox("高级设置")
        advanced_layout = QVBoxLayout()
        replace_layout = QVBoxLayout()
        use_regex_layout = QHBoxLayout()
        self.use_regex = QCheckBox("启用正则表达式")
        self.use_regex.setToolTip("启用后，替换文本将使用正则表达式匹配")
        use_regex_layout.addWidget(self.use_regex)
        regex_help_btn = QToolButton()
        regex_help_btn.setText("?")
        regex_help_btn.setToolTip("点击查看正则表达式用法帮助")
        regex_help_btn.clicked.connect(self.show_regex_help)
        use_regex_layout.addWidget(regex_help_btn)
        use_regex_layout.addStretch()
        replace_layout.addLayout(use_regex_layout)
        replace_from_layout = QHBoxLayout()
        replace_from_layout.addWidget(QLabel("替换:"))
        self.replace_from = QLineEdit()
        self.replace_from.setPlaceholderText("要替换的文本(多个用/分隔)")
        replace_from_layout.addWidget(self.replace_from)
        replace_layout.addLayout(replace_from_layout)
        replace_to_layout = QHBoxLayout()
        replace_to_layout.addWidget(QLabel("为:"))
        self.replace_to = QLineEdit()
        self.replace_to.setPlaceholderText("替换后的文本(多个用/分隔，:表示同上一项)")
        replace_to_layout.addWidget(self.replace_to)
        replace_layout.addLayout(replace_to_layout)
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        replace_layout.addWidget(line)
        advanced_layout.addLayout(replace_layout)
        edit_layout = QHBoxLayout()
        edit_layout.addWidget(QLabel("从位置:"))
        self.position_spin = QSpinBox()
        self.position_spin.setRange(0, 999)
        edit_layout.addWidget(self.position_spin)
        self.operation_type = QComboBox()
        self.operation_type.addItems(["插入", "删除"])
        edit_layout.addWidget(self.operation_type)
        self.edit_text = QLineEdit()
        self.edit_text.setPlaceholderText("要插入的文本")
        edit_layout.addWidget(self.edit_text)
        self.delete_count = QSpinBox()
        self.delete_count.setRange(1, 999)
        self.delete_count.setVisible(False)
        edit_layout.addWidget(self.delete_count)
        conflict_layout = QHBoxLayout()
        conflict_layout.addWidget(QLabel("文件名冲突:"))
        self.auto_resolve_conflicts = QCheckBox("自动解决冲突")
        self.auto_resolve_conflicts.setChecked(True)
        conflict_layout.addWidget(self.auto_resolve_conflicts)
        self.operation_type.currentTextChanged.connect(self.toggle_edit_controls)
        help_btn.clicked.connect(self.show_naming_help)
        advanced_layout.addLayout(edit_layout)
        advanced_layout.addLayout(conflict_layout)
        advanced_group.setLayout(advanced_layout)
        left_layout.addWidget(advanced_group)
      
        button_layout = QHBoxLayout()
        preview_btn = QPushButton("预览重命名结果")
        preview_btn.setMinimumHeight(45)
        bigger_font = preview_btn.font()
        bigger_font.setPointSize(bigger_font.pointSize() + 1)
        bigger_font.setBold(True)
        preview_btn.setFont(bigger_font)
        preview_btn.clicked.connect(self.preview_rename)
        button_layout.addWidget(preview_btn)
        execute_btn = QPushButton("执行重命名")
        execute_btn.setMinimumHeight(45)
        execute_btn.setFont(bigger_font)
        execute_btn.clicked.connect(self.execute_rename)
        button_layout.addWidget(execute_btn)
        left_layout.addLayout(button_layout)
        left_layout.setContentsMargins(9, 9, 9, 0)
        watermark_label = QLabel("yumumao@medu.cc")
        watermark_label.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
        watermark_label.setStyleSheet("color: rgba(128, 128, 128, 100); margin-top: 0px; margin-bottom: 0px; padding: 0px;")
        watermark_label.setAlignment(Qt.AlignLeft | Qt.AlignBottom)
        watermark_label.setFixedHeight(15)
        left_layout.addWidget(watermark_label)
      
        # 右侧面板：原文件列表与预览结果
        right_panel = QWidget()
        right_panel.setMinimumWidth(300)
        right_layout = QVBoxLayout(right_panel)
        file_btns_layout = QHBoxLayout()
        add_files_btn = QPushButton("添加文件")
        add_files_btn.clicked.connect(self.add_files_dialog)
        file_btns_layout.addWidget(add_files_btn)
        add_folders_btn = QPushButton("添加文件/文件夹")
        add_folders_btn.clicked.connect(self.add_folders_and_files_dialog)
        file_btns_layout.addWidget(add_folders_btn)
        clear_btn = QPushButton("清空列表")
        clear_btn.clicked.connect(self.clear_files)
        file_btns_layout.addWidget(clear_btn)
        right_layout.addLayout(file_btns_layout)

        # 修改：调整控件布局，直接去除暗色模式按钮，同时将排序下拉框和上移、下移按钮统一固定高度30
        control_layout = QHBoxLayout()
        # 右侧放置排序相关控件
        sort_layout = QHBoxLayout()
        sort_layout.addWidget(QLabel("排序:"))
        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["名称升序", "名称降序", "日期升序", "日期降序", "大小升序", "大小降序"])
        self.sort_combo.setCurrentIndex(2)
        self.sort_combo.currentIndexChanged.connect(self.sort_files)
        self.sort_combo.setFixedHeight(26)
        sort_layout.addWidget(self.sort_combo)
        move_up_btn = QPushButton("上移")
        move_up_btn.clicked.connect(self.move_file_up)
        move_up_btn.setFixedHeight(26)
        sort_layout.addWidget(move_up_btn)
        move_down_btn = QPushButton("下移")
        move_down_btn.clicked.connect(self.move_file_down)
        move_down_btn.setFixedHeight(26)
        sort_layout.addWidget(move_down_btn)
        control_layout.addLayout(sort_layout)

        right_layout.addLayout(control_layout)

        list_header_layout = QHBoxLayout()
        list_header_layout.addWidget(QLabel("原文件:"))
        selection_tip = QLabel("(支持Shift/Ctrl多选)")
        selection_tip.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        smaller_font = selection_tip.font()
        smaller_font.setPointSize(smaller_font.pointSize() - 1)
        selection_tip.setFont(smaller_font)
        list_header_layout.addWidget(selection_tip)
        right_layout.addLayout(list_header_layout)
        self.file_list = FileListWidget(show_empty_tip=True, is_preview_list=False)
        right_layout.addWidget(self.file_list)
        preview_header_layout = QHBoxLayout()
        preview_header_layout.addWidget(QLabel("预览结果:"))
        self.revert_btn = QPushButton("回退重命名")
        self.revert_btn.setToolTip("先点击此按钮回退预览结果，再点击‘执行重命名’正式回退")
        self.revert_btn.setEnabled(False)
        self.revert_btn.setMaximumHeight(25)
        self.revert_btn.clicked.connect(self.revert_last_rename)
        preview_header_layout.addStretch()
        preview_header_layout.addWidget(self.revert_btn)
        right_layout.addLayout(preview_header_layout)
        self.preview_list = FileListWidget(show_empty_tip=False, is_preview_list=True)
        right_layout.addWidget(self.preview_list)
        self.splitter.addWidget(left_panel)
        self.splitter.addWidget(right_panel)
        total_width = 1000
        self.splitter.setSizes([int(total_width * 0.4), int(total_width * 0.6)])
        main_layout.addWidget(self.splitter)
        self.splitter.splitterMoved.connect(self.update_help_label_elide)
        self.update_help_label_elide()
        self.preview_list.set_editing_enabled(True)
        self.update_default_tag_preview()
  
    def on_splitter_moved(self, pos, index):
        try:
            self.is_splitter_moving = True
            self.update_help_label_elide()
        except Exception as e:
            print(f"分割器移动错误: {str(e)}")
  
    # 修改后的同步选择：当一个列表中直接点击选中某项时，在另一列表中对应项显示浅黄色背景
    def sync_selection(self, sender, row):
        try:
            if sender == self.file_list:
                if self.preview_list.count() > 0:
                    self.preview_list.clearSelection()
                    self.preview_list.clear_sync_highlight()
                    if 0 <= row < self.preview_list.count():
                        item = self.preview_list.item(row)
                        item.setSelected(False)
                        item.setBackground(QColor("#ffffa8"))  # 浅黄色背景
            elif sender == self.preview_list:
                self.file_list.clearSelection()
                self.file_list.clear_sync_highlight()
                if 0 <= row < self.file_list.count():
                    item = self.file_list.item(row)
                    item.setSelected(False)
                    item.setBackground(QColor("#ffffa8"))
        except Exception as e:
            print(f"同步选择时出错: {str(e)}")
  
    def show_regex_help(self):
        help_text = (
            "正则表达式简明指南：\n\n"
            "基本匹配：\n"
            "  . - 匹配任意单个字符\n"
            "  \\d - 匹配任意数字\n"
            "  \\w - 匹配任意字母、数字或下划线\n"
            "  \\s - 匹配任意空白字符\n\n"
            "数量限定：\n"
            "  * - 匹配前面的表达式0次或多次\n"
            "  + - 匹配前面的表达式1次或多次\n"
            "  ? - 匹配前面的表达式0次或1次\n"
            "  {n} - 匹配前面的表达式n次\n"
            "  {n,m} - 匹配前面的表达式n到m次\n\n"
            "特殊字符：\n"
            "  ^ - 匹配字符串开头\n"
            "  $ - 匹配字符串结尾\n"
            "  [] - 匹配括号内的任意一个字符\n"
            "  [^] - 匹配除括号内字符外的任意一个字符\n"
            "  | - 或操作\n"
            "  () - 分组\n\n"
            "替换中的反向引用：\n"
            "  \\1, \\2, ... - 引用第1,2,...个捕获组\n\n"
            "示例：\n"
            "  匹配：\\d+ 替换为：# 将所有连续数字替换为#\n"
            "  匹配：(\\w+)_(\\d+) 替换为：\\2_\\1 互换数字和单词的位置\n"
            "  匹配：(?i)abc 替换为：XYZ 不区分大小写地匹配abc\n"
        )
        help_dialog = HelpDialog(help_text, self)
        help_dialog.setWindowTitle("正则表达式帮助")
        help_dialog.exec_()
  
    def update_default_tag_preview(self):
        tag_content = self.default_tag_edit.text().strip()
        self.default_empty_tag = tag_content
        try:
            test_text = f"<{tag_content}>"
            processed_text = self.process_date_time_tags(test_text)
            self.default_tag_preview.setText(f"预览: {processed_text}")
        except Exception as e:
            self.default_tag_preview.setText(f"格式错误: {str(e)}")
  
    def is_valid_empty_tag(self):
        tag_content = self.default_empty_tag.strip()
        if not tag_content:
            return False
        test_text = f"<{tag_content}>"
        processed_text = self.process_date_time_tags(test_text)
        return processed_text != test_text
  
    def add_files_dialog(self):
        try:
            initial_dir = self.last_folder_path if self.last_folder_path else os.getcwd()
            file_paths, _ = QFileDialog.getOpenFileNames(self, "选择文件", initial_dir)
            if file_paths:
                self.last_folder_path = os.path.dirname(file_paths[0])
                self.add_files(file_paths)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"选择文件时发生错误: {str(e)}")
            traceback.print_exc()
  
    def add_folders_and_files_dialog(self):
        try:
            dialog = EnhancedFileDialog(self, self.last_folder_path)
            if dialog.exec_() == QDialog.Accepted:
                selected_paths = dialog.get_selected_paths()
                if selected_paths:
                    self.last_folder_path = dialog.get_current_directory()
                    self.add_files(selected_paths)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"选择文件/文件夹时发生错误: {str(e)}")
            traceback.print_exc()
  
    def delete_file(self, row):
        if 0 <= row < len(self.files):
            file_path = self.files.pop(row)
            if file_path in self.single_rules:
                del self.single_rules[file_path]
            self.update_file_list()
            if self.preview_list.count() > 0:
                self.preview_rename()
  
    def delete_files(self, rows):
        rows.sort(reverse=True)
        for row in rows:
            if 0 <= row < len(self.files):
                file_path = self.files.pop(row)
                if file_path in self.single_rules:
                    del self.single_rules[file_path]
        self.update_file_list()
        if self.preview_list.count() > 0:
            self.preview_rename()
  
    def eventFilter(self, watched, event):
        if event.type() == QEvent.MouseButtonRelease:
            if self.is_splitter_moving:
                self.is_splitter_moving = False
                self.check_splitter_snap_on_release()
        return super().eventFilter(watched, event)
  
    def check_splitter_snap_on_release(self):
        try:
            total_width = self.splitter.width()
            target = int(total_width * 0.4)
            current_sizes = self.splitter.sizes()
            if abs(current_sizes[0] - target) < total_width * 0.4:
                self.splitter.setSizes([target, total_width - target])
        except Exception as e:
            print(f"分割条吸附错误: {str(e)}")
  
    def update_help_label_elide(self):
        if not hasattr(self, 'help_label'):
            return
        width = self.help_label.width()
        if width <= 0:
            return
        font_metrics = self.help_label.fontMetrics()
        text = "提示: * 原文件名, # 序号, $ 随机数, ? 单字符, <*-n> 去尾, <-n*> 去头, \\n 第n个字符, <-> 日期"
        if font_metrics.horizontalAdvance(text) > width:
            elided_text = text
            for i in range(len(text) - 1, 0, -1):
                elided_text = text[:i] + "..."
                if font_metrics.horizontalAdvance(elided_text) < width:
                    break
            self.help_label.setText(elided_text)
        else:
            self.help_label.setText(text)
  
    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_help_label_elide()
  
    def get_naming_help_text(self):
        return (
            "命名规则说明：\n\n"
            "* - 代表原文件名\n"
            "# - 代表序号\n"
            "$ - 代表随机数字\n"
            "? - 代表原文件名中的单个字符\n"
            "<*-n> - 原文件名去掉末尾n个字符\n"
            "<-n*> - 原文件名去掉开头n个字符\n"
            "\\n - 原文件名的第n个字符\n"
            "使用\\$ \\# \\? 可直接输出这些字符\n\n"
            "日期时间格式化：\n"
            "  yyyy/yy   - 年份 (四位/两位数字) YYYY/YY - 年份 (中文表示)\n"
            "  mm/m      - 月份 (补零/不补零)   MM/M    - 月份 (中文表示)\n"
            "  dd/d      - 日期 (补零/不补零)   DD/D    - 日期 (中文长/短表示)\n"
            "  hh/h      - 小时 (24/12小时制)   HH/H   - 小时 (中文表示)\n"
            "  w/ddd     - 星期 (阿拉伯数字)    W/DDD  - 星期 (中文数字)\n"
            "  tt/t      - 分钟 (补零/不补零)   TT/T   - 分钟 (中文数字，补/不补零)\n"
            "  ss        - 秒钟 (阿拉伯数字)\n"
            "  在<>中可以混合文字和代码，例如：<星期W>，\n"
            "  其中文字将保留，而代码自动替换。\n\n"
            "简化日期格式：\n"
            "  <->    - 2025-3-19\n"
            "  <-->   - 2025年3月19日\n"
            "  <|>    - 二〇二五年三月十九日\n"
            "  <||>   - 农历二月二十日\n"
            "  <:>    - 0557\n"
            "  <::>   - 055712（时分秒连写）\n"
            "  <-:>   - 2025-3-19 0444\n"
            "  <.>    - 20250319\n"
            "  <>     - 空标签，使用默认规则\n\n"
            "示例：\n"
            "File_001.jpg 使用 <*-4>_新 得到 File_新.jpg\n"
            "File_001.jpg 使用 \\1\\2\\3 得到 Fil.jpg\n"
            "File_001.jpg 使用 Doc_# 得到 Doc_01.jpg\n"
            "文件名可以使用多个规则组合\n\n"
            "-- by yumumao@medu.cc"
        )
  
    def show_naming_help(self):
        help_dialog = HelpDialog(self.get_naming_help_text(), self)
        help_dialog.exec_()
  
    def move_file_up(self):
        current_row = self.file_list.currentRow()
        if current_row > 0:
            current_item_widget = self.file_list.itemWidget(self.file_list.item(current_row))
            if current_item_widget:
                current_item_text = current_item_widget.get_filename()
                file_path = self.files.pop(current_row)
                self.files.insert(current_row - 1, file_path)
                self.update_file_list()
                for i in range(self.file_list.count()):
                    item_widget = self.file_list.itemWidget(self.file_list.item(i))
                    if item_widget and item_widget.get_filename() == current_item_text:
                        self.file_list.setCurrentRow(i)
                        break
  
    def move_file_down(self):
        current_row = self.file_list.currentRow()
        if current_row < self.file_list.count() - 1 and current_row >= 0:
            current_item_widget = self.file_list.itemWidget(self.file_list.item(current_row))
            if current_item_widget:
                current_item_text = current_item_widget.get_filename()
                file_path = self.files.pop(current_row)
                self.files.insert(current_row + 1, file_path)
                self.update_file_list()
                for i in range(self.file_list.count()):
                    item_widget = self.file_list.itemWidget(self.file_list.item(i))
                    if item_widget and item_widget.get_filename() == current_item_text:
                        self.file_list.setCurrentRow(i)
                        break
  
    def toggle_edit_controls(self, text):
        if text == "插入":
            self.edit_text.setVisible(True)
            self.delete_count.setVisible(False)
            self.edit_text.setPlaceholderText("要插入的文本")
        else:
            self.edit_text.setVisible(False)
            self.delete_count.setVisible(True)
  
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
  
    def dropEvent(self, event: QDropEvent):
        paths = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path) or os.path.isdir(path):
                paths.append(path)
        if paths:
            self.add_files(paths)
  
    def add_files(self, file_paths=None):
        if isinstance(file_paths, list):
            for file_path in file_paths:
                if file_path not in self.files and (os.path.isfile(file_path) or os.path.isdir(file_path)):
                    self.files.append(file_path)
            self.update_file_list()
            self.sort_files()
  
    def clear_files(self):
        self.files = []
        self.update_file_list()
        self.preview_list.clear()
        self.random_strings = {}
        self.has_previewed = False
        self.preview_list.reset_manual_renamed_items()
        self.single_rules = {}
  
    def update_file_list(self):
        self.file_list.clear()
        for i, file_path in enumerate(self.files):
            is_folder = os.path.isdir(file_path)
            has_single_rule = file_path in self.single_rules
            self.file_list.add_file_item(i + 1, os.path.basename(file_path), is_folder, has_single_rule)
  
    def sort_files(self):
        sort_option = self.sort_combo.currentText()
        file_info = []
        for file_path in self.files:
            try:
                stat = os.stat(file_path)
                file_info.append({
                    'path': file_path,
                    'name': os.path.basename(file_path),
                    'mtime': stat.st_mtime,
                    'size': stat.st_size,
                    'is_dir': os.path.isdir(file_path)
                })
            except (FileNotFoundError, PermissionError) as e:
                QMessageBox.warning(self, "文件访问错误", f"无法访问 {file_path}: {str(e)}")
        if sort_option == "名称升序":
            file_info.sort(key=lambda x: x['name'])
        elif sort_option == "名称降序":
            file_info.sort(key=lambda x: x['name'], reverse=True)
        elif sort_option == "日期升序":
            file_info.sort(key=lambda x: x['mtime'])
        elif sort_option == "日期降序":
            file_info.sort(key=lambda x: x['mtime'], reverse=True)
        elif sort_option == "大小升序":
            file_info.sort(key=lambda x: x['size'])
        elif sort_option == "大小降序":
            file_info.sort(key=lambda x: x['size'], reverse=True)
        self.files = [info['path'] for info in file_info]
        self.update_file_list()
  
    def generate_random_string(self, index):
        from_index = self.same_random_from.value()
        to_index = self.same_random_to.value()
        if from_index < to_index and from_index <= index <= to_index:
            key = f"group_{from_index}_{to_index}"
            if key not in self.random_strings:
                length = self.random_digits.value()
                random_type = self.random_type.currentText()
                chars = ""
                if random_type == "纯数字":
                    chars = string.digits
                elif random_type == "数字+小写字母":
                    chars = string.digits + string.ascii_lowercase
                elif random_type == "数字+大写字母":
                    chars = string.digits + string.ascii_uppercase
                else:
                    chars = string.digits + string.ascii_letters
                self.random_strings[key] = ''.join(random.choice(chars) for _ in range(length))
            return self.random_strings[key]
        else:
            key = f"single_{index}"
            if key not in self.random_strings:
                length = self.random_digits.value()
                random_type = self.random_type.currentText()
                chars = ""
                if random_type == "纯数字":
                    chars = string.digits
                elif random_type == "数字+小写字母":
                    chars = string.digits + string.ascii_lowercase
                elif random_type == "数字+大写字母":
                    chars = string.digits + string.ascii_uppercase
                else:
                    chars = string.digits + string.ascii_letters
                self.random_strings[key] = ''.join(random.choice(chars) for _ in range(length))
            return self.random_strings[key]
  
    def generate_sequence_string(self, index):
        seq_num = self.start_num.value() + (index - 1) * self.increment_num.value()
        seq_str = ""
        if self.number_radio.isChecked():
            digits = self.digits_num.value()
            if self.pad_zeros.isChecked():
                seq_str = str(seq_num).zfill(digits)
            else:
                seq_str = str(seq_num)
        elif self.letter_radio.isChecked():
            char_offset = ord('a') if self.letter_case.currentText() == "小写字母" else ord('A')
            digits = self.digits_num.value()
            seq_str = ""
            temp_num = seq_num
            while temp_num > 0 or not seq_str:
                temp_num -= 1
                char = chr(char_offset + (temp_num % 26))
                seq_str = char + seq_str
                temp_num //= 26
            if self.pad_zeros.isChecked() and len(seq_str) < digits:
                pad_char = chr(char_offset)
                seq_str = pad_char * (digits - len(seq_str)) + seq_str
        else:
            number_part = (seq_num - 1) // 26 + 1
            letter_part = (seq_num - 1) % 26
            char_offset = ord('a') if self.letter_case.currentText() == "小写字母" else ord('A')
            letter_char = chr(char_offset + letter_part)
            digits = self.digits_num.value() - 1
            if digits < 1:
                digits = 1
            if self.pad_zeros.isChecked():
                number_str = str(number_part).zfill(digits)
            else:
                number_str = str(number_part)
            seq_str = number_str + letter_char
        return seq_str
  
    def process_question_marks(self, pattern, original_name):
        question_mark_groups = []
        current_group = ''
        pattern = self.process_direct_chars(pattern)
        for char in pattern:
            if char == '?':
                current_group += '?'
            elif current_group:
                question_mark_groups.append(current_group)
                current_group = ''
        if current_group:
            question_mark_groups.append(current_group)
        result = pattern
        for group in question_mark_groups:
            count = len(group)
            if count <= len(original_name):
                replacement = original_name[:count]
                original_name = original_name[count:]
                result = result.replace(group, replacement, 1)
            else:
                replacement = original_name
                original_name = ''
                result = result.replace(group, replacement, 1)
        return result
  
    # -------------------------------------------------------------------------
    # 修改后的日期时间标签处理：
    # 1. 先调用 process_empty_tags 和 process_simplified_date_formats
    # 2. 用正则表达式把 <> 内的内容提取出来，再通过 format_datetime() 把代码替换上去，
    #    这样支持混合文字和代码的组合。
    # -------------------------------------------------------------------------
    def process_date_time_tags(self, text):
        now = datetime.datetime.now()
        text = self.process_empty_tags(text)
        text = self.process_simplified_date_formats(text)
        def replace_tag(match):
            content = match.group(1)
            return self.format_datetime(content, now)
        return re.sub(r'<([^>]+)>', replace_tag, text)
  
    def format_datetime(self, template, now):
        dt_codes = {
            "yyyy": lambda: now.strftime('%Y'),
            "yy": lambda: now.strftime('%y'),
            "YYYY": lambda: ''.join(self.cn_num[d] for d in now.strftime('%Y')),
            "YY": lambda: ''.join(self.cn_num[d] for d in now.strftime('%y')),
            "mm": lambda: f"{now.month:02d}",
            "m": lambda: str(now.month),
            "MM": lambda: self.get_chinese_month(now.month, True),
            "M": lambda: self.get_chinese_month(now.month, False),
            "dd": lambda: f"{now.day:02d}",
            "d": lambda: str(now.day),
            "DD": lambda: self.get_chinese_day(now.day, True),
            "D": lambda: self.get_chinese_day(now.day, False),
            "hh": lambda: f"{now.hour:02d}",
            "h": lambda: str(now.hour if now.hour <= 12 else now.hour - 12),
            "HH": lambda: self.get_chinese_number(now.hour),
            "H": lambda: self.get_chinese_number(now.hour if now.hour <= 12 else now.hour - 12),
            "w": lambda: str(now.weekday() + 1),
            "W": lambda: self.weekday_map[now.weekday()],
            "ddd": lambda: str(now.weekday() + 1),
            "DDD": lambda: self.weekday_map[now.weekday()],
            "tt": lambda: now.strftime('%M'),
            "t": lambda: str(now.minute),
            "TT": lambda: ''.join(self.cn_num[d] for d in f"{now.minute:02d}"),
            "T": lambda: ''.join(self.cn_num[d] for d in str(now.minute)),
            "ss": lambda: now.strftime('%S'),
        }
        # 为避免较短的代码影响较长代码的替换，先对代码按长度降序排序
        for code in sorted(dt_codes.keys(), key=len, reverse=True):
            template = template.replace(code, dt_codes[code]())
        return template
    # -------------------------------------------------------------------------
  
    def process_empty_tags(self, text):
        empty_tag_pattern = r'<>'
        if re.search(empty_tag_pattern, text):
            default_tag = self.default_empty_tag
            text = re.sub(empty_tag_pattern, f'<{default_tag}>', text)
        return text
  
    def process_simplified_date_formats(self, text):
        now = datetime.datetime.now()
        if '<-:>' in text:
            replacement = f"{now.year}-{now.month}-{now.day} {now.hour:02d}{now.minute:02d}"
            text = text.replace('<-:>', replacement)
        simplified_formats = {
            '<->': f"{now.year}-{now.month}-{now.day}",
            '<-->': f"{now.year}年{now.month}月{now.day}日",
            '<|>': f"{''.join(self.cn_num[d] for d in str(now.year))}年{self.get_chinese_month(now.month, False)}月{self.get_chinese_day(now.day, True)}日",
            '<||>': f"农历{self.lunar_month_names[self.get_approx_lunar_month()-1]}月{self.get_chinese_day(self.get_approx_lunar_day(), True)}",
            '<:>': f"{now.hour:02d}{now.minute:02d}",
            '<::>': f"{now.hour:02d}{now.minute:02d}{now.second:02d}",
            '<.>': f"{now.year}{now.month:02d}{now.day:02d}",
        }
        for pattern, replacement in simplified_formats.items():
            if pattern in text:
                text = text.replace(pattern, replacement)
        return text
  
    def get_approx_lunar_month(self):
        month = datetime.datetime.now().month - 1
        return month if month >= 1 else 12
  
    def get_approx_lunar_day(self):
        today = datetime.datetime.now()
        return (today.day + 15) % 30 or 30
  
    def get_chinese_month(self, month, long_format=True):
        if month <= 10:
            return self.cn_num[str(month)]
        elif month == 11:
            return "十一"
        else:
            return "十二"
  
    def get_chinese_day(self, day, long_format=True):
        if day <= 10:
            return "一十" if (long_format and day == 10) else self.cn_num[str(day)]
        elif day < 20:
            return ("一十" if long_format else "十") + self.cn_num[str(day)[1]]
        elif day == 20:
            return "二十" if long_format else "廿"
        elif day < 30:
            return ("二十" if long_format else "廿") + self.cn_num[str(day)[1]]
        else:
            return ("三十" if long_format else "卅") + ("" if day == 30 else self.cn_num["1"])
  
    def get_chinese_number(self, num):
        if num <= 10:
            return self.cn_num[str(num)]
        elif num < 20:
            return "十" + (self.cn_num[str(num)[1]] if num > 10 else "")
        else:
            return self.cn_num[str(num)[0]] + "十" + (self.cn_num[str(num)[1]] if num % 10 else "")
  
    def process_direct_chars(self, pattern):
        pattern = pattern.replace('\\$', '___DOLLAR___')
        pattern = pattern.replace('\\#', '___HASH___')
        pattern = pattern.replace('\\?', '___QUESTION___')
        return pattern
  
    def restore_direct_chars(self, text):
        text = text.replace('___DOLLAR___', '$')
        text = text.replace('___HASH___', '#')
        text = text.replace('___QUESTION___', '?')
        return text
  
    def process_truncate_pattern(self, pattern, original_name):
        truncate_end_matches = re.findall(r'<\*-(\d+)>', pattern)
        for match in truncate_end_matches:
            n = int(match)
            replacement = original_name[:-n] if n < len(original_name) else ""
            pattern = pattern.replace(f"<*-{match}>", replacement)
        truncate_start_matches = re.findall(r'<-(\d+)\*>', pattern)
        for match in truncate_start_matches:
            n = int(match)
            replacement = original_name[n:] if n < len(original_name) else ""
            pattern = pattern.replace(f"<-{match}*>", replacement)
        return pattern
  
    def process_specific_char_pattern(self, pattern, original_name):
        char_matches = re.findall(r'\\(\d+)', pattern)
        for match in char_matches:
            n = int(match)
            replacement = original_name[n-1] if 1 <= n <= len(original_name) else ""
            pattern = pattern.replace(f"\\{match}", replacement)
        return pattern
  
    def apply_multiple_replacements(self, text, replace_froms, replace_tos):
        use_regex = self.use_regex.isChecked()
        for i, replace_from in enumerate(replace_froms):
            if not replace_from:
                continue
            replace_to = replace_tos[i] if i < len(replace_tos) and replace_tos[i] != ':' else (replace_tos[i-1] if i > 0 else "")
            try:
                if use_regex:
                    text = re.sub(replace_from, replace_to, text)
                else:
                    text = text.replace(replace_from, replace_to)
            except Exception as e:
                error_msg = f"替换文本时出错: {str(e)}\n模式: '{replace_from}'，替换为: '{replace_to}'"
                print(error_msg)
                QMessageBox.warning(self, "替换错误", error_msg)
                continue
        return text
  
    def generate_new_name(self, file_path, index, custom_rule=None):
        base_name = os.path.basename(file_path)
        if os.path.isdir(file_path):
            name = base_name
            ext = ""
        else:
            name, ext = os.path.splitext(base_name)
        if custom_rule is not None:
            pattern = custom_rule
        elif file_path in self.single_rules:
            pattern = self.single_rules[file_path]
        else:
            pattern = self.pattern_edit.text()
      
        pattern = self.process_direct_chars(pattern)
        try:
            pattern = self.process_date_time_tags(pattern)
        except Exception as e:
            print(f"日期时间处理异常: {str(e)}")
        try:
            if "<*-" in pattern or ("<-" in pattern and "*>" in pattern):
                pattern = self.process_truncate_pattern(pattern, name)
        except Exception as e:
            print(f"截断模式处理异常: {str(e)}")
        try:
            if "\\" in pattern:
                pattern = self.process_specific_char_pattern(pattern, name)
        except Exception as e:
            print(f"特定字符处理异常: {str(e)}")
        try:
            if '?' in pattern:
                pattern = self.process_question_marks(pattern, name)
        except Exception as e:
            print(f"问号处理异常: {str(e)}")
        try:
            seq_str = self.generate_sequence_string(index + 1)
        except Exception as e:
            print(f"序号生成异常: {str(e)}")
            seq_str = str(index + 1)
        try:
            random_str = self.generate_random_string(index + 1)
        except Exception as e:
            print(f"随机字符串生成异常: {str(e)}")
            random_str = ''.join(random.choices(string.digits, k=4))
      
        try:
            new_name = pattern
            new_name = re.sub(r"(?<!\\)\*", name, new_name)
            new_name = re.sub(r"(?<!\\)#", lambda m: seq_str, new_name)
            new_name = re.sub(r"(?<!\\)\$", lambda m: random_str, new_name)
        except Exception as e:
            print(f"通配符替换异常: {str(e)}")
            new_name = name
        new_name = new_name.replace('\\*', '*')
        new_name = new_name.replace('\\#', '#')
        new_name = new_name.replace('\\$', '$')
        new_name = new_name.replace('\\?', '?')
        new_name = new_name.replace('___DOLLAR___', '$')
        new_name = new_name.replace('___HASH___', '#')
        new_name = new_name.replace('___QUESTION___', '?')
      
        if os.path.isdir(file_path):
            return new_name
        else:
            return new_name + ext
  
    def preview_rename(self):
        if not self.files:
            QMessageBox.warning(self, "警告", "请先添加文件!")
            return
        pattern = self.pattern_edit.text()
        if '<>' in pattern and not self.is_valid_empty_tag():
            QMessageBox.warning(
                self, 
                "空标签警告", 
                f"您的命名模式中包含空<>标签，但当前设置的默认值 '{self.default_empty_tag}' 不是有效的日期时间格式。\n\n"
                "请修改空<>默认值为有效格式，例如：\n- '.' 表示紧凑日期格式如 20250319\n- ':' 表示时间格式如 05:57\n- '::' 表示带秒的时间格式如 05:58:12\n- '-' 表示日期格式如 2025-3-19"
            )
            return
        self.random_strings = {}
        self.preview_list.clear()
        new_names = []
        for i, file_path in enumerate(self.files):
            try:
                is_folder = os.path.isdir(file_path)
                new_name = self.generate_new_name(file_path, i)
                new_names.append(new_name)
            except Exception as e:
                print(f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
                new_names.append(os.path.basename(file_path))
                continue
        for i, file_path in enumerate(self.files):
            if file_path in self.preview_list.manual_renamed_items:
                new_names[i] = self.preview_list.manual_renamed_items[file_path]
        for i, (file_path, new_name) in enumerate(zip(self.files, new_names)):
            is_folder = os.path.isdir(file_path)
            has_single_rule = file_path in self.single_rules
            is_manually_renamed = file_path in self.preview_list.manual_renamed_items
            self.preview_list.add_file_item(i + 1, new_name, is_folder, has_single_rule, is_manually_renamed)
        for i in range(self.preview_list.count()):
            item = self.preview_list.item(i)
            widget = self.preview_list.itemWidget(item)
            if widget and has_illegal_chars(new_names[i]):
                widget.filename_label.setStyleSheet("padding: 2px; font-size: 12pt; font-weight: bold; color: red;")
                widget.filename_label.setText(widget.current_filename + " [包含无效字符]")
        self.preview_list.set_editing_enabled(True)
        self.has_previewed = True
  
    def execute_rename(self):
        if not self.files:
            QMessageBox.warning(self, "警告", "请先添加文件!")
            return
        if not self.has_previewed or self.preview_list.count() == 0:
            QMessageBox.warning(self, "警告", "请先点击「预览重命名结果」按钮预览更改!")
            return
        new_names = []
        for i in range(self.preview_list.count()):
            item = self.preview_list.item(i)
            if item:
                custom_widget = self.preview_list.itemWidget(item)
                if custom_widget:
                    new_names.append(custom_widget.current_filename)
        illegal_found = any(has_illegal_chars(name) for name in new_names)
        if illegal_found:
            reply = QMessageBox.question(
                self, 
                "存在无效字符", 
                "预览结果中存在包含无效字符的文件名，是否返回修改？\n\n点击“是”返回修改，点击“否”忽略错误继续执行。",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                return
        has_conflicts = False
        conflict_items = []
        for i, new_name in enumerate(new_names):
            dir_path = os.path.dirname(self.files[i])
            conflicts = []
            for j, other_file in enumerate(self.files):
                if i != j:
                    other_dir = os.path.dirname(other_file)
                    other_new_name = new_names[j]
                    if dir_path == other_dir and new_name == other_new_name:
                        conflicts.append(j)
            if conflicts:
                has_conflicts = True
                conflict_items.append({
                    'index': i + 1,
                    'original_name': os.path.basename(self.files[i]),
                    'new_name': new_name,
                    'conflicts': conflicts
                })
        if has_conflicts:
            if self.auto_resolve_conflicts.isChecked():
                try:
                    new_names = self.auto_resolve_name_conflicts(new_names)
                    for i, name in enumerate(new_names):
                        if i < self.preview_list.count():
                            item = self.preview_list.item(i)
                            if item:
                                custom_widget = self.preview_list.itemWidget(item)
                                if custom_widget:
                                    custom_widget.set_filename(name)
                except Exception as e:
                    print(f"自动解决冲突时出错: {str(e)}")
            else:
                try:
                    conflict_dialog = RenameConflictDialog(conflict_items, self)
                    if conflict_dialog.exec_() == QDialog.Accepted:
                        new_name_dict = conflict_dialog.get_new_names()
                        for idx, name in new_name_dict.items():
                            new_names[idx - 1] = name
                            if idx - 1 < self.preview_list.count():
                                item = self.preview_list.item(idx - 1)
                                if item:
                                    custom_widget = self.preview_list.itemWidget(item)
                                    if custom_widget:
                                        custom_widget.set_filename(name)
                    else:
                        return
                except Exception as e:
                    print(f"手动解决冲突对话框出错: {str(e)}")
                    QMessageBox.warning(self, "冲突解决错误", 
                        f"处理文件名冲突时出错: {str(e)}\n请尝试启用自动解决冲突功能。")
                    return
        if QMessageBox.question(self, "确认重命名", 
            f"确定要重命名 {len(self.files)} 个文件/文件夹吗？", 
            QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        self.last_rename_before_files = self.files.copy()
        self.last_rename_after_files = []
        success_count = 0
        error_count = 0
        unchanged_count = 0
        error_messages = []
        self.last_rename_operations = []
        for i, file_path in enumerate(self.files):
            try:
                dir_path = os.path.dirname(file_path)
                new_name = new_names[i]
                new_path = os.path.join(dir_path, new_name)
                self.last_rename_after_files.append(new_path)
                is_folder = os.path.isdir(file_path)
                if os.path.basename(file_path) == new_name:
                    unchanged_count += 1
                    continue
                if file_path in self.single_rules:
                    del self.single_rules[file_path]
                if file_path in self.preview_list.manual_renamed_items:
                    del self.preview_list.manual_renamed_items[file_path]
                self.last_rename_operations.append({
                    'old_path': file_path,
                    'new_path': new_path,
                    'old_name': os.path.basename(file_path),
                    'new_name': new_name
                })
                os.rename(file_path, new_path)
                success_count += 1
                self.files[i] = new_path
            except Exception as e:
                error_count += 1
                item_type = "文件夹" if is_folder else "文件"
                error_messages.append(f"无法重命名{item_type} {os.path.basename(file_path)}: {str(e)}")
            # 若重命名异常，也继续处理后续文件
        self.update_file_list()
        self.revert_btn.setEnabled(True)
        self.preview_rename()
        result_dialog = QMessageBox(self)
        result_dialog.setWindowTitle("重命名结果")
        if error_count > 0:
            result_dialog.setIcon(QMessageBox.Warning)
            result_dialog.setText(f"成功修改: {success_count} 个文件/文件夹\n"
                                  f"未修改: {unchanged_count} 个文件/文件夹\n"
                                  f"失败: {error_count} 个文件/文件夹\n\n" +
                                  "\n".join(error_messages[:10]) +
                                  ("\n..." if len(error_messages) > 10 else ""))
        elif success_count > 0:
            result_dialog.setIcon(QMessageBox.Information)
            result_dialog.setText(f"成功修改: {success_count} 个文件/文件夹\n"
                                  f"未修改: {unchanged_count} 个文件/文件夹")
        else:
            result_dialog.setIcon(QMessageBox.Information)
            result_dialog.setText("所有文件名未发生变化，没有文件被修改")
        result_dialog.setStandardButtons(QMessageBox.Ok)
        if success_count > 0:
            undo_button = result_dialog.addButton("撤销重命名", QMessageBox.ActionRole)
            undo_button.clicked.connect(self.undo_rename)
        result_dialog.exec_()
  
    def undo_rename(self):
        if not self.last_rename_operations:
            QMessageBox.information(self, "撤销", "没有可撤销的重命名操作")
            return
        success_count = 0
        error_count = 0
        error_messages = []
        for op in reversed(self.last_rename_operations):
            try:
                if os.path.exists(op['new_path']):
                    for i, file_path in enumerate(self.files):
                        if file_path == op['new_path']:
                            self.files[i] = op['old_path']
                    os.rename(op['new_path'], op['old_path'])
                    success_count += 1
                else:
                    error_count += 1
                    error_messages.append(f"找不到文件: {op['new_path']}")
            except Exception as e:
                error_count += 1
                error_messages.append(f"无法撤销重命名 {op['new_name']} 到 {op['old_name']}: {str(e)}")
        self.update_file_list()
        self.preview_rename()
        if error_count > 0:
            QMessageBox.warning(self, "撤销结果", 
                f"成功撤销: {success_count} 个文件/文件夹\n"
                f"失败: {error_count} 个文件/文件夹\n\n" +
                "\n".join(error_messages[:10]) +
                ("\n..." if len(error_messages) > 10 else ""))
        else:
            QMessageBox.information(self, "撤销结果", f"成功撤销了 {success_count} 个文件/文件夹的重命名")
        self.last_rename_operations = []
  
    def revert_last_rename(self):
        if not self.last_rename_before_files or not self.last_rename_after_files:
            QMessageBox.information(self, "回退重命名", "没有可回退的重命名操作")
            return
        self.files = self.last_rename_after_files.copy()
        self.update_file_list()
        self.preview_list.clear()
        for i, file_path in enumerate(self.last_rename_before_files):
            is_folder = os.path.isdir(file_path)
            self.preview_list.add_file_item(i + 1, os.path.basename(file_path), is_folder)
        self.has_previewed = True
        QMessageBox.information(self, "回退重命名", 
            "已回退预览。\n文件列表显示当前文件名，预览结果显示上一次命名前的名称。\n请点击‘执行重命名’按钮正式回退。")
  
    def preview_manual_rename(self, row):
        if row < 0 or row >= len(self.files):
            return
        file_path = self.files[row]
        if file_path in self.single_rules:
            reply = QMessageBox.question(self, "提示",
                "该文件已设置单条规则，手动重命名将撤销该规则，是否继续？",
                QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
            else:
                del self.single_rules[file_path]
                self.update_file_list()
                self.preview_rename()
        current_text = self.preview_list.manual_renamed_items.get(file_path, None)
        if current_text is None:
            item = self.preview_list.item(row)
            if item:
                widget = self.preview_list.itemWidget(item)
                if widget:
                    current_text = widget.current_filename
                else:
                    current_text = os.path.basename(file_path)
            else:
                current_text = os.path.basename(file_path)
        dialog = RenameDialog(current_text, self)
        if dialog.exec_() != QDialog.Accepted:
            return
        new_text = dialog.get_new_name()
        if has_illegal_chars(new_text):
            QMessageBox.warning(self, "错误", "手动重命名不能包含无效字符，请修改！")
            return
        self.preview_list.manual_renamed_items[file_path] = new_text
        self.preview_rename()
  
    def edit_single_rule(self, row, file_path):
        if file_path in self.preview_list.manual_renamed_items:
            reply = QMessageBox.question(self, "提示",
                "该文件已在预览结果中手动重命名，是否先撤销手动重命名再设置单条规则？",
                QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                if file_path in self.preview_list.manual_renamed_items:
                    del self.preview_list.manual_renamed_items[file_path]
                    self.preview_rename()
            else:
                return
        default_rule = self.pattern_edit.text()
        current_rule = self.single_rules.get(file_path, default_rule)
        dialog = SingleRuleDialog(file_path, current_rule, self)
        if dialog.exec_() == QDialog.Accepted:
            rule = dialog.get_rule()
            if rule:
                self.single_rules[file_path] = rule
                self.update_file_list()
                if self.has_previewed:
                    self.preview_rename()
  
    def clear_single_rule(self, file_path):
        if file_path in self.single_rules:
            del self.single_rules[file_path]
            self.update_file_list()
            if self.has_previewed:
                self.preview_rename()
  
    def clear_all_single_rules(self):
        if not self.single_rules:
            return
        confirm = QMessageBox.question(self, "确认清除", "确定要清除所有单条规则吗？", QMessageBox.Yes | QMessageBox.No)
        if confirm == QMessageBox.Yes:
            self.single_rules = {}
            self.update_file_list()
            if self.has_previewed:
                self.preview_rename()
  
    def reset_manual_renamed(self, row):
        if hasattr(self, 'preview_list'):
            file_path = self.files[row]
            if file_path in self.preview_list.manual_renamed_items:
                del self.preview_list.manual_renamed_items[file_path]
                self.preview_rename()
  
    def reset_all_manual_renamed(self):
        if not hasattr(self, 'preview_list') or not self.preview_list.manual_renamed_items:
            return
        confirm = QMessageBox.question(self, "确认撤销", "确定要撤销所有手动重命名吗？", QMessageBox.Yes | QMessageBox.No)
        if confirm == QMessageBox.Yes:
            self.preview_list.manual_renamed_items = {}
            self.preview_rename()


# =============================================================================
# 全局异常处理器
# =============================================================================
def excepthook(exc_type, exc_value, exc_traceback):
    error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    print(error_msg)
    dialog = QMessageBox()
    dialog.setWindowTitle("错误")
    dialog.setText(f"程序发生错误: {str(exc_value)}")
    dialog.setDetailedText(error_msg)
    dialog.setIcon(QMessageBox.Critical)
    dialog.exec_()

# =============================================================================
# 程序入口
# =============================================================================
if __name__ == "__main__":
    sys.excepthook = excepthook
    app = QApplication(sys.argv)
    window = FileRenamerApp()
    window.show()
    sys.exit(app.exec_())