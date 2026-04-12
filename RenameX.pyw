"""
批量文件重命名工具
功能：支持多种命名规则（通配符、序号、随机数、日期时间、正则替换等），
     支持单条规则、手动重命名、冲突自动解决、撤销/回退等。
"""

import sys
import os
import re
import random
import string
import traceback
import datetime
import calendar
import math
import uuid
from functools import partial

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
# 辅助函数
# =============================================================================

def has_illegal_chars(name):
    """检测文件名中是否包含 Windows 不允许的字符"""
    illegal_chars = ['"', '*', '<', '>', '?', '\\', '/', '|', ':']
    return any(ch in name for ch in illegal_chars)


# =============================================================================
# RenameDialog：手动重命名对话框
# =============================================================================

class RenameDialog(QDialog):
    """用于在预览列表中手动输入新文件名的对话框"""

    def __init__(self, current_text, parent=None):
        super().__init__(parent)
        self.setWindowTitle("重命名")
        self.setMinimumWidth(500)

        layout = QVBoxLayout()
        label = QLabel("输入新名称:")
        layout.addWidget(label)

        self.name_edit = QLineEdit(current_text)
        self.name_edit.setMinimumWidth(480)
        self.name_edit.selectAll()
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
    """显示命名规则或正则表达式帮助文本的只读对话框"""

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
# SingleRuleDialog：单条规则编辑对话框（带实时预览）
# =============================================================================

class SingleRuleDialog(QDialog):
    """为单个文件设置独立命名规则的对话框，支持实时预览结果"""

    def __init__(self, file_path, default_rule, parent=None):
        super().__init__(parent)
        self.setWindowTitle("单条规则编辑")
        self.setMinimumWidth(600)
        self.setMinimumHeight(200)

        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.parent_window = parent

        layout = QVBoxLayout()

        # 当前文件名显示
        file_info = QLabel(f"为文件设置单独规则: {self.file_name}")
        file_info.setStyleSheet("font-weight: bold;")
        layout.addWidget(file_info)

        # 规则输入
        rule_layout = QHBoxLayout()
        rule_layout.addWidget(QLabel("单条规则:"))
        self.rule_edit = QLineEdit(default_rule)
        self.rule_edit.setPlaceholderText("使用 * 替代原文件名，详见帮助")
        self.rule_edit.textChanged.connect(self.update_preview)
        rule_layout.addWidget(self.rule_edit)

        help_btn = QToolButton()
        help_btn.setText("?")
        help_btn.clicked.connect(partial(self.show_help, self._get_help_text()))
        rule_layout.addWidget(help_btn)
        layout.addLayout(rule_layout)

        # 预览结果
        preview_layout = QHBoxLayout()
        preview_layout.addWidget(QLabel("预览结果:"))
        self.preview_label = QLabel("")
        self.preview_label.setStyleSheet("font-weight: bold; color: #0066cc;")
        preview_layout.addWidget(self.preview_label)
        layout.addLayout(preview_layout)

        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)
        self.update_preview()

    def update_preview(self):
        """根据当前规则实时更新预览结果"""
        rule = self.rule_edit.text()
        name, ext = os.path.splitext(self.file_name)
        if not rule:
            preview = self.file_name
        else:
            try:
                preview = rule.replace('*', name) + ext
                if hasattr(self.parent_window, 'generate_new_name'):
                    file_index = -1
                    for i, fp in enumerate(self.parent_window.files):
                        if os.path.basename(fp) == self.file_name:
                            file_index = i
                            break
                    if file_index != -1:
                        preview = self.parent_window.generate_new_name(
                            self.file_path, file_index, custom_rule=rule
                        )
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

    def _get_help_text(self):
        return (
            "命名规则说明：\n\n"
            "* - 代表原文件名\n"
            "# - 代表序号\n"
            "$ - 代表随机数字\n"
            "? - 代表原文件名中的单个字符\n"
            "<*-n> - 原文件名去掉末尾n个字符\n"
            "<-n*> - 原文件名去掉开头n个字符\n"
            "\\n - 原文件名的第n个字符\n"
            "使用\\$ \\# \\? 可直接输出这些字符\n"
        )

    def show_help(self, help_text):
        HelpDialog(help_text, self).exec_()


# =============================================================================
# RenameConflictDialog：文件名冲突解决对话框
# =============================================================================

class RenameConflictDialog(QDialog):
    """当存在同名冲突时，让用户手动修改冲突文件名"""

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
            name_edit.textChanged.connect(self._on_name_changed)
            item_layout.addWidget(name_edit)
            layout.addLayout(item_layout)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        self.setLayout(layout)

        for item in conflict_items:
            self.new_names[item['index']] = item['new_name']

    def _on_name_changed(self):
        sender = self.sender()
        if hasattr(sender, 'item_index'):
            self.new_names[sender.item_index] = sender.text()

    def get_new_names(self):
        return self.new_names


# =============================================================================
# EnhancedFileDialog：支持同时选择文件和文件夹的多选对话框
# =============================================================================

class EnhancedFileDialog(QDialog):
    """左侧文件夹树 + 右侧文件列表，可同时选择文件和文件夹"""

    def __init__(self, parent=None, last_folder=None):
        super().__init__(parent)
        self.setWindowTitle("选择文件和文件夹")
        self.setMinimumWidth(900)
        self.setMinimumHeight(600)

        self.selected_paths = []
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(6, 6, 6, 6)
        main_layout.setSpacing(4)

        # --- 地址栏 ---
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

        # --- 左右分栏 ---
        self.splitter = QSplitter(Qt.Horizontal)

        # 左侧：文件夹导航树
        folder_widget = QWidget()
        folder_layout = QVBoxLayout(folder_widget)
        folder_layout.setContentsMargins(0, 0, 0, 0)
        folder_layout.setSpacing(2)
        folder_layout.addWidget(QLabel("文件夹导航:"))
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

        # 右侧：文件和文件夹列表（多选）
        file_widget = QWidget()
        file_layout = QVBoxLayout(file_widget)
        file_layout.setContentsMargins(0, 0, 0, 0)
        file_layout.setSpacing(2)
        file_header = QHBoxLayout()
        file_header.addWidget(QLabel("文件和文件夹 (支持多选):"))
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

        # --- 底部按钮 ---
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

        # --- 信号连接 ---
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
            self._add_path(path)
            self.accept()

    def add_selected_items(self):
        selected_indexes = self.list_view.selectedIndexes()
        for index in selected_indexes:
            if index.column() == 0:
                path = self.list_model.filePath(index)
                self._add_path(path)
        if self.selected_paths:
            self.accept()

    def _add_path(self, path):
        if path not in self.selected_paths:
            self.selected_paths.append(path)

    def get_selected_paths(self):
        return self.selected_paths

    def get_current_directory(self):
        current_index = self.list_view.rootIndex()
        return self.list_model.filePath(current_index)


# =============================================================================
# CustomListItem：自定义列表项组件（序号 + 文件名）
# =============================================================================

class CustomListItem(QWidget):
    """列表中每一行的自定义控件，左侧序号，右侧文件名"""

    def __init__(self, index, filename, is_folder=False,
                 has_single_rule=False, is_manually_renamed=False, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 序号标签
        self.index_label = QLabel(str(index))
        self.index_label.setFixedWidth(30)
        self.index_label.setAlignment(Qt.AlignVCenter | Qt.AlignRight)
        self.index_label.setStyleSheet(
            "background-color: #e6e6e6; padding: 2px 1px 0px 2px; "
            "color: #696969; font-size: 9pt; font-weight: bold; "
            'font-family: "Microsoft YaHei", sans-serif; '
            "border-right: 1px solid #cccccc;"
        )

        # 状态标志
        self.has_single_rule = has_single_rule
        self.is_manually_renamed = is_manually_renamed
        self.is_folder = is_folder
        self.original_filename = filename
        self.current_filename = filename

        # 文件名标签
        display_name = self._build_display_name(filename)
        self.filename_label = QLabel(display_name)
        self._update_label_style()

        layout.addWidget(self.index_label)
        layout.addWidget(self.filename_label)
        self.setFixedHeight(26)

    def _build_display_name(self, filename):
        """根据文件夹/单条规则状态构造显示名"""
        display = filename
        if self.is_folder:
            display = f"[文件夹] {display}"
        if self.has_single_rule:
            display = f"({display})"
        return display

    def _update_label_style(self):
        """根据是否手动/单条规则设置字体样式"""
        bold = "font-weight: bold;" if (self.is_manually_renamed or self.has_single_rule) else ""
        self.filename_label.setStyleSheet(f"padding: 2px; font-size: 12pt; {bold}")

    def get_filename(self):
        return self.filename_label.text()

    def get_original_filename(self):
        return self.current_filename

    def set_filename(self, filename):
        self.current_filename = filename
        self.filename_label.setText(self._build_display_name(filename))

    def set_single_rule(self, has_rule):
        self.has_single_rule = has_rule
        self._update_label_style()
        self.set_filename(self.current_filename)

    def set_manually_renamed(self, is_renamed):
        self.is_manually_renamed = is_renamed
        self._update_label_style()
        self.set_filename(self.current_filename)


# =============================================================================
# FileListWidget：文件列表控件
# =============================================================================

class FileListWidget(QListWidget):

    def __init__(self, parent=None, show_empty_tip=True, is_preview_list=False):
        super().__init__(parent)
        self.setAlternatingRowColors(True)
        self.setViewMode(QListWidget.ListMode)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.is_preview_list = is_preview_list
        self.manual_renamed_items = {}
        self.editing_enabled = False

        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        self.itemDoubleClicked.connect(self.on_double_click)

        self.setStyleSheet(
            "QListWidget::item { height: 26px; padding: 0; margin: 0; }"
        )
        self.installEventFilter(self)
        self.show_empty_tip = show_empty_tip

        # 空列表提示
        if show_empty_tip:
            tip_text = "可将需改名的文件/文件夹拖动至此处"
        else:
            tip_text = "文件/文件夹将按此处显示进行重命名"
        self.empty_tip = QLabel(tip_text, self)
        self.empty_tip.setAlignment(Qt.AlignCenter)
        self.empty_tip.setStyleSheet(
            'font-family: "Microsoft YaHei", sans-serif; '
            "font-size: 12pt; font-weight: bold; color: #999999;"
        )
        self.empty_tip.hide()

        self.itemSelectionChanged.connect(self.on_selection_changed)

    # ---- 选中同步 ----

    def on_selection_changed(self):
        selected_items = self.selectedItems()
        if not selected_items:
            return
        main_window = self._find_main_window()
        if main_window:
            row = self.row(selected_items[0])
            main_window.sync_selection(self, row)

    def _find_main_window(self):
        """沿父级链查找 FileRenamerApp"""
        parent = self.parent()
        while parent and not isinstance(parent, FileRenamerApp):
            parent = parent.parent()
        return parent

    # ---- 尺寸与提示 ----

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'empty_tip'):
            self.empty_tip.resize(self.width(), self.height())

    def showEvent(self, event):
        super().showEvent(event)
        self._update_empty_tip()

    def _update_empty_tip(self):
        if hasattr(self, 'empty_tip'):
            self.empty_tip.setVisible(self.count() == 0)

    # ---- 键盘事件 ----

    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress and event.key() == Qt.Key_Delete:
            self._delete_selected()
            return True
        return super().eventFilter(obj, event)

    def _delete_selected(self):
        main_window = self._find_main_window()
        selected_items = self.selectedItems()
        if not selected_items or not main_window:
            return
        rows = sorted(set(self.row(item) for item in selected_items))
        if len(rows) == 1:
            main_window.delete_file(rows[0])
        else:
            main_window.delete_files(rows)

    # ---- 列表项操作 ----

    def add_file_item(self, index, filename, is_folder=False,
                      has_single_rule=False, is_manually_renamed=False):
        """【修复】不再将 self 传给 QListWidgetItem 构造器，避免重复添加"""
        item = QListWidgetItem()                       # ← 修复点：去掉 self
        custom_widget = CustomListItem(
            index, filename, is_folder, has_single_rule, is_manually_renamed
        )
        item.setSizeHint(custom_widget.sizeHint())
        self.addItem(item)
        self.setItemWidget(item, custom_widget)
        self._update_empty_tip()

    def clear(self):
        super().clear()
        self._update_empty_tip()

    def get_item_filename(self, item):
        widget = self.itemWidget(item)
        return widget.get_filename() if widget else ""

    def get_item_original_filename(self, item):
        widget = self.itemWidget(item)
        return widget.get_original_filename() if widget else ""

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

    # ---- 双击事件 ----

    def on_double_click(self, item):
        row = self.row(item)
        main_window = self._find_main_window()
        if not main_window:
            return
        if self.is_preview_list and self.editing_enabled:
            main_window.preview_manual_rename(row)
        elif not self.is_preview_list:
            if row < len(main_window.files):
                file_path = main_window.files[row]
                if file_path in self.manual_renamed_items:
                    QMessageBox.information(
                        self, "提示",
                        "此文件已在预览结果中手动重命名，请先撤销手动重命名再设置单条规则。"
                    )
                    return
                main_window.edit_single_rule(row, file_path)

    # ---- 右键菜单 ----

    def show_context_menu(self, position):
        menu = QMenu()
        selected_items = self.selectedItems()
        item_at_pos = self.itemAt(position)
        main_window = self._find_main_window()
        if not main_window:
            return

        if not self.is_preview_list:
            has_single_rules = bool(getattr(main_window, 'single_rules', {}))
            if not item_at_pos:
                if has_single_rules:
                    action = QAction("清除所有单条规则", self)
                    action.triggered.connect(main_window.clear_all_single_rules)
                    menu.addAction(action)
                action = QAction("清空列表", self)
                action.triggered.connect(main_window.clear_files)
                menu.addAction(action)
            elif selected_items:
                if len(selected_items) > 1:
                    action = QAction(f"删除选中的 {len(selected_items)} 项", self)
                    action.triggered.connect(self._delete_selected)
                    menu.addAction(action)
                else:
                    item = selected_items[0]
                    row = self.row(item)
                    file_path = main_window.files[row]

                    single_rule_action = QAction("单条规则", self)
                    if file_path in self.manual_renamed_items:
                        single_rule_action.setEnabled(False)
                    else:
                        single_rule_action.triggered.connect(
                            lambda checked=False, r=row, fp=file_path:
                                main_window.edit_single_rule(r, fp)
                        )
                    menu.addAction(single_rule_action)

                    if file_path in main_window.single_rules:
                        action = QAction("清除单条规则", self)
                        action.triggered.connect(
                            lambda checked=False, fp=file_path:
                                main_window.clear_single_rule(fp)
                        )
                        menu.addAction(action)

                    action = QAction("删除该条", self)
                    action.triggered.connect(self._delete_selected)
                    menu.addAction(action)

                if has_single_rules:
                    menu.addSeparator()
                    action = QAction("清除所有单条规则", self)
                    action.triggered.connect(main_window.clear_all_single_rules)
                    menu.addAction(action)
                menu.addSeparator()
                action = QAction("清空列表", self)
                action.triggered.connect(main_window.clear_files)
                menu.addAction(action)
        else:
            if not item_at_pos or not selected_items:
                if self.manual_renamed_items:
                    action = QAction("撤销所有手动命名", self)
                    action.triggered.connect(main_window.reset_all_manual_renamed)
                    menu.addAction(action)
                action = QAction("清空列表", self)
                action.triggered.connect(main_window.clear_files)
                menu.addAction(action)
            elif selected_items:
                item = selected_items[0]
                row = self.row(item)
                file_path = main_window.files[row]

                if self.editing_enabled:
                    action = QAction("手动重命名", self)
                    action.triggered.connect(
                        lambda checked=False, r=row: main_window.preview_manual_rename(r)
                    )
                    menu.addAction(action)

                if file_path in self.manual_renamed_items:
                    action = QAction("撤销手动命名", self)
                    action.triggered.connect(
                        lambda checked=False, r=row: main_window.reset_manual_renamed(r)
                    )
                    menu.addAction(action)

                action = QAction("删除该条", self)
                action.triggered.connect(self._delete_selected)
                menu.addAction(action)

                if self.manual_renamed_items:
                    menu.addSeparator()
                    action = QAction("撤销所有手动命名", self)
                    action.triggered.connect(main_window.reset_all_manual_renamed)
                    menu.addAction(action)

                menu.addSeparator()
                action = QAction("清空列表", self)
                action.triggered.connect(main_window.clear_files)
                menu.addAction(action)

        if not menu.isEmpty():
            menu.exec_(self.mapToGlobal(position))

    # ---- 编辑与高亮 ----

    def set_editing_enabled(self, enabled):
        self.editing_enabled = enabled

    def reset_manual_renamed_items(self):
        for i in range(self.count()):
            item = self.item(i)
            if item:
                self.set_item_manually_renamed(item, False)
        self.manual_renamed_items = {}

    def clear_sync_highlight(self):
        for i in range(self.count()):
            self.item(i).setBackground(QBrush())

    def mousePressEvent(self, event):
        self.clear_sync_highlight()
        super().mousePressEvent(event)

    def keyPressEvent(self, event):
        self.clear_sync_highlight()
        super().keyPressEvent(event)


# =============================================================================
# FileRenamerApp：主窗口
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
        except Exception:
            pass

        # ---- 数据状态 ----
        self.files = []
        self.random_strings = {}
        self.is_splitter_moving = False
        self.last_folder_path = None
        self.default_empty_tag = "."
        self.has_previewed = False
        self.single_rules = {}
        self.last_rename_operations = []
        self.last_rename_before_files = []
        self.last_rename_after_files = []

        # ---- 中文数字映射 ----
        self.cn_num = {
            '0': '〇', '1': '一', '2': '二', '3': '三', '4': '四',
            '5': '五', '6': '六', '7': '七', '8': '八', '9': '九',
            '10': '十', '20': '廿', '30': '卅'
        }
        self.weekday_map = {
            0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'
        }
        self.lunar_month_names = [
            '正', '二', '三', '四', '五', '六',
            '七', '八', '九', '十', '冬', '腊'
        ]

        # ---- 构建 UI ----
        try:
            self._init_ui()
            print("[启动] UI 初始化完成")
        except Exception as e:
            print(f"[错误] UI 初始化失败: {e}")
            traceback.print_exc()

    # =====================================================================
    # UI 初始化
    # =====================================================================

    def _init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setHandleWidth(6)
        self.splitter.setChildrenCollapsible(False)

        # ---------- 左侧面板 ----------
        left_panel = self._build_left_panel()

        # ---------- 右侧面板 ----------
        right_panel = self._build_right_panel()

        # 组装分割器
        self.splitter.addWidget(left_panel)
        self.splitter.addWidget(right_panel)
        total_width = 1000
        self.splitter.setSizes([int(total_width * 0.4), int(total_width * 0.6)])
        main_layout.addWidget(self.splitter)

        # 信号绑定放在所有控件都创建完毕之后
        self._connect_signals()

        # 初始化预览文字
        self.preview_list.set_editing_enabled(True)
        self._update_default_tag_preview()
        self._update_help_label_elide()

    # ---------- 左侧面板构建 ----------

    def _build_left_panel(self):
        left_panel = QWidget()
        left_panel.setMinimumWidth(300)
        left_layout = QVBoxLayout(left_panel)

        # ---- 基本命名规则 ----
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
        help_btn.setToolTip("点击查看命名规则帮助")
        help_btn.clicked.connect(self._show_naming_help)
        pattern_layout.addWidget(help_btn)
        naming_layout.addLayout(pattern_layout)

        # 简要提示
        self.help_label = QLabel(
            "提示: * 原文件名, # 序号, $ 随机数, ? 单字符, "
            "<*-n> 去尾, <-n*> 去头, \\n 第n个字符, <-> 日期"
        )
        smaller_font = self.help_label.font()
        smaller_font.setPointSize(smaller_font.pointSize() - 1)
        self.help_label.setFont(smaller_font)
        self.help_label.setTextFormat(Qt.PlainText)
        self.help_label.setWordWrap(False)
        self.help_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.help_label.setToolTip(self._get_naming_help_text())
        self.help_label.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        self.help_label.setMinimumWidth(100)
        naming_layout.addWidget(self.help_label)

        # 空<>默认值
        default_tag_layout = QHBoxLayout()
        default_tag_layout.addWidget(QLabel("空<>默认值:"))
        self.default_tag_edit = QLineEdit(self.default_empty_tag)
        self.default_tag_edit.setMinimumWidth(120)
        self.default_tag_edit.setToolTip("空<>标签的默认规则")
        default_tag_layout.addWidget(self.default_tag_edit)
        self.default_tag_preview = QLabel("")
        self.default_tag_preview.setStyleSheet("color: #666;")
        default_tag_layout.addWidget(self.default_tag_preview)
        default_tag_layout.addStretch()
        naming_layout.addLayout(default_tag_layout)

        # 扩展名
        ext_layout = QHBoxLayout()
        ext_layout.addWidget(QLabel("扩展名:"))
        self.ext_edit = QLineEdit()
        self.ext_edit.setPlaceholderText("留空保持原扩展名, 也可用 *, #, $, ? 规则")
        ext_layout.addWidget(self.ext_edit)
        naming_layout.addLayout(ext_layout)

        naming_group.setLayout(naming_layout)
        left_layout.addWidget(naming_group)

        # ---- 序号设置 ----
        sequence_group = QGroupBox("序号设置")
        sequence_layout = QGridLayout()

        sequence_layout.addWidget(QLabel("起始序号:"), 0, 0)
        self.start_num = QSpinBox()
        self.start_num.setRange(0, 99999)
        self.start_num.setValue(1)
        sequence_layout.addWidget(self.start_num, 0, 1)

        sequence_layout.addWidget(QLabel("序号增量:"), 1, 0)
        self.increment_num = QSpinBox()
        self.increment_num.setRange(-999, 999)
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

        # ---- 随机数设置 ----
        random_group = QGroupBox("随机数设置")
        random_layout = QGridLayout()

        random_layout.addWidget(QLabel("随机数位数:"), 0, 0)
        self.random_digits = QSpinBox()
        self.random_digits.setRange(1, 20)
        self.random_digits.setValue(7)
        random_layout.addWidget(self.random_digits, 0, 1)

        random_layout.addWidget(QLabel("随机数类型:"), 1, 0)
        self.random_type = QComboBox()
        self.random_type.addItems([
            "纯数字", "数字+小写字母", "数字+大写字母", "数字+大小写字母"
        ])
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

        # ---- 高级设置 ----
        advanced_group = QGroupBox("高级设置")
        advanced_layout = QVBoxLayout()

        # 正则/替换
        use_regex_layout = QHBoxLayout()
        self.use_regex = QCheckBox("启用正则表达式")
        use_regex_layout.addWidget(self.use_regex)
        regex_help_btn = QToolButton()
        regex_help_btn.setText("?")
        regex_help_btn.clicked.connect(self._show_regex_help)
        use_regex_layout.addWidget(regex_help_btn)
        use_regex_layout.addStretch()
        advanced_layout.addLayout(use_regex_layout)

        replace_from_layout = QHBoxLayout()
        replace_from_layout.addWidget(QLabel("替换:"))
        self.replace_from = QLineEdit()
        self.replace_from.setPlaceholderText("要替换的文本(多个用/分隔)")
        replace_from_layout.addWidget(self.replace_from)
        advanced_layout.addLayout(replace_from_layout)

        replace_to_layout = QHBoxLayout()
        replace_to_layout.addWidget(QLabel("为:"))
        self.replace_to = QLineEdit()
        self.replace_to.setPlaceholderText("替换后的文本(多个用/分隔)")
        replace_to_layout.addWidget(self.replace_to)
        advanced_layout.addLayout(replace_to_layout)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        advanced_layout.addWidget(line)

        # 插入/删除
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
        advanced_layout.addLayout(edit_layout)

        # 冲突处理
        conflict_layout = QHBoxLayout()
        conflict_layout.addWidget(QLabel("文件名冲突:"))
        self.auto_resolve_conflicts = QCheckBox("自动解决冲突")
        self.auto_resolve_conflicts.setChecked(True)
        conflict_layout.addWidget(self.auto_resolve_conflicts)
        advanced_layout.addLayout(conflict_layout)

        advanced_group.setLayout(advanced_layout)
        left_layout.addWidget(advanced_group)

        # ---- 操作按钮 ----
        button_layout = QHBoxLayout()
        bigger_font = QFont()
        bigger_font.setPointSize(QApplication.font().pointSize() + 1)
        bigger_font.setBold(True)

        preview_btn = QPushButton("预览重命名结果")
        preview_btn.setMinimumHeight(45)
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
        watermark_label = QLabel("yumumao@medu.cc 3.8.1")
        watermark_label.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
        watermark_label.setStyleSheet(
            "color: rgba(128, 128, 128, 100); margin-top: 0px; "
            "margin-bottom: 0px; padding: 0px;"
        )
        watermark_label.setAlignment(Qt.AlignLeft | Qt.AlignBottom)
        watermark_label.setFixedHeight(15)
        left_layout.addWidget(watermark_label)

        return left_panel

    # ---------- 右侧面板构建 ----------

    def _build_right_panel(self):
        right_panel = QWidget()
        right_panel.setMinimumWidth(300)
        right_layout = QVBoxLayout(right_panel)

        # 添加/清空按钮
        file_btns_layout = QHBoxLayout()
        add_files_btn = QPushButton("添加文件")
        add_files_btn.clicked.connect(self._add_files_dialog)
        file_btns_layout.addWidget(add_files_btn)
        add_folders_btn = QPushButton("添加文件/文件夹")
        add_folders_btn.clicked.connect(self._add_folders_and_files_dialog)
        file_btns_layout.addWidget(add_folders_btn)
        clear_btn = QPushButton("清空列表")
        clear_btn.clicked.connect(self.clear_files)
        file_btns_layout.addWidget(clear_btn)
        right_layout.addLayout(file_btns_layout)

        # 排序与移动
        control_layout = QHBoxLayout()
        control_layout.addWidget(QLabel("排序:"))
        self.sort_combo = QComboBox()
        self.sort_combo.addItems([
            "名称升序", "名称降序", "日期升序", "日期降序", "大小升序", "大小降序"
        ])
        self.sort_combo.setCurrentIndex(2)
        self.sort_combo.setFixedHeight(26)
        control_layout.addWidget(self.sort_combo)
        move_up_btn = QPushButton("上移")
        move_up_btn.clicked.connect(self._move_file_up)
        move_up_btn.setFixedHeight(26)
        control_layout.addWidget(move_up_btn)
        move_down_btn = QPushButton("下移")
        move_down_btn.clicked.connect(self._move_file_down)
        move_down_btn.setFixedHeight(26)
        control_layout.addWidget(move_down_btn)
        right_layout.addLayout(control_layout)

        # 原文件列表
        list_header = QHBoxLayout()
        list_header.addWidget(QLabel("原文件:"))
        sel_tip = QLabel("(支持Shift/Ctrl多选)")
        sel_tip.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        sel_font = sel_tip.font()
        sel_font.setPointSize(sel_font.pointSize() - 1)
        sel_tip.setFont(sel_font)
        list_header.addWidget(sel_tip)
        right_layout.addLayout(list_header)

        self.file_list = FileListWidget(show_empty_tip=True, is_preview_list=False)
        right_layout.addWidget(self.file_list)

        # 预览结果列表
        preview_header = QHBoxLayout()
        preview_header.addWidget(QLabel("预览结果:"))
        self.revert_btn = QPushButton("回退重命名")
        self.revert_btn.setToolTip("先回退预览，再点击'执行重命名'正式回退")
        self.revert_btn.setEnabled(False)
        self.revert_btn.setMaximumHeight(25)
        self.revert_btn.clicked.connect(self._revert_last_rename)
        preview_header.addStretch()
        preview_header.addWidget(self.revert_btn)
        right_layout.addLayout(preview_header)

        self.preview_list = FileListWidget(show_empty_tip=False, is_preview_list=True)
        right_layout.addWidget(self.preview_list)

        return right_panel

    # ---------- 信号连接（全部集中，确保控件已创建） ----------

    def _connect_signals(self):
        """【修复】所有信号连接集中到此处，保证所有控件都已创建完毕"""
        self.sort_combo.currentIndexChanged.connect(self._sort_files)
        self.default_tag_edit.textChanged.connect(self._update_default_tag_preview)
        self.operation_type.currentTextChanged.connect(self._toggle_edit_controls)
        self.splitter.splitterMoved.connect(self._on_splitter_moved)

    # =====================================================================
    # 分割器逻辑（修复卡死问题）
    # =====================================================================

    def _on_splitter_moved(self, pos, index):
        """【修复】移除 is_splitter_moving 标志，改为直接处理"""
        try:
            self._update_help_label_elide()
        except Exception as e:
            print(f"[警告] 分割器移动处理出错: {e}")

    def _update_help_label_elide(self):
        """根据可用宽度自动省略提示文本"""
        if not hasattr(self, 'help_label'):
            return
        width = self.help_label.width()
        if width <= 0:
            return
        full_text = (
            "提示: * 原文件名, # 序号, $ 随机数, ? 单字符, "
            "<*-n> 去尾, <-n*> 去头, \\n 第n个字符, <-> 日期"
        )
        fm = self.help_label.fontMetrics()
        if fm.horizontalAdvance(full_text) > width:
            elided = full_text
            for i in range(len(full_text) - 1, 0, -1):
                elided = full_text[:i] + "..."
                if fm.horizontalAdvance(elided) < width:
                    break
            self.help_label.setText(elided)
        else:
            self.help_label.setText(full_text)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_help_label_elide()

    # =====================================================================
    # 同步选择
    # =====================================================================

    def sync_selection(self, sender, row):
        try:
            if sender == self.file_list:
                if self.preview_list.count() > 0:
                    self.preview_list.clearSelection()
                    self.preview_list.clear_sync_highlight()
                    if 0 <= row < self.preview_list.count():
                        item = self.preview_list.item(row)
                        item.setSelected(False)
                        item.setBackground(QColor("#ffffa8"))
            elif sender == self.preview_list:
                self.file_list.clearSelection()
                self.file_list.clear_sync_highlight()
                if 0 <= row < self.file_list.count():
                    item = self.file_list.item(row)
                    item.setSelected(False)
                    item.setBackground(QColor("#ffffa8"))
        except Exception as e:
            print(f"[警告] 同步选择出错: {e}")

    # =====================================================================
    # 帮助文本
    # =====================================================================

    def _get_naming_help_text(self):
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
            "  yyyy/yy - 年份(四位/两位)   YYYY/YY - 年份(中文)\n"
            "  mm/m  - 月份(补零/不补零)   MM/M - 月份(中文)\n"
            "  dd/d  - 日期(补零/不补零)   DD/D - 日期(中文)\n"
            "  hh/h  - 小时(24/12小时制)   HH/H - 小时(中文)\n"
            "  w/ddd - 星期(数字)           W/DDD - 星期(中文)\n"
            "  tt/t  - 分钟(补零/不补零)   TT/T - 分钟(中文)\n"
            "  ss    - 秒钟\n\n"
            "简化日期格式：\n"
            "  <->  2025-3-19    <--> 2025年3月19日\n"
            "  <|>  中文日期     <||> 农历日期\n"
            "  <:>  0557         <::> 055712\n"
            "  <-:> 日期+时间   <.>  20250319\n"
            "  <>   空标签使用默认规则\n\n"
            "-- by yumumao@medu.cc"
        )

    def _show_naming_help(self):
        HelpDialog(self._get_naming_help_text(), self).exec_()

    def _show_regex_help(self):
        help_text = (
            "正则表达式简明指南：\n\n"
            ". - 匹配任意单个字符\n"
            "\\d - 匹配数字  \\w - 匹配字母数字下划线\n"
            "* + ? - 0+次 / 1+次 / 0或1次\n"
            "{n,m} - n到m次\n"
            "^ $ - 开头 / 结尾\n"
            "[] - 字符集  () - 分组\n"
            "\\1 \\2 - 反向引用\n\n"
            "示例：\n"
            "  \\d+ → # 替换所有数字为#\n"
            "  (\\w+)_(\\d+) → \\2_\\1 互换位置\n"
        )
        dlg = HelpDialog(help_text, self)
        dlg.setWindowTitle("正则表达式帮助")
        dlg.exec_()

    # =====================================================================
    # 空<>默认标签
    # =====================================================================

    def _update_default_tag_preview(self):
        tag_content = self.default_tag_edit.text().strip()
        self.default_empty_tag = tag_content
        try:
            test_text = f"<{tag_content}>"
            processed = self._process_date_time_tags(test_text)
            self.default_tag_preview.setText(f"预览: {processed}")
        except Exception as e:
            self.default_tag_preview.setText(f"格式错误: {e}")

    def _is_valid_empty_tag(self):
        tag_content = self.default_empty_tag.strip()
        if not tag_content:
            return False
        test_text = f"<{tag_content}>"
        processed = self._process_date_time_tags(test_text)
        return processed != test_text

    # =====================================================================
    # 文件添加 / 删除 / 排序 / 移动
    # =====================================================================

    def _add_files_dialog(self):
        try:
            initial_dir = self.last_folder_path or os.getcwd()
            file_paths, _ = QFileDialog.getOpenFileNames(self, "选择文件", initial_dir)
            if file_paths:
                self.last_folder_path = os.path.dirname(file_paths[0])
                self._add_files(file_paths)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"选择文件时发生错误: {e}")
            traceback.print_exc()

    def _add_folders_and_files_dialog(self):
        try:
            dialog = EnhancedFileDialog(self, self.last_folder_path)
            if dialog.exec_() == QDialog.Accepted:
                selected = dialog.get_selected_paths()
                if selected:
                    self.last_folder_path = dialog.get_current_directory()
                    self._add_files(selected)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"选择文件/文件夹时发生错误: {e}")
            traceback.print_exc()

    def _add_files(self, file_paths):
        if not isinstance(file_paths, list):
            return
        for fp in file_paths:
            if fp not in self.files and (os.path.isfile(fp) or os.path.isdir(fp)):
                self.files.append(fp)
        self._update_file_list()
        self._sort_files()

    def clear_files(self):
        self.files = []
        self._update_file_list()
        self.preview_list.clear()
        self.random_strings = {}
        self.has_previewed = False
        self.preview_list.reset_manual_renamed_items()
        self.single_rules = {}

    def _update_file_list(self):
        self.file_list.clear()
        for i, fp in enumerate(self.files):
            is_folder = os.path.isdir(fp)
            has_rule = fp in self.single_rules
            self.file_list.add_file_item(i + 1, os.path.basename(fp), is_folder, has_rule)

    def delete_file(self, row):
        if 0 <= row < len(self.files):
            fp = self.files.pop(row)
            self.single_rules.pop(fp, None)
            self._update_file_list()
            if self.preview_list.count() > 0:
                self.preview_rename()

    def delete_files(self, rows):
        for row in sorted(rows, reverse=True):
            if 0 <= row < len(self.files):
                fp = self.files.pop(row)
                self.single_rules.pop(fp, None)
        self._update_file_list()
        if self.preview_list.count() > 0:
            self.preview_rename()

    def _sort_files(self):
        if not self.files:
            return
        sort_option = self.sort_combo.currentText()
        file_info = []
        for fp in self.files:
            try:
                stat = os.stat(fp)
                file_info.append({
                    'path': fp, 'name': os.path.basename(fp),
                    'mtime': stat.st_mtime, 'size': stat.st_size,
                })
            except (FileNotFoundError, PermissionError) as e:
                print(f"[警告] 无法访问 {fp}: {e}")

        sort_map = {
            "名称升序": (lambda x: x['name'].lower(), False),
            "名称降序": (lambda x: x['name'].lower(), True),
            "日期升序": (lambda x: x['mtime'], False),
            "日期降序": (lambda x: x['mtime'], True),
            "大小升序": (lambda x: x['size'], False),
            "大小降序": (lambda x: x['size'], True),
        }
        key_func, reverse = sort_map.get(sort_option, (lambda x: x['name'].lower(), False))
        file_info.sort(key=key_func, reverse=reverse)
        self.files = [info['path'] for info in file_info]
        self._update_file_list()

    def _move_file_up(self):
        row = self.file_list.currentRow()
        if row > 0:
            fp = self.files.pop(row)
            self.files.insert(row - 1, fp)
            self._update_file_list()
            self.file_list.setCurrentRow(row - 1)

    def _move_file_down(self):
        row = self.file_list.currentRow()
        if 0 <= row < self.file_list.count() - 1:
            fp = self.files.pop(row)
            self.files.insert(row + 1, fp)
            self._update_file_list()
            self.file_list.setCurrentRow(row + 1)

    # =====================================================================
    # 拖放
    # =====================================================================

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
            self._add_files(paths)

    # =====================================================================
    # 控件联动
    # =====================================================================

    def _toggle_edit_controls(self, text):
        if text == "插入":
            self.edit_text.setVisible(True)
            self.delete_count.setVisible(False)
        else:
            self.edit_text.setVisible(False)
            self.delete_count.setVisible(True)

    # =====================================================================
    # 转义字符保护与恢复
    # =====================================================================

    def _protect_escapes(self, pattern):
        """将 \\$, \\#, \\? 替换为占位符"""
        pattern = pattern.replace('\\$', '\x00DOLLAR\x00')
        pattern = pattern.replace('\\#', '\x00HASH\x00')
        pattern = pattern.replace('\\?', '\x00QUESTION\x00')
        return pattern

    def _restore_escapes(self, text):
        """将占位符还原为实际字符"""
        text = text.replace('\x00DOLLAR\x00', '$')
        text = text.replace('\x00HASH\x00', '#')
        text = text.replace('\x00QUESTION\x00', '?')
        return text

    # =====================================================================
    # 截断模式 <*-n> / <-n*>
    # =====================================================================

    def _process_truncate(self, pattern, original_name):
        for match in re.findall(r'<\*-(\d+)>', pattern):
            n = int(match)
            repl = original_name[:-n] if n < len(original_name) else ""
            pattern = pattern.replace(f"<*-{match}>", repl)
        for match in re.findall(r'<-(\d+)\*>', pattern):
            n = int(match)
            repl = original_name[n:] if n < len(original_name) else ""
            pattern = pattern.replace(f"<-{match}*>", repl)
        return pattern

    # =====================================================================
    # \\n（取原文件名第n个字符）
    # =====================================================================

    def _process_char_index(self, pattern, original_name):
        for match in re.findall(r'\\(\d+)', pattern):
            n = int(match)
            repl = original_name[n - 1] if 1 <= n <= len(original_name) else ""
            pattern = pattern.replace(f"\\{match}", repl)
        return pattern

    # =====================================================================
    # ? 问号通配符
    # =====================================================================

    def _process_questions(self, pattern, original_name):
        """连续 ? 依次映射为原文件名中对应字符（不重复调用转义保护）"""
        question_groups = []
        current = ''
        for ch in pattern:
            if ch == '?':
                current += '?'
            else:
                if current:
                    question_groups.append(current)
                    current = ''
        if current:
            question_groups.append(current)

        result = pattern
        name_cursor = original_name
        for group in question_groups:
            count = len(group)
            replacement = name_cursor[:count]
            name_cursor = name_cursor[count:]
            result = result.replace(group, replacement, 1)
        return result

    # =====================================================================
    # 日期时间标签 <...>
    # =====================================================================

    def _process_date_time_tags(self, text):
        now = datetime.datetime.now()
        text = self._process_empty_tags(text)
        text = self._process_simplified_dates(text, now)

        def replace_tag(match):
            return self._format_datetime(match.group(1), now)

        return re.sub(r'<([^>]+)>', replace_tag, text)

    def _format_datetime(self, template, now):
        dt_codes = {
            "yyyy": lambda: now.strftime('%Y'),
            "yy":   lambda: now.strftime('%y'),
            "YYYY": lambda: ''.join(self.cn_num[d] for d in now.strftime('%Y')),
            "YY":   lambda: ''.join(self.cn_num[d] for d in now.strftime('%y')),
            "mm":   lambda: f"{now.month:02d}",
            "m":    lambda: str(now.month),
            "MM":   lambda: self._cn_month(now.month, True),
            "M":    lambda: self._cn_month(now.month, False),
            "dd":   lambda: f"{now.day:02d}",
            "d":    lambda: str(now.day),
            "DD":   lambda: self._cn_day(now.day, True),
            "D":    lambda: self._cn_day(now.day, False),
            "hh":   lambda: f"{now.hour:02d}",
            "h":    lambda: str(now.hour if now.hour <= 12 else now.hour - 12),
            "HH":   lambda: self._cn_number(now.hour),
            "H":    lambda: self._cn_number(now.hour if now.hour <= 12 else now.hour - 12),
            "w":    lambda: str(now.weekday() + 1),
            "W":    lambda: self.weekday_map[now.weekday()],
            "ddd":  lambda: str(now.weekday() + 1),
            "DDD":  lambda: self.weekday_map[now.weekday()],
            "tt":   lambda: now.strftime('%M'),
            "t":    lambda: str(now.minute),
            "TT":   lambda: ''.join(self.cn_num[d] for d in f"{now.minute:02d}"),
            "T":    lambda: ''.join(self.cn_num[d] for d in str(now.minute)),
            "ss":   lambda: now.strftime('%S'),
        }
        for code in sorted(dt_codes, key=len, reverse=True):
            template = template.replace(code, dt_codes[code]())
        return template

    def _process_empty_tags(self, text):
        return re.sub(r'<>', f'<{self.default_empty_tag}>', text)

    def _process_simplified_dates(self, text, now):
        # 先处理 <-:> 避免与 <-> 冲突
        if '<-:>' in text:
            text = text.replace(
                '<-:>',
                f"{now.year}-{now.month}-{now.day} {now.hour:02d}{now.minute:02d}"
            )
        mapping = {
            '<->':  f"{now.year}-{now.month}-{now.day}",
            '<-->': f"{now.year}年{now.month}月{now.day}日",
            '<|>':  (f"{''.join(self.cn_num[d] for d in str(now.year))}年"
                     f"{self._cn_month(now.month, False)}月"
                     f"{self._cn_day(now.day, True)}日"),
            '<||>': (f"农历{self.lunar_month_names[self._approx_lunar_month()-1]}月"
                     f"{self._cn_day(self._approx_lunar_day(), True)}"),
            '<:>':  f"{now.hour:02d}{now.minute:02d}",
            '<::>': f"{now.hour:02d}{now.minute:02d}{now.second:02d}",
            '<.>':  f"{now.year}{now.month:02d}{now.day:02d}",
        }
        for tag, repl in mapping.items():
            if tag in text:
                text = text.replace(tag, repl)
        return text

    # =====================================================================
    # 中文数字 / 农历近似
    # =====================================================================

    def _approx_lunar_month(self):
        m = datetime.datetime.now().month - 1
        return m if m >= 1 else 12

    def _approx_lunar_day(self):
        return (datetime.datetime.now().day + 15) % 30 or 30

    def _cn_month(self, month, long_format=True):
        if month <= 10:
            return self.cn_num[str(month)]
        elif month == 11:
            return "十一"
        else:
            return "十二"

    def _cn_day(self, day, long_format=True):
        if day <= 10:
            if long_format and day == 10:
                return "一十"
            return self.cn_num[str(day)]
        elif day < 20:
            return ("一十" if long_format else "十") + self.cn_num[str(day)[1]]
        elif day == 20:
            return "二十" if long_format else "廿"
        elif day < 30:
            return ("二十" if long_format else "廿") + self.cn_num[str(day)[1]]
        else:
            suffix = "" if day == 30 else self.cn_num["1"]
            return ("三十" if long_format else "卅") + suffix

    def _cn_number(self, num):
        if num <= 10:
            return self.cn_num[str(num)]
        elif num < 20:
            return "十" + (self.cn_num[str(num)[1]] if num > 10 else "")
        else:
            tens = self.cn_num[str(num)[0]] + "十"
            ones = self.cn_num[str(num)[1]] if num % 10 else ""
            return tens + ones

    # =====================================================================
    # 序号 / 随机数
    # =====================================================================

    def _gen_sequence(self, index):
        """index 从 1 开始"""
        seq_num = self.start_num.value() + (index - 1) * self.increment_num.value()
        digits = self.digits_num.value()

        if self.number_radio.isChecked():
            if self.pad_zeros.isChecked():
                return str(seq_num).zfill(digits)
            return str(seq_num)

        elif self.letter_radio.isChecked():
            offset = ord('a') if self.letter_case.currentText() == "小写字母" else ord('A')
            s = ""
            temp = seq_num
            while temp > 0 or not s:
                temp -= 1
                s = chr(offset + (temp % 26)) + s
                temp //= 26
            if self.pad_zeros.isChecked() and len(s) < digits:
                s = chr(offset) * (digits - len(s)) + s
            return s

        else:
            number_part = (seq_num - 1) // 26 + 1
            letter_part = (seq_num - 1) % 26
            offset = ord('a') if self.letter_case.currentText() == "小写字母" else ord('A')
            letter_char = chr(offset + letter_part)
            num_digits = max(digits - 1, 1)
            if self.pad_zeros.isChecked():
                number_str = str(number_part).zfill(num_digits)
            else:
                number_str = str(number_part)
            return number_str + letter_char

    def _gen_random(self, index):
        from_idx = self.same_random_from.value()
        to_idx = self.same_random_to.value()
        if from_idx < to_idx and from_idx <= index <= to_idx:
            key = f"group_{from_idx}_{to_idx}"
        else:
            key = f"single_{index}"
        if key not in self.random_strings:
            length = self.random_digits.value()
            rt = self.random_type.currentText()
            if rt == "纯数字":
                chars = string.digits
            elif rt == "数字+小写字母":
                chars = string.digits + string.ascii_lowercase
            elif rt == "数字+大写字母":
                chars = string.digits + string.ascii_uppercase
            else:
                chars = string.digits + string.ascii_letters
            self.random_strings[key] = ''.join(random.choice(chars) for _ in range(length))
        return self.random_strings[key]

    # =====================================================================
    # 文本替换
    # =====================================================================

    def _apply_replacements(self, text, from_list, to_list):
        use_regex = self.use_regex.isChecked()
        pairs = sorted(zip(from_list, to_list), key=lambda x: len(x[0]), reverse=True)
        for rf, rt in pairs:
            if not rf:
                continue
            try:
                if use_regex:
                    text = re.sub(rf, rt, text)
                else:
                    text = text.replace(rf, rt)
            except Exception as e:
                print(f"[警告] 替换出错: {e}")
        return text

    # =====================================================================
    # 扩展名规则处理
    # =====================================================================

    def _process_ext_rule(self, ext_rule, original_ext, index):
        """对扩展名规则应用完整的规则引擎"""
        new_ext = ext_rule
        new_ext = self._protect_escapes(new_ext)

        try:
            new_ext = self._process_date_time_tags(new_ext)
        except Exception:
            pass

        if "<*-" in new_ext or ("<-" in new_ext and "*>" in new_ext):
            new_ext = self._process_truncate(new_ext, original_ext)
        if "\\" in new_ext:
            new_ext = self._process_char_index(new_ext, original_ext)
        if '?' in new_ext:
            new_ext = self._process_questions(new_ext, original_ext)

        try:
            seq_str = self._gen_sequence(index + 1)
        except Exception:
            seq_str = str(index + 1)
        try:
            random_str = self._gen_random(index + 1)
        except Exception:
            random_str = ''.join(random.choices(string.digits, k=4))

        new_ext = re.sub(r"(?<!\\)\*", original_ext, new_ext)
        new_ext = re.sub(r"(?<!\\)#", lambda m: seq_str, new_ext)
        new_ext = re.sub(r"(?<!\\)\$", lambda m: random_str, new_ext)

        new_ext = new_ext.replace('\\*', '*')
        new_ext = new_ext.replace('\\#', '#')
        new_ext = new_ext.replace('\\$', '$')
        new_ext = new_ext.replace('\\?', '?')
        new_ext = self._restore_escapes(new_ext)

        if new_ext and not new_ext.startswith('.'):
            new_ext = '.' + new_ext
        return new_ext

    # =====================================================================
    # 核心：生成新文件名
    # =====================================================================

    def generate_new_name(self, file_path, index, custom_rule=None):
        base_name = os.path.basename(file_path)
        if os.path.isdir(file_path):
            name, ext = base_name, ""
        else:
            name, ext = os.path.splitext(base_name)

        # 确定规则来源
        if custom_rule is not None:
            pattern = custom_rule
        elif file_path in self.single_rules:
            pattern = self.single_rules[file_path]
        else:
            pattern = self.pattern_edit.text()

        # 转义保护
        pattern = self._protect_escapes(pattern)

        # 日期时间标签
        try:
            pattern = self._process_date_time_tags(pattern)
        except Exception as e:
            print(f"[警告] 日期处理异常: {e}")

        # 截断模式
        try:
            if "<*-" in pattern or ("<-" in pattern and "*>" in pattern):
                pattern = self._process_truncate(pattern, name)
        except Exception as e:
            print(f"[警告] 截断处理异常: {e}")

        # 特定字符
        try:
            if "\\" in pattern:
                pattern = self._process_char_index(pattern, name)
        except Exception as e:
            print(f"[警告] 字符索引异常: {e}")

        # 问号通配符
        try:
            if '?' in pattern:
                pattern = self._process_questions(pattern, name)
        except Exception as e:
            print(f"[警告] 问号处理异常: {e}")

        # 序号与随机数
        try:
            seq_str = self._gen_sequence(index + 1)
        except Exception:
            seq_str = str(index + 1)
        try:
            random_str = self._gen_random(index + 1)
        except Exception:
            random_str = ''.join(random.choices(string.digits, k=4))

        # 通配符替换
        try:
            new_name = pattern
            new_name = re.sub(r"(?<!\\)\*", name, new_name)
            new_name = re.sub(r"(?<!\\)#", lambda m: seq_str, new_name)
            new_name = re.sub(r"(?<!\\)\$", lambda m: random_str, new_name)
        except Exception as e:
            print(f"[警告] 通配符替换异常: {e}")
            new_name = name

        # 恢复转义
        new_name = new_name.replace('\\*', '*')
        new_name = new_name.replace('\\#', '#')
        new_name = new_name.replace('\\$', '$')
        new_name = new_name.replace('\\?', '?')
        new_name = self._restore_escapes(new_name)

        # 文本替换
        if self.replace_from.text().strip():
            from_list = self.replace_from.text().split('/')
            to_list = self.replace_to.text().split('/')
            for i in range(len(to_list)):
                if to_list[i] == ":" and i > 0:
                    to_list[i] = to_list[i - 1]
            if from_list and len(to_list) < len(from_list):
                if not to_list:
                    to_list = [''] * len(from_list)
                else:
                    to_list += [to_list[-1]] * (len(from_list) - len(to_list))
            new_name = self._apply_replacements(new_name, from_list, to_list)

        # 拼接扩展名
        if os.path.isdir(file_path):
            return new_name
        ext_rule = self.ext_edit.text().strip()
        if ext_rule:
            original_ext = ext[1:] if ext.startswith('.') else ext
            new_ext = self._process_ext_rule(ext_rule, original_ext, index)
            return new_name + new_ext
        return new_name + ext

    # =====================================================================
    # 预览
    # =====================================================================

    def preview_rename(self):
        if not self.files:
            QMessageBox.warning(self, "警告", "请先添加文件!")
            return

        pattern = self.pattern_edit.text()
        if '<>' in pattern and not self._is_valid_empty_tag():
            QMessageBox.warning(
                self, "空标签警告",
                f"默认值 '{self.default_empty_tag}' 不是有效的日期时间格式。"
            )
            return

        self.random_strings = {}
        self.preview_list.clear()

        new_names = []
        for i, fp in enumerate(self.files):
            try:
                new_names.append(self.generate_new_name(fp, i))
            except Exception as e:
                print(f"[错误] 处理 {os.path.basename(fp)}: {e}")
                traceback.print_exc()
                new_names.append(os.path.basename(fp))

        # 覆盖手动重命名项
        for i, fp in enumerate(self.files):
            if fp in self.preview_list.manual_renamed_items:
                new_names[i] = self.preview_list.manual_renamed_items[fp]

        # 填充预览
        for i, (fp, nn) in enumerate(zip(self.files, new_names)):
            is_folder = os.path.isdir(fp)
            has_rule = fp in self.single_rules
            is_manual = fp in self.preview_list.manual_renamed_items
            self.preview_list.add_file_item(i + 1, nn, is_folder, has_rule, is_manual)

        # 标记非法字符
        for i in range(self.preview_list.count()):
            item = self.preview_list.item(i)
            widget = self.preview_list.itemWidget(item)
            if widget and has_illegal_chars(new_names[i]):
                widget.filename_label.setStyleSheet(
                    "padding: 2px; font-size: 12pt; font-weight: bold; color: red;"
                )
                widget.filename_label.setText(widget.current_filename + " [包含无效字符]")

        self.preview_list.set_editing_enabled(True)
        self.has_previewed = True

    # =====================================================================
    # 执行重命名
    # =====================================================================

    def execute_rename(self):
        if not self.files:
            QMessageBox.warning(self, "警告", "请先添加文件!")
            return
        if not self.has_previewed or self.preview_list.count() == 0:
            QMessageBox.warning(self, "警告", "请先点击「预览重命名结果」!")
            return

        new_names = []
        for i in range(self.preview_list.count()):
            item = self.preview_list.item(i)
            widget = self.preview_list.itemWidget(item)
            if widget:
                new_names.append(widget.current_filename)

        # 检查非法字符
        if any(has_illegal_chars(n) for n in new_names):
            reply = QMessageBox.question(
                self, "存在无效字符",
                "预览中存在无效字符，是否返回修改？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                return

        # 冲突检测
        has_conflicts, conflict_items = self._detect_conflicts(new_names)
        if has_conflicts:
            new_names = self._handle_conflicts(new_names, conflict_items)
            if new_names is None:
                return

        if QMessageBox.question(
            self, "确认", f"确定重命名 {len(self.files)} 个文件/文件夹？",
            QMessageBox.Yes | QMessageBox.No
        ) != QMessageBox.Yes:
            return

        self.last_rename_before_files = self.files.copy()
        self.last_rename_after_files = []
        self.last_rename_operations = []

        success, errors, unchanged, msgs = self._perform_rename(new_names)
        self._update_file_list()
        self.revert_btn.setEnabled(True)
        self.preview_rename()
        self._show_rename_result(success, errors, unchanged, msgs)

    def _detect_conflicts(self, new_names):
        has_conflicts = False
        conflict_items = []
        for i, nn in enumerate(new_names):
            dp = os.path.dirname(self.files[i])
            for j in range(len(self.files)):
                if i != j:
                    if dp == os.path.dirname(self.files[j]) and nn == new_names[j]:
                        has_conflicts = True
                        conflict_items.append({
                            'index': i + 1,
                            'original_name': os.path.basename(self.files[i]),
                            'new_name': nn,
                        })
                        break
        return has_conflicts, conflict_items

    def _handle_conflicts(self, new_names, conflict_items):
        if self.auto_resolve_conflicts.isChecked():
            try:
                new_names = self._auto_resolve(new_names)
                for i, name in enumerate(new_names):
                    if i < self.preview_list.count():
                        item = self.preview_list.item(i)
                        widget = self.preview_list.itemWidget(item)
                        if widget:
                            widget.set_filename(name)
                return new_names
            except Exception as e:
                print(f"[警告] 自动冲突解决出错: {e}")

        try:
            dialog = RenameConflictDialog(conflict_items, self)
            if dialog.exec_() == QDialog.Accepted:
                for idx, name in dialog.get_new_names().items():
                    new_names[idx - 1] = name
                return new_names
            return None
        except Exception as e:
            print(f"[警告] 冲突对话框出错: {e}")
            return None

    def _show_rename_result(self, success, errors, unchanged, msgs):
        dlg = QMessageBox(self)
        dlg.setWindowTitle("重命名结果")
        if errors > 0:
            dlg.setIcon(QMessageBox.Warning)
            text = f"成功: {success}  未修改: {unchanged}  失败: {errors}\n\n"
            text += "\n".join(msgs[:10])
            if len(msgs) > 10:
                text += "\n..."
            dlg.setText(text)
        elif success > 0:
            dlg.setIcon(QMessageBox.Information)
            dlg.setText(f"成功: {success}  未修改: {unchanged}")
        else:
            dlg.setIcon(QMessageBox.Information)
            dlg.setText("所有文件名未变化，没有文件被修改")
        dlg.setStandardButtons(QMessageBox.Ok)
        if success > 0:
            undo_btn = dlg.addButton("撤销重命名", QMessageBox.ActionRole)
            undo_btn.clicked.connect(self._undo_rename)
        dlg.exec_()

    # =====================================================================
    # 撤销 / 回退
    # =====================================================================

    def _undo_rename(self):
        if not self.last_rename_operations:
            QMessageBox.information(self, "撤销", "没有可撤销的操作")
            return
        success = 0
        errors = 0
        msgs = []
        for op in reversed(self.last_rename_operations):
            try:
                if os.path.exists(op['new_path']):
                    for i, fp in enumerate(self.files):
                        if fp == op['new_path']:
                            self.files[i] = op['old_path']
                    os.rename(op['new_path'], op['old_path'])
                    success += 1
                else:
                    errors += 1
                    msgs.append(f"找不到: {op['new_path']}")
            except Exception as e:
                errors += 1
                msgs.append(str(e))
        self._update_file_list()
        self.preview_rename()
        if errors > 0:
            QMessageBox.warning(self, "撤销结果",
                f"成功: {success}  失败: {errors}\n\n" + "\n".join(msgs[:10]))
        else:
            QMessageBox.information(self, "撤销结果", f"成功撤销 {success} 项")
        self.last_rename_operations = []

    def _revert_last_rename(self):
        if not self.last_rename_before_files or not self.last_rename_after_files:
            QMessageBox.information(self, "回退", "没有可回退的操作")
            return
        self.files = self.last_rename_after_files.copy()
        self._update_file_list()
        self.preview_list.clear()
        for i, fp in enumerate(self.last_rename_before_files):
            is_folder = os.path.isdir(fp)
            self.preview_list.add_file_item(i + 1, os.path.basename(fp), is_folder)
        self.has_previewed = True
        QMessageBox.information(self, "回退", "已回退预览，请点击'执行重命名'正式回退。")

    # =====================================================================
    # 手动重命名 / 单条规则
    # =====================================================================

    def preview_manual_rename(self, row):
        if row < 0 or row >= len(self.files):
            return
        file_path = self.files[row]

        if file_path in self.single_rules:
            reply = QMessageBox.question(
                self, "提示", "该文件已设单条规则，手动重命名将撤销该规则，继续？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return
            del self.single_rules[file_path]
            self._update_file_list()
            self.preview_rename()

        current_text = self.preview_list.manual_renamed_items.get(file_path)
        if current_text is None:
            item = self.preview_list.item(row)
            widget = self.preview_list.itemWidget(item) if item else None
            current_text = widget.current_filename if widget else os.path.basename(file_path)

        dialog = RenameDialog(current_text, self)
        if dialog.exec_() != QDialog.Accepted:
            return
        new_text = dialog.get_new_name()
        if has_illegal_chars(new_text):
            QMessageBox.warning(self, "错误", "不能包含无效字符！")
            return

        self.preview_list.manual_renamed_items[file_path] = new_text
        self.preview_rename()

    def edit_single_rule(self, row, file_path):
        if file_path in self.preview_list.manual_renamed_items:
            reply = QMessageBox.question(
                self, "提示", "该文件已手动重命名，撤销后再设单条规则？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.preview_list.manual_renamed_items.pop(file_path, None)
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
                self._update_file_list()
                if self.has_previewed:
                    self.preview_rename()

    def clear_single_rule(self, file_path):
        if file_path in self.single_rules:
            del self.single_rules[file_path]
            self._update_file_list()
            if self.has_previewed:
                self.preview_rename()

    def clear_all_single_rules(self):
        if not self.single_rules:
            return
        if QMessageBox.question(
            self, "确认", "清除所有单条规则？",
            QMessageBox.Yes | QMessageBox.No
        ) == QMessageBox.Yes:
            self.single_rules = {}
            self._update_file_list()
            if self.has_previewed:
                self.preview_rename()

    def reset_manual_renamed(self, row):
        if row < len(self.files):
            fp = self.files[row]
            if fp in self.preview_list.manual_renamed_items:
                del self.preview_list.manual_renamed_items[fp]
                self.preview_rename()

    def reset_all_manual_renamed(self):
        if not self.preview_list.manual_renamed_items:
            return
        if QMessageBox.question(
            self, "确认", "撤销所有手动重命名？",
            QMessageBox.Yes | QMessageBox.No
        ) == QMessageBox.Yes:
            self.preview_list.manual_renamed_items = {}
            self.preview_rename()

    # =====================================================================
    # 冲突自动解决
    # =====================================================================

    def _auto_resolve(self, new_names):
        """重名项添加后缀 _1, _2, ..."""
        groups = {}
        for i, nn in enumerate(new_names):
            dp = os.path.dirname(self.files[i])
            key = f"{dp}#{nn}"
            groups.setdefault(key, []).append(i)
        resolved = new_names.copy()
        for key, indices in groups.items():
            if len(indices) > 1:
                for j, idx in enumerate(indices):
                    if j > 0:
                        base, ext_part = os.path.splitext(resolved[idx])
                        resolved[idx] = f"{base}_{j}{ext_part}"
        return resolved

    # =====================================================================
    # 多步骤重命名执行
    # =====================================================================

    def _perform_rename(self, new_names):
        success_count = 0
        error_count = 0
        unchanged_count = 0
        error_messages = []

        dir_groups = {}
        for i, fp in enumerate(self.files):
            dp = os.path.dirname(fp)
            dir_groups.setdefault(dp, []).append(i)

        for dp, indices in dir_groups.items():
            current_names = [os.path.basename(self.files[i]) for i in indices]
            target_names = [new_names[i] for i in indices]

            rename_needed = []
            temp_mappings = {}

            for j, idx in enumerate(indices):
                cn = current_names[j]
                tn = target_names[j]
                if cn == tn:
                    unchanged_count += 1
                    continue
                rename_needed.append((idx, cn, tn))

            if not rename_needed:
                continue

            # 第一步：冲突项先改临时名
            for idx, cn, tn in rename_needed:
                fp = self.files[idx]
                if tn in current_names and tn != cn:
                    temp_name = f"__TEMP_{uuid.uuid4().hex[:8]}_{tn}"
                    temp_path = os.path.join(dp, temp_name)
                    try:
                        os.rename(fp, temp_path)
                        temp_mappings[idx] = (temp_path, tn, fp)
                        self.files[idx] = temp_path
                        current_names[indices.index(idx)] = temp_name
                        self.last_rename_operations.append({
                            'old_path': fp, 'new_path': temp_path,
                            'old_name': cn, 'new_name': temp_name
                        })
                    except Exception as e:
                        error_count += 1
                        error_messages.append(f"临时重命名失败 {cn}: {e}")

            # 第二步：不冲突项直接改名
            for idx, cn, tn in rename_needed:
                if idx in temp_mappings:
                    continue
                fp = self.files[idx]
                new_path = os.path.join(dp, tn)
                try:
                    self.single_rules.pop(fp, None)
                    self.preview_list.manual_renamed_items.pop(fp, None)
                    self.last_rename_operations.append({
                        'old_path': fp, 'new_path': new_path,
                        'old_name': cn, 'new_name': tn
                    })
                    os.rename(fp, new_path)
                    self.files[idx] = new_path
                    self.last_rename_after_files.append(new_path)
                    success_count += 1
                    current_names[indices.index(idx)] = tn
                except Exception as e:
                    error_count += 1
                    error_messages.append(f"重命名失败 {cn}: {e}")

            # 第三步：临时名改为最终名
            for idx, (temp_path, tn, original_path) in temp_mappings.items():
                final_path = os.path.join(dp, tn)
                try:
                    self.single_rules.pop(original_path, None)
                    self.preview_list.manual_renamed_items.pop(original_path, None)
                    self.last_rename_operations.append({
                        'old_path': temp_path, 'new_path': final_path,
                        'old_name': os.path.basename(temp_path), 'new_name': tn
                    })
                    os.rename(temp_path, final_path)
                    self.files[idx] = final_path
                    self.last_rename_after_files.append(final_path)
                    success_count += 1
                except Exception as e:
                    error_count += 1
                    error_messages.append(f"最终重命名失败 → {tn}: {e}")

        return success_count, error_count, unchanged_count, error_messages


# =============================================================================
# 全局异常处理
# =============================================================================

def excepthook(exc_type, exc_value, exc_traceback):
    error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    print(error_msg)
    try:
        dialog = QMessageBox()
        dialog.setWindowTitle("程序错误")
        dialog.setText(f"程序发生错误: {exc_value}")
        dialog.setDetailedText(error_msg)
        dialog.setIcon(QMessageBox.Critical)
        dialog.exec_()
    except Exception:
        pass


# =============================================================================
# 入口
# =============================================================================

if __name__ == "__main__":
    print("[启动] 正在初始化...")
    sys.excepthook = excepthook
    app = QApplication(sys.argv)
    print("[启动] QApplication 已创建")
    try:
        window = FileRenamerApp()
        print("[启动] 主窗口已创建")
        window.show()
        print("[启动] 窗口已显示，进入事件循环")
        sys.exit(app.exec_())
    except Exception as e:
        print(f"[致命错误] {e}")
        traceback.print_exc()