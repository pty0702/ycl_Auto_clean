from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTableWidget,
                             QTableWidgetItem, QPushButton, QComboBox, QLabel,
                             QHeaderView, QMessageBox)
from PyQt5.QtCore import Qt
import pandas as pd


class UnknownItemsDialog(QDialog):
    def __init__(self, unknown_df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("发现新科目 - 请指派分类")
        self.resize(600, 450)
        self.unknown_df = unknown_df
        self.result_df = None
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("以下科目在字典中不存在，请选择归属类别："))

        self.table = QTableWidget(len(unknown_df), 3)
        self.table.setHorizontalHeaderLabels(['科目编码', '科目名称', '归属分类'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.row_widgets = []
        for i, (idx, row) in enumerate(unknown_df.iterrows()):
            self.table.setItem(i, 0, QTableWidgetItem(str(row['科目编码'])))
            self.table.setItem(i, 1, QTableWidgetItem(str(row['科目名称'])))
            combo = QComboBox()
            # 统一使用“返工旧药”
            combo.addItems(['原材料', '包材', '返工旧药', '其他'])
            self.table.setCellWidget(i, 2, combo)
            self.row_widgets.append((row['科目编码'], row['科目名称'], combo))

        layout.addWidget(self.table)
        confirm_btn = QPushButton("确定并更新数据库")
        confirm_btn.clicked.connect(self.collect_and_accept)
        layout.addWidget(confirm_btn)

    def collect_and_accept(self):
        data = [{'科目编码': str(c), '科目名称': n, '分类': cb.currentText()} for c, n, cb in self.row_widgets]
        self.result_df = pd.DataFrame(data)
        self.accept()


class DictCRUDDialog(QDialog):
    def __init__(self, dict_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("本地数据库在线维护")
        self.resize(700, 550)
        self.dict_manager = dict_manager
        layout = QVBoxLayout(self)
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['科目编码', '科目名称', '分类'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        btn_layout = QHBoxLayout()
        add_btn = QPushButton("新增")
        del_btn = QPushButton("删除")
        save_btn = QPushButton("保存修改")
        add_btn.clicked.connect(lambda: self.table.insertRow(self.table.rowCount()))
        del_btn.clicked.connect(lambda: self.table.removeRow(self.table.currentRow()))
        save_btn.clicked.connect(self.save_data)
        btn_layout.addWidget(add_btn);
        btn_layout.addWidget(del_btn);
        btn_layout.addStretch();
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)
        self.load_data()

    def load_data(self):
        df = self.dict_manager.load_dict()
        self.table.setRowCount(len(df))
        for i, row in df.iterrows():
            self.table.setItem(i, 0, QTableWidgetItem(str(row['科目编码'])))
            self.table.setItem(i, 1, QTableWidgetItem(str(row['科目名称'])))
            self.table.setItem(i, 2, QTableWidgetItem(str(row['分类'])))

    def save_data(self):
        data = []
        for i in range(self.table.rowCount()):
            code = self.table.item(i, 0)
            if code and code.text().strip():
                data.append({
                    '科目编码': code.text().strip(),
                    '科目名称': self.table.item(i, 1).text() if self.table.item(i, 1) else "",
                    '分类': self.table.item(i, 2).text() if self.table.item(i, 2) else ""
                })
        if self.dict_manager.save_dict(pd.DataFrame(data)):
            QMessageBox.information(self, "成功", "数据库已更新")
            self.accept()