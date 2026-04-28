import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
                             QMessageBox, QTabWidget, QAction)
from converter import ExcelConverter
from dictionary import DictManager
from dialogs import UnknownItemsDialog, DictCRUDDialog


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("财务报表自动分类工具")
        self.resize(1000, 750)

        self.converter = ExcelConverter()
        self.dict_manager = DictManager("dictionary.db")
        self.init_ui()
        self.init_menu()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        self.btn_process = QPushButton("上传财务报表 (.xls) 并处理", self)
        self.btn_process.setFixedHeight(60)
        self.btn_process.clicked.connect(self.process_file)
        layout.addWidget(self.btn_process)

        self.tabs = QTabWidget()
        self.table_mat = QTableWidget()
        self.table_pkg = QTableWidget()
        self.tabs.addTab(self.table_mat, "原材料及返工旧药")
        self.tabs.addTab(self.table_pkg, "包装材料")
        layout.addWidget(self.tabs)

        self.btn_export = QPushButton("导出为分表 Excel", self)
        self.btn_export.setFixedHeight(45)
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self.export_result)
        layout.addWidget(self.btn_export)

        self.df_materials = None
        self.df_packaging = None

    def init_menu(self):
        dict_menu = self.menuBar().addMenu('字典管理')
        dict_menu.addAction('下载字典模板').triggered.connect(self.download_template)
        dict_menu.addAction('上传 Excel 导入数据库').triggered.connect(self.upload_dict)
        dict_menu.addAction('备份数据库为 Excel').triggered.connect(self.backup_dict)
        dict_menu.addAction('在线维护数据库').triggered.connect(self.open_crud_dialog)

    def process_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择导出的 xls 文件", "", "Excel (*.xls)")
        if not file_path: return

        temp_xlsx = None
        try:
            temp_xlsx = self.converter.convert_to_xlsx(file_path)
            df = pd.read_excel(temp_xlsx, dtype={'科目编码': str})

            # 1. 清理数据
            df = df[~df.astype(str).apply(lambda x: x.str.contains('合计')).any(axis=1)]
            cols = ['科目编码', '科目名称', '本期贷方发生数量']
            df = df[cols].copy()
            df['本期贷方发生数量'] = pd.to_numeric(df['本期贷方发生数量'], errors='coerce')
            df = df[df['本期贷方发生数量'] != 0].dropna(subset=['本期贷方发生数量'])

            # 2. 匹配数据库
            dict_df = self.dict_manager.load_dict()
            df['科目编码'] = df['科目编码'].astype(str)
            dict_df['科目编码'] = dict_df['科目编码'].astype(str)

            unknown_mask = ~df['科目编码'].isin(dict_df['科目编码'].tolist())
            unknown_items = df[unknown_mask][['科目编码', '科目名称']].drop_duplicates()

            if not unknown_items.empty:
                dialog = UnknownItemsDialog(unknown_items, self)
                if dialog.exec_():
                    self.dict_manager.add_new_items(dialog.result_df)
                    dict_df = self.dict_manager.load_dict()
                else:
                    return

            # 3. 分表逻辑
            final_df = pd.merge(df, dict_df[['科目编码', '分类']], on='科目编码', how='left')
            final_df['分类'] = final_df['分类'].astype(str).str.strip()

            # 原材料 + 返工旧药
            self.df_materials = final_df[final_df['分类'].isin(['原材料', '返工旧药'])]
            # 包装材料
            self.df_packaging = final_df[final_df['分类'] == '包材']

            self.update_table(self.table_mat, self.df_materials)
            self.update_table(self.table_pkg, self.df_packaging)
            self.btn_export.setEnabled(True)
            QMessageBox.information(self, "成功", "报表处理完成，请检查预览。")

        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
        finally:
            if temp_xlsx and os.path.exists(temp_xlsx): os.remove(temp_xlsx)

    def update_table(self, table, df):
        table.clear()
        if df is None or df.empty: return
        table.setColumnCount(len(df.columns));
        table.setRowCount(len(df))
        table.setHorizontalHeaderLabels(df.columns)
        for i in range(len(df)):
            for j in range(len(df.columns)):
                table.setItem(i, j, QTableWidgetItem(str(df.iloc[i, j])))

    def export_result(self):
        path, _ = QFileDialog.getSaveFileName(self, "导出 Excel", "结果.xlsx", "Excel (*.xlsx)")
        if path:
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                if self.df_materials is not None:
                    self.df_materials.to_excel(writer, sheet_name='原材料及返工旧药', index=False)
                if self.df_packaging is not None:
                    self.df_packaging.to_excel(writer, sheet_name='包材', index=False)
            QMessageBox.information(self, "成功", "文件已成功导出。")

    def download_template(self):
        p, _ = QFileDialog.getSaveFileName(self, "下载模板", "字典模板.xlsx", "Excel (*.xlsx)")
        if p: pd.DataFrame(columns=['科目编码', '科目名称', '分类']).to_excel(p, index=False)

    def upload_dict(self):
        p, _ = QFileDialog.getOpenFileName(self, "上传字典", "", "Excel (*.xlsx *.xls)")
        if p and self.dict_manager.import_from_excel(p):
            QMessageBox.information(self, "成功", "字典已导入 SQLite")

    def backup_dict(self):
        p, _ = QFileDialog.getSaveFileName(self, "备份字典", "备份.xlsx", "Excel (*.xlsx)")
        if p: self.dict_manager.load_dict().to_excel(p, index=False)

    def open_crud_dialog(self):
        DictCRUDDialog(self.dict_manager, self).exec_()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow();
    window.show()
    sys.exit(app.exec_())