import sqlite3
import pandas as pd
import os

class DictManager:
    """SQLite 字典管理器"""
    def __init__(self, db_name="dictionary.db"):
        self.db_path = db_name
        self.table_name = "subjects"
        self._init_db()

    def _init_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {self.table_name} (
                code TEXT PRIMARY KEY,
                name TEXT,
                category TEXT
            )
        ''')
        conn.commit()
        conn.close()

    def load_dict(self):
        conn = sqlite3.connect(self.db_path)
        try:
            df = pd.read_sql(f"SELECT code AS 科目编码, name AS 科目名称, category AS 分类 FROM {self.table_name}", conn)
            return df
        finally:
            conn.close()

    def save_dict(self, df):
        conn = sqlite3.connect(self.db_path)
        try:
            save_df = df.copy()
            save_df.columns = ['code', 'name', 'category']
            save_df.to_sql(self.table_name, conn, if_exists='replace', index=False)
            return True
        except Exception:
            return False
        finally:
            conn.close()

    def add_new_items(self, new_items_df):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            for _, row in new_items_df.iterrows():
                cursor.execute(f'''
                    INSERT OR REPLACE INTO {self.table_name} (code, name, category)
                    VALUES (?, ?, ?)
                ''', (str(row['科目编码']), str(row['科目名称']), str(row['分类'])))
            conn.commit()
            return True
        finally:
            conn.close()

    def import_from_excel(self, excel_path):
        try:
            df = pd.read_excel(excel_path, dtype={'科目编码': str})
            df = df[['科目编码', '科目名称', '分类']]
            return self.add_new_items(df)
        except Exception:
            return False