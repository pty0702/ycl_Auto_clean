import pandas as pd
import xlrd
import os


class ExcelConverter:
    """Excel 格式转换工具：支持 .xls 到 .xlsx 的安全转换并修复乱码"""

    def __init__(self):
        self.encodings_to_try = [None, 'gbk', 'gb18030', 'gb2312', 'cp1252']
        self.garble_patterns = ['锟斤拷', '烫烫烫', '锟', '']

    def _is_garbled(self, text):
        if pd.isna(text): return False
        text = str(text)
        if '\ufffd' in text: return True
        for p in self.garble_patterns:
            if p and p in text: return True
        latin1_high_count = sum(1 for ch in text if '\u00c0' <= ch <= '\u00ff')
        if latin1_high_count > len(text) * 0.3 and len(text) > 5:
            return True
        return False

    def _check_df_garbled(self, df):
        for col in df.columns:
            for val in df[col].head(10):
                if self._is_garbled(val):
                    return True
        return False

    def _read_xls_with_encoding(self, filepath, encoding=None):
        wb = xlrd.open_workbook(filepath, encoding_override=encoding)
        sheet = wb.sheet_by_index(0)
        data = [sheet.row_values(i) for i in range(sheet.nrows)]
        if not data: return pd.DataFrame()
        return pd.DataFrame(data[1:], columns=data[0])

    def convert_to_xlsx(self, xls_path):
        if not xls_path.endswith('.xls'): return xls_path
        new_path = xls_path + 'x'
        for enc in self.encodings_to_try:
            try:
                if enc is None:
                    df = pd.read_excel(xls_path, engine='xlrd')
                else:
                    df = self._read_xls_with_encoding(xls_path, encoding=enc)
                if not self._check_df_garbled(df):
                    df.to_excel(new_path, index=False, engine='openpyxl')
                    return new_path
            except Exception:
                continue
        raise ValueError(f"无法正确解析文件 {xls_path}")