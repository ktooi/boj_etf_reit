import re
import xlrd
from datetime import datetime, timedelta


class EtfReitXlsDriver(object):

    # 100億円単位の記載なので、100億倍する。
    mul = 10000000000

    def __init__(self, filename, **kwargs):
        self.__filename = filename

    def read(self):
        """
        Returns:
            list: 読み込んだ Excel の全シートの全内容。

        Examples:
            >>> erxd = EtfReitXlsDriver(filename="/path/to/etfreit.xlsx")
            >>> erxd.read()
            >>> [
            >>>     {
            >>>         "trade_date": "2020/01/01",
            >>>         "etf": {
            >>>             "support": 10000000000,
            >>>             "other": 0
            >>>         },
            >>>         "reit": 20000000000
            >>>     },
            >>>     ...
            >>> ]
        """
        wb = self.workbook

        r = []
        for ws in wb.sheets():

            min_row = self.get_min_row(ws)
            max_row = self.get_max_row(ws)
            min_col = self.get_min_column(ws)
            # max_col = self.get_max_column(ws)
            for row in range(min_row, max_row):
                record = {}
                date = self.excel_date_to_iso(ws.cell(row, min_col).value)
                other = self.to_int(ws.cell(row, min_col + 1).value)
                support = self.to_int(ws.cell(row, min_col + 2).value)
                reit = self.to_int(ws.cell(row, min_col + 3).value)
                record['trade_date'] = date
                record['etf'] = {'other': other,
                                 'support': support}
                record['reit'] = reit
                r.append(record)
        return r

    def get_min_row(self, ws):
        """ 表のデータ部の開始行番号を戻す関数。

        Returns:
            int: 0を先頭行とした場合の、表のデータ部分の開始行番号。
        """
        return 10  # 現状は決まったセルから開始なので、固定値を戻す。

    def get_min_column(self, ws):
        """ 表のデータ部の開始列番号を戻す関数。

        Returns:
            int: 0を先頭列とした場合の、表のデータ部分の開始列番号。
        """
        return 1  # 現状は決まったセルから開始なので、固定値を戻す。

    def get_max_row(self, ws):
        """ 表のデータ部の最終行番号を戻す関数。

        Returns:
            int: 0を先頭行とした場合の、表のデータ部分の最終行番号。
        """
        return ws.nrows

    def get_max_column(self, ws):
        """ 表のデータ部の最終列番号を戻す関数。

        Returns:
            int: 0を先頭列とした場合の、表のデータ部分の最終列番号。
        """
        return ws.ncols

    def excel_date_to_iso(self, excel_date):
        """ Excel 形式の日付情報を、 ISO8601 拡張形式の日付(YYYY-MM-DD)にして戻す関数。

        Args:
            excel_date(int): Excel 形式の日付情報。

        Returns:
            str: 渡された日付情報を ISO8601 拡張形式の日付(YYYY-MM-DD)に変換した文字列。
        """
        return (datetime(1899, 12, 30) + timedelta(days=excel_date)).date().isoformat()

    def to_int(self, value):
        if value:
            return int(value) * self.mul
        else:
            return 0

    @property
    def workbook(self):
        if not hasattr(self, "_EtfReitXlsDriver__workbook"):
            self.__workbook = xlrd.open_workbook(self.filename)
        return self.__workbook

    @property
    def filename(self):
        return self.__filename