import xlwings as xw
import xlwings.constants


class JissaFile:
    def __init__(self, path):
        """
        実査数入力対象ファイルを開く
        :param path: 対象エクセルファイルのパス
        """
        self.wb = xw.Book(path)
        self.ws = self.wb.sheets.active
        self.is_paid = "(有)" in self.wb.name

    def get_jissa_su(self, zuban: str) -> float:
        """
        実査登録確認リストを検索し、発見したらセルを黄色く塗り、実査数を返す
        :param zuban: 検索対象図番
        :return: float: 実査数
        """
        try:
            found = self.ws.api.Columns("A").Find(What=zuban, LookAt=xw.constants.LookAt.xlWhole)
            jissa_su = self.ws.range(f"I{found.Row}")
            jissa_su.select()
            jissa_su.color = 255, 255, 0
            return jissa_su.value
        except AttributeError:
            print(f"{zuban} is not found.")
