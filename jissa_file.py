import xlwings as xw
import xlwings.constants


class JissaFile:
    def __init__(self, path):
        """
        実査数入力対象ファイルを開く。
        有償・無償でPI-Netの画面を切り替えるためファイル名でis_paidを設定する。
        :param path: 対象エクセルファイルのパス
        """
        self.wb = xw.Book(path)
        self.ws = self.wb.sheets.active
        self.is_paid = "(有)" in self.wb.name

    def get_inventory_quantity(self, zuban: str) -> float:
        """
        実査登録確認リストを検索し、発見したらセルを黄色く塗り、実査数を返す
        :param zuban: 検索対象図番
        :return: float: 実査数
        """
        try:
            found = self.ws.api.Columns("A").Find(What=zuban, LookAt=xw.constants.LookAt.xlWhole)
            inventory_quantiry_cell = self.ws.range(f"I{found.Row}")
            inventory_quantiry_cell.select()
            inventory_quantiry_cell.color = 255, 255, 0
            return inventory_quantiry_cell.value
        except AttributeError:
            print(f"{zuban} is not found.")
