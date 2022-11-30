import openpyxl
from pathlib import Path


class AutoStundenBerechnung:
    def __init__(self, wochenarbeitszeit_soll = 31) -> None:
        self.root_path = Path.cwd()
        self.wochenarbeitszeit_soll = wochenarbeitszeit_soll

    def get_excellist_path(self) -> list:
        """returns the path of all excel sgvm files

        Returns:
            list: paths
        """
        for xlsm_path in self.root_path.glob("*.xlsx"):
            yield xlsm_path

    def lese_ist_wochenarbeitszeit(self, excel_path, excel_sheet = "Stundenaufstellung"):
        stundenaufstellung = []
        kalenderwoche = None
        wochenarbeitszeit = None
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
        except Exception:
            RuntimeError(f"Bitte Excel {excel_path} schließen!")
        max_col = wb[excel_sheet].max_column + 1
        max_row = wb[excel_sheet].max_row + 1
        for i_row in range(1, max_row):
            for i_col in range(1, max_col):
                cell_obj = wb[excel_sheet].cell(row=i_row, column=i_col)
                if cell_obj.value == "Wochenarbeitszeit":
                    wochenarbeitszeit = wb[excel_sheet].cell(row=i_row+1, column=i_col).value
                    kalenderwoche = wb[excel_sheet].cell(row=i_row, column=i_col-1).value
                if kalenderwoche and wochenarbeitszeit:
                    if (dic := {"kalenderwoche":kalenderwoche, "wochenarbeitszeit":wochenarbeitszeit}) not in stundenaufstellung:
                        stundenaufstellung.append(dic)
        return stundenaufstellung

    def berechne_wochenarbeitszeit(self):
        for excel_path in self.get_excellist_path():
            zeiten = self.lese_ist_wochenarbeitszeit(excel_path)
            print(zeiten)

if __name__ == "__main__":
    asb = AutoStundenBerechnung()
    asb.berechne_wochenarbeitszeit()