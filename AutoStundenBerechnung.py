import os
import re
from pathlib import Path

import openpyxl

RESET = "\033[0m"
BOLD = "\033[01m"
UNDERLINE = "\033[04m"
BLACK = "\033[30m"
RED = "\033[31m"
GREEN = "\033[32m"
ORANGE = "\033[33m"


class AutoStundenBerechnung:
    def __init__(self, wochenarbeitszeit_soll=31.0) -> None:
        self.root_path = Path.cwd()
        self.wochenarbeitszeit_soll = wochenarbeitszeit_soll
        os.system("color")

    def get_excellist_path(self) -> Path:
        """_summary_

        Returns:
            Path: _description_

        Yields:
            Iterator[Path]: _description_
        """
        for xlsm_path in self.root_path.glob("*.xlsx"):
            pattern = re.search(".*KW [\d]*-[\d]*.xlsx", str(xlsm_path))
            if pattern:
                yield xlsm_path

    def lese_ist_wochenarbeitszeit(
        self, excel_path: Path, excel_sheet="Stundenaufstellung"
    ) -> list:
        """_summary_

        Args:
            excel_path (Path): _description_
            excel_sheet (str, optional): _description_. Defaults to "Stundenaufstellung".

        Returns:
            list: _description_
        """
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
                    wochenarbeitszeit = (
                        wb[excel_sheet].cell(row=i_row + 1, column=i_col).value
                    )
                    kalenderwoche = (
                        wb[excel_sheet].cell(row=i_row, column=i_col - 1).value
                    )
                if kalenderwoche and wochenarbeitszeit:
                    if (
                        dic := {
                            "kalenderwoche": kalenderwoche,
                            "wochenarbeitszeit": wochenarbeitszeit,
                        }
                    ) not in stundenaufstellung:
                        stundenaufstellung.append(dic)
        return stundenaufstellung

    def berechne_wochenarbeitszeit(self):
        """_summary_"""
        print(f"{UNDERLINE}{BOLD}Stunden werden berechnet ...{RESET}")
        ueberstunden_liste = []
        for excel_path in self.get_excellist_path():
            zeiten = self.lese_ist_wochenarbeitszeit(excel_path)
            for zeit in zeiten:
                ueberstunden = round(
                    float(zeit["wochenarbeitszeit"])
                    - float(self.wochenarbeitszeit_soll),
                    2,
                )
                ueberstunden_liste.append(ueberstunden)
                if ueberstunden > 0:
                    print(
                        f"{GREEN}Du hast in KW {zeit['kalenderwoche']} "
                        f"-> +{BOLD}{ueberstunden:.2f}h{RESET}{GREEN} "
                        f"zu VIEL gearbeitet.{RESET}"
                    )
                else:
                    print(
                        f"{RED}Du hast in KW {zeit['kalenderwoche']} "
                        f"-> {BOLD}{ueberstunden:.2f}h{RESET}{RED} "
                        f"zu WENIG arbeitet.{RESET}"
                    )
        total_ueberstunden = sum(ueberstunden_liste)
        print("-" * 50)
        if total_ueberstunden > 0:
            print(
                f"{GREEN}Deine gesamten Überstunde sind: "
                f"{UNDERLINE}{BOLD}{total_ueberstunden:.2f}h.{RESET}"
            )
        else:
            print(
                f"{RED}Du bist insgesamt mit "
                f"{UNDERLINE}{BOLD}{total_ueberstunden:.2f}h{RESET}{RED} im minus.{RESET}")
        print("-" * 50)


if __name__ == "__main__":
    asb = AutoStundenBerechnung()
    asb.berechne_wochenarbeitszeit()
