import openpyxl
import pandas as pd
import os


class CompareExcelFile:
    def __init__(self, file1, file2) -> None:
        self.file1 = file1
        self.file2 = file2
        self.compare_files()

    def compare_files(self):
        cols_var = ["Value"]
        cols_merge = ["Full address", "Sheet", "Address"]
        cols_df = cols_merge + cols_var

        df1 = pd.DataFrame(self.parse_file(self.file1), columns=cols_df)
        df2 = pd.DataFrame(self.parse_file(self.file2), columns=cols_df)
        var_df = df1.merge(df2, how="left", on=cols_merge, suffixes=("_1", "_2"))
        var_df = var_df.fillna("")
        var_df["difference"] = var_df[cols_var[0] + "_1"] != var_df[cols_var[0] + "_2"]
        var_df = var_df[var_df["difference"]]

        if len(var_df) == 0:
            print("OK")
            return

        err_file = "errors.xlsx"

        try:
            var_df.to_excel(err_file, index=False)
        except PermissionError as e:
            print(f"Ошибка. Закройте файл {err_file}")
        except Exception as e:
            print(e)
        else:
            print(f"Ошибка. Проверьте файл: {err_file}")

    def parse_file(self, file):
        wb = openpyxl.load_workbook(file, data_only=True)
        data = list()
        for sheet in wb.worksheets:
            for row in sheet:
                for cell in row:
                    data.append(
                        [
                            sheet.title + "!" + cell.coordinate,
                            sheet.title,
                            cell.coordinate,
                            cell.value,
                        ]
                    )
        return data


if __name__ == "__main__":
    CompareExcelFile(
        "MASTER FILE_1.xlsx",
        "MASTER FILE_2.xlsx",
    )
