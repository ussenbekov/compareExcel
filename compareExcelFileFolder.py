import openpyxl
import pandas as pd
import glob
import os


class CompareExcelFile:
    def __init__(self, folder1, folder2) -> None:
        self.base_dir = os.getcwd()
        self.errors_dir = os.path.join(self.base_dir, "errors")
        self.folder1 = folder1
        self.folder2 = folder2
        self.compare_folder_files()

    def compare_folder_files(self):
        folder1_files = self.get_folder_files(self.folder1)
        for file_1 in folder1_files:
            filename_1 = file_1.split("\\")[-1]
            file_2 = self.folder2 + "\\" + filename_1
            self.compare_files(file_1, file_2)

    def get_folder_files(self, folder, type_file="*"):
        files = glob.glob(folder + "\\" + type_file)
        return files

    def compare_files(self, file_1, file_2):
        message = file_1.split("\\")[-1]
        cols_var = ["Value"]
        cols_merge = ["Full address", "Sheet", "Address"]
        cols_df = cols_merge + cols_var

        df1 = pd.DataFrame(self.parse_file(file_1), columns=cols_df)
        df2 = pd.DataFrame(self.parse_file(file_2), columns=cols_df)

        var_df = df1.merge(df2, how="left", on=cols_merge, suffixes=("_1", "_2"))
        var_df = var_df.fillna("")
        var_df["difference"] = var_df[cols_var[0] + "_1"] != var_df[cols_var[0] + "_2"]
        var_df = var_df[var_df["difference"]]

        if len(var_df) == 0:
            print("OK")
            return

        err_file = os.path.join(self.errors_dir, f"{message}_errors.xlsx")

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
        r"do",
        r"posle",
    )
