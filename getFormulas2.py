import pandas as pd
import glob
import os
import openpyxl


class getFormulas:
    def __init__(self, folder):
        self.base_dir = os.getcwd()
        self.errors_dir = os.path.join(self.base_dir, "errors")
        self.folder = folder
        self.get_formulas()

    def get_formulas(self):
        folder_files = self.get_folder_files(self.folder)
        for file in folder_files:
            file_name = file.split("\\")[-1]
            file_name = file_name[0 : file_name.find(".xlsx")]
            print(file_name)
            exit()
            data = self.parse_file(file)
            df = pd.DataFrame(
                data,
                columns=["Sheet", "Address", "Formula"],
            )
            df.to_csv(file_name + ".csv", index=False)

    def parse_file(self, file):
        wb = openpyxl.load_workbook(file)
        data = list()
        for sheet in wb.worksheets:
            sheet_title = sheet.title
            for row in sheet:
                for cell in row:
                    cell_address = cell.coordinate
                    cell_formula = str(wb[sheet_title][cell_address].value)

                    if cell_formula.startswith("=") == False:
                        continue

                    cell_formula = cell_formula.replace("_xll.", "").replace("=", "", 1)
                    data.append(
                        [
                            sheet_title,
                            cell_address,
                            cell_formula,
                        ]
                    )
        return data

    def get_folder_files(self, folder, type_file="*"):
        files = glob.glob(folder + "\\" + type_file)
        return files

    def get_df_from_file(self, file):
        df = pd.read_csv(
            file,
            encoding="ISO-8859-1",
            low_memory=False,
        )

        return df


if __name__ == "__main__":
    getFormulas(
        r"C:\Users\talgat.ussenbekov\OneDrive - Kaz Minerals Management LLP\Документы\_python projects\OTHER\compareExcel\files",
    )
