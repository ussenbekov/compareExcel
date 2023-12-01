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
        data = list()
        for file in folder_files:
            data += self.parse_file(file)

        df = pd.DataFrame(
            data, columns=["File name", "Full address", "Sheet", "Address", "Formula"]
        )
        df.drop_duplicates(subset=["Formula"], inplace=True)
        df.to_excel("formulas.xlsx", index=False)

    def parse_file(self, file):
        wb = openpyxl.load_workbook(file)
        file_name = file.split("\\")[-1]
        data = list()
        for sheet in wb.worksheets:
            sheet_title = sheet.title
            for row in sheet:
                for cell in row:
                    cell_address = cell.coordinate
                    cell_formula = str(wb[sheet_title][cell_address].value)
                    cell_formula_source = cell_formula

                    if (
                        cell_formula.find("SUBNM") > -1
                        or cell_formula.find("DBRA") > -1
                        or cell_formula.find("ELPAR") > -1
                        or cell_formula.find("TM1USER") > -1
                    ):
                        continue

                    index_view = cell_formula.find("TM1RPTVIEW")
                    index_dbr = cell_formula.find("DBR")
                    index_sn = cell_formula.find("ServerName")

                    if index_view > -1:
                        cell_formula = cell_formula[index_view:]
                    elif index_dbr > -1:
                        cell_formula = cell_formula[index_dbr:]
                    elif index_sn > -1:
                        cell_formula = cell_formula[index_sn:]
                    else:
                        continue

                    cell_formula = cell_formula.split(",")[0]
                    cell_formula = cell_formula.replace("_xll.", "")

                    data.append(
                        [
                            file_name,
                            sheet_title + "!" + cell_address,
                            sheet_title,
                            cell_address,
                            cell_formula,
                            # "'" + cell_formula_source,
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
        r"C:\Users\talgat.ussenbekov\OneDrive - Kaz Minerals Management LLP\Документы\python projects\OTHER\compareExcel\files",
    )
