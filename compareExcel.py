from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfile
import openpyxl
import pandas as pd
from tkinter.messagebox import showinfo, showerror


class CompareExcel:
    files = [None, None]
    labels = []
    ws = Tk()

    def __init__(self):
        self.ws.title("ExcelCompare")
        w = self.ws.winfo_screenwidth() // 2
        h = self.ws.winfo_screenheight() // 2
        self.ws.columnconfigure(index=1, weight=1)
        self.ws.geometry(f"{w}x{h}")
        self.create_widgets()

    def open_file(self, btn_name):
        index = int(btn_name)

        file_path = askopenfile(mode="r", filetypes=[("Excel Files", "*xlsx")])
        if file_path is not None:
            self.files[index] = file_path.name
            self.labels[index].set(file_path.name)

    def parse_file(self, file):
        wb = openpyxl.load_workbook(file, data_only=True)
        wb_with_formula = openpyxl.load_workbook(file)
        data = list()
        for sheet in wb.worksheets:
            sheet_title = sheet.title
            for row in sheet:
                for cell in row:
                    cell_address = cell.coordinate
                    cell_value = cell.value
                    cell_formula = wb_with_formula[sheet_title][cell_address].value
                    cell_formula = (
                        cell_formula
                        if cell_formula is not None and str(cell_formula).find("=") > -1
                        else ""
                    )

                    data.append(
                        [
                            sheet_title + "!" + cell_address,
                            sheet_title,
                            cell_address,
                            cell_formula,
                            cell_value,
                        ]
                    )
        return data

    def compare_files(self):
        loading = ttk.Progressbar(self.ws, orient="horizontal", mode="indeterminate")
        loading.grid(row=5, column=0, columnspan=3, pady=[30, 10])
        loading.start()

        first_file, second_file = self.files

        if first_file is None or second_file is None:
            showerror(title="Ошибка", message="Select files")
            loading.grid_forget()
            return

        if first_file == second_file:
            showerror(title="Ошибка", message="The files are the same")
            loading.grid_forget()
            return

        cols_var = ["Formula", "Value"]
        cols_merge = ["Full address", "Sheet", "Address"]
        df_cols = cols_merge + cols_var
        df1 = pd.DataFrame(self.parse_file(first_file), columns=df_cols)
        df2 = pd.DataFrame(self.parse_file(second_file), columns=df_cols)

        var_df = df1.merge(df2, how="outer", on=cols_merge, suffixes=("_1", "_2"))
        var_df = var_df.fillna("")

        var_df["var_formula"] = var_df[cols_var[0] + "_1"] != var_df[cols_var[0] + "_2"]
        var_df["var_value"] = var_df[cols_var[1] + "_1"] != var_df[cols_var[1] + "_2"]
        var_df["var_total"] = (var_df["var_formula"] == TRUE) | (
            var_df["var_value"] == TRUE
        )

        # to show norm formulas in the report
        var_df[cols_var[0] + "_1"] = var_df[cols_var[0] + "_1"].apply(
            lambda x: "'" + x if x != "" and str(x)[0] == "=" else x
        )
        var_df[cols_var[0] + "_2"] = var_df[cols_var[0] + "_2"].apply(
            lambda x: "'" + x if x != "" and str(x)[0] == "=" else x
        )

        # difference filter
        var_df = var_df[var_df["var_total"]]

        # drop columns
        var_df = var_df.drop(["var_total", "var_formula", "var_value"], axis=1)

        # rearrange columns
        var_df = var_df[
            [
                "Full address",
                "Sheet",
                "Address",
                "Formula_1",
                "Formula_2",
                # "var_formula",
                "Value_1",
                "Value_2",
                # "var_value",
                # "var_total",
            ]
        ]

        # hide progress bar
        loading.grid_forget()

        if len(var_df) == 0:
            showinfo("Result", "No differences found, files match")
            return

        err_file = "errors.xlsx"
        try:
            var_df.to_excel(err_file, index=False)
        except PermissionError as e:
            showerror(title="Error", message=f"Close the file {err_file}")
            return
        except Exception as e:
            showerror(title="Error", message=e)
            return
        else:
            showinfo(
                "Result",
                f"Differences found in {len(var_df)} rows. Please check error file: {err_file}",
            )

    def create_widgets(self):
        k = 0
        for i in range(len(self.files)):
            label_text = StringVar()
            label_text.set("Choose file in xlsx format")
            self.labels.append(label_text)
            label = ttk.Label(self.ws, textvariable=label_text)
            label.grid(row=k, column=0, columnspan=3, pady=[20, 5])
            button = ttk.Button(
                self.ws, text="Choose File", command=lambda btn=i: self.open_file(btn)
            )
            button.grid(row=k + 1, column=0, columnspan=3)
            k = k + 2

        check_btn = ttk.Button(
            self.ws, text="Compare Files", command=self.compare_files
        )
        check_btn.grid(row=k + 2, column=0, columnspan=3, pady=10)

        self.ws.mainloop()


if __name__ == "__main__":
    CompareExcel()
