from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfile
import openpyxl
import pandas as pd
from tkinter.messagebox import showinfo, showerror
import os


class CompareExcel:
    files = [None, None]
    labels = []
    root = Tk()
    compare_formulas = IntVar()
    compare_values = IntVar()

    def __init__(self):
        self.root.title("ExcelCompare")
        w = self.root.winfo_screenwidth() // 2
        h = self.root.winfo_screenheight() // 2
        self.root.geometry(f"{w}x{h}")
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
                    temp = [
                        sheet_title + "!" + cell_address,
                        sheet_title,
                        cell_address,
                    ]
                    if self.compare_values.get() == 1:
                        cell_value = cell.value
                        temp.append(cell_value)
                    if self.compare_formulas.get() == 1:
                        cell_formula = wb_with_formula[sheet_title][cell_address].value
                        cell_formula = (
                            cell_formula
                            if cell_formula is not None
                            and str(cell_formula).find("=") > -1
                            else ""
                        )
                        temp.append(cell_formula)

                    data.append(temp)
        return data

    def compare_files(self):
        first_file, second_file = self.files

        if first_file is None or second_file is None:
            showerror(title="Error", message="Select files")
            return

        if first_file == second_file:
            showerror(title="Error", message="The files are the same")
            return

        if self.compare_formulas.get() == 0 and self.compare_values.get() == 0:
            showerror(title="Error", message="Select compare options")
            return

        cols_merge = ["Full address", "Sheet", "Address"]
        cols_var = []
        if self.compare_values.get() == 1:
            cols_var.append("Value")
        if self.compare_formulas.get() == 1:
            cols_var.append("Formula")

        df_cols = cols_merge + cols_var
        df1 = pd.DataFrame(self.parse_file(first_file), columns=df_cols)
        df2 = pd.DataFrame(self.parse_file(second_file), columns=df_cols)

        var_df = df1.merge(df2, how="outer", on=cols_merge, suffixes=("_1", "_2"))
        var_df = var_df.fillna("")

        if self.compare_formulas.get() == 1:
            var_df["var_formula"] = var_df[["Formula_1", "Formula_2"]].apply(
                lambda x: int(x.Formula_1 != x.Formula_2), axis=1
            )
            var_df[["Formula_1", "Formula_2"]] = var_df[
                ["Formula_1", "Formula_2"]
            ].replace(regex={"\=_xll.|\=": ""})

        if self.compare_values.get() == 1:
            var_df[["Value_1", "Value_2"]] = var_df[["Value_1", "Value_2"]].replace(
                regex={",": "."}
            )
            var_df[["Value_1", "Value_2"]] = var_df[["Value_1", "Value_2"]].fillna(0)
            var_df["var_value"] = var_df[["Value_1", "Value_2"]].apply(
                lambda x: (
                    round(round(float(x.Value_1), 3) - round(float(x.Value_2), 3), 3)
                    if (isinstance(x.Value_1, str) and x.Value_1.isdigit() == True)
                    or isinstance(x.Value_1, (int, float))
                    else int(x.Value_2 != x.Value_1)
                ),
                axis=1,
            )

        if self.compare_formulas.get() == 1 and self.compare_values.get() == 1:
            var_df = var_df[(var_df["var_value"] != 0) | (var_df["var_formula"] != 0)]
        elif self.compare_formulas.get() == 1:
            var_df = var_df[var_df["var_formula"] != 0]
        else:
            var_df = var_df[var_df["var_value"] != 0]

        if len(var_df) == 0:
            showinfo("Result", "No differences found, files match")
            return

        err_file = "errors.xlsx"
        try:
            var_df.to_excel(err_file, index=False)
            os.system(f"start EXCEL.EXE {err_file}")
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
        # block choose files
        for i in range(len(self.files)):
            label_text = StringVar()
            label_text.set(f"Choose {i+1} file in xlsx format")
            self.labels.append(label_text)
            label = ttk.Label(self.root, textvariable=label_text)
            label.pack()
            button = ttk.Button(
                self.root,
                text="Choose file",
                command=lambda btn=i: self.open_file(btn),
            )
            button.pack(pady=[0, 20])

        # block compare options
        frame_compare_options = ttk.LabelFrame(self.root, text="Compare")
        frame_compare_options.pack(pady=[0, 20])

        btn_compare_formulas = ttk.Checkbutton(
            frame_compare_options,
            text="Formulas",
            variable=self.compare_formulas,
        )
        btn_compare_formulas.pack(side="left", padx=10, pady=5)

        btn_compare_values = ttk.Checkbutton(
            frame_compare_options,
            text="Values",
            variable=self.compare_values,
        )
        btn_compare_values.pack(side="left", padx=10, pady=5)

        # block compare btn
        btn_compare = ttk.Button(
            self.root, text="Compare Files", command=self.compare_files
        )
        btn_compare.pack()

        self.root.mainloop()


if __name__ == "__main__":
    CompareExcel()
