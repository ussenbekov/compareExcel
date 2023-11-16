from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfile
from tkinter.messagebox import showinfo, showerror
import pandas as pd


class CompareTxt:
    root = Tk()
    files = [None, None]
    data_in_column = {1: "Первый", 2: "Последний"}
    labels = []
    data_in_column_selected = StringVar(value="")
    skip_header = IntVar()

    def __init__(self):
        self.root.title("Сравнение файлов")
        w = self.root.winfo_screenwidth() // 2
        h = self.root.winfo_screenheight() // 2
        self.root.geometry(f"{w}x{h}")
        self.init_ui()

    def open_file(self, btn_name):
        index = int(btn_name)

        file_path = askopenfile(mode="r")
        if file_path is not None:
            self.files[index] = file_path.name
            self.labels[index].set(file_path.name)

    def get_df_from_file(self, file, skip_header):
        df = pd.read_csv(
            file,
            header=None,
            skiprows=skip_header,
            decimal=",",
            encoding="ISO-8859-1",
            low_memory=False,
        )
        return df

    def compare_files(self):
        first_file, second_file = self.files

        if first_file is None or second_file is None:
            showerror(title="Ошибка", message="Выберите файлы")
            return

        if first_file == second_file:
            showerror(title="Ошибка", message="Файлы одинаковые")
            return

        df1 = self.get_df_from_file(first_file, self.skip_header.get())
        df2 = self.get_df_from_file(second_file, self.skip_header.get())

        len_cols_df1 = len(df1.columns.values.tolist())
        len_cols_df2 = len(df2.columns.values.tolist())

        if len_cols_df1 != len_cols_df2:
            showerror(title="Ошибка", message="Количество столбцов разные")
            return

        values_column = (
            0 if self.data_in_column_selected.get() == "1" else len_cols_df1 - 1
        )

        join_columns = list(df1.columns.values)
        join_columns.remove(values_column)

        var_df = df1.merge(df2, how="left", on=join_columns, suffixes=("_1", "_2"))
        vals_cols = [str(values_column) + "_1", str(values_column) + "_2"]
        var_df[vals_cols[0]] = pd.to_numeric(var_df[vals_cols[0]])
        var_df[vals_cols[1]] = pd.to_numeric(var_df[vals_cols[1]])
        var_df["var"] = round(var_df[vals_cols[0]] - var_df[vals_cols[1]], 5)

        var_df = var_df[var_df["var"] != 0]

        if len(var_df) == 0:
            showinfo(title="Сравнение файлов", message="Разницы не найдены")
            return

        err_file = "errors.xlsx"

        try:
            var_df.to_excel(err_file, index=False)
        except PermissionError as e:
            showerror(title="Ошибка", message=f"Закройте файл {err_file}")
            return
        except Exception as e:
            showerror(title="Ошибка", message=e)
            return
        else:
            showinfo(
                "Сравнение файлов",
                f"Разницы найдены {len(var_df)} строках. Проверьте файл: {err_file}",
            )

    def init_ui(self):
        # выбор файлов
        for i in range(len(self.files)):
            file_label_value = StringVar()
            file_label_value.set("Выберите файл")
            self.labels.append(file_label_value)
            file_label = ttk.Label(self.root, textvariable=file_label_value)
            file_label.pack(padx=5, pady=5)
            file_button = ttk.Button(
                self.root,
                text=f"Выберите файл {i+1}",
                command=lambda btn=i: self.open_file(btn),
            )
            file_button.pack()

        # настройки
        data_in_column_frame = ttk.LabelFrame(self.root, text="Выберите столбец данных")
        data_in_column_frame.pack(pady=[20, 0])

        for item in self.data_in_column.items():
            data_in_column_btn = ttk.Radiobutton(
                data_in_column_frame,
                text=item[1],
                value=item[0],
                variable=self.data_in_column_selected,
            )
            data_in_column_btn.pack(side="left", padx=10)

        skip_header_btn = ttk.Checkbutton(
            self.root,
            text="Пропустить первый строку",
            variable=self.skip_header,
        )
        skip_header_btn.pack(pady=10)

        # кнопки
        compare_btn = ttk.Button(
            self.root, text="Сравнить файлы", command=self.compare_files
        )
        compare_btn.pack(pady=20)

        self.root.mainloop()


if __name__ == "__main__":
    CompareTxt()
