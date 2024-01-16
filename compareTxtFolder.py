import pandas as pd
import glob
import os


class CompareTxtFolder:
    def __init__(self, folder1, folder2, skip_header=1, values_column="last"):
        self.base_dir = os.getcwd()
        self.errors_dir = os.path.join(self.base_dir, "errors")
        self.folder1 = folder1
        self.folder2 = folder2
        self.values_column = values_column
        self.skip_header = skip_header
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
        filename_1 = file_1.split("\\")[-1]
        message = f"{filename_1}"

        df1 = self.get_df_from_file(file_1, self.skip_header)
        df2 = self.get_df_from_file(file_2, self.skip_header)

        len_cols_df1 = len(df1.columns.values.tolist())
        len_cols_df2 = len(df2.columns.values.tolist())

        if len_cols_df1 != len_cols_df2:
            print(f"{message}: количество столбцов разные")
            return

        values_column = 0 if self.values_column == "first" else len_cols_df1 - 1

        join_columns = list(df1.columns.values)
        join_columns.remove(values_column)

        var_df = df1.merge(df2, how="left", on=join_columns, suffixes=("_1", "_2"))
        vals_cols = [str(values_column) + "_1", str(values_column) + "_2"]
        var_df[vals_cols] = var_df[vals_cols].replace(regex={"\+|\(|\)|\/|\ |NaN": ""})
        var_df[vals_cols] = var_df[vals_cols].replace(regex={",": "."})
        var_df[vals_cols] = var_df[vals_cols].fillna(0)
        var_df["var"] = var_df.apply(
            lambda x: round(float(x[vals_cols[0]]) - float(x[vals_cols[1]]), 5)
            if (isinstance(x[vals_cols[0]], str) and x[vals_cols[0]].isdigit() == True)
            or isinstance(x[vals_cols[0]], (int, float))
            else int(x[vals_cols[1]] != x[vals_cols[0]]),
            axis=1,
        )

        var_df = var_df[var_df["var"] != 0]

        if len(var_df) == 0:
            print(f"{message}: OK")
            return

        err_file = os.path.join(self.errors_dir, f"{message}_errors.xlsx")

        try:
            var_df.to_excel(err_file, index=False)
        except PermissionError as e:
            print(f"{message}: закройте файл {err_file}")
            return
        except Exception as e:
            print(f"{message}: {e}")
            return
        else:
            print(f"{message}: Ошибка. Проверьте файл: {err_file}")

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


if __name__ == "__main__":
    CompareTxtFolder(
        r"C:\Users\talgat.ussenbekov\Downloads\shs\do",
        r"C:\Users\talgat.ussenbekov\Downloads\shs\posle",
    )
