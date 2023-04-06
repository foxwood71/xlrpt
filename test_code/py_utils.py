import os
import pathlib
import tkinter
import tkinter.filedialog
import tkinter.messagebox
from datetime import datetime

from dateutil.relativedelta import relativedelta


def select_folder():
    path = tkinter.filedialog.askdirectory()
    return path


def select_file():
    filename = tkinter.filedialog.askopenfilename(
        title="Select file", filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*"))
    )
    return filename


def work_file(file):
    try:
        # path, name = os.path.split(src_path)
        directory = file.parent  # rename
        file_name_full = file.name  # test10.png
        file_name = file.stem  # test10
        file_extension = file.suffix  # .png

        # file rename
        old_date = datetime.strptime(file.stem, "%Y%m%d").date()
        new_date = old_date - relativedelta(days=1)
        old_name = file.name
        new_name = new_date.strftime("%Y%m%d") + file.suffix
        print(old_name, " ==> ", new_name)
        file.rename(directory / new_name)

    except Exception as e:
        tkinter.messagebox.showinfo("Exception!", str(e))


def work_folder(folder_path, is_sub):
    try:
        path = pathlib.Path(folder_path)
        print("print path : ", path)
        file_count = len([f for f in path.iterdir()])
        file_count_len = len(str(file_count))
        print(f"file_count: {file_count}\n len: {file_count_len}")

        cnt = 1
        # for file in path.iterdir():
        # pathlib.glob regex pattern 사용불가
        # glob.glob regex patten 사용가능
        for file in path.glob("**/????????.xlsx"):  # ** -> recursive
            if not file.is_dir():  # rename\test10.png
                directory = file.parent  # rename
                file_name_full = file.name  # test10.png
                file_name = file.stem  # test10
                file_extension = file.suffix  # .png
                work_file(file)

                # if file.is_file():
                #     new_filename = file_name + str(cnt).zfill(file_count_len) + file_extension
                #     file.rename(path / new_filename)

                # cnt = +1

    except Exception as e:
        tkinter.messagebox.showinfo("Exception!", str(e))


def work_xl_file(file: pathlib):
    try:
        directory = file.parent  # rename
        file_name_full = file.name  # test10.png
        file_name = file.stem  # test10
        file_extension = file.suffix  # .png

        print(file_name_full)

    except Exception as e:
        tkinter.messagebox.showinfo("Exception!", str(e))


if __name__ == "__main__":

    try:
        # work_path = select_folder()
        # work_folder(work_path, True)
        filename = select_file()
        work_xl_file(pathlib.Path(filename))
        os.system("pause")

    except Exception as e:
        tkinter.messagebox.showinfo("Exception!", str(e))
        # with open('Log.txt', 'a', encoding='utf-8', errors='ignore') as f:
        #    f.write(e)
