from datetime import datetime, date, timedelta
from tkinter import filedialog
from tkinter import ttk
from tkinter import *
import tkinter as tk
import openpyxl
import random
import os

root = tk.Tk()
root.title("ExcelAccounts")

# Frames -----------------------------------------------
main_frame = Frame(root, width=630, height=300)
help_frame = Frame(root, width=630, height=300)

# Entry ------------------------------------------------
show_folder_path_entry = Entry(main_frame, width=40, borderwidth=2)
show_save_folder_entry = Entry(main_frame, width=40, borderwidth=2)

end_year_entry = Entry(main_frame, width=10, borderwidth=2)
end_month_entry = Entry(main_frame, width=10, borderwidth=2)
end_day_entry = Entry(main_frame, width=10, borderwidth=2)

start_year_entry = Entry(main_frame, width=10, borderwidth=2)
start_month_entry = Entry(main_frame, width=10, borderwidth=2)
start_day_entry = Entry(main_frame, width=10, borderwidth=2)

# Labels -----------------------------------------------
start_year_label = Label(main_frame, text="Set starting year:")
start_month_label = Label(main_frame, text="Set starting month(1-12):")
start_day_label = Label(main_frame, text="Set starting day(1-31):")

end_year_label = Label(main_frame, text="Set ending year:", anchor="w")
end_month_label = Label(main_frame, text="Set ending month(1-12):", anchor="w")
end_day_label = Label(main_frame, text="Set ending day(1-31):", anchor="w")

info_label = Label(help_frame, text="""-Starting date cannot be after ending date
-Starting date is automatically included in process making
-In case you want accounts made till a certain date, just insert day after
-Please use selecting buttons for finding paths, otherwise program won't work
-When you fill out all information program needs you can start program
-First number for working hours must be smaller than the second one
-Button 'Delete all' will delete all info inserted and you can start again""", justify=LEFT)

format_rand_input_label = Label(main_frame, text="Select FROM and TO random numbers to insert for work hours")
format_date_label = Label(main_frame, text="Format date: (day(d), month(m), year(Y))")
done_label = Label(root, text="FINISHED!!!")

# Format date -------------------------------------------
date_format_options = ["d", "m", "Y"]

# Drop Menus --------------------------------------------
drop_menu_1 = ttk.Combobox(main_frame, value=date_format_options, width=5)
drop_menu_1.current(1)
drop_menu_1.bind("<<ComboboxSelected>>")

drop_menu_2 = ttk.Combobox(main_frame, value=date_format_options, width=5)
drop_menu_2.current(0)
drop_menu_2.bind("<<ComboboxSelected>>")

drop_menu_3 = ttk.Combobox(main_frame, value=date_format_options, width=5)
drop_menu_3.current(2)
drop_menu_3.bind("<<ComboboxSelected>>")

rand_format_options = list(range(16))

rand_menu_1 = ttk.Combobox(main_frame, value=rand_format_options, width=5)
rand_menu_1.current(0)
rand_menu_1.bind("<<ComboboxSelected>>")

rand_menu_2 = ttk.Combobox(main_frame, value=rand_format_options, width=5)
rand_menu_2.current(5)
rand_menu_2.bind("<<ComboboxSelected>>")


def find_excel_file():
    """User selects excel account it wants to change and function returns path """
    global path
    path = root.filename = filedialog.askopenfilename(initialdir="/", title="Select Excel File",
                                                      filetypes=(("xlsx files", "*.xlsx"), ("xlsm files", "*.xlsm"),
                                                                 ("xltx files", "*.xltx"), ("xltm files", "*.xltm"),
                                                                 ("All files", "*.*")))
    show_folder_path_entry.insert(0, f"{path}")


def save_folder():
    """User selects folder to save all accounts made by program"""
    global save
    save = root.filename = filedialog.askdirectory(initialdir='/', title="Select save folder")
    os.chdir(str(save))
    show_save_folder_entry.insert(0, f"{save}")


def look_in_help():
    """Changes frame for basic tutorial/help"""
    hide_all_frames()
    help_frame.place(x=0, y=0)


def get_back():
    """Returns user to main frame"""
    hide_all_frames()
    main_frame.place(x=0, y=0)


def hide_all_frames():
    """Deletes both frames so new frame can be drawn of function call"""
    help_frame.place_forget()
    main_frame.place_forget()


def delete_inside():
    """Deletes all inserted information by user"""
    show_folder_path_entry.delete(0, END)
    show_save_folder_entry.delete(0, END)
    end_year_entry.delete(0, END)
    end_month_entry.delete(0, END)
    end_day_entry.delete(0, END)
    start_year_entry.delete(0, END)
    start_month_entry.delete(0, END)
    start_day_entry.delete(0, END)
    done_label.place_forget()


def day_counter(start_day):
    """ Function that will print out every work day from certain date
    excluding weekend (Saturday and Sunday). """
    day = datetime.strptime(start_day, '%m/%d/%Y').date()
    until_date = date(int(end_year_entry.get()), int(end_month_entry.get()), int(end_day_entry.get()))
    # Format: date(year, month, day) - set by openpyxl
    while day != until_date:
        if day.weekday() not in {5, 6}:
            yield day.strftime(f'%{str(drop_menu_1.get())}/%{str(drop_menu_2.get())}/%{str(drop_menu_3.get())}')
        day += timedelta(days=1)


def make_excel_copy():
    """Executes program when start button is pressed"""
    workbook = openpyxl.load_workbook(path)
    for day in day_counter(f'{str(start_month_entry.get())}/{str(start_day_entry.get())}/{str(start_year_entry.get())}'):  # odavde pravi
        sheet = workbook.get_sheet_by_name('Sheet1')
        n = sheet['C5'].value
        n += 1
        sheet['C5'] = n
        sheet['F5'] = str(day)
        sheet['E9'] = random.randint(int(rand_menu_1.get()), int(rand_menu_2.get()))
        workbook.save(f"{str(n) + '.xlsx'}")
    done_label.place(x=540, y=145)


find_button = Button(main_frame, text="Search for excel file", padx=33, pady=10, command=find_excel_file)
save_folder_button = Button(main_frame, text="Select save folder", padx=38, pady=10, command=save_folder)
start_button = Button(main_frame, text="Start", padx=51, pady=10, command=make_excel_copy)
back_to_main_button = Button(help_frame, text="Back", padx=15, pady=5, command=get_back)
delete_all_button = Button(main_frame, text="Delete all", padx=39, pady=10, command=delete_inside)

# Menu ---------------------------------------------------
help_menu = Menu(root)
root.config(menu=help_menu)
sub_help_menu = Menu(help_menu, tearoff='off')
help_menu.add_cascade(label="Help", menu=sub_help_menu)
# Command to change the frame
sub_help_menu.add_command(label='View help', command=look_in_help)

# Return to main frame -----------------------------------
back_to_main_button.place(x=0, y=130)

# Entry boxes placed -------------------------------------
show_save_folder_entry.place(x=210, y=70)
show_folder_path_entry.place(x=210, y=20)

start_day_entry.place(x=160, y=115)
start_month_entry.place(x=160, y=145)
start_year_entry.place(x=160, y=175)

end_day_entry.place(x=400, y=115)
end_month_entry.place(x=400, y=145)
end_year_entry.place(x=400, y=175)
# Buttons placed ------------------------------------------
find_button.place(x=0, y=10)
save_folder_button.place(x=0, y=60)
start_button.place(x=480, y=10)
delete_all_button.place(x=480, y=60)
# Drop menu boxes placed ----------------------------------
drop_menu_1.place(x=10, y=240)
drop_menu_2.place(x=75, y=240)
drop_menu_3.place(x=140, y=240)

rand_menu_1.place(x=240, y=240)
rand_menu_2.place(x=310, y=240)
# Label placed -------------------------------------------
start_day_label.place(x=0, y=115)
start_month_label.place(x=0, y=145)
start_year_label.place(x=0, y=175)

end_day_label.place(x=240, y=115)
end_month_label.place(x=240, y=145)
end_year_label.place(x=240, y=175)

format_date_label.place(x=0, y=210)
format_rand_input_label.place(x=240, y=210)

info_label.place(x=0, y=0)
# Main loop and start ------------------------------------

main_frame.pack()
root.mainloop()
