from openpyxl import load_workbook
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import random
import pandas
import os

# <-----------------------------------WORKING-----------------------------------> #


def search_file():
    '''Get credentials by searching for the website.'''
    website_search = website_input.get().title()
    path = path_input.get()
    # Check if the file exits on the computer
    if os.path.exists(path):
        data = pandas.read_excel(
            path, index_col=False)
        df = pandas.DataFrame(data)
        email = df["Email"][df["Website"] == website_search].values
        password = df["Password"][df["Website"] == website_search].values
        username = df["Username"][df["Website"] == website_search].values
        for i in range(len(email)):
            messagebox.showinfo(
                f"{website_search}", f"Email: {email[i]}\nUsername: {username[i]}\nPassword: {password[i]}")
    else:
        messagebox.showwarning(
            "Warning", "File not found.\nMake sure to specify correct file path.")


def generate_pwd():
    '''Generate random password.'''
    password = ""
    letters = "q w e r t y u i o p a s d f g h j k l z x c v b n m".split(
        " ")
    characters = "! @ # $ % ^ & * ( ) > <".split(" ")
    numbers = "1 2 3 4 5 6 7 8 9 0".split(" ")
    for _ in range(4):
        password += random.choice(letters)
    for _ in range(3):
        password += random.choice(characters)
    for _ in range(3):
        password += random.choice(numbers)
    password = ''.join(random.sample(password, len(password)))
    # Clear contents of the Entry widget and insert the password onto it
    password_input.delete(0, END)
    password_input.insert(0, password)
    # Copy password to Clipboard
    root.clipboard_clear()
    root.clipboard_append(password)


def browse_button():
    '''Allow user to select a directory and store it in global var called '"folder_path".'''
    global folder_path
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Excel Files",
                                                      "*.xlsx*"),
                                                     ("All Files",
                                                      "*.*")))
    path_input.delete(0, END)
    path_input.insert(0, filename)


def save_to_file():
    '''Save data to file.'''
    if website_input.get() and email_input.get() and password_input.get() and path_input.get():
        website = website_input.get().title()
        email = email_input.get()
        username = username_input.get()
        password = password_input.get()
        path = path_input.get()
        data_dict = {"Website": [website], "Email": [
            email], "Username": [username], "Password": [password]}

        df = pandas.DataFrame(data_dict)
        # Read existing file
        reader = pandas.read_excel(r'{0}'.format(path))
        # Check if file is empty
        if reader.empty:
            # Create a Pandas Excel writer using XlsxWriter as the engine.
            writer = pandas.ExcelWriter(path, engine='xlsxwriter')
            # Convert the dataframe to an Excel object.
            df.to_excel(writer, sheet_name='Sheet1', index=False, header=True)
            # Close the Pandas Excel writer and output the Excel file.
            writer.save()
        else:
            writer = pandas.ExcelWriter(path, engine='openpyxl')
            # Try to open an existing workbook
            writer.book = load_workbook(path)
            # Copy existing sheets
            writer.sheets = dict((ws.title, ws)
                                 for ws in writer.book.worksheets)
            # Append to the Excel Sheet
            df.to_excel(writer, index=False, header=False,
                        startrow=len(reader)+1)

            writer.close()

    else:
        messagebox.showwarning(
            "Warning", "Please make sure to fill all details.")

# <-------------------------------------UI-------------------------------------> #


BLUE = "#caf0f8"
# Screen Setup
root = Tk()
root.resizable(0, 0)
# Set position of window on screen
root.geometry('+%d+%d' % (430, 120))
root.title("Password Manager")
# root.minsize(width=500, height=500)
root.config(padx=50, pady=30, bg=BLUE)
root.iconbitmap('password.ico')

folder_path = StringVar()

# Adding Image
canvas = Canvas(root, width=128, height=128,
                bg=BLUE, highlightthickness=0)
img = PhotoImage(file="password.png")
canvas.create_image(64, 64, image=img)
canvas.grid(row=0, column=1, pady=(0, 30), padx=(18, 0))

# Path Input
path_label = Label(root, text="Path: ", font=(
    "Montserrat", 10), bg=BLUE, highlightthickness=0)
path_label.grid(row=1, column=0)
path_input = Entry(root)
path_input.grid(row=1, column=1, sticky=W+N, padx=5, pady=5)
path_btn = Button(root, text="Browse", font=(
    "Montserrat", 7, "bold"), command=browse_button)
path_btn.grid(row=1, column=2, sticky=W+E, padx=(0, 5))

# Website Input
website_label = Label(root, text="Website: ", font=(
    "Montserrat", 10), bg=BLUE, highlightthickness=0)
website_label.grid(row=2, column=0)
website_input = Entry(root)
website_input.grid(row=2, column=1,
                   sticky=W+N, padx=5, pady=5)
website_search_btn = Button(root, text="Search", font=(
    "Montserrat", 7, "bold"), command=search_file)
website_search_btn.grid(row=2, column=2, sticky=W+E, padx=(0, 5))

# Email Input
email_label = Label(root, text="Email: ", font=(
    "Montserrat", 10), bg=BLUE, highlightthickness=0)
email_label.grid(row=3, column=0)
email_input = Entry(root)
email_input.grid(row=3, column=1, columnspan=2, sticky=W+E+N, padx=5, pady=5)

# Username Input
username_label = Label(root, text="Username: ", font=(
    "Montserrat", 10), bg=BLUE, highlightthickness=0)
username_label.grid(row=4, column=0)
username_input = Entry(root)
username_input.grid(row=4, column=1, columnspan=2,
                    sticky=W+E+N, padx=5, pady=5)

# Password Input
password_label = Label(root, text="Password: ", font=(
    "Montserrat", 10), bg=BLUE, highlightthickness=0)
password_label.grid(row=5, column=0)
password_input = Entry(root)
password_input.grid(row=5, column=1, sticky=W+N, padx=5, pady=5)
generate_pwd_btn = Button(root, text="Generate Password", font=(
    "Montserrat", 7, "bold"), command=generate_pwd)
generate_pwd_btn.grid(row=5, column=2, sticky=W+E, padx=(0, 5))

# Save Button
save_btn = Button(root, text="Save", font=(
    "Montserrat", 7, "bold"), command=save_to_file)
save_btn.grid(row=6, column=1, columnspan=2, sticky=W+E, padx=5, pady=5)

path_info = Label(root, text="Make sure to use .xlsx file format to save data", font=(
    "Montserrat", 10), bg=BLUE, highlightthickness=0)
path_info.grid(row=7, column=0, columnspan=3)

root.mainloop()
