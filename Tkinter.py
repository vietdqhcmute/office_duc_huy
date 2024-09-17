import os
import tkinter as tk
import xlwings as xlw

# Set up xlwings
wb = xlw.books.active
ws = wb.sheets["Tong_hop"]

# Create the main window
window = tk.Tk()
window.geometry("200x250")

# Function to perform the task
def process_row():
    row_id = nametb.get()

    if "sơ nhân thân" in str(ws["Y" + str(row_id)].value):
        clone = [text for text in str(ws["Y" + str(row_id)].value).split("\n") if "sơ nhân thân" not in text]
        ws["Y" + str(row_id)].value = "\n".join(clone)

    clone = str(ws["Z" + str(row_id)].value).split("\n")
    clone.append("Hồ sơ nhân thân sao y")
    ws["Z" + str(row_id)].value = "\n".join(clone)

# Create labels, input entry, and button
label_row_id = tk.Label(window, text="Số hàng: ")
entry_row_id = tk.Entry(window)
button_run = tk.Button(window, text="Chạy", command=process_row)

# Set grid positions
label_row_id.grid(row=1, column=0)
entry_row_id.grid(row=1, column=1)
entry_row_id.focus()
button_run.grid(row=2, columnspan=2)

# Run the GUI
window.mainloop()
