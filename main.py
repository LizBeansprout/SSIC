import tkinter as tk
from tksheet import Sheet
import controller

# Intiate tkinter entity
app = tk.Tk()

# Setting window size
app.geometry("1080x720")
app.minsize(1080,720)
app.maxsize(1080,720)

# Setting title
app.title("Single-Stage Inventory Control")

# Creating menu bar
menu_bar = tk.Menu(app)
app.config(menu = menu_bar)

# Adding major menu
file_menu = tk.Menu(menu_bar, tearoff = 0)
edit_menu = tk.Menu(menu_bar, tearoff = 0)
view_menu = tk.Menu(menu_bar, tearoff = 0)
option_menu = tk.Menu(menu_bar, tearoff = 0)
menu_bar.add_cascade(label = "File", menu = file_menu)
menu_bar.add_cascade(label = "Edit", menu = edit_menu)
menu_bar.add_cascade(label = "View", menu = view_menu)
menu_bar.add_cascade(label = "Option", menu = option_menu)

# Adding command for each major menu
  # File menu
file_menu.add_cascade(label = 'Something')
file_menu.add_separator()
file_menu.add_cascade(label = 'Exit')
  # Edit menu
edit_menu.add_cascade(label = 'Something')
  # View menu
view_menu.add_cascade(label = 'Something')
  # Option menu
option_menu.add_cascade(label = 'Something')

# Frame management
left_frame = tk.Frame(app, bg = "lightblue", width = 360, padx = 25, pady = 10)
left_frame.grid(row=0, column=0)
right_frame = tk.Frame(app, bg = "red", width = 720, padx = 10, pady = 10)
right_frame.grid(row=0, column=1)

# Adding import button
import_button = tk.Button(left_frame, text = "Import", padx = 150, pady = 8, command = controller.importExcel)
import_button.grid(row=0,column=0)

# Adding tkSheet1
main_sheet = Sheet(right_frame,data = [[1],[2],[3]])
main_sheet.pack()
def main():
    # Intiate mainloop
    app.mainloop()

if __name__ == "__main__":
    main()