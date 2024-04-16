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
left_frame = tk.Frame(app, bg = "lightblue",padx = 25, pady = 8)
left_frame.grid(row=0, column=0, sticky="n")
right_frame = tk.Frame(app, bg = "red")
right_frame.grid(row=0, column=1, sticky="n")

# Adding import button
import_button = tk.Button(left_frame, text = "Import",width = 50, height = 2, command = controller.importExcel)
import_button.grid(row=0, column=0, pady=(0,8) )

# Adding tkSheet1
initial_sheet = Sheet(right_frame,
              #headers=["Row", "Column"],  # Show row and column headers
              page_up_down_select_row=True,  # Allow selecting rows using page up/down keys
              empty_vertical = 0,  # Prevents empty space between headers and first row
              empty_horizontal = 0,  # Prevents empty space between headers and first column
              column_width=120,  # Width of columns
              row_index_width=60,  # Width of the row index column
              total_rows=5000,  # Total number of rows
              total_columns=100,
              width = 670,
              height = 720
              )

initial_sheet.enable_bindings((
                       "toggle_select",
                       "select_all",  # Select all cells with Ctrl+A,
                       "drag_select",
                       "copy",  # Copy selected cells with Ctrl+C
                       "cut",  # Cut selected cells with Ctrl+X
                       "paste",  # Paste copied/cut cells with Ctrl+V
                       "undo",  # Undo action with Ctrl+Z
                       "redo",  # Redo action with Ctrl+Y
                       "resize",  # Resize column/row with right mouse button
                       "column_select",  # Select entire column with Ctrl+left mouse button
                       "row_select",  # Select entire row with Ctrl+left mouse button
                       "external_drag_drop",  # Enable external drag and drop
                       "box_select",  # Select rectangular area with Ctrl+drag
                       "open_editor",  # Open editor with double click or Enter key
                       "delete",  # Delete selected cells with Delete or Backspace key
                       "deselect",  # Deselect cells with Esc key
                       "edit_cell"
                       ))
initial_sheet.pack()

analyze_button = tk.Button(left_frame, text = "Analyze",width = 50, height = 2, command =lambda: controller.IsAnyData(initial_sheet))
analyze_button.grid(row=1, column=0)

def main():
    # Intiate mainloop
    app.mainloop()

if __name__ == "__main__":
    main()