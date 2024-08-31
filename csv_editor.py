import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from tkinter import ttk
import os
from tkinter.font import Font

class CSVEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Editor")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")  # Light gray background

        # Variable to store the next column index
        self.next_col_index = None

        # Import button
        self.import_btn = tk.Button(root, text="Import File", command=self.import_file, font=('Arial', 12), bg="#007bff", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.import_btn.pack(pady=10)

        # Treeview for displaying the data
        self.tree = ttk.Treeview(root, style="Custom.Treeview", selectmode='browse')
        self.tree.pack(expand=True, fill=tk.BOTH, side=tk.LEFT)

        # Style configuration for Treeview
        style = ttk.Style()
        style.configure("Custom.Treeview", background="#ffffff", foreground="#000000", rowheight=25, fieldbackground="#ffffff")
        style.configure("Custom.Treeview.Heading", background="#007bff", foreground="#ffffff", font=('Arial', 12, 'bold'))

        # Add vertical scrollbar
        self.vsb = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.vsb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=self.vsb.set)

        # Add horizontal scrollbar
        self.hsb = ttk.Scrollbar(root, orient="horizontal", command=self.tree.xview)
        self.hsb.pack(side='bottom', fill='x')
        self.tree.configure(xscrollcommand=self.hsb.set)

        # Frame for buttons (Edit, Delete, Add Row, Add Column)
        self.button_frame = tk.Frame(root, bg="#f0f0f0")
        self.button_frame.pack(fill=tk.X, pady=5)

        self.edit_btn = tk.Button(self.button_frame, text="Edit", command=self.edit_row, font=('Arial', 12), bg="#28a745", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.edit_btn.pack(side=tk.LEFT, padx=5)

        self.delete_btn = tk.Button(self.button_frame, text="Delete", command=self.delete_row, font=('Arial', 12), bg="#dc3545", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.delete_btn.pack(side=tk.LEFT, padx=5)

        self.add_row_btn = tk.Button(self.button_frame, text="Add Row", command=self.add_row, font=('Arial', 12), bg="#17a2b8", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.add_row_btn.pack(side=tk.LEFT, padx=5)

        self.add_col_btn = tk.Button(self.button_frame, text="Add Column", command=self.add_column, font=('Arial', 12), bg="#ffc107", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.add_col_btn.pack(side=tk.LEFT, padx=5)

        # Separate Save button
        self.save_frame = tk.Frame(root, bg="#f0f0f0")
        self.save_frame.pack(fill=tk.X, pady=5)

        self.save_btn = tk.Button(self.save_frame, text="Save", command=self.save_file, font=('Arial', 12), bg="#28a745", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.save_btn.pack()

        # Bind keyboard events
        self.tree.bind("<Delete>", self.delete_column)

    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if file_path:
            if file_path.endswith(".csv"):
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)
            self.display_data()

    def display_data(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["column"] = list(self.df.columns)
        self.tree["show"] = "headings"

        # Create a Font object
        font = Font()

        # Adjust the column headers
        for col in self.tree["column"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=font.measure(col))

        # Insert rows and adjust the column widths based on content
        for index, row in self.df.iterrows():
            values = list(row)
            self.tree.insert("", "end", values=values)
            for i, value in enumerate(values):
                col_width = font.measure(str(value)) + 10  # Adjust for padding
                if self.tree.column(self.tree["columns"][i], 'width') < col_width:
                    self.tree.column(self.tree["columns"][i], width=col_width)

    def delete_row(self):
        selected_item = self.tree.selection()
        if selected_item:
            for item in selected_item:
                row_index = self.tree.index(item)
                self.df = self.df.drop(self.df.index[row_index])
                self.df.reset_index(drop=True, inplace=True)
                self.tree.delete(item)

    def add_row(self):
        # Logic to add a new row
        pass

    def add_column(self):
        # Logic to add a new column
        pass

    def edit_row(self):
        # Logic to edit a row
        pass

    def save_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if file_path:
            _, ext = os.path.splitext(file_path)
            if ext == ".csv":
                self.df.to_csv(file_path, index=False)
            elif ext == ".xlsx":
                self.df.to_excel(file_path, index=False)
            messagebox.showinfo("File Saved", f"File has been saved as {file_path}")
        else:
            messagebox.showwarning("Save Cancelled", "No file was saved.")

    def delete_column(self, event):
        selected_column = self.tree.identify_column(event.x)
        if selected_column:
            col_index = int(selected_column.replace("#", "")) - 1
            col_name = self.tree["columns"][col_index]
            
            if col_name in self.df.columns:
                # Remove the column from DataFrame
                self.df.drop(columns=[col_name], inplace=True)
                self.display_data()

                # Calculate the next column index
                remaining_columns = self.tree["columns"]
                num_columns = len(remaining_columns)
                
                # Set the next column index
                if num_columns > 0:
                    next_col_index = min(col_index, num_columns - 1)
                    self.next_col_index = next_col_index

                    # Select and focus the next column
                    next_col_name = remaining_columns[next_col_index]
                    self.tree.heading(next_col_name, text=next_col_name)
                    self.tree.column(next_col_name, width=100)  # Adjust width as needed

                    # Focus the first row if available
                    if self.tree.get_children():
                        self.tree.selection_set(self.tree.get_children()[0])
                        self.tree.focus(self.tree.get_children()[0])

    def __init__(self, root):
        self.root = root
        self.root.title("CSV Editor")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")  # Light gray background

        # Import button
        self.import_btn = tk.Button(root, text="Import File", command=self.import_file, font=('Arial', 12), bg="#007bff", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.import_btn.pack(pady=10)

        # Treeview for displaying the data
        self.tree = ttk.Treeview(root, style="Custom.Treeview", selectmode='browse')
        self.tree.pack(expand=True, fill=tk.BOTH, side=tk.LEFT)

        # Style configuration for Treeview
        style = ttk.Style()
        style.configure("Custom.Treeview", background="#ffffff", foreground="#000000", rowheight=25, fieldbackground="#ffffff")
        style.configure("Custom.Treeview.Heading", background="#007bff", foreground="#ffffff", font=('Arial', 12, 'bold'))

        # Add vertical scrollbar
        self.vsb = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.vsb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=self.vsb.set)

        # Add horizontal scrollbar
        self.hsb = ttk.Scrollbar(root, orient="horizontal", command=self.tree.xview)
        self.hsb.pack(side='bottom', fill='x')
        self.tree.configure(xscrollcommand=self.hsb.set)

        # Frame for buttons (Edit, Delete, Add Row, Add Column)
        self.button_frame = tk.Frame(root, bg="#f0f0f0")
        self.button_frame.pack(fill=tk.X, pady=5)

        self.edit_btn = tk.Button(self.button_frame, text="Edit", command=self.edit_row, font=('Arial', 12), bg="#28a745", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.edit_btn.pack(side=tk.LEFT, padx=5)

        self.delete_btn = tk.Button(self.button_frame, text="Delete", command=self.delete_row, font=('Arial', 12), bg="#dc3545", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.delete_btn.pack(side=tk.LEFT, padx=5)

        self.add_row_btn = tk.Button(self.button_frame, text="Add Row", command=self.add_row, font=('Arial', 12), bg="#17a2b8", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.add_row_btn.pack(side=tk.LEFT, padx=5)

        self.add_col_btn = tk.Button(self.button_frame, text="Add Column", command=self.add_column, font=('Arial', 12), bg="#ffc107", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.add_col_btn.pack(side=tk.LEFT, padx=5)

        # Separate Save button
        self.save_frame = tk.Frame(root, bg="#f0f0f0")
        self.save_frame.pack(fill=tk.X, pady=5)

        self.save_btn = tk.Button(self.save_frame, text="Save", command=self.save_file, font=('Arial', 12), bg="#28a745", fg="#ffffff", padx=10, pady=5, relief="flat")
        self.save_btn.pack()

        # Bind keyboard events
        self.tree.bind("<Delete>", self.delete_column)

    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if file_path:
            if file_path.endswith(".csv"):
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)
            self.display_data()

    def display_data(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["column"] = list(self.df.columns)
        self.tree["show"] = "headings"

        # Create a Font object
        font = Font()

        # Adjust the column headers
        for col in self.tree["column"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=font.measure(col))

        # Insert rows and adjust the column widths based on content
        for index, row in self.df.iterrows():
            values = list(row)
            self.tree.insert("", "end", values=values)
            for i, value in enumerate(values):
                col_width = font.measure(str(value)) + 10  # Adjust for padding
                if self.tree.column(self.tree["columns"][i], 'width') < col_width:
                    self.tree.column(self.tree["columns"][i], width=col_width)

    def delete_row(self):
        selected_item = self.tree.selection()
        if selected_item:
            for item in selected_item:
                row_index = self.tree.index(item)
                self.df = self.df.drop(self.df.index[row_index])
                self.df.reset_index(drop=True, inplace=True)
                self.tree.delete(item)

    def add_row(self):
        # Logic to add a new row
        pass

    def add_column(self):
        # Logic to add a new column
        pass

    def edit_row(self):
        # Logic to edit a row
        pass

    def save_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if file_path:
            _, ext = os.path.splitext(file_path)
            if ext == ".csv":
                self.df.to_csv(file_path, index=False)
            elif ext == ".xlsx":
                self.df.to_excel(file_path, index=False)
            messagebox.showinfo("File Saved", f"File has been saved as {file_path}")
        else:
            messagebox.showwarning("Save Cancelled", "No file was saved.")

    def delete_column(self, event):
        selected_column = self.tree.identify_column(event.x)
        if selected_column:
            col_index = int(selected_column.replace("#", "")) - 1
            col_name = self.tree["columns"][col_index]
            
            if col_name in self.df.columns:
                # Remove the column from DataFrame
                self.df.drop(columns=[col_name], inplace=True)
                self.display_data()

                # Find the next column index
                remaining_columns = self.tree["columns"]
                next_col_index = min(col_index, len(remaining_columns) - 1)

                # Ensure the Treeview is updated and focused on the next column
                if remaining_columns:
                    self.tree.heading(remaining_columns[next_col_index], text=remaining_columns[next_col_index])
                    self.tree.column(remaining_columns[next_col_index], width=100)  # Adjust width as needed
                    self.tree.selection_set(self.tree.get_children()[0])  # Select the first row if available
                    self.tree.focus(self.tree.get_children()[0])  # Focus the first row if available


if __name__ == "__main__":
    root = tk.Tk()
    app = CSVEditorApp(root)
    root.mainloop()
