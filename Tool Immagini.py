import tkinter as tk
from tkinter import filedialog
from app import transform_excel


class App:
    def __init__(self, master):
        self.master = master
        master.title("Modifica Excel")
        master.geometry("600x400")

        # Create a frame to hold the listbox and add some padding to it
        self.listbox_frame = tk.Frame(master, pady=10, padx=10)
        self.listbox_frame.pack(fill=tk.BOTH, expand=True)

        # Create the listbox and add some items
        self.listbox = tk.Listbox(self.listbox_frame, selectmode=tk.MULTIPLE)
        self.listbox.pack(fill=tk.BOTH, expand=True)

        # Create a new frame to hold the buttons and add some padding to it
        self.button_frame = tk.Frame(master, pady=10, padx=10)
        self.button_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # Create the "Open" button and bind it to the open_folder method
        self.open_button = tk.Button(self.button_frame, text="Open", command=self.open_folder)
        self.open_button.pack(side=tk.LEFT)

        # Create the "Start" button and bind it to the test method
        self.test_button = tk.Button(self.button_frame, text="Start", command=self.elaborate_files)
        self.test_button.pack(side=tk.RIGHT)

        master.bind("<Delete>", self.delete_selected)

    def open_folder(self):
        filetypes = [("Excel files", "*.xls")]
        filenames = filedialog.askopenfilenames(filetypes=filetypes,initialdir='~/Desktop')
        for filename in filenames:
            self.listbox.insert(tk.END, filename)

    def elaborate_files(self):
        self.master.config(cursor="wait")
        for i in range(self.listbox.size()):
            path = self.listbox.get(i)
            transform_excel(path)
        self.master.config(cursor="")
        self.progress_label.config(text="Elaborazione completata")

    def delete_selected(self, event):
        # Delete the selected items in the listbox
        for idx in reversed(self.listbox.curselection()):
            self.listbox.delete(idx)


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
