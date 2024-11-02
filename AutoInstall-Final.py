import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import subprocess

class ExeManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Program Installer Manager")
        self.root.geometry("650x400")
        self.root.resizable(False, False)
        self.exe_list = self.load_programs()
        self.current_index = 0  # Track the current program index
        self.skipped_indices = set()  # Track skipped programs

        # Main frame with padding
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Label
        ttk.Label(self.main_frame, text="Program Installer Manager", font=("Helvetica", 16)).pack(pady=5)

        # Container frame for arrow buttons and listbox
        container_frame = ttk.Frame(self.main_frame)
        container_frame.pack(fill=tk.BOTH, expand=True)

        # Arrow button frame
        arrow_button_frame = ttk.Frame(container_frame)
        arrow_button_frame.pack(side=tk.LEFT, padx=5)

        self.up_button = ttk.Button(arrow_button_frame, text="↑", command=self.move_up, width=2)
        self.up_button.pack(pady=5)

        self.down_button = ttk.Button(arrow_button_frame, text="↓", command=self.move_down, width=2)
        self.down_button.pack(pady=5)

        # Listbox and scrollbar
        listbox_frame = ttk.Frame(container_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=5, side=tk.LEFT)

        self.listbox = tk.Listbox(listbox_frame, width=80, height=10, selectmode=tk.SINGLE, font=("Helvetica", 10))
        self.listbox.grid(row=0, column=0, sticky='nsew')

        scrollbar = ttk.Scrollbar(listbox_frame, command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.listbox.config(yscrollcommand=scrollbar.set)

        listbox_frame.grid_columnconfigure(0, weight=1)
        listbox_frame.grid_rowconfigure(0, weight=1)

        # Button frame for adding and removing programs
        button_frame_top = ttk.Frame(self.main_frame)
        button_frame_top.pack(anchor='e', pady=5)

        ttk.Button(button_frame_top, text="Add Program", command=self.add_program).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(button_frame_top, text="Remove Selected", command=self.remove_program).grid(row=0, column=1, padx=5, pady=5)

        # Button frame for start and skip actions
        button_frame_bottom = ttk.Frame(self.main_frame)
        button_frame_bottom.pack(anchor='e', pady=5)

        ttk.Button(button_frame_bottom, text="Start", command=self.start_running_programs).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(button_frame_bottom, text="Skip", command=self.skip_program).grid(row=0, column=1, padx=5, pady=5)

        # Load initial list items into the listbox
        for exe in self.exe_list:
            self.listbox.insert(tk.END, exe)

    def add_program(self):
        if len(self.exe_list) >= 10:
            messagebox.showwarning("Limit Reached", "Cannot add more than 10 programs.")
            return
        filepath = filedialog.askopenfilename(filetypes=[("Executable and Installation files", "*.exe;*.msi;*.inf")])
        if filepath:
            normalized_path = os.path.normpath(filepath).lower()  # Normalize for case-insensitive comparison
            if any(os.path.normpath(path).lower() == normalized_path for path in self.exe_list):
                messagebox.showinfo("Duplicate Entry", "This program is already in the list.")
            else:
                self.exe_list.append(filepath)
                self.listbox.insert(tk.END, filepath)
                self.save_programs()

    def remove_program(self):
        selected = self.listbox.curselection()
        if selected:
            index = selected[0]
            del self.exe_list[index]
            self.listbox.delete(index)
            self.save_programs()

    def move_up(self):
        selected = self.listbox.curselection()
        if selected and selected[0] > 0:
            index = selected[0]
            self.exe_list[index], self.exe_list[index - 1] = self.exe_list[index - 1], self.exe_list[index]
            self.update_listbox()
            self.listbox.selection_set(index - 1)
            self.save_programs()

    def move_down(self):
        selected = self.listbox.curselection()
        if selected and selected[0] < len(self.exe_list) - 1:
            index = selected[0]
            self.exe_list[index], self.exe_list[index + 1] = self.exe_list[index + 1], self.exe_list[index]
            self.update_listbox()
            self.listbox.selection_set(index + 1)
            self.save_programs()

    def start_running_programs(self):
        selected = self.listbox.curselection()
        if selected:
            self.current_index = selected[0]
        else:
            self.current_index = 0  # Start from the first item if nothing is selected

        if self.current_index < len(self.exe_list):
            self.run_program_at_index(self.current_index)
        else:
            messagebox.showinfo("No Programs", "The list is empty or selection is out of range.")

    def skip_program(self):
        if self.current_index + 1 < len(self.exe_list):
            self.current_index += 1
            self.skipped_indices.add(self.current_index)
            self.mark_next_program()  # Mark the next program to indicate it will run next
        else:
            messagebox.showinfo("End of List", "There are no more programs to skip.")

    def run_program_at_index(self, index):
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(index)
        self.listbox.activate(index)
        filepath = self.exe_list[index]
        extension = os.path.splitext(filepath)[1].lower()

        if extension == '.exe' or extension == '.msi':
            # Run without showing the command prompt window
            try:
                subprocess.run(filepath, shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to run {filepath}: {e}")
        elif extension == '.inf':
            try:
                subprocess.run(
                    ['rundll32', 'setupapi,InstallHinfSection', 'DefaultInstall', '128', filepath],
                    check=True,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            except subprocess.CalledProcessError:
                messagebox.showerror("Error", f"Failed to install {filepath}")
        else:
            messagebox.showerror("Unsupported File", f"Unsupported file type: {extension}")

        self.mark_next_program()  # Mark the next program

    def mark_next_program(self):
        next_index = self.current_index + 1
        if next_index < len(self.exe_list) and next_index not in self.skipped_indices:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(next_index)
            self.listbox.activate(next_index)

    def update_listbox(self):
        """Refreshes the listbox to reflect changes in the exe_list."""
        self.listbox.delete(0, tk.END)
        for exe in self.exe_list:
            self.listbox.insert(tk.END, exe)

    def save_programs(self):
        with open("exe_list.txt", "w") as file:
            for exe in self.exe_list:
                file.write(exe + "\n")

    def load_programs(self):
        try:
            with open("exe_list.txt", "r") as file:
                return [line.strip() for line in file.readlines()]
        except FileNotFoundError:
            return []

if __name__ == "__main__":
    root = tk.Tk()
    app = ExeManagerApp(root)
    root.mainloop()
