import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser

class appxtool:
    def __init__(self, master):
        self.master = master
        master.title("appx-tools (by peasoup)")

        self.appx_list = []
        self.create_widgets()

        self.create_uninstall_widgets()

    def create_uninstall_widgets(self):
        # uninstall appx files button
        self.uninstall_appx_button = tk.Button(self.master, text="uninstall packages and bundles", command=self.show_installed_appx_files)
        self.uninstall_appx_button.grid(row=7, column=0, columnspan=50, padx=5, pady=5, sticky="w")

    def show_installed_appx_files(self):
        installed_appx_files = self.get_installed_appx_files()
        self.show_installed_list("installed packages and bundles", installed_appx_files, self.uninstall_appx)

    def show_installed_list(self, title, items, uninstall_callback):
        installed_list_window = tk.Toplevel(self.master)
        installed_list_window.title(title)

        # appx files list
        installed_list = tk.Listbox(installed_list_window, selectmode=tk.MULTIPLE)
        scrollbar = ttk.Scrollbar(installed_list_window, command=installed_list.yview)
        installed_list.config(yscrollcommand=scrollbar.set)

        for item in items:
            installed_list.insert(tk.END, item)

        installed_list.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # uninstall selected button
        uninstall_button = tk.Button(installed_list_window, text="uninstall selected", command=lambda: uninstall_callback(installed_list))
        uninstall_button.grid(row=1, column=0, columnspan=2, pady=10)

        # Configure rows and columns to expand with the window size
        installed_list_window.rowconfigure(0, weight=1)
        installed_list_window.columnconfigure(0, weight=1)

    def get_installed_appx_files(self):
        # Retrieve the list of installed appx files using PowerShell
        try:
            result = subprocess.run(
                ["powershell", "Get-AppxPackage | Select-Object Name"],
                check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
            )
            appx_packages = [name.strip().replace('\r', '').replace('\n', '').rstrip('. ') for name in result.stdout.strip().split('\n')[2:]]
            print(appx_packages)

            return appx_packages
        except subprocess.CalledProcessError:
            return []


    def uninstall_appx(self, installed_list):
        self.uninstall_selected(installed_list, "appx")


    def uninstall_selected(self, installed_list, type):
        selected_items = [installed_list.get(i) for i in installed_list.curselection()]

        if not selected_items:
            messagebox.showinfo("you need to select something!!!!", f"no packages selected for uninstallation.")
            return

        for item in selected_items:
            try:
                subprocess.run(["powershell", "Get-AppxPackage", "-Name", item, "| Remove-AppxPackage"], check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
                messagebox.showinfo("hooray", f"{type.capitalize()} '{item}' uninstalled successfully")
                installed_list.delete(installed_list.curselection())
            except subprocess.CalledProcessError as e:
                messagebox.showerror("ohhh noooo", f"error uninstalling {type} '{item}': {e.stderr}")


    def create_widgets(self):
        # add appx files button
        self.add_button = tk.Button(self.master, text="add package files", command=self.add_appx_files)
        self.add_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # download appx files button
        self.download_button = tk.Button(self.master, text="download appx files online", command=self.open_download_link)
        self.download_button.grid(row=0, column=0, padx=5, pady=5, sticky="n", columnspan=2)

        # clear all button
        self.clear_button = tk.Button(self.master, text="clear all", command=self.clear_appx_list)
        self.clear_button.grid(row=0, column=1, padx=5, pady=5, sticky="e")

        # separator
        ttk.Separator(self.master, orient='horizontal').grid(row=1, column=0, columnspan=2, sticky='ew', pady=5)

        # appx files list
        self.appx_table = ttk.Treeview(self.master, columns=("name", "path"), show="headings", selectmode="extended")
        self.appx_table.heading("name", text="name")
        self.appx_table.heading("path", text="path")
        self.appx_table.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

        # separator
        ttk.Separator(self.master, orient='horizontal').grid(row=3, column=0, columnspan=2, sticky='ew', pady=5)

        # install selected appx files button
        self.install_button = tk.Button(self.master, text="install listed package files", command=self.install_selected_appx)
        self.install_button.grid(row=4, column=0, padx=5, pady=5, sticky="w")

        # install bundle button
        self.bundle_button = tk.Button(self.master, text="install bundle", command=self.install_bundle)
        self.bundle_button.grid(row=4, column=1, padx=5, pady=5, sticky="e")

        # progress bar
        self.progress_bar = ttk.Progressbar(self.master, orient="horizontal", length=800, mode="determinate")
        self.progress_bar.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        # console text
        self.console_text = tk.Text(self.master, height=5, wrap="word", state="disabled")
        self.console_text.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        # set row and column weights
        for i in range(6):
            self.master.rowconfigure(i, weight=1)
        self.master.columnconfigure(0, weight=1)
        self.master.columnconfigure(1, weight=1)
        self.master.columnconfigure(2, weight=1)

    def add_appx_files(self):
        files = filedialog.askopenfilenames(title="select package files", filetypes=[("appx and msix files", "*.appx;*.msix")])

        for file_path in files:
            # check if the file is already in the list
            if any(appx["path"] == file_path for appx in self.appx_list):
                duplicates = [appx["name"] for appx in self.appx_list if appx["path"] == file_path]
                messagebox.showerror("that already is there!!!", f"these appx packages already exist:\n- {', '.join(duplicates)}")
            else:
                self.appx_list.append({"name": file_path.split("/")[-1], "path": file_path})
                self.refresh_appx_table()

    def open_download_link(self):
        webbrowser.open("https://store.rg-adguard.net/")

    def refresh_appx_table(self):
        # clear existing items
        for item in self.appx_table.get_children():
            self.appx_table.delete(item)
        # insert new items
        for appx in self.appx_list:
            self.appx_table.insert("", "end", values=(appx["name"], appx["path"]))

    def clear_appx_list(self):
        self.appx_list = []
        self.refresh_appx_table()

    def install_selected_appx(self):
        self.clear_console()
        for appx in self.appx_list:
            try:
                subprocess.run(["powershell", "add-appxpackage", appx["path"]], check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
                self.console_output(f"installed {appx['name']} successfully.\n")
            except subprocess.CalledProcessError as e:
                self.console_output(f"error installing {appx['name']}: {e.stderr}\n")
            self.progress_bar.step(100 / len(self.appx_list))
            self.master.update()
        # reset progress bar
        self.progress_bar.step(0)

    def install_bundle(self):
        self.clear_console()
        bundle_path = filedialog.askopenfilename(title="select bundle file", filetypes=[("bundle files", "*.appxbundle;*.msixbundle")])
        if bundle_path:
            try:
                subprocess.run(["powershell", "add-appxpackage", "-path", bundle_path], check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
                self.console_output(f"installed bundle {bundle_path.split('/')[-1]} successfully.\n")
            except subprocess.CalledProcessError as e:
                self.console_output(f"error installing bundle: {e.stderr}\n")
            self.progress_bar.step(100)
            self.master.update()
        # reset progress bar
        self.progress_bar.step(0)

    def console_output(self, message):
        self.console_text.config(state="normal")
        self.console_text.insert("end", message.encode('utf-8').decode('utf-8'), 'tag_stdout')
        self.console_text.insert("end", '\n', 'tag_newline')  # Add a newline after each message
        self.console_text.config(state="disabled")
        self.console_text.see("end")

    def clear_console(self):
        self.console_text.config(state="normal")
        self.console_text.delete(1.0, "end")
        self.console_text.config(state="disabled")

if __name__ == "__main__":
    root = tk.Tk()
    app = appxtool(root)
    root.mainloop()
