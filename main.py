import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser

class appxtool:
    def __init__(self, master):
        self.master = master
        master.title("appx deployment tool")

        self.appx_list = []
        self.create_widgets()

    def create_widgets(self):
        # add appx files button
        self.add_button = tk.Button(self.master, text="add appx files", command=self.add_appx_files)
        self.add_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # download appx files button
        self.download_button = tk.Button(self.master, text="download appx files from link", command=self.open_download_link)
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
        self.install_button = tk.Button(self.master, text="install selected appx files", command=self.install_selected_appx)
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
        files = filedialog.askopenfilenames(title="select appx files", filetypes=[("appx files", "*.appx")])

        for file_path in files:
            # check if the file is already in the list
            if any(appx["path"] == file_path for appx in self.appx_list):
                duplicates = [appx["name"] for appx in self.appx_list if appx["path"] == file_path]
                messagebox.showerror("duplicates found!", f"these appx packages already exist:\n- {', '.join(duplicates)}")
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
                subprocess.run(["powershell", "add-appxpackage", appx["path"]], check=True, capture_output=True)
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
                subprocess.run(["powershell", "add-appxpackage", "-path", bundle_path], check=True, capture_output=True)
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
