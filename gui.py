from tkinter import END, messagebox, Tk, ttk
from tkinter import filedialog, PhotoImage
import pywinstyles, sys
import sv_ttk
import os
from scripts.excel.get_sheet_names import get_sheet_names


class Gui:
    def __init__(self, root):
        # Basic config
        self.root = root
        self.root.title("Report Generator")
        self.root.geometry("800x600+560+240")

 
        self.container = ttk.Frame(root)
        self.container.pack(fill="both", expand=True)


        self.select_view = SelectView(self.container, self.show_sheet_view)
        self.sheet_view = SheetView(self.container, self.show_select_view)

     
        self.show_select_view()

        
        sv_ttk.set_theme("dark")
        self.apply_theme_to_titlebar(root, sv_ttk)
        root.mainloop()

    def show_select_view(self):
        self.sheet_view.frame.pack_forget()
        self.select_view.frame.pack(fill="both", expand=True)

    def show_sheet_view(self, file_path):
        sheet_names = get_sheet_names(file_path)
        if sheet_names:
            self.sheet_view.update_sheet(sheet_names)
            self.select_view.frame.pack_forget()
            self.sheet_view.frame.pack(fill="both", expand=True)

    def apply_theme_to_titlebar(self, root, sv_ttk):
        version = sys.getwindowsversion()
        if version.major == 10 and version.build >= 22000:
            pywinstyles.change_header_color(root, "black" if sv_ttk.get_theme() == 'dark' else 'normal')
        elif version.major == 10:
            pywinstyles.apply_style(root, "dark" if sv_ttk.get_theme() == 'dark' else 'normal')
            root.wm_attributes("-alpha", 0.99)
            root.wm_attributes("-alpha", 1)


class SelectView:
    def __init__(self, parent, change_view):
        self.frame = ttk.Frame(parent)

        self.main_title = ttk.Label(self.frame, text="Report Generator", foreground="white", background="", font=("Arial", 30), padding=20)
        self.label_message_charge = ttk.Label(self.frame, text="Please upload the excel file", foreground="white", background="", font=("Arial", 15))
        self.select_button = ttk.Button(self.frame, text="Select file", style="FileSelectButton.TButton", cursor="hand2", command=lambda: self.select_file(change_view), padding=15)

        self.main_title.pack(expand=True, pady=20)
        self.label_message_charge.pack(expand=True, pady=10)
        self.select_button.pack(expand=True, pady=10)

    def select_file(self, change_view):
        file_route = filedialog.askopenfilename(
            title="Select the file",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_route:
            confirm  = messagebox.askyesno("Confirmación", f"¿Desea abrir el archivo: {os.path.basename(file_route)}?")
            if confirm:
                change_view(file_route)


class SheetView:
    def __init__(self, parent, change_view):
        self.frame = ttk.Frame(parent)

        self.label = ttk.Label(self.frame, text="Nombres de las hojas:", foreground="white", background="", font=("Arial", 20))
        self.label.pack(expand=True, pady=20)
        self.return_button = ttk.Button(self.frame, text="Volver", command=change_view)
        self.return_button.pack(expand=True)

    def update_sheet(self, sheet_names):
        self.label.config(text=f"Nombres de las hojas:\n{', '.join(sheet_names)}")


def main():
    root = Tk()
    app = Gui(root)
    root.mainloop()


if __name__ == "__main__":
    main()