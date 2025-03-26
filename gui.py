import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sv_ttk
import pywinstyles
from PIL import Image, ImageTk
from scripts.excel.get_sheet_names import get_sheet_names


class Gui:
    def __init__(self, root):
        # Basic config
        self.root = root
        self.root.title("Report Generator")
        self.root.geometry("900x700+500+200")
        self.root.configure(bg="#1E1E1E")  # Dark background

        # Create main container with gradient background
        self.container = tk.Frame(root, bg="#1E1E1E")
        self.container.pack(fill="both", expand=True)



        # Load and resize logo
        self.load_logo()

        # Set window icon
        if self.logo:
            self.root.iconphoto(True, self.logo)

        # Create views
        self.select_view = SelectView(self.container, self.show_sheet_view, self.logo)
        self.sheet_view = SheetView(self.container, self.show_select_view)

        # Show initial view
        self.show_select_view()

        # Apply theme
        sv_ttk.set_theme("dark")
        self.apply_theme_to_titlebar(root, sv_ttk)
        root.mainloop()

    def load_logo(self):
        try:
            # Load logo image
            logo_path = os.path.join(r"C:\Users\Duban Serrano\Desktop\REPORTES PYTHON\assets\images", "LOGO SRL FINAL.png")
            original_logo = Image.open(logo_path)

            # Resize logo (adjust size as needed)
            logo_width = 180
            logo_height = 150
            resized_logo = original_logo.resize((logo_width, logo_height), Image.LANCZOS)

            # Convert to PhotoImage
            self.logo = ImageTk.PhotoImage(resized_logo)
        except Exception as e:
            print(f"Logo loading error: {e}")
            self.logo = None

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
    def __init__(self, parent, change_view, logo):
        self.frame = tk.Frame(parent, bg="#1E1E1E")

        # Logo display
        if logo:
            self.logo_label = tk.Label(self.frame, image=logo, bg="#1E1E1E")
            self.logo_label.pack(expand=True, pady=(40, 20))


        # Main title with improved styling
        self.main_title = tk.Label(
            self.frame,
            text="Report Generator",
            fg="#FFFFFF",
            bg="#1E1E1E",
            font=("Segoe UI", 36, "bold"),
            pady=20
        )
        self.main_title.pack(expand=True)

        # Subtitle
        self.label_message_charge = tk.Label(
            self.frame,
            text="Upload Excel File",
            fg="#A0A0A0",
            bg="#1E1E1E",
            font=("Segoe UI", 18),
            pady=10
        )
        self.label_message_charge.pack(expand=True)

        # Styled select file button
        self.select_button = tk.Button(
            self.frame,
            text="Select File",
            command=lambda: self.select_file(change_view),
            bg="#2C3E50",  # Dark blue-gray
            fg="white",
            activebackground="#34495E",
            activeforeground="white",
            relief=tk.FLAT,
            font=("Segoe UI", 14),
            padx=20,
            pady=10,
            cursor="hand2"
        )
        self.select_button.pack(expand=True, pady=20)

    def select_file(self, change_view):
        file_route = filedialog.askopenfilename(
            title="Select the file",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_route:
            confirm = messagebox.askyesno("Confirmation",
                                          f"Do you want to open the file: {os.path.basename(file_route)}?")
            if confirm:
                change_view(file_route)


class SheetView:
    def __init__(self, parent, change_view):
        self.frame = tk.Frame(parent, bg="#1E1E1E")

        self.label = tk.Label(
            self.frame,
            text="Sheet Names:",
            fg="#FFFFFF",
            bg="#1E1E1E",
            font=("Segoe UI", 24, "bold"),
            pady=20
        )
        self.label.pack(expand=True)

        self.return_button = tk.Button(
            self.frame,
            text="Back",
            command=change_view,
            bg="#2C3E50",  # Dark blue-gray
            fg="white",
            activebackground="#34495E",
            activeforeground="white",
            relief=tk.FLAT,
            font=("Segoe UI", 14),
            padx=20,
            pady=10,
            cursor="hand2"
        )
        self.return_button.pack(expand=True)

    def update_sheet(self, sheet_names):
        self.label.config(text=f"Sheet Names:\n{', '.join(sheet_names)}")


def main():
    root = tk.Tk()
    app = Gui(root)
    root.mainloop()


if __name__ == "__main__":
    main()