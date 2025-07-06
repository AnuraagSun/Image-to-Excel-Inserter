import os
import string
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
import csv


def column_letter_to_index(col_letter):
    """Convert Excel column letter (e.g., 'A', 'AA') to 1-based index."""
    col_letter = col_letter.upper()
    result = 0
    for char in col_letter:
        if char in string.ascii_uppercase:
            result = result * 26 + (ord(char) - ord('A') + 1)
        else:
            return -1  # Invalid character
    return result


class ImageToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🖼️ Image to Excel Inserter")
        self.root.geometry("600x400")

        self.image_folder = tk.StringVar()
        self.output_file = tk.StringVar()
        self.start_row = tk.StringVar(value="1")
        self.start_col = tk.StringVar(value="A")
        self.status_text = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # Image folder
        tk.Label(self.root, text="1. Select Image Folder:").pack(anchor='w', padx=10, pady=(10, 0))
        tk.Button(self.root, text="📁 Browse Folder", command=self.browse_folder).pack(padx=10, anchor='w')
        tk.Label(self.root, textvariable=self.image_folder, fg="blue").pack(anchor='w', padx=20)

        # Start row
        tk.Label(self.root, text="2. Starting Row Number:").pack(anchor='w', padx=10, pady=(10, 0))
        tk.Entry(self.root, textvariable=self.start_row).pack(anchor='w', padx=20)

        # Start column
        tk.Label(self.root, text="3. Starting Column (e.g., A, B, AA):").pack(anchor='w', padx=10, pady=(10, 0))
        tk.Entry(self.root, textvariable=self.start_col).pack(anchor='w', padx=20)

        # Output file
        tk.Label(self.root, text="4. Select Output File (.xlsx or .csv):").pack(anchor='w', padx=10, pady=(10, 0))
        tk.Button(self.root, text="📂 Choose Save Location", command=self.choose_output_file).pack(anchor='w', padx=10)
        tk.Label(self.root, textvariable=self.output_file, fg="green").pack(anchor='w', padx=20)

        # Generate button
        tk.Button(self.root, text="▶️ Generate File", command=self.generate_excel, bg="green", fg="white").pack(pady=15)

        # Status
        tk.Label(self.root, textvariable=self.status_text, fg="blue").pack()

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.image_folder.set(folder)

    def choose_output_file(self):
        file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel File (.xlsx)", "*.xlsx"),
                ("CSV File (.csv)", "*.csv"),
            ]
        )
        if file:
            self.output_file.set(file)

    def validate_inputs(self):
        if not self.image_folder.get() or not os.path.isdir(self.image_folder.get()):
            messagebox.showerror("Error", "Please select a valid image folder.")
            return False

        try:
            row = int(self.start_row.get())
            if row < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Starting row must be a positive integer.")
            return False

        col = self.start_col.get().strip().upper()
        if column_letter_to_index(col) == -1:
            messagebox.showerror("Error", "Invalid Excel column name.")
            return False

        if not self.output_file.get():
            messagebox.showerror("Error", "Please select an output file path.")
            return False

        return True

    def generate_excel(self):
        if not self.validate_inputs():
            return

        folder = self.image_folder.get()
        output_path = self.output_file.get()
        start_row = int(self.start_row.get())
        start_col_letter = self.start_col.get().strip().upper()
        start_col_idx = column_letter_to_index(start_col_letter)

        image_files = sorted(
            [f for f in os.listdir(folder) if f.lower().endswith((".jpg", ".jpeg", ".png"))]
        )

        if not image_files:
            messagebox.showerror("Error", "No image files found in the selected folder.")
            return

        # If saving as .csv
        if output_path.lower().endswith(".csv"):
            confirm = messagebox.askyesno(
                "CSV Notice",
                "CSV files do not support images.\n\nOnly filenames and paths will be saved.\nProceed?"
            )
            if not confirm:
                return

            try:
                with open(output_path, mode='w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["Image Filename", "Image Path"])

                    for idx, img_file in enumerate(image_files, start=1):
                        img_path = os.path.join(folder, img_file)
                        writer.writerow([img_file, img_path])
                        self.status_text.set(f"Writing entry {idx} of {len(image_files)}...")
                        self.root.update_idletasks()

                self.status_text.set("✅ CSV file saved successfully!")
                messagebox.showinfo("Success", "CSV file generated successfully!")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to save CSV file:\n{str(e)}")
            return

        # Else: Generate Excel (.xlsx)
        wb = Workbook()
        ws = wb.active
        current_row = start_row

        for idx, img_file in enumerate(image_files, start=1):
            img_path = os.path.join(folder, img_file)
            self.status_text.set(f"Inserting image {idx} of {len(image_files)}...")
            self.root.update_idletasks()

            try:
                # Validate image with Pillow
                with PILImage.open(img_path) as pil_img:
                    pil_img.verify()

                xl_img = XLImage(img_path)
                cell_coord = f"{start_col_letter}{current_row}"
                ws.add_image(xl_img, cell_coord)
                current_row += 1

            except Exception as e:
                messagebox.showwarning("Warning", f"Failed to insert {img_file}: {str(e)}")

        try:
            wb.save(output_path)
            self.status_text.set("✅ Excel file saved successfully!")
            messagebox.showinfo("Success", "Excel file generated successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel file:\n{str(e)}")


# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    app = ImageToExcelApp(root)
    root.mainloop()
