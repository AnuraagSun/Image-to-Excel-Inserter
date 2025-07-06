# 🖼️ Image to Excel Inserter (Python GUI)

A Python GUI application that lets you bulk insert images into an Excel spreadsheet — **in full original resolution** — starting from a user-defined row and column. Perfect for photographers, designers, engineers, and anyone needing to embed images into Excel without resizing.

---

## 🚀 Features

- 📁 Select a folder containing `.jpg`, `.jpeg`, or `.png` images
- 🔢 Choose a **starting row number** (e.g., 1, 10)
- 🔠 Choose a **starting column** (e.g., A, B, AA, etc.)
- 📂 Select the output `.xlsx` Excel file path
- 🖼️ Inserts images **exactly as they are** — no scaling or resizing
- ↕️ Places one image per row, column stays fixed
- 🔃 Sorted insertion order based on filenames
- ✅ Progress indicator for image processing
- 🛡️ Robust error handling and friendly popups

---

## 💻 GUI Preview

![GUI Screenshot](preview.png) <!-- Optional: Add a real screenshot here -->

---

## 📦 Requirements

- Python 3.7+
- `openpyxl` (for Excel handling)
- `Pillow` (for image processing)

### 🔧 Install Dependencies

```bash
pip install openpyxl pillow
```

---

## 🧠 How It Works

1. **Choose your image folder**  
2. **Set the starting row and column** in the Excel sheet  
3. **Pick where to save the output file**  
4. Click **"Generate Excel"**  
5. Watch as your images are inserted — one per row — into Excel  
6. Open your `.xlsx` file and see your images embedded pixel-perfect!

---

## 📝 Excel Output

- Each image is inserted into the specified column
- The row number increases for each image (e.g., A1, A2, A3...)
- Images maintain **original resolution** (no compression or scaling)
- Excel rows auto-adjust based on image layout (let Excel handle spacing)

---

## ⚠️ Error Handling

The app will display helpful error messages if:

- The image folder is missing or empty
- The starting row is not a valid positive integer
- The starting column is not a valid Excel column label (e.g., "Z", "AA")
- The output path is not selected
- An image fails to load or insert

---

## 📁 Usage Example

1. Add your images to a folder (e.g., `/Users/you/Pictures`)
2. Run the script:
   ```bash
   python image_to_excel_gui.py
   ```
3. Select the folder, row, column, and output file
4. Click **Generate Excel**
5. Done! 🎉

---

## 🔧 Optional Enhancements (Coming Soon)

- [ ] Insert image filename in adjacent column
- [ ] Theme selector (light/dark mode)
- [ ] Preview total images before insertion
- [ ] Drag-and-drop support

---

## 🙌 Contributions

Pull requests are welcome! If you find bugs or have feature ideas, feel free to open an issue.

---

## 📄 License

This project is licensed under the MIT License.

---

## ❤️ Acknowledgments

- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [Pillow (PIL)](https://pillow.readthedocs.io/)
- Built with love using Python and Tkinter 🐍

```
