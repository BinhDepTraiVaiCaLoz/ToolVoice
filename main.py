import tkinter as tk
from tkinter import filedialog, messagebox
import speech_recognition as sr
import openpyxl

# Hàm để chọn file Excel
def select_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, filepath)

# Hàm chuyển giọng nói sang văn bản
def recognize_speech():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            status_label.config(text="Listening...", fg="blue")
            audio = recognizer.listen(source)
            status_label.config(text="Processing...", fg="orange")
            text = recognizer.recognize_google(audio, language="vi-VN")
            text_entry.delete(0, tk.END)
            text_entry.insert(0, text)
            status_label.config(text="Recognition complete.", fg="green")
        except sr.UnknownValueError:
            status_label.config(text="Could not understand the audio.", fg="red")
        except sr.RequestError as e:
            status_label.config(text=f"Error with the service: {e}", fg="red")

# Hàm để ghi văn bản vào ô Excel
def write_to_excel():
    filepath = file_path_entry.get()
    if not filepath:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    text = text_entry.get()
    if not text:
        messagebox.showerror("Error", "No text to write.")
        return

    sheet_name = sheet_entry.get()
    cell = cell_entry.get()

    try:
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            messagebox.showerror("Error", f"Sheet '{sheet_name}' not found in the workbook.")
            return

        sheet = wb[sheet_name]
        sheet[cell] = text
        wb.save(filepath)
        messagebox.showinfo("Success", f"Text written to {cell} in sheet '{sheet_name}'.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write to Excel: {e}")

# Tạo giao diện chính
root = tk.Tk()
root.title("Speech to Text to Excel")

# Phần chọn file Excel
tk.Label(root, text="Select Excel File:").grid(row=0, column=0, sticky="w")
file_path_entry = tk.Entry(root, width=40)
file_path_entry.grid(row=0, column=1)
file_button = tk.Button(root, text="Browse", command=select_file)
file_button.grid(row=0, column=2)

# Phần nhập tên sheet và ô
sheet_label = tk.Label(root, text="Sheet Name:")
sheet_label.grid(row=1, column=0, sticky="w")
sheet_entry = tk.Entry(root)
sheet_entry.grid(row=1, column=1, sticky="w")

cell_label = tk.Label(root, text="Cell (e.g., A1):")
cell_label.grid(row=2, column=0, sticky="w")
cell_entry = tk.Entry(root)
cell_entry.grid(row=2, column=1, sticky="w")

# Phần chuyển giọng nói sang văn bản
record_button = tk.Button(root, text="Start Recording", command=recognize_speech)
record_button.grid(row=3, column=0, columnspan=3, pady=10)

text_label = tk.Label(root, text="Recognized Text:")
text_label.grid(row=4, column=0, sticky="w")
text_entry = tk.Entry(root, width=40)
text_entry.grid(row=4, column=1)

# Nút ghi vào Excel
write_button = tk.Button(root, text="Write to Excel", command=write_to_excel)
write_button.grid(row=5, column=0, columnspan=3, pady=10)

# Status label
status_label = tk.Label(root, text="", fg="green")
status_label.grid(row=6, column=0, columnspan=3)

# Chạy giao diện
root.mainloop()
