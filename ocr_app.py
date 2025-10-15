import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from PIL import Image, ImageEnhance, ImageFilter
import easyocr
import pandas as pd
import difflib
import re
from datetime import datetime
import threading

# Initialize EasyOCR reader
reader = easyocr.Reader(['en'], gpu=True)  # Set gpu=False if no GPU

# Create main window
root = tk.Tk()
root.title("OCR Speed Test Extractor")
root.geometry("500x400")

# Variables to store file paths
image_files = []
output_folder = ""

# Function to select image files
def select_images():
    global image_files
    image_files = filedialog.askopenfilenames(
        title="Select Image Files",
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")]
    )
    image_label.config(text=f"Selected {len(image_files)} images")

def is_similar(target, line, threshold=0.9):
    return difflib.SequenceMatcher(None, target.lower(), line.lower()).ratio() > threshold

# Function to select output folder
def select_output_folder():
    global output_folder
    output_folder = filedialog.askdirectory(title="Select Output Folder")
    output_label.config(text=f"Output Folder: {output_folder}")

def clean_ocr_number(text):
    # Remove stray commas between digits (e.g., '3.,43' → '3.43')
    text = re.sub(r'(?<=\d)[,](?=\d)', '', text)
    
    # Remove any non-digit characters except one dot
    text = re.sub(r'[^\d.]', '', text)
    
    # If multiple dots, keep only the first one
    parts = text.split('.')
    if len(parts) > 2:
        text = parts[0] + '.' + ''.join(parts[1:])
    
    return text

def update_progress(i, total, percent):
    progress["value"] = i
    status_label.config(text=f"Progress: {i}/{total} ({percent}%)")

# Function to process images and save Excel
def process_images():
    if not image_files or not output_folder:
        messagebox.showerror("Error", "Please select images and output folder.")
        return

    results_data = []
    total = len(image_files)
    progress["maximum"] = total

    for i, file_path in enumerate(reversed(image_files), start=1):  # earliest first
        file_name = os.path.basename(file_path)
        try:
            match = re.search(r"_(\d{8})-(\d{6})_", file_path)
            if match:
                date_str = match.group(1)
                time_str = match.group(2)
                date_obj = datetime.strptime(date_str, "%Y%m%d").date()
                time_obj = datetime.strptime(time_str, "%H%M%S").time()

            results = reader.readtext(file_path)
            text = '\n'.join([result[1] for result in results])
            lines = text.splitlines()

            wanted_lines = []
            for j, line in enumerate(lines):
                if any(is_similar(keyword, line) for keyword in ["EAMbps", "EfJMbps", "Upload Mbps", "EfMbps", "ENMbps"]):
                    wanted_lines = lines[j+1:j+3]
                    wanted_lines[0] = clean_ocr_number(wanted_lines[0])
                    wanted_lines[1] = clean_ocr_number(wanted_lines[1])
                    break

            if wanted_lines:
                if not wanted_lines[0].replace('.', '', 1).isdigit() or not wanted_lines[1].replace('.', '', 1).isdigit():
                    wanted_lines = ["Not Found", "Not Found"]
                result = "\n".join(wanted_lines)
            else:
                wanted_lines = ["Not Found", "Not Found"]
                result = "\n".join(wanted_lines)

            results_data.append([file_name, date_obj, time_obj] + wanted_lines)

        except Exception as e:
            print(f"Error processing {file_name}: {e}")
            results_data.append([file_name, date_obj, time_obj, "Error", "Error"])

        # Update progress bar and label
        percent = int((i / total) * 100)
        root.after(0, lambda i=i, percent=percent: update_progress(i, total, percent))


    df = pd.DataFrame(results_data, columns=["File name", "Date", "Time", "UplinkMbps", "DownlinkMbps"])
    excel_path = os.path.join(output_folder, "aggregated_results.xlsx")
    df.to_excel(excel_path, index=False)
    messagebox.showinfo("Done", f"Results saved to:\n{excel_path}")

def reset_all():
    global image_files, output_folder
    image_files = []
    output_folder = ""

    # Reset labels
    image_label.config(text="No images selected")
    image_label2.config(text="Search <speedtest> in your pic folder, drag box to open them.")
    output_label.config(text="No output folder selected")

    # Reset progress bar and status
    progress["value"] = 0
    status_label.config(text="Progress: 0%")

    # Reset Run button
    run_button.config(text="Run OCR and Save Excel", state="normal")

# GUI layout
tk.Button(root, text="Select Images", command=select_images).pack(pady=10)
image_label = tk.Label(root, text="No images selected")
image_label2 = tk.Label(root, text="Search <speedtest> in your pic folder, drag box to open them.")
image_label.pack()
image_label2.pack()

# Progress bar and status label
progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress.pack(pady=10)

status_label = tk.Label(root, text="Progress: 0%")
status_label.pack()

tk.Button(root, text="Select Output Folder", command=select_output_folder).pack(pady=10)
output_label = tk.Label(root, text="No output folder selected")
output_label.pack()

def start_ocr():
    # Change button text immediately when clicked
    run_button.config(text="Processing...", state="disabled")

    def task():
        process_images()  # your OCR function
        # When finished, reset button text
        root.after(0, lambda: run_button.config(text="Run OCR and Save Excel", state="normal"))

    t = threading.Thread(target=task, daemon=True)
    t.start()

# Create the button with a variable name
run_button = tk.Button(root, text="Run OCR and Save Excel", command=start_ocr)
run_button.pack(pady=20)

reset_button = tk.Button(root, text="Reset", command=reset_all)
reset_button.pack(pady=10)

root.mainloop()