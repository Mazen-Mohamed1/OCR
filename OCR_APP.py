import easyocr
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, OptionMenu, StringVar
import threading
import datetime
import socket  # For checking internet connection

# Function to check internet connection
def check_internet():
    try:
        # Connect to a well-known server to verify internet connection
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        return True
    except OSError:
        return False

# Function to perform OCR on an image and save text to DataFrame
def ocr_to_excel(image_path, num_columns, language):
    reader = easyocr.Reader([language])
    result = reader.readtext(image_path)

    # Process text into rows based on the specified number of columns
    structured_data = []
    row = []
    for idx, res in enumerate(result):
        text = res[1]
        row.append(text)
        if (idx + 1) % num_columns == 0:
            structured_data.append(row)
            row = []
    if row:
        structured_data.append(row)

    df = pd.DataFrame(structured_data).fillna('')
    return df

# Function to handle single image upload
def upload_single():
    if not check_internet():
        messagebox.showerror("Error", "No internet connection. Please check your connection and try again.")
        return

    try:
        num_columns = int(column_entry.get())
        language = language_option.get()
        if num_columns > 0:
            file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png")])
            if file_path:
                threading.Thread(target=process_single_image, args=(file_path, num_columns, language)).start()
        else:
            messagebox.showerror("Error", "Please enter a valid number of columns.")
    except ValueError:
        messagebox.showerror("Error", "Please enter a valid integer for the number of columns.")

# Function to process a single image in a thread
def process_single_image(file_path, num_columns, language):
    df = ocr_to_excel(file_path, num_columns, language)
    save_to_excel(df)

# Function to save DataFrame to Excel with a default name suggestion
def save_to_excel(df):
    default_name = f"OCR_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    save_path = filedialog.asksaveasfilename(initialfile=default_name, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Success", "Data saved to Excel successfully!")
        root.after(100, root.destroy)  # Waits 100 ms before closing

# Set up the GUI with CustomTkinter for a modern look
ctk.set_appearance_mode("dark")  # Use "dark" mode for a black background
ctk.set_default_color_theme("green")  # Customize to a color theme

root = ctk.CTk()
root.title("OCR to Excel Converter")
root.geometry("500x400")

# Internet connection status
connection_status = "Connected" if check_internet() else "No Connection"
status_label = ctk.CTkLabel(root, text=f"Internet Status: {connection_status}", font=("Arial", 12))
status_label.pack(pady=(5, 0))

# Title Label
title_label = ctk.CTkLabel(root, text="OCR to Excel Converter", font=("Arial", 18, "bold"))
title_label.pack(pady=15)

# Instructions label for column entry
instructions_label = ctk.CTkLabel(root, text="Enter the number of columns in your data:")
instructions_label.pack(pady=(10, 5))

# Column count entry
column_entry = ctk.CTkEntry(root, width=60, placeholder_text="e.g., 3")
column_entry.pack(pady=5)

# Language selection dropdown
language_option = StringVar(value="en")  # Internal code for OCR language
language_display = StringVar(value="English")  # Displayed value for the user

language_label = ctk.CTkLabel(root, text="Select the language of the image:")
language_label.pack(pady=(10, 5))

# Use standard tkinter OptionMenu for compatibility and user language display
language_menu = OptionMenu(root, language_display, "English", "Arabic", command=lambda _: update_language_display())
language_menu.configure(bg="black", fg="white", highlightbackground="black", activebackground="gray")
language_menu.pack(pady=5)

# Update language_option based on language_display
def update_language_display(*args):
    language_option.set("en" if language_display.get() == "English" else "ar")
language_display.trace("w", update_language_display)

# Upload button
btn_single = ctk.CTkButton(root, text="Upload Photo", command=upload_single, width=200, fg_color="darkgreen")
btn_single.pack(pady=(20, 10))

# Run the app
root.mainloop()
