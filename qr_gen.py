import cv2
import qrcode
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from datetime import datetime

HISTORY_FILE = "scan_history.csv"

# --------------------------------------------------------------
# Save Scanned Data to CSV
# --------------------------------------------------------------
def save_history(data):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not os.path.exists(HISTORY_FILE):
        df = pd.DataFrame(columns=["Timestamp", "Data"])
        df.to_csv(HISTORY_FILE, index=False)

    df = pd.read_csv(HISTORY_FILE)
    df.loc[len(df)] = [timestamp, data]
    df.to_csv(HISTORY_FILE, index=False)


# --------------------------------------------------------------
# Delete History
# --------------------------------------------------------------
def delete_history():
    if os.path.exists(HISTORY_FILE):
        os.remove(HISTORY_FILE)
    messagebox.showinfo("Deleted", "History cleared successfully!")


# --------------------------------------------------------------
# Export Excel
# --------------------------------------------------------------
def export_excel():
    if not os.path.exists(HISTORY_FILE):
        messagebox.showerror("Error", "No history to export.")
        return

    df = pd.read_csv(HISTORY_FILE)
    filename = "QR_History.xlsx"
    df.to_excel(filename, index=False)

    messagebox.showinfo("Exported", f"History exported as {filename}")


# --------------------------------------------------------------
# QR Code Generator
# --------------------------------------------------------------
def generate_qr():
    data = text_input.get()

    if data.strip() == "":
        messagebox.showerror("Error", "Please enter text or URL!")
        return

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")
    img.save("my_qr.png")

    messagebox.showinfo("Success", "QR Code generated as my_qr.png")


# --------------------------------------------------------------
# QR Scan From Camera
# --------------------------------------------------------------
def scan_from_camera():
    detector = cv2.QRCodeDetector()

    for cam_index in [0, 1, 2]:
        cap = cv2.VideoCapture(cam_index)
        if cap.isOpened():
            break

    if not cap.isOpened():
        messagebox.showerror("Error", "Cannot access any camera.")
        return

    cap.set(3, 1280)
    cap.set(4, 720)

    messagebox.showinfo("Info", "Camera started. Show QR.\nPress 'q' to exit.")

    while True:
        ret, frame = cap.read()
        if not ret:
            continue

        data, points, _ = detector.detectAndDecode(frame)

        if data:
            save_history(data)
            messagebox.showinfo("QR Found", f"Decoded:\n{data}")
            break

        cv2.imshow("QR Scanner - Press q to exit", frame)

        if cv2.waitKey(1) == ord("q"):
            break

    cap.release()
    cv2.destroyAllWindows()


# --------------------------------------------------------------
# Scan From Image
# --------------------------------------------------------------
def scan_from_image():
    file_path = filedialog.askopenfilename(
        title="Select QR Code Image",
        filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp")]
    )

    if not file_path:
        return

    img = cv2.imread(file_path)
    if img is None:
        messagebox.showerror("Error", "Cannot open image")
        return

    detector = cv2.QRCodeDetector()
    data, points, _ = detector.detectAndDecode(img)

    if data:
        save_history(data)
        messagebox.showinfo("QR Data Found", f"Decoded:\n{data}")
    else:
        messagebox.showerror("Error", "No QR code detected!")


# --------------------------------------------------------------
# View History (FINAL UPGRADED VERSION)
# --------------------------------------------------------------
def view_history():
    if not os.path.exists(HISTORY_FILE):
        messagebox.showinfo("History", "No scan history found.")
        return

    hist_window = tk.Toplevel(root)
    hist_window.title("Scan History")
    hist_window.geometry("600x450")

    tk.Label(hist_window, text="QR Scan History", font=("Arial", 16, "bold")).pack(pady=10)

    search_frame = tk.Frame(hist_window)
    search_frame.pack(pady=5)

    tk.Label(search_frame, text="Search:", font=("Arial", 12)).pack(side=tk.LEFT)

    search_entry = tk.Entry(search_frame, width=30, font=("Arial", 12))
    search_entry.pack(side=tk.LEFT, padx=10)

    table_frame = tk.Frame(hist_window)
    table_frame.pack(fill="both", expand=True)

    columns = ("ID", "Timestamp", "Data")
    tree = ttk.Treeview(table_frame, columns=columns, show="headings")

    tree.heading("ID", text="ID")
    tree.heading("Timestamp", text="Timestamp")
    tree.heading("Data", text="Scanned Data")

    tree.column("ID", width=50)
    tree.column("Timestamp", width=150)
    tree.column("Data", width=350)

    y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)

    tree.configure(yscroll=y_scroll.set, xscroll=x_scroll.set)

    tree.grid(row=0, column=0, sticky="nsew")
    y_scroll.grid(row=0, column=1, sticky="ns")
    x_scroll.grid(row=1, column=0, sticky="ew")

    table_frame.grid_rowconfigure(0, weight=1)
    table_frame.grid_columnconfigure(0, weight=1)

    df = pd.read_csv(HISTORY_FILE)
    for i, row in df.iterrows():
        tree.insert("", "end", values=(i + 1, row["Timestamp"], row["Data"]))

    def search_history():
        query = search_entry.get().lower()
        for item in tree.get_children():
            tree.delete(item)

        for i, row in df.iterrows():
            if query in row["Timestamp"].lower() or query in row["Data"].lower():
                tree.insert("", "end", values=(i + 1, row["Timestamp"], row["Data"]))

    tk.Button(hist_window, text="Search", bg="skyblue",
              command=search_history).pack(pady=5)

    def copy_item(event):
        selected = tree.focus()
        if selected:
            values = tree.item(selected, "values")
            root.clipboard_clear()
            root.clipboard_append(values[2])
            messagebox.showinfo("Copied", "Scanned data copied to clipboard!")

    tree.bind("<Double-1>", copy_item)

    button_frame = tk.Frame(hist_window)
    button_frame.pack(pady=10)

    tk.Button(button_frame, text="Delete History", bg="red", fg="white",
              width=15, command=delete_history).grid(row=0, column=0, padx=10)

    tk.Button(button_frame, text="Export to Excel", bg="green", fg="white",
              width=15, command=export_excel).grid(row=0, column=1, padx=10)


# --------------------------------------------------------------
# TKINTER UI
# --------------------------------------------------------------
root = tk.Tk()
root.title("QR Code Generator + Scanner with History")
root.geometry("550x420")
root.resizable(False, False)

title = tk.Label(root, text="QR Generator & Scanner App", font=("Arial", 18, "bold"))
title.pack(pady=10)

tk.Label(root, text="Enter text or URL for QR:", font=("Arial", 12)).pack()

text_input = tk.Entry(root, width=45, font=("Arial", 12))
text_input.pack(pady=10)

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

tk.Button(btn_frame, text="Generate QR", width=15, bg="green", fg="white",
          command=generate_qr).grid(row=0, column=0, padx=10)

tk.Button(btn_frame, text="Scan (Camera)", width=15, bg="blue", fg="white",
          command=scan_from_camera).grid(row=0, column=1, padx=10)

tk.Button(btn_frame, text="Scan (Image)", width=15, bg="purple", fg="white",
          command=scan_from_image).grid(row=0, column=2, padx=10)

tk.Button(root, text="View History", width=20, bg="orange", fg="black",
          command=view_history).pack(pady=20)

tk.Label(root, text="Created by You — Advanced Python Project",
         font=("Arial", 10)).pack(pady=10)

root.mainloop()
