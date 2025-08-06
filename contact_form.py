import tkinter as tk
from tkinter import messagebox, filedialog
import csv
import os
import subprocess
import platform

def open_csv():
    # Open the contacts.csv with the default program (Excel or other)
    if os.path.isfile('contacts.csv'):
        try:
            if platform.system() == 'Windows':
                os.startfile('contacts.csv')
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(['open', 'contacts.csv'])
            else:  # Linux and others
                subprocess.call(['xdg-open', 'contacts.csv'])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")
    else:
        messagebox.showinfo("Info", "No contacts found. The contacts.csv file does not exist.")

def save_contact():
    name = entry_name.get().strip()
    email = entry_email.get().strip()
    phone = entry_phone.get().strip()
    address = entry_address.get("1.0", tk.END).strip()

    if not name or not email or not phone or not address:
        messagebox.showerror("Error", "Please fill all fields.")
        return

    file_exists = os.path.isfile('contacts.csv')

    with open('contacts.csv', 'a', newline='') as csvfile:
        fieldnames = ['Name', 'Email', 'Phone', 'Address']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        if not file_exists:
            writer.writeheader()

        writer.writerow({'Name': name, 'Email': email, 'Phone': phone, 'Address': address})

    messagebox.showinfo("Success", "Contact saved successfully!")

    entry_name.delete(0, tk.END)
    entry_email.delete(0, tk.END)
    entry_phone.delete(0, tk.END)
    entry_address.delete("1.0", tk.END)

def on_exit():
    if messagebox.askokcancel("Quit", "Do you really want to exit?"):
        root.destroy()

# Create main window
root = tk.Tk()
root.title("Contact Form")
root.geometry("450x450")
root.config(bg="#f0f4f8")  # Light background color

# Menu bar
menubar = tk.Menu(root)
filemenu = tk.Menu(menubar, tearoff=0)
filemenu.add_command(label="View Contacts", command=open_csv)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=on_exit)
menubar.add_cascade(label="File", menu=filemenu)
root.config(menu=menubar)

# Styling helpers
label_opts = {'bg': "#f0f4f8", 'font': ('Arial', 12), 'anchor': 'w'}
entry_opts = {'font': ('Arial', 12)}

# Widgets
tk.Label(root, text="Name:", **label_opts).pack(padx=20, pady=(20, 0), fill='x')
entry_name = tk.Entry(root, **entry_opts)
entry_name.pack(padx=20, pady=5, fill='x')

tk.Label(root, text="Email:", **label_opts).pack(padx=20, pady=(10, 0), fill='x')
entry_email = tk.Entry(root, **entry_opts)
entry_email.pack(padx=20, pady=5, fill='x')

tk.Label(root, text="Phone Number:", **label_opts).pack(padx=20, pady=(10, 0), fill='x')
entry_phone = tk.Entry(root, **entry_opts)
entry_phone.pack(padx=20, pady=5, fill='x')

tk.Label(root, text="Address:", **label_opts).pack(padx=20, pady=(10, 0), fill='x')
entry_address = tk.Text(root, font=('Arial', 12), height=5)
entry_address.pack(padx=20, pady=5, fill='both')

save_btn = tk.Button(root, text="Save Contact", bg="#4a90e2", fg="white", font=('Arial', 12, 'bold'), command=save_contact)
save_btn.pack(pady=20, ipadx=10, ipady=5)

# Handle window close
root.protocol("WM_DELETE_WINDOW", on_exit)

root.mainloop()
