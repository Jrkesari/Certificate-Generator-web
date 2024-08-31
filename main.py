import subprocess
import os
from tkinter import Tk, Label, Button, messagebox

# Replace with the path to your Python executable
python_executable = r"C:\Users\JAYESH R KESARI\.virtualenvs\hotel-lskrodsd\Scripts\python.exe"

def run_certificate_generator():
    try:
        subprocess.run([python_executable, "generate_certificates.py"], check=True)
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to run Certificate Generator: {e}")

def run_email_sender():
    try:
        subprocess.run([python_executable, "email_sender.py"], check=True)
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to run Email Sender: {e}")

# Create the main application window
root = Tk()
root.title("Main Menu")
root.geometry("400x200")
root.configure(bg="#f0f4f8")

# Custom font
font = ("Arial", 12, "bold")

# Create and place buttons
certificate_button = Button(root, text="Run Certificate Generator", command=run_certificate_generator, bg="#4CAF50", fg="white", font=font)
certificate_button.pack(pady=20)

email_button = Button(root, text="Run Email Sender", command=run_email_sender, bg="#4CAF50", fg="white", font=font)
email_button.pack(pady=20)

# Run the application
root.mainloop()
