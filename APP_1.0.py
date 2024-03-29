import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from tkinter import messagebox
import csv
from PIL import Image, ImageTk
def read_recipient_emails(excel_file):
    try:
        df = pd.read_excel(excel_file)
        recipients = df.iloc[:, 0].tolist()  # Assuming the email addresses are in the first column
        return recipients
    except Exception as e:
        print("Error reading Excel file:", e)
        return []

def send_email(sender_email, sender_password, recipients, email_title, email_body, attached_files):
    try:
        # Setup the SMTP server
        smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
        smtp_server.starttls()
        smtp_server.login(sender_email, sender_password)

        # Send the email to each recipient
        for recipient in recipients:
            # Create message container
            msg = MIMEMultipart()
            msg['Subject'] = email_title
            msg['From'] = sender_email
            msg['To'] = recipient
            
            # Attach email body
            msg.attach(MIMEText(email_body, 'plain'))

            # Attach files
            for file_path in attached_files:
                if file_path:
                    filename = os.path.basename(file_path)  # Extract filename from the full path
                    with open(file_path, "rb") as attachment:
                        part = MIMEApplication(attachment.read())
                        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                        msg.attach(part)

            # Send the email
            smtp_server.sendmail(sender_email, recipient, msg.as_string())
            
            # Reset the message container for the next recipient
            msg = None

        # Quit SMTP server
        smtp_server.quit()
        status_label.config(text="Email sent successfully!", fg="green")
    except Exception as e:
        status_label.config(text="Error: " + str(e), fg="red")

def send_emails():
    sender_email = sender_email_entry.get()
    sender_password = sender_password_entry.get()
    excel_file = excel_file_entry.get()
    email_title = email_title_entry.get()
    email_body = email_body_text.get("1.0", tk.END).strip()  # Get email body from text widget
    attached_files = file_list_entry.get("1.0",tk.END).strip().split('\n')

    recipients = read_recipient_emails(excel_file)
    if not recipients:
        status_label.config(text="No recipients found!", fg="red")
        return

    send_email(sender_email, sender_password, recipients, email_title, email_body, attached_files)


def load_credentials_gui():
    top = tk.Toplevel(root)
    top.title("Load Credentials")

    credentials_list = tk.Listbox(top, width=80, height=40)
    credentials_list.pack(padx=10, pady=10)

    load_credentials_csv(credentials_list)

    select_button = tk.Button(top, text="Select", command=lambda: select_credentials(credentials_list))
    select_button.pack(pady=5)

def select_credentials(listbox):
    selected_index = listbox.curselection()
    if selected_index:
        selected_index = int(selected_index[0])
        selected_item = listbox.get(selected_index)
        email, *password = selected_item.split(" - ")
        sender_email_entry.delete(0, tk.END)  # Clear existing entry
        sender_email_entry.insert(tk.END, email)
        sender_password_entry.delete(0, tk.END)  # Clear existing entry
        sender_password_entry.insert(tk.END, ' '.join(password))
    else:
        messagebox.showerror("Error", "Please select a credential.")


def load_credentials_csv(listbox):
    listbox.delete(0, tk.END)  # Clear the listbox
    try:
        if not os.path.exists('credentials.csv'):
            with open('credentials.csv', 'w') as file:
                file.write("Email,Password\n")  # Write header row if file doesn't exist

        with open('credentials.csv', 'r') as file:
            reader = csv.reader(file)
            next(reader)  # Skip header row
            for row in reader:
                if len(row) >= 2:  # Ensure row has at least 2 values
                    email, *password = row  # Unpack the row, password can be more than one word
                    listbox.insert(tk.END, f"{email.strip()} - {' '.join(password).strip()}")
                else:
                    print(f"Ignoring row: {row}")  # Optional: Print the problematic row
    except Exception as e:
        messagebox.showerror('Error', 'Error loading credentials: ' + str(e))
    

def save_credentials():
    sender_email = sender_email_entry.get()
    sender_password = sender_password_entry.get()
    credentials_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])

    if not credentials_file:
        return

    with open(credentials_file, "a") as file:
        file.write(f"{sender_email},{sender_password}\n")

# Create GUI window
root = tk.Tk()
root.title("Email Sender")
root.configure(bg="#f0f0f0")
root.resizable(0,0)

# Create labels and entry widgets
sender_email_label = tk.Label(root, text="Sender Email:", bg="#f0f0f0")
sender_email_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
sender_email_entry = tk.Entry(root, width=70)
sender_email_entry.grid(row=0, column=1, padx=5, pady=5)

sender_password_label = tk.Label(root, text="Sender Password:", bg="#f0f0f0")
sender_password_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
sender_password_entry = tk.Entry(root, show="*", width=70)
sender_password_entry.grid(row=1, column=1, padx=5, pady=5)


excel_file_label = tk.Label(root, text="Excel File:", bg="#f0f0f0")
excel_file_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
excel_file_entry = tk.Entry(root, width=70)
excel_file_entry.grid(row=2, column=1, padx=5, pady=5)
excel_file_button = tk.Button(root, text="Browse", command=lambda: excel_file_entry.insert(tk.END, filedialog.askopenfilename()))
excel_file_button.grid(row=2, column=2, padx=5, pady=5)

email_title_label = tk.Label(root, text="Email Title:", bg="#f0f0f0")
email_title_label.grid(row=3, column=0, padx=5, pady=5, sticky="e")
email_title_entry = tk.Entry(root ,width=70)
email_title_entry.grid(row=3, column=1, padx=5, pady=5)

email_body_label = tk.Label(root, text="Email Body:", bg="#f0f0f0")
email_body_label.grid(row=4, column=0, padx=5, pady=5, sticky="e")
email_body_text = tk.Text(root, height=10, width=70)
email_body_text.grid(row=4, column=1, padx=5, pady=5)

file_list_label = tk.Label(root, text="Attached Files (separated by '+'):", bg="#f0f0f0")
file_list_label.grid(row=5, column=0, padx=5, pady=5, sticky="e")

file_list_entry = tk.Text(root, height=5, width=70)  # Use Text widget instead of Entry
file_list_entry.grid(row=5, column=1, padx=5, pady=5)

"""def browse_files():
    file_paths = filedialog.askopenfilenames()
    if file_paths:
        # Extract filenames from full paths
        filenames = [os.path.basename(file_path) for file_path in file_paths]
        
        # Insert new filenames underneath existing content
        file_list_entry.insert(tk.END, "\n".join(filenames) + "\n")"""

#file_list_button = tk.Button(root, text="Browse Files", command=browse_files)
file_list_button = tk.Button(root, text="Browse Files", command=lambda: browse_files(file_list_entry))
file_list_button.grid(row=5, column=2, padx=5, pady=5)

def browse_files(entry_widget):
    file_paths = filedialog.askopenfilenames()
    if file_paths:
        for path in file_paths:
            entry_widget.insert(tk.END, path + "\n")

send_button = tk.Button(root, text="Send Email", command=send_emails, padx=20, pady=5)
send_button.grid(row=7, column=0, columnspan=2, padx=5, pady=5)

status_label = tk.Label(root, text="", fg="black", bg="#f0f0f0")
status_label.grid(row=8, column=1, columnspan=2, padx=5, pady=5)

save_button = tk.Button(root, text="Save Credentials", command=save_credentials, padx=20, pady=5)
save_button.grid(row=7, column=1, columnspan=2, padx=5, pady=5)

load_credentials_button = tk.Button(root, text="Load Credentials", command=load_credentials_gui, padx=20, pady=5)
load_credentials_button.grid(row=7, column=2, columnspan=2, padx=5, pady=5)

image_path = "back.jpg"
image = Image.open(image_path)
photo = ImageTk.PhotoImage(image)

# Create a Label to display the image
image_label = tk.Label(root, image=photo, height=300, width=900)
image_label.grid(row=11, column=0, columnspan=3, padx=0, pady=0,sticky="ew")

root.mainloop()