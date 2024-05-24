import tkinter as tk
import pyperclip
import os
import win32com.client as win32
import datetime

# Copy the customer's email address to the clipboard and run this script to send an email to the customer.

# To properly set up this script,
    # 1. Make sure the email_body.htm file is located in the same directory as this script.
    # 2. Change the email body content in the email_body.htm file to match your personal email body.
    # 3. Change the signature file path to match your own signature file (usually located in the Signatures folder in AppData).
    # 4. Change the email subject, sign off, and name in the create_outlook_email function to match your own email preferences.
    

# Save the user's name to a text file
def save_name(name): 
    try:
        with open("user_name.txt", "w") as file:
            file.write(name)
    except Exception as e:
        print(f"Error saving name: {e}")

# Getting the user's name from the text file, getting the time of day, and attaching it to the email body (email_body.htm)
def get_email_body(name):
    try:
        with open("email_body.htm", "r", encoding="utf-8") as f:
            email_body = f.read()
    except Exception as e:
        print(f"Error reading email body: {e}")
        return ""

    # Get the current time and determine if it's morning or afternoon
    current_time = datetime.datetime.now().time()
    time_of_day = "morning" if current_time.hour < 12 else "afternoon"
    email_body = f"{name}, \n\n" + email_body.replace("{time}", f"{time_of_day}")
    email_body = email_body.strip()
    return email_body

# Create an Outlook email with the user's name and the email body
def create_outlook_email(name):
    recipient_email = pyperclip.paste().strip()
    signature_file = os.path.expanduser("~\\AppData\\Roaming\\Microsoft\\Signatures\\ppa (hhazzard@ppas.com).htm")
    signature = ""

    if os.path.exists(signature_file):
        try:
            with open(signature_file, "r", encoding="utf-8") as f:
                signature = f.read()
        except Exception as e:
            print(f"Error reading signature file: {e}")
    else:
        print("Signature file not found.")

    # Create an Outlook email 
    outlook = win32.Dispatch('Outlook.Application')
    email = outlook.CreateItem(0)
    email.Subject = "Premier Pools and Spas"
    body_text = get_email_body(name)
    body_text += "<br>Best regards, <br><br>Hunter"
    email.HTMLBody = f"{body_text}<br><br>{signature}"
    email.To = recipient_email

    try:
        email.Display()
    except Exception as e:
        print(f"Error displaying email: {e}")

# Submit the user's name and create the Outlook email
def submit_name():
    name = name_entry.get()
    if name:
        save_name(name)
        create_outlook_email(name)
        root.destroy()

# Submit the user's name when Enter key is pressed
def submit_on_enter(event):
    submit_name()

# Set the location of the working directory (where the email body file is located)
def set_working_directory():
    try:
        os.chdir(r"C:\Users\hhazz\Desktop\auto email")
    except Exception as e:
        print(f"Error setting working directory: {e}")

# Main function, create the GUI and run the program
if __name__ == "__main__":
    set_working_directory()
    
    root = tk.Tk()
    root.title("Name Input")

    name_label = tk.Label(root, text="Please enter your name:")
    name_label.pack()

    name_entry = tk.Entry(root)
    name_entry.pack()

    submit_button = tk.Button(root, text="Submit", command=submit_name)
    submit_button.pack()

    name_entry.bind("<Return>", submit_on_enter)
    name_entry.focus_set()

    root.mainloop()
