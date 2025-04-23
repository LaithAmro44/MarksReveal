# Imports
import time
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import tkinter.font as tkFont
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By as by
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.service import Service
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from bs4 import BeautifulSoup

#####################################################################################################################
# Global variables
usernamee = "32115333034" # الرقم الجامعي
passwordd = "Laioth44@"  # كلمة سر البوابة

receiver_email = "laithamro17@gmail.com"  # الإيميل الذي سيتم إرسال الإشعار إليه

# Email setup
sender_email = "laithamroq@gmail.com" # Use your own email
password = "xjcw bxpi opwj wmds" # Use your own app password or OAuth credentials

email_body = ""  # The body of the email to be sent
#####################################################################################################################
"""
# Function to handle the submission of the form
def on_submit():
    if (username_entry.get().strip() == ""
        or entry_password.get().strip() == ""):
        messagebox.showerror("هناك نقص في المدخلات", "الرجاء ملء جميع الحقول الإجبارية")
        valid = 0
        return
    
    global passwordd, usernamee, receiver_email
    valid = 1
    try:
        global usernamee
        #student_name = student_name_entry.get().strip()
        usernamee = username_entry.get().strip()
        passwordd = entry_password.get().strip()
        receiver_email = entry_email.get().strip()
    except Exception as e:
        pass
    usernamee = username_entry.get().strip()
    passwordd = entry_password.get().strip()
    receiver_email = entry_email.get().strip()
    root.destroy()  # Close the window when submission is successful
############################################################################################################

# Create the main window ( GUI Frame )
root = tk.Tk()
root.title("برنامج علامات الميد")
root.geometry("400x475")
root.configure(bg="#EAEAEA")  # Light gray background

# Fonts and styles
font1 = tkFont.Font(family="Calibri", size=13, weight="bold")
custom_font = ("Calibri", 11)  # Define the desired font family and size and bold

style = ttk.Style()

# Styling the buttons
style.configure("TButton", font=font1, padding=10, relief="flat",
                background="#0078D7", foreground="white")  # Button with modern blue color

# Adding rounded corners
style.map("TButton", background=[('active', '#005A9E')])  # Darker blue on hover

# Styling labels and entries
style.configure("TLabel", background="#EAEAEA", font=font1)
style.configure("TEntry", padding=5, relief="flat", borderwidth=2, 
                highlightthickness=0, background="white")

# Helper function to style ttk.Entry widgets (add rounded borders)
def create_rounded_entry(parent, **kwargs):
    entry = ttk.Entry(parent, **kwargs)
    entry.configure(style="Rounded.TEntry")
    return entry
# empty label for spacing
ttk.Label(root, text="", style="TLabel").pack(pady=3)

ttk.Label(root, text="المعلومات محفوظة للمستخدم فقط والخصوصية %100", style="TLabel", anchor="center", justify="center").pack(pady=3)

# empty label for spacing
ttk.Label(root, text="", style="TLabel").pack(pady=3)

ttk.Label(root, text=": الرقم الجامعي", style="TLabel").pack(pady=3)
username_entry = create_rounded_entry(root, width=40, font=custom_font)  # Default is hidden with '*'
username_entry.pack(pady=3)

def toggle_password_visibility():
    if entry_password.cget('show') == '*':
        entry_password.config(show='')  # Show the password
        show_password_button.config(text="إخفاء كلمة السر")  # Change button text
    else:
        entry_password.config(show='*')  # Hide the password
        show_password_button.config(text="إظهار كلمة السر")  # Change button text

# Create a new style for the button
style.configure("Block.TButton", font=font1, padding=5, relief="flat", background="#0078D7", foreground="black")

# Adding hover effect for the button
style.map("Block.TButton", background=[('active', '#005A9E')], foreground=[('active', "black")])

# Password entry field
ttk.Label(root, text=": كلمة سر البوابة", style="TLabel").pack(pady=3)
entry_password = create_rounded_entry(root, width=40, show="*", font=custom_font)  # Default is hidden with '*'
entry_password.pack(pady=3)


style = ttk.Style()

# First, ensure you are extending from an existing style (like "TButton")
style.configure("Block.TButton1.TButton", padding=5, relief="flat", background="#0078D7", foreground="gray")

# Add a hover effect to change background and foreground colors when the button is active
style.map("Block.TButton1.TButton", background=[('active', 'white')], foreground=[('active', "black")])

# Button to toggle password visibility
show_password_button = ttk.Button(root, text="إظهار كلمة السر", command=toggle_password_visibility, style="Block.TButton1.TButton")
show_password_button.pack(pady=3)


def add_placeholder1(entry, placeholder_text):
    entry.insert(0, placeholder_text)  # Insert placeholder text initially
    entry.config(foreground="grey", justify="right")    # Set the placeholder color to grey

    # Remove placeholder on focus
    def on_focus_in(event):
        if entry.get() == placeholder_text:
            entry.delete(0, tk.END)    # Clear the placeholder text
            entry.config(foreground="black")  # Reset to normal text color

    # Add placeholder back if empty
    def on_focus_out(event):
        if entry.get() == "":
            entry.insert(0, placeholder_text)  # Add the placeholder text back
            entry.config(foreground="grey")    # Set the placeholder color to grey

    entry.bind("<FocusIn>", on_focus_in)
    entry.bind("<FocusOut>", on_focus_out)

# Email entry field
ttk.Label(root, text=": إيميلك (حتى يوصلك إشعار بالعلامة الجديدة)", style="TLabel").pack(pady=3)
entry_email = create_rounded_entry(root, width=40)
entry_email.pack(pady=3)
add_placeholder1(entry_email, "example@mail.com")


# Style for the button
style = ttk.Style()
# First, ensure you are extending from an existing style (like "TButton")
style.configure("Block.TButton1.TButton", padding=5, relief="flat", background="#0078D7", foreground="black")
# Add a hover effect to change background and foreground colors when the button is active
style.map("Block.TButton1.TButton", background=[('active', 'white')], foreground=[('active', "black")])

# Use the new style in the button
submit_button = ttk.Button(root, text="بدء البرنامج", command=on_submit, style="Block.TButton1.TButton")
submit_button.pack(pady=2)

ttk.Label(root, text=": تم كتابة البرنامج وتصميمه بواسطة\nم.ليث عمرو", style="TLabel", anchor="center", justify="center").pack(pady=3)

root.mainloop()
"""
############################################################################################################
# Email setup
# Function to send an email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
def send_email(title,body):
    try:
        # Email setup
        sender_email = "laithamroq@gmail.com"
        password = "xjcw bxpi opwj wmds"  # Use your own app password or OAuth credentials
        smtp_server = "smtp.gmail.com"
        smtp_port = 587

        def create_and_send_email(to_email, subj, body):
            # Email message setup
            msg = MIMEMultipart()
            msg['From'] = Header("برنامج علامات الميد", 'utf-8')
            msg['To'] = to_email
            msg['Subject'] = Header(subj, 'utf-8')

            # Create HTML formatted body
            html_body = f"""
            <html>
            <head></head>
            <body>
                <div style="font-family: 'Calibri', sans-serif; font-size: 22px; text-align: center;">
                    {body}
                </div>
            </body>
            </html>
            """
            # Attach the HTML body with the right encoding for Arabic text
            msg.attach(MIMEText(html_body, 'html', 'utf-8'))

            # Send the email
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, password)
                server.sendmail(sender_email, to_email, msg.as_string())

        # Send emails
        if title == "علامة جديدة":
            create_and_send_email(receiver_email, "علامة الميد",body)
    except Exception as e:
        pass
############################################################################################################

############################################################################################################
# Web scraping
if passwordd != "" and usernamee != "":
    try:
        driver=webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()))
    except :
        try:
            driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        except:
            messagebox.showerror("Error", "Please install Microsoft Edge or Google Chrome browser")
    ############################################################################################################
    driver.maximize_window()
    driver.get('http://appserver.fet.edu.jo:7778/reg_new/index.jsp')

    WebDriverWait(driver, 100).until(EC.presence_of_element_located ((by.NAME,"username")))

    driver.find_element(by.NAME,"username").send_keys(usernamee)
    driver.find_element(by.NAME,"password").send_keys(passwordd)
    time.sleep(1)
    driver.find_element(by.NAME,"password").send_keys(Keys.RETURN)
    while True:
        try:
            # refresh every 4 second until the class has id='classes' is loaded without timeout , if loaded break the loop ( waiting until the registration is available )
            WebDriverWait(driver, 4).until(EC.presence_of_element_located ((by.XPATH,'//*[@id="workspace"]/div[5]/center/table'))) 
            break
        except:
            driver.refresh()

# Function to get the HTML of the specified table
def get_table_html():
        try:
            # Locate the table
            table_element = driver.find_element(by.XPATH, '//*[@id="workspace"]/div[5]')
            # Get the HTML of the table
            return table_element.get_attribute('outerHTML')
        except Exception as e:
            print(f"Error fetching table HTML: {e}")
            return None
def parse_table_rows(table_html):
    """Parse table rows and return a list of rows (each row is a list of column values)."""
    soup = BeautifulSoup(table_html, 'html.parser')
    rows = []
    for row in soup.find_all('tr'):  # Iterate over each table row
        columns = [col.get_text(strip=True) for col in row.find_all('td')]  # Extract text from each column
        if columns:  # Skip empty rows
            rows.append(columns)
    return rows
previous_table_html = get_table_html()

previous_rows = parse_table_rows(previous_table_html)
while True:
        time.sleep(3)  # Wait 3 seconds before refreshing
        current_table_html = get_table_html()

        if current_table_html is None:
            continue

        current_rows = parse_table_rows(current_table_html)

        # Compare rows to detect changes
        if current_rows != previous_rows:
            # Find the changed row(s)
            for i, (prev_row, curr_row) in enumerate(zip(previous_rows, current_rows)):
                if prev_row != curr_row:
                    column_2 = curr_row[1] if len(curr_row) > 1 else "N/A"
                    column_4 = curr_row[3] if len(curr_row) > 3 else "N/A"
                    column_5 = curr_row[4] if len(curr_row) > 4 else "N/A"
                    # cut the column_5 to get first 5 characters
                    if column_5 == "/ 20/ 0":
                        column_5 = "لم تصدر بعد"
                    else:
                        column_5 = column_5[:7]
                    send_email("علامة جديدة", f"""
                    علامة مادة جديدة : {column_2}
                    <br>
                    العلامة : {column_4}
                    <br>
                    علامة المشاركة : {column_5}
                    """)
                    # save the email body in a variable
                    email_body = f"""
                    علامة مادة جديدة : {column_2}
                    العلامة : {column_4}
                    علامة المشاركة : {column_5}
                    
                    : تم إرسال هذا الإشعار للإيميل التالي
                    {receiver_email}
                    """
                    break
            break
        driver.refresh()  # Refresh the page


driver.quit()  # Close the browser

# Show message box containing the variable email_body
messagebox.showinfo("تحديث في جدول علامات الميد", email_body)