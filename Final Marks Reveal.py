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
from email.mime.image import MIMEImage
from bs4 import BeautifulSoup
from PIL import Image
from io import BytesIO

#####################################################################################################################
# Global variables
usernamee = "32115333034" # الرقم الجامعي
passwordd = "Laioth44@"  # كلمة سر البوابة
number_of_previous_rows = 0  # عدد الصفوف السابقة
number_of_max_rows = 0 # عدد الصفوف الاقصى
receiver_email = "laithamro17@gmail.com"  # الإيميل الذي سيتم إرسال الإشعار إليه

# Email setup
sender_email = "laithamroq@gmail.com" # Use your own email
password = "xjcw bxpi opwj wmds" # Use your own app password or OAuth credentials

#####################################################################################################################
# Function to handle the submission of the form
def on_submit():
    if (username_entry.get().strip() == ""
        or entry_password.get().strip() == ""):
        messagebox.showerror("هناك نقص في المدخلات", "الرجاء ملء جميع الحقول الإجبارية")
        valid = 0
        return
    
    global passwordd, usernamee, receiver_email, number_of_previous_rows, number_of_max_rows
    valid = 1
    try:
        global usernamee
        #student_name = student_name_entry.get().strip()
        usernamee = username_entry.get().strip()
        passwordd = entry_password.get().strip()
        receiver_email = entry_email.get().strip()
        number_of_previous_rows = int(entry_number_of_previous_rows.get().strip())
        number_of_max_rows = int(entry_number_of_max_rows.get().strip())
    except Exception as e:
        pass
    usernamee = username_entry.get().strip()
    passwordd = entry_password.get().strip()
    receiver_email = entry_email.get().strip()
    root.destroy()  # Close the window when submission is successful
############################################################################################################

# Create the main window ( GUI Frame )
root = tk.Tk()
root.title("برنامج علامات الفاينل")
root.geometry("400x555")
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

# Label for the number of previous rows input
ttk.Label(root, text=": عدد المواد أو اللابات اللي نزلت علامتهم هذا الفصل", style="TLabel").pack(pady=3)
ttk.Label(root, text=": (إذا لسا ما نزلك ولا مادة حط رقم 0)", style="TLabel").pack(pady=3)
entry_number_of_previous_rows = create_rounded_entry(root, width=40)
entry_number_of_previous_rows.pack(pady=3)
add_placeholder1(entry_number_of_previous_rows, "0")

# Label for the number of previous rows input
ttk.Label(root, text=": عدد المواد أو اللابات اللي منزلهم الفصل هذا", style="TLabel").pack(pady=3)
entry_number_of_max_rows = create_rounded_entry(root, width=40)
entry_number_of_max_rows.pack(pady=3)
add_placeholder1(entry_number_of_max_rows, "10")

# Email entry field
ttk.Label(root, text=": إيميلك (حتى يوصلك صورة بالرموز الجديدة)", style="TLabel").pack(pady=3)
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

############################################################################################################
def take_screenshot_and_email(driver, table_element, email_subject, email_body):
    # Take a screenshot of the full page
    png_data = driver.get_screenshot_as_png()
    img = Image.open(BytesIO(png_data))

    # Get the table's location and size
    location = table_element.location
    size = table_element.size

    # Calculate the bounding box for the table with padding and margin
    left = location['x'] - 10
    top = location['y'] - 10
    right = location['x'] + size['width'] + 10
    bottom = location['y'] + size['height'] + 10

    # Crop the screenshot to the table
    table_screenshot = img.crop((left, top, right, bottom))

    # Save the cropped image to a buffer
    buffer = BytesIO()
    table_screenshot.save(buffer, format="PNG")
    buffer.seek(0)

    # Send the email with the image as an attachment
    send_email_with_attachment(receiver_email, email_subject, email_body, buffer)

def send_email_with_attachment(to_email, subject, body, attachment):
    try:
        # Email message setup
        msg = MIMEMultipart()
        msg['From'] = "برنامج رموز الفاينل"
        msg['To'] = to_email
        msg['Subject'] = subject

        # Add the body
        body = f"""
        <html>
        <head></head>
        <body>
            <div style="font-family: 'Calibri', sans-serif; font-size: 18px; text-align: center; line-height: 1.8;">
                <p><strong>إسم الطالب :{student_name}</strong></p>
                <p><strong>: علاماتك لهذه الفصل الدراسي بالصورة المرفقة
                <br>
                هذا البرنامج تم تصميمه بواسطة الطالب : ليث قاسم مفيد عمرو
                </strong><p>
                
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(body, 'html', 'utf-8'))

        # Attach the image
        img = MIMEImage(attachment.read())
        img.add_header('Content-Disposition', 'attachment', filename='table.png')
        msg.attach(img)

        # Send the email
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, to_email, msg.as_string())

    except Exception as e:
        pass

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
    driver.get('http://appsrv1.fet.edu.jo:7777/reg_new/index.jsp')

    WebDriverWait(driver, 100).until(EC.presence_of_element_located ((by.NAME,"username")))

    driver.find_element(by.NAME,"username").send_keys(usernamee)
    driver.find_element(by.NAME,"password").send_keys(passwordd)
    time.sleep(1)
    driver.find_element(by.NAME,"password").send_keys(Keys.RETURN)
    while True:
        try:
            # refresh every 4 second until the class has id='classes' is loaded without timeout , if loaded break the loop ( waiting until the registration is available )
            WebDriverWait(driver, 4).until(EC.presence_of_element_located ((by.XPATH,'//*[@id="navmenu"]/li[5]/a')))
            break
        except:
            driver.refresh()
    driver.find_element(by.ID,"navmenu")
    driver.find_element(by.LINK_TEXT,"كشف العلامات").click()
    time.sleep(2)
    student_name_element = driver.find_element(by.XPATH, '//*[@id="workspace"]/div/table/tbody/tr[1]/td[2]/label[2]')
    # Store the text of the label into the variable
    student_name = student_name_element.text.strip()  # Clean up extra whitespace

    try:
        while True:
            driver.refresh()
            time.sleep(2)

            try:
                # Wait for the page to load and check for "null"
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((by.XPATH, '//*[@id="workspace"]/div')))
                page_source = driver.page_source

                if "null" in page_source:
                    # Find the table
                    table_element = driver.find_element(by.XPATH, '//*[@id="workspace"]/table[1]/tbody/tr[2]/td/table')
                    rows = table_element.find_elements(by.XPATH, './/tbody/tr')
                    print(len(rows) - 1)  # Exclude header row
                    # Compare row count
                    current_row_count = len(rows) - 1  # Exclude header row
                    if current_row_count == number_of_max_rows:
                        take_screenshot_and_email(driver, table_element, "دفعة رموز، إضغط للتعرف عليها", "الرجاء مراجعة الجدول.")
                        break
                    if current_row_count != number_of_previous_rows:
                        # Update the row count
                        number_of_previous_rows = current_row_count
                        # Take a screenshot and email it
                        take_screenshot_and_email(driver, table_element, "دفعة رموز، إضغط للتعرف عليها", "تم تحديث العلامات. الرجاء مراجعة الجدول.")
                    else:
                        pass

            except Exception as e:
                pass
    finally:
        driver.quit()
