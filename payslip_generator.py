import pandas as pd
from fpdf import FPDF
import os
import yagmail
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# Create output directory for payslips
os.makedirs("payslips", exist_ok=True)

# Read Excel data
try:
    df = pd.read_excel("employees.xlsx")
except Exception as e:
    print("Error reading Excel file:", e)
    exit()

# Calculate net salary
df["Net Salary"] = df["Basic Salary"] + df["Allowances"] - df["Deductions"]

# Initialize email client
try:
    yag = yagmail.SMTP(user=EMAIL_USER, password=EMAIL_PASS)
except Exception as e:
    print("Error setting up email client:", e)
    exit()

# Function to generate PDF
def generate_payslip(employee):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.set_title(f"Payslip - {employee['Name']}")
    pdf.cell(200, 10, txt=f"Payslip for {employee['Name']}", ln=True, align="C")
    pdf.ln(10)

    pdf.cell(200, 10, txt=f"Employee ID: {employee['Employee ID']}", ln=True)
    pdf.cell(200, 10, txt=f"Name: {employee['Name']}", ln=True)
    pdf.cell(200, 10, txt=f"Basic Salary: ${employee['Basic Salary']:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Allowances: ${employee['Allowances']:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Deductions: ${employee['Deductions']:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Net Salary: ${employee['Net Salary']:.2f}", ln=True)

    filename = f"payslips/{employee['Employee ID']}.pdf"
    pdf.output(filename)
    return filename

# Function to send email with attachment
def send_email(recipient, pdf_path, employee_name):
    try:
        yag.send(
            to=recipient,
            subject="Your Payslip for This Month",
            contents=f"Dear {employee_name},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nHR Department",
            attachments=pdf_path
        )
        print(f"Email sent to {recipient}")
    except Exception as e:
        print(f"Failed to send email to {recipient}: {e}")

# Process each employee
for index, row in df.iterrows():
    try:
        payslip_file = generate_payslip(row)
        send_email(row["Email"], payslip_file, row["Name"])
    except Exception as e:
        print(f"Error processing employee ID {row['Employee ID']}: {e}")
