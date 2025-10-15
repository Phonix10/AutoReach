import pandas as pd
import yagmail
import time
import sys
import os

# --- STEP 1: Load Excel sheet ---
file_path = r"C:\Users\uditr\Downloads\hr.xlsx" #change the path

try:
    df = pd.read_excel(file_path)
except Exception as e:
    print(f"❌ Error reading Excel file: {e}")
    sys.exit()

# Normalize column names
df.columns = df.columns.str.strip().str.lower()

print("✅ Excel file loaded successfully!")
print("📊 Columns found:", df.columns.tolist())

# Detect correct column names automatically
name_col = None
email_col = None
company_col = None

for col in df.columns:
    if "name" in col and "company" not in col:
        name_col = col
    elif "email" in col:
        email_col = col
    elif "company" in col:
        company_col = col

if not (name_col and email_col and company_col):
    print("❌ Required columns (Name, Email, Company) not found.")
    sys.exit()

# Show sample
print("\n📄 Sample data:")
print(df[[name_col, email_col, company_col]].sample(5))

# --- STEP 2: Ask user for company name ---
company = input("\nEnter company name: ").strip().lower()

# Filter matching companies (case-insensitive, partial match)
filtered_df = df[df[company_col].str.lower().str.contains(company, na=False)]

if filtered_df.empty:
    print(f"⚠️ No contacts found for '{company}'.")
    sys.exit()

print(f"\n✅ Found {len(filtered_df)} contact(s) for '{company}':")
print(filtered_df[[name_col, email_col, company_col]])

# --- STEP 3: Email setup ---
sender_email = "smtppython2@gmail.com" #change the email
app_password = "**** **** **** ****"  #change the app password


try:
    yag = yagmail.SMTP(sender_email, app_password)
except Exception as e:
    print(f"❌ Failed to initialize email client: {e}")
    sys.exit()

# --- STEP 4: Ask subject & use fixed CV path ---
subject = input("\nEnter email subject: ").strip()
cv_path = r"C:\Users\uditr\Downloads\Udit_Ranjan_RESUME.pdf"

# Validate CV file path
if not os.path.exists(cv_path):
    print(f"❌ CV file not found: {cv_path}")
    sys.exit()

# --- STEP 5: Generate email body ---
def generate_email_body(hr_name, company_name):
    return f"""
Dear {hr_name},

I hope you're doing well. My name is Udit Ranjan, and I’m reaching out to express my interest in any suitable roles or internship opportunities at {company_name}.

I’m currently pursuing my B.Tech in Information Technology at NMIMS University and have hands-on experience with backend development using Java (Spring Boot), MySQL, and REST APIs. I’m passionate about building scalable applications and contributing to innovative projects.

I’ve attached my CV for your reference. I would be grateful for an opportunity to discuss how my skills can align with {company_name}’s goals.

Thank you for your time and consideration.

Best regards,  
Udit Ranjan  
📧 uditranjan@example.com  
📞 +91-XXXXXXXXXX  
🌐 https://www.linkedin.com/in/udit-ranjan/
"""

# --- STEP 6: Send personalized emails ---
for _, row in filtered_df.iterrows():
    receiver_name = str(row[name_col]).strip() if not pd.isna(row[name_col]) else "HR"
    receiver_email = str(row[email_col]).strip()

    if not receiver_email or receiver_email.lower() == "nan":
        print(f"⚠️ Missing email for {receiver_name} — skipping")
        continue

    body = generate_email_body(receiver_name, company)

    try:
        yag.send(
            to=receiver_email,
            subject=subject,
            contents=body,
            attachments=cv_path
        )
        print(f"✅ Email sent to {receiver_name} ({receiver_email})")
        time.sleep(2)  # Delay to avoid Gmail rate limits
    except Exception as e:
        print(f"❌ Failed to send email to {receiver_name}: {e}")

print("\n🎯 All emails sent successfully!")


