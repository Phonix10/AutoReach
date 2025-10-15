import pandas as pd
import yagmail
import time

# --- STEP 1: Load Excel sheet ---
file_path = r"C:\Users\uditr\Downloads\_1800+ Talent Acquisition Database .xlsx"  # Use raw string (r"...")
df = pd.read_excel(file_path)


print(df.sample(5))


# --- STEP 2: Ask user for company name ---
company = input("Enter company name: ").strip()

# Filter contacts of that company
filtered_df = df[df["Company Name"].str.lower() == company.lower()]

if filtered_df.empty:
    print(f"No contacts found for {company}")
    exit()


# --- STEP 3: Email setup ---
sender_email = "your_email@gmail.com"        # Replace with your Gmail
app_password = "your_app_password"           # Replace with your Gmail App Password

# Initialize Yagmail client
yag = yagmail.SMTP(sender_email, app_password)




# --- STEP 4: Get email subject and CV file path ---
subject = input("Enter email subject: ")
cv_path = input("Enter path to your CV file (e.g., C:/Users/user/Documents/CV.pdf): ")


def generate_email_body(hr_name, company):
    return f"""
Dear {hr_name},

I hope you're doing well. My name is Udit Ranjan, and I’m reaching out to express my interest in any suitable roles or internship opportunities at {company}.

I’m currently pursuing my B.Tech in Information Technology at NMIMS University and have hands-on experience with backend development using Java (Spring Boot), MySQL, and REST APIs. I’m passionate about building scalable applications and contributing to innovative projects.

I’ve attached my CV for your reference. I would be grateful for an opportunity to discuss how my skills can align with {company}’s goals.

Thank you for your time and consideration.

Best regards,
Udit Ranjan
📧 uditranjan@example.com
📞 +91-XXXXXXXXXX
🌐 https://www.linkedin.com/in/udit-ranjan/
"""

# --- STEP 6: Send personalized emails ---
for _, row in filtered_df.iterrows():
    receiver_name = row["Name"]
    receiver_email = row["Email"]

    if pd.isna(receiver_email):
        print(f"⚠️ Missing email for {receiver_name} ({company}) — skipping")
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

print("🎯 All emails sent successfully!")
