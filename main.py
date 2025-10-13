import pandas as pd
import yagmail

# --- STEP 1: Load Excel sheet ---
file_path = r"C:\Users\uditr\Downloads\_1800+ Talent Acquisition Database .xlsx"  # Use raw string (r"...")
df = pd.read_excel(file_path)


print(df.sample(5))
