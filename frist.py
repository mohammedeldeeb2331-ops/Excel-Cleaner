import pandas as pd 
import os 

if not os.path.exists("user.xlsx"):
    print("Error: Excel file not found!")
    raise SystemExit

try:
    df = pd.read_excel("user.xlsx")
except Exception:
    print("Error: Could not read the Excel file!")
    raise SystemExit

if df.empty :
    print("The Excel file is empty!")
    raise SystemExit

required_colums = ["name" , "email"]

missing_columns = [
    col for col in required_colums
    if col not in df.columns
]
if missing_columns :
    print("Missing required columns!")

    for col in missing_columns:
        print(f"Missing column: '{col}'")

    raise SystemExit

df["name"] = df["name"].str.strip()
df["email"] = df["email"].str.strip()
df["name"] = df["name"].str.title()
df["email"] = df["email"].str.lower()

original_rows = len(df)
df = df.drop_duplicates()
df = df.dropna()


invalid_emails = 0 
invalid_indices = []

print("-------------------")
for  index , email in enumerate(df["email"]):
    if  str(email).count("@") != 1 :
            print(f"Invalid '@' count in email: {email}")
            invalid_emails += 1
            invalid_indices.append(index)
    elif email.split("@")[0].strip() == "" :
        print(f"Missing 'data' before '@' in email: {email}")
        invalid_emails += 1
        invalid_indices.append(index)
    elif "." not in str(email.split("@")[-1]):
        print(f"Missing '.' after '@' in email: {email}")
        invalid_emails += 1 
        invalid_indices.append(index)
    elif email.split("@")[-1].startswith("."):
        print(f"Eror because after '@' in : {email} starts with '.'") 
        invalid_emails += 1 
        invalid_indices.append(index)
    elif ".." in email:
        print(f"Invalid consecutive dots in email: {email}")
        invalid_emails += 1
        invalid_indices.append(index)
    elif email.split("@")[-1].endswith("."):
        print(f"Invalid email: domain ends with '.' in {email}")
        invalid_emails += 1
        invalid_indices.append(index)
df = df.drop(index=invalid_indices)
fixed_rows = len(df)
removed_rows = original_rows - fixed_rows

print("-------------------")
print(f"Invalid emails: {invalid_emails}")
print("-------------------")
print(f"Rows removed: {removed_rows}")
print(f"Rows remaining: {fixed_rows}")
print("-------------------")

try:
    df.to_excel("cleaning_user.xlsx", index=False)
    print("File saved successfully!")
except Exception:
    print("Error: Could not save the Excel file!")
    raise SystemExit    
print("-------------------")

print(df)
print("-------------------")
print("Cleaning completed!")



# py -3.14 frist.py