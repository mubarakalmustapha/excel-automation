from tqdm import tqdm
import pandas as pd
import subprocess
import re

# Load the Excel file
file_path = r"C:\Users\USER\Desktop\excel_file.xlsx"
sheet_name = "Sheet1"
df = pd.read_excel(file_path, sheet_name=sheet_name)

def get_latency(ip):
    try:
        result = subprocess.run(["ping", "-n", "4", ip], capture_output=True, text=True)
        matches = re.findall(r"Average = (\d+)ms", result.stdout)
        if matches:
            return int(matches[-1])
        return "Timeout"
    except Exception as e:
        return f"Error: {e}"

print("Processing...")

# Process each AP IP with a progress bar
for index, row in tqdm(df.iterrows(), total=df.shape[0], desc="Pinging APs"):
    ap_ip = row["AP IP"]
    ip_match = re.search(r"(\d+\.\d+\.\d+\.\d+)", ap_ip)
    if ip_match:
        ip_address = ip_match.group(0)
        latency = get_latency(ip_address)
        df.at[index, "LATENCY(ms)"] = latency

# Save the updated file
df.to_excel(file_path, sheet_name=sheet_name, index=False)
print("Latency values updated successfully!")
