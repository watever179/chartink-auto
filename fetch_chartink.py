import pandas as pd
from datetime import datetime
import requests

URL = "https://chartink.com/dashboard/160633"
FILE = "Market_Breadth.xlsx"

# Fetch page
html = requests.get(URL).text

# Read table
tables = pd.read_html(html)
df_live = tables[0]   # Market Breadth table

# Latest row (top one)
latest = df_live.iloc[0]

# Prepare new row
new_row = {
    "Date": datetime.now().strftime("%d %b"),
    "Up 4.5%+ today": latest["Up 4.5%+ today"],
    "Down 4.5%+ today": latest["Down 4.5%+ today"],
    "Up 20%+ in 5d": latest["Up 20%+ in 5d"],
    "Down 20%+ in 5d": latest["Down 20%+ in 5d"],
    "Above 20dma": latest["Above 20dma"],
    "Below 20dma": latest["Below 20dma"],
    "Above 50dma": latest["Above 50dma"],
    "Below 50dma": latest["Below 50dma"],
    "Above 200dma": latest["Above 200dma"],
    "Below 200dma": latest["Below 200dma"]
}

# Load existing excel
old = pd.read_excel(FILE)

# Insert new row on top
final = pd.concat([pd.DataFrame([new_row]), old], ignore_index=True)

# Save
final.to_excel(FILE, index=False)

print("Excel updated successfully")
