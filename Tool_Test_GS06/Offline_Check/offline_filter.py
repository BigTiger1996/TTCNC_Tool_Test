import pandas as pd
from pathlib import Path

# -------- CONFIG --------
INPUT_FOLDER = Path("input")
OUTPUT_FOLDER = Path("output")
OUTPUT_SUFFIX = "_DesDeviceID_Count"
# ------------------------

# Create output folder if not exists
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# Process all Excel files
for excel_file in INPUT_FOLDER.glob("*.xlsx"):
    print(f"📄 Processing: {excel_file.name}")

    # Read Excel
    df = pd.read_excel(excel_file)

    # Count desdeviceid appearances
    result = (
        df["DesDeviceID"]
        .value_counts()
        .reset_index()
    )
    result.columns = ["DesDeviceID", "Count"]

    # Output file name
    output_file = OUTPUT_FOLDER / f"{excel_file.stem}{OUTPUT_SUFFIX}.xlsx"

    # Save result
    result.to_excel(output_file, index=False)

    print(f"✅ Saved: {output_file.name}")

print("\n🎉 All files processed.")
