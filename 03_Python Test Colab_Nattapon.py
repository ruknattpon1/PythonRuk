import pandas as pd
import os

def sum_shopee_order_by_grass_region(folder_path):
    # ค้นหาไฟล์ที่มีชื่อขึ้นต้นด้วย 'pricing_project_dataset' และเป็น .xlsx
    excel_files = [f for f in os.listdir(folder_path) if f.startswith('pricing_project_dataset') and f.endswith('.xlsx')]

    if not excel_files:
        print("No matching Excel files found in the folder.")
        return None

    all_data = []

    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        
        try:
            df = pd.read_excel(file_path, engine="openpyxl")  # อ่านไฟล์ Excel
        except Exception as e:
            print(f"Error reading {file}: {e}")
            continue

        # ตรวจสอบว่ามีคอลัมน์ 'grass_region' และ 'shopee_order' หรือไม่
        if 'grass_region' in df.columns and 'shopee_order' in df.columns:
            df_filtered = df[['grass_region', 'shopee_order']].dropna()  # เอาค่าที่ไม่เป็น NaN
            all_data.append(df_filtered)
        else:
            print(f"\nColumns 'grass_region' or 'shopee_order' not found in {file}")

    # รวมข้อมูลทุกไฟล์เข้าเป็น DataFrame เดียว
    if all_data:
        combined_df = pd.concat(all_data)

        # คำนวณผลรวมของ 'shopee_order' แยกตาม 'grass_region'
        result = combined_df.groupby('grass_region')['shopee_order'].sum().reset_index()

        # คำนวณ "# of items" โดยคูณ 0.3 กับ shopee_order
        result["# of items"] = result["shopee_order"] * 0.3

        # กำหนดลำดับของ grass_region ตามที่ต้องการ
        custom_order = ["SG", "TH", "VN", "ID", "PH", "MY"]
        result['grass_region'] = pd.Categorical(result['grass_region'], categories=custom_order, ordered=True)

        # เรียงลำดับตาม custom_order
        result = result.sort_values('grass_region')

        # แสดงผลลัพธ์
        print("\n=== Summed Shopee Orders by Grass Region ===")
        print(result.to_string(index=False))  # แสดงผลแบบอ่านง่าย
    else:
        print("No valid data found.")

# ระบุพาธของโฟลเดอร์
folder_path = r"C:\Users\R&J\Desktop\Shoppe\Python_DAT"
sum_shopee_order_by_grass_region(folder_path)
