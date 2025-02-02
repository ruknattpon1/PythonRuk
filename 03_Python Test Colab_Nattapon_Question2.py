import pandas as pd
import os

def process_shopee_competitiveness(folder_path, file_name="pricing_project_dataset.xlsx", output_file="processed_shopee_data.xlsx"):
    # กำหนดพาธของไฟล์
    file_path = os.path.join(folder_path, file_name)
    
    # ตรวจสอบว่าไฟล์มีอยู่หรือไม่
    if not os.path.exists(file_path):
        print(f"File '{file_name}' not found in the folder '{folder_path}'.")
        return None

    # อ่านไฟล์ Excel
    df = pd.read_excel(file_path)

    # ตรวจสอบว่าคอลัมน์ที่ต้องการมีอยู่หรือไม่
    required_columns = ["shopee_item_id", "shopee_model_id", "shopee_model_competitiveness_status"]
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        print(f"Missing columns in dataset: {missing_columns}")
        return None

    # สร้าง DataFrame ใหม่เก็บค่า shopee_item_id โดยไม่ซ้ำ
    unique_shopee_items = df[['shopee_item_id']].drop_duplicates().reset_index(drop=True)

    # ฟังก์ชันกำหนดค่าตามเงื่อนไขที่ระบุ
    def assign_competitiveness_status(item_id):
        subset = df[df["shopee_item_id"] == item_id]["shopee_model_competitiveness_status"]

        if "Shopee < CPT" in subset.values:
            return "Shopee < CPT"
        elif "Shopee = CPT" in subset.values:
            return "Shopee = CPT"
        elif "Shopee > CPT" in subset.values:
            return "Shopee > CPT"
        else:
            return "Others"

    # ใช้ฟังก์ชันเพื่อตรวจสอบและกำหนดค่าคอลัมน์ใหม่
    unique_shopee_items["shopee_model_competitiveness_status"] = unique_shopee_items["shopee_item_id"].apply(assign_competitiveness_status)

    # สร้างพาธสำหรับไฟล์ที่ Export ออกมา
    output_path = os.path.join(folder_path, output_file)

    # Export ข้อมูลเป็นไฟล์ Excel
    unique_shopee_items.to_excel(output_path, index=False)

    print(f"✅ Exported processed data to '{output_path}'")

    return unique_shopee_items

# ระบุโฟลเดอร์ที่เก็บไฟล์
folder_path = r"C:\Users\R&J\Desktop\Shoppe\Python_DAT"

# เรียกใช้ฟังก์ชันและ Export ไฟล์
result_data = process_shopee_competitiveness(folder_path)
