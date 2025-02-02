import pandas as pd
import os

def get_order_coverage_and_competitiveness(folder_path):
    # ค้นหาไฟล์ที่มีชื่อขึ้นต้นด้วย 'pricing_project_dataset' และ 'platform_number'
    pricing_files = [f for f in os.listdir(folder_path) if f.startswith('pricing_project_dataset') and f.endswith('.xlsx')]
    platform_files = [f for f in os.listdir(folder_path) if f.startswith('platform_number') and f.endswith('.xlsx')]

    if not pricing_files:
        print("No matching pricing_project_dataset files found.")
        return
    if not platform_files:
        print("No matching platform_number files found.")
        return

    # อ่านไฟล์ pricing_project_dataset
    pricing_data = pd.read_excel(os.path.join(folder_path, pricing_files[0]))
    if 'grass_region' not in pricing_data.columns or 'shopee_order' not in pricing_data.columns or 'shopee_model_competitiveness_status' not in pricing_data.columns:
        print("Missing required columns in pricing_project_dataset file.")
        return

    # คำนวณ SUM ของ shopee_order ตาม grass_region
    pricing_summary = pricing_data.groupby('grass_region')['shopee_order'].sum().reset_index()

    # อ่านไฟล์ platform_number
    platform_data = pd.read_excel(os.path.join(folder_path, platform_files[0]))
    if 'region' not in platform_data.columns or 'platform order' not in platform_data.columns:
        print("Missing required columns in platform_number file.")
        return

    # รวมข้อมูลโดยจับคู่ grass_region กับ region
    merged_data = pricing_summary.merge(platform_data, left_on='grass_region', right_on='region', how='left')

    # แทนค่า NaN ใน platform order เป็น 0 เพื่อป้องกันการหารด้วยศูนย์
    merged_data['platform order'].fillna(0, inplace=True)

    # คำนวณ Order Coverage (by Item)
    merged_data['Order Coverage (by Item)'] = merged_data.apply(
        lambda row: row['shopee_order'] / row['platform order'] if row['platform order'] != 0 else None, axis=1
    )

    # นับจำนวน Shopee > CPT และ Shopee < CPT ตาม grass_region
    competitiveness_counts = pricing_data.groupby('grass_region')['shopee_model_competitiveness_status'].value_counts().unstack(fill_value=0)

    # รวมข้อมูล Shopee > CPT และ Shopee < CPT กับ merged_data
    merged_data = merged_data.merge(competitiveness_counts, left_on='grass_region', right_index=True, how='left')

    # แทนค่า NaN เป็น 0 ถ้าไม่มีค่า
    merged_data[['Shopee > CPT', 'Shopee < CPT']] = merged_data[['Shopee > CPT', 'Shopee < CPT']].fillna(0)

    # คำนวณ # of Item โดยรวม Shopee > CPT และ Shopee < CPT
    merged_data['# of Item'] = merged_data['Shopee > CPT'] + merged_data['Shopee < CPT']

    # คำนวณ Net Competitiveness (by Item)
    merged_data['Net Competitiveness (by Item)'] = merged_data.apply(
        lambda row: (row['Shopee < CPT'] - row['Shopee > CPT']) / row['# of Item'] if row['# of Item'] != 0 else None, axis=1
    )

    # กำหนดลำดับของ grass_region ตามที่ต้องการ
    grass_order = ["SG", "TH", "VN", "ID", "PH", "MY"]
    merged_data['grass_region'] = pd.Categorical(merged_data['grass_region'], categories=grass_order, ordered=True)

    # เรียงข้อมูลตาม grass_region
    merged_data = merged_data.sort_values('grass_region')

    # แสดงผลลัพธ์
    print("\n=== Final Merged Data with Custom Grass Region Sorting ===")
    print(merged_data[['grass_region', 'shopee_order', 'platform order', 'Order Coverage (by Item)', 'Shopee > CPT', 'Shopee < CPT', '# of Item', 'Net Competitiveness (by Item)']])

# ระบุพาธของโฟลเดอร์
folder_path = r"C:\Users\R&J\Desktop\Shoppe\Python_DAT"
get_order_coverage_and_competitiveness(folder_path)
