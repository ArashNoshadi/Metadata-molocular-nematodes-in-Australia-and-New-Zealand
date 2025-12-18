import pandas as pd
import os

# ==========================================
# 1. تنظیمات ورودی (Input Configuration)
# ==========================================
# مسیر پوشه حاوی فایل‌های اکسل
input_dir = r'G:\Paper\nema-Nanopore-Sequencing\zoology new zealand and australia\data\Suspect'

# لیست فایل‌ها و حد نصاب طول (Length Threshold) برای هرکدام
files_info = [
    ('5.8S.xlsx', 160),
    ('ITS2.xlsx', 400),
    ('ITS1.xlsx', 730),
    ('18S.xlsx', 1700),
    ('28S.xlsx', 3500),
    ('COX1.xlsx', 700),
]

# کشورهای مورد نظر
target_countries = ['Australia', 'New Zealand']

# لیست‌های ذخیره‌سازی داده‌ها
summary_data = []    # برای شیت اول (آمار کلی)
frequency_data = []  # برای شیت دوم (فراوانی مقادیر cp و f-h - همه داده‌ها)
filtered_frequency_data = [] # [NEW] برای شیت سوم (فراوانی مقادیر cp و f-h با شرط طول)

print("--- Start Processing Analysis ---")

# ==========================================
# 2. حلقه پردازش فایل‌ها (Processing Loop)
# ==========================================
for file_name, length_threshold in files_info:
    file_path = os.path.join(input_dir, file_name)
    
    # بررسی وجود فایل
    if not os.path.exists(file_path):
        print(f"[Warning] File not found: {file_name}")
        continue

    print(f"Processing: {file_name} (Threshold: {length_threshold})...")

    try:
        # خواندن فایل اکسل
        df = pd.read_excel(file_path)
        
        # حذف فاصله‌های اضافی از نام ستون‌ها (Standardization)
        df.columns = [c.strip() for c in df.columns]
        
        # بررسی ستون‌های ضروری
        required_cols = ['Geo Loc Name', 'Length (RNA)', 'Family', 'Main_Organism', 'Species']
        if not all(col in df.columns for col in required_cols):
            print(f"  -> Skipping {file_name}: Missing required columns.")
            continue

        # بررسی وجود ستون‌های اختیاری cp و f-h
        has_cp = 'cp' in df.columns
        has_fh = 'f-h' in df.columns

        # تبدیل ستون طول به عدد و حذف مقادیر نامعتبر
        df['Length (RNA)'] = pd.to_numeric(df['Length (RNA)'], errors='coerce')
        df = df.dropna(subset=['Length (RNA)'])

        # پردازش به تفکیک کشور
        for country in target_countries:
            # فیلتر کردن ردیف‌ها بر اساس نام کشور (جستجوی جزئی)
            country_df = df[df['Geo Loc Name'].str.contains(country, case=False, na=False)]
            
            if country_df.empty:
                continue

            # ---------------------------------------------------------
            # بخش A: محاسبات آماری کلی (Main Statistics)
            # ---------------------------------------------------------
            
            # تعداد کل و تعداد معتبر از نظر طول
            total_rows = len(country_df)
            valid_len_count = len(country_df[country_df['Length (RNA)'] > length_threshold])
            
            # تعداد خانه‌های پر شده در ستون‌های cp و f-h
            count_cp = country_df['cp'].count() if has_cp else 0
            count_fh = country_df['f-h'].count() if has_fh else 0

            # تابع داخلی برای محاسبه آمار تاکسونومی (Family, Genus, Species)
            def get_taxonomy_stats(dataframe, column_name, threshold):
                if column_name not in dataframe.columns:
                    return 0, 0, 0, 0
                
                total_items = dataframe[column_name].count()
                
                # گروه‌بندی برای پیدا کردن یونیک‌ها و بررسی طول آن‌ها
                # ماکزیمم طول هر گروه را می‌گیریم
                grouped = dataframe.groupby(column_name)['Length (RNA)'].max()
                
                unique_total = len(grouped)
                # اگر ماکزیمم طول یک گروه > حد نصاب باشد، در دسته >Thr قرار می‌گیرد
                unique_gt = (grouped > threshold).sum()
                unique_le = (grouped <= threshold).sum()
                
                return total_items, unique_total, unique_gt, unique_le

            # اجرای محاسبات
            f_tot, f_uniq, f_gt, f_le = get_taxonomy_stats(country_df, 'Family', length_threshold)
            g_tot, g_uniq, g_gt, g_le = get_taxonomy_stats(country_df, 'Main_Organism', length_threshold)
            s_tot, s_uniq, s_gt, s_le = get_taxonomy_stats(country_df, 'Species', length_threshold)

            # افزودن به لیست خلاصه
            summary_data.append({
                'File Name': file_name,
                'Country': country,
                'Threshold': length_threshold,
                'Total Rows': total_rows,
                'Valid Length (>Thr)': valid_len_count,
                
                'Count Filled (cp)': count_cp,
                'Count Filled (f-h)': count_fh,
                
                'Family (Unique)': f_uniq,
                'Fam Uniq >Thr': f_gt,
                'Fam Uniq <=Thr': f_le,
                
                'Genus (Unique)': g_uniq,
                'Gen Uniq >Thr': g_gt,
                'Gen Uniq <=Thr': g_le,
                
                'Species (Unique)': s_uniq,
                'Spc Uniq >Thr': s_gt,
                'Spc Uniq <=Thr': s_le
            })

            # ---------------------------------------------------------
            # بخش B: استخراج فراوانی مقادیر (Frequencies Breakdown - ALL Data)
            # ---------------------------------------------------------
            
            def extract_frequencies(col_name):
                if col_name in country_df.columns:
                    # شمارش تعداد تکرار هر مقدار (مثلا عدد 5 یا حرف A)
                    counts = country_df[col_name].value_counts(dropna=True)
                    for val, freq in counts.items():
                        frequency_data.append({
                            'File Name': file_name,
                            'Country': country,
                            'Column': col_name,
                            'Value': val,
                            'Frequency': freq
                        })

            if has_cp: extract_frequencies('cp')
            if has_fh: extract_frequencies('f-h')

            # ---------------------------------------------------------
            # [NEW] بخش C: استخراج فراوانی مقادیر فقط برای داده‌های معتبر (Filtered Frequencies)
            # شرط: Length (RNA) >= Threshold
            # ---------------------------------------------------------
            
            # فیلتر کردن دیتافریم بر اساس شرط طول (بزرگتر یا مساوی)
            filtered_df = country_df[country_df['Length (RNA)'] >= length_threshold]

            def extract_filtered_frequencies(col_name):
                if col_name in filtered_df.columns:
                    # شمارش تعداد تکرار هر مقدار در دیتافریم فیلتر شده
                    counts = filtered_df[col_name].value_counts(dropna=True)
                    for val, freq in counts.items():
                        filtered_frequency_data.append({
                            'File Name': file_name,
                            'Country': country,
                            'Column': col_name,
                            'Value': val,
                            'Frequency (Filtered >= Thr)': freq,
                            'Applied Threshold': length_threshold
                        })

            if not filtered_df.empty:
                if has_cp: extract_filtered_frequencies('cp')
                if has_fh: extract_filtered_frequencies('f-h')

    except Exception as e:
        print(f"  -> Error processing {file_name}: {e}")

# ==========================================
# 3. ذخیره‌سازی خروجی (Saving Output)
# ==========================================
print("--- Saving Results to Excel ---")

if summary_data:
    # نام فایل خروجی
    output_filename = 'Final_Complete_Analysissss.xlsx'
    output_path = os.path.join(input_dir, output_filename)
    
    # ساخت دیتافریم‌ها
    df_summary = pd.DataFrame(summary_data)
    df_freq = pd.DataFrame(frequency_data)
    df_freq_filtered = pd.DataFrame(filtered_frequency_data) # [NEW] دیتافریم جدید

    try:
        # ایجاد فایل اکسل با سه شیت
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            
            # شیت 1: خلاصه آمار
            df_summary.to_excel(writer, sheet_name='Main_Statistics', index=False)
            
            # شیت 2: ریز فراوانی‌ها (همه داده‌ها)
            if not df_freq.empty:
                # مرتب‌سازی برای زیبایی
                df_freq = df_freq.sort_values(by=['File Name', 'Country', 'Column', 'Frequency'], ascending=[True, True, True, False])
                df_freq.to_excel(writer, sheet_name='CP_FH_Frequencies', index=False)
            else:
                pd.DataFrame({'Info': ['No data found in cp/f-h columns']}).to_excel(writer, sheet_name='CP_FH_Frequencies', index=False)
            
            # شیت 3 [NEW]: ریز فراوانی‌ها (فیلتر شده با طول)
            if not df_freq_filtered.empty:
                # مرتب‌سازی
                df_freq_filtered = df_freq_filtered.sort_values(by=['File Name', 'Country', 'Column', 'Frequency (Filtered >= Thr)'], ascending=[True, True, True, False])
                df_freq_filtered.to_excel(writer, sheet_name='Filtered_CP_FH_Frequencies', index=False)
            else:
                pd.DataFrame({'Info': ['No data matched length criteria for cp/f-h']}).to_excel(writer, sheet_name='Filtered_CP_FH_Frequencies', index=False)

        print(f"\n[SUCCESS] File successfully created at:\n{output_path}")
        print("-" * 50)
        print("Sheet 1: 'Main_Statistics' (Counts and Taxonomy info)")
        print("Sheet 2: 'CP_FH_Frequencies' (All data counts for cp/f-h)")
        print("Sheet 3: 'Filtered_CP_FH_Frequencies' (Counts for cp/f-h WHERE Length >= Threshold)")
        print("-" * 50)

    except PermissionError:
        print(f"\n[ERROR] Could not save. Please close '{output_filename}' if it is open.")
else:
    print("\n[INFO] No matching data found to save.")