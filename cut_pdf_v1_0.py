import streamlit as st
import fitz  # PyMuPDF
import os
import pandas as pd
from pathlib import Path

# --- 設定區 (修正路徑問題) ---
# 使用相對路徑 "." 代表 app.py 所在的資料夾
BASE_DIR = Path(__file__).parent.absolute() 
GRAPH_DIR = BASE_DIR / "graph"
LOG_FILE = GRAPH_DIR / "processing_checkpoint.txt"
DATA_EXPORT = GRAPH_DIR / "extraction_results.xlsx"

# 確保輸出資料夾存在
if not os.path.exists(GRAPH_DIR):
    os.makedirs(GRAPH_DIR)

# --- 功能函數 ---
def extract_assets_from_pdf(pdf_path, year):
    doc = fitz.open(pdf_path)
    file_name = Path(pdf_path).stem
    
    # 建立 graph/年份/檔名 資料夾
    specific_output_dir = GRAPH_DIR / year / file_name
    os.makedirs(specific_output_dir, exist_ok=True)
    
    results = []
    total_images_in_file = 0
    
    for page_index, page in enumerate(doc):
        page_num = page_index + 1
        page_area = page.rect.width * page.rect.height
        image_info_list = page.get_images(full=True)
        
        for img_index, img_item in enumerate(image_info_list):
            xref = img_item[0]
            try:
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                ext = base_image["ext"]
                
                # 計算面積占比
                img_rects = page.get_image_rects(xref)
                img_area_pct = 0
                if img_rects:
                    r = img_rects[0]
                    img_area_pct = round(((r.width * r.height) / page_area) * 100, 4)

                total_images_in_file += 1
                img_file_name = f"{file_name}_p{page_num}_{total_images_in_file}.{ext}"
                img_save_path = specific_output_dir / img_file_name
                
                with open(img_save_path, "wb") as f:
                    f.write(image_bytes)
                
                results.append({
                    "年份": year,
                    "原始PDF檔名": file_name,
                    "PDF總頁數": len(doc),
                    "圖片所在頁碼": page_num,
                    "圖片存檔名稱": img_file_name,
                    "圖片面積占比(%)": img_area_pct
                })
            except:
                continue # 跳過損毀的圖片物件
            
    doc.close()
    return results

# --- UI 介面 ---
st.title("🚀 永續報告書圖片自動裁切系統")

# 自動偵測年份資料夾
exclude = ['graph', '__pycache__', 'venv', '.git']
years = sorted([d for d in os.listdir(BASE_DIR) if os.path.isdir(BASE_DIR / d) and d not in exclude and d.isdigit()])

if not years:
    st.error(f"在 {BASE_DIR} 找不到年份資料夾 (如 2024)")
    st.stop()

selected_year = st.sidebar.selectbox("選擇年份", years)

# 讀取進度
processed_files = set()
if os.path.exists(LOG_FILE):
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        processed_files = set(line.strip() for line in f.readlines())

# 準備檔案清單
target_path = BASE_DIR / selected_year
all_pdfs = [str(target_path / f) for f in os.listdir(target_path) if f.lower().endswith(".pdf")]
pending_pdfs = [f for f in all_pdfs if f not in processed_files]

st.metric("待處理檔案", len(pending_pdfs))

if st.button("開始執行"):
    if not pending_pdfs:
        st.success("這年份都處理完囉！")
    else:
        progress = st.progress(0)
        status = st.empty()
        all_data = []
        
        # 載入現有 Excel
        if os.path.exists(DATA_EXPORT):
            all_data = pd.read_excel(DATA_EXPORT).to_dict('records')

        for i, pdf_path in enumerate(pending_pdfs):
            fname = os.path.basename(pdf_path)
            status.text(f"處理中: {fname}")
            
            file_results = extract_assets_from_pdf(pdf_path, selected_year)
            all_data.extend(file_results)
            
            # 存紀錄
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"{pdf_path}\n")
            
            pd.DataFrame(all_data).to_excel(DATA_EXPORT, index=False)
            progress.progress((i + 1) / len(pending_pdfs))
            
        st.success("完成！")