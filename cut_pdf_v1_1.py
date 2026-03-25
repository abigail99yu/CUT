import streamlit as st
import fitz  # PyMuPDF
import os
import pandas as pd
from pathlib import Path

# --- 路徑設定 ---
BASE_DIR = Path(__file__).parent.absolute()
GRAPH_DIR = BASE_DIR / "graph"
LOG_FILE = GRAPH_DIR / "processing_checkpoint.txt"
DATA_EXPORT = GRAPH_DIR / "extraction_results.xlsx"

if not os.path.exists(GRAPH_DIR):
    os.makedirs(GRAPH_DIR)

def extract_assets_from_pdf(pdf_path, year):
    """
    處理單份 PDF：將每一頁轉為圖片以確保能辨識向量圖表 (圓餅圖/長條圖)
    """
    doc = fitz.open(pdf_path)
    file_name = Path(pdf_path).stem
    specific_output_dir = GRAPH_DIR / year / file_name
    os.makedirs(specific_output_dir, exist_ok=True)
    
    results = []
    
    for page_index, page in enumerate(doc):
        page_num = page_index + 1
        page_rect = page.rect
        page_area = page_rect.width * page_rect.height
        
        # 1. 將整頁渲染為圖片 (DPI 200 兼顧清晰度與檔案大小)
        # 這是為了確保 CNN 能辨識那些由「路徑」組成的圓餅圖/長條圖
        zoom = 2  # 放大倍率
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        
        img_name = f"{file_name}_p{page_num}_all.png"
        img_save_path = specific_output_dir / img_name
        pix.save(str(img_save_path))
        
        # 2. 統計數據 (全頁圖片面積占比預設為 100%)
        results.append({
            "年份": year,
            "原始PDF檔名": file_name,
            "PDF總頁數": len(doc),
            "圖片所在頁碼": page_num,
            "圖片存檔名稱": img_name,
            "圖片面積占比(%)": 100.0,
            "類型": "FullPage_Rendering"
        })
        
        pix = None  # 釋放記憶體
            
    doc.close()
    return results

# --- Streamlit UI ---
st.set_page_config(page_title="永續報告書圖表裁切系統", layout="wide")
st.title("📊 永續報告書：全頁面圖表自動轉檔工具")
st.markdown("""
* **解決 .jpx 問題**：自動轉為 PNG
* **解決圖表抓不到問題**：採用全頁渲染，確保圓餅圖、長條圖不遺漏
* **斷點續傳**：自動紀錄處理進度
""")

# 自動偵測年份資料夾
exclude = ['graph', '__pycache__', 'venv', '.git']
years = sorted([d for d in os.listdir(BASE_DIR) if os.path.isdir(BASE_DIR / d) and d not in exclude and d.isdigit()])

if not years:
    st.error(f"⚠️ 在目錄 {BASE_DIR} 下找不到年份資料夾 (如 2024)。請確認 PDF 存放位置。")
    st.stop()

selected_year = st.sidebar.selectbox("請選擇要處理的年份", years)

# 讀取進度紀錄
processed_files = set()
if os.path.exists(LOG_FILE):
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        processed_files = set(line.strip() for line in f.readlines())

# 掃描 PDF 檔案
target_path = BASE_DIR / selected_year
all_pdfs = [str(target_path / f) for f in os.listdir(target_path) if f.lower().endswith(".pdf")]
pending_pdfs = [f for f in all_pdfs if f not in processed_files]

# 儀表板
c1, c2, c3 = st.columns(3)
c1.metric("當前年份", selected_year)
c2.metric("待處理檔案", len(pending_pdfs))
c3.metric("已完成檔案", len(processed_files))

if st.button(f"開始執行 {selected_year} 年數據轉檔"):
    if not pending_pdfs:
        st.success("✨ 此年份所有檔案皆已處理完成！")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        all_data_list = []
        
        # 如果 Excel 已存在，先讀取舊資料避免覆蓋
        if os.path.exists(DATA_EXPORT):
            all_data_list = pd.read_excel(DATA_EXPORT).to_dict('records')

        for i, pdf_path in enumerate(pending_pdfs):
            fname = os.path.basename(pdf_path)
            status_text.text(f"正在處理 ({i+1}/{len(pending_pdfs)}): {fname}")
            
            try:
                # 呼叫正確的函數名稱
                file_results = extract_assets_from_pdf(pdf_path, selected_year)
                all_data_list.extend(file_results)
                
                # 寫入斷點紀錄
                with open(LOG_FILE, "a", encoding="utf-8") as f:
                    f.write(f"{pdf_path}\n")
                
                # 每處理完一個檔案就存一次 Excel，確保安全
                pd.DataFrame(all_data_list).to_excel(DATA_EXPORT, index=False)
                
                # 更新進度條
                progress_bar.progress((i + 1) / len(pending_pdfs))
            except Exception as e:
                st.error(f"❌ 檔案 {fname} 處理出錯: {e}")
                continue
                
        st.balloons()
        st.success(f"✅ {selected_year} 年處理完畢！請查看 graph 資料夾。")

# 預覽數據
if os.path.exists(DATA_EXPORT):
    with st.expander("查看 Excel 處理紀錄 (前 100 筆)"):
        st.dataframe(pd.read_excel(DATA_EXPORT).head(100))