import pandas as pd
import folium
import geopandas as gpd
from pathlib import Path
import re
import fiona
from shapely.geometry import mapping
import sys

# ==========================================
# 設定：檔案路徑與自動掃描
# ==========================================

# 1. 系統/背景圖層 (這些檔名固定)
COUNTY_SHP = Path("測站與工業區邊界距離/縣市邊界圖層/COUNTY_MOI_1130718.shp")
INDUSTRIAL_SHP = Path("測站與工業區邊界距離/產業園區範圍圖_114110更新/產業園區範圍圖.shp")
SCHOOL_EXCEL = Path("測站與工業區邊界距離/111學年度各級學校名錄（含經緯度）20230825.xlsx")
CENTER_EXCEL = Path("測站與工業區邊界距離/園區名單及座標_114.06.05.xlsx")

# 2. 輸出路徑
OUTPUT_DIR = Path("data")  # 存放生成的詳細頁面
OUTPUT_HTML = "index.html" # 首頁地圖

# 3. 定義「不要」被當作目標工業區掃描的檔案 (避免誤讀系統檔)
EXCLUDE_FILES = {
    SCHOOL_EXCEL.name, 
    CENTER_EXCEL.name, 
    "requirements.txt",
    ".DS_Store"
}

# ==========================================
# 工具函式
# ==========================================

def safe_slug(text: str) -> str:
    """將檔名轉為安全網址格式"""
    text = str(text).strip()
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^\w\u4e00-\u9fff\-_.]", "_", text)
    return text

def read_park_xlsx(path: Path, data_dir: Path):
    """
    嘗試讀取 Excel。
    如果它有「工業區基本資料」這個 Sheet，就視為目標工業區並處理。
    如果沒有，就回傳 None (跳過)。
    """
    try:
        # 檢查 Sheet 是否存在，避免讀取非目標 Excel 報錯
        xl = pd.ExcelFile(path)
        if "工業區基本資料" not in xl.sheet_names:
            return None # 這不是我們要的格式，跳過

        df = pd.read_excel(xl, sheet_name="工業區基本資料")
        # 轉成字典方便取值
        key_col = df.columns[0]
        val_col = df.columns[1]
        m = df.set_index(key_col)[val_col].to_dict()
    except Exception:
        return None # 讀取失敗，跳過

    def get_str(key: str, default: str = "") -> str:
        v = m.get(key, default)
        return "" if v is None else str(v)

    park_name = get_str("工業區名稱", path.stem)
    try:
        lon = float(m.get("工業區中心經度"))
        lat = float(m.get("工業區中心緯度"))
    except:
        print(f"⚠️ {park_name} ({path.name}) 經緯度格式錯誤，跳過。")
        return None

    # 處理量測資料 (生成 HTML)
    raw_page_href = ""
    if "量測資料" in xl.sheet_names:
        try:
            meas_df = pd.read_excel(xl, sheet_name="量測資料")
            if "StartTime" in meas_df.columns:
                meas_df = meas_df.sort_values("StartTime")
            
            data_dir.mkdir(parents=True, exist_ok=True)
            meas_filename = f"{safe_slug(park_name)}_量測資料.html"
            meas_path = data_dir / meas_filename
            
            table_html = meas_df.to_html(index=False, border=0, classes="table")
            # 簡單美化 HTML
            page_html = f"""<!doctype html>
            <html lang="zh-Hant">
            <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>{park_name} - 量測資料</title>
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
            </head>
            <body class="p-3">
            <h3>{park_name}｜量測資料</h3>
            <div class="table-responsive" style="max-height: 90vh;">
                {table_html}
            </div>
            </body></html>"""
            meas_path.write_text(page_html, encoding="utf-8")
            raw_page_href = f"./data/{meas_filename}"
        except Exception as e:
            print(f"⚠️ {park_name} 量測資料生成失敗: {e}")

    return {
        "park_name": park_name,
        "lon": lon,
        "lat": lat,
        "monitoring_period": get_str("監測期間", "（未填）"),
        "data_type": get_str("資料類型", "（未填）"),
        "note": get_str("備註", "（未填）"),
        "raw_page_href": raw_page_href,
    }

def create_popup_html(park):
    """建立互動視窗內容"""
    pid = safe_slug(park["park_name"])
    
    # 按鈕
    btn_html = ""
    if park['raw_page_href']:
        btn_html = f"""<a href="{park['raw_page_href']}" target="_blank" style="color:white;background:#0d6efd;padding:4px 8px;text-decoration:none;border-radius:4px;font-size:12px;">查看原始資料</a>"""

    # 回饋區塊
    feedback_html = f"""
    <div style="margin-top:8px;border-top:1px solid #ccc;padding-top:8px;">
        <textarea id="fb_{pid}" rows="2" style="width:100%;font-size:12px;" placeholder="輸入回饋..."></textarea>
        <button onclick="sendFeedback('{pid}')" style="margin-top:4px;font-size:12px;cursor:pointer;">送出</button>
        <span id="msg_{pid}" style="font-size:11px;color:green;"></span>
    </div>
    """

    return f"""
    <div style="font-family:sans-serif;font-size:13px;min-width:250px;">
        <h5 style="margin:0 0 8px 0;">{park['park_name']}</h5>
        <div><b>監測期間:</b> {park['monitoring_period']}</div>
        <div><b>備註:</b> {park['note']}</div>
        <div style="margin-top:6px;">{btn_html}</div>
        {feedback_html}
        <div id="meta_{pid}" data-park="{park['park_name']}" data-lat="{park['lat']}" data-lon="{park['lon']}" style="display:none;"></div>
    </div>
    """

# ==========================================
# 主程式
# ==========================================
def main():
    print("🚀 開始建立地圖...")
    
    # 1. 建立地圖
    m = folium.Map(location=[23.6, 121], zoom_start=8, tiles="OpenStreetMap")
    
    # 2. 加入背景圖層 (若檔案存在)
    # (縣市邊界)
    if COUNTY_SHP.exists():
        try:
            with fiona.open(COUNTY_SHP) as src:
                # 簡單轉 GeoJSON
                geojson = {"type": "FeatureCollection", "features": [{"type": "Feature", "geometry": mapping(f["geometry"]), "properties": dict(f["properties"])} for f in src]}
            folium.GeoJson(geojson, name="縣市邊界", style_function=lambda x: {"fill": False, "color": "#666", "weight": 1}).add_to(m)
        except Exception as e: print(f"⚠️ 載入縣市邊界失敗: {e}")

    # (產業園區範圍)
    if INDUSTRIAL_SHP.exists():
        try:
            gdf = gpd.read_file(INDUSTRIAL_SHP).to_crs(epsg=4326)
            folium.GeoJson(gdf, name="產業園區範圍", style_function=lambda x: {"color": "orange", "weight": 1, "fillOpacity": 0.2}).add_to(m)
        except: pass

    # (學校)
    fg_school = folium.FeatureGroup(name="學校")
    if SCHOOL_EXCEL.exists():
        try:
            sdf = pd.read_excel(SCHOOL_EXCEL)
            for _, r in sdf.iterrows():
                if pd.notnull(r.get("N")) and pd.notnull(r.get("E")):
                    folium.CircleMarker([r["N"], r["E"]], radius=2, color="red", popup=r.get("學校名稱")).add_to(fg_school)
        except: pass
    fg_school.add_to(m)

    # (全台工業區中心點)
    fg_center = folium.FeatureGroup(name="全台工業區中心點")
    if CENTER_EXCEL.exists():
        try:
            cdf = pd.read_excel(CENTER_EXCEL)
            for _, r in cdf.iterrows():
                lat, lon = r.get("座標(緯度)"), r.get("座標(經度)")
                if pd.notnull(lat) and pd.notnull(lon):
                    folium.CircleMarker([lat, lon], radius=3, color="purple", popup=r.get("園區名稱(比對)")).add_to(fg_center)
        except: pass
    fg_center.add_to(m)

    # 3. 🔥 核心：自動掃描並處理目標工業區
    fg_target = folium.FeatureGroup(name="📌 分析目標 (含回饋)", show=True)
    
    # 抓取當前目錄下所有的 .xlsx
    all_excels = list(Path(".").glob("*.xlsx"))
    print(f"📂 找到 {len(all_excels)} 個 Excel 檔，開始掃描...")

    count = 0
    for p_file in all_excels:
        # 排除系統檔案
        if p_file.name in EXCLUDE_FILES:
            continue
        
        # 嘗試讀取
        data = read_park_xlsx(p_file, OUTPUT_DIR)
        if data:
            print(f"  ✅ 成功載入: {data['park_name']}")
            popup = folium.Popup(create_popup_html(data), max_width=350)
            folium.Marker(
                [data["lat"], data["lon"]],
                popup=popup,
                tooltip=data["park_name"],
                icon=folium.Icon(color="green", icon="info-sign")
            ).add_to(fg_target)
            count += 1
    
    fg_target.add_to(m)
    print(f"🎉 處理完成！共加入 {count} 個目標工業區。")

    # 4. 注入 JS (回饋功能)
    feedback_js = """
    <script>
    const GAS_URL = "https://script.google.com/macros/s/AKfycby5yDZnSrExZyGm3xZzgpFwZbS-877qCAVUsn8BPe9-BuY0ZkzvAC_r04p39GXv9rUs_A/exec";
    async function sendFeedback(pid){
        const meta = document.getElementById("meta_"+pid);
        const txt = document.getElementById("fb_"+pid).value;
        const msg = document.getElementById("msg_"+pid);
        if(!txt) return alert("請輸入內容");
        
        msg.innerText = "傳送中...";
        const form = new URLSearchParams();
        form.append("timestamp", new Date().toISOString());
        form.append("park", meta.dataset.park);
        form.append("lat", meta.dataset.lat);
        form.append("lon", meta.dataset.lon);
        form.append("feedback", txt);
        form.append("page_url", location.href);
        
        try {
            await fetch(GAS_URL, {method:"POST", mode:"no-cors", body:form});
            msg.innerText = "✅ 已送出";
            msg.style.color = "green";
            document.getElementById("fb_"+pid).value = "";
        } catch(e) {
            msg.innerText = "❌ 失敗";
            msg.style.color = "red";
        }
    }
    </script>
    """
    m.get_root().html.add_child(folium.Element(feedback_js))

    folium.LayerControl().add_to(m)
    m.save(OUTPUT_HTML)
    print(f"💾 地圖已存為 {OUTPUT_HTML}")

if __name__ == "__main__":
    main()