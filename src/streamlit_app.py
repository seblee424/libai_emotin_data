import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from folium.plugins import HeatMap
import requests
from openai import OpenAI
import json
import ast
from io import BytesIO

# --- 0. 版本兼容性检查与页面配置 ---
try:
    import streamlit.version as st_version
    st_version = st_version.__version__
    if st_version < "1.28.0":
        st.warning(f"检测到 Streamlit 版本过旧（{st_version}），可能导致功能异常，建议升级：pip install --upgrade streamlit")
except:
    pass

st.set_page_config(
    page_title="李白生平GIS与RAG整合系统",
    page_icon="🐉",
    layout="wide"
)

# --- 1. 初始化OpenAI客户端 ---
client = OpenAI(
    api_key="sk-72997944466a4af2bcd52a068895f8cf",
    base_url="https://api.deepseek.com"
)

# --- 2. 全局变量与配置 ---
XLSX_FILENAME = "https://raw.githubusercontent.com/seblee424/libai_emotin_data/main/%E6%9D%8E%E7%99%BD%E4%BA%BA%E7%94%9F%E9%87%8D%E8%A6%81%E8%8A%82%E7%82%B9%E4%B8%8E%E4%BB%A3%E8%A1%A8%E4%BD%9C%E5%9C%B0%E7%90%86%E4%BD%8D%E7%BD%AE.xlsx"
EMOTION_DATA_URL = "https://raw.githubusercontent.com/seblee424/libai_emotin_data/main/%E6%9D%8E%E7%99%BD%E8%AF%97%E6%AD%8C%E6%95%B0%E6%8D%AE%E6%95%B4%E7%90%86%20%E5%B9%B4%E4%BB%BD%2B%E5%9C%B0%E7%82%B9%2B%E7%AE%80%E4%BD%93%2B%E7%BB%8F%E7%BA%AC%E5%BA%A6%2B%E6%83%85%E6%84%9F%EF%BC%88%E7%AE%80%E4%BD%93%2B%E7%B9%81%E4%BD%93%E6%A0%87%E9%A2%98%2B%E7%B9%81%E4%BD%93%E6%AD%A3%E6%96%87%EF%BC%89.xlsx"

location_col = '地点（古称/今称）'
summary_col = '诗作/事件摘要'

LOCATION_COORDS = {
    "碎叶城": {"lat": 42.8447, "lon": 75.1648, "match_keys": ["碎叶城"]},
    "峨眉山": {"lat": 29.5807, "lon": 103.3592, "match_keys": ["峨眉山"]},
    "蜀中": {"lat": 31.7828, "lon": 104.7570, "match_keys": ["蜀中", "江油"]},
    "荆门/南津关": {"lat": 30.5667, "lon": 111.4500, "match_keys": ["荆门", "南津关"]},
    "岳阳楼": {"lat": 29.3879, "lon": 113.1092, "match_keys": ["岳阳楼", "岳阳"]},
    "安陆": {"lat": 31.3653, "lon": 113.7077, "match_keys": ["安陆"]},
    "黄鹤楼": {"lat": 30.5484, "lon": 114.3168, "match_keys": ["黄鹤楼", "武汉"]},
    "金陵（凤凰台）": {"lat": 32.0415, "lon": 118.7781, "match_keys": ["金陵", "凤凰台", "南京"]},
    "庐山": {"lat": 29.5910, "lon": 115.9922, "match_keys": ["庐山", "九江"]},
    "天姥山": {"lat": 29.5000, "lon": 120.8900, "match_keys": ["天姥山"]},
    "金陵/长干里": {"lat": 32.0298, "lon": 118.7900, "match_keys": ["长干里"]},
    "长安": {"lat": 34.2652, "lon": 108.9500, "match_keys": ["长安", "西安"]},
    "长安/宫廷": {"lat": 34.2652, "lon": 108.9500, "match_keys": ["宫廷"]},
    "长安/洛阳": {"lat": 34.6859, "lon": 112.4600, "match_keys": ["洛阳"]},
    "桃花潭": {"lat": 30.4079, "lon": 118.4230, "match_keys": ["桃花潭", "泾县"]},
    "敬亭山": {"lat": 30.9822, "lon": 118.7844, "match_keys": ["敬亭山", "宣城"]},
    "天门山": {"lat": 31.4285, "lon": 118.3970, "match_keys": ["天门山", "芜湖"]},
    "扬州/旅店": {"lat": 32.3934, "lon": 119.4290, "match_keys": ["扬州"]},
    "夜郎": {"lat": 27.6888, "lon": 106.3773, "match_keys": ["夜郎", "桐梓"]},
    "白帝城": {"lat": 31.0450, "lon": 109.5780, "match_keys": ["白帝城", "奉节"]},
    "秋浦": {"lat": 30.6500, "lon": 117.4800, "match_keys": ["秋浦", "池州"]},
    "当涂": {"lat": 31.5453, "lon": 118.4870, "match_keys": ["当涂", "马鞍山"]},
    "蜀道": {"lat": 31.0000, "lon": 107.0000, "match_keys": ["蜀道"]},
    "月下独酌": {"lat": 34.2652, "lon": 108.9500, "match_keys": ["独酌", "月下"]},
    "静夜思": {"lat": 32.3934, "lon": 119.4290, "match_keys": ["静夜思"]},
    "长江沿线": {"lat": 30.5928, "lon": 114.3055, "match_keys": ["长江"]},
    "战城南": {"lat": 35.0000, "lon": 100.0000, "match_keys": ["边塞", "战争"]},
    "送友人": {"lat": 30.5928, "lon": 114.3055, "match_keys": ["送友人"]},
    "将进酒": {"lat": 34.2652, "lon": 108.9500, "match_keys": ["将进酒", "豪饮"]},
    "行路难": {"lat": 34.2652, "lon": 108.9500, "match_keys": ["行路难"]},
}

# --- 3. 数据加载逻辑 (增强防崩溃版) ---

@st.cache_data(ttl=3600, show_spinner="正在从 GitHub 下载主页数据...")
def load_main_data(url):
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        df = pd.read_excel(BytesIO(response.content), sheet_name=0)
        df.columns = df.columns.str.strip()
    except Exception as e:
        st.error(f"❌ 主页数据下载失败: {str(e)}")
        return pd.DataFrame()
    
    df['coords_key'] = '未知'
    df['Latitude'] = 34.0478
    df['Longitude'] = 108.4357
    
    for idx, row in df.iterrows():
        location_str = str(row.get(location_col, '')).strip()
        for key, data in LOCATION_COORDS.items():
            if location_str == key or any(k in location_str for k in data.get('match_keys', [])):
                df.at[idx, 'coords_key'] = key
                df.at[idx, 'Latitude'] = data['lat']
                df.at[idx, 'Longitude'] = data['lon']
                break
    return df

@st.cache_data(ttl=3600, show_spinner="正在从 GitHub 下载情感数据...")
def load_emotion_data_from_github():
    try:
        response = requests.get(EMOTION_DATA_URL, timeout=15)
        response.raise_for_status()
        df = pd.read_excel(BytesIO(response.content))
        # 强制转字符串并去空格，防止隐形字符
        df.columns = df.columns.astype(str).str.strip()
        
        def get_first_matching_col(keywords):
            for col in df.columns:
                if any(k in col.lower() for k in keywords): return col
            return None

        rename_map = {}
        if c := get_first_matching_col(['经', 'lon', 'longitude']): rename_map[c] = 'Longitude'
        if c := get_first_matching_col(['纬', 'lat', 'latitude']): rename_map[c] = 'Latitude'
        if c := get_first_matching_col(['year', '年', 'time']): rename_map[c] = 'Year'
        # 增加对 Title 的匹配关键词
        if c := get_first_matching_col(['诗名', 'title', '题', '标题', 'name', '诗歌']): rename_map[c] = 'Title'
        if c := get_first_matching_col(['地点', 'location', 'place', 'city']): rename_map[c] = 'Location'

        emo_col = get_first_matching_col(['emotion_top3', 'top3']) or get_first_matching_col(['emotion', '情', 'sentiment'])
        if emo_col: rename_map[emo_col] = 'Emotion_Raw'

        df = df.rename(columns=rename_map)
        
        # --- 关键修复：强制检查必要列，防止 KeyError ---
        required_cols = ['Title', 'Location', 'Emotion', 'Year', 'Latitude', 'Longitude']
        for col in required_cols:
            if col not in df.columns:
                # 如果没找到，创建一个默认列，防止后续程序崩溃
                df[col] = '未知' 
        # ---------------------------------------------

        # 情感解析
        def extract_primary_emotion(val):
            try:
                if isinstance(val, str):
                    parsed = ast.literal_eval(val)
                    if isinstance(parsed, list) and len(parsed) > 0: return parsed[0][0]
                elif isinstance(val, list) and len(val) > 0: return val[0][0]
            except: pass
            return str(val).split(' ')[0] if val else "未知"

        if 'Emotion_Raw' in df.columns:
            df['Emotion'] = df['Emotion_Raw'].apply(extract_primary_emotion)
        
        df = df.dropna(subset=['Latitude', 'Longitude'])
        return df

    except Exception as e:
        st.error(f"❌ 情感数据下载失败: {str(e)}")
        return pd.DataFrame()

# --- 4. RAG Chatbot 逻辑 ---

@st.cache_data(ttl=3600)
def get_cbdb_data(name="李白"):
    try:
        url = f"https://cbdb.fas.harvard.edu/cbdbapi/person.php?name={name}&o=json"
        r = requests.get(url, headers={"User-Agent": "Streamlit App"}, timeout=5)
        return r.json() if r.status_code == 200 else None
    except: return None

def generate_poem_analysis(year, location, emotion, title, cbdb_data):
    """专门用于诗歌卡片的生成逻辑"""
    cbdb_text = json.dumps(cbdb_data, ensure_ascii=False)[:1000] if cbdb_data else "无"
    
    system_prompt = (
        "你是一位精通唐代文学与李白生平的专家AI。\n"
        f"参考史料：{cbdb_text}\n"
        "任务：用户将提供李白的一首诗及其背景（年份、地点、情感标签）。\n"
        "请按以下格式输出（使用Markdown）：\n"
        "### 📜 全诗呈现\n"
        "（请默写全诗，若不确定则注明）\n\n"
        "### 🎭 情感深度解析\n"
        "（结合标签分析诗句如何体现该情感）\n\n"
        "### 🌍 时空与历史背景\n"
        f"（简述李白在{year}年于{location}的人生境遇）\n"
    )
    
    user_prompt = f"请分析李白在 {year} 年，于 {location} 创作的《{title}》。该诗被标记为【{emotion}】情感。"
    
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.7
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"AI 分析服务暂时不可用: {str(e)}"

def run_main_chatbot(cbdb_data, prompt):
    if not prompt: return "请输入有效的问题"
    cbdb_text = json.dumps(cbdb_data, ensure_ascii=False)[:3000] if cbdb_data else "无CBDB资料"
    system_prompt = f"你是李白研究专家。史料参考：{cbdb_text}"
    try:
        messages = [{"role": "system", "content": system_prompt}]
        messages.extend([msg for msg in st.session_state.chat_history[-5:] if msg.get("role") in ["user", "assistant"]])
        messages.append({"role": "user", "content": prompt})
        response = client.chat.completions.create(model="deepseek-chat", messages=messages, temperature=0.7)
        answer = response.choices[0].message.content.strip()
        
        highlight_key = None
        if not st.session_state.data_df.empty:
            for key in st.session_state.data_df['coords_key'].unique():
                if key != '未知' and key in answer:
                    highlight_key = key
                    break
        st.session_state.highlight_location_key = highlight_key
        return answer
    except Exception as e:
        return f"Chatbot错误：{str(e)}"

# --- 5. 地图绘制逻辑 ---

def create_main_map(df, highlight_key):
    if df.empty: return folium.Map(location=[34.0, 108.0], zoom_start=4)
    try:
        center_lat = df['Latitude'].mean()
        center_lon = df['Longitude'].mean()
    except: center_lat, center_lon = 34.0, 108.0
    m = folium.Map(location=[center_lat, center_lon], zoom_start=4.5, tiles="cartodbdarkmatter")
    
    points = df[['Latitude', 'Longitude']].dropna().values.tolist()
    if len(points) > 1: folium.PolyLine(points, color="#00AEEF", weight=3, opacity=0.5).add_to(m)

    for idx, row in df.iterrows():
        try:
            if pd.isna(row['Latitude']): continue
            is_highlighted = (row['coords_key'] == highlight_key)
            color = 'orange' if is_highlighted else 'blue'
            icon = 'fire' if is_highlighted else 'user'
            popup_html = f"<b>{row.get(location_col, '未知')}</b><br>{row.get(summary_col, '')}"
            folium.Marker([row['Latitude'], row['Longitude']], popup=folium.Popup(popup_html, max_width=300), icon=folium.Icon(color=color, icon=icon, prefix='fa')).add_to(m)
        except: continue
    return m

def create_emotion_heatmap(df, period_name):
    if df.empty: return folium.Map(location=[34.0, 108.0], zoom_start=4)
    try:
        center_lat = df['Latitude'].mean()
        center_lon = df['Longitude'].mean()
    except: center_lat, center_lon = 34.0, 108.0
    m = folium.Map(location=[center_lat, center_lon], zoom_start=5, tiles="cartodbdarkmatter")
    
    heatmap_gradients = {
        "豪放与激昂": {0.2: 'orange', 0.6: 'red', 1.0: 'darkred'},
        "喜悦与欢快": {0.2: 'yellow', 0.6: 'orange', 1.0: '#d35400'},
        "哀怨与悲伤": {0.2: 'cyan', 0.6: 'blue', 1.0: 'navy'},
        "忧愁与苦闷": {0.2: 'lightblue', 0.6: 'royalblue', 1.0: '#1a5276'},
        "孤独与寂寞": {0.2: 'plum', 0.6: 'purple', 1.0: '#4a235a'},
        "思乡与怀古": {0.2: '#d7bde2', 0.6: '#8e44ad', 1.0: '#5b2c6f'},
        "友情与知己": {0.2: 'lightgreen', 0.6: 'green', 1.0: 'darkgreen'},
        "闲适与隐逸": {0.2: '#a3e4d7', 0.6: '#16a085', 1.0: '#0e6251'},
        "未知": {0.4: 'gray', 0.8: 'white', 1.0: 'white'}
    }
    marker_colors = {
        "豪放与激昂": "#e74c3c", "喜悦与欢快": "#e67e22", "哀怨与悲伤": "#3498db",
        "忧愁与苦闷": "#2980b9", "孤独与寂寞": "#9b59b6", "思乡与怀古": "#8e44ad",
        "友情与知己": "#2ecc71", "闲适与隐逸": "#1abc9c", "未知": "#95a5a6"
    }

    unique_emotions = df['Emotion'].fillna("未知").unique()

    for emotion in unique_emotions:
        fg = folium.FeatureGroup(name=str(emotion))
        subset = df[df['Emotion'] == emotion]
        if subset.empty: continue
        
        heat_data = [[row['Latitude'], row['Longitude'], 1] for _, row in subset.iterrows()]
        HeatMap(heat_data, radius=20, blur=15, min_opacity=0.4, gradient=heatmap_gradients.get(emotion, None), name=f"{emotion} (热力)").add_to(fg)
        
        marker_color = marker_colors.get(emotion, "#ecf0f1")
        for _, row in subset.iterrows():
            folium.CircleMarker(
                location=[row['Latitude'], row['Longitude']], radius=3, color=marker_color,
                fill=True, fill_color=marker_color, fill_opacity=0.8, weight=0,
                popup=folium.Popup(f"<b>{row.get('Title', '无题')}</b><br>{emotion}", max_width=200),
                tooltip=f"{row.get('Title', '无题')}"
            ).add_to(fg)
        fg.add_to(m)

    folium.LayerControl(collapsed=False).add_to(m)
    return m

# --- 6. 页面渲染组件 ---

def render_ai_analysis_card(df, period_name):
    """渲染诗歌智能分析卡片"""
    st.markdown("---")
    st.subheader("🤖 智能诗歌检索与情感解析")
    st.caption("选择下方的年份、地点与情感，AI 将为您深度解读李白的心境。")
    
    with st.container(border=True):
        col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
        
        # 1. 筛选年份
        available_years = sorted(df['Year'].dropna().unique())
        with col1:
            selected_year = st.selectbox("1️⃣ 选择年份", available_years, key=f"year_{period_name}")
        
        # 2. 筛选地点 (基于年份联动)
        year_subset = df[df['Year'] == selected_year]
        available_locs = sorted(year_subset['Location'].dropna().unique())
        with col2:
            selected_loc = st.selectbox("2️⃣ 选择地点", available_locs, key=f"loc_{period_name}")
            
        # 3. 筛选情感 (基于地点联动)
        loc_subset = year_subset[year_subset['Location'] == selected_loc]
        available_emotions = sorted(loc_subset['Emotion'].dropna().unique())
        with col3:
            selected_emotion = st.selectbox("3️⃣ 选择情感", available_emotions, key=f"emo_{period_name}")
            
        # 4. 获取对应诗歌列表
        final_subset = loc_subset[loc_subset['Emotion'] == selected_emotion]
        # 修复：强制转字符串并去重
        available_titles = sorted(final_subset['Title'].astype(str).unique().tolist())
        
        with col4:
            if not available_titles:
                st.warning("该组合下暂无数据")
                selected_title = None
            else:
                selected_title = st.selectbox("4️⃣ 选择诗歌", available_titles, key=f"title_{period_name}")
        
        # 5. 提交按钮与生成
        # 使用 container_width 让按钮更醒目
        if st.button("✨ 生成 AI 深度解析", key=f"btn_{period_name}", use_container_width=True):
            if selected_title and selected_title != 'nan' and selected_title != '未知':
                with st.spinner(f"DeepSeek 正在阅读《{selected_title}》并分析历史背景..."):
                    cbdb_data = get_cbdb_data("李白")
                    analysis = generate_poem_analysis(selected_year, selected_loc, selected_emotion, selected_title, cbdb_data)
                    
                    st.markdown("---")
                    st.success("✅ 分析完成")
                    st.markdown(analysis)
            else:
                st.error("请先选择一首有效的诗歌。")

def render_home_page():
    st.header("🐉 李白生平 GIS 地图与 Chatbot 交互系统")
    cbdb_data = get_cbdb_data("李白")
    
    if st.session_state.data_df.empty:
        st.error("主页数据加载失败，请刷新重试。")
        if st.button("🔄 重试"): st.rerun()
        return

    with st.container():
        col1, col2 = st.columns([1, 1.5], gap="large")
        with col1:
            st.subheader("💬 CBDB-RAG 李白 Chatbot")
            if not cbdb_data: st.warning("CBDB 连接失败，使用通用知识库。")
            for msg in st.session_state.chat_history: st.chat_message(msg["role"]).markdown(msg["content"])
            if prompt := st.chat_input("请输入问题（例如：李白在安陆有哪些经历？）"):
                st.session_state.chat_history.append({"role": "user", "content": prompt})
                st.chat_message("user").markdown(prompt)
                with st.chat_message("assistant"):
                    with st.spinner("AI正在思考..."):
                        answer = run_main_chatbot(cbdb_data, prompt)
                        st.markdown(answer)
                        st.session_state.chat_history.append({"role": "assistant", "content": answer})
                        if st.session_state.highlight_location_key: st.success(f"地图已高亮：{st.session_state.highlight_location_key}")
                st.rerun()
        with col2:
            st.subheader("🗺️ 李白一生完整足迹可视化")
            current_map = create_main_map(st.session_state.data_df, st.session_state.highlight_location_key)
            st_folium(current_map, width=800, height=700)

def render_emotion_page(period):
    titles = {
        "youth": "🌱 青年期情感 GIS 热力图 ( < 742年 )",
        "middle": "🔥 中年期情感 GIS 热力图 ( 742-755年 )",
        "old": "🍂 晚年期情感 GIS 热力图 ( > 755年 )"
    }
    st.header(titles[period])
    df = load_emotion_data_from_github()
    if df.empty:
        st.error("❌ 情感数据下载失败。")
        if st.button("🔄 重试"): st.rerun()
        return

    df['Year'] = pd.to_numeric(df['Year'], errors='coerce')
    df = df.dropna(subset=['Year'])

    if period == "youth": filtered_df = df[df['Year'] < 742]
    elif period == "middle": filtered_df = df[(df['Year'] >= 742) & (df['Year'] <= 755)]
    else: filtered_df = df[df['Year'] > 755]

    st.info(f"共检索到 {len(filtered_df)} 首相关诗作。右上角可切换不同情感热力层。")

    # 1. 地图区域
    emotion_map = create_emotion_heatmap(filtered_df, period)
    st_folium(emotion_map, width="100%", height=600)
    
    # 2. AI 分析卡片区域
    render_ai_analysis_card(filtered_df, period)

# --- 7. 主程序入口 ---

def main():
    if "chat_history" not in st.session_state: st.session_state.chat_history = []
    if "highlight_location_key" not in st.session_state: st.session_state.highlight_location_key = None
    if "data_df" not in st.session_state: st.session_state.data_df = load_main_data(XLSX_FILENAME)

    with st.sidebar:
        st.title("📜 导航菜单")
        page = st.radio("选择视图模式:", ["🏠 主页 (RAG & 足迹)", "🌱 青年期情感热力图", "🔥 中年期情感热力图", "🍂 晚年期情感热力图"])
        st.markdown("---")
        st.caption("数据源：CBDB / GitHub")

    if page == "🏠 主页 (RAG & 足迹)": render_home_page()
    elif page == "🌱 青年期情感热力图": render_emotion_page("youth")
    elif page == "🔥 中年期情感热力图": render_emotion_page("middle")
    elif page == "🍂 晚年期情感热力图": render_emotion_page("old")

if __name__ == "__main__":
    main()
