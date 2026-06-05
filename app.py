import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date
import plotly.graph_objects as go
import io

# Setup page configuration
st.set_page_config(
    page_title="Quản Lý Việc Nhà & Lương - Bond & Sushi",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Constants
DB_PATH = "database.json"
EXCEL_PATH = "bang_gia.xlsx"

# ----------------------------------------------------
# DATABASE OPERATIONS
# ----------------------------------------------------
def init_db():
    if not os.path.exists(DB_PATH):
        default_data = {
            "settings": {
                "base_salary": {"Bond": 0, "Sushi": 0}
            },
            "daily_logs": {}
        }
        with open(DB_PATH, "w", encoding="utf-8") as f:
            json.dump(default_data, f, ensure_ascii=False, indent=4)

def load_db():
    init_db()
    try:
        with open(DB_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        st.error(f"Lỗi tải cơ sở dữ liệu JSON: {e}")
        return {"settings": {"base_salary": {"Bond": 0, "Sushi": 0}}, "daily_logs": {}}

def save_db(data):
    try:
        with open(DB_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        st.error(f"Lỗi ghi dữ liệu JSON: {e}")
        return False

# ----------------------------------------------------
# EXCEL OPERATIONS (Price Sheet)
# ----------------------------------------------------
def load_prices():
    default_prices = {
        "lau nhà, quét nhà, hút bụi": 20000,
        "rửa chén": 10000,
        "phơi đồ, gắp đồ": 10000,
        "không tắt đèn": -20000,
        "không tắt quạt": -20000,
        "quên uống sữa": -20000
    }
    
    if os.path.exists(EXCEL_PATH):
        try:
            df = pd.read_excel(EXCEL_PATH)
            # Normalize column names to match
            df.columns = [str(col).strip() for col in df.columns]
            
            # Find the columns
            job_col = None
            price_col = None
            for col in df.columns:
                if 'công việc' in col.lower() or 'job' in col.lower():
                    job_col = col
                if 'bảng giá' in col.lower() or 'giá' in col.lower() or 'price' in col.lower():
                    price_col = col
            
            if job_col and price_col:
                prices = {}
                for _, row in df.iterrows():
                    job = str(row[job_col]).strip()
                    try:
                        price = int(row[price_col])
                    except:
                        price = 0
                    if job:
                        prices[job] = price
                # Ensure we have fallback if some items are missing
                for k, v in default_prices.items():
                    if k not in prices:
                        prices[k] = v
                return prices
        except Exception as e:
            st.warning(f"Không thể đọc file Excel '{EXCEL_PATH}' ({e}). Sử dụng bảng giá mặc định.")
            
    return default_prices

def save_prices_to_excel(prices_dict):
    try:
        # Create Excel
        df = pd.DataFrame(list(prices_dict.items()), columns=['công việc', 'bảng giá'])
        df.to_excel(EXCEL_PATH, index=False)
        return True
    except Exception as e:
        st.error(f"Không thể ghi bảng giá ra file Excel: {e}")
        return False

# ----------------------------------------------------
# DATES AND WEEKDAYS UTILS
# ----------------------------------------------------
def get_vietnamese_weekday(dt_date):
    # Mon=0, Tue=1, ..., Sun=6
    wd = dt_date.weekday()
    wd_names = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"]
    return wd_names[wd]

def get_default_assigned_person(dt_date):
    wd = dt_date.weekday()
    if wd in [0, 2, 4]:  # Thứ 2, 4, 6 (Mon, Wed, Fri) -> Sushi
        return "Sushi"
    elif wd in [1, 3, 5]:  # Thứ 3, 5, 7 (Tue, Thu, Sat) -> Bond
        return "Bond"
    else:  # Chủ nhật (Sun) -> Không có
        return "Không có"

# ----------------------------------------------------
# WAGE & BREAKDOWN CALCULATION LOGIC
# ----------------------------------------------------
def calculate_daily_breakdown(date_str, log_data, prices):
    reward_lau_nha = prices.get("lau nhà, quét nhà, hút bụi", 20000)
    reward_rua_chen = prices.get("rửa chén", 10000)
    reward_phoi_do = prices.get("phơi đồ, gắp đồ", 10000)
    
    # Penalties are negative in Excel, we convert to absolute to subtract or add signed values
    penalty_light = abs(prices.get("không tắt đèn", -20000))
    penalty_fan = abs(prices.get("không tắt quạt", -20000))
    penalty_milk = abs(prices.get("quên uống sữa", -20000))
    
    assigned = log_data.get("assigned_to", "Không có")
    chores = log_data.get("chores", {})
    milk = log_data.get("milk", {"Bond": True, "Sushi": True})
    appliances = log_data.get("appliances", {"Bond": {"light": 0, "fan": 0}, "Sushi": {"light": 0, "fan": 0}})
    
    breakdown = {
        "Bond": {"chores_earned": 0, "chores_penalties": 0, "milk_penalty": 0, "appliance_penalty": 0, "total": 0, "details": []},
        "Sushi": {"chores_earned": 0, "chores_penalties": 0, "milk_penalty": 0, "appliance_penalty": 0, "total": 0, "details": []}
    }
    
    for name in ["Bond", "Sushi"]:
        details = []
        earned_chores = 0
        penalty_chores = 0
        p_milk = 0
        p_app = 0
        
        # 1. Chores calculation
        if "Bond" in chores or "Sushi" in chores:
            person_chores = chores.get(name, {})
            has_chores = True
        else:
            # Old format backward compatibility
            person_chores = chores if assigned == name else {}
            has_chores = (assigned == name)
            
        if has_chores:
            # Lau nha
            status_ln = person_chores.get("lau_nha", "Không giao")
            if status_ln == "Hoàn thành":
                earned_chores += reward_lau_nha
                details.append(f"Làm đúng combo lau/quét nhà: +{reward_lau_nha:,}đ")
            elif status_ln == "Chưa làm / Trễ":
                penalty_chores += 20000
                details.append(f"Trễ/Chưa làm quét dọn (sau 11h30): -20,000đ")
                
            # Rua chen
            status_rc = person_chores.get("rua_chen", "Không giao")
            if status_rc == "Hoàn thành":
                earned_chores += reward_rua_chen
                details.append(f"Làm đúng việc rửa chén: +{reward_rua_chen:,}đ")
            elif status_rc == "Chưa làm / Trễ":
                penalty_chores += 20000
                details.append(f"Trễ/Chưa rửa chén (sau 11h30): -20,000đ")
                
            # Phoi do
            status_pd = person_chores.get("phoi_do", "Không giao")
            if status_pd == "Hoàn thành":
                earned_chores += reward_phoi_do
                details.append(f"Làm đúng việc phơi/gấp đồ: +{reward_phoi_do:,}đ")
            elif status_pd == "Chưa làm / Trễ":
                penalty_chores += 20000
                details.append(f"Trễ/Chưa phơi đồ (sau 11h30): -20,000đ")
        
        # 2. Morning Milk Check
        drunk_milk = milk.get(name, True)
        if not drunk_milk:
            p_milk = penalty_milk
            details.append(f"Quên uống sữa sáng: -{penalty_milk:,}đ")
            
        # 3. Appliances Penalty (Light/Fan)
        app_forgot = appliances.get(name, {"light": 0, "fan": 0})
        lights = app_forgot.get("light", 0)
        fans = app_forgot.get("fan", 0)
        
        if lights > 0:
            p_light_total = lights * penalty_light
            p_app += p_light_total
            details.append(f"Quên tắt đèn ({lights} lần): -{p_light_total:,}đ")
        if fans > 0:
            p_fan_total = fans * penalty_fan
            p_app += p_fan_total
            details.append(f"Quên tắt quạt ({fans} lần): -{p_fan_total:,}đ")
            
        total_change = earned_chores - penalty_chores - p_milk - p_app
        
        breakdown[name]["chores_earned"] = earned_chores
        breakdown[name]["chores_penalties"] = penalty_chores
        breakdown[name]["milk_penalty"] = p_milk
        breakdown[name]["appliance_penalty"] = p_app
        breakdown[name]["total"] = total_change
        breakdown[name]["details"] = details
        
    return breakdown

def calculate_monthly_summary(year, month, db_data, prices):
    target_prefix = f"{year:04d}-{month:02d}-"
    bond_summary = {"base": db_data["settings"]["base_salary"].get("Bond", 0), "chores_earned": 0, "chores_penalties": 0, "milk_penalty": 0, "appliance_penalty": 0, "total": 0, "logs_count": 0, "chores_done_count": 0, "violations_count": 0}
    sushi_summary = {"base": db_data["settings"]["base_salary"].get("Sushi", 0), "chores_earned": 0, "chores_penalties": 0, "milk_penalty": 0, "appliance_penalty": 0, "total": 0, "logs_count": 0, "chores_done_count": 0, "violations_count": 0}
    
    daily_details = []
    
    for date_str, log_data in sorted(db_data["daily_logs"].items()):
        if date_str.startswith(target_prefix):
            day_breakdown = calculate_daily_breakdown(date_str, log_data, prices)
            
            # Extract day number
            day_num = date_str.split("-")[2]
            weekday_str = get_vietnamese_weekday(datetime.strptime(date_str, "%Y-%m-%d").date())
            
            # Accumulate
            for name, summary in [("Bond", bond_summary), ("Sushi", sushi_summary)]:
                summary["chores_earned"] += day_breakdown[name]["chores_earned"]
                summary["chores_penalties"] += day_breakdown[name]["chores_penalties"]
                summary["milk_penalty"] += day_breakdown[name]["milk_penalty"]
                summary["appliance_penalty"] += day_breakdown[name]["appliance_penalty"]
                summary["logs_count"] += 1
                
                # count chores done
                chores = log_data.get("chores", {})
                if "Bond" in chores or "Sushi" in chores:
                    person_chores = chores.get(name, {})
                else:
                    person_chores = chores if log_data.get("assigned_to") == name else {}
                for c_status in person_chores.values():
                    if c_status == "Hoàn thành":
                        summary["chores_done_count"] += 1
                
                # count violations (milk, light, fan)
                if not log_data.get("milk", {}).get(name, True):
                    summary["violations_count"] += 1
                app_forgot = log_data.get("appliances", {}).get(name, {"light": 0, "fan": 0})
                summary["violations_count"] += app_forgot.get("light", 0) + app_forgot.get("fan", 0)
                
            assigned_val = log_data.get("assigned_to", "Không có")
            if "Bond" in log_data.get("chores", {}) or "Sushi" in log_data.get("chores", {}):
                assigned_val = "Cả hai"
                
            daily_details.append({
                "Ngày": f"Ngày {day_num} ({weekday_str})",
                "Phân công": assigned_val,
                "Bond_Chore": day_breakdown["Bond"]["chores_earned"] - day_breakdown["Bond"]["chores_penalties"],
                "Bond_Milk": -day_breakdown["Bond"]["milk_penalty"],
                "Bond_Elec": -day_breakdown["Bond"]["appliance_penalty"],
                "Bond_Day_Total": day_breakdown["Bond"]["total"],
                "Sushi_Chore": day_breakdown["Sushi"]["chores_earned"] - day_breakdown["Sushi"]["chores_penalties"],
                "Sushi_Milk": -day_breakdown["Sushi"]["milk_penalty"],
                "Sushi_Elec": -day_breakdown["Sushi"]["appliance_penalty"],
                "Sushi_Day_Total": day_breakdown["Sushi"]["total"],
                "Ghi chú": log_data.get("notes", "")
            })
            
    bond_summary["total"] = bond_summary["base"] + bond_summary["chores_earned"] - bond_summary["chores_penalties"] - bond_summary["milk_penalty"] - bond_summary["appliance_penalty"]
    sushi_summary["total"] = sushi_summary["base"] + sushi_summary["chores_earned"] - sushi_summary["chores_penalties"] - sushi_summary["milk_penalty"] - sushi_summary["appliance_penalty"]
    
    return bond_summary, sushi_summary, daily_details

# Load data and prices
db = load_db()
prices = load_prices()

# ----------------------------------------------------
# APPLICATION STYLING & CUSTOM ELEMENTS
# ----------------------------------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    .app-title {
        font-size: 2.5rem;
        font-weight: 800;
        background: linear-gradient(135deg, #1e3a8a 0%, #ec4899 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 5px;
        text-align: center;
    }
    
    .app-subtitle {
        font-size: 1.1rem;
        color: #64748b;
        text-align: center;
        margin-bottom: 25px;
    }
    
    /* Styled container cards for the kids */
    .card-container {
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 4px 15px -3px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
        color: #1e293b;
    }
    
    .card-bond {
        background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
        border: 1px solid #bfdbfe;
        border-left: 8px solid #3b82f6;
    }
    
    .card-sushi {
        background: linear-gradient(135deg, #fdf2f8 0%, #fce7f3 100%);
        border: 1px solid #fbcfe8;
        border-left: 8px solid #ec4899;
    }
    
    .card-title {
        font-size: 1.4rem;
        font-weight: 700;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .card-metric-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 12px;
        margin-top: 15px;
    }
    
    .card-metric-box {
        background: rgba(255, 255, 255, 0.6);
        padding: 10px;
        border-radius: 8px;
        text-align: center;
    }
    
    .card-metric-label {
        font-size: 0.8rem;
        color: #64748b;
        font-weight: 600;
        text-transform: uppercase;
    }
    
    .card-metric-val {
        font-size: 1.1rem;
        font-weight: 700;
    }
    
    .card-salary-box {
        background: rgba(255, 255, 255, 0.9);
        padding: 15px;
        border-radius: 12px;
        text-align: center;
        margin-top: 15px;
        border: 1px dashed rgba(0, 0, 0, 0.1);
    }
    
    .card-salary-label {
        font-size: 0.9rem;
        color: #475569;
        font-weight: 700;
    }
    
    .card-salary-val {
        font-size: 1.8rem;
        font-weight: 800;
    }
    
    .color-bond { color: #1d4ed8; }
    .color-sushi { color: #be185d; }
    
    /* Utility colors */
    .text-success { color: #16a34a; font-weight: 600; }
    .text-danger { color: #dc2626; font-weight: 600; }
    .text-muted { color: #94a3b8; }
</style>
""", unsafe_allow_html=True)

# App header
st.markdown("<div class='app-title'>🏠 Quản Lý Việc Nhà & Lương Thưởng</div>", unsafe_allow_html=True)
st.markdown("<div class='app-subtitle'>Hệ thống theo dõi nhiệm vụ trực nhật, điểm danh uống sữa và phạt thiết bị điện của Bond & Sushi</div>", unsafe_allow_html=True)

# ----------------------------------------------------
# SIDEBAR
# ----------------------------------------------------
st.sidebar.markdown("### 📅 Chọn thời gian báo cáo")
today = date.today()
current_year = today.year

selected_month = st.sidebar.selectbox(
    "Chọn tháng",
    options=list(range(1, 13)),
    index=today.month - 1,
    format_func=lambda m: f"Tháng {m}"
)

selected_year = st.sidebar.number_input(
    "Chọn năm",
    min_value=2020,
    max_value=2100,
    value=current_year,
    step=1
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Tổng quan nhanh")
# Fetch metrics for the sidebar summary
bond_sum, sushi_sum, df_details = calculate_monthly_summary(selected_year, selected_month, db, prices)

st.sidebar.markdown(f"**Bond**: <span class='text-success'>{bond_sum['total']:,}đ</span>", unsafe_allow_html=True)
st.sidebar.markdown(f"**Sushi**: <span class='text-success'>{sushi_sum['total']:,}đ</span>", unsafe_allow_html=True)
st.sidebar.markdown(f"Số ngày đã ghi nhận: **{bond_sum['logs_count']} ngày**")

st.sidebar.markdown("---")
st.sidebar.info("💡 **Mẹo:** Vào tab **Cài đặt & Bảng giá** để đổi tiền thưởng trong file Excel hoặc cài đặt lương cứng.")

# ----------------------------------------------------
# MAIN DASHBOARD TABS
# ----------------------------------------------------
tab_dash, tab_checkin, tab_report, tab_history, tab_settings = st.tabs([
    "📊 Bảng điều khiển (Dashboard)", 
    "📝 Điểm danh hàng ngày", 
    "📋 Báo cáo lương chi tiết", 
    "📜 Lịch sử nhật ký",
    "⚙️ Cài đặt & Bảng giá Excel"
])

# ----------------------------------------------------
# TAB 1: DASHBOARD
# ----------------------------------------------------
with tab_dash:
    st.markdown(f"### 📈 Kết quả tháng {selected_month}/{selected_year}")
    
    col1, col2 = st.columns(2)
    
    # Bond Card
    with col1:
        st.markdown(f"""
        <div class="card-container card-bond">
            <div class="card-title"><span style="font-size:1.8rem;">👦</span> Anh hai Bond</div>
            <hr style="border: 0; border-top: 1px solid #bfdbfe; margin: 10px 0;">
            <div class="card-metric-grid">
                <div class="card-metric-box">
                    <div class="card-metric-label">Lương cơ bản</div>
                    <div class="card-metric-val">{bond_sum['base']:,}đ</div>
                </div>
                <div class="card-metric-box">
                    <div class="card-metric-label">Thưởng việc nhà</div>
                    <div class="card-metric-val text-success">+{bond_sum['chores_earned']:,}đ</div>
                </div>
                <div class="card-metric-box">
                    <div class="card-metric-label">Phạt việc nhà</div>
                    <div class="card-metric-val text-danger">-{bond_sum['chores_penalties']:,}đ</div>
                </div>
                <div class="card-metric-box">
                    <div class="card-metric-label">Phạt uống sữa</div>
                    <div class="card-metric-val text-danger">-{bond_sum['milk_penalty']:,}đ</div>
                </div>
                <div class="card-metric-box" style="grid-column: span 2;">
                    <div class="card-metric-label">Phạt thiết bị điện</div>
                    <div class="card-metric-val text-danger">-{bond_sum['appliance_penalty']:,}đ</div>
                </div>
            </div>
            <div class="card-salary-box">
                <div class="card-salary-label color-bond">TỔNG LƯƠNG NHẬN ĐƯỢC</div>
                <div class="card-salary-val color-bond">{bond_sum['total']:,}đ</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    # Sushi Card
    with col2:
        st.markdown(f"""
        <div class="card-container card-sushi">
            <div class="card-title"><span style="font-size:1.8rem;">👧</span> Em gái Sushi</div>
            <hr style="border: 0; border-top: 1px solid #fbcfe8; margin: 10px 0;">
            <div class="card-metric-grid">
                <div class="card-metric-box">
                    <div class="card-metric-label">Lương cơ bản</div>
                    <div class="card-metric-val">{sushi_sum['base']:,}đ</div>
                </div>
                <div class="card-metric-box">
                    <div class="card-metric-label">Thưởng việc nhà</div>
                    <div class="card-metric-val text-success">+{sushi_sum['chores_earned']:,}đ</div>
                </div>
                <div class="card-metric-box">
                    <div class="card-metric-label">Phạt việc nhà</div>
                    <div class="card-metric-val text-danger">-{sushi_sum['chores_penalties']:,}đ</div>
                </div>
                <div class="card-metric-box">
                    <div class="card-metric-label">Phạt uống sữa</div>
                    <div class="card-metric-val text-danger">-{sushi_sum['milk_penalty']:,}đ</div>
                </div>
                <div class="card-metric-box" style="grid-column: span 2;">
                    <div class="card-metric-label">Phạt thiết bị điện</div>
                    <div class="card-metric-val text-danger">-{sushi_sum['appliance_penalty']:,}đ</div>
                </div>
            </div>
            <div class="card-salary-box">
                <div class="card-salary-label color-sushi">TỔNG LƯƠNG NHẬN ĐƯỢC</div>
                <div class="card-salary-val color-sushi">{sushi_sum['total']:,}đ</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # Visualizations
    st.markdown("### 📊 Biểu đồ thống kê")
    if bond_sum['logs_count'] > 0:
        # Create a Plotly chart showing earnings comparison
        categories = ['Thưởng việc nhà', 'Phạt việc nhà', 'Phạt uống sữa', 'Phạt điện/quạt']
        
        fig = go.Figure()
        
        # Bond Trace
        fig.add_trace(go.Bar(
            name='Bond',
            x=categories,
            y=[bond_sum['chores_earned'], -bond_sum['chores_penalties'], -bond_sum['milk_penalty'], -bond_sum['appliance_penalty']],
            marker_color='#3b82f6'
        ))
        
        # Sushi Trace
        fig.add_trace(go.Bar(
            name='Sushi',
            x=categories,
            y=[sushi_sum['chores_earned'], -sushi_sum['chores_penalties'], -sushi_sum['milk_penalty'], -sushi_sum['appliance_penalty']],
            marker_color='#ec4899'
        ))
        
        fig.update_layout(
            barmode='group',
            title='So sánh giá trị thưởng và các khoản phạt của 2 bạn (VND)',
            yaxis_title='Số tiền (VND)',
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            legend_font_family='Plus Jakarta Sans',
            font_family='Plus Jakarta Sans'
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # Simple pie chart for chore completion
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            fig_bond_chores = go.Figure(data=[go.Pie(
                labels=['Hoàn thành đúng giờ', 'Chưa hoàn thành / Trễ hạn / Lỗi khác'],
                values=[bond_sum['chores_done_count'], max(0, bond_sum['violations_count'] - bond_sum['chores_done_count'])],
                hole=.3,
                marker_colors=['#3b82f6', '#f87171']
            )])
            fig_bond_chores.update_layout(title_text='Tỉ lệ hoàn thành nhiệm vụ & Vi phạm - Bond', font_family='Plus Jakarta Sans')
            st.plotly_chart(fig_bond_chores, use_container_width=True)
            
        with col_c2:
            fig_sushi_chores = go.Figure(data=[go.Pie(
                labels=['Hoàn thành đúng giờ', 'Chưa hoàn thành / Trễ hạn / Lỗi khác'],
                values=[sushi_sum['chores_done_count'], max(0, sushi_sum['violations_count'] - sushi_sum['chores_done_count'])],
                hole=.3,
                marker_colors=['#ec4899', '#f87171']
            )])
            fig_sushi_chores.update_layout(title_text='Tỉ lệ hoàn thành nhiệm vụ & Vi phạm - Sushi', font_family='Plus Jakarta Sans')
            st.plotly_chart(fig_sushi_chores, use_container_width=True)
    else:
        st.info("Chưa có dữ liệu nhật ký cho tháng này để hiển thị biểu đồ.")

# ----------------------------------------------------
# TAB 2: DAILY CHECK-IN
# ----------------------------------------------------
with tab_checkin:
    st.markdown("### 📝 Điểm danh & Ghi nhận việc nhà hàng ngày")
    
    # Date Input
    check_date = st.date_input("Chọn ngày điểm danh", value=date.today())
    date_str = check_date.strftime("%Y-%m-%d")
    weekday_vn = get_vietnamese_weekday(check_date)
    
    # Check if log already exists
    existing_log = db["daily_logs"].get(date_str, {})
    if existing_log:
        st.warning(f"⚠️ Đã có dữ liệu điểm danh cho ngày {date_str} ({weekday_vn}). Nếu bạn sửa đổi và lưu, dữ liệu cũ sẽ bị ghi đè.")
    
    # Pre-calculated default values
    default_assignee = get_default_assigned_person(check_date)
    
    # Load existing chores or set defaults
    chores_data = existing_log.get("chores", {})
    
    # Check if existing log is new or old format
    if "Bond" in chores_data or "Sushi" in chores_data:
        bond_chores_existing = chores_data.get("Bond", {})
        sushi_chores_existing = chores_data.get("Sushi", {})
    else:
        # Old format or no existing log
        old_assigned = existing_log.get("assigned_to", default_assignee)
        if old_assigned == "Bond":
            bond_chores_existing = chores_data
            sushi_chores_existing = {}
        elif old_assigned == "Sushi":
            sushi_chores_existing = chores_data
            bond_chores_existing = {}
        else:
            bond_chores_existing = {}
            sushi_chores_existing = {}

    # Determine default values for new entries based on weekday
    if not existing_log:
        if default_assignee == "Bond":
            default_bond_status = "Hoàn thành"
            default_sushi_status = "Không giao"
        elif default_assignee == "Sushi":
            default_bond_status = "Không giao"
            default_sushi_status = "Hoàn thành"
        else:
            default_bond_status = "Không giao"
            default_sushi_status = "Không giao"
    else:
        default_bond_status = "Không giao"
        default_sushi_status = "Không giao"

    # Get status for each chore
    default_ln_bond = bond_chores_existing.get("lau_nha", default_bond_status)
    default_rc_bond = bond_chores_existing.get("rua_chen", default_bond_status)
    default_pd_bond = bond_chores_existing.get("phoi_do", default_bond_status)
    
    default_ln_sushi = sushi_chores_existing.get("lau_nha", default_sushi_status)
    default_rc_sushi = sushi_chores_existing.get("rua_chen", default_sushi_status)
    default_pd_sushi = sushi_chores_existing.get("phoi_do", default_sushi_status)

    # Form for check-in
    with st.form("daily_checkin_form"):
        st.markdown(f"#### Thông tin điểm danh ngày **{date_str} ({weekday_vn})**")
        
        # Chores inputs for both kids
        st.markdown("---")
        st.markdown("#### 🧹 Tình hình làm việc nhà (Cả 2 bạn đều được phân công hàng ngày)")
        
        col_bond_chores, col_sushi_chores = st.columns(2)
        
        with col_bond_chores:
            st.markdown("##### 👦 Việc nhà của Bond")
            status_ln_bond = st.radio(
                "Lau, quét nhà & hút bụi (Bond - Combo: 20k):",
                options=["Hoàn thành", "Chưa làm / Trễ", "Không giao"],
                index=["Hoàn thành", "Chưa làm / Trễ", "Không giao"].index(default_ln_bond),
                key="ln_bond"
            )
            status_rc_bond = st.radio(
                "Rửa chén (Bond - 10k):",
                options=["Hoàn thành", "Chưa làm / Trễ", "Không giao"],
                index=["Hoàn thành", "Chưa làm / Trễ", "Không giao"].index(default_rc_bond),
                key="rc_bond"
            )
            status_pd_bond = st.radio(
                "Phơi đồ & gấp đồ (Bond - 10k):",
                options=["Hoàn thành", "Chưa làm / Trễ", "Không giao"],
                index=["Hoàn thành", "Chưa làm / Trễ", "Không giao"].index(default_pd_bond),
                key="pd_bond"
            )
            
        with col_sushi_chores:
            st.markdown("##### 👧 Việc nhà của Sushi")
            status_ln_sushi = st.radio(
                "Lau, quét nhà & hút bụi (Sushi - Combo: 20k):",
                options=["Hoàn thành", "Chưa làm / Trễ", "Không giao"],
                index=["Hoàn thành", "Chưa làm / Trễ", "Không giao"].index(default_ln_sushi),
                key="ln_sushi"
            )
            status_rc_sushi = st.radio(
                "Rửa chén (Sushi - 10k):",
                options=["Hoàn thành", "Chưa làm / Trễ", "Không giao"],
                index=["Hoàn thành", "Chưa làm / Trễ", "Không giao"].index(default_rc_sushi),
                key="rc_sushi"
            )
            status_pd_sushi = st.radio(
                "Phơi đồ & gấp đồ (Sushi - 10k):",
                options=["Hoàn thành", "Chưa làm / Trễ", "Không giao"],
                index=["Hoàn thành", "Chưa làm / Trễ", "Không giao"].index(default_pd_sushi),
                key="pd_sushi"
            )
            
        st.markdown("---")
        st.markdown("#### 🥛 Tình hình uống sữa sáng (Áp dụng cho cả 2)")
        
        milk_data = existing_log.get("milk", {"Bond": True, "Sushi": True})
        
        col_mb, col_ms = st.columns(2)
        with col_mb:
            milk_bond = st.checkbox("Bond đã uống sữa sáng đầy đủ", value=milk_data.get("Bond", True))
        with col_ms:
            milk_sushi = st.checkbox("Sushi đã uống sữa sáng đầy đủ", value=milk_data.get("Sushi", True))
            
        st.markdown("---")
        st.markdown("#### 🔌 Quên tắt thiết bị điện (Phạt 20k / lần quên)")
        
        app_data = existing_log.get("appliances", {"Bond": {"light": 0, "fan": 0}, "Sushi": {"light": 0, "fan": 0}})
        
        col_ab, col_as = st.columns(2)
        with col_ab:
            st.markdown("**Anh hai Bond**")
            light_bond = st.number_input("Số lần quên tắt đèn (Bond):", min_value=0, max_value=10, value=app_data.get("Bond", {}).get("light", 0), step=1)
            fan_bond = st.number_input("Số lần quên tắt quạt (Bond):", min_value=0, max_value=10, value=app_data.get("Bond", {}).get("fan", 0), step=1)
            
        with col_as:
            st.markdown("**Em gái Sushi**")
            light_sushi = st.number_input("Số lần quên tắt đèn (Sushi):", min_value=0, max_value=10, value=app_data.get("Sushi", {}).get("light", 0), step=1)
            fan_sushi = st.number_input("Số lần quên tắt quạt (Sushi):", min_value=0, max_value=10, value=app_data.get("Sushi", {}).get("fan", 0), step=1)
            
        st.markdown("---")
        notes = st.text_area("Ghi chú thêm (nếu có):", value=existing_log.get("notes", ""))
        
        # Validation checks
        warnings = []
        if status_ln_bond == "Hoàn thành" and status_ln_sushi == "Hoàn thành":
            warnings.append("⚠️ Cảnh báo: Lau/quét nhà đang được đánh dấu Hoàn thành cho cả 2 bạn.")
        if status_rc_bond == "Hoàn thành" and status_rc_sushi == "Hoàn thành":
            warnings.append("⚠️ Cảnh báo: Rửa chén đang được đánh dấu Hoàn thành cho cả 2 bạn.")
        if status_pd_bond == "Hoàn thành" and status_pd_sushi == "Hoàn thành":
            warnings.append("⚠️ Cảnh báo: Phơi/gấp đồ đang được đánh dấu Hoàn thành cho cả 2 bạn.")
            
        for warning_msg in warnings:
            st.warning(warning_msg)
            
        # Submit button
        submit_btn = st.form_submit_button("💾 Lưu Nhật Ký Ngày Hôm Nay")
        
        if submit_btn:
            # Prepare data
            new_log = {
                "assigned_to": "Cả hai",
                "chores": {
                    "Bond": {
                        "lau_nha": status_ln_bond,
                        "rua_chen": status_rc_bond,
                        "phoi_do": status_pd_bond
                    },
                    "Sushi": {
                        "lau_nha": status_ln_sushi,
                        "rua_chen": status_rc_sushi,
                        "phoi_do": status_pd_sushi
                    }
                },
                "milk": {
                    "Bond": milk_bond,
                    "Sushi": milk_sushi
                },
                "appliances": {
                    "Bond": {"light": int(light_bond), "fan": int(fan_bond)},
                    "Sushi": {"light": int(light_sushi), "fan": int(fan_sushi)}
                },
                "notes": notes
            }
            
            db["daily_logs"][date_str] = new_log
            if save_db(db):
                st.success(f"🎉 Đã lưu thành công nhật ký ngày {date_str}!")
                st.rerun()

# ----------------------------------------------------
# TAB 3: SALARY REPORT
# ----------------------------------------------------
with tab_report:
    st.markdown(f"### 📋 Chi tiết bảng lương Tháng {selected_month}/{selected_year}")
    
    bond_sum, sushi_sum, daily_details = calculate_monthly_summary(selected_year, selected_month, db, prices)
    
    if len(daily_details) == 0:
        st.info(f"Chưa có bản ghi nhật ký nào trong tháng {selected_month}/{selected_year}!")
    else:
        df_report = pd.DataFrame(daily_details)
        
        # Format the columns for nice visualization
        styled_df = df_report.copy()
        
        # Rename columns to Vietnamese
        styled_df.columns = [
            "Ngày", "Trực nhật", 
            "Bond - Việc nhà", "Bond - Phạt Sữa", "Bond - Phạt Điện/Quạt", "Bond - Tổng cộng ngày",
            "Sushi - Việc nhà", "Sushi - Phạt Sữa", "Sushi - Phạt Điện/Quạt", "Sushi - Tổng cộng ngày",
            "Ghi chú"
        ]
        
        # Display the table
        st.dataframe(
            styled_df.style.format({
                "Bond - Việc nhà": "{:+,}đ", "Bond - Phạt Sữa": "{:+,}đ", "Bond - Phạt Điện/Quạt": "{:+,}đ", "Bond - Tổng cộng ngày": "{:+,}đ",
                "Sushi - Việc nhà": "{:+,}đ", "Sushi - Phạt Sữa": "{:+,}đ", "Sushi - Phạt Điện/Quạt": "{:+,}đ", "Sushi - Tổng cộng ngày": "{:+,}đ"
            }),
            use_container_width=True
        )
        
        # Calculation formula walkthrough
        st.markdown("#### 🧮 Công thức tính lương chi tiết")
        col_f1, col_f2 = st.columns(2)
        
        with col_f1:
            st.markdown(f"""
            <div class="card-container card-bond">
                <div class="card-title">👦 Chi tiết Bond</div>
                <ul style="font-size: 0.95rem; line-height: 1.6;">
                    <li>Lương cơ bản: <b>{bond_sum['base']:,}đ</b></li>
                    <li>Tổng tiền thưởng việc nhà: <span class="text-success">+{bond_sum['chores_earned']:,}đ</span></li>
                    <li>Tổng tiền phạt chưa làm việc nhà: <span class="text-danger">-{bond_sum['chores_penalties']:,}đ</span></li>
                    <li>Tổng tiền phạt quên uống sữa: <span class="text-danger">-{bond_sum['milk_penalty']:,}đ</span></li>
                    <li>Tổng tiền phạt quên tắt điện/quạt: <span class="text-danger">-{bond_sum['appliance_penalty']:,}đ</span></li>
                    <li style="font-size:1.1rem; margin-top: 10px; border-top: 1px dashed rgba(0,0,0,0.2); padding-top:5px;">
                        <b>Tổng lương thực nhận: <span class="color-bond">{bond_sum['total']:,}đ</span></b>
                    </li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
            
        with col_f2:
            st.markdown(f"""
            <div class="card-container card-sushi">
                <div class="card-title">👧 Chi tiết Sushi</div>
                <ul style="font-size: 0.95rem; line-height: 1.6;">
                    <li>Lương cơ bản: <b>{sushi_sum['base']:,}đ</b></li>
                    <li>Tổng tiền thưởng việc nhà: <span class="text-success">+{sushi_sum['chores_earned']:,}đ</span></li>
                    <li>Tổng tiền phạt chưa làm việc nhà: <span class="text-danger">-{sushi_sum['chores_penalties']:,}đ</span></li>
                    <li>Tổng tiền phạt quên uống sữa: <span class="text-danger">-{sushi_sum['milk_penalty']:,}đ</span></li>
                    <li>Tổng tiền phạt quên tắt điện/quạt: <span class="text-danger">-{sushi_sum['appliance_penalty']:,}đ</span></li>
                    <li style="font-size:1.1rem; margin-top: 10px; border-top: 1px dashed rgba(0,0,0,0.2); padding-top:5px;">
                        <b>Tổng lương thực nhận: <span class="color-sushi">{sushi_sum['total']:,}đ</span></b>
                    </li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
            
        # Download Report Option
        st.markdown("#### 📥 Xuất báo cáo")
        
        # Generate Excel in-memory
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            # Details sheet
            styled_df.to_excel(writer, sheet_name=f"Chi tiết Tháng {selected_month}", index=False)
            
            # Summary sheet
            summary_data = {
                "Khoản mục": [
                    "Lương cơ bản", 
                    "Thưởng làm việc nhà (+)", 
                    "Phạt chưa làm việc nhà (-)", 
                    "Phạt quên uống sữa (-)", 
                    "Phạt quên tắt điện/quạt (-)", 
                    "TỔNG LƯƠNG NHẬN ĐƯỢC"
                ],
                "Bond (đ)": [
                    bond_sum['base'], 
                    bond_sum['chores_earned'], 
                    -bond_sum['chores_penalties'], 
                    -bond_sum['milk_penalty'], 
                    -bond_sum['appliance_penalty'], 
                    bond_sum['total']
                ],
                "Sushi (đ)": [
                    sushi_sum['base'], 
                    sushi_sum['chores_earned'], 
                    -sushi_sum['chores_penalties'], 
                    -sushi_sum['milk_penalty'], 
                    -sushi_sum['appliance_penalty'], 
                    sushi_sum['total']
                ]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name="Tổng hợp lương", index=False)
            
        excel_data = output_buffer.getvalue()
        
        st.download_button(
            label="📥 Tải Báo Cáo Excel (.xlsx)",
            data=excel_data,
            file_name=f"Bao_cao_luong_Nha_Bond_Sushi_{selected_year}_{selected_month:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ----------------------------------------------------
# TAB 4: HISTORY LOGS
# ----------------------------------------------------
with tab_history:
    st.markdown("### 📜 Lịch sử nhật ký điểm danh")
    
    logs = db.get("daily_logs", {})
    
    if not logs:
        st.info("Chưa có nhật ký nào được lưu.")
    else:
        # Create a list of logs
        log_list = []
        for d_str, log_info in sorted(logs.items(), reverse=True):
            wd = get_vietnamese_weekday(datetime.strptime(d_str, "%Y-%m-%d").date())
            
            # Formatted chores summary supporting both formats
            chores = log_info.get("chores", {})
            if "Bond" in chores or "Sushi" in chores:
                def format_chores_summary(chores_dict):
                    name_map = {"lau_nha": "Lau nhà", "rua_chen": "Rửa chén", "phoi_do": "Phơi đồ"}
                    done = []
                    late = []
                    for k, v in chores_dict.items():
                        vn_name = name_map.get(k, k)
                        if v == "Hoàn thành":
                            done.append(vn_name)
                        elif v == "Chưa làm / Trễ":
                            late.append(f"{vn_name} (Trễ)")
                    
                    parts = []
                    if done:
                        parts.append(f"Xong: {', '.join(done)}")
                    if late:
                        parts.append(f"Trễ: {', '.join(late)}")
                    return " + ".join(parts) if parts else "Không làm"

                bond_str = format_chores_summary(chores.get("Bond", {}))
                sushi_str = format_chores_summary(chores.get("Sushi", {}))
                chores_str = f"Bond ({bond_str}) | Sushi ({sushi_str})"
                assigned_to_val = "Cả hai"
            else:
                assigned_to_val = log_info.get("assigned_to", "Không có")
                if assigned_to_val != "Không có":
                    name_map = {"lau_nha": "Lau nhà", "rua_chen": "Rửa chén", "phoi_do": "Phơi đồ"}
                    done = [name_map.get(k, k) for k, v in chores.items() if v == "Hoàn thành"]
                    late = [f"{name_map.get(k, k)} (Trễ)" for k, v in chores.items() if v == "Chưa làm / Trễ"]
                    parts = []
                    if done:
                        parts.append(f"Xong: {', '.join(done)}")
                    if late:
                        parts.append(f"Trễ: {', '.join(late)}")
                    chores_str = f"{assigned_to_val} (" + (" + ".join(parts) if parts else "Không làm") + ")"
                else:
                    chores_str = "Không giao"
            
            milk_str = f"Bond: {'Đã uống' if log_info.get('milk', {}).get('Bond', True) else 'Quên'}, Sushi: {'Đã uống' if log_info.get('milk', {}).get('Sushi', True) else 'Quên'}"
            elec_str = f"Bond: {log_info.get('appliances', {}).get('Bond', {}).get('light', 0)} đèn/{log_info.get('appliances', {}).get('Bond', {}).get('fan', 0)} quạt, Sushi: {log_info.get('appliances', {}).get('Sushi', {}).get('light', 0)} đèn/{log_info.get('appliances', {}).get('Sushi', {}).get('fan', 0)} quạt"
            
            log_list.append({
                "Ngày": d_str,
                "Thứ": wd,
                "Người trực": assigned_to_val,
                "Việc nhà": chores_str,
                "Uống sữa": milk_str,
                "Quên tắt điện/quạt": elec_str,
                "Ghi chú": log_info.get("notes", "")
            })
            
        df_hist = pd.DataFrame(log_list)
        st.dataframe(df_hist, use_container_width=True)
        
        st.markdown("---")
        st.markdown("#### ✏️ Chỉnh sửa hoặc Xóa nhật ký ngày")
        
        edit_col1, edit_col2 = st.columns(2)
        
        with edit_col1:
            date_to_manage = st.selectbox(
                "Chọn ngày cần xử lý:",
                options=sorted(list(logs.keys()), reverse=True)
            )
            
        with edit_col2:
            action = st.radio(
                "Hành động:",
                options=["Xóa nhật ký ngày này", "Sửa nhật ký ngày này"]
            )
            
        if action == "Xóa nhật ký ngày này":
            confirm_delete = st.button("🚨 Xác nhận Xóa Vĩnh Viễn", type="primary")
            if confirm_delete:
                if date_to_manage in db["daily_logs"]:
                    del db["daily_logs"][date_to_manage]
                    if save_db(db):
                        st.success(f"Đã xóa thành công nhật ký ngày {date_to_manage}!")
                        st.rerun()
        else:
            st.info(f"💡 Để sửa nhật ký ngày **{date_to_manage}**, hãy chuyển sang tab **'Điểm danh hàng ngày'**, chọn ngày đó và thực hiện thay đổi rồi nhấn lưu.")

# ----------------------------------------------------
# TAB 5: SETTINGS & EXCEL
# ----------------------------------------------------
with tab_settings:
    st.markdown("### ⚙️ Bảng giá Excel & Cấu hình tiền lương")
    
    st.markdown("#### 💵 Đơn giá thưởng và phạt (Đồng bộ với file Excel)")
    
    # Load prices into a DataFrame for display and editing
    prices_list = [{"Công việc": k, "Đơn giá (VND)": v} for k, v in prices.items()]
    df_prices = pd.DataFrame(prices_list)
    
    st.info("💡 Nhấp đúp vào ô bất kỳ để sửa. Nhấn nút **'+ Thêm hàng'** ở cuối bảng để thêm công việc mới — điền tên công việc và đơn giá (số âm = phạt, số dương = thưởng), rồi bấm lưu.")
    
    edited_df = st.data_editor(
        df_prices,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Công việc": st.column_config.TextColumn(
                "Công việc",
                help="Tên công việc hoặc vi phạm",
                required=True,
            ),
            "Đơn giá (VND)": st.column_config.NumberColumn(
                "Đơn giá (VND)",
                help="Số dương = thưởng, số âm = phạt",
                format="%d đ",
                required=True,
            ),
        }
    )
    
    save_excel_btn = st.button("💾 Cập nhật bảng giá Excel (bang_gia.xlsx)")
    
    if save_excel_btn:
        # Convert edited dataframe back to dictionary, skip empty rows
        updated_prices = {}
        for _, row in edited_df.iterrows():
            ten_cv = str(row["Công việc"]).strip() if pd.notna(row["Công việc"]) else ""
            don_gia = row["Đơn giá (VND)"]
            if ten_cv and pd.notna(don_gia):
                try:
                    updated_prices[ten_cv] = int(don_gia)
                except:
                    pass
        
        if not updated_prices:
            st.error("❌ Bảng giá trống hoặc chưa điền đầy đủ, vui lòng kiểm tra lại!")
        elif save_prices_to_excel(updated_prices):
            st.success(f"🎉 Đã lưu {len(updated_prices)} mục vào file Excel 'bang_gia.xlsx' thành công!")
            st.rerun()
            
    st.markdown("---")
    st.markdown("#### 🏢 Thiết lập Lương cơ bản tháng (Lương cứng bắt đầu)")
    
    col_s1, col_s2 = st.columns(2)
    
    current_bond_base = db["settings"].get("base_salary", {}).get("Bond", 0)
    current_sushi_base = db["settings"].get("base_salary", {}).get("Sushi", 0)
    
    with col_s1:
        new_bond_base = st.number_input(
            "Lương cơ bản của anh hai Bond (đ):",
            min_value=0,
            value=int(current_bond_base),
            step=10000
        )
        
    with col_s2:
        new_sushi_base = st.number_input(
            "Lương cơ bản của em gái Sushi (đ):",
            min_value=0,
            value=int(current_sushi_base),
            step=10000
        )
        
    save_settings_btn = st.button("💾 Lưu Cài Đặt Lương Cơ Bản")
    if save_settings_btn:
        db["settings"]["base_salary"]["Bond"] = int(new_bond_base)
        db["settings"]["base_salary"]["Sushi"] = int(new_sushi_base)
        if save_db(db):
            st.success("🎉 Đã lưu cài đặt lương cơ bản thành công!")
            st.rerun()
