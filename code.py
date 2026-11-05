import os
import re
import difflib
import warnings
import pyodbc
import numpy as np
import pandas as pd
import geopandas as gpd
import statsmodels.api as sm
from sklearn.neighbors import NearestNeighbors
import scipy.stats as stats
import libpysal
import esda
import matplotlib as mpl
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.colors import ListedColormap
import matplotlib.patches as mpatches
from mpl_toolkits.axes_grid1.inset_locator import inset_axes
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

warnings.filterwarnings('ignore')
plt.rcParams['font.family'] = 'Tahoma'
plt.rcParams['axes.unicode_minus'] = False 


# =====================================================================
# بخش ۱: استخراج داده‌ها از دیتابیس (IWI، هزینه‌ها و درآمدها)
# =====================================================================
print("\n[۱/۸] در حال استخراج و محاسبه داده‌ها از پایگاه داده...")
db_path = r'F:\dgg\HB1403.mdb'
conn_str = rf'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};'
conn = pyodbc.connect(conn_str)

cols = """Address, DyCol03, DyCol06, DyCol07, DyCol08, DyCol09, DyCol10, DyCol12, 
          DyCol13, DyCol15, DyCol16, DyCol17, DyCol18, DyCol19, DyCol22, DyCol24, 
          DyCol26, DyCol30, DyCol31, DyCol33, DyCol35, DyCol41, DyCol42"""
df_iwi = pd.read_sql(f"SELECT {cols} FROM U1403P2 UNION ALL SELECT {cols} FROM R1403P2", conn)

def has_asset(series):
    s = series.astype(str).str.strip()
    return np.where((s == '') | (s == 'nan') | (s == 'None') | (s == 'NaN'), 0, 1)

df_iwi['Address'] = df_iwi['Address'].astype(str).str.strip()
df_iwi['province_code'] = df_iwi['Address'].str[1:3]
df_iwi['county_code']   = df_iwi['Address'].str[3:5]
df_iwi['full_geo_code'] = df_iwi['Address'].str[1:5]

df_iwi['iwi_tv']     = has_asset(df_iwi['DyCol12']) | has_asset(df_iwi['DyCol13'])
df_iwi['iwi_fridge'] = has_asset(df_iwi['DyCol17']) | has_asset(df_iwi['DyCol18']) | has_asset(df_iwi['DyCol19'])
df_iwi['iwi_phone']  = has_asset(df_iwi['DyCol16']) | has_asset(df_iwi['DyCol33'])
df_iwi['iwi_car']    = has_asset(df_iwi['DyCol07'])
df_iwi['iwi_bike']   = has_asset(df_iwi['DyCol09'])
df_iwi['iwi_cheap']  = has_asset(df_iwi['DyCol10']) | has_asset(df_iwi['DyCol24'])
df_iwi['iwi_exp']    = has_asset(df_iwi['DyCol08']) | has_asset(df_iwi['DyCol15']) | has_asset(df_iwi['DyCol22']) | has_asset(df_iwi['DyCol26']) | has_asset(df_iwi['DyCol41'])
df_iwi['iwi_elec']   = has_asset(df_iwi['DyCol31'])

rooms = pd.to_numeric(df_iwi['DyCol03'].astype(str).str.strip(), errors='coerce').fillna(1)
df_iwi['w_rooms'] = np.select([rooms <= 1, rooms == 2, rooms >= 3], [-3.699681, 0.384050, 3.445009], default=-3.699681)
material = df_iwi['DyCol06'].astype(str).str.strip()
df_iwi['w_floor'] = np.select([material.isin(['1', '3', '4']), material.isin(['2', '5', '8']), material.isin(['6', '7'])], [6.107428, 1.227531, -7.558471], default=1.227531)
sewage, bathroom = has_asset(df_iwi['DyCol42']), has_asset(df_iwi['DyCol35'])
df_iwi['w_toilet'] = np.select([sewage == 1, (sewage == 0) & (bathroom == 1)], [8.140637, -1.090393], default=-7.439841)
df_iwi['w_water'] = np.where(has_asset(df_iwi['DyCol30']) == 1, 7.952443, -2.302023)

df_iwi['iwi_exp'] = np.where((df_iwi['iwi_exp'] == 0) & (df_iwi['iwi_car'] == 1), 1, df_iwi['iwi_exp'])
df_iwi['iwi_cheap'] = np.where((df_iwi['iwi_cheap'] == 0) & ((df_iwi['iwi_tv'] == 1) | (df_iwi['iwi_fridge'] == 1) | (df_iwi['iwi_phone'] == 1) | (df_iwi['iwi_bike'] == 1) | (df_iwi['iwi_car'] == 1) | (df_iwi['iwi_exp'] == 1) | (df_iwi['w_toilet'] == 8.140637) | (df_iwi['w_floor'] == 6.107428)), 1, df_iwi['iwi_cheap'])

df_iwi['IWI_Score'] = (25.004470 + (df_iwi['iwi_tv'] * 8.612657) + (df_iwi['iwi_fridge'] * 8.429076) + (df_iwi['iwi_phone'] * 7.127699) + (df_iwi['iwi_car'] * 4.651382) + (df_iwi['iwi_bike'] * 1.846860) + (df_iwi['iwi_cheap'] * 4.118394) + (df_iwi['iwi_exp'] * 6.507283) + (df_iwi['iwi_elec'] * 8.056664) + df_iwi['w_rooms'] + df_iwi['w_floor'] + df_iwi['w_toilet'] + df_iwi['w_water']).round(1)
df_iwi_final = df_iwi[['Address', 'full_geo_code', 'IWI_Score']]

df_health = pd.read_sql("SELECT Address, DyCol01 AS item_code, DyCol03 AS cost FROM U1403P3S06 UNION ALL SELECT Address, DyCol01 AS item_code, DyCol03 AS cost FROM R1403P3S06", conn)
df_health['Address'] = df_health['Address'].astype(str).str.strip()
df_health['cost'] = pd.to_numeric(df_health['cost'], errors='coerce').fillna(0)
codes_G = ['062117', '062119', '062122', '062128', '062223', '062225', '062334', '062336', '062338', '062341', '062343', '062345', '062347']
codes_NG = ['062118', '062121', '062123', '062129', '062224', '062226', '062335', '062337', '062339', '062342', '062344', '062346', '062348']
df_health['sector'] = np.select([df_health['item_code'].isin(codes_G), df_health['item_code'].isin(codes_NG)], ['cost_public', 'cost_private'], default='Other')
hh_health_costs = df_health[df_health['sector'].isin(['cost_public', 'cost_private'])].groupby(['Address', 'sector'])['cost'].sum().unstack(fill_value=0).reset_index()
for col in ['cost_public', 'cost_private']: 
    if col not in hh_health_costs.columns: hh_health_costs[col] = 0
hh_health_costs['total_OOP'] = hh_health_costs['cost_public'] + hh_health_costs['cost_private']

df_w = pd.read_sql("SELECT Address, DyCol15 FROM U1403P4S01 UNION ALL SELECT Address, DyCol15 FROM R1403P4S01", conn)
df_w['inc_w'] = pd.to_numeric(df_w['DyCol15'], errors='coerce').fillna(0)
df_s = pd.read_sql("SELECT Address, DyCol15 FROM U1403P4S02 UNION ALL SELECT Address, DyCol15 FROM R1403P4S02", conn)
df_s['inc_s'] = pd.to_numeric(df_s['DyCol15'], errors='coerce').fillna(0)
df_m = pd.read_sql("SELECT Address, DyCol03, DyCol04, DyCol05, DyCol06, DyCol07, DyCol08 FROM U1403P4S03 UNION ALL SELECT Address, DyCol03, DyCol04, DyCol05, DyCol06, DyCol07, DyCol08 FROM R1403P4S03", conn)
cols = ['DyCol03', 'DyCol04', 'DyCol05', 'DyCol06', 'DyCol07', 'DyCol08']
df_m[cols] = df_m[cols].apply(pd.to_numeric, errors='coerce').fillna(0)
df_m['inc_m'] = df_m[cols].sum(axis=1)
df_y = pd.read_sql("SELECT Address, DyCol05 FROM U1403P4S04 UNION ALL SELECT Address, DyCol05 FROM R1403P4S04", conn)
df_y['inc_y'] = pd.to_numeric(df_y['DyCol05'], errors='coerce').fillna(0)
conn.close()

inc_df = df_w.groupby('Address')['inc_w'].sum().reset_index()\
    .merge(df_s.groupby('Address')['inc_s'].sum().reset_index(), on='Address', how='outer')\
    .merge(df_m.groupby('Address')['inc_m'].sum().reset_index(), on='Address', how='outer')\
    .merge(df_y.groupby('Address')['inc_y'].sum().reset_index(), on='Address', how='outer').fillna(0)
inc_df['Address'] = inc_df['Address'].astype(str).str.strip()
inc_df['total_income'] = (inc_df['inc_w'] + inc_df['inc_s'] + inc_df['inc_m'] + inc_df['inc_y']) / 12
inc_df['total_income'] = np.where(inc_df['total_income'] <= 0, np.nan, inc_df['total_income'])

df_master = df_iwi_final.merge(hh_health_costs, on='Address', how='left').merge(inc_df[['Address', 'total_income']], on='Address', how='left')
df_master[['cost_public', 'cost_private', 'total_OOP']] = df_master[['cost_public', 'cost_private', 'total_OOP']].fillna(0)

df_master['private_income_ratio'] = (df_master['cost_private'] / df_master['total_income']) * 100
df_master['total_income_ratio'] = (df_master['total_OOP'] / df_master['total_income']) * 100
df_master['is_che'] = np.where(df_master['total_income_ratio'] > 10, 1, 0)

# =====================================================================
# بخش ۲: اجرای مدل‌های علّی PSM
# =====================================================================
print("\n[۲/۸] در حال اجرای PSM و تولید گزارش Word با فواصل اطمینان 95%...")
word_doc = Document()
word_doc.add_heading('گزارش تحلیل علّی همسان‌سازی امتیاز گرایش (PSM) - سال ۱۴۰۳', 0)

def mean_confidence_interval(data, confidence=0.95):
    a = 1.0 * np.array(data)
    n = len(a)
    m, se = np.mean(a), stats.sem(a)
    h = se * stats.t.ppf((1 + confidence) / 2., n-1)
    return m, m-h, m+h

def calc_att_with_ci(treated_array, control_array, confidence=0.95):
    diffs = treated_array - control_array
    n = len(diffs)
    att = np.mean(diffs)
    se = stats.sem(diffs)
    ci = stats.t.interval(confidence, df=n-1, loc=att, scale=se)
    return att, ci

def run_psm_advanced_word(data, treatment_col, report_title, doc, exclude_healthy=False):
    df_model = data.copy()
    if exclude_healthy:
        df_model = df_model[df_model['total_OOP'] > 0]
        report_title += " (فقط خانوارهای درگیر بیماری)"
        
    df_model[treatment_col] = pd.to_numeric(df_model[treatment_col], errors='coerce')
    df_model = df_model.dropna(subset=[treatment_col, 'total_income', 'total_income_ratio', 'IWI_Score'])
    
    treated, control = df_model[df_model[treatment_col] == 1].copy(), df_model[df_model[treatment_col] == 0].copy()
    if len(treated) == 0 or len(control) == 0: return
        
    X = sm.add_constant(df_model[['IWI_Score']])
    y = df_model[treatment_col]
    logit_model = sm.Logit(y, X).fit(disp=0)
    df_model['pscore'] = logit_model.predict(X)
    
    treated, control = df_model[df_model[treatment_col] == 1].copy(), df_model[df_model[treatment_col] == 0].copy()
    nn = NearestNeighbors(n_neighbors=1, algorithm='ball_tree').fit(control[['pscore']].values)
    _, indices = nn.kneighbors(treated[['pscore']].values)
    matched_control = control.iloc[indices.flatten()].copy()
    
    m_t_cost, l_t_cost, u_t_cost = mean_confidence_interval(treated['cost_private'].values)
    m_c_cost, l_c_cost, u_c_cost = mean_confidence_interval(matched_control['cost_private'].values)
    att_cost, ci_cost = calc_att_with_ci(treated['cost_private'].values, matched_control['cost_private'].values)
    
    m_t_ratio, l_t_ratio, u_t_ratio = mean_confidence_interval(treated['total_income_ratio'].values)
    m_c_ratio, l_c_ratio, u_c_ratio = mean_confidence_interval(matched_control['total_income_ratio'].values)
    att_ratio, ci_ratio = calc_att_with_ci(treated['total_income_ratio'].values, matched_control['total_income_ratio'].values)
    
    doc.add_heading(report_title, level=2)
    doc.add_paragraph(f"حجم نمونه: {len(treated)} خانوار در گروه هدف و {len(matched_control)} خانوار در گروه پایه")
    doc.add_paragraph(f"شاخص ثروت (IWI) همسان‌سازی شده: گروه هدف {treated['IWI_Score'].mean():.1f} | گروه پایه {matched_control['IWI_Score'].mean():.1f}")
    doc.add_paragraph(f"میانگین درآمد ماهانه: گروه هدف {treated['total_income'].mean():,.0f} ریال | گروه پایه {matched_control['total_income'].mean():,.0f} ریال")
    
    doc.add_paragraph(f"پرداخت از جیب خصوصی (گروه هدف): میانگین {m_t_cost:,.0f} ریال [95% CI: {l_t_cost:,.0f} تا {u_t_cost:,.0f}]")
    doc.add_paragraph(f"پرداخت از جیب خصوصی (گروه پایه): میانگین {m_c_cost:,.0f} ریال [95% CI: {l_c_cost:,.0f} تا {u_c_cost:,.0f}]")
    
    doc.add_paragraph(f"نسبت کل هزینه به درآمد (گروه هدف): میانگین {m_t_ratio:.1f}٪ [95% CI: {l_t_ratio:.1f} تا {u_t_ratio:.1f}]")
    doc.add_paragraph(f"نسبت کل هزینه به درآمد (گروه پایه): میانگین {m_c_ratio:.1f}٪ [95% CI: {l_c_ratio:.1f} تا {u_c_ratio:.1f}]")
    
    p = doc.add_paragraph(f"نتیجه‌گیری اثر خالص (ATT): متغیر مداخله باعث شده گروه هدف به طور میانگین {att_cost:,.0f} ریال [95% CI: {ci_cost[0]:,.0f} تا {ci_cost[1]:,.0f}] بیشتر در بخش خصوصی هزینه کند. در مجموع، بار مالی درمان برای گروه هدف، {att_ratio:.1f} درصد [95% CI: {ci_ratio[0]:.1f} تا {ci_ratio[1]:.1f}] از درآمد آن‌ها را بیشتر می‌بلعد.")
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_paragraph("-" * 50)

province_centers = [
    '0001', '0105', '0207', '0303', '0401', '0502', '0603', '0707', '0808', '0916', 
    '1002', '1105', '1204', '1304', '1402', '1503', '1601', '1701', '1801', '1904', 
    '2002', '2105', '2202', '2301', '2401', '2501', '2603', '2705', '2802', '2901', '3001'
]
df_master['is_periphery'] = np.where(df_master['full_geo_code'].isin(province_centers), 0, 1)
df_master['is_rural'] = np.where(df_master['Address'].str[0] == '2', 1, 0)

for is_sick in [True, False]:
    run_psm_advanced_word(df_master, 'is_periphery', 'تاوان جغرافیایی - حاشیه در برابر مرکز', word_doc, is_sick)
    run_psm_advanced_word(df_master, 'is_rural', 'شکاف شهری-روستایی در کوچ درمانی', word_doc, is_sick)

word_doc.save(r'F:\dgg\maps\PSM_Causal_Analysis_Report.docx')
df_master.to_excel(r'F:\dgg\Master_Dataset_1403_Final.xlsx', index=False)



# =====================================================================
# بخش ۳: بارگذاری نقشه‌های جغرافیایی (اصلاح‌شده برای حذف نقاط و مرز آبی)
# =====================================================================
print("\n[۳/۸] در حال بارگذاری و برش نقشه‌های جغرافیایی...")
def normalize_persian_text(text):
    if pd.isna(text): return ""
    text = str(text).replace('شهرستان', '').replace('ویژه', '').replace('فرمانداری', '')
    text = text.replace('ي', 'ی').replace('ك', 'ک').replace('آ', 'ا').replace('أ', 'ا')
    return re.sub(r'\s+', ' ', text.replace('\u200c', ' ')).strip()

# ۱. بارگذاری لایه پایه (شهرستان‌ها)
gdf = gpd.read_file(r'F:\dgg\map.geojson')

# تابعی برای استخراج فقط پلی‌گون‌ها (حذف نقاط و خطوط مزاحم از فایل‌های OSM)
def extract_polygons(geometry):
    from shapely.geometry import Polygon, MultiPolygon, GeometryCollection
    if isinstance(geometry, (Polygon, MultiPolygon)):
        return geometry
    elif isinstance(geometry, GeometryCollection):
        # اگر ترکیبی از نقطه و پلی‌گون بود، فقط پلی‌گون‌هایش را بردار
        polys = [geom for geom in geometry.geoms if isinstance(geom, (Polygon, MultiPolygon))]
        if polys:
            from shapely.ops import unary_union
            return unary_union(polys)
    return None

# اعمال فیلتر روی لایه شهرستان‌ها
gdf['geometry'] = gdf['geometry'].apply(extract_polygons)
gdf = gdf.dropna(subset=['geometry'])
gdf = gdf.set_geometry('geometry')
gdf.geometry = gdf.geometry.make_valid()

land_path = r'F:\dgg\iran_land.geojson'
water_path = r'F:\dgg\water.geojson'

# کلیپ کردن لایه شهرستان‌ها با مرز خاکی
gdf_land = None
if os.path.exists(land_path): 
    gdf_land = gpd.read_file(land_path)
    gdf = gpd.clip(gdf, gdf_land.to_crs(gdf.crs))

if os.path.exists(water_path):
    gdf_water = gpd.read_file(water_path)
    gdf_water = gdf_water[gdf_water.geometry.type.isin(['Polygon', 'MultiPolygon'])]
    gdf = gpd.overlay(gdf, gdf_water.set_geometry(gdf_water.geometry.make_valid()), how='difference')

# ۲. بارگذاری لایه استان‌ها (با اعمال فیلتر نقاط و مرز آبی)
gdf_ostan = None
if os.path.exists(r'F:\dgg\ostan.geojson'):
    gdf_ostan = gpd.read_file(r'F:\dgg\ostan.geojson')
    # دور ریختن تمام نقاط و نگه داشتن فقط پلی‌گون‌ها
    gdf_ostan['geometry'] = gdf_ostan['geometry'].apply(extract_polygons)
    gdf_ostan = gdf_ostan.dropna(subset=['geometry'])
    gdf_ostan = gdf_ostan.set_geometry('geometry')
    gdf_ostan.geometry = gdf_ostan.geometry.make_valid()
    # کلیپ کردن لایه استان با مرز خاکی (حذف خطوط در دریا)
    if gdf_land is not None:
        gdf_ostan = gpd.clip(gdf_ostan, gdf_land.to_crs(gdf_ostan.crs))
        
# ۳. بارگذاری لایه مراکز استان‌ها (فقط نقاط)
gdf_capitals = None
if os.path.exists(r'F:\dgg\capitals.geojson'):
    gdf_capitals = gpd.read_file(r'F:\dgg\capitals.geojson')
    # فیلتر برای نگه‌داشتن فقط نقطه‌ها
    gdf_capitals = gdf_capitals[gdf_capitals.geometry.type == 'Point']

map_name_col = 'name' if 'name' in gdf.columns else ('name:fa' if 'name:fa' in gdf.columns else gdf.columns[0])
gdf['Map_Name_Cleaned'] = gdf[map_name_col].apply(normalize_persian_text)

bridge_file_path = r'F:\dgg\Map_Bridge_Final.xlsx'
bridge_df = pd.read_excel(bridge_file_path, dtype={'Final_Code': str})
bridge_df['Final_Code'] = bridge_df['Final_Code'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(4)
bridge_df = bridge_df[~bridge_df['Final_Code'].str.contains('nan|000n', case=False, na=False)]
gdf = gdf.merge(bridge_df[['Map_Name_Cleaned', 'Final_Code']], on='Map_Name_Cleaned', how='left')

county_data_all = df_master.groupby('full_geo_code').agg(IWI_Mean=('IWI_Score', 'mean')).reset_index()
df_sick = df_master[df_master['total_OOP'] > 0]
county_data_sick = df_sick.groupby('full_geo_code').agg(
    avg_private_m=('cost_private', lambda x: x.mean() / 1_000_000),
    avg_public_m=('cost_public', lambda x: x.mean() / 1_000_000),
    avg_ratio=('private_income_ratio', 'mean'),
    avg_total_ratio=('total_income_ratio', 'mean'),
    che_rate=('is_che', lambda x: x.mean() * 100),
    patient_count=('total_OOP', 'count')
).reset_index()

county_data = pd.merge(county_data_sick, county_data_all, on='full_geo_code', how='left')
gdf_final = gdf.merge(county_data, left_on='Final_Code', right_on='full_geo_code', how='left')
gdf_spatial = gdf_final.dropna(subset=['avg_private_m', 'avg_public_m', 'che_rate', 'avg_total_ratio', 'IWI_Mean']).copy()
gdf_spatial = gdf_spatial[gdf_spatial.geometry.notnull() & ~gdf_spatial.geometry.is_empty].reset_index(drop=True)



# =====================================================================
# بخش ۴: محاسبه Moran's I
# =====================================================================
print("\n[۴/۸] در حال محاسبه شاخص‌های اکونومتری فضایی (Moran's I)...")
w = libpysal.weights.KNN.from_dataframe(gdf_spatial, k=5)
w.transform = 'R'

variables_to_analyze = {
    'avg_private_m': ('Private_Cost', 'میانگین هزینه خصوصی (میلیون ریال)', 'YlOrRd'),
    'avg_public_m': ('Public_Cost', 'میانگین هزینه دولتی (میلیون ریال)', 'Blues'),
    'avg_ratio': ('Private_Ratio', 'نسبت هزینه خصوصی به درآمد (درصد)', 'YlOrRd'),
    'avg_total_ratio': ('Total_Ratio', 'نسبت کل هزینه‌های درمان به درآمد (درصد)', 'OrRd'),
    'che_rate': ('CHE_Rate', 'نرخ مواجهه با هزینه‌های کمرشکن سلامت (درصد خانوارها)', 'Reds'),
    'IWI_Mean': ('Wealth_IWI', 'میانگین شاخص ثروت بین‌المللی (IWI)', 'RdYlGn')
}

for var_col, (eng_name, fa_label, _) in variables_to_analyze.items():
    mi = esda.Moran(gdf_spatial[var_col], w)
    print(f"📌 {eng_name}: Moran I = {mi.I:.4f} (p={mi.p_sim:.4f})")

# =====================================================================
# بخش ۵: رسم نقشه‌های توزیع فضایی (با لایه‌های استانی)
# =====================================================================
print("\n[۵/۸] در حال رسم نقشه‌های توزیع و افزودن ۱۰ شهرستان برتر...")

for var_col, (eng_name, fa_label, cmap_choro) in variables_to_analyze.items():
    fig, ax = plt.subplots(1, 1, figsize=(14, 13))
    plt.subplots_adjust(bottom=0.2)
    
    gdf_final.plot(
        column=var_col, cmap=cmap_choro, linewidth=0.2, ax=ax, edgecolor='0.6',
        legend=True, legend_kwds={'label': fa_label, 'orientation': "horizontal", 'shrink': 0.5, 'pad': 0.02},
        missing_kwds={'color': 'whitesmoke'}
    )
    
    # اضافه کردن مرز استان‌ها و ستاره مراکز
    if gdf_ostan is not None:
        gdf_ostan.plot(ax=ax, facecolor="none", edgecolor="black", linewidth=0.8, alpha=0.7)
    if gdf_capitals is not None:
        gdf_capitals.plot(ax=ax, marker='*', color='gold', markersize=60, edgecolor='black', linewidth=0.5, zorder=5)
        
    ax.set_title(f'نقشه توزیع فضایی: {fa_label}', fontdict={'fontsize': 16, 'fontweight': 'bold'}, pad=15)
    ax.set_axis_off()
    
    if var_col == 'IWI_Mean':
        top10 = gdf_spatial.nsmallest(10, var_col)
        box_title = "۱۰ شهرستان با پایین‌ترین شاخص ثروت (محروم‌ترین‌ها):"
    else:
        top10 = gdf_spatial.nlargest(10, var_col)
        box_title = "۱۰ شهرستان با بالاترین مقدار (بحرانی‌ترین‌ها):"
        
    col1 = " | ".join([f"{i+1}. {row['Map_Name_Cleaned']} ({row[var_col]:.1f})" for i, row in enumerate(top10.iloc[:5].to_dict('records'))])
    col2 = " | ".join([f"{i+6}. {row['Map_Name_Cleaned']} ({row[var_col]:.1f})" for i, row in enumerate(top10.iloc[5:10].to_dict('records'))])
    footer_text = f"{box_title}\n{col1}\n{col2}"
    
    fig.text(0.5, 0.04, footer_text, ha='center', va='center', fontsize=11, fontweight='bold',
             bbox=dict(facecolor='whitesmoke', edgecolor='gray', boxstyle='round,pad=0.8', alpha=0.9))
    plt.savefig(rf'F:\dgg\maps\Map_Raw_{eng_name}.png', dpi=300, bbox_inches='tight')
    plt.close()


# =====================================================================
# بخش ۶: رسم نقشه‌های ارتباط‌سنجی دو متغیره  
# =====================================================================
print("\n[۶/۸] در حال رسم نقشه‌های ارتباط‌سنجی دو متغیره (Bivariate)...")

# اطمینان از وجود پوشه خروجی
output_dir = r'F:\dgg\maps'
os.makedirs(output_dir, exist_ok=True)

biv_colors = {
    '11': '#e8e8e8', '12': '#dfb0d6', '13': '#be64ac',
    '21': '#ace4e4', '22': '#a5add3', '23': '#8c62aa',
    '31': '#5ac8c8', '32': '#5698b9', '33': '#3b4994'
}

def create_biv_class(val_x, val_y, bx, by):
    if pd.isna(val_x) or pd.isna(val_y): return None
    cx = 1 if val_x <= bx[0] else (2 if val_x <= bx[1] else 3)
    cy = 1 if val_y <= by[0] else (2 if val_y <= by[1] else 3)
    return f"{cx}{cy}"

x_col = 'IWI_Mean'
breaks_x = [gdf_spatial[x_col].quantile(0.333), gdf_spatial[x_col].quantile(0.667)]

# تعریف متغیرها برای ذخیره‌سازی
y_cols_biv = {
    'avg_private_m': ('Private_Cost', 'هزینه درمان خصوصی'),
    'avg_total_ratio': ('Total_Ratio', 'نسبت کل هزینه‌ها به درآمد'),
    'che_rate': ('CHE_Rate', 'نرخ هزینه‌های کمرشکن سلامت')
}

for y_col, (eng_y, fa_y) in y_cols_biv.items():
    breaks_y = [gdf_spatial[y_col].quantile(0.333), gdf_spatial[y_col].quantile(0.667)]
    gdf_final['biv_class'] = gdf_final.apply(lambda r: create_biv_class(r[x_col], r[y_col], breaks_x, breaks_y), axis=1)
    
    fig, ax = plt.subplots(1, 1, figsize=(14, 12))
    gdf_final.plot(ax=ax, color='whitesmoke', edgecolor='0.8', linewidth=0.2)
    
    for class_code, color in biv_colors.items():
        subset = gdf_final[gdf_final['biv_class'] == class_code]
        if not subset.empty:
            subset.plot(ax=ax, color=color, linewidth=0.2, edgecolor='0.4')
            
    if gdf_ostan is not None:
        gdf_ostan.plot(ax=ax, facecolor="none", edgecolor="black", linewidth=0.8, alpha=0.7)
    if gdf_capitals is not None:
        gdf_capitals.plot(ax=ax, marker='*', color='gold', markersize=60, edgecolor='black', linewidth=0.5, zorder=5)
            
    ax.set_title(f'ارتباط‌سنجی فضایی: تقابل ثروت منطقه‌ای (IWI) با {fa_y}', fontdict={'fontsize': 16, 'fontweight': 'bold'})
    ax.set_axis_off()
    
    ax_leg = inset_axes(ax, width="25%", height="25%", loc='lower left', borderpad=2)
    ax_leg.set_aspect('equal')
    for i in range(3):
        for j in range(3):
            rect = plt.Rectangle((i, j), 1, 1, facecolor=biv_colors[f"{i+1}{j+1}"], edgecolor='white', linewidth=0.8)
            t = mpl.transforms.Affine2D().rotate_deg_around(1.5, 1.5, 45) + ax_leg.transData
            rect.set_transform(t)
            ax_leg.add_patch(rect)
            
    ax_leg.set_xlim(-1, 4)
    ax_leg.set_ylim(-1, 4)
    ax_leg.axis('off')
    
    ax_leg.text(1.5, 3.8, 'ثروت بالا\nهزینه بالا', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#3b4994')
    ax_leg.text(1.5, -0.8, 'ثروت کم\nهزینه کم', ha='center', va='top', fontsize=9, color='#555555')
    ax_leg.text(3.7, 1.5, 'ثروت بالا\nهزینه کم', ha='left', va='center', fontsize=9, fontweight='bold', color='#5ac8c8')
    ax_leg.text(-0.7, 1.5, 'ثروت کم\nهزینه بالا', ha='right', va='center', fontsize=9, fontweight='bold', color='#be64ac')
    
    clean_eng_y = re.sub(r'[^A-Za-z0-9_]', '', str(eng_y))
    
    safe_filename = f"Map_Bivariate_{clean_eng_y}_vs_IWI.png"
    save_path = os.path.join(output_dir, safe_filename)
    
    plt.savefig(save_path, dpi=300, bbox_inches='tight')
    plt.close()