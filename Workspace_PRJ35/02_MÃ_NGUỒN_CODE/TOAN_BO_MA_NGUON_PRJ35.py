# -*- coding: utf-8 -*-
# Toàn bộ mã PRJ-35: phân tích chính và đối chiếu B1.

# PHÂN TÍCH CHÍNH

# %%
import pandas as pd
import numpy as np
from pathlib import Path

PATH_CSV1_RAW = 'data/nguon_goc/26_2521001040_Triệu Hoàng Lan Anh_CSV1.csv'
PATH_CSV2_RAW = 'data/nguon_goc/26_2521001040_Triệu Hoàng Lan Anh_CSV2.csv'
PATH_CSV2_CLEAN = 'data/weather_clean_csv2.csv'
COT_DIEN = [
    'Global_active_power', 'Global_reactive_power', 'Voltage',
    'Global_intensity', 'Sub_metering_1', 'Sub_metering_2', 'Sub_metering_3'
]

WEATHER_START = "2006-12-16 00:00:00"
WEATHER_END = "2010-11-26 23:00:00"
WEATHER_NUMERIC_COLS = [
    "apparent_temperature_c",
    "relative_humidity_pct",
    "dew_point_c",
    "wind_speed_kmh",
    "precipitation_mm",
]

# %%
def tim_so_dong_metadata(file_path):
    """Tự động quét các dòng đầu của file CSV để tìm vị trí Header chính thức chứa cột thời gian."""
    with open(file_path, "r", encoding="utf-8-sig") as f:
        for idx, line in enumerate(f):
            cols = [c.strip().lower() for c in line.split(",")]
            if "time" in cols or "date" in cols:
                return idx
    return 0

SKIPROWS_CSV2 = tim_so_dong_metadata(PATH_CSV2_RAW)

# %%
def lam_sach_thoi_tiet(bang_tho):
    bang = bang_tho.rename(columns={
        'RELATIVE_HUMIDITY_2M (%)': 'relative_humidity_pct',
        'dew_point_2m (°C)': 'dew_point_c',
        'apparent_temperature (°C)': 'apparent_temperature_c',
        'wind_speed_10m (km/h)': 'wind_speed_kmh',
        'precipitation (mm)': 'precipitation_mm'
    }).copy()
    gio_tho = pd.to_datetime(bang['time'], format='%Y-%m-%dT%H:%M', errors='coerce')
    gio_du_kien = pd.Series(pd.date_range(WEATHER_START, WEATHER_END, freq='h'))
    assert len(bang) == len(gio_du_kien), 'Khác số dòng; dừng để kiểm tra thời gian.'
    doc_duoc = gio_tho.notna()
    assert (gio_tho[doc_duoc] == gio_du_kien[doc_duoc]).all(), 'Mốc đọc được lệch lưới giờ.'
    bang['datetime'] = gio_tho.fillna(gio_du_kien)
    co = pd.DataFrame({'datetime': bang['datetime'], 'gio_suy_ra': ~doc_duoc})
    for cot in WEATHER_NUMERIC_COLS:
        bang[cot] = pd.to_numeric(bang[cot], errors='coerce')
        co[cot + '_thieu_truoc_dien'] = bang[cot].isna()
    cot_lien_tuc = WEATHER_NUMERIC_COLS[:-1]
    bang[cot_lien_tuc] = bang[cot_lien_tuc].ffill()
    bang['precipitation_mm'] = bang['precipitation_mm'].fillna(0)
    bang['relative_humidity_pct'] = bang['relative_humidity_pct'].clip(0, 100)
    bang['wind_speed_kmh'] = bang['wind_speed_kmh'].clip(lower=0)
    bang['precipitation_mm'] = bang['precipitation_mm'].clip(lower=0)
    assert bang['datetime'].is_unique
    assert bang['datetime'].is_monotonic_increasing
    return bang[['datetime'] + WEATHER_NUMERIC_COLS], co

# %%
so_dong_raw = 0
thieu_raw = pd.Series(dtype='int64')
chuoi_loi_raw = []
cuc_tri_raw = []
ram_chunk_max = 0
for chunk in pd.read_csv(PATH_CSV1_RAW, sep=';', dtype=str, chunksize=100_000):
    so_dong_raw += len(chunk)
    thieu_raw = thieu_raw.add(chunk.isna().sum(), fill_value=0)
    ram_chunk_max = max(ram_chunk_max, chunk.memory_usage(deep=True).sum())
    so = pd.to_numeric(chunk['Global_active_power'], errors='coerce')
    chuoi_loi_raw.extend(chunk.loc[chunk['Global_active_power'].notna() & so.isna(), 'Global_active_power'].unique())
    cuc_tri_raw.extend(so.loc[(so < 0) | (so > 1000)].unique())
doc_chunk_31 = pd.DataFrame({'so_dong': [so_dong_raw], 'chunksize': [100_000],
              'ram_chunk_max_byte': [ram_chunk_max]})
del chunk
print(doc_chunk_31.to_string(index=False))

# %%
df_weather_raw = pd.read_csv(PATH_CSV2_RAW, skiprows=SKIPROWS_CSV2)
print('Điện:', so_dong_raw, 'dòng; thời tiết:', len(df_weather_raw), 'dòng')
print(pd.read_csv(PATH_CSV2_RAW, nrows=1).to_string(index=False))

# %%
power = pd.read_csv(r"data/power_clean_csv1.csv")
weather, co_thoi_tiet = lam_sach_thoi_tiet(df_weather_raw)
weather.to_csv(PATH_CSV2_CLEAN, index=False)
power_raw_time = pd.read_csv(
    r"data/nguon_goc/26_2521001040_Triệu Hoàng Lan Anh_CSV1.csv",
    sep=';', usecols=['Date', 'Time']
)
weather_raw_time = df_weather_raw[['time']]
power['datetime'] = pd.to_datetime(power['datetime'], errors='coerce')
weather['datetime'] = pd.to_datetime(weather['datetime'], errors='coerce')

power_raw_datetime = pd.to_datetime(
    power_raw_time['Date'] + ' ' + power_raw_time['Time'],
    format='%d/%m/%Y %H:%M:%S', errors='coerce'
)
weather_raw_datetime = pd.to_datetime(
    weather_raw_time['time'], format='%Y-%m-%dT%H:%M', errors='coerce'
)

power_bad = power_raw_datetime.isna()
weather_bad = weather_raw_datetime.isna()
power_shift = power_raw_datetime.notna() & (power_raw_datetime != power['datetime'])
weather_shift = weather_raw_datetime.notna() & (weather_raw_datetime != weather['datetime'])
assert len(power_raw_datetime) == len(power)
assert not power_shift.any() and not weather_shift.any()
assert power['datetime'].diff().dropna().eq(pd.Timedelta(minutes=1)).all()
assert weather['datetime'].diff().dropna().eq(pd.Timedelta(hours=1)).all()
print('Mốc không đọc được (điện, thời tiết):', power_bad.sum(), weather_bad.sum())
cong_suat_goc = pd.read_csv(PATH_CSV1_RAW, sep=';', usecols=['Global_active_power'],
                          dtype=str, keep_default_na=False)['Global_active_power']
cong_suat_so = pd.to_numeric(cong_suat_goc.str.replace(r'[^\d,.\-]', '', regex=True)
                            .str.replace(',', '.', regex=False), errors='coerce')
power['cong_suat_du_bao_kw'] = cong_suat_so.where(cong_suat_so.between(0, 11))
power['can_dien_du_bao'] = power['cong_suat_du_bao_kw'].isna()
power['moc_suy_ra'] = power_bad.to_numpy()
power['cong_suat_du_bao_kw'] = power.groupby(power['datetime'].dt.floor('D'))['cong_suat_du_bao_kw'].ffill()


# %%
loi_power = pd.DataFrame({
    'dong_csv': power_raw_time.index + 2,
    'Date_tho': power_raw_time['Date'],
    'Time_tho': power_raw_time['Time'],
    'datetime_sach': power['datetime'],
    'khong_doc_duoc': power_bad,
    'lech_nhan': power_shift
})
loi_power = loi_power[power_bad | power_shift]

loi_weather = pd.DataFrame({
    'dong_csv': weather_raw_time.index + 4,
    'time_tho': weather_raw_time['time'],
    'datetime_sach': weather['datetime'],
    'khong_doc_duoc': weather_bad,
    'lech_nhan': weather_shift
})
loi_weather = loi_weather[weather_bad | weather_shift]


power['hour_key'] = power['datetime'].dt.floor('h')
weather['hour_key'] = weather['datetime'].dt.floor('h')

assert power['datetime'].notna().all(), "Bảng điện có thời gian thiếu."
assert weather['datetime'].notna().all(), "Bảng thời tiết có thời gian thiếu."
assert (weather['datetime'] == weather['hour_key']).all(), "Thời tiết có mốc không đúng đầu giờ."

weather = weather.rename(columns={'datetime': 'weather_datetime'})
merged = pd.merge(
    power, weather, on='hour_key', how='left',
    validate='many_to_one', indicator=True
)

# %%
assert len(merged) == len(power) >= 100_000
assert merged[power.columns].equals(power)
assert merged['datetime'].is_unique
assert merged['_merge'].eq('both').all()
assert merged[WEATHER_NUMERIC_COLS].notna().all().all()
assert merged['datetime'].dt.floor('h').eq(merged['weather_datetime']).all()

merged = merged.drop(columns='_merge')
print('Bảng ghép:', merged.shape)

# %%
bang_b3 = merged

ram_cot_b3 = pd.DataFrame({
    'kieu_du_lieu': bang_b3.dtypes.astype(str),
    'bo_nho_byte': bang_b3.memory_usage(index=False, deep=True)
})
ram_cot_b3['bo_nho_MiB'] = ram_cot_b3['bo_nho_byte'] / 1024 ** 2


cot_so_b3 = bang_b3.select_dtypes(include='float64').columns
bang_float32_b3 = bang_b3.copy()
for cot in cot_so_b3:
    bang_float32_b3[cot] = bang_b3[cot].astype('float32')

# %%
sai_so_b3 = []
for cot in cot_so_b3:
    goc = bang_b3[cot]
    moi = bang_float32_b3[cot].astype('float64')
    lech = np.abs(moi - goc)
    sai_so_b3.append({
        'cot': cot,
        'sai_so_tuyet_doi_max': lech.max(),
        'sai_so_tuong_doi_max': (lech / goc.abs().replace(0, np.nan)).max(),
        'so_gia_tri_thay_doi': (lech > 0).sum(),
        'so_gia_tri_vuot_dung_sai': (~np.isclose(moi, goc, rtol=1e-6, atol=1e-8, equal_nan=True)).sum()
    })
sai_so_b3 = pd.DataFrame(sai_so_b3)

assert bang_float32_b3.shape == bang_b3.shape
assert bang_float32_b3.index.equals(bang_b3.index)
assert bang_float32_b3.isna().equals(bang_b3.isna())
cot_gio_b3 = ['datetime', 'hour_key', 'weather_datetime']
assert bang_float32_b3[cot_gio_b3].equals(bang_b3[cot_gio_b3])

assert (sai_so_b3['so_gia_tri_vuot_dung_sai'] == 0).all()

# %%
ram_truoc_b3 = bang_b3.memory_usage(deep=True).sum()
ram_sau_b3 = bang_float32_b3.memory_usage(deep=True).sum()
so_sanh_ram_b3 = pd.DataFrame({
    'phien_ban': ['float64', 'float32'],
    'bo_nho_byte': [ram_truoc_b3, ram_sau_b3]
})
so_sanh_ram_b3['bo_nho_MiB'] = so_sanh_ram_b3['bo_nho_byte'] / 1024 ** 2

print(so_sanh_ram_b3.to_string(index=False))

# %%
bang_float32_b3.to_parquet(
    r'data/merged_minute_float32_B3.parquet',
    engine='pyarrow', compression='snappy', index=False
)

bang_doc_lai_b3 = pd.read_parquet(
    r'data/merged_minute_float32_B3.parquet',
    engine='pyarrow'
)
kiem_tra_parquet_b3 = pd.DataFrame({
    'noi_dung': ['Số dòng và cột', 'Tên và thứ tự cột', 'Kiểu dữ liệu', 'Chỉ mục dòng', 'Toàn bộ giá trị'],
    'khop': [
        bang_doc_lai_b3.shape == bang_float32_b3.shape,
        bang_doc_lai_b3.columns.equals(bang_float32_b3.columns),
        bang_doc_lai_b3.dtypes.equals(bang_float32_b3.dtypes),
        bang_doc_lai_b3.index.equals(bang_float32_b3.index),
        bang_doc_lai_b3.equals(bang_float32_b3)
    ]
})
assert kiem_tra_parquet_b3['khop'].all()
print(kiem_tra_parquet_b3.to_string(index=False))


dung_luong_parquet_b3 = Path('data/merged_minute_float32_B3.parquet').stat().st_size / 1024 ** 2
print('Dung lượng bảng ghép Parquet (MiB):', dung_luong_parquet_b3)
del bang_doc_lai_b3, bang_float32_b3, bang_b3, merged

# %%
lich = pd.DataFrame({
    'thu_trong_tuan': power['datetime'].dt.dayofweek + 2,
    'thang': power['datetime'].dt.month
})
ram_lich_truoc = lich.memory_usage(index=False, deep=True)
for cot in lich.columns:
    lich[cot] = lich[cot].astype('category')
ram_lich_sau = lich.memory_usage(index=False, deep=True)
so_sanh_category = pd.DataFrame({'truoc_byte': ram_lich_truoc, 'category_byte': ram_lich_sau})
so_sanh_category['giam_pct'] = (1 - so_sanh_category['category_byte'] / so_sanh_category['truoc_byte']) * 100
del lich
print(so_sanh_category)

# %%
print(thieu_raw.astype('int64').to_frame('so_o_thieu_dien'))
print(df_weather_raw.isna().sum().to_frame('so_o_thieu_thoi_tiet'))

print('Chuỗi lỗi:', pd.Series(chuoi_loi_raw).unique()[:6])
print('Cực trị:', pd.Series(cuc_tri_raw).unique()[:4])

# %%
def dem_loi_tho(bang, cot_so, moc, dinh_dang):
    du_lieu = bang[cot_so]
    thieu = du_lieu.isna()
    khong_doc = du_lieu.apply(pd.to_numeric, errors='coerce').isna() & ~thieu
    loi_gio = pd.to_datetime(moc, format=dinh_dang, errors='coerce').isna()
    loi_so = thieu.any(axis=1) | khong_doc.any(axis=1)
    return pd.Series({
        'so_dong': len(bang), 'so_o_so': du_lieu.size,
        'o_thieu': thieu.sum().sum(), 'o_khong_doc': khong_doc.sum().sum(),
        'loi_thoi_gian': loi_gio.sum(),
        'dong_co_van_de': (loi_gio | loi_so).sum(),
        'chi_loi_gio': (loi_gio & ~loi_so).sum(),
        'chi_loi_so': (~loi_gio & loi_so).sum(),
        'loi_ca_hai': (loi_gio & loi_so).sum()
    })

cac_phan = []
for phan in pd.read_csv(PATH_CSV1_RAW, sep=';', dtype=str, chunksize=100_000):
    moc = phan['Date'] + ' ' + phan['Time']
    cac_phan.append(dem_loi_tho(phan, COT_DIEN, moc, '%d/%m/%Y %H:%M:%S'))
dien_tho = pd.DataFrame(cac_phan).sum()
thoi_tiet_tho = dem_loi_tho(df_weather_raw, df_weather_raw.columns.drop('time'),
                            df_weather_raw['time'], '%Y-%m-%dT%H:%M')
chat_luong_tho = pd.DataFrame([dien_tho, thoi_tiet_tho], index=['Điện', 'Thời tiết'])
chat_luong_tho.index.name = 'bang'
Path('dashboard_csv').mkdir(exist_ok=True)
chat_luong_tho.to_csv('dashboard_csv/chat_luong_ban_dau.csv', encoding='utf-8-sig')
print(chat_luong_tho.to_string())

# %%
print(co_thoi_tiet.iloc[:, 1:].sum().to_frame('so_dong_can_xu_ly'))

df_weather = weather.rename(columns={'weather_datetime': 'datetime'}).drop(columns='hour_key').copy()

print(df_weather[WEATHER_NUMERIC_COLS].agg(['count', 'min', 'mean', 'max']))

# %%
raw_b2 = pd.read_csv(
    r"data/nguon_goc/26_2521001040_Triệu Hoàng Lan Anh_CSV1.csv",
    sep=';', dtype=str, keep_default_na=False,
    usecols=COT_DIEN
)
time_b2 = power[['datetime']].copy()
assert len(raw_b2) == len(time_b2), "Khác số dòng, chưa thể đối chiếu theo vị trí."

nguong_b2 = {
    'Global_active_power': (0.0, 11.0),
    'Global_reactive_power': (0.0, 2.0),
    'Voltage': (210.0, 255.0),
    'Global_intensity': (0.0, 50.0),
    'Sub_metering_1': (0.0, 90.0),
    'Sub_metering_2': (0.0, 82.0),
    'Sub_metering_3': (0.0, 35.0)
}

# %%
thong_ke_b2 = []
co_power_b2 = time_b2.copy()

for col, (can_duoi, can_tren) in nguong_b2.items():
    chuoi = raw_b2[col].replace(['?', 'null', 'NULL', 'None', 'N/A', '-', ''], np.nan)
    chuoi = chuoi.astype(str).str.replace(r'[^\d,.\-]', '', regex=True)
    chuoi = chuoi.str.replace(',', '.', regex=False)
    so = pd.to_numeric(chuoi, errors='coerce')

    khong_doc_so = so.isna()
    ngoai_nguong = (so < can_duoi) | (so > can_tren)
    can_dien = khong_doc_so | ngoai_nguong
    thong_ke_b2.append({
        'cot': col,
        'so_dong': len(so),
        'khong_doc_so_sau_chuan_hoa': int(khong_doc_so.sum()),
        'ngoai_nguong_phuong_an': int(ngoai_nguong.sum()),
        'can_dien_theo_phuong_an': int(can_dien.sum()),
        'ty_le_can_dien_pct': can_dien.mean() * 100
    })

    if col == 'Global_active_power':
        co_power_b2['khong_doc_so'] = khong_doc_so
        co_power_b2['ngoai_nguong'] = ngoai_nguong
        co_power_b2['can_dien'] = can_dien

chat_luong_cot_b2 = pd.DataFrame(thong_ke_b2)
print(chat_luong_cot_b2.to_string(index=False))

# %%
chat_luong_ngay_b2 = co_power_b2.groupby(co_power_b2['datetime'].dt.floor('D')).agg(
    so_phut=('can_dien', 'size'),
    so_phut_khong_doc_so=('khong_doc_so', 'sum'),
    so_phut_ngoai_nguong=('ngoai_nguong', 'sum'),
    so_phut_can_dien=('can_dien', 'sum')
)
chat_luong_ngay_b2['ty_le_can_dien_pct'] = (
    chat_luong_ngay_b2['so_phut_can_dien'] / chat_luong_ngay_b2['so_phut'] * 100
)
chat_luong_ngay_b2['du_1440_phut'] = chat_luong_ngay_b2['so_phut'] == 1440
assert chat_luong_ngay_b2['so_phut'].sum() == len(raw_b2)
assert chat_luong_ngay_b2['so_phut_can_dien'].sum() == co_power_b2['can_dien'].sum()

print(chat_luong_ngay_b2['ty_le_can_dien_pct'].describe())

# %%
chat_luong_ngay_c1 = chat_luong_ngay_b2.rename_axis('ngay').reset_index()
ngay_du_c1 = chat_luong_ngay_c1['so_phut'] == 1440
so_sanh_nguong_c1 = []
for nguong in [0, 5, 10, 15, 20, 50, 100]:
    dat_nguong = ngay_du_c1 & (chat_luong_ngay_c1['ty_le_can_dien_pct'] <= nguong)
    so_sanh_nguong_c1.append({
        'nguong_can_dien_pct': nguong,
        'so_ngay_du_thoi_luong_dat_nguong': dat_nguong.sum(),
        'so_ngay_du_thoi_luong_vuot_nguong': (ngay_du_c1 & ~dat_nguong).sum()
    })
so_sanh_nguong_c1 = pd.DataFrame(so_sanh_nguong_c1)

print(so_sanh_nguong_c1.to_string(index=False))

# %%
dien_c2 = pd.read_parquet(
    r'data/merged_minute_float32_B3.parquet',
    columns=['datetime', 'Global_active_power'], engine='pyarrow'
)
dien_c2['Global_active_power'] = dien_c2['Global_active_power'].astype('float64')
chat_luong_c2 = chat_luong_ngay_c1.copy()

dien_c2['ngay'] = dien_c2['datetime'].dt.floor('D')
dien_ngay_c2 = dien_c2.groupby('ngay').agg(
    so_phut=('Global_active_power', 'size'),
    so_phut_co_cong_suat=('Global_active_power', 'count'),
    cong_suat_tb_kw=('Global_active_power', 'mean'),
    tong_cong_suat=('Global_active_power', 'sum')
).reset_index()
dien_ngay_c2['dien_nang_trong_khoang_du_lieu_kwh'] = dien_ngay_c2['tong_cong_suat'] / 60
dien_ngay_c2 = dien_ngay_c2.drop(columns='tong_cong_suat')
assert dien_ngay_c2['so_phut'].sum() == len(dien_c2)
assert (dien_ngay_c2['so_phut'] == dien_ngay_c2['so_phut_co_cong_suat']).all()
assert dien_ngay_c2['so_phut'].equals(chat_luong_c2['so_phut'])

# %%
dien_ngay_c2 = dien_ngay_c2.merge(
    chat_luong_c2.drop(columns='so_phut'), on='ngay', how='left', validate='one_to_one'
)
assert dien_ngay_c2['ty_le_can_dien_pct'].notna().all()
dien_ngay_c2['dat_nguong_20pct'] = dien_ngay_c2['du_1440_phut'] & (dien_ngay_c2['ty_le_can_dien_pct'] <= 20)
dien_ngay_c2['dat_nguong_15pct'] = dien_ngay_c2['du_1440_phut'] & (dien_ngay_c2['ty_le_can_dien_pct'] <= 15)
dien_ngay_c2['dien_nang_ngay_kwh'] = dien_ngay_c2['dien_nang_trong_khoang_du_lieu_kwh']
dien_ngay_c2.loc[~dien_ngay_c2['du_1440_phut'], 'dien_nang_ngay_kwh'] = np.nan
ngay_du_c2 = dien_ngay_c2['du_1440_phut']
assert np.isclose(
    dien_ngay_c2.loc[ngay_du_c2, 'dien_nang_ngay_kwh'],
    dien_ngay_c2.loc[ngay_du_c2, 'cong_suat_tb_kw'] * 24
).all()

print(dien_ngay_c2[['dat_nguong_20pct', 'dat_nguong_15pct']].sum())

# %%
dien_c3 = dien_c2[['datetime', 'Global_active_power']].merge(
    co_power_b2[['datetime', 'can_dien']],
    on='datetime', how='left', validate='one_to_one'
)
assert dien_c3['can_dien'].notna().all()

dien_c3['gio'] = dien_c3['datetime'].dt.floor('h')
dien_gio_c3 = dien_c3.groupby('gio').agg(
    so_phut=('Global_active_power', 'size'),
    so_phut_co_cong_suat=('Global_active_power', 'count'),
    cong_suat_tb_kw=('Global_active_power', 'mean'),
    tong_cong_suat=('Global_active_power', 'sum'),
    so_phut_can_dien=('can_dien', 'sum')
).reset_index()
dien_gio_c3['dien_nang_trong_khoang_du_lieu_kwh'] = dien_gio_c3['tong_cong_suat'] / 60
dien_gio_c3 = dien_gio_c3.drop(columns='tong_cong_suat')
dien_gio_c3['du_60_phut'] = dien_gio_c3['so_phut'] == 60
dien_gio_c3['ty_le_can_dien_pct'] = dien_gio_c3['so_phut_can_dien'] / dien_gio_c3['so_phut'] * 100
dien_gio_c3['dien_nang_gio_kwh'] = dien_gio_c3['dien_nang_trong_khoang_du_lieu_kwh']
dien_gio_c3.loc[~dien_gio_c3['du_60_phut'], 'dien_nang_gio_kwh'] = np.nan
assert dien_gio_c3['so_phut'].sum() == len(dien_c3)
assert (dien_gio_c3['so_phut'] == dien_gio_c3['so_phut_co_cong_suat']).all()
assert np.isclose(
    dien_gio_c3.loc[dien_gio_c3['du_60_phut'], 'dien_nang_gio_kwh'],
    dien_gio_c3.loc[dien_gio_c3['du_60_phut'], 'cong_suat_tb_kw']
).all()

# %%
dien_ngay_c3 = dien_ngay_c2.copy()
dien_gio_c3['ngay'] = dien_gio_c3['gio'].dt.floor('D')
ngay_tu_gio_c3 = dien_gio_c3.groupby('ngay').agg(
    so_phut_tu_gio=('so_phut', 'sum'),
    so_phut_can_dien_tu_gio=('so_phut_can_dien', 'sum'),
    dien_nang_tu_gio_kwh=('dien_nang_trong_khoang_du_lieu_kwh', 'sum')
).reset_index()
doi_chieu_c3 = ngay_tu_gio_c3.merge(
    dien_ngay_c3[['ngay', 'so_phut', 'so_phut_can_dien', 'dien_nang_trong_khoang_du_lieu_kwh']],
    on='ngay', how='left', validate='one_to_one'
)
doi_chieu_c3['lech_dien_nang_kwh'] = doi_chieu_c3['dien_nang_tu_gio_kwh'] - doi_chieu_c3['dien_nang_trong_khoang_du_lieu_kwh']
assert len(doi_chieu_c3) == len(dien_ngay_c3)
assert (doi_chieu_c3['so_phut_tu_gio'] == doi_chieu_c3['so_phut']).all()
assert (doi_chieu_c3['so_phut_can_dien_tu_gio'] == doi_chieu_c3['so_phut_can_dien']).all()
assert np.isclose(doi_chieu_c3['lech_dien_nang_kwh'], 0, rtol=0, atol=1e-10).all()

print('Số giờ:', len(dien_gio_c3), '| Đủ 60 phút:', dien_gio_c3['du_60_phut'].sum())

# %%
thoi_tiet_c4 = df_weather.copy()
dien_gio_c4 = dien_gio_c3.copy()
dien_ngay_c4 = dien_ngay_c2.copy()

thoi_tiet_c4['ngay'] = thoi_tiet_c4['datetime'].dt.floor('D')
thoi_tiet_ngay_c4 = thoi_tiet_c4.groupby('ngay').agg(
    so_gio_thoi_tiet=('datetime', 'size'),
    nhiet_do_cam_nhan_tb_c=('apparent_temperature_c', 'mean'),
    do_am_tb_pct=('relative_humidity_pct', 'mean'),
    diem_suong_tb_c=('dew_point_c', 'mean'),
    toc_do_gio_tb_kmh=('wind_speed_kmh', 'mean'),
    tong_mua_mm=('precipitation_mm', 'sum')
).reset_index()
so_gio_co_so_c4 = thoi_tiet_c4.groupby('ngay')[WEATHER_NUMERIC_COLS].count()
thoi_tiet_ngay_c4['du_24_gio_thoi_tiet'] = so_gio_co_so_c4.eq(24).all(axis=1).to_numpy()
assert thoi_tiet_ngay_c4['so_gio_thoi_tiet'].sum() == len(thoi_tiet_c4)
assert np.isclose(thoi_tiet_ngay_c4['tong_mua_mm'].sum(), thoi_tiet_c4['precipitation_mm'].sum())

# %%
dien_thoi_tiet_gio_c4 = dien_gio_c4.merge(
    thoi_tiet_c4.drop(columns='ngay').rename(columns={'datetime': 'gio'}),
    on='gio', how='left', validate='one_to_one', indicator=True
)
dien_thoi_tiet_ngay_c4 = dien_ngay_c4.merge(
    thoi_tiet_ngay_c4, on='ngay', how='left', validate='one_to_one', indicator=True
)
assert (dien_thoi_tiet_gio_c4['_merge'] == 'both').all()
assert (dien_thoi_tiet_ngay_c4['_merge'] == 'both').all()
assert dien_thoi_tiet_gio_c4[dien_gio_c4.columns].equals(dien_gio_c4)
assert dien_thoi_tiet_ngay_c4[dien_ngay_c4.columns].equals(dien_ngay_c4)
dien_thoi_tiet_gio_c4 = dien_thoi_tiet_gio_c4.drop(columns='_merge')
dien_thoi_tiet_ngay_c4 = dien_thoi_tiet_ngay_c4.drop(columns='_merge')

print('Bảng giờ:', dien_thoi_tiet_gio_c4.shape, '| Bảng ngày:', dien_thoi_tiet_ngay_c4.shape)

# %%
ngay_d1 = dien_thoi_tiet_ngay_c4.copy()
nhom20_d1 = ngay_d1.loc[ngay_d1['dat_nguong_20pct']].copy()
nhom15_d1 = ngay_d1.loc[ngay_d1['dat_nguong_15pct']].copy()
cot_d1 = ['cong_suat_tb_kw', 'dien_nang_ngay_kwh']
assert nhom20_d1[cot_d1].notna().all().all()
assert nhom15_d1[cot_d1].notna().all().all()

thong_ke20_d1 = nhom20_d1[cot_d1].describe()
thong_ke15_d1 = nhom15_d1[cot_d1].describe()
so_sanh_d1 = pd.DataFrame({
    'cong_suat_20pct_kw': thong_ke20_d1['cong_suat_tb_kw'],
    'cong_suat_15pct_kw': thong_ke15_d1['cong_suat_tb_kw'],
    'dien_nang_20pct_kwh': thong_ke20_d1['dien_nang_ngay_kwh'],
    'dien_nang_15pct_kwh': thong_ke15_d1['dien_nang_ngay_kwh']
})
print(so_sanh_d1)

# %%
cot_xem_d1 = ['ngay', 'cong_suat_tb_kw', 'dien_nang_ngay_kwh', 'ty_le_can_dien_pct', 'dat_nguong_15pct']
top10_d1 = nhom20_d1.sort_values('dien_nang_ngay_kwh', ascending=False).head(10)[cot_xem_d1]
print(top10_d1.to_string(index=False))

# %%
ngay_d2 = ngay_d1.copy()
gio_d2 = dien_thoi_tiet_gio_c4.copy()
gio_d2 = gio_d2.merge(
    ngay_d2[['ngay', 'dat_nguong_20pct', 'dat_nguong_15pct']],
    on='ngay', how='left', validate='many_to_one'
)
assert gio_d2[['dat_nguong_20pct', 'dat_nguong_15pct']].notna().all().all()
gio_d2['gio_trong_ngay'] = gio_d2['gio'].dt.hour
ngay_d2['thu_trong_tuan'] = ngay_d2['ngay'].dt.dayofweek + 2
ngay_d2['thang'] = ngay_d2['ngay'].dt.month

for cot in ['thu_trong_tuan', 'thang']:
    ngay_d2[cot] = ngay_d2[cot].astype('category')

# %%
cac_bang_gio_d2 = []
cac_bang_thu_d2 = []
cac_bang_thang_d2 = []
for nguong in [20, 15]:
    cot_co = f'dat_nguong_{nguong}pct'
    chon_ngay = ngay_d2.loc[ngay_d2[cot_co]]
    chon_gio = gio_d2.loc[gio_d2[cot_co] & gio_d2['du_60_phut']]
    assert len(chon_gio) == len(chon_ngay) * 24

    bang_gio = chon_gio.groupby('gio_trong_ngay').agg(
        so_gio=('cong_suat_tb_kw', 'size'),
        cong_suat_tb_kw=('cong_suat_tb_kw', 'mean'),
        dien_nang_tb_kwh=('dien_nang_gio_kwh', 'mean')
    ).reset_index()
    bang_thu = chon_ngay.groupby('thu_trong_tuan').agg(
        so_ngay=('cong_suat_tb_kw', 'size'),
        cong_suat_tb_kw=('cong_suat_tb_kw', 'mean'),
        dien_nang_tb_kwh=('dien_nang_ngay_kwh', 'mean')
    ).reset_index()
    bang_thang = chon_ngay.groupby('thang').agg(
        so_ngay=('cong_suat_tb_kw', 'size'),
        cong_suat_tb_kw=('cong_suat_tb_kw', 'mean'),
        dien_nang_tb_kwh=('dien_nang_ngay_kwh', 'mean')
    ).reset_index()
    bang_gio['nguong_can_dien_pct'] = nguong
    bang_thu['nguong_can_dien_pct'] = nguong
    bang_thang['nguong_can_dien_pct'] = nguong
    cac_bang_gio_d2.append(bang_gio)
    cac_bang_thu_d2.append(bang_thu)
    cac_bang_thang_d2.append(bang_thang)

chu_ky_gio_d2 = pd.concat(cac_bang_gio_d2, ignore_index=True)
chu_ky_thu_d2 = pd.concat(cac_bang_thu_d2, ignore_index=True)
chu_ky_thang_d2 = pd.concat(cac_bang_thang_d2, ignore_index=True)

print(chu_ky_gio_d2.to_string(index=False))
print(chu_ky_thu_d2.to_string(index=False))
print(chu_ky_thang_d2.to_string(index=False))

# %%
cot_d3 = ['Global_active_power', 'Sub_metering_1', 'Sub_metering_2', 'Sub_metering_3']
dien_d3 = pd.read_parquet(
    r'data/merged_minute_float32_B3.parquet',
    columns=['datetime'] + cot_d3, engine='pyarrow'
)
for cot in cot_d3:
    dien_d3[cot] = dien_d3[cot].astype('float64')
assert dien_d3[cot_d3].notna().all().all()
dien_d3['tong_kwh'] = dien_d3['Global_active_power'] / 60
dien_d3['nhom1_kwh'] = dien_d3['Sub_metering_1'] / 1000
dien_d3['nhom2_kwh'] = dien_d3['Sub_metering_2'] / 1000
dien_d3['nhom3_kwh'] = dien_d3['Sub_metering_3'] / 1000
dien_d3['chenh_lech_kwh'] = dien_d3['tong_kwh'] - dien_d3['nhom1_kwh'] - dien_d3['nhom2_kwh'] - dien_d3['nhom3_kwh']
dien_d3['chenh_lech_am'] = dien_d3['chenh_lech_kwh'] < -1e-8
dien_d3['ngay'] = dien_d3['datetime'].dt.floor('D')

# %%
cot_nang_luong_d3 = ['tong_kwh', 'nhom1_kwh', 'nhom2_kwh', 'nhom3_kwh', 'chenh_lech_kwh']
co_cau_ngay_d3 = dien_d3.groupby('ngay')[cot_nang_luong_d3].sum()
co_cau_ngay_d3['so_phut_chenh_lech_am'] = dien_d3.groupby('ngay')['chenh_lech_am'].sum()
co_cau_ngay_d3 = co_cau_ngay_d3.reset_index()
ngay_d3 = dien_ngay_c2.copy()
co_cau_ngay_d3 = co_cau_ngay_d3.merge(
    ngay_d3[['ngay', 'du_1440_phut', 'dat_nguong_20pct', 'dat_nguong_15pct', 'dien_nang_trong_khoang_du_lieu_kwh']],
    on='ngay', how='left', validate='one_to_one'
)
assert co_cau_ngay_d3['du_1440_phut'].notna().all()
assert np.isclose(co_cau_ngay_d3['tong_kwh'], co_cau_ngay_d3['dien_nang_trong_khoang_du_lieu_kwh'], rtol=0, atol=1e-10).all()

# %%
cac_bang_d3 = []
for nguong in [20, 15]:
    chon = co_cau_ngay_d3.loc[co_cau_ngay_d3[f'dat_nguong_{nguong}pct']]
    tong = chon[cot_nang_luong_d3].sum()
    bang = pd.DataFrame({
        'thanh_phan': cot_nang_luong_d3[1:],
        'dien_nang_kwh': tong[cot_nang_luong_d3[1:]].to_numpy()
    })
    bang['ty_trong_pct'] = bang['dien_nang_kwh'] / tong['tong_kwh'] * 100
    bang['nguong_cong_suat_pct'] = nguong
    bang['so_ngay'] = len(chon)
    assert np.isclose(bang['dien_nang_kwh'].sum(), tong['tong_kwh'])
    cac_bang_d3.append(bang)
co_cau_d3 = pd.concat(cac_bang_d3, ignore_index=True)

print(co_cau_d3.to_string(index=False))
print('Phút chênh lệch âm:', co_cau_ngay_d3['so_phut_chenh_lech_am'].sum())

# %%
ngay_d4 = ngay_d1.copy()
cot_thoi_tiet_d4 = ['nhiet_do_cam_nhan_tb_c', 'do_am_tb_pct', 'diem_suong_tb_c', 'toc_do_gio_tb_kmh', 'tong_mua_mm']
cot_d4 = ['cong_suat_tb_kw'] + cot_thoi_tiet_d4

cac_bang_d4 = []
for nguong in [20, 15]:
    chon = ngay_d4.loc[ngay_d4[f'dat_nguong_{nguong}pct'] & ngay_d4['du_24_gio_thoi_tiet'], cot_d4]
    assert chon.notna().all().all()
    ma_tran = chon.corr(method='pearson')
    bang = pd.DataFrame({
        'bien_thoi_tiet': cot_thoi_tiet_d4,
        'r_pearson': ma_tran.loc[cot_thoi_tiet_d4, 'cong_suat_tb_kw'].to_numpy()
    })
    bang['nguong_can_dien_pct'] = nguong
    bang['so_cap_ngay'] = len(chon)
    assert bang['r_pearson'].notna().all()
    assert (bang['r_pearson'].abs() <= 1).all()
    cac_bang_d4.append(bang)
tuong_quan_d4 = pd.concat(cac_bang_d4, ignore_index=True)

print(tuong_quan_d4.to_string(index=False))

# %%
thu_muc_csv = Path('dashboard_csv')
thu_muc_csv.mkdir(exist_ok=True)

def xuat_csv(bang, ten):
    bang.to_csv(thu_muc_csv / (ten + '.csv'), index=False, encoding='utf-8-sig')

# %%
cot_ngay = ['ngay', 'du_1440_phut', 'dat_nguong_20pct', 'dat_nguong_15pct',
            'dien_nang_ngay_kwh', 'thu_trong_tuan', 'cong_suat_tb_kw', 'nhiet_do_cam_nhan_tb_c']
xuat_csv(ngay_d2[cot_ngay], 'dien_ngay')
xuat_csv(chu_ky_gio_d2, 'chu_ky_gio')
xuat_csv(chu_ky_thu_d2, 'chu_ky_thu')
xuat_csv(chu_ky_thang_d2, 'chu_ky_thang')

# %%
xuat_csv(tuong_quan_d4, 'thoi_tiet')
xuat_csv(co_cau_d3, 'co_cau')

# %%
cac_ma_tran = []
for nguong in [20, 15]:
    chon = ngay_d4[f'dat_nguong_{nguong}pct'] & ngay_d4['du_24_gio_thoi_tiet']
    bang = ngay_d4.loc[chon, cot_d4].corr().rename_axis('bien').reset_index()
    bang['nguong_pct'] = nguong
    cac_ma_tran.append(bang)
xuat_csv(pd.concat(cac_ma_tran), 'ma_tran')

# %%
ngay_36 = ngay_d1[['ngay', 'dien_nang_ngay_kwh', 'dat_nguong_20pct']].copy()
ngay_36['dau_tuan'] = ngay_36['ngay'] - pd.to_timedelta(ngay_36['ngay'].dt.dayofweek, unit='D')
ngay_36['cuoi_tuan'] = ngay_36['ngay'].dt.dayofweek >= 5
ngay_36 = ngay_36.loc[ngay_36['dat_nguong_20pct']]

thuong_36 = ngay_36.loc[~ngay_36['cuoi_tuan']].groupby('dau_tuan').agg(
    so_ngay_thuong=('dien_nang_ngay_kwh', 'count'),
    tb_ngay_thuong_kwh=('dien_nang_ngay_kwh', 'mean')
)
cuoi_36 = ngay_36.loc[ngay_36['cuoi_tuan']].groupby('dau_tuan').agg(
    so_ngay_cuoi_tuan=('dien_nang_ngay_kwh', 'count'),
    tb_cuoi_tuan_kwh=('dien_nang_ngay_kwh', 'mean')
)
tuan_36 = thuong_36.merge(cuoi_36, on='dau_tuan', how='outer', validate='one_to_one')
tuan_36['du_7_ngay'] = tuan_36['so_ngay_thuong'].eq(5) & tuan_36['so_ngay_cuoi_tuan'].eq(2)

# %%
cap_tuan_36 = tuan_36.loc[tuan_36['du_7_ngay']].copy()
cap_tuan_36['chenh_lech_kwh_ngay'] = cap_tuan_36['tb_cuoi_tuan_kwh'] - cap_tuan_36['tb_ngay_thuong_kwh']
assert cap_tuan_36[['tb_ngay_thuong_kwh', 'tb_cuoi_tuan_kwh']].notna().all().all()
print('Số tuần đủ 7 ngày đạt ngưỡng 20%:', len(cap_tuan_36))
print(cap_tuan_36[['tb_ngay_thuong_kwh', 'tb_cuoi_tuan_kwh', 'chenh_lech_kwh_ngay']].describe())

# %%
from scipy import stats

d_36 = cap_tuan_36['chenh_lech_kwh_ngay'].sort_index().asfreq('W-MON')
n_36 = d_36.count()
tb_36 = d_36.mean()
print('Tuần trong lịch:', len(d_36), '| Tuần được chọn:', n_36)
print('Tuần bị loại:', d_36.isna().sum())
print(pd.Series({k: d_36.corr(d_36.shift(k)) for k in [1, 2, 3, 4]},
                name='Tương quan theo độ trễ tuần'))

# %%
# Tuần bị loại đóng góp 0 vào tổng sai lệch, không điền điện năng bằng 0.
e_36 = (d_36 - tb_36).fillna(0).to_numpy()
ket_qua_36 = []
for L in [4, 8, 12]:
    tong_36 = np.sum(e_36 ** 2)
    for k in range(1, L + 1):
        trong_so = 1 - k / (L + 1)
        tong_36 += 2 * trong_so * np.sum(e_36[k:] * e_36[:-k])
    se_36 = np.sqrt(tong_36) / n_36
    z_36 = tb_36 / se_36
    bien_36 = stats.norm.ppf(0.975) * se_36
    ket_qua_36.append([L, n_36, tb_36, se_36, z_36,
                       2 * stats.norm.sf(abs(z_36)), tb_36 - bien_36, tb_36 + bien_36])
hac_36 = pd.DataFrame(ket_qua_36, columns=[
    'tre_tuan', 'so_tuan', 'chenh_lech_tb', 'sai_so_chuan',
    'z', 'p_value', 'can_duoi_95', 'can_tren_95'])
print(hac_36.to_string(index=False))

# %%
t_36 = stats.ttest_rel(cap_tuan_36['tb_cuoi_tuan_kwh'],
                       cap_tuan_36['tb_ngay_thuong_kwh'], nan_policy='raise')
ci_36 = t_36.confidence_interval(confidence_level=0.95)
ttest_36 = pd.DataFrame({'so_tuan': [n_36], 'chenh_lech_tb': [tb_36],
    't': [t_36.statistic], 'df': [t_36.df], 'p_value': [t_36.pvalue],
    'can_duoi_95': [ci_36.low], 'can_tren_95': [ci_36.high]})
print(ttest_36.to_string(index=False))

# %%
dien_41 = pd.read_parquet(
    r'data/merged_minute_float32_B3.parquet',
    columns=['datetime', 'cong_suat_du_bao_kw', 'can_dien_du_bao', 'moc_suy_ra',
             'apparent_temperature_c', 'relative_humidity_pct', 'dew_point_c',
             'wind_speed_kmh', 'precipitation_mm']
)
dien_41['ngay'] = dien_41['datetime'].dt.floor('D')
print(len(dien_41), 'dòng;', dien_41['ngay'].nunique(), 'ngày')

# %%
ngay_41 = dien_41.groupby('ngay').agg(
    so_phut=('datetime', 'size'),
    so_phut_co_so=('cong_suat_du_bao_kw', 'count'),
    so_phut_can_dien=('can_dien_du_bao', 'sum'),
    so_moc_suy_ra=('moc_suy_ra', 'sum'),
    cong_suat_tb_kw=('cong_suat_du_bao_kw', 'mean'),
    nhiet_do_cam_nhan_tb_c=('apparent_temperature_c', 'mean'),
    do_am_tb_pct=('relative_humidity_pct', 'mean'),
    diem_suong_tb_c=('dew_point_c', 'mean'),
    toc_do_gio_tb_kmh=('wind_speed_kmh', 'mean')
).asfreq('D')
mua_gio_41 = dien_41[['datetime', 'precipitation_mm']].copy()
mua_gio_41['gio'] = mua_gio_41['datetime'].dt.floor('h')
mua_gio_41 = mua_gio_41.drop_duplicates('gio')
mua_ngay_41 = mua_gio_41.groupby(mua_gio_41['gio'].dt.floor('D'))['precipitation_mm'].agg(['sum', 'count'])
ngay_41['tong_mua_mm'] = mua_ngay_41['sum'].where(mua_ngay_41['count'].eq(24))
ngay_41['ty_le_can_dien_pct'] = ngay_41['so_phut_can_dien'] / ngay_41['so_phut'] * 100
ngay_41['dat_quy_tac_thu'] = (ngay_41['so_phut'].eq(1440)
    & ngay_41['so_phut_co_so'].eq(1440) & ngay_41['ty_le_can_dien_pct'].le(20))
ngay_41['muc_tieu_kw'] = ngay_41['cong_suat_tb_kw'].where(ngay_41['dat_quy_tac_thu'])
ngay_41['muc_tieu_kwh'] = ngay_41['muc_tieu_kw'] * 24
print(ngay_41[['dat_quy_tac_thu', 'so_moc_suy_ra']].sum())

# %%
dac_trung_41 = ngay_41[['muc_tieu_kw', 'muc_tieu_kwh', 'ty_le_can_dien_pct',
                        'so_moc_suy_ra']].copy()
dac_trung_41['dien_tre_1_kw'] = ngay_41['muc_tieu_kw'].shift(1)
dac_trung_41['dien_tre_7_kw'] = ngay_41['muc_tieu_kw'].shift(7)
dac_trung_41['tb_7_ngay_truoc_kw'] = ngay_41['muc_tieu_kw'].shift(1).rolling(7, min_periods=7).mean()
dac_trung_41['nhiet_do_truoc_1'] = ngay_41['nhiet_do_cam_nhan_tb_c'].shift(1)
dac_trung_41['nhiet_do_tb_7_truoc'] = (
    ngay_41['nhiet_do_cam_nhan_tb_c'].shift(1).rolling(7, min_periods=7).mean())
dac_trung_41['thu'] = dac_trung_41.index.dayofweek
dac_trung_41['thang'] = dac_trung_41.index.month
cot_dau_vao_41 = ['dien_tre_1_kw', 'dien_tre_7_kw', 'tb_7_ngay_truoc_kw',
                  'nhiet_do_truoc_1', 'nhiet_do_tb_7_truoc', 'thu', 'thang']
dac_trung_41['du_du_lieu_thu'] = dac_trung_41[cot_dau_vao_41 + ['muc_tieu_kw']].notna().all(axis=1)
print('Ngày trong lịch:', len(dac_trung_41))
print('Ngày đủ đầu vào và nhãn thử:', dac_trung_41['du_du_lieu_thu'].sum())
print(dac_trung_41[cot_dau_vao_41 + ['muc_tieu_kw']].head(10))

# %%
chia_42 = dac_trung_41.copy()
moc_test_42 = int(len(chia_42) * 0.8)
moc_val_42 = int(moc_test_42 * 0.8)
chia_42['tap'] = 'train'
chia_42.loc[chia_42.index[moc_val_42:moc_test_42], 'tap'] = 'validation'
chia_42.loc[chia_42.index[moc_test_42:], 'tap'] = 'test'

pham_vi_42 = chia_42.reset_index().groupby('tap', sort=False).agg(
    tu_ngay=('ngay', 'min'), den_ngay=('ngay', 'max'),
    so_ngay_lich=('ngay', 'size'), so_ngay_du=('du_du_lieu_thu', 'sum'))
pham_vi_42['so_ngay_loai'] = pham_vi_42['so_ngay_lich'] - pham_vi_42['so_ngay_du']
print(pham_vi_42.to_string())

# %%
lich_cham = chia_42.loc[chia_42['tap'] != 'train'].copy()
lich_cham['thang_lich'] = lich_cham.index.strftime('%Y-%m')
bao_phu = lich_cham.groupby(['tap', 'thang_lich'], sort=False).agg(
    ngay_lich=('muc_tieu_kw', 'size'), du_nhan=('muc_tieu_kw', 'count'),
    du_dau_vao_va_nhan=('du_du_lieu_thu', 'sum'))
print(bao_phu.to_string())

# %%
train_42 = chia_42.loc[chia_42['tap'].eq('train') & chia_42['du_du_lieu_thu']]
val_42 = chia_42.loc[chia_42['tap'].eq('validation') & chia_42['du_du_lieu_thu']]
test_42 = chia_42.loc[chia_42['tap'].eq('test') & chia_42['du_du_lieu_thu']]

X_train = train_42[cot_dau_vao_41].copy()
y_train = train_42['muc_tieu_kw'].copy()
X_val = val_42[cot_dau_vao_41].copy()
y_val = val_42['muc_tieu_kw'].copy()
X_test = test_42[cot_dau_vao_41].copy()
y_test = test_42['muc_tieu_kw'].copy()

assert train_42.index.max() < val_42.index.min() < test_42.index.min()
assert val_42.index.max() < test_42.index.min()
assert len(train_42) + len(val_42) + len(test_42) == chia_42['du_du_lieu_thu'].sum()
print('Số dòng train / validation / test:', len(X_train), len(X_val), len(X_test))
print('Đầu vào:', cot_dau_vao_41)
print('Có biến nhiệt độ quá khứ:', ['nhiet_do_truoc_1', 'nhiet_do_tb_7_truoc'])

# %%
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline

cot_so_423 = ['dien_tre_1_kw', 'dien_tre_7_kw', 'tb_7_ngay_truoc_kw',
              'nhiet_do_truoc_1', 'nhiet_do_tb_7_truoc']
cot_lich_423 = ['thu', 'thang']
tien_xu_ly_423 = Pipeline([('xu_ly_cot', ColumnTransformer([
    ('so', StandardScaler(), cot_so_423),
    ('lich', OneHotEncoder(drop='first', sparse_output=False), cot_lich_423)
]))])

# %%
X_train_423 = tien_xu_ly_423.fit_transform(X_train)
X_val_423 = tien_xu_ly_423.transform(X_val)
X_test_423 = tien_xu_ly_423.transform(X_test)

ten_cot_423 = tien_xu_ly_423.get_feature_names_out()
X_train_423 = pd.DataFrame(X_train_423, columns=ten_cot_423, index=X_train.index)
X_val_423 = pd.DataFrame(X_val_423, columns=ten_cot_423, index=X_val.index)
X_test_423 = pd.DataFrame(X_test_423, columns=ten_cot_423, index=X_test.index)
print('Kích thước train / validation / test:',
      X_train_423.shape, X_val_423.shape, X_test_423.shape)
print(X_train_423.iloc[:5, :6].to_string())

# %%
scaler_423 = tien_xu_ly_423.named_steps['xu_ly_cot'].named_transformers_['so']
thong_so_423 = pd.DataFrame({'cot': cot_so_423,
    'trung_binh_train': scaler_423.mean_, 'do_lech_chuan_train': scaler_423.scale_})

dau_vao_423 = pd.concat([X_train_423, X_val_423, X_test_423])
dau_vao_423['tap'] = chia_42.loc[dau_vao_423.index, 'tap']
assert np.allclose(scaler_423.mean_, X_train[cot_so_423].mean())
assert np.allclose(scaler_423.scale_, X_train[cot_so_423].std(ddof=0))
assert np.isfinite(dau_vao_423[ten_cot_423].to_numpy()).all()
print(thong_so_423.to_string(index=False))

# %%
ten_dac_trung_423 = []
for ten in tien_xu_ly_423.named_steps['xu_ly_cot'].get_feature_names_out():
    ten_dac_trung_423.append(ten.replace('so__', '').replace('lich__', ''))
print(len(ten_dac_trung_423), 'cột'); print(ten_dac_trung_423[:8])

# %%
baseline_431 = pd.DataFrame({
    'thuc_te_thu_kw': y_val,
    'du_bao_kw': X_val['dien_tre_7_kw']
})
baseline_431['sai_so_kw'] = baseline_431['du_bao_kw'] - baseline_431['thuc_te_thu_kw']
print(baseline_431.head().to_string())

# %%
from sklearn.linear_model import LinearRegression

ols_433 = Pipeline([
    ('tien_xu_ly', ColumnTransformer([
        ('so', StandardScaler(), cot_so_423),
        ('lich', OneHotEncoder(drop='first', sparse_output=False), cot_lich_423)
    ])),
    ('hoi_quy', LinearRegression())
])
ols_433.fit(X_train, y_train)
du_bao_433 = pd.DataFrame({'thuc_te_thu_kw': y_val,
    'du_bao_ols_kw': ols_433.predict(X_val)}, index=y_val.index)
du_bao_433['sai_so_kw'] = du_bao_433['du_bao_ols_kw'] - du_bao_433['thuc_te_thu_kw']
print(du_bao_433.head().to_string())

# %%
from sklearn.ensemble import RandomForestRegressor

rf_433 = Pipeline([
    ('tien_xu_ly', ColumnTransformer([
        ('so', 'passthrough', cot_so_423),
        ('lich', OneHotEncoder(sparse_output=False), cot_lich_423)
    ])),
    ('hoi_quy', RandomForestRegressor(n_estimators=200, max_depth=5,
                                    min_samples_leaf=5, random_state=42))
])
rf_433.fit(X_train, y_train)
du_bao_rf_433 = pd.DataFrame({'thuc_te_thu_kw': y_val,
    'du_bao_rf_kw': rf_433.predict(X_val)}, index=y_val.index)
print(du_bao_rf_433.head().to_string())

# %%
du_bao_44 = pd.DataFrame({'nhan_thu_kw': y_val,
    'Baseline': baseline_431['du_bao_kw'], 'OLS': du_bao_433['du_bao_ols_kw'],
    'Random Forest': du_bao_rf_433['du_bao_rf_kw']})
assert du_bao_44.notna().all().all() and len(du_bao_44) == 68

# %%
def tinh_chi_so_44(du_bao, nhan):
    ket_qua = []
    for ten in ['Baseline', 'OLS', 'Random Forest']:
        sai_so = du_bao[ten] - nhan
        mae = sai_so.abs().mean()
        rmse = np.sqrt((sai_so ** 2).mean())
        r2 = 1 - (sai_so ** 2).sum() / ((nhan - nhan.mean()) ** 2).sum()
        chon_mape = nhan.abs() >= 0.1
        mape = (sai_so.loc[chon_mape] / nhan.loc[chon_mape]).abs().mean() * 100
        ket_qua.append([ten, len(nhan), mae, rmse, r2, mape, chon_mape.sum(), mae * 24])
    bang = pd.DataFrame(ket_qua, columns=['phuong_phap', 'so_ngay', 'MAE_kw', 'RMSE_kw',
                                          'R2', 'MAPE_pct', 'so_ngay_MAPE', 'MAE_kwh_ngay'])
    bang['giam_MAE_pct'] = (1 - bang['MAE_kw'] / bang.loc[0, 'MAE_kw']) * 100
    bang['giam_RMSE_pct'] = (1 - bang['RMSE_kw'] / bang.loc[0, 'RMSE_kw']) * 100
    return bang

so_sanh_44 = tinh_chi_so_44(du_bao_44, y_val)
print(so_sanh_44.round(6).to_string(index=False))

# %%
from sklearn.linear_model import LinearRegression

cot_cu_44 = ['dien_tre_1_kw', 'dien_tre_7_kw', 'tb_7_ngay_truoc_kw', 'thu', 'thang']
ols_khong_nhiet_44 = Pipeline([('tien_xu_ly', ColumnTransformer([
    ('so', StandardScaler(), cot_cu_44[:3]),
    ('lich', OneHotEncoder(drop='first', sparse_output=False), cot_lich_423)])),
    ('hoi_quy', LinearRegression())])
ols_khong_nhiet_44.fit(X_train[cot_cu_44], y_train)
mae_khong_nhiet = (ols_khong_nhiet_44.predict(X_val[cot_cu_44]) - y_val).abs().mean()
mae_co_nhiet = (du_bao_433['du_bao_ols_kw'] - y_val).abs().mean()
so_sanh_so_bien_44 = pd.DataFrame({
    'phuong_an': ['OLS 5 biến (không nhiệt độ)', 'OLS 7 biến (có nhiệt độ quá khứ)'],
    'MAE_kw': [mae_khong_nhiet, mae_co_nhiet]})
so_sanh_so_bien_44['chenh_MAE_kw'] = so_sanh_so_bien_44['MAE_kw'] - mae_khong_nhiet
print(so_sanh_so_bien_44.round(6).to_string(index=False))
from sklearn.base import clone
rf_khong_nhiet_44 = clone(rf_433)
rf_khong_nhiet_44.set_params(tien_xu_ly__transformers=[
    ('so', 'passthrough', cot_cu_44[:3]),
    ('lich', OneHotEncoder(sparse_output=False), cot_lich_423)])
rf_khong_nhiet_44.fit(X_train[cot_cu_44], y_train)
so_sanh_so_bien_44 = pd.DataFrame({
    'mo_hinh': ['OLS', 'OLS', 'Random Forest', 'Random Forest'], 'so_bien': [5, 7, 5, 7],
    'MAE_kw': [mae_khong_nhiet, mae_co_nhiet,
        (rf_khong_nhiet_44.predict(X_val[cot_cu_44]) - y_val).abs().mean(),
        (du_bao_rf_433['du_bao_rf_kw'] - y_val).abs().mean()]})
xuat_csv(so_sanh_so_bien_44, 'doi_chung')

# %%
xuat_csv(du_bao_44.rename_axis('ngay').reset_index(), 'du_bao_validation')
xuat_csv(so_sanh_44, 'chi_so_validation')

# %%
quy_tac_test_44 = pd.DataFrame({
    'phien_ban': ['7 biến, sửa sau lần đánh giá test trước'],
    'test_doc_lap_moi': [False],
    'tieu_chi_chinh': ['MAE_kw'], 'ung_vien_tu_validation': ['Random Forest'],
    'du_lieu_fit': ['331 ngày train, không gộp validation'],
    'rf_so_cay': [200], 'rf_do_sau': [5], 'rf_la_toi_thieu': [5],
    'random_state': [42], 'nguong_MAPE_kw': [0.1]
})
print(quy_tac_test_44.to_string(index=False))

# %%
from sklearn.model_selection import TimeSeriesSplit
from sklearn.base import clone

truoc_test = chia_42.loc[chia_42['tap'].ne('test')]
chia_cua_so = TimeSeriesSplit(n_splits=3, test_size=180)
ket_qua_cua_so = []
for luot, (vi_tri_hoc, vi_tri_cham) in enumerate(chia_cua_so.split(truoc_test), start=1):
    hoc = truoc_test.iloc[vi_tri_hoc]
    cham = truoc_test.iloc[vi_tri_cham]
    hoc = hoc.loc[hoc['du_du_lieu_thu']]
    cham = cham.loc[cham['du_du_lieu_thu']]
    assert hoc.index.max() < cham.index.min() < X_test.index.min()
    for ten, mau in [('Baseline', None), ('OLS', ols_433), ('Random Forest', rf_433)]:
        if mau is None:
            du_bao = cham['dien_tre_7_kw']
        else:
            ban_sao = clone(mau)
            ban_sao.fit(hoc[cot_dau_vao_41], hoc['muc_tieu_kw'])
            du_bao = ban_sao.predict(cham[cot_dau_vao_41])
        sai_so = du_bao - cham['muc_tieu_kw']
        ket_qua_cua_so.append([luot, ten, len(hoc), len(cham),
            truoc_test.index[vi_tri_cham[0]], truoc_test.index[vi_tri_cham[-1]],
            sai_so.abs().mean(), np.sqrt((sai_so ** 2).mean())])
cua_so_44 = pd.DataFrame(ket_qua_cua_so, columns=['luot', 'mo_hinh', 'so_ngay_hoc',
    'so_ngay_cham', 'tu_ngay', 'den_ngay', 'MAE_kw', 'RMSE_kw'])
print(cua_so_44.round({'MAE_kw': 4, 'RMSE_kw': 4}).to_string(index=False))

# %%
du_bao_test_44 = pd.DataFrame({'nhan_thu_kw': y_test,
    'Baseline': X_test['dien_tre_7_kw'], 'OLS': ols_433.predict(X_test),
    'Random Forest': rf_433.predict(X_test)}, index=y_test.index)
assert len(du_bao_test_44) == 66 and du_bao_test_44.notna().all().all()
assert X_train.index.max() < X_val.index.min() < X_test.index.min()
assert X_val.index.max() < X_test.index.min()
print(du_bao_test_44.head().to_string())

# %%
so_sanh_test_44 = tinh_chi_so_44(du_bao_test_44, y_test)
print(so_sanh_test_44.round(6).to_string(index=False))

# %%
sai_so_ngay = du_bao_test_44[['Baseline', 'OLS', 'Random Forest']].sub(
    du_bao_test_44['nhan_thu_kw'], axis=0).abs()
nhom_ngay = np.where(sai_so_ngay.index.dayofweek < 5, 'Ngày thường', 'Cuối tuần')
mae_nhom = sai_so_ngay.groupby(nhom_ngay).mean()
mae_nhom['so_ngay'] = sai_so_ngay.groupby(nhom_ngay).size()
print(mae_nhom.round(4).to_string())

# %%
from scipy import stats

sai_so_val_44 = du_bao_44[['Baseline', 'OLS', 'Random Forest']].sub(y_val, axis=0).abs()
ket_qua_kd_44 = []
for mo_hinh, so_voi in [('Random Forest', 'OLS'), ('Random Forest', 'Baseline'), ('OLS', 'Baseline')]:
    d = (sai_so_val_44[mo_hinh] - sai_so_val_44[so_voi]).asfreq('D')
    n = d.count()
    tb = d.mean()
    e = (d - tb).fillna(0).to_numpy()
    for L in [7, 14, 28]:
        tong = np.sum(e ** 2)
        for k in range(1, L + 1):
            tong += 2 * (1 - k / (L + 1)) * np.sum(e[k:] * e[:-k])
        se = np.sqrt(tong) / n
        p = 2 * stats.norm.sf(abs(tb / se))
        bien = stats.norm.ppf(0.975) * se
        ket_qua_kd_44.append([mo_hinh, so_voi, L, n, tb, se, p, min(p * 3, 1), tb - bien, tb + bien])
kd_44 = pd.DataFrame(ket_qua_kd_44, columns=['mo_hinh', 'so_voi', 'tre_ngay', 'so_ngay',
    'chenh_MAE_kw', 'sai_so_chuan', 'p_value', 'p_bonferroni_3', 'can_duoi_95', 'can_tren_95'])
print(kd_44.round(6).to_string(index=False))

# %%
import sys, subprocess, pickle, json
from sklearn.base import clone

ma_do_44 = """
import sys, pickle, json
from time import perf_counter
import psutil
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
from threadpoolctl import threadpool_limits

ten, mo_hinh, X_train, y_train, X_test, du_bao_cu = pickle.loads(sys.stdin.buffer.read())
with threadpool_limits(limits=1):
    fit_ms = None
    if mo_hinh is not None:
        bat_dau = perf_counter()
        mo_hinh.fit(X_train, y_train)
        fit_ms = (perf_counter() - bat_dau) * 1000
    bat_dau = perf_counter()
    if mo_hinh is None:
        du_bao = X_test['dien_tre_7_kw'].to_numpy(copy=True)
    else:
        du_bao = mo_hinh.predict(X_test)
    predict_ms = (perf_counter() - bat_dau) * 1000
    ram_mib = psutil.Process().memory_info().peak_wset / 1024**2
assert np.allclose(du_bao, du_bao_cu, rtol=1e-9, atol=1e-10)
print(json.dumps({'phuong_phap': ten, 'fit_ms': fit_ms,
                  'predict_ms': predict_ms, 'ram_dinh_tien_trinh_MiB': ram_mib}))
"""

# %%
cac_mo_hinh_44 = {'Baseline': None, 'OLS': ols_433, 'Random Forest': rf_433}
ket_qua_do_44 = []
for luot in range(1, 4):
    for ten, mo_hinh in cac_mo_hinh_44.items():
        ban_sao = None if mo_hinh is None else clone(mo_hinh)
        dau_vao = pickle.dumps((ten, ban_sao, X_train, y_train, X_test, du_bao_test_44[ten]))
        chay = subprocess.run([sys.executable, '-c', ma_do_44], input=dau_vao,
                             capture_output=True, check=True)
        dong = json.loads(chay.stdout)
        dong['luot'] = luot
        ket_qua_do_44.append(dong)

chi_tiet_do_44 = pd.DataFrame(ket_qua_do_44)
hieu_nang_44 = chi_tiet_do_44.groupby('phuong_phap', sort=False)[
    ['fit_ms', 'predict_ms', 'ram_dinh_tien_trinh_MiB']].median().reset_index()
print(hieu_nang_44.round(3).to_string(index=False))

# %%
xuat_csv(du_bao_test_44.rename_axis('ngay').reset_index(), 'du_bao_test')
xuat_csv(so_sanh_test_44, 'chi_so_test')
xuat_csv(hieu_nang_44, 'hieu_nang')

# %%
xu_ly_45 = ols_433.named_steps['tien_xu_ly']
he_so_45 = pd.DataFrame({
    'dac_trung': xu_ly_45.get_feature_names_out(),
    'he_so_kw': ols_433.named_steps['hoi_quy'].coef_
})
print('Hệ số chặn (kW):', ols_433.named_steps['hoi_quy'].intercept_)
print(he_so_45.round(6).to_string(index=False))

# %%
xu_ly_rf_45 = rf_433.named_steps['tien_xu_ly']
quan_trong_45 = pd.DataFrame({
    'dac_trung': xu_ly_rf_45.get_feature_names_out(),
    'quan_trong_pct': rf_433.named_steps['hoi_quy'].feature_importances_ * 100
})
quan_trong_45['nhom'] = quan_trong_45['dac_trung'].str.replace('so__', '', regex=False)
quan_trong_45.loc[quan_trong_45['dac_trung'].str.startswith('lich__thu_'), 'nhom'] = 'thu'
quan_trong_45.loc[quan_trong_45['dac_trung'].str.startswith('lich__thang_'), 'nhom'] = 'thang'
nhom_rf_45 = quan_trong_45.groupby('nhom')['quan_trong_pct'].sum().sort_values(ascending=False)
print(nhom_rf_45.round(3).to_string())

# %%
from sklearn.inspection import permutation_importance

hoan_vi_45 = permutation_importance(rf_433, X_val, y_val,
    scoring='neg_mean_absolute_error', n_repeats=10, random_state=42, n_jobs=1)
bang_hoan_vi_45 = pd.DataFrame({'dac_trung': X_val.columns,
    'tang_MAE_tb_kw': hoan_vi_45.importances_mean,
    'do_lech_chuan_lap_kw': hoan_vi_45.importances_std}).sort_values('tang_MAE_tb_kw', ascending=False)
print(bang_hoan_vi_45.round(5).to_string(index=False))

# %%
ten_45 = {'dien_tre_1_kw': 'Điện ngày trước', 'dien_tre_7_kw': 'Điện tuần trước',
          'tb_7_ngay_truoc_kw': 'TB 7 ngày trước', 'nhiet_do_truoc_1': 'Nhiệt độ hôm qua',
          'nhiet_do_tb_7_truoc': 'TB nhiệt độ 7 ngày', 'thu': 'Thứ trong tuần', 'thang': 'Tháng'}
he_so_so_45 = he_so_45[he_so_45['dac_trung'].str.startswith('so__')].copy()
he_so_so_45['ten'] = he_so_so_45['dac_trung'].str.removeprefix('so__').map(ten_45)
xuat_csv(he_so_so_45, 'he_so_ols')
bang_rf = nhom_rf_45.rename_axis('dac_trung').reset_index()
bang_rf['ten'] = bang_rf['dac_trung'].map(ten_45)
xuat_csv(bang_rf, 'do_quan_trong_rf')

# %%
# Kiểm tra chất lượng từng đồng hồ phụ trên cùng ngày; không đổi mô hình.
co_tung_cot = pd.DataFrame({'ngay': power['datetime'].dt.floor('D')})
for cot in ['Global_active_power', 'Sub_metering_1', 'Sub_metering_2', 'Sub_metering_3']:
    so = pd.to_numeric(raw_b2[cot].str.replace(r'[^\d,.\-]', '', regex=True).str.replace(',', '.', regex=False), errors='coerce')
    duoi, tren = nguong_b2[cot]
    co_tung_cot[cot] = (~so.between(duoi, tren)).to_numpy()
chat_luong_sub_ngay = co_tung_cot.groupby('ngay').mean() * 100
kiem_tra_sub = []
for nguong in [20, 30, 40]:
    dat = chat_luong_sub_ngay.le(nguong).all(axis=1)
    chon = co_cau_ngay_d3.set_index('ngay')
    chon = chon.loc[dat & chon['du_1440_phut']]
    tong = chon[cot_nang_luong_d3].sum()
    kiem_tra_sub.append({'nguong_tung_cot_pct': nguong, 'so_ngay': len(chon),
        'ty_trong_sub3_pct': tong['nhom3_kwh'] / tong['tong_kwh'] * 100 if len(chon) else np.nan})
kiem_tra_sub = pd.DataFrame(kiem_tra_sub)
xuat_csv(kiem_tra_sub, 'kiem_tra_chat_luong_sub')
print(kiem_tra_sub.to_string(index=False))

# Phân biệt khả năng phát dự báo và khả năng chấm sau khi có nhãn.
lich_test = chia_42.loc[chia_42['tap'].eq('test')].copy()
du_X = lich_test[cot_dau_vao_41].notna().all(axis=1)
du_y = lich_test['muc_tieu_kw'].notna()
trang_thai = pd.DataFrame({'trang_thai': ['Đủ đầu vào và nhãn', 'Đủ đầu vào nhưng thiếu nhãn',
    'Thiếu đầu vào nhưng đủ nhãn', 'Thiếu cả hai'],
    'so_ngay': [(du_X & du_y).sum(), (du_X & ~du_y).sum(), (~du_X & du_y).sum(), (~du_X & ~du_y).sum()]})
assert trang_thai['so_ngay'].sum() == len(lich_test)
xuat_csv(trang_thai, 'bao_phu_van_hanh')
xuat_csv(bao_phu.reset_index(), 'coverage_by_month')
print(trang_thai.to_string(index=False))


# ĐỐI CHIẾU B1

# %%
import pandas as pd
import numpy as np

df = pd.read_csv(
    r"data/nguon_goc/26_2521001040_Triệu Hoàng Lan Anh_CSV1.csv",
    sep=";",
    dtype=str,
    keep_default_na=False
)

num_cols = [
    'Global_active_power', 'Global_reactive_power', 'Voltage',
    'Global_intensity', 'Sub_metering_1', 'Sub_metering_2', 'Sub_metering_3'
]

MISSING_STRS = ['?', 'null', 'NULL', 'None', 'N/A', '-', '']
gio_tho_b1 = pd.to_datetime(df['Date'] + ' ' + df['Time'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
df.replace(MISSING_STRS, np.nan, inplace=True)

for col in num_cols:
    df[col] = df[col].astype(str).str.replace(r'[^\d,.\-]', '', regex=True)
    df[col] = df[col].str.replace(',', '.', regex=False)
    df[col] = pd.to_numeric(df[col], errors='coerce')

PHYSICAL_LIMITS = {
    'Global_active_power':   (0.0, 11.0),
    'Global_reactive_power': (0.0, 2.0),
    'Voltage':               (210.0, 255.0),
    'Global_intensity':      (0.0, 50.0),
    'Sub_metering_1':        (0.0, 90.0),
    'Sub_metering_2':        (0.0, 82.0),
    'Sub_metering_3':        (0.0, 35.0),
}

for col, (lower_bound, upper_bound) in PHYSICAL_LIMITS.items():
    df.loc[(df[col] < lower_bound) | (df[col] > upper_bound), col] = np.nan

full_date_range = pd.date_range(start='2006-12-16 17:24', end='2010-11-26 21:02', freq='1min')
assert len(df) == len(full_date_range)
doc_duoc_b1 = gio_tho_b1.notna()
assert (gio_tho_b1[doc_duoc_b1].to_numpy() == full_date_range[doc_duoc_b1]).all()
df.index = full_date_range

df[num_cols] = df[num_cols].interpolate(method='time')
df[num_cols] = df[num_cols].ffill().bfill()


# %%
power_da_co = pd.read_csv(r"data/power_clean_csv1.csv", parse_dates=['datetime'])
assert len(df) == len(power_da_co), "Hai bản khác số dòng, chưa thể đối chiếu theo vị trí."
print("Số dòng:", len(df))
print("Nhãn thời gian trùng theo vị trí:", (df.index == power_da_co['datetime']).all())
ket_qua = []
for col in num_cols:
    lech = (df[col].to_numpy() - power_da_co[col].to_numpy())
    khop = np.isclose(df[col].to_numpy(), power_da_co[col].to_numpy(), rtol=1e-9, atol=1e-9, equal_nan=True)
    ket_qua.append({
        'cot': col,
        'so_dong_khac': int((~khop).sum()),
        'so_dong_khac_nguong_1e6': int((~np.isclose(df[col].to_numpy(), power_da_co[col].to_numpy(), rtol=1e-6, atol=1e-6, equal_nan=True)).sum()),
        'ty_le_khac_pct': (~khop).mean() * 100,
        'chenh_lech_lon_nhat': np.nanmax(np.abs(lech))
    })
doi_chieu_b1 = pd.DataFrame(ket_qua)
print(doi_chieu_b1.to_string(index=False))


# %%
print(doi_chieu_b1.to_string(index=False))
