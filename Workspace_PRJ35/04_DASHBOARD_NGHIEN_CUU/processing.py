"""Các phép kiểm tra và xử lý thuần Pandas; không thay đổi bảng đầu vào."""
import io
import numpy as np
import pandas as pd

MISSING = ['', '?', 'null', 'NULL', 'None', 'N/A', '-']


def missing_mask(s):
    """Tách thiếu theo mã quy ước khỏi chuỗi không chuyển được thành số."""
    return s.isna() | s.astype('string').str.strip().isin(MISSING)


def numeric(s):
    """Chỉ đọc số trực tiếp; không tự xóa đơn vị của dữ liệu tùy ý."""
    if pd.api.types.is_datetime64_any_dtype(s) or pd.api.types.is_bool_dtype(s):
        return pd.Series(np.nan, index=s.index)
    return pd.to_numeric(s.where(~missing_mask(s)), errors='coerce').replace([np.inf, -np.inf], np.nan)


def bounds_for(df):
    """Ngưỡng IQR của bảng tham chiếu; giữ cố định khi so sánh trước/sau."""
    result = {}
    for col in df:
        s = numeric(df[col]).dropna()
        if len(s) >= 4:
            q1, q3 = s.quantile([.25, .75])
            result[col] = [float(q1-1.5*(q3-q1)), float(q3+1.5*(q3-q1))]
    return result


def profile(df, bounds=None):
    """Thống kê toàn bộ dữ liệu, không lấy mẫu để tính KPI."""
    if bounds is None:
        bounds = bounds_for(df)
    rows = []
    for col in df:
        s = df[col]; mask = missing_mask(s); x = numeric(s)
        good = x.dropna(); limit = bounds.get(col)
        suspect = int(((x < limit[0]) | (x > limit[1])).sum()) if limit else 0
        rows.append({'Biến': col, 'Kiểu': str(s.dtype), 'Thiếu': int(mask.sum()),
                     'Thiếu (%)': float(mask.mean()*100) if len(s) else 0.,
                     'Số đọc trực tiếp': int(x.count()),
                     'Không đọc thành số': int((~mask & x.isna()).sum()),
                     'Ô ngoài IQR': suspect,
                     'Count': int(good.size), 'Mean': good.mean(), 'Median': good.median(),
                     'Std': good.std(), 'Min': good.min(), 'Q1': good.quantile(.25),
                     'Q3': good.quantile(.75), 'Max': good.max()})
    table = pd.DataFrame(rows)
    cells = df.size
    missing = int(table['Thiếu'].sum()) if rows else 0
    return {'rows': len(df), 'cols': len(df.columns), 'missing': missing,
            'missing_pct': 100*missing/cells if cells else 0.,
            'duplicates': int(df.duplicated().sum()),
            'outliers': int(table['Ô ngoài IQR'].sum()) if rows else 0,
            'nan': int(df.isna().sum().sum()), 'table': table.to_dict('records')}


def change_summary(before, after):
    """Đối chiếu bằng index dòng gốc; NaN–NaN không tính là thay đổi."""
    ids = before.index.intersection(after.index)
    cols = before.columns.intersection(after.columns)
    changed = 0; changed_rows = pd.Series(False, index=ids)
    filled_missing = 0; introduced_missing = 0
    examples = []
    for col in cols:
        a, b = before.loc[ids, col], after.loc[ids, col]
        ma, mb = missing_mask(a), missing_mask(b)
        filled_missing += int((ma & ~mb).sum())
        introduced_missing += int((~ma & mb).sum())
        equal = a.eq(b).fillna(False) | (a.isna() & b.isna())
        diff = ~equal
        changed += int(diff.sum()); changed_rows |= diff
        for idx in diff[diff].index[:max(0, 30-len(examples))]:
            examples.append({'Dòng gốc (từ 1)': int(idx)+1, 'Biến': col,
                             'Trước': str(a.loc[idx]), 'Sau': str(b.loc[idx])})
    removed = before.index.difference(after.index)
    for idx in removed[:max(0, 30-len(examples))]:
        examples.append({'Dòng gốc (từ 1)': int(idx)+1, 'Biến': '(toàn dòng)',
                         'Trước': 'Có bản ghi', 'Sau': 'Đã loại'})
    return {'changed_cells': changed, 'changed_rows': int(changed_rows.sum()),
            'filled_missing': filled_missing, 'introduced_missing': introduced_missing,
            'removed_rows': len(removed), 'added_columns': list(after.columns.difference(before.columns)),
            'examples': examples}


def compare_profiles(a, b):
    labels = [('rows','Số dòng'), ('cols','Số cột'), ('missing','Ô thiếu (kể cả mã thiếu)'),
              ('missing_pct','Tỷ lệ thiếu (%) — chênh lệch là điểm phần trăm'),
              ('duplicates','Dòng trùng hoàn toàn, không tính lần đầu'),
              ('outliers','Ô ngoài IQR cố định'), ('nan','Ô NaN/NaT thực')]
    return pd.DataFrame([{'Chỉ tiêu': label, 'Trước': a[k], 'Sau': b[k],
                          'Thay đổi': b[k]-a[k]} for k, label in labels])


def time_audit(s, fmt=None, freq=None):
    t = pd.to_datetime(s, format=fmt or 'mixed', errors='coerce')
    clean = t.dropna()
    result = {'Mốc không đọc được': int(t.isna().sum()),
              'Mốc trùng (sau lần đầu)': int(clean.duplicated().sum()),
              'Đã tăng dần': bool(clean.is_monotonic_increasing),
              'Đầu': str(clean.min()), 'Cuối': str(clean.max())}
    if freq and len(clean):
        delta = pd.Timedelta(freq)
        if delta <= pd.Timedelta(0):
            raise ValueError('Tần suất phải dương.')
        unique = clean.drop_duplicates().sort_values()
        # Không dựng một date_range khổng lồ cho tệp tùy ý.
        gaps = unique.diff().dropna()
        result['Mốc thiếu ước tính theo tần suất đã chọn'] = int(((gaps/delta).apply(np.floor)-1).clip(lower=0).sum())
        result['Khoảng cách không khớp bội tần suất'] = int((gaps % delta != pd.Timedelta(0)).sum())
    return result


def apply_step(df, method, cols, value=None, fit_mask=None, split_col=None):
    """Tạo trạng thái mới. Các tham số học từ dữ liệu chỉ dùng fit_mask (train)."""
    out = df.copy(deep=True)
    train = df if fit_mask is None else df.loc[fit_mask]
    if method in ['Điền trung vị', 'Điền trung bình', 'Chặn theo IQR train'] and train.empty:
        raise ValueError('Không có dòng train để học tham số.')
    params = {}
    if method == 'Loại dòng trùng hoàn toàn':
        return out.loc[~out.duplicated()].copy(), params
    if not cols:
        raise ValueError('Chọn ít nhất một cột.')
    for col in cols:
        if method == 'Chuẩn hóa mã thiếu':
            out[col] = out[col].mask(missing_mask(out[col]))
        elif method == 'Chuyển thành số':
            out[col] = pd.to_numeric(out[col], errors='coerce')
        elif method == 'Chuyển thành thời gian':
            if not value:
                raise ValueError('Cần định dạng thời gian, ví dụ %Y-%m-%d.')
            out[col] = pd.to_datetime(out[col], format=value, errors='coerce')
        elif method == 'Điền hằng số':
            fill = float(value) if pd.api.types.is_numeric_dtype(out[col]) else value
            out[col] = out[col].mask(missing_mask(out[col]), fill)
        elif method in ['Điền trung vị', 'Điền trung bình', 'Chặn theo IQR train']:
            if not pd.api.types.is_numeric_dtype(out[col]):
                raise ValueError(f'{col}: cần chuyển kiểu số trước.')
            x = numeric(train[col]).dropna()
            if x.empty:
                raise ValueError(f'{col}: train không có số hữu hạn.')
            if method == 'Chặn theo IQR train':
                if len(x) < 4:
                    raise ValueError('Cần ít nhất 4 giá trị train để tính IQR.')
                q1,q3=x.quantile([.25,.75]); lo,hi=q1-1.5*(q3-q1),q3+1.5*(q3-q1)
                out[col] = out[col].clip(lo,hi); params[col]=[float(lo),float(hi)]
            else:
                fill=x.median() if method=='Điền trung vị' else x.mean()
                out[col]=out[col].fillna(fill);params[col]=float(fill)
        elif method == 'Điền tiến theo thứ tự hiện tại':
            out[col]=out[col].mask(missing_mask(out[col]))
            out[col]=out.groupby(split_col,dropna=False)[col].ffill() if split_col else out[col].ffill()
        elif method == 'Loại dòng thiếu ở cột chọn':
            out=out.loc[~missing_mask(out[col])].copy()
        else:
            raise ValueError('Phương pháp chưa hỗ trợ.')
    return out,params


def read_upload(content, filename, sep='auto', encoding='utf-8-sig', skiprows=0, sheet=0):
    """Đọc bảng bất kỳ; giữ chuỗi gốc để người dùng quyết định chuyển kiểu."""
    if filename.lower().endswith('.xlsx'):
        df=pd.read_excel(io.BytesIO(content),sheet_name=sheet,dtype=str,keep_default_na=False,skiprows=skiprows)
    else:
        df=pd.read_csv(io.BytesIO(content),sep=None if sep=='auto' else sep,
                       engine='python' if sep=='auto' else 'c',encoding=encoding,
                       dtype=str,keep_default_na=False,skiprows=skiprows)
    df.columns=df.columns.astype(str)
    if not df.columns.is_unique:
        raise ValueError('Tên cột không duy nhất; hãy sửa tiêu đề trước khi phân tích.')
    return df.reset_index(drop=True)
