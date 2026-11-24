"""Dashboard nghiên cứu PRJ-35. Chạy: python -m streamlit run app.py"""
from pathlib import Path
import hashlib
import io
import json
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from processing import (profile, bounds_for, compare_profiles, change_summary,
                        missing_mask, numeric, time_audit, apply_step, read_upload)

BASE=Path(__file__).resolve().parent
EVIDENCE=BASE/'evidence'
ORANGE='#df782f';BLUE='#1979a9'
PAGES=['1. Tổng quan dữ liệu','2. Trước xử lý','3. Quá trình xử lý',
       '4. So sánh trước – sau','5. Đánh giá kết quả']
METHODS=['Chuẩn hóa mã thiếu','Chuyển thành số','Chuyển thành thời gian',
         'Điền trung vị','Điền trung bình','Điền hằng số','Điền tiến theo thứ tự hiện tại',
         'Loại dòng trùng hoàn toàn','Loại dòng thiếu ở cột chọn','Chặn theo IQR train']


@st.cache_data(max_entries=2,show_spinner='Đang đọc trạng thái dữ liệu thực…')
def read_state(file, modified):
    return pd.read_parquet(file)


def table(df):
    st.dataframe(df,width='stretch',hide_index=True)


def download_csv(label,df,name):
    st.download_button(label,df.to_csv(index=False).encode('utf-8-sig'),name,'text/csv')


def plot(fig,key):
    fig.update_layout(template='plotly_white',font={'family':'Arial','size':14},
                      margin={'l':20,'r':20,'t':45,'b':25},height=350)
    st.plotly_chart(fig,width='stretch',key=key)


def kpis(p):
    columns=st.columns(4)
    for box,label,value in zip(columns,['▤ Quan sát','▥ Biến','◌ Thiếu (kể cả mã)','⧉ Dòng trùng'],
                               [f'{p["rows"]:,}',p['cols'],f'{p["missing_pct"]:.2f}%',f'{p["duplicates"]:,}']):
        box.metric(label,value)


def step_log(stages):
    rows=[]
    for i,r in enumerate(stages):
        a=stages[max(0,i-1)]['profile'];b=r['profile'];c=r['change']
        rows.append({'Bước':i,'Vấn đề / trạng thái':r['name'],'Phương pháp':r['method'],
                     'Lý do / đánh đổi':r['reason'],
                     'Trước xử lý':f'{a["rows"]:,} dòng; {a["missing"]:,} ô thiếu',
                     'Sau xử lý':f'{b["rows"]:,} dòng; {b["missing"]:,} ô thiếu',
                     'Ô thay đổi trên cột chung':c['changed_cells'],
                     'Dòng có ô thay đổi':c['changed_rows'], 'Dòng bị loại':c['removed_rows'],
                     'Ô thiếu thành có giá trị':c.get('filled_missing',0),
                     'Ô có giá trị thành thiếu':c.get('introduced_missing',0),
                     'Cột mới':', '.join(c['added_columns']), 'Căn cứ':r.get('reference','Thử nghiệm người dùng')})
    return pd.DataFrame(rows)


def choose_numeric(p,key):
    cols=[r['Biến'] for r in p['table'] if r['Count']>0]
    if not cols:
        st.info('Chưa có giá trị đọc được thành số. Chuyển kiểu có kiểm soát ở phần 3 nếu phù hợp.')
        return None
    return st.selectbox('Biến số cần phân tích',cols,key=key)


def timeseries_controls(df,key):
    options=['Không sử dụng']+list(df.columns)
    if 'Date' in df and 'Time' in df:options.append('Ghép Date + Time')
    default=options.index('datetime') if 'datetime' in options else options.index('Ghép Date + Time') if 'Ghép Date + Time' in options else 0
    choice=st.selectbox('Cột thời gian (chỉ chọn khi có ý nghĩa thời gian)',options,index=default,key=key+'_time')
    if choice=='Không sử dụng':return None
    fmt=st.text_input('Định dạng thời gian; để trống để suy luận',value='%d/%m/%Y %H:%M:%S' if choice=='Ghép Date + Time' else '',key=key+'_fmt')
    raw=df.Date.astype(str)+' '+df.Time.astype(str) if choice=='Ghép Date + Time' else df[choice]
    freq=st.selectbox('Tần suất kỳ vọng để kiểm tra khoảng trống',['Không giả định','1min','1h','1D'],key=key+'_freq')
    try:
        info=time_audit(raw,fmt or None,None if freq=='Không giả định' else freq)
        st.json(info)
        return pd.to_datetime(raw,format=fmt or 'mixed',errors='coerce')
    except (ValueError,TypeError,OverflowError) as e:
        st.warning(f'Không phân tích được thời gian: {e}')
        return None


def compare_charts(a,b,col,kind,key,timecol=None):
    """Cùng biến, cùng biên bin, cùng miền trục; histogram tính trên toàn dữ liệu."""
    x=numeric(a[col]).dropna();y=numeric(b[col]).dropna()
    if x.empty and y.empty:
        st.info('Cả hai trạng thái đều không có số hữu hạn để vẽ.');return
    low=min(s.min() for s in [x,y] if len(s));high=max(s.max() for s in [x,y] if len(s))
    if low==high:low-=.5;high+=.5
    bins=np.linspace(low,high,41)
    counts=[np.histogram(s,bins=bins)[0] for s in [x,y]]
    ymax=max(1,max(int(c.max()) for c in counts))*1.12
    box_limits=[]
    for values in [x,y]:
        if len(values):
            q1,q3=values.quantile([.25,.75]);spread=q3-q1
            inner=values[values.between(q1-1.5*spread,q3+1.5*spread)]
            box_limits.extend([inner.min(),inner.max()])
    box_low,box_high=min(box_limits),max(box_limits)
    box_pad=max((box_high-box_low)*.08,.01)
    cells=st.columns(2)
    for i,(host,s,label,color) in enumerate(zip(cells,[x,y],['TRƯỚC','SAU'],[ORANGE,BLUE])):
        with host:
            if kind=='Histogram':
                fig=go.Figure(go.Bar(x=(bins[:-1]+bins[1:])/2,y=counts[i],width=np.diff(bins),marker_color=color,name=label))
                fig.update_xaxes(range=[low,high],title=col);fig.update_yaxes(range=[0,ymax],title='Số quan sát')
            elif kind=='Boxplot':
                fig=go.Figure()
                if len(s):
                    q1,med,q3=s.quantile([.25,.5,.75]);iqr=q3-q1
                    inside=s[s.between(q1-1.5*iqr,q3+1.5*iqr)]
                    fig.add_trace(go.Box(q1=[q1],median=[med],q3=[q3],lowerfence=[inside.min()],upperfence=[inside.max()],name=label,marker_color=color))
                fig.update_yaxes(range=[box_low-box_pad,box_high+box_pad],title=col)
            else:
                values=s.sort_values().to_numpy()
                indices=np.unique(np.linspace(0,max(0,len(values)-1),min(1000,len(values))).astype(int))
                fig=go.Figure(go.Scatter(x=values[indices],y=(indices+1)/len(values) if len(values) else [],mode='lines',line_color=color,name=label))
                fig.update_xaxes(range=[low,high],title=col);fig.update_yaxes(range=[0,1],title='Tỷ lệ tích lũy')
            fig.update_layout(title=f'{label} · {col} · n={len(s):,}')
            plot(fig,f'{key}_{i}')
    if kind=='Boxplot':st.caption('Phân vị và râu tính trên toàn bộ số hữu hạn. Hai trục dùng chung phạm vi râu hộp của cả hai trạng thái; các điểm ngoài râu không được vẽ và không bị xóa. Min/Max đầy đủ nằm trong bảng phía trên. Bảng IQR dùng ngưỡng gốc cố định; râu hộp dùng phân vị riêng của từng trạng thái.')
    if kind=='Phân bố tích lũy':st.caption('Phân bố tính từ toàn bộ dữ liệu; đường biểu diễn tối đa 1.000 vị trí phân vị để giao diện nhẹ.')


def overview(df,p,metadata):
    st.subheader('Dữ liệu ban đầu có gì?');kpis(p)
    st.write(metadata)
    st.caption('Thiếu gồm NaN/NaT và mã rỗng, ?, null, NULL, None, N/A, -. Mã thiếu được đếm nhưng chưa sửa trong bảng gốc. “Không đọc thành số” chỉ có ý nghĩa lỗi khi biến được kỳ vọng là số.')
    table(pd.DataFrame(p['table'])[['Biến','Kiểu','Thiếu','Thiếu (%)','Số đọc trực tiếp','Không đọc thành số']])
    st.subheader('20 dòng đầu — dữ liệu gốc');table(df.head(20))
    with st.expander('Kiểm tra mốc và khoảng thời gian',expanded=False):timeseries_controls(df,'overview')


def before_view(df,p,bounds):
    st.subheader('Nhận diện vấn đề trước khi xử lý');kpis(p)
    t=pd.DataFrame(p['table'])
    st.markdown('**A. Giá trị thiếu**')
    plot(px.bar(t,x='Biến',y='Thiếu (%)',color_discrete_sequence=[ORANGE],title='Tỷ lệ thiếu theo biến'),'missing_before')
    table(t[['Biến','Thiếu','Thiếu (%)']])
    if len(df):
        heat=pd.DataFrame({c:missing_mask(df[c].head(400)).astype(int) for c in df.columns[:30]})
        plot(px.imshow(heat.T,aspect='auto',color_continuous_scale=['#edf2f5',ORANGE],zmin=0,zmax=1,title='Vị trí thiếu · tối đa 400 dòng đầu, 30 cột đầu'),'heat')
        st.caption('Heatmap là vùng xem trước có giới hạn; KPI và thống kê sử dụng toàn bộ bảng.')
    st.markdown('**B. Trùng lặp hoàn toàn**')
    st.write(f'{p["duplicates"]:,} dòng dư sau lần xuất hiện đầu tiên, tương đương {100*p["duplicates"]/max(1,p["rows"]):.3f}% tổng dòng.')
    duplicates=df.duplicated(keep=False)
    if duplicates.any():table(df.loc[duplicates].head(20))
    else:st.success('Không có bản ghi trùng hoàn toàn; không cần loại dòng vì lý do này.')
    st.markdown('**C. IQR và phân bố — chỉ là dấu hiệu cần kiểm tra**')
    st.info('IQR dùng số đọc trực tiếp ở trạng thái gốc và giữ cố định cho mọi so sánh trong nhánh. Ngoài IQR không mặc nhiên là sai. Không dùng ngưỡng IQR khám phá này để học tiền xử lý từ test.')
    table(t[['Biến','Count','Ô ngoài IQR']].assign(**{'Tỷ lệ trên số hữu hạn (%)':t['Ô ngoài IQR']/t['Count'].replace(0,np.nan)*100}))
    col=choose_numeric(p,'before_col')
    if col:
        kind=st.selectbox('Loại biểu đồ',['Histogram','Boxplot','Phân bố tích lũy'],key='before_kind')
        # Một biểu đồ khám phá, không trình bày hai bản giống nhau là hiệu quả xử lý.
        x=numeric(df[col]).dropna()
        if kind=='Histogram':
            count,edges=np.histogram(x,bins=40)
            fig=go.Figure(go.Bar(x=(edges[:-1]+edges[1:])/2,y=count,width=np.diff(edges),marker_color=ORANGE))
        elif kind=='Boxplot':
            q1,med,q3=x.quantile([.25,.5,.75]);inner=x[x.between(q1-1.5*(q3-q1),q3+1.5*(q3-q1))]
            fig=go.Figure(go.Box(q1=[q1],median=[med],q3=[q3],lowerfence=[inner.min()],upperfence=[inner.max()],name=col,marker_color=ORANGE))
        else:
            q=np.linspace(0,1,min(1000,len(x)));fig=go.Figure(go.Scatter(x=x.quantile(q),y=q,line_color=ORANGE))
        fig.update_layout(title=f'{col} · {len(x):,} số hữu hạn');plot(fig,'before_numeric')
    st.markdown('**D. Thống kê mô tả trên số đọc trực tiếp**')
    table(t[['Biến','Count','Mean','Median','Std','Min','Q1','Q3','Max']])
    with st.expander('Diễn biến thời gian của biến chọn'):
        times=timeseries_controls(df,'before')
        if times is not None and col:
            line=pd.DataFrame({'Thời gian':times,'Giá trị':numeric(df[col])}).dropna().set_index('Thời gian').resample('D').mean().reset_index()
            if len(line):plot(px.line(line,x='Thời gian',y='Giá trị',title='Trung bình ngày từ các mốc đọc được',color_discrete_sequence=[ORANGE]),'before_time_chart')
            else:st.info('Không có cặp thời gian–giá trị hợp lệ.')


def editor(states,bounds):
    current=states[-1]['data']
    st.markdown('**Thử nghiệm có lựa chọn — chỉ áp dụng khi bấm nút**')
    st.warning('Đây là nhánh thử nghiệm trên tệp tải lên, không phải kết quả đã báo cáo trong tiểu luận. Chuyển kiểu có thể tạo NaN; loại dòng/chặn IQR có thể mất thông tin. Mọi bước đều giữ bản trước đó.')
    protected=st.checkbox('Dữ liệu có train/validation/test — chỉ học tham số trên train',value=True)
    split=None;fit=None
    if protected:
        split=st.selectbox('Cột phân chia tập',['Chưa chọn']+list(current.columns))
        if split!='Chưa chọn':
            vals=current[split].dropna().astype(str).unique().tolist()
            val=st.selectbox('Giá trị biểu thị train',vals) if vals else None
            fit=current[split].astype(str).eq(val) if val is not None else pd.Series(False,index=current.index)
            st.caption(f'Chỉ {int(fit.sum()):,} dòng train được dùng để học trung vị/trung bình/ngưỡng IQR.')
    method=st.selectbox('Phương pháp',METHODS)
    cols=st.multiselect('Cột áp dụng',list(current.columns))
    value=st.text_input('Hằng số hoặc định dạng thời gian (nếu phương pháp yêu cầu)')
    reason=st.text_input('Lý do lựa chọn và ý nghĩa biến')
    confirmed=st.checkbox('Tôi đã xem cảnh báo, kiểm tra ý nghĩa biến và chấp nhận đánh đổi của thao tác')
    if method=='Điền tiến theo thứ tự hiện tại':st.info('Không tự sắp xếp. Cần xác minh thứ tự thời gian; khi bật bảo vệ train/test, điền riêng trong từng tập, không nối qua ranh giới.')
    if st.button('Áp dụng bước xử lý',type='primary'):
        try:
            if not confirmed or not reason.strip():raise ValueError('Cần ghi lý do và xác nhận đánh đổi trước khi xử lý.')
            if len(states)>=13:raise ValueError('Giới hạn 12 bước để tránh đầy RAM; tải kết quả hoặc hoàn tác trước.')
            learned=method in ['Điền trung vị','Điền trung bình','Chặn theo IQR train']
            if protected and (learned or method=='Điền tiến theo thứ tự hiện tại') and split=='Chưa chọn':
                raise ValueError('Chọn cột phân chia train/test trước khi xử lý.')
            if protected and split in cols:raise ValueError('Không biến đổi cột phân chia tập khi bật bảo vệ.')
            after,params=apply_step(current,method,cols,value,fit if protected else None,
                                    split if protected and split!='Chưa chọn' else None)
            record={'name':method,'method':method+'; cột: '+', '.join(cols),
                    'reason':reason+'; tham số: '+json.dumps(params,ensure_ascii=False),
                    'reference':'Thử nghiệm; chưa thuộc kết quả tiểu luận',
                    'profile':profile(after,bounds),'change':change_summary(current,after),'data':after}
            states.append(record);st.rerun()
        except (ValueError,TypeError,KeyError) as e:st.error(str(e))
    if len(states)>1 and st.button('Hoàn tác bước cuối'):
        states.pop();st.rerun()


def pipeline(stages,get_data,custom,bounds):
    st.subheader('Từ vấn đề đến phương pháp và kết quả')
    st.write(' → '.join(f'{i}. {s["name"]}' for i,s in enumerate(stages)))
    st.caption('Chỉ các trạng thái đã tính mới xuất hiện trong timeline. Kiểm tra trùng và IQR không có nghĩa đã xóa dòng hay điều chỉnh điểm.')
    if custom:editor(stages,bounds)
    log=step_log(stages);table(log);download_csv('Xuất nhật ký xử lý',log,'nhat_ky_xu_ly.csv')
    index=st.selectbox('Xem chi tiết một bước',range(len(stages)),format_func=lambda x:f'{x}. {stages[x]["name"]}',index=len(stages)-1)
    r=stages[index]
    st.write('**Phương pháp:** '+r['method']);st.write('**Lý do và giới hạn:** '+r['reason']);st.caption('Căn cứ: '+r.get('reference','Thử nghiệm'))
    c=r['change'];st.write(f'{c["changed_cells"]:,} ô trên cột chung thay đổi; {c["changed_rows"]:,} dòng có ô đổi; {c["removed_rows"]:,} dòng bị loại. Cột mới: {", ".join(c["added_columns"]) or "không"}.')
    st.write(f'{c.get("filled_missing",0):,} ô thiếu thành có giá trị; {c.get("introduced_missing",0):,} ô có nội dung thành thiếu. Đây là số ô thực tế trên cột chung, không phải số đo đã được xác thực.')
    st.caption('Ô đổi là giá trị/biểu diễn khác trên cùng dòng gốc; cột mới được báo riêng, không cộng vào số ô đổi.')
    table(compare_profiles(stages[max(0,index-1)]['profile'],r['profile']))
    left,right=st.columns(2)
    with left:st.markdown('**Trước bước này · 20 dòng đầu**');table(get_data(max(0,index-1)).head(20))
    with right:st.markdown('**Sau bước này · 20 dòng đầu**');table(get_data(index).head(20))
    st.markdown('**Ví dụ thay đổi (tối đa 30 ô/dòng)**')
    if c['examples']:table(pd.DataFrame(c['examples']))
    else:st.info('Không có thay đổi giá trị trên các cột chung; xem riêng cột mới hoặc thay đổi kiểu dữ liệu.')
    if r.get('time'):st.json(r['time'])


def comparison(stages,get_data):
    st.subheader('TRƯỚC và SAU — thay đổi cụ thể, cùng thang đo')
    mode=st.radio('Phạm vi đối chiếu',['Gốc với cuối cùng','Trước và sau một bước'],horizontal=True)
    j=len(stages)-1
    if mode=='Trước và sau một bước':
        j=st.selectbox('Bước cần so sánh',range(1,len(stages)),format_func=lambda i:stages[i]['name']) if len(stages)>1 else 0
    i=0 if mode=='Gốc với cuối cùng' else max(0,j-1)
    a,b=stages[i]['profile'],stages[j]['profile']
    st.caption(f'Trước: {stages[i]["name"]} | Sau: {stages[j]["name"]}. Số liệu toàn bộ dữ liệu, không phải mẫu biểu đồ.')
    cmp=compare_profiles(a,b);table(cmp);download_csv('Tải bảng so sánh tổng hợp',cmp,'so_sanh_tong_hop.csv')
    common_cols=set(r['Biến'] for r in a['table']) & set(r['Biến'] for r in b['table'])
    ma=sum(r['Thiếu'] for r in a['table'] if r['Biến'] in common_cols)
    mb=sum(r['Thiếu'] for r in b['table'] if r['Biến'] in common_cols)
    pa=100*ma/max(1,a['rows']*len(common_cols));pb=100*mb/max(1,b['rows']*len(common_cols))
    st.info(f'Trên cùng {len(common_cols)} cột chung: tỷ lệ thiếu {pa:.3f}% → {pb:.3f}% ({pb-pa:+.3f} điểm phần trăm). Tỷ lệ toàn bảng phía trên có thể chịu ảnh hưởng của cột được thêm; không coi riêng việc thêm datetime là xử lý ô thiếu nguồn.')
    if a['duplicates']!=b['duplicates'] and a['rows']==b['rows']:
        st.caption('Số dòng giữ nguyên: thay đổi số dòng trùng không có nghĩa đã loại trùng. Thêm datetime duy nhất hoặc thay đổi giá trị có thể làm các dòng không còn giống hoàn toàn.')
    if a['rows']!=b['rows']:st.warning(f'Kích thước mẫu thay đổi: {a["rows"]:,} → {b["rows"]:,}; giảm {a["rows"]-b["rows"]:,} dòng. So phân bố cần xét sai lệch chọn mẫu.')
    frames=[]
    for p,label in [(a,'Trước'),(b,'Sau')]:
        for k,name in [('missing','Ô thiếu'),('duplicates','Dòng trùng'),('outliers','Ô ngoài IQR'),('rows','Quan sát')]:
            frames.append({'Chỉ tiêu':name,'Số lượng':p[k],'Trạng thái':label})
    plot(px.bar(pd.DataFrame(frames),x='Chỉ tiêu',y='Số lượng',color='Trạng thái',barmode='group',
                color_discrete_map={'Trước':ORANGE,'Sau':BLUE},title='Chỉ tiêu trước và sau (đơn vị ghi tại từng nhóm)'),'grouped')
    ta,tb=pd.DataFrame(a['table']),pd.DataFrame(b['table'])
    maxmissing=max(1,int(max(ta['Thiếu'].max(),tb['Thiếu'].max())))
    for cell,t,label,color in zip(st.columns(2),[ta,tb],['Trước','Sau'],[ORANGE,BLUE]):
        with cell:
            fig=px.bar(t,x='Biến',y='Thiếu',title=label+' · ô thiếu',color_discrete_sequence=[color]);fig.update_yaxes(range=[0,maxmissing*1.1]);plot(fig,'missing_'+label)
    common=[r['Biến'] for r in a['table'] if r['Biến'] in set(tb['Biến']) and (r['Count']>0 or tb.set_index('Biến').loc[r['Biến'],'Count']>0)]
    if not common:st.info('Không có biến số chung để so phân bố.');return
    col=st.selectbox('Biến so sánh',common)
    unit=st.text_input('Đơn vị chung của biến (không đổi giữa hai trạng thái)',value='kW' if col=='Global_active_power' else '')
    st.caption(f'{col}'+(f' · {unit}' if unit else ' · đơn vị theo dữ liệu nguồn'))
    stats=pd.DataFrame({'Chỉ tiêu':['Count','Mean','Median','Std','Min','Q1','Q3','Max']})
    for label,t in [('Trước',ta),('Sau',tb)]:stats[label]=[t.set_index('Biến').loc[col,k] for k in stats['Chỉ tiêu']]
    stats['Thay đổi']=stats['Sau']-stats['Trước'];table(stats)
    download_csv('Tải thống kê biến chọn',stats,'so_sanh_bien.csv')
    left,right=get_data(i),get_data(j)
    kind=st.selectbox('Biểu đồ so sánh',['Histogram','Boxplot','Phân bố tích lũy'])
    compare_charts(left,right,col,kind,'compare')
    with st.expander('Chuỗi thời gian trước/sau với cùng trục'):
        common_time=[c for c in left.columns.intersection(right.columns) if 'time' in c.lower() or 'date' in c.lower()]
        choice=st.selectbox('Cột thời gian chung',['Không sử dụng']+list(left.columns.intersection(right.columns)),index=0,key='compare_time')
        fmt=st.text_input('Định dạng thời gian chung; để trống để suy luận',key='compare_time_fmt')
        if choice!='Không sử dụng':
            series=[]
            for frame in [left,right]:
                t=pd.to_datetime(frame[choice],format=fmt or 'mixed',errors='coerce')
                daily=pd.DataFrame({'Ngày':t,'Giá trị':numeric(frame[col])}).dropna().set_index('Ngày').resample('D').mean().reset_index()
                series.append(daily)
            valid=[d for d in series if len(d)]
            if valid:
                xr=[min(d.Ngày.min() for d in valid),max(d.Ngày.max() for d in valid)]
                vals=pd.concat(valid)['Giá trị'].dropna();lo,hi=vals.min(),vals.max();pad=max((hi-lo)*.05,.01)
                for host,d,label,color in zip(st.columns(2),series,['Trước','Sau'],[ORANGE,BLUE]):
                    with host:
                        fig=px.line(d,x='Ngày',y='Giá trị',title=label+' · trung bình ngày từ mốc đọc được',color_discrete_sequence=[color])
                        fig.update_xaxes(range=xr);fig.update_yaxes(range=[lo-pad,hi+pad]);plot(fig,'time_'+label)
                st.caption('Chỉ dùng mốc đọc được của mỗi trạng thái, tổng hợp trung bình ngày; không tự lấp khoảng trống hoặc coi mốc lỗi là ngày thật.')
            else:st.info('Không có cặp thời gian–giá trị hợp lệ.')
    if st.checkbox('Xem các ô thay đổi giữa hai trạng thái (có thể mất thời gian với CSV1)'):
        with st.spinner('Đối chiếu theo dòng gốc…'):changed=change_summary(left,right)
        table(pd.DataFrame(changed['examples']))
        st.caption(f'{changed["changed_cells"]:,} ô khác trên cột chung; {changed["removed_rows"]:,} dòng bị loại; hiển thị tối đa 30 ví dụ.')


def assessment(stages,get_data,research):
    st.subheader('Đầy đủ hơn không đồng nghĩa chính xác hơn')
    a,b=stages[0]['profile'],stages[-1]['profile'];kpis(b)
    transitions=sum(s['change'].get('filled_missing',0) for s in stages[1:])
    introduced=sum(s['change'].get('introduced_missing',0) for s in stages[1:])
    actual_removed=sum(s['change']['removed_rows'] for s in stages[1:])
    st.write(f'Nhật ký ghi {transitions:,} lượt ô thiếu thành có giá trị và {introduced:,} lượt ô có giá trị thành thiếu; thực tế loại {actual_removed:,} dòng. Cùng một ô có thể xuất hiện ở nhiều bước, vì vậy số lượt không phải số ô duy nhất.')
    removed=a['rows']-b['rows'];delta=a['missing']-b['missing']
    st.write(f'Sau {len(stages)-1} bước, số ô thiếu thay đổi từ {a["missing"]:,} xuống/còn {b["missing"]:,} (chênh lệch {b["missing"]-a["missing"]:+,}); số quan sát từ {a["rows"]:,} thành {b["rows"]:,}. Tỷ lệ dòng bị loại là {100*removed/max(1,a["rows"]):.3f}%. Số ô ngoài ngưỡng IQR tham chiếu thay đổi từ {a["outliers"]:,} thành {b["outliers"]:,}.')
    st.info('Giảm ô thiếu phản ánh tính đầy đủ, không xác nhận số được điền là số đo thật. Thay đổi số dòng trùng có thể do chuyển kiểu/điền; chỉ nhật ký thao tác mới xác định có loại trùng hay không. IQR đo dấu hiệu phân bố, không đo số lỗi đã sửa.')
    ta=pd.DataFrame(a['table']).set_index('Biến');tb=pd.DataFrame(b['table']).set_index('Biến')
    changes=[]
    for col in ta.index.intersection(tb.index):
        if ta.loc[col,'Count'] or tb.loc[col,'Count']:
            changes.append({'Biến':col,'Mean trước':ta.loc[col,'Mean'],'Mean sau':tb.loc[col,'Mean'],
                            'Δ Mean':tb.loc[col,'Mean']-ta.loc[col,'Mean'],
                            'Δ Median':tb.loc[col,'Median']-ta.loc[col,'Median'],
                            'Δ Std':tb.loc[col,'Std']-ta.loc[col,'Std'],
                            '|Δ Mean| / Std trước':abs(tb.loc[col,'Mean']-ta.loc[col,'Mean'])/ta.loc[col,'Std'] if ta.loc[col,'Std']>0 else np.nan})
    if changes:
        changed=pd.DataFrame(changes).sort_values('|Δ Mean| / Std trước',ascending=False,na_position='last');table(changed)
        st.caption('Xếp hạng mức dịch chuyển tương đối; không gọi là có ý nghĩa thống kê và không dùng một ngưỡng tùy ý để khẳng định cải thiện.')
    st.markdown('**Đánh đổi và vấn đề còn lại**')
    for s in stages[1:]:st.write('• '+s['name']+': '+s['reason'])
    if len(stages)==1:st.warning('Chưa xử lý; chưa có bằng chứng để kết luận cải thiện.')
    if b['missing']:st.warning(f'Còn {b["missing"]:,} ô thiếu theo mã quy ước. Trong PRJ-35, cột thời gian thô được giữ riêng có thể còn lỗi dù datetime đã khôi phục.')
    if research:
        st.caption('Chương 4: test đã dùng lại, nhãn có giá trị điền; kết quả ở đây không chứng minh độ chính xác dự báo hay tiết kiệm điện. Dashboard HTML cũ vẫn giữ phần đánh giá mô hình.')
    download_csv('Tải bảng đánh giá trước/sau',compare_profiles(a,b),'danh_gia.csv')
    download_csv('Xuất nhật ký toàn bộ bước',step_log(stages),'nhat_ky_xu_ly.csv')
    st.caption('Dữ liệu xuất là trạng thái cuối của nhánh đang chọn; không ghi đè CSV gốc. Với CSV1 lớn, tạo CSV tải xuống có thể cần thêm RAM và thời gian.')
    if st.button('Chuẩn bị CSV dữ liệu sau xử lý'):
        with st.spinner('Đang tạo CSV…'):
            st.session_state['export_csv']=get_data(len(stages)-1).to_csv(index=True,index_label='row_id_goc').encode('utf-8-sig')
    if 'export_csv' in st.session_state:
        st.download_button('Tải dữ liệu sau xử lý',st.session_state['export_csv'],'du_lieu_sau_xu_ly.csv','text/csv')


def main():
    st.set_page_config(page_title='PRJ-35 · Nghiên cứu và xử lý dữ liệu',page_icon='📊',layout='wide')
    st.markdown('<style>.stApp{background:#f6f8fb}h1{color:#153b55}div[data-testid="stMetric"]{background:white;border:1px solid #d5e0e9;padding:16px;border-radius:8px}section[data-testid="stSidebar"]{background:#eaf0f6}</style>',unsafe_allow_html=True)
    st.title('NGHIÊN CỨU VÀ XỬ LÝ DỮ LIỆU')
    st.write('Trực quan hóa quá trình tiền xử lý và đánh giá sự thay đổi · PRJ-35, Nhóm 26')
    mode=st.sidebar.radio('Nguồn phân tích',['Tiểu luận PRJ-35','Tệp CSV/Excel tải lên'])
    page=st.sidebar.radio('Nội dung',PAGES)
    research=mode=='Tiểu luận PRJ-35'
    if research:
        manifest=EVIDENCE/'manifest.json'
        if not manifest.exists():
            st.info('Chưa tạo minh chứng. Chạy python prepare_research.py một lần; quá trình đọc toàn bộ CSV thật và lưu trạng thái vào evidence/. Không dùng dữ liệu mẫu thay thế.')
            st.stop()
        meta=json.loads(manifest.read_text(encoding='utf-8'))
        branch=st.sidebar.selectbox('Nhánh theo tiểu luận',list(meta['branches']))
        data=meta['branches'][branch];stages=data['stages'];bounds=data['bounds']
        st.info(data['note'])
        st.caption('Nguồn: CSV1 điện và CSV2 thời tiết được bàn giao trong folder (1); tài liệu tham chiếu UCI/Open-Meteo không chứng nhận các tệp này là số đo gốc đã xác minh. Quy tắc: Chương 3, mục 3.3; Chương 4, mục 4.1; mã B1 và lam_sach_thoi_tiet.')
        st.caption('Các thống kê IQR, số ô đổi và biểu đồ trước–sau được tính lại để minh họa quy trình của bài; không được trình bày như kết quả kiểm định hay phương pháp loại ngoại lai đã áp dụng trong tiểu luận.')
        def get_data(i):
            p=EVIDENCE/stages[i]['file'];return read_state(str(p),p.stat().st_mtime_ns)
        token='research:'+branch
        with st.sidebar.expander('Đối chiếu với hồ sơ nộp'):
            st.write('Điện EDA khớp bảng bàn giao trong dung sai 1e-6:',all(v['matches_tolerance'] for v in meta['power_check'].values()))
            st.write('Thời tiết khớp bảng sạch:',all(meta['weather_check'].values()))
            st.json(meta['merge_check']);st.write(meta['weather_metadata'])
            st.write('Ngưỡng thử B1 theo từng biến (không phải ngưỡng IQR):')
            st.json({'Global_active_power':[0,11],'Global_reactive_power':[0,2],
                     'Voltage':[210,255],'Global_intensity':[0,50],
                     'Sub_metering_1':[0,90],'Sub_metering_2':[0,82],'Sub_metering_3':[0,35]})
    else:
        st.info('Nhánh tải lên dùng cùng cách kiểm tra trước–sau; thao tác người dùng là thử nghiệm riêng, không tự coi là phương pháp hoặc kết quả của tiểu luận.')
        upload=st.sidebar.file_uploader('CSV hoặc Excel .xlsx',type=['csv','xlsx'])
        with st.sidebar.expander('Tùy chọn đọc tệp'):
            sep=st.selectbox('Dấu phân cách',['auto',',',';','\t','|'])
            encoding=st.selectbox('Mã hóa CSV',['utf-8-sig','utf-8','cp1258','cp1252'])
            skip=int(st.number_input('Dòng đầu cần bỏ qua',min_value=0,max_value=100,value=0))
            sheet=int(st.number_input('Chỉ số sheet Excel (từ 0)',min_value=0,value=0))
        if upload is None:st.info('Chọn CSV/Excel để bắt đầu, hoặc chuyển sang dữ liệu thật của tiểu luận ở sidebar.');st.stop()
        content=upload.getvalue()
        if len(content)>50*1024**2:st.error('Chế độ thử nghiệm giới hạn tệp tải lên 50 MiB để lưu nhiều trạng thái trong RAM. CSV1 đầy đủ có sẵn ở chế độ Tiểu luận, không bị lấy mẫu.');st.stop()
        token=hashlib.sha256(content+str((sep,encoding,skip,sheet)).encode()).hexdigest()
        if st.session_state.get('upload_token')!=token:
            try:
                df=read_upload(content,upload.name,sep,encoding,skip,sheet)
                if df.empty:st.warning('Tệp không có dòng dữ liệu. Hãy kiểm tra header/sheet.');st.stop()
                if len(df)>200_000:st.error('Tệp thử nghiệm vượt 200.000 dòng; không lấy mẫu âm thầm. Dùng nhánh PRJ-35 để xem đầy đủ 2 triệu dòng hoặc chia tệp thử nghiệm có chủ đích.');st.stop()
                bounds=bounds_for(df);p=profile(df,bounds)
                st.session_state['uploaded']=[{'name':'Gốc tải lên','method':'Chỉ đọc, giữ nguyên',
                    'reason':'Không xử lý tự động','profile':p,'data':df,'change':{'changed_cells':0,'changed_rows':0,'removed_rows':0,'added_columns':[],'examples':[]}}]
                st.session_state['upload_bounds']=bounds;st.session_state['upload_token']=token
            except (ValueError,UnicodeError,pd.errors.ParserError,pd.errors.EmptyDataError,ImportError) as e:
                st.error(f'Không đọc được tệp: {e}');st.stop()
        stages=st.session_state['uploaded'];bounds=st.session_state['upload_bounds']
        def get_data(i):return stages[i]['data']
    export_token=(token,len(stages))
    if st.session_state.get('export_token')!=export_token:
        st.session_state.pop('export_csv',None);st.session_state['export_token']=export_token
    if page==PAGES[0]:overview(get_data(0),stages[0]['profile'],'Dữ liệu gốc được giữ bất biến; các trạng thái sau được lưu riêng, nhận diện dòng bằng index gốc.')
    elif page==PAGES[1]:before_view(get_data(0),stages[0]['profile'],bounds)
    elif page==PAGES[2]:pipeline(stages,get_data,not research,bounds)
    elif page==PAGES[3]:comparison(stages,get_data)
    else:assessment(stages,get_data,research)
    st.divider();st.caption('Màu cam: trước xử lý · Màu xanh: sau xử lý. Số liệu được tính từ tệp thật; không mặc định dữ liệu sau xử lý tốt hơn ở mọi khía cạnh.')


if __name__=='__main__':main()
