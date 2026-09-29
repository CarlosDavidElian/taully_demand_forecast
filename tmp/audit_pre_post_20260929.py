import sys,json,hashlib,csv,math
from pathlib import Path
from collections import defaultdict
from datetime import date,timedelta
from statistics import mean
R=Path(__file__).resolve().parents[1];sys.path.insert(0,str(R))
import pandas as pd
from infrastructure.repositories.catalog_repository import ExcelCatalogRepository
from infrastructure.readers.excel_reader import ExcelReader
from infrastructure.repositories.csv_repository import CSVDemandRepository
from application.services.posttest_evaluation_service import PosttestEvaluationService
from application.services.inventory_service import InventoryService
from application.services.history_fingerprint import history_fingerprint

out=R/'tmp/auditoria_20260929';out.mkdir(exist_ok=True)
h=pd.read_csv(R/'data/historial_demanda.csv');h['date']=pd.to_datetime(h['date'])
lookup={(r.date.strftime('%Y-%m-%d'),r.category):r.quantity for r in h.itertuples()}
catalog=ExcelCatalogRepository(); products=catalog.get_all_products();cats=sorted({p.category for p in products})
reader=ExcelReader();rebuild=defaultdict(float);unmatched=set();units_unmatched=0;count=0;report_dates=[]
files=sorted((R/'data').glob('reporte_ventas_2026-*.xlsx'))
for path in files:
    sales=reader.read_sales(str(path));count+=len(sales)
    report_dates.append(sales[0].date.strftime('%Y-%m-%d'))
    for sale in sales:
        product=catalog.get_product(sale.product_name)
        if product:rebuild[(sale.date.strftime('%Y-%m-%d'),product.category)]+=sale.quantity
        else:unmatched.add(sale.product_name);units_unmatched+=sale.quantity
diffs=[(k,v,rebuild.get(k,0)) for k,v in lookup.items() if abs(v-rebuild.get(k,0))>1e-8]
extra=[(k,v) for k,v in rebuild.items() if k not in lookup]
invpath=R/'data/inventario_producto_Taully_2026-04-01_a_2026-09-15.xlsx'
inv=pd.read_excel(invpath,sheet_name='Datos inventario'); inv['Fecha']=pd.to_datetime(inv['Fecha'])
def score(rows):
    a=[r[0] for r in rows];p=[r[1] for r in rows];e=[x-y for x,y in rows]
    return dict(actual=sum(a),prediction=sum(p),mae=mean(abs(x) for x in e),rmse=math.sqrt(mean(x*x for x in e)),mape=mean(abs((x-y)/x) for x,y in rows if x)*100,wape=sum(abs(x) for x in e)/sum(a)*100)
periods={}
for label,start,end in [('pre','2026-09-02','2026-09-08'),('post','2026-09-09','2026-09-15')]:
    cut=date.fromisoformat(start)-timedelta(days=1)
    base=[];bycat={};invent={}
    for c in cats:
        series=h[h.category==c].sort_values('date'); past=series[series.date.dt.date<=cut].quantity.tolist()[-7:]
        actual=series[(series.date>=start)&(series.date<=end)].quantity.tolist();pred=[]
        for _ in actual:pred.append(mean(past[-7:]));past.append(pred[-1])
        paired=list(zip(actual,pred));base.extend(paired);bycat[c]=score(paired)
        window=inv[(inv['Fecha']>=start)&(inv['Fecha']<=end)&(inv['Categoría']==c)].sort_values(['Producto','Fecha'])
        si=en=sf=0; missing_start=missing_end=continuity=0
        for _,g in window.groupby('Producto'):
            si+=g.iloc[0]['Stock_Inicial'];en+=g.Entradas.sum();sf+=g.iloc[-1]['Stock_Final']
            missing_start+=g.iloc[0]['Fecha']!=pd.Timestamp(start);missing_end+=g.iloc[-1]['Fecha']!=pd.Timestamp(end)
            continuity+=int((g.Stock_Inicial.iloc[1:].to_numpy()!=g.Stock_Final.iloc[:-1].to_numpy()).sum())
        dqs=window.loc[window.Dias_Sin_Stock>0,'Fecha'].nunique()
        invent[c]=dict(si=float(si),en=float(en),sf=float(sf),cd=float(si+en-sf),isi=float((si+en-sf)/(si+en)*100),dqs=int(dqs),tqs=float(dqs/7*100),sales=float(window.Cantidad_Vendida.sum()),products=int(window.Producto.nunique()),missing_start=int(missing_start),missing_end=int(missing_end),continuity=int(continuity))
    agg={k:sum(v[k] for v in invent.values()) for k in ['si','en','sf','cd','dqs','sales','products','missing_start','missing_end','continuity']}
    agg['isi']=agg['cd']/(agg['si']+agg['en'])*100;agg['tqs']=agg['dqs']/35*100
    periods[label]=dict(start=start,end=end,baseline=score(base),baseline_by_category=bycat,inventory=invent,inventory_total=agg)
repo=CSVDemandRepository(R/'data/historial_demanda.csv')
savedpath=R/'data/posttests/postest_20260927t191247z_20260908_7d_70578bbb.json';saved=json.loads(savedpath.read_text(encoding='utf-8'))['result']
fresh=PosttestEvaluationService(repo).run(date(2026,9,8),7)
delta=max(abs(a['model_prediction']-b['model_prediction']) for a,b in zip(saved['observations'],fresh['observations']))
ind=score([(r['actual_demand'],r['model_prediction']) for r in fresh['observations']])
all_period=(h.date.max()-h.date.min()).days+1
modelpath=R/'data/models/models_by_category.pkl'
result=dict(history=dict(rows=len(h),days=int(h.date.nunique()),calendar_days=all_period,start=str(h.date.min().date()),end=str(h.date.max().date()),categories=cats,products=len(products),duplicates=int(h.duplicated(['date','category']).sum()),units=float(h.quantity.sum()),zero_rows=int((h.quantity==0).sum())),reports=dict(files=len(files),sales_rows=count,unique_dates=len(set(report_dates)),unmatched_products=len(unmatched),unmatched_units=units_unmatched,mismatches=len(diffs),extra_keys=len(extra),examples=diffs[:5]),periods=periods,post_model=ind,selected_models=fresh['training']['validation_by_category'],replay_max_prediction_difference=delta,history_fingerprint_matches=saved['methodology']['history_fingerprint']==history_fingerprint(repo.get_all_demands()),inventory_duplicates=int(inv.duplicated(['Fecha','Categoría','Producto']).sum()),inventory_rows=len(inv),saved_result=str(savedpath))
(out/'resultados.json').write_text(json.dumps(result,indent=2,ensure_ascii=False),encoding='utf-8')
print(json.dumps(result,ensure_ascii=False,indent=2))
