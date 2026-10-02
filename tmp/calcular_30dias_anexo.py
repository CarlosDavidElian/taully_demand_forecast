from pathlib import Path
from datetime import date,timedelta
from statistics import mean
from collections import defaultdict
from copy import deepcopy
import urllib.request,json,csv,hashlib,unicodedata
from openpyxl import load_workbook
R=Path('C:/taully_demand_forecast');O=R/'outputs/anexo2_30dias_hasta_15septiembre';O.mkdir(exist_ok=True)
def api(path,body):
 req=urllib.request.Request('http://127.0.0.1:5000'+path,data=json.dumps(body).encode(),headers={'Content-Type':'application/json'},method='POST')
 return urllib.request.urlopen(req,timeout=240).read()
for label,cut in [('pre','2026-07-17'),('post','2026-08-16')]:
 fp=O/(label+'_ejecucion_programa.json')
 if not fp.exists():
  print('EJECUTANDO',label,cut,flush=True)
  fp.write_bytes(api('/api/posttest',{'cutoff_date':cut,'days':30}))
 result=json.loads(fp.read_text(encoding='utf-8'))
 assert result['cutoff_date']==cut and result['horizon_days']==30 and result['observation_count']==150
 print(label,result['run_id'],result['model']['metrics'],flush=True)
 (O/('Exportacion_programa_'+label+'_30dias.xlsx')).write_bytes(api('/api/posttest/export',{'run_id':result['run_id']}))

# Independent reading of the source reports, history, and inventory.
old=(R/'tmp/verificacion_independiente_30.py').read_text(encoding='utf-8')
ns={};exec(old[:old.index('savedpath=')],ns)
H=ns['H'];cats=ns['cats'];inv=ns['inv'];checks=ns['checks'];manifest=ns['manifest'];assert not ns['mismatches']
reference={c:sum(v for (day,k),v in H.items() if day.startswith('2026-06') and k==c)/30 for c in cats}
detail=[];periods={};products={}
def test(name,valid):
 assert valid,name
 checks.append({'check':name,'passed':True})
for label in ['pre','post']:
 live=json.loads((O/(label+'_ejecucion_programa.json')).read_text(encoding='utf-8'))
 start=live['period']['start_date'];end=live['period']['end_date'];cut=live['cutoff_date']
 days=[(date.fromisoformat(start)+timedelta(days=i)).isoformat() for i in range(30)]
 score={};inventory={};pms={};products[label]=defaultdict(float)
 # Preserve product-level observations for the original Producto instrument.
 for day in days:
  rows=ns['read_sheet'](R/'data'/f'reporte_ventas_{day}.xlsx')
  idx=next(i for i,row in enumerate(rows) if row and row[0]=='PRODUCTO')
  for row in rows[idx+1:]:
   if row and row[0] and isinstance(row[1],(float,int)) and ns['norm'](row[0]) not in ['TOTAL','TOTALES']:
    products[label][str(row[0]).strip()]+=row[1]
 for c in cats:
  past=[v for (day,k),v in sorted(H.items()) if k==c and day<=cut][-7:];pred=[]
  for day in days:
   b=mean(past[-7:]);past.append(b)
   row=next(x for x in live['observations'] if x['category']==c and x['date']==day)
   test(label+' real '+c+' '+day,row['actual_demand']==H[day,c])
   test(label+' PMS7 '+c+' '+day,abs(row['pms_7_prediction']-b)<0.000001)
   dp=b if label=='pre' else row['model_prediction'];dr=H[day,c]
   assert dr>0
   pred.append(dp)
   detail.append({'period':label,'date':day,'category':c,'actual':dr,'prediction':dp,'ape_percent':abs(dr-dp)/dr*100,'pms7':b,'VD':dr,'VPM':reference[c],'IE_daily':dr/reference[c]})
  drs=[H[day,c] for day in days]
  score[c]={'DR':sum(drs),'DP':sum(pred),'MAPE':mean(abs(a-b)/a*100 for a,b in zip(drs,pred)),'VP':sum(drs),'VPR':reference[c]*30,'IE':sum(drs)/(reference[c]*30)}
  expected=live['baseline' if label=='pre' else 'model']['by_category'][c]
  test(label+' MAPE '+c,abs(score[c]['MAPE']-expected['mape'])<0.000001)
  rows=[r for r in inv if start<=r['date']<=end and r['Categoría']==c];groups=defaultdict(list)
  for r in rows:groups[r['Producto']].append(r)
  si=en=sf=ms=me=gaps=0
  for rr in groups.values():
   rr.sort(key=lambda r:r['date']);si+=rr[0]['Stock_Inicial'];sf+=rr[-1]['Stock_Final'];en+=sum(r['Entradas'] for r in rr)
   ms+=rr[0]['date']!=start;me+=rr[-1]['date']!=end;gaps+=sum(a['Stock_Final']!=b['Stock_Inicial'] for a,b in zip(rr,rr[1:]))
  stockout=sorted({r['date'] for r in rows if r['Dias_Sin_Stock']>0});dqs=len(stockout)
  im={'SI':si,'EN':en,'SF':sf,'CD':si+en-sf,'ISI':(si+en-sf)/(si+en)*100,'DQS':dqs,'DD':30,'TQS':dqs/30*100,'missing_start':ms,'missing_end':me,'gaps':gaps,'products':len(groups),'stockout_dates':stockout}
  inventory[c]=im
  checkim=next(x for x in live['inventory']['metrics'] if x['category']==c)
  for k in ['SI','EN','SF','CD','ISI','DQS','DD','TQS']:
   test(label+' inventario '+c+' '+k,abs(im[k]-checkim[k])<0.00001)
 total={k:sum(v[k] for v in inventory.values()) for k in ['SI','EN','SF','CD','DQS','DD','missing_start','missing_end','gaps','products']}
 total.update(ISI=total['CD']/(total['SI']+total['EN'])*100,TQS=total['DQS']/total['DD']*100)
 global_={k:sum(v[k] for v in score.values()) for k in ['DR','DP','VP','VPR']}
 global_.update(MAPE=mean(v['MAPE'] for v in score.values()),IE=global_['VP']/global_['VPR'])
 test(label+' producto volumen total',sum(products[label].values())==global_['DR'])
 total['sales_difference']=global_['DR']-total['CD']
 # Check the program export against the JSON and independent daily calculation.
 w=load_workbook(O/('Exportacion_programa_'+label+'_30dias.xlsx'),data_only=True)
 for row in list(w['Métricas categoría'].values)[1:]:
  actual=score[row[0]]
  test(label+' Excel categoría '+row[0],row[1]==actual['DR'] and abs(row[8 if label=='pre' else 4]-actual['MAPE'])<=0.0000005)
 for row in list(w['Detalle diario'].values)[1:]:
  x=next(x for x in detail if x['period']==label and x['date']==row[0] and x['category']==row[1])
  test(label+' Excel día '+row[0]+' '+row[1],row[2]==x['actual'] and abs(row[4 if label=='pre' else 3]-x['prediction'])<=0.0000005)
 periods[label]={'start':start,'end':end,'cutoff':cut,'days':days,'by_category':score,'global':global_,'inventory':inventory,'inventory_total':total,'run_id':live['run_id'],'baseline_same_period':live['baseline']['metrics']['mape'],'model_methods':live['training']['validation_by_category']}
 print('VERIFICADO',label,json.dumps({'global':global_,'inventory':total,'pms_same_period':periods[label]['baseline_same_period']},ensure_ascii=True),flush=True)
out={'periods':periods,'reference_month':'2026-06','reference_daily':reference,'categories':cats,'products':products,'checks':checks,'manifest':manifest,'reports':ns['result'] if 'result' in ns else {'files':168,'rows':ns['sale_rows'],'differences':ns['diff']}}
(O/'calculos_30dias_verificados.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
with (O/'detalle_30dias_verificado.csv').open('w',encoding='utf-8-sig',newline='') as fp:
 writer=csv.DictWriter(fp,fieldnames=list(detail[0]));writer.writeheader();writer.writerows(detail)
print('COMPROBACIONES',len(checks),flush=True)
