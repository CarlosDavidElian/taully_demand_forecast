from pathlib import Path
import json,urllib.request
R=Path('C:/taully_demand_forecast');O=R/'outputs/anexo2_30dias_hasta_15septiembre';rows=[]
for label in ['pre','post']:
 old=json.loads((O/(label+'_ejecucion_programa.json')).read_text(encoding='utf-8'))
 req=urllib.request.Request('http://127.0.0.1:5000/api/posttest',data=json.dumps({'cutoff_date':old['cutoff_date'],'days':30}).encode(),headers={'Content-Type':'application/json'},method='POST')
 new=json.loads(urllib.request.urlopen(req,timeout=240).read())
 assert old['observations']==new['observations'],label+' predicciones diferentes'
 assert old['model']==new['model'] and old['baseline']==new['baseline']
 assert old['inventory']['metrics']==new['inventory']['metrics']
 rows.append({'period':label,'run_id':new['run_id'],'same_observations':True,'same_metrics':True})
 print(label,'COINCIDE',new['run_id'],flush=True)
(O/'repeticion_independiente_programa.json').write_text(json.dumps(rows,indent=2),encoding='utf-8')
