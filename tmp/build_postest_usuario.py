from pathlib import Path

root=Path('C:/taully_demand_forecast')
s=(root/'tmp/pretest_corregido_usuario.py').read_text(encoding='utf-8')
s=s.replace('C:/Users/Acer/Downloads/Pretest.docx','C:/taully_demand_forecast/outputs/postest_visualizado/Pretest_corregido.docx')
s=s.replace('for i in range(0,16,2)', 'for i in [0,2,4,6,9,11,13,15]')
# Retain the user's original instrument formulas and the corrected pretest layout.
s=s.replace("['periods']['pre']", "['periods']['post']")
s=s.replace('date(2026,9,2)', 'date(2026,9,9)')
s=s.replace("p('Pretest ", "p('Postest ")
s=s.replace("'Pretest. Del 02 al 08 de septiembre de 2026. Siete días.'", "'Postest. Del 09 al 15 de septiembre de 2026. Siete días.'")
s=s.replace('del 02 al 08/09/2026','del 09 al 15/09/2026')
s=s.replace('1784','1671').replace('1411','1311').replace('373','360')
s=s.replace('0,4367 equivale a 43,67 %','0,5054 equivale a 50,54 %')
s=s.replace("['','TOTAL','2','35 categoría-días','5,71 %']", "['','TOTAL','3','35 categoría-días','8,57 %']")
s=s.replace("'27,69 %'", "'33,89 %'")
s=s.replace("out=root/'outputs/postest_visualizado/Pretest_corregido.docx'", "out=root/'outputs/postest_visualizado/Postest_corregido.docx'")
old='Fuente: historial_demanda.csv; referencia PMS-7 reconstruida con corte al 01/09/2026. Cada pronóstico usa los siete valores previos y se incorpora al cálculo del día siguiente. No se ha acreditado que PMS-7 fuera el método operativo anterior de la tienda.'
new='Fuente: ejecución guardada postest_20260927t191247z_20260908_7d_70578bbb.json e historial_demanda.csv. El modelo pronostica los siete días posteriores al corte del 08/09/2026. Esta evaluación es histórica y no demuestra por sí sola una mejora frente al método previo.'
assert old in s
s=s.replace(old,new)
inject='''
saved=json.loads((root/'data/posttests/postest_20260927t191247z_20260908_7d_70578bbb.json').read_text(encoding='utf-8'))['result']
assert saved['cutoff_date']=='2026-09-08' and saved['horizon_days']==7
observations=saved['observations']
assert len(observations)==35
assert len({(o['date'],o['category']) for o in observations})==35
assert all(o['date'] in days and o['actual_demand']==lookup[o['date'],o['category']] and o['actual_demand']>0 for o in observations)
def model_score(rows):
    return {'actual':sum(r['actual_demand'] for r in rows),'prediction':sum(r['model_prediction'] for r in rows),'mape':mean(abs(r['actual_demand']-r['model_prediction'])/r['actual_demand'] for r in rows)*100}
audit['baseline_by_category']={c:model_score([r for r in observations if r['category']==c]) for c in cats}
audit['baseline']=model_score(observations)
assert round(audit['baseline']['mape'],2)==33.89
assert round(audit['baseline']['prediction'],2)==1931.38
'''
s=s.replace('assert sum(pre.values())==1671','assert sum(pre.values())==1671\n'+inject)
s=s.replace("assert txt.count('33,89 %')==2", "assert txt.count('33,89 %')==2\nassert 'Pretest' not in txt\nassert '27,69' not in txt")
out=root/'tmp/postest_corregido_usuario.py'
out.write_text(s,encoding='utf-8')
exec(compile(s,str(out),'exec'))
