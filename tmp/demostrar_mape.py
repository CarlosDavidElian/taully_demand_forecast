from pathlib import Path
import csv
from statistics import mean
from datetime import date
R=Path('C:/taully_demand_forecast')
exec((R/'tmp/scrum_plan/metodologia.py').read_text(encoding='utf-8').split("d.add_heading('Metodología Scrum")[0])
rows=list(csv.DictReader((R/'outputs/anexo2_30dias_hasta_15septiembre/detalle_30dias_verificado.csv').open(encoding='utf-8-sig')))
obs=[r for r in rows if r['period']=='pre' and r['category']=='ABARROTES']
assert len(obs)==30
real=sum(float(r['actual']) for r in obs);pred=sum(float(r['prediction']) for r in obs)
errors=[abs(float(r['actual'])-float(r['prediction']))/float(r['actual'])*100 for r in obs]
assert real==1604 and round(pred,2)==1636.18 and round(mean(errors),2)==21.20
def f(v,n=2):return f'{v:,.{n}f}'.replace(',',' ').replace('.',',')
d.add_heading('Cómo comprobar el MAPE de Abarrotes en Excel',0)
p('Pretest del 18 de julio al 16 de agosto de 2026\nMinimarket Taully · Guía de cálculo · 6 de octubre de 2026')
h('Por qué aparecen dos resultados distintos')
p('La fórmula de tu celda F15 es correcta para el error del acumulado mensual. Con demanda real de 1 604,00 y pronóstico de 1 636,18, produce 2,01 %. El 21,20 % de la ficha corresponde al promedio de los errores de los 30 días. No se obtiene a partir de esos dos totales.')
table(['Cálculo','Datos utilizados','Resultado'],[['Error del acumulado mensual','Dos cantidades: 1 604,00 y 1 636,18','2,01 %'],['MAPE diario de Abarrotes','30 pares diarios de demanda real y pronóstico PMS-7','21,20 %']],[5.1,8.0,3.5])
h('Comprobación de tu fórmula actual')
p('En F15: =ABS((D15-C15)/C15)\nDiferencia: 1 636,18 − 1 604,00 = 32,18.\nDivisión: 32,18 ÷ 1 604,00 = 0,020062344…\nCon formato Porcentaje: 2,01 %.')
p('ABS convierte la diferencia en positiva. Cambiar D15-C15 por C15-D15 no modifica el resultado al utilizar ABS. Cambiar el formato tampoco transforma 2,01 % en 21,20 %.')
h('Qué estaba incompleto en la tabla anterior')
p('La ficha mostraba demanda real y pronosticada acumuladas junto con un MAPE calculado a partir del detalle diario, sin mostrar ese detalle. Para poder reproducirlo, se incluye a continuación cada observación utilizada. Los valores proceden del detalle de cálculo conservado del proyecto; esta guía comprueba su aritmética, no certifica su origen empresarial.')
p('Conserva los encabezados: DR acumulada de 30 días, DP acumulada de 30 días y MAPE calculado con 30 pares diarios. Si mantienes tu columna F, denomínala Error del acumulado mensual.')

page('1 Los 30 pares diarios que explican el resultado')
p('Cada fila es un día de Abarrotes. DR y DP se expresan en unidades. La última columna es ABS((DP − DR) ÷ DR) expresado en porcentaje. DP se muestra con seis decimales; el cálculo original conserva toda la precisión.')
table(['Fila Excel','Fecha','DR','DP PMS-7','Error diario'],[[str(i+2),date.fromisoformat(r['date']).strftime('%d/%m/%Y'),f(float(r['actual']),0),f(float(r['prediction']),6),f(e,6)+' %'] for i,(r,e) in enumerate(zip(obs,errors))],[1.9,3.5,2.2,4.7,4.3])
for row in d.tables[-1].rows:
 for cell in row.cells:
  for pp in cell.paragraphs:pp.paragraph_format.space_before=Pt(3);pp.paragraph_format.space_after=Pt(3)
p(f'Suma DR: {f(real)}. Suma DP: {f(pred,6)}. Suma de errores porcentuales diarios: {f(sum(errors),6)}. MAPE: {f(sum(errors),6)} ÷ 30 = {f(mean(errors),6)} %, que se muestra como 21,20 %.')

page('2 Cómo reproducirlo paso a paso en Excel')
p('Utiliza una hoja nueva, para no confundir los datos diarios con los totales mensuales de tu Hoja5. Los encabezados estarán en la fila 1 y los 30 días en las filas 2 a 31.')
table(['Columna','Encabezado','Qué debes colocar'],[['A','Fecha','Las fechas de la tabla anterior, en orden.'],['B','Demanda real diaria','Los valores DR de la tabla anterior.'],['C','Pronóstico diario PMS-7','Los valores DP. Ingresa números, no texto.'],['D','Error porcentual absoluto','La fórmula para cada día.']],[1.8,5.0,9.8])
h('Paso 1 Calcular el error de cada día')
p('En D2 escribe: =ABS((C2-B2)/B2)\nArrastra la fórmula hasta D31. Aplica formato Porcentaje con dos decimales a D2:D31. No multipliques también por 100.')
p('Ejemplo de la primera fila: |54,285714… − 56| ÷ 56 = 0,03061224…; Excel mostrará 3,06 %. Este es un error diario, no el MAPE de los 30 días.')
h('Paso 2 Calcular el MAPE de la categoría')
p('En D32 escribe: =PROMEDIO(D2:D31)\nAplica formato Porcentaje con dos decimales. Resultado esperado: 21,20 %.\nComprobación del número de errores: =CONTAR(D2:D31) debe dar 30.')
h('Paso 3 Comprobar los acumulados')
p('En B32: =SUMA(B2:B31) → 1 604,00.\nEn C32: =SUMA(C2:C31) → 1 636,18.\nEn D34: =ABS((C32-B32)/B32) → 2,01 % con formato Porcentaje.\nD32 y D34 son resultados distintos porque las operaciones se aplican en un orden diferente.')
p('Al copiar DP con seis decimales, puedes obtener diferencias mínimas en los últimos decimales internos; el MAPE presentado a dos decimales se mantiene en 21,20 %. Para reproducción exacta utiliza los valores completos del detalle original. En estos 30 días no hay DR igual a cero; en otros datos, una división por cero necesita tratamiento explícito.')

page('3 Por qué los errores diarios no se compensan')
p('Al calcular el error mensual, primero se suman todos los pronósticos y todas las ventas. Los días pronosticados de más pueden compensar días pronosticados de menos. El MAPE diario toma el valor absoluto de cada error antes de promediar; por eso conserva ambas desviaciones.')
h('Ejemplo didáctico independiente de tus datos')
table(['Día','DR','DP','Error diario'],[['1','100','120','20 %'],['2','100','80','20 %'],['TOTAL','200','200','MAPE diario: 20 %']],[3.0,4.0,4.0,5.6])
p('Los totales del ejemplo coinciden: error acumulado = 0 %. Sin embargo, en cada día hubo 20 % de error: MAPE = 20 %. Por eso ambos resultados pueden ser correctos sin coincidir.')
h('Referencia para las demás categorías del pretest')
summary=[]
for cat in ['ABARROTES','BEBIDAS','GOLOSINAS','HELADOS','LIMPIEZA']:
 rr=[r for r in rows if r['period']=='pre' and r['category']==cat]
 a=sum(float(r['actual']) for r in rr);b=sum(float(r['prediction']) for r in rr)
 m=mean(abs(float(r['prediction'])-float(r['actual']))/float(r['actual'])*100 for r in rr)
 summary.append([cat.title(),f(m)+' %',f(abs(b-a)/a*100)+' %'])
table(['Categoría','MAPE de 30 pares diarios','Error del acumulado mensual'],summary,[4.2,6.2,6.2])
p('Para comprobar otra categoría, repite el procedimiento con sus propios 30 pares diarios. No promedies ni mezcles el MAPE diario con el error del acumulado. Un valor menor en la última columna no demuestra que el pronóstico diario sea mejor.')
h('Fuente y alcance')
p('Fuente: detalle_30dias_verificado.csv, carpeta outputs/anexo2_30dias_hasta_15septiembre del proyecto. Se seleccionaron las filas period = pre y category = ABARROTES y se recalcularon los errores. El método es PMS-7 reconstruido; esta guía no confirma que fuese el procedimiento utilizado históricamente por la tienda.')
out=R/'outputs/Comprobacion_MAPE_Abarrotes_30_dias.docx';d.save(out);print(out);print('MAPE',mean(errors),'TOTAL',pred)
