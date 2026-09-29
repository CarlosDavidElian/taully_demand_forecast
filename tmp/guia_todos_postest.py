import ast, csv, json, math
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from statistics import mean
import openpyxl
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

R=Path('C:/taully_demand_forecast'); O=R/'outputs/postest_visualizado'
doc=Document(O/'Guia_verificacion_inventario_postest.docx')
for el in list(doc._element.body):
    if el.tag!=qn('w:sectPr'):doc._element.body.remove(el)
for node in ast.parse((R/'tmp/build_guia_verificacion.py').read_text(encoding='utf-8')).body:
    if isinstance(node,ast.FunctionDef) and node.name in {'p','table'}:exec(compile(ast.Module(body=[node],type_ignores=[]),'<helper>','exec'))
def page(title):doc.add_page_break();p(title,'Heading 1')
def f(x,n=2):return f'{x:.{n}f}'.replace('.',',')
def step(t):p(t,'List Number')
hist=list(csv.DictReader((R/'data/historial_demanda.csv').open(encoding='utf-8-sig')))
lookup={(r['date'],r['category']):float(r['quantity']) for r in hist}
saved=R/'data/posttests/postest_20260927t191247z_20260908_7d_70578bbb.json'
res=json.loads(saved.read_text(encoding='utf-8'))['result']; obs=res['observations']; cats=res['categories_evaluated']
assert len(obs)==35
for r in obs:assert lookup[(r['date'],r['category'])]==r['actual_demand']
wb=openpyxl.load_workbook(R/'data/inventario_producto_Taully_2026-04-01_a_2026-09-15.xlsx',read_only=True,data_only=True)
raw=[r for r in list(wb['Datos inventario'].values)[1:] if datetime(2026,9,9)<=r[0]<=datetime(2026,9,15)]
wb.close(); groups=defaultdict(list)
for r in raw:groups[(r[4],r[1])].append(r)
inv={c:dict(si=0,en=0,sf=0,sold=0,dates=set(),coverage=set()) for c in cats}
for (cat,prod),rows in groups.items():
    rows.sort(key=lambda r:r[0]); v=inv[cat];v['si']+=rows[0][6];v['en']+=sum(r[7] for r in rows);v['sf']+=rows[-1][9];v['sold']+=sum(r[5] for r in rows)
    v['dates'].update(r[0].strftime('%d/%m') for r in rows if r[10]>0);v['coverage'].update(r[0] for r in rows)
for v in inv.values():v['cd']=v['si']+v['en']-v['sf'];v['isi']=100*v['cd']/(v['si']+v['en'])
aug={c:sum(v for (d,cat),v in lookup.items() if cat==c and '2026-08-01'<=d<='2026-08-31') for c in cats}
sales={c:sum(r['actual_demand'] for r in obs if r['category']==c) for c in cats}
def metrics(rows,field='model_prediction'):
    a=[r['actual_demand'] for r in rows];b=[r[field] for r in rows];e=[abs(x-y) for x,y in zip(a,b)]
    return dict(mae=mean(e),rmse=math.sqrt(mean(z*z for z in e)),mape=mean(abs((x-y)/x) for x,y in zip(a,b) if x)*100,wape=sum(e)/sum(a)*100)
for key,val in metrics(obs).items():assert abs(val-res['model']['metrics'][key])<.00001
assert sum(v['cd'] for v in inv.values())==1311
assert sum(sales.values())==1671

p('Cómo comprobar todos los resultados del postest','Title')
p('Minimarket Taully  |  Del 9 al 15 de septiembre de 2026')
p('Esta guía explica cómo reproducir los números de las fichas del postest y del Excel del programa. Incluye los cinco indicadores y las tres dimensiones del anexo. Cada comprobación identifica el archivo de origen, la operación y el resultado esperado.')
p('1 Preparar los archivos y el período','Heading 1')
step('Abre http://127.0.0.1:5000. Selecciona Corte de prueba 08/09/2026 y Horizonte 7 días. Pulsa Ejecutar Postest y confirma que el período sea del 09/09/2026 al 15/09/2026.')
step('Pulsa Descargar Postest. Conserva una copia sin editar del Excel y crea otra para revisar fórmulas. Las hojas que se usan son Inventario, Detalle diario, Métricas categoría y Resumen postest.')
step('Para comprobar el origen, abre data/inventario_producto_Taully_2026-04-01_a_2026-09-15.xlsx, hoja Datos inventario, y data/historial_demanda.csv. Guarda también los reportes diarios POS del período y de agosto para contrastarlos con el historial.')
table(['Ficha del anexo','Qué se verifica','Dónde se explica'],[
('Cantidad demandada por inventario','CD = SI + EN − SF','Página 2'),('Índice de salida de inventario','ISI = CD / (SI + EN)','Página 2'),('Tasa de quiebre de stock','TQS = DQS / DD','Página 3'),('Estacionalidad','IE = VD / VPM','Página 5'),('Error de pronóstico','MAPE y errores diarios','Páginas 6 y 7'),('Volumen de demanda','CD = suma de unidades vendidas','Página 4'),('Patrón de demanda','IE = VP / VPR','Página 5'),('Precisión del pronóstico','MAPE, MAE y RMSE','Páginas 6 y 7')],[2.3,3,1.5])
p('Alcance de los resultados','Heading 2')
p('El inventario disponible está marcado como demostración y tiene diferencias respecto de las ventas. Verificar las operaciones no valida el origen empresarial de los datos. Los pronósticos corresponden a una evaluación histórica; no acreditan por sí solos una mejora después de una intervención real. Los valores coinciden con la ejecución guardada del 27/09/2026 usada en el Word anterior.')

page('2 Cantidad demandada e índice de salida de inventario')
p('Filtra el archivo original por Fecha, del 09 al 15/09/2026, y por Categoría. Sus columnas son A Fecha, B Producto, E Categoría, F Cantidad_Vendida, G Stock_Inicial, H Entradas, J Stock_Final y K Dias_Sin_Stock.')
step('Ordena por producto y fecha. Para cada producto toma una sola vez el primer Stock_Inicial disponible y el último Stock_Final disponible dentro del período. Suma sus entradas de todos los días.')
step('En una hoja auxiliar conserva una fila por producto y las fechas de sus saldos. Suma los stocks iniciales, entradas y stocks finales de los productos de cada categoría. No sumes los saldos de todos los días.')
step('Calcula CD = SI + EN − SF. Luego calcula ISI = CD / (SI + EN) × 100. Compara con la hoja Inventario del Excel descargado.')
table(['Categoría','SI','EN','SF','CD','ISI'],[[c,f(v['si'],0),f(v['en'],0),f(v['sf'],0),f(v['cd'],0),f(v['isi'])+' %'] for c,v in inv.items()]+[['TOTAL','1965','786','1440','1311','47,66 %']],[1.55,.9,.9,.9,.9,1.85])
p('Ejemplo de ABARROTES','Heading 2')
p('CD = 383 + 175 − 276 = 282 unidades.\nStock disponible del período = 383 + 175 = 558 unidades.\nISI = 282 / 558 × 100 = 50,54 %.')
p('Comprobación en Excel','Heading 2')
p('En Inventario, los encabezados están en la fila 6: B SI, C EN, D SF, E CD y F ISI. Copia A6:F11 a una hoja auxiliar en las mismas posiciones. Escribe =B7+C7-D7 en G7 y =G7/(B7+C7) en H7. Copia hasta la fila 11. G debe coincidir con E; aplica Porcentaje a H.')
p('El programa exporta ISI como 50,54 y añade el signo % con formato visual. La fórmula H devuelve 0,5054; para comparar numéricamente usa H7*100 frente a F7. El total se calcula como 1311/(1965+786) × 100 = 47,66 %, no como promedio de los porcentajes.')
p('Si faltan saldos de los días límite, registra la limitación: el programa toma el primer y último registro disponible. Si SI + EN es cero, el programa muestra 0; para la revisión metodológica, señala que el cociente no es calculable con denominador cero.')

page('3 Tasa de quiebre de stock')
p('DQS cuenta fechas distintas en las que al menos un producto de la categoría tiene Dias_Sin_Stock mayor que cero. DD es la duración del período: 15 − 9 + 1 = 7 días. TQS = DQS / DD × 100.')
step('Filtra Datos inventario al período del postest, a la categoría elegida y a Dias_Sin_Stock mayor que cero.')
step('Copia solamente las fechas visibles a una hoja auxiliar y usa Datos > Quitar duplicados en esa copia. Cuenta las fechas únicas; ese es DQS. Dos productos agotados el mismo día cuentan como un solo día para esa categoría.')
step('Divide DQS entre 7 y aplica formato Porcentaje. Repite la comprobación por categoría.')
table(['Categoría','Fechas con alerta','DQS','DD','TQS'],[[c,', '.join(sorted(v['dates'])) or 'Ninguna registrada',len(v['dates']),7,f(len(v['dates'])/7*100)+' %'] for c,v in inv.items()]+[['TOTAL','Categoría y día',3,35,'8,57 %']],[1.5,2,.65,.65,2.2])
p('Ejemplo de GOLOSINAS','Heading 2')
p('El archivo registra dos fechas distintas con alerta: '+', '.join(sorted(inv['GOLOSINAS']['dates']))+'. Por tanto, TQS = 2/7 × 100 = 28,57 %. HELADOS registra una fecha: 1/7 × 100 = 14,29 %.')
p('El total del anexo usa 35 combinaciones de categoría y día: 5 categorías × 7 días. TQS total = (0 + 0 + 2 + 1 + 0) / 35 × 100 = 8,57 %. No representa el porcentaje de días en que toda la tienda tuvo algún faltante.')
p('Comprobación en Excel','Heading 2')
p('En la hoja Inventario exportada, G es DQS, H es DD e K es TQS. En una hoja auxiliar divide G7/H7 y aplica formato Porcentaje. Debe mostrarse igual que K7. Igual que ISI, K contiene un valor de 0 a 100, no una fracción de 0 a 1.')
p('Comprueba también DD REGISTRADOS y COBERTURA. Tener una fila de una categoría por día no garantiza el registro de todos sus productos. La falta de ventas no demuestra un quiebre; el cálculo usa exclusivamente las alertas del inventario.')

page('4 Volumen de demanda a partir de ventas')
p('En esta dimensión, CD = Σ ventas realizadas, expresada en unidades. Es distinta de la CD calculada por balance de inventario. La ficha se resume por categoría para coincidir con el modelo; el respaldo por producto debe conservarse desde los reportes POS.')
p('Comprobación desde el historial','Heading 2')
step('Abre historial_demanda.csv. Filtra date del 09 al 15/09/2026. Agrupa category y suma quantity. Puedes usar una tabla dinámica: category en Filas y Suma de quantity en Valores.')
step('En el Excel del postest, Detalle diario contiene 35 filas: siete días por cinco categorías. La columna C, DEMANDA REAL, debe dar las mismas sumas por categoría.')
step('Contrasta además con los reportes POS: suma unidades vendidas de cada producto y usa el catálogo para asignar la categoría. Comprueba anulaciones, devoluciones y productos sin categoría según los criterios de depuración; no sumes importes monetarios.')
table(['Categoría','Ventas y CD por ventas','CD por inventario','Diferencia'],[[c,f(sales[c],0),f(inv[c]['cd'],0),f(sales[c]-inv[c]['cd'],0)] for c in cats]+[['TOTAL','1671','1311','360']],[1.7,1.8,1.7,1.6])
p('Ejemplo de ABARROTES','Heading 2')
ab=sorted([r for r in obs if r['category']=='ABARROTES'],key=lambda r:r['date'])
p('Las ventas de los siete días son '+ ' + '.join(f(r['actual_demand'],0) for r in ab)+' = 330 unidades. Esas 330 unidades aparecen como demanda real del período; no deben reemplazarse por las 282 unidades del balance de stock.')
p('En Detalle diario, =SUMA(C2:C36) debe devolver 1671. Para sumar solo ABARROTES, en una celda libre escribe =SUMAR.SI(B2:B36;"ABARROTES";C2:C36). Para otra categoría cambia el nombre.')
p('La diferencia de 360 unidades entre ventas y balance requiere revisar el kardex, las fechas disponibles, las entradas, ajustes, devoluciones y mermas. No demuestra por sí sola una pérdida concreta. Los registros de venta miden ventas atendidas; no incluyen necesariamente demanda que no pudo atenderse.')

page('5 Estacionalidad y patrón de demanda')
p('Las dos fichas usan un cociente frente a una referencia. Para reproducir el anexo anterior, se conserva agosto de 2026 como referencia anterior al corte y se compara el período completo de siete días. El programa no presenta este índice en su panel actual; se obtiene del historial de ventas.')
p('Cómo obtener la referencia','Heading 2')
step('En historial_demanda.csv, filtra del 01 al 31/08/2026 y suma quantity por categoría. Comprueba 31 fechas por categoría. Divide el total de agosto entre 31 para obtener el promedio diario de referencia.')
step('Suma las ventas del 09 al 15/09. Para Patrón de demanda usa VP = ventas de la semana / 7 y VPR = ventas de agosto / 31. Calcula IE = VP / VPR.')
step('Para reproducir la ficha Estacionalidad del Word anterior, VD es la venta semanal y VPM es el promedio diario de agosto multiplicado por 7. Calcula IE = VD / VPM. Las unidades deben ser comparables en numerador y denominador.')
table(['Categoría','Agosto total','VD semanal','VPM equivalente semanal','IE'],[[c,f(aug[c],0),f(sales[c],0),f(aug[c]/31*7),f(sales[c]/(aug[c]/31*7))] for c in cats]+[['TOTAL',f(sum(aug.values()),0),'1671',f(sum(aug.values())/31*7),f(1671/(sum(aug.values())/31*7))]],[1.45,1.1,1.1,1.95,1.2])
table(['Categoría','VP diario','VPR diario','IE'],[[c,f(sales[c]/7),f(aug[c]/31),f((sales[c]/7)/(aug[c]/31))] for c in cats],[1.7,1.7,1.7,1.7])
p('Ejemplo ABARROTES: 1748/31 = 56,387096… unidades/día de referencia. Su equivalente semanal es 394,709677…; IE = 330/394,709677… = 0,84. Con promedios diarios: (330/7)/(1748/31) = 0,84. Redondea al final, no los valores intermedios.')
p('Aclaración de la ficha: la tesis llama VD a la venta diaria y VPM al promedio mensual. El Word anterior resumió una semana y lo explicó en su criterio de cálculo. Si necesitas la medición diaria estricta, divide la venta de cada fecha entre el promedio diario de agosto y registra siete valores por categoría. El cociente semanal resume la variación frente a agosto; no demuestra por sí solo una estacionalidad recurrente.')

page('6 Error de pronóstico y precisión del pronóstico')
p('Ambas fichas comprueban el MAPE con el mismo detalle diario. DR procede de las ventas registradas y DP del modelo del programa. DP no se obtiene del inventario. La ejecución usa datos hasta el 08/09/2026 para pronosticar del 09 al 15/09/2026.')
p('Ejemplo completo de ABARROTES','Heading 2')
table(['Fecha','DR','DP modelo','Error absoluto','Error porcentual'],[[r['date'][8:10]+'/09',f(r['actual_demand'],0),f(r['model_prediction']),f(abs(r['actual_demand']-r['model_prediction'])),f(abs((r['actual_demand']-r['model_prediction'])/r['actual_demand'])*100)+' %'] for r in ab],[.75,.7,1.25,1.6,2.5])
first=ab[0]
p(f'Para el 09/09: |31 − 47,75| = 16,75 unidades de error. Error porcentual = 16,75/31 × 100 = {f(16.75/31*100)} %. El MAPE de ABARROTES es el promedio de los siete errores porcentuales sin redondear: 39,51 %.')
p('Resultados que deben coincidir','Heading 2')
table(['Categoría','DR semanal','DP semanal','MAPE'],[[c,f(sales[c],0),f(res['model']['by_category'][c]['prediction_total']),f(res['model']['by_category'][c]['mape'])+' %'] for c in cats]+[['GLOBAL','1671','1931,38','33,89 %']],[1.7,1.6,1.6,1.9])
p('El MAPE global es el promedio de los 35 errores porcentuales diarios. No se calcula con la diferencia entre 1671 y 1931,38 ni con el error de los totales de cada categoría. Aquí las 35 demandas son positivas. Si hay DR = 0, el programa excluye esa observación del MAPE; informa cuántas se excluyeron. MAE y RMSE sí utilizan todos los errores.')
p('Para reproducir DP exactamente hace falta conservar los datos, el corte y la versión del modelo. Esta guía contrasta los pronósticos guardados; no reemplaza el entrenamiento por una fórmula manual. Las cantidades pronosticadas pueden tener decimales.')

page('7 Fórmulas de Excel para todas las métricas de error')
p('En una copia de Detalle diario conserva A Fecha, B Categoría, C Demanda real, D Modelo predictivo, E PMS-7, F Error absoluto modelo y G Error absoluto PMS-7. Las observaciones van de la fila 2 a la 36. Añade columnas H, I y J para comprobar el modelo.')
table(['Celda','Fórmula','Uso'],[
('H2','=ABS(C2-D2)','Error absoluto en unidades.'),('I2','=(C2-D2)^2','Error elevado al cuadrado.'),('J2','=SI(C2=0;"";ABS((C2-D2)/C2))','Error porcentual como fracción. Aplica formato Porcentaje.')],[.65,3.35,2.8])
p('Copia H2:J2 hasta la fila 36. Después escribe estas fórmulas en celdas libres. Para WAPE y MAPE aplica formato Porcentaje con dos decimales; no multipliques por 100 si usas ese formato.')
table(['Métrica','Fórmula global','Resultado'],[
('MAE','=PROMEDIO(H2:H36)','10,41 unidades'),('RMSE','=RAIZ(PROMEDIO(I2:I36))','14,17 unidades'),('MAPE','=PROMEDIO(J2:J36)','33,89 %'),('WAPE','=SUMA(H2:H36)/SUMA(C2:C36)','21,81 %')],[.8,3.9,2.1])
p('Por categoría, usa PROMEDIO.SI con la columna B. Ejemplo del MAPE: =PROMEDIO.SI(B2:B36;"ABARROTES";J2:J36). Para RMSE, calcula la raíz del promedio de I de esa categoría. WAPE por categoría divide la suma de H entre la suma de C de la misma categoría; no promedia errores porcentuales.')
table(['Categoría','MAE','RMSE','WAPE'],[[c,f(res['model']['by_category'][c]['mae']),f(res['model']['by_category'][c]['rmse']),f(res['model']['by_category'][c]['wape'])+' %'] for c in cats],[1.7,1.7,1.7,1.7])
p('PMS-7 del panel','Heading 2')
p('Para verificar el método base, repite las columnas auxiliares sustituyendo D por E. En la misma semana: MAE 9,94; RMSE 13,26; MAPE 31,54 %; WAPE 20,82 %. PMS-7 tiene menor error global en esta ejecución. Este PMS-7 de la semana 09 al 15/09 no es el pretest de la semana 02 al 08/09.')
p('Control final','Heading 2')
p('Comprueba el período, las cinco categorías y las 35 observaciones; conserva los archivos originales y la ejecución exportada. Investiga toda diferencia antes de cambiar un resultado. Las fórmulas deben coincidir al redondear a dos decimales. La evidencia de inventario continúa siendo demostrativa hasta reemplazarla y conciliarla con registros reales.')

for style in doc.styles:
    for border in list(style.element.iter(qn('w:pBdr'))):border.getparent().remove(border)
doc.core_properties.title='Cómo comprobar todos los resultados del postest'
dest=O/'Guia_comprobar_todos_los_resultados_postest.docx';doc.save(dest)
print(dest)
print('DQS:',{c:sorted(v['dates']) for c,v in inv.items()})
