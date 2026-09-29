import ast,json
from pathlib import Path
from docx import Document
from docx.shared import Inches,Pt,RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
R=Path('C:/taully_demand_forecast');O=R/'outputs/postest_visualizado'
a=json.loads((R/'tmp/auditoria_20260929/resultados.json').read_text(encoding='utf-8'))
doc=Document(O/'Guia_verificacion_inventario_postest.docx')
for el in list(doc._element.body):
    if el.tag!=qn('w:sectPr'):doc._element.body.remove(el)
for node in ast.parse((R/'tmp/build_guia_verificacion.py').read_text(encoding='utf-8')).body:
    if isinstance(node,ast.FunctionDef) and node.name in {'p','table'}:exec(compile(ast.Module(body=[node],type_ignores=[]),'<helper>','exec'))
def f(x,n=2):return f'{x:.{n}f}'.replace('.',',')
def page(title):doc.add_page_break();p(title,'Heading 1')
p('Verificación del pretest y postest','Title')
p('Minimarket Taully  |  Revisión del 29 de septiembre de 2026')
p('Dictamen','Heading 1')
p('Los cálculos principales de las fichas recientes se reproducen con los archivos del proyecto. Sin embargo, no corresponde aprobarlos como resultados definitivos de una intervención real: falta justificar el diseño y el período, validar el origen de los datos y sustituir o conciliar el inventario demostrativo. La comparación del programa sobre la misma semana favorece a PMS-7 en las cuatro métricas globales de error.')
p('Qué se comprobó','Heading 1')
table(['Comprobación','Resultado'],[
('Reportes frente al historial','168 archivos diarios, 11 748 filas de venta leídas y 43 105 unidades. Cero diferencias al consolidar por fecha y categoría con el catálogo actual.'),
('Cobertura del historial','01/04/2026 al 15/09/2026: 168 fechas consecutivas, cinco categorías y 840 filas categoría-día. Sin claves fecha-categoría duplicadas. Un registro histórico con cantidad cero.'),
('Catálogo','203 productos. Todas las ventas leídas se clasifican con el catálogo actual. Estos conteos no equivalen al número de tickets o transacciones.'),
('Repetición del postest','Corte 08/09/2026, horizonte 7 días. Diferencia máxima frente a los pronósticos guardados: 0,00. La huella del historial coincide con la de la ejecución guardada.'),
('Cálculos independientes','PMS-7 pretest, PMS-7 en la semana del postest, MAE, RMSE, MAPE, WAPE e indicadores de inventario recalculados desde sus entradas.')
],[1.55,5.25])
p('Alcance de la revisión','Heading 2')
p('Se contrastaron la tesis original, los anexos y guías entregados recientemente para septiembre, el código de cálculo y exportación, los reportes disponibles y el resultado guardado del 27/09/2026. La consistencia entre archivos no acredita que las ventas provengan de operaciones reales de la empresa. Las versiones antiguas de mayo o de múltiples cortes no deben mezclarse con estas fichas semanales.')

page('1 Pronóstico y comparación de resultados')
p('Las fichas llamadas pretest y postest usan semanas distintas. Sus cifras son reproducibles, pero la diferencia entre ellas no permite separar el cambio del método de las variaciones de demanda entre semanas.')
pre=a['periods']['pre']['baseline'];base=a['periods']['post']['baseline'];mod=a['post_model']
table(['Medición','Período','MAE','RMSE','MAPE','WAPE'],[
('Pretest PMS-7','02 al 08/09',f(pre['mae']),f(pre['rmse']),f(pre['mape'])+' %',f(pre['wape'])+' %'),
('Modelo postest','09 al 15/09',f(mod['mae']),f(mod['rmse']),f(mod['mape'])+' %',f(mod['wape'])+' %'),
('PMS-7 de control','09 al 15/09',f(base['mae']),f(base['rmse']),f(base['mape'])+' %',f(base['wape'])+' %')
],[1.5,1.25,.85,.85,1.2,1.15])
p('La comparación pertinente del desempeño predictivo es modelo contra PMS-7 del 09 al 15/09, con las mismas 35 observaciones y el mismo corte. El MAPE del modelo supera al de PMS-7 en 2,35 puntos porcentuales; MAE, RMSE y WAPE también son mayores. No hay evidencia de superioridad global del modelo en esa semana. Esto no determina su desempeño en todas las semanas.')
p('Volumen y precisión','Heading 2')
table(['Concepto','Pretest 02 al 08/09','Postest 09 al 15/09'],[
('Ventas observadas','1784 unidades','1671 unidades'),('Pronóstico acumulado','2016,97 unidades de PMS-7','1931,38 unidades del modelo'),('Observaciones para el error','35 registros diarios por categoría','35 registros diarios por categoría')
],[2.1,2.35,2.35])
p('Las tablas semanales son resúmenes; la tesis plantea demanda diaria. Debe conservarse el detalle de los siete días. MAPE se promedia sobre los errores porcentuales diarios, MAE sobre errores absolutos y RMSE es la raíz del promedio de los errores cuadrados. No se obtiene el RMSE global promediando los RMSE de las categorías. La repetición de esos valores en la ficha de precisión no constituye una muestra adicional.')
p('Métodos que realmente selecciona el programa','Heading 2')
p('En esta ejecución, ABARROTES, BEBIDAS, HELADOS y LIMPIEZA usan promedio estacional; GOLOSINAS usa Random Forest. Por ello, el conjunto debe describirse como un sistema de selección por categoría, no como cinco modelos de machine learning. El pretest PMS-7 es una referencia construida: no se encontró evidencia en los archivos revisados de que fuera el método operativo previo de la tienda.')

page('2 Inventario y estacionalidad')
p('Los valores de SI, EN, SF, CD e ISI coinciden con la regla del programa: primer saldo por producto, suma de entradas y último saldo disponible del período. Esa regla se reproduce, pero los registros no constituyen un kardex completo y conciliado.')
table(['Indicador total','Pretest','Postest'],[
('Stock inicial SI','1988','1965'),('Entradas EN','778','786'),('Stock final SF','1355','1440'),('CD por balance','1411','1311'),('Ventas observadas','1784','1671'),('Ventas menos balance','373','360'),('ISI','51,01 %','47,66 %'),('DQS / categoría-días','2 / 35','3 / 35'),('TQS','5,71 %','8,57 %')
],[2.6,2.1,2.1])
p('Problemas concretos del archivo','Heading 2')
p('Entre los productos presentes en cada semana, 107 de 167 no tienen fila en la fecha inicial del pretest y 107 no tienen fila en la fecha final. En el postest ocurre con 111 de 169 al inicio y 101 al final. Se detectan 31 cambios entre el saldo final de un registro y el inicial del siguiente en el pretest, y 23 en el postest. Las fechas intermedias pueden faltar; estas diferencias requieren movimientos y saldos verificables, no asumir que son ventas o pérdidas.')
p('Las alertas del postest son GOLOSINAS el 11 y 14/09 y HELADOS el 13/09. Se cuentan fechas únicas por categoría. Tener datos de una categoría durante siete días no demuestra cobertura de todos sus productos. El total usa 35 categoría-días, no siete días de toda la tienda. El inventario es demostrativo, por lo que tampoco permite afirmar una reducción real de quiebres.')
p('Estacionalidad y patrón de demanda','Heading 2')
p('Los cocientes entregados coinciden con ventas de la semana frente al equivalente semanal del promedio diario de agosto. Para ABARROTES: 330 / (1748/31 × 7) = 0,84. Los índices globales son 0,93 en el pretest y 0,87 en el postest.')
p('Debe corregirse la definición en el anexo: VD no puede denominarse venta diaria cuando contiene la suma semanal; VPM debe explicar que es una referencia de agosto ajustada a siete días. Alternativamente, registrar siete índices diarios por categoría. Un cociente frente a agosto describe variación, pero no prueba por sí solo estacionalidad recurrente. El panel actual no calcula ni muestra ese índice.')

page('3 Coherencia con la tesis y verificaciones técnicas')
table(['Tema','Tesis o presentación anterior','Evidencia actual y corrección necesaria'],[
('Período y población','Abril a diciembre de 2026, aproximadamente 8500 transacciones, 320 productos y 10 categorías.','Los archivos llegan al 15/09, catálogo de 203 productos y cinco categorías. Revisar la descripción o completar la recolección prevista; no confundir filas de venta, tickets y categoría-días.'),
('Diseño','Preexperimental y aplicación del modelo.','Las fichas entregadas son evaluaciones históricas. No se documenta aquí una fecha efectiva de implementación ni decisiones de reposición basadas en el modelo.'),
('Horizonte','Siete días empleados en las fichas.','Puede ser una ventana operativa de prueba, pero no justifica por sí sola el período definitivo de la tesis ni una conclusión general.'),
('Algoritmos','ARIMA, Random Forest, CatBoost y XGBoost.','La ruta actual compara Random Forest y promedio estacional. Ajustar lo declarado a lo realizado o implementar y evaluar los métodos faltantes.'),
('Validación','Validación hacia adelante.','La ruta del panel usa un corte externo y una división interna 80/20 cronológica. Una ejecución no equivale a validar múltiples cortes; las evaluaciones antiguas requieren su propia trazabilidad.'),
('Significancia','Mejora significativamente superior al método base.','No hay contraste inferencial documentado en estas fichas. Además, las cuatro métricas globales de esta semana favorecen a PMS-7.')
],[1.1,2.35,3.35])
p('Pruebas y observaciones del código','Heading 2')
p('Se ejecutaron 30 pruebas de pronóstico, postest, inventario y web: 29 pasaron y una falló por leer la selección de inventario activa del usuario. Al aislar esa selección en una carpeta temporal, la prueba fallida pasó. Es un problema de aislamiento de pruebas; no una prueba de que los cálculos del postest estén mal. No se modificó el código productivo ni las fuentes.')
p('La revisión del entrenamiento muestra que Random Forest se ajusta al 80 % inicial de la ventana interna y el modelo elegido se guarda sin reajustarlo con toda la información anterior al corte. La validación interna usa rezagos observados, mientras el pronóstico de siete días es recursivo. Debe documentarse y revisarse este diseño antes de afirmar que se optimizó el horizonte de siete días; cualquier cambio exige generar nuevas métricas, sin reemplazar las anteriores de forma silenciosa.')

page('4 Qué conservar y qué corregir antes de presentar')
p('Se pueden conservar las cifras reproducidas como resultados de una evaluación histórica sobre los archivos disponibles. Deben mantenerse identificados el corte, la semana, el método y el origen de cada indicador. Esta revisión no aprueba las fichas como evidencia definitiva de mejora del negocio.')
for t in [
'Definir con el asesor si el estudio evaluará precisión histórica o una intervención real. Para precisión, usar modelo y referencia en las mismas fechas. Para intervención, documentar implementación y mediciones reales antes y después.',
'Justificar el horizonte de siete días por la decisión operativa y evaluar varios cortes cronológicos, sin elegir únicamente las semanas favorables. Si hay ajustes posteriores al análisis, reservar una evaluación final que no haya intervenido en la selección.',
'Reemplazar el inventario demostrativo por un kardex validado y conciliar los saldos y movimientos. No corregir las diferencias de 373 y 360 unidades introduciendo cantidades inventadas.',
'Actualizar la metodología y las matrices con el período, catálogo, categorías, modelos y variables efectivamente utilizados. Diferenciar el volumen vendido de la salida por balance y definir la escala de ISI y TQS.',
'Corregir las etiquetas de estacionalidad y conservar las observaciones diarias. Un resumen semanal no cambia la unidad de análisis ni crea observaciones independientes.',
'Conservar por separado el pretest PMS-7 del 02 al 08/09 y la comparación PMS-7 contra modelo del 09 al 15/09. Identificar las versiones anteriores de los anexos para no mezclar períodos o métricas.',
'Resolver el aislamiento de la prueba web y revisar el entrenamiento antes de una nueva versión del modelo. Ejecutar de nuevo la evaluación y conservar ambos resultados si se hacen cambios.'
]:p(t,'List Number')
p('Corrección a la orientación anterior','Heading 2')
p('No era suficiente presentar las tablas como pretest y postest sin distinguir la prueba histórica de una intervención real. Tampoco basta con que los números coincidan entre Word y programa para validar la tesis. El resultado verificable es la reproducción técnica; la validez metodológica y el origen real de los datos siguen pendientes.')
p('Fuentes de verificación','Heading 2')
p('Tesis original LIMA-NORTE_PI_GURRERRO.docx; reportes de ventas y catálogo de data; historial_demanda.csv; inventario_producto_Taully_2026-04-01_a_2026-09-15.xlsx; ejecución postest_20260927t191247z_20260908_7d_70578bbb.json; código de servicios de entrenamiento, pronóstico e inventario. Recalculo y resultados detallados: tmp/auditoria_20260929/resultados.json.')
p('Referencia metodológica: Hyndman y Athanasopoulos, Forecasting Principles and Practice, secciones 5.8 y 5.10. Evaluar sobre datos no usados en el ajuste y repetir orígenes temporales permite estimar el error a lo largo de distintas ventanas.\nhttps://otexts.com/fpp3/accuracy.html\nhttps://otexts.com/fpp3/tscv.html')
for style in doc.styles:
    for el in list(style.element.iter(qn('w:pBdr'))):el.getparent().remove(el)
doc.core_properties.title='Verificación del pretest y postest'
doc.save(O/'Informe_verificacion_pretest_postest.docx')
print(O/'Informe_verificacion_pretest_postest.docx')
