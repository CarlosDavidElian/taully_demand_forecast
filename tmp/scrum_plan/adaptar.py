from pathlib import Path
base=Path('tmp/scrum_plan/metodologia.py').read_text(encoding='utf-8')
extra=r'''
page('8 Trazabilidad y tareas por Sprint')
p('La estructura del ejemplo se adapta al producto Taully mediante una relación entre historias, requerimientos, tareas y evidencias. Las actividades siguientes constituyen el plan de trabajo; su cumplimiento se registra al ejecutarlas.')
table(['Sprint e historias','Requerimientos','Tareas y evidencia de cierre'],[
['1 · HU01','RF01 y RF02','Revisar formatos y catálogo; comprobar una carga controlada. Conservar archivo fuente, registros procesados y productos sin asociación.'],
['2 · HU02','RF03 y RNF01 a RNF03','Conciliar unidades por fecha y categoría; revisar duplicados y fechas ausentes. Conservar matriz de conciliación e incidencias.'],
['3 · HU03','RF04 y RNF06','Revisar corte temporal y selección del modelo por categoría. Guardar configuración, métricas y motivos de exclusión.'],
['4 · HU04','RF05 y RNF02 y RNF04','Probar horizonte y exportación. Comparar fechas y cantidades en pantalla y archivo; revisar etiquetas de estimaciones por producto.'],
['5 · HU05 y HU06','RF06 y RF08','Reproducir el postest e implementar los índices pendientes. Conservar detalle diario, referencia de junio y comparación independiente.'],
['6 · HU07 y HU08','RF07 y RF09 y RNF05','Conciliar kardex real, probar guía y respaldo, revisar entrega. Conservar pruebas, acta y pendientes con responsables.']],[3.2,4.0,9.4])
h('Registro del Sprint Backlog')
p('Durante la planificación se asigna cada tarea a un responsable y se acuerda el esfuerzo. El estado se actualiza como Por hacer, En curso, En revisión o Terminado. La aceptación requiere las evidencias y los criterios definidos para la historia.')
table(['Tarea y requisito','Responsable','Estimación','Estado'],[['________________','____________','________','________'],['________________','____________','________','________']],[6.4,4.0,3.1,3.1])
p('No se reconstruyen tiempos reales ni gráficos de trabajo pendiente sin registros diarios. Si se usa un burndown, se conservará la unidad de esfuerzo y el saldo registrado cada día.')

page('9 Pruebas funcionales de la entrega')
p('Casos propuestos de caja negra: el revisor utiliza entradas conocidas y compara la salida con el resultado esperado. Ejecute las cargas en una copia de prueba. Los resultados obtenidos deben registrarse después de ejecutar cada caso sobre una versión identificada.')
table(['Caso y requisito','Entrada y procedimiento','Resultado esperado'],[
['CP01 · RF01 y RF02','Cargar un reporte autorizado y contrastarlo con su contenido y el catálogo.','Fechas, unidades y categorías conciliadas; registros sin asociación identificados.'],
['CP02 · RF03','Seleccionar un intervalo conocido y sumar sus ventas fuente.','Cobertura y totales coinciden con el historial, sin duplicaciones no justificadas.'],
['CP03 · RF04','Entrenar con un corte definido y revisar fechas y selección por categoría.','No se utilizan ventas posteriores al corte en el entrenamiento; método y exclusiones documentados.'],
['CP04 · RF05','Generar y descargar un pronóstico con horizonte conocido.','Fechas, categorías y cantidades coinciden entre consulta y exportación.'],
['CP05 · RF06','Usar corte 16/08/2026 y 30 días sobre la misma base y versión de la evaluación conservada.','Período 17/08 al 15/09; 150 pares diarios; DR 7851 y DP 7684,29; MAPE 26,21 %.'],
['CP06 · RF08','Después de implementar, contrastar VP, VPR e IE con ventas y junio de referencia.','VP 7851; VPR 7842; IE global redondeado 1,00. La ficha completa se visualiza y exporta. Pendiente.'],
['CP07 · RF07','Conciliar saldos y movimientos de un kardex real y recalcular indicadores.','Resultados coinciden con la fuente. Un archivo demostrativo se identifica como tal.'],
['CP08 · RF09 y RNF05','Un usuario sigue la guía; verificar restauración de una copia autorizada.','Puede reproducir el resultado y recuperar la información; registrar las incidencias.']],[2.6,7.0,7.0])
p('Los valores de CP05 son una referencia de la evaluación conservada, no una nueva prueba aprobada. Si cambian el código, las fuentes o el modelo, se analiza la diferencia antes de actualizar el valor esperado.')

page('10 Registro de pruebas y retrospectiva')
h('Resultados de la ejecución')
p('Versión: __________________  Fecha: __________________\nRevisor: __________________  Archivos: __________________')
table(['Caso','Resultado obtenido','Estado y evidencia'],[[f'CP0{i}','\n','\n'] for i in range(1,9)],[2.0,7.3,7.3])
p('Estado permitido: Conforme, No conforme o No ejecutado. Adjuntar archivo, captura o cálculo que permita reproducir cada conclusión. Las pruebas de funciones pendientes permanecen sin aprobación.')
h('Retrospectiva del Sprint')
p('Sprint: __________  Fecha: __________  Participantes: __________________')
table(['Aspecto','Acuerdo o registro'],[
['Qué funcionó','\n'],['Qué dificultó el trabajo','\n'],['Mejora para el próximo Sprint','\n'],['Responsable y fecha de seguimiento','\n']],[6.0,10.6])
p('La retrospectiva revisa la forma de trabajar. Sus acuerdos se comprueban en el siguiente ciclo y no sustituyen la revisión del incremento con el negocio.')

page('11 Acta de planificación del Sprint')
p('Sistema de pronóstico de demanda del Minimarket Taully. Este formato registra el objetivo y el trabajo acordado al iniciar cada Sprint. Reproducirlo para los seis ciclos propuestos y completar las fechas reales.')
table(['Campo','Registro'],[
['Número de acta y Sprint','________________________________________'],
['Fecha y participantes','________________________________________'],
['Inicio y fin del Sprint','________________________________________'],
['Objetivo del Sprint','\n\n'],
['Historias y requisitos seleccionados','\n\n'],
['Capacidad y estimación acordadas','________________________________________'],
['Dependencias o impedimentos','\n'],
['Criterios y evidencia de finalización','\n\n']],[6.1,10.5])
p('Los participantes acuerdan el objetivo y el plan consignados. Esta acta acredita la planificación una vez completada y firmada; la revisión de los resultados se registra en el acta de aceptación de las páginas siguientes.')
table(['Representante del negocio','Responsable del desarrollo'],[
['Firma:\n\n________________________','Firma:\n\n________________________'],
['Nombre: __________________\nCargo: ___________________\nFecha: ___________________','Guerrero Mora, Carlos David Elian\nCargo: ___________________\nFecha: ___________________']],[8.3,8.3])
p('Facilitador o Scrum Master por designar\nNombre: ____________________  Firma: ____________________')
'''
base=base.replace("page('8 Acta de revisión y aceptación del sistema')",extra+"\npage('12 Acta de revisión y aceptación del sistema')")
base=base.replace("page('9 Observaciones compromisos y firmas')","page('13 Observaciones compromisos y firmas')")
base=base.replace("Este documento adopta una estructura propia de metodología del proyecto.","Se tomó como referencia de organización el apartado Metodología Scrum del archivo Modelo_CHACON_SANCHO_TIFINAL_FICHA.pdf, con sus cuadros de backlog, planificación, pruebas, retrospectiva y actas. El contenido se adapta al sistema Taully. Este documento adopta una estructura propia de metodología del proyecto.")
base=base.replace('outputs/Metodologia_Scrum_Taully_Con_Requerimientos_y_Acta.docx','outputs/Metodologia_Scrum_Taully_Adaptada_al_Ejemplo.docx')
exec(compile(base,'metodologia_adaptada','exec'))
