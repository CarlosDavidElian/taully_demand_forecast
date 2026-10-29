import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const input=JSON.parse(await fs.readFile(new URL('./input.json',import.meta.url),'utf8'));
const wb=Workbook.create();
const summary=wb.worksheets.add('Resultados Hoja5');
const pre=wb.worksheets.add('Diario pretest');
const post=wb.worksheets.add('Diario postest');
const cats=['ABARROTES','BEBIDAS','GOLOSINAS','HELADOS','LIMPIEZA'];
function style(s,range){s.showGridLines=false;s.getRange(range).format.font.name='Arial';s.getRange(range).format.font.size=11;s.getRange(range).format.rowHeight=23;}
function header(s,r){s.getRange(r).format.fill='#17365D';s.getRange(r).format.font.color='#FFFFFF';s.getRange(r).format.font.bold=true;s.getRange(r).format.wrapText=true;s.getRange(r).format.rowHeight=44;}
for(const [s,key,label] of [[pre,'pre','PMS-7 reconstruido'],[post,'post','Modelo del programa']]){
 style(s,'A1:E155');s.getRange('A:A').format.columnWidth=15;s.getRange('B:B').format.columnWidth=20;s.getRange('C:D').format.columnWidth=22;s.getRange('E:E').format.columnWidth=25;
 s.getRange('A1:E1').merge();s.getRange('A1').values=[[`${s.name || (key==='pre'?'Pretest':'Postest')} — ${label}`]];
 s.getRange('A2:E2').merge();s.getRange('A2').values=[['Fuente: detalle_30dias_verificado.csv conservado del proyecto.']];
 s.getRange('A4:E4').values=[['Fecha','Categoría','Demanda real diaria','Pronóstico diario','Error absoluto porcentual']];header(s,'A4:E4');
 const rows=input.detail.filter(r=>r.period===key).sort((a,b)=>cats.indexOf(a.category)-cats.indexOf(b.category)||a.date.localeCompare(b.date));
 s.getRange('A5:D154').values=rows.map(r=>[new Date(r.date+'T00:00:00Z'),r.category,Number(r.actual),Number(r.prediction)]);
 s.getRange('A5:A154').setNumberFormat('dd/mm/yyyy');s.getRange('C5:D154').setNumberFormat('#,##0.000000');
 s.getRange('E5:E154').formulas=rows.map((r,i)=>[`=IF(C${i+5}=0,"",ABS((D${i+5}-C${i+5})/C${i+5}))`]);s.getRange('E5:E154').setNumberFormat('0.00%');
 s.freezePanes.freezeRows(4);
}
style(summary,'A1:H38');summary.getRange('A:A').format.columnWidth=19;summary.getRange('B:C').format.columnWidth=18;summary.getRange('D:H').format.columnWidth=21;
summary.getRange('A1:H1').merge();summary.getRange('A1').values=[['VERIFICACIÓN DEL PRETEST Y POSTEST — HOJA 5']];summary.getRange('A1').format.font.size=16;summary.getRange('A1:H1').format.rowHeight=32;
for(const [r,t] of [[2,'MAPE diario = promedio de los errores de cada día. Error mensual = error de los acumulados.'],[3,'Datos diarios conservados del proyecto; valores de referencia tomados de Libro1.xlsx, Hoja5.'],[24,'La columna D se calcula desde los datos diarios. La columna E reproduce el cálculo con totales de tu Hoja5.'],[25,'Doble clic en cualquier resultado para ver la fórmula. En las hojas Diario puedes revisar los 30 pares de cada categoría.'],[26,'Las diferencias de pronóstico de hasta 0,005 unidades son redondeo. El MAPE de referencia se compara a 2 decimales porcentuales.'],[27,'La demanda cero se excluye del MAPE, no se transforma en error cero. El detalle original de estas ventanas no contiene demanda cero.'],[36,'Fuentes: Libro1.xlsx (Hoja5 C15:E19 y C43:E47) y outputs/anexo2_30dias_hasta_15septiembre/detalle_30dias_verificado.csv.'],[37,'Evaluación retrospectiva. PMS-7 reconstruido en pretest; modelo en postest. La comparación entre ventanas no acredita por sí sola causalidad.']]){
summary.getRange(`A${r}:H${r}`).merge();summary.getRange(`A${r}`).values=[[t]];summary.getRange(`A${r}:H${r}`).format.wrapText=true;summary.getRange(`A${r}:H${r}`).format.rowHeight=32;
}
for(const [key,start,sheet,title] of [['pre',6,'Diario pretest','PRETEST · 18/07/2026 al 16/08/2026 · PMS-7'],['post',16,'Diario postest','POSTEST · 17/08/2026 al 15/09/2026 · Modelo']]){
 summary.getRange(`A${start-1}:H${start-1}`).merge();summary.getRange(`A${start-1}`).values=[[title]];summary.getRange(`A${start-1}`).format.font.bold=true;
 summary.getRange(`A${start}:H${start}`).values=[['Categoría','DR acumulada','DP acumulada','MAPE diario calculado','Error mensual acumulado','MAPE original Hoja5','Pares válidos','Verificación Hoja5']];header(summary,`A${start}:H${start}`);
 cats.forEach((c,i)=>{const r=start+1+i;const lo=5+30*i,hi=lo+29;const ref=`'${sheet}'!`;
 summary.getRange(`A${r}`).values=[[c]];summary.getRange(`B${r}:E${r}`).formulas=[[`=SUM(${ref}C${lo}:C${hi})`,`=SUM(${ref}D${lo}:D${hi})`,`=AVERAGE(${ref}E${lo}:E${hi})`,`=ABS((C${r}-B${r})/B${r})`]];
 summary.getRange(`F${r}`).values=[[input.original[key][i][2]]];summary.getRange(`G${r}`).formulas=[[`=COUNT(${ref}E${lo}:E${hi})`]];
 summary.getRange(`H${r}`).formulas=[[`=IF(AND(ABS(B${r}-${input.original[key][i][0]})<0.005,ABS(C${r}-${input.original[key][i][1]})<0.005,ROUND(D${r},4)=ROUND(F${r},4)),"Coincide","Revisar")`]];
 });
 const r=start+6;summary.getRange(`A${r}`).values=[['TOTAL / GLOBAL']];summary.getRange(`B${r}:G${r}`).formulas=[[`=SUM(B${start+1}:B${r-1})`,`=SUM(C${start+1}:C${r-1})`,`=AVERAGE('${sheet}'!E5:E154)`,`=ABS((C${r}-B${r})/B${r})`,`=AVERAGE(F${start+1}:F${r-1})`,`=SUM(G${start+1}:G${r-1})`]];
 summary.getRange(`B${start+1}:C${r}`).setNumberFormat('#,##0.00');summary.getRange(`D${start+1}:F${r}`).setNumberFormat('0.00%');summary.getRange(`A${r}:H${r}`).format.fill='#E6EEF5';summary.getRange(`A${r}:H${r}`).format.font.bold=true;
}
summary.getRange('A29:E29').values=[['Comparación global','Pretest','Postest','Cambio en puntos porcentuales','Lectura']];header(summary,'A29:E29');
summary.getRange('A30').values=[['MAPE diario']];summary.getRange('B30:E30').formulas=[['=D12','=D22','=(C30-B30)*100','=IF(C30<B30,"Menor error",IF(C30>B30,"Mayor error","Sin cambio"))']];
summary.getRange('A31').values=[['Promedio de errores mensuales']];summary.getRange('A31').format.wrapText=true;summary.getRange('A31:H31').format.rowHeight=38;
summary.getRange('B31:E31').formulas=[['=AVERAGE(E7:E11)','=AVERAGE(E17:E21)','=(C31-B31)*100','=IF(C31<B31,"Menor error",IF(C31>B31,"Mayor error","Sin cambio"))']];summary.getRange('B30:C31').setNumberFormat('0.00%');summary.getRange('D30:D31').setNumberFormat('0.00');
summary.getRange('A33:H33').merge();summary.getRange('A33').values=[['TOTAL en la columna E es el error del total del negocio; fila 31 es el promedio de los cinco errores mensuales. Son cálculos distintos.']];summary.getRange('A33:H33').format.wrapText=true;summary.getRange('A33:H33').format.rowHeight=32;
summary.freezePanes.freezeRows(6);
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:"'Resultados Hoja5'!A6:H22",include:'values,formulas',tableMaxRows:18,tableMaxCols:8,maxChars:2200})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!',options:{useRegex:true,maxResults:20},summary:'Errores'})).ndjson);
await fs.mkdir('C:/taully_demand_forecast/outputs/verificacion_hoja5',{recursive:true});
await (await SpreadsheetFile.exportXlsx(wb)).save('C:/taully_demand_forecast/outputs/verificacion_hoja5/Verificacion_Pretest_Postest_Hoja5.xlsx');
for(const [s,r,n] of [['Resultados Hoja5','A5:H22','resumen'],['Diario pretest','A1:E12','pre'],['Diario postest','A1:E12','post']]){
 const img=await wb.render({sheetName:s,range:r,scale:1.3,format:'png'});await fs.writeFile(new URL(`./${n}.png`,import.meta.url),new Uint8Array(await img.arrayBuffer()));
}
