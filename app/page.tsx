"use client";
import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import QRCode from "qrcode";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

type Row = Record<string, string | number | boolean | null>;
const mm=(n:number)=>n*2.834645669;
const presets=[{name:"11 × 11 cm Etikette",w:110,h:110},{name:"10 × 10 cm Etikette",w:100,h:100},{name:"10 × 5 cm Etikette",w:100,h:50},{name:"7 × 5 cm Etikette",w:70,h:50},{name:"A4",w:210,h:297},{name:"Benutzerdefiniert",w:70,h:55}];

export default function Home(){
 const [rows,setRows]=useState<Row[]>([]),[headers,setHeaders]=useState<string[]>([]),[qrCol,setQrCol]=useState(""),[labelPreset,setLabelPreset]=useState("11 × 11 cm Etikette"),[labelW,setLabelW]=useState(110),[labelH,setLabelH]=useState(110),[cols,setCols]=useState(1),[showName,setShowName]=useState(true),[nameCol,setNameCol]=useState(""),[showLocation,setShowLocation]=useState(true),[locationCol,setLocationCol]=useState(""),[qrSize,setQrSize]=useState(65),[qrTop,setQrTop]=useState(8),[qrXOffset,setQrXOffset]=useState(0),[qrYOffset,setQrYOffset]=useState(0),[textXOffset,setTextXOffset]=useState(0),[textYOffset,setTextYOffset]=useState(0),[titleSize,setTitleSize]=useState(9),[detailSize,setDetailSize]=useState(7.5),[textGap,setTextGap]=useState(6),[textAlign,setTextAlign]=useState<"left"|"center"|"right">("center"),[boldTitle,setBoldTitle]=useState(true),[busy,setBusy]=useState(false),[message,setMessage]=useState(""),[previewQr,setPreviewQr]=useState(""),[dragging,setDragging]=useState(false);
 const fileInput=useRef<HTMLInputElement>(null);
 const preview=useMemo(()=>rows.slice(0,6),[rows]),firstPreview=preview[0];

 useEffect(()=>{let cancelled=false;if(!firstPreview||!qrCol){setPreviewQr("");return;}const value=String(firstPreview[qrCol]??"");if(!value){setPreviewQr("");return;}QRCode.toDataURL(value,{width:500,margin:1,errorCorrectionLevel:"M"}).then(url=>{if(!cancelled)setPreviewQr(url)}).catch(()=>{if(!cancelled)setPreviewQr("")});return()=>{cancelled=true}},[firstPreview,qrCol]);

 function selectPreset(value:string){setLabelPreset(value);const p=presets.find(x=>x.name===value);if(p){setLabelW(p.w);setLabelH(p.h);setCols(1)}}
 async function load(file:File){const valid=/\.(xlsx|xls|csv)$/i.test(file.name);if(!valid){setMessage("Bitte eine Excel- oder CSV-Datei auswählen.");return}try{const data=await file.arrayBuffer();const wb=XLSX.read(data,{type:"array"});const sheet=wb.Sheets[wb.SheetNames[0]];const json=XLSX.utils.sheet_to_json<Row>(sheet,{defval:""});const hs=json.length?Object.keys(json[0]):[];setRows(json);setHeaders(hs);setQrCol(hs[0]||"");setNameCol(hs[1]||"");setLocationCol(hs[2]||"");setMessage(json.length+" Datensätze geladen: "+file.name)}catch{setMessage("Excel-Datei konnte nicht gelesen werden.")}}
 function onDrop(e:React.DragEvent<HTMLLabelElement>){e.preventDefault();e.stopPropagation();setDragging(false);const file=e.dataTransfer.files?.[0];if(file)load(file)}
 function onDragOver(e:React.DragEvent<HTMLLabelElement>){e.preventDefault();e.stopPropagation();e.dataTransfer.dropEffect="copy";setDragging(true)}
 function onDragLeave(e:React.DragEvent<HTMLLabelElement>){e.preventDefault();e.stopPropagation();setDragging(false)}

 function textX(x:number,w:number,align:"left"|"center"|"right",text:string,font:any,size:number){const tw=font.widthOfTextAtSize(text,size);if(align==="center")return x+(w-tw)/2;if(align==="right")return x+w-tw-mm(5);return x+mm(5)}
 async function makePdf(){
   if(!rows.length||!qrCol){setMessage("Bitte Excel-Datei und QR-Spalte auswählen.");return}
   const pageW=mm(210),pageH=mm(297),margin=mm(10),gap=mm(4),perRow=Math.max(1,cols);
   if(labelW*perRow+4*(perRow-1)>190){setMessage("Die Etiketten sind zu breit für A4. Bitte Breite oder Anzahl pro Zeile reduzieren.");return}
   if(labelW>190||labelH>277){setMessage("Das Etikett ist größer als der bedruckbare A4-Bereich. Bitte ein kleineres Format wählen.");return}
   if(qrSize>Math.min(labelW-10,labelH-25)){setMessage("Der QR-Code ist für dieses Etikett zu groß.");return}
   setBusy(true);setMessage("PDF wird erstellt …");
   try{
     const pdf=await PDFDocument.create(),regular=await pdf.embedFont(StandardFonts.Helvetica),bold=await pdf.embedFont(StandardFonts.HelveticaBold);
     const rowsPerPage=Math.max(1,Math.floor((pageH-margin*2+gap)/(mm(labelH)+gap))),maxPerPage=perRow*rowsPerPage;
     for(let start=0;start<rows.length;start+=maxPerPage){
       const page=pdf.addPage([pageW,pageH]);
       for(let i=0;i<Math.min(maxPerPage,rows.length-start);i++){
         const row=rows[start+i],c=i%perRow,r=Math.floor(i/perRow),x=margin+c*(mm(labelW)+gap),y=pageH-margin-(r+1)*mm(labelH)-r*gap,value=String(row[qrCol]??"");if(!value)continue;
         const png=await QRCode.toDataURL(value,{width:500,margin:1,errorCorrectionLevel:"M"}),img=await pdf.embedPng(png),qr=mm(qrSize);
         page.drawRectangle({x,y,width:mm(labelW),height:mm(labelH),borderColor:rgb(.78,.84,.87),borderWidth:.7});
         const drawQrX=x+(mm(labelW)-qr)/2+mm(qrXOffset),drawQrY=y+mm(labelH)-qr-mm(qrTop)+mm(qrYOffset);
         page.drawImage(img,{x:drawQrX,y:drawQrY,width:qr,height:qr});
         let ty=drawQrY-mm(textGap)-titleSize;
         const titleFont=boldTitle?bold:regular,title=value.slice(0,60);
         page.drawText(title,{x:textX(x,mm(labelW),textAlign,title,titleFont,titleSize)+mm(textXOffset),y:ty+mm(textYOffset),size:titleSize,font:titleFont,maxWidth:mm(labelW)-mm(10)});
         ty-=mm(6)+titleSize;
         if(showName&&nameCol){const t=String(row[nameCol]??"").slice(0,60);page.drawText(t,{x:textX(x,mm(labelW),textAlign,t,regular,detailSize)+mm(textXOffset),y:ty+mm(textYOffset),size:detailSize,font:regular,maxWidth:mm(labelW)-mm(10)});ty-=mm(5)+detailSize}
         if(showLocation&&locationCol){const t=String(row[locationCol]??"").slice(0,60);page.drawText(t,{x:textX(x,mm(labelW),textAlign,t,regular,detailSize),y:ty,size:detailSize,font:regular,maxWidth:mm(labelW)-mm(10)})}
       }
     }
     const bytes=await pdf.save(),safeBytes=new Uint8Array(bytes),blob=new Blob([safeBytes.buffer],{type:"application/pdf"}),url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download="QR-Generator47-Etiketten.pdf";a.click();URL.revokeObjectURL(url);setMessage("PDF fertig – "+rows.length+" QR-Etiketten.")
   }catch(e){console.error(e);setMessage("PDF-Erstellung fehlgeschlagen.")}finally{setBusy(false)}
 }

 return <main className="min-h-screen"><header className="border-b bg-white"><div className="mx-auto max-w-6xl px-6 py-6"><div className="text-2xl font-bold">QR Generator<span className="text-sky-600">47</span></div><div className="text-sm text-slate-500">Excel → QR-Codes → individuell gestaltbare Etiketten</div></div></header>
 <div className="mx-auto grid max-w-6xl gap-6 px-6 py-8 lg:grid-cols-[1fr_400px]"><section className="space-y-6">
 <div className="card p-7"><h1 className="text-2xl font-bold">Excel-Datei importieren</h1><p className="mt-1 text-slate-500">Jede Zeile wird zu einem QR-Code-Etikett.</p>
 <label onDrop={onDrop} onDragOver={onDragOver} onDragEnter={onDragOver} onDragLeave={onDragLeave} onClick={()=>fileInput.current?.click()} className={"mt-6 flex cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed px-6 py-12 text-center transition "+(dragging?"border-sky-500 bg-sky-100 scale-[1.01]":"border-sky-200 bg-sky-50 hover:bg-sky-100")}><div className="text-4xl">{dragging?"📥":"📊"}</div><div className="mt-3 font-semibold">{dragging?"Datei hier loslassen":"Excel hier ablegen oder Datei auswählen"}</div><div className="mt-1 text-sm text-slate-500">.xlsx / .xls / .csv</div><input ref={fileInput} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={e=>e.target.files?.[0]&&load(e.target.files[0])}/></label>{message&&<div className="mt-4 rounded-xl bg-slate-50 px-4 py-3 text-sm">{message}</div>}</div>
 {rows.length>0&&<div className="card p-7"><h2 className="text-xl font-bold">Daten & Etikett</h2>
 <div className="mt-5 grid gap-4 sm:grid-cols-2">
 <label className="text-sm font-semibold">QR-Code-Inhalt<select value={qrCol} onChange={e=>setQrCol(e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal">{headers.map(h=><option key={h}>{h}</option>)}</select></label>
 <label className="text-sm font-semibold">Etikettenformat<select value={labelPreset} onChange={e=>selectPreset(e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal">{presets.map(p=><option key={p.name}>{p.name}</option>)}</select></label>
 <label className="text-sm font-semibold">Etiketten pro Zeile<select value={cols} onChange={e=>setCols(+e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal"><option value={1}>1</option><option value={2}>2</option><option value={3}>3</option></select></label>
 <label className="text-sm font-semibold">Breite (mm)<input type="number" min="20" value={labelW} onChange={e=>{setLabelPreset("Benutzerdefiniert");setLabelW(+e.target.value)}} className="mt-2 w-full rounded-xl border p-3 font-normal"/></label>
 <label className="text-sm font-semibold">Höhe (mm)<input type="number" min="25" value={labelH} onChange={e=>{setLabelPreset("Benutzerdefiniert");setLabelH(+e.target.value)}} className="mt-2 w-full rounded-xl border p-3 font-normal"/></label>
 <div className="sm:col-span-2 mt-2 rounded-2xl border bg-slate-50 p-5"><h3 className="font-bold">QR-Code & Text individuell</h3><div className="mt-4 grid gap-4 sm:grid-cols-2">
 <label className="text-sm font-semibold">QR-Größe (mm)<input type="number" min="15" max="180" value={qrSize} onChange={e=>setQrSize(+e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal"/></label>
 <label className="text-sm font-semibold">QR oben/unten (mm)<input type="number" min="-50" max="50" value={qrTop} onChange={e=>setQrTop(+e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal"/><span className="mt-1 block text-xs font-normal text-slate-500">Positiver Wert = weiter nach unten</span></label>
 <label className="text-sm font-semibold">QR links/rechts (mm)<input type="number" min="-50" max="50" value={qrXOffset} onChange={e=>setQrXOffset(+e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal"/><span className="mt-1 block text-xs font-normal text-slate-500">− links · + rechts</span></label>
 <label className="text-sm font-semibold">QR vertikal (mm)<input type="number" min="-50" max="50" value={qrYOffset} onChange={e=>setQrYOffset(+e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal"/><span className="mt-1 block text-xs font-normal text-slate-500">− oben · + unten</span></label>
 <label className="text-sm font-semibold">Text links/rechts (mm)<input type="number" min="-50" max="50" value={textXOffset} onChange={e=>setTextXOffset(+e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal"/><span className="mt-1 block text-xs font-normal text-slate-500">− links · + rechts</span></label>
 <label className="text-sm font-semibold">Text oben/unten (mm)<input type="number" min="-50" max="50" value={textYOffset} onChange={e=>setTextYOffset(+e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal"/><span className="mt-1 block text-xs font-normal text-slate-500">− oben · + unten</span></label>
 <label className="text-sm font-semibold">Haupttext Größe (pt)<input type="number" min="4" max="30" step="0.5" value={titleSize} onChange={e=>setTitleSize(+e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal"/></label>
 <label className="text-sm font-semibold">Name/Details Größe (pt)<input type="number" min="4" max="24" step="0.5" value={detailSize} onChange={e=>setDetailSize(+e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal"/></label>
 <label className="text-sm font-semibold">Abstand QR → Text (mm)<input type="number" min="0" max="30" value={textGap} onChange={e=>setTextGap(+e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal"/></label>
 <label className="text-sm font-semibold">Textausrichtung<select value={textAlign} onChange={e=>setTextAlign(e.target.value as "left"|"center"|"right")} className="mt-2 w-full rounded-xl border p-3 font-normal"><option value="left">Links</option><option value="center">Zentriert</option><option value="right">Rechts</option></select></label>
 <label className="flex items-center gap-3 text-sm font-semibold"><input type="checkbox" checked={boldTitle} onChange={e=>setBoldTitle(e.target.checked)}/> Haupttext fett</label>
 </div></div>
 <label className="flex items-center gap-3 text-sm font-semibold"><input type="checkbox" checked={showName} onChange={e=>setShowName(e.target.checked)}/> Name/Bezeichnung anzeigen</label>
 <label className="text-sm font-semibold">Name-Spalte<select value={nameCol} onChange={e=>setNameCol(e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal">{headers.map(h=><option key={h}>{h}</option>)}</select></label>
 <label className="flex items-center gap-3 text-sm font-semibold"><input type="checkbox" checked={showLocation} onChange={e=>setShowLocation(e.target.checked)}/> Standort anzeigen</label>
 <label className="text-sm font-semibold">Standort-Spalte<select value={locationCol} onChange={e=>setLocationCol(e.target.value)} className="mt-2 w-full rounded-xl border p-3 font-normal">{headers.map(h=><option key={h}>{h}</option>)}</select></label>
 </div><button disabled={busy} onClick={makePdf} className="mt-7 w-full rounded-xl bg-sky-600 px-5 py-4 font-bold text-white hover:bg-sky-700 disabled:opacity-50">{busy?"PDF wird erstellt …":"📄 Druckfertiges PDF erstellen"}</button></div>}</section>
 <aside className="card h-fit p-6 lg:sticky lg:top-6"><div className="flex items-center justify-between"><div><h2 className="text-xl font-bold">Druckvorschau</h2><p className="text-sm text-slate-500">QR oben, Text darunter</p></div><span className="rounded-full bg-sky-50 px-3 py-1 text-xs font-semibold text-sky-700">{labelW} × {labelH} mm</span></div>
 <div className="mt-5 rounded-2xl bg-slate-100 p-5"><div className="mx-auto overflow-hidden rounded-lg border bg-white p-3 shadow-sm" style={{width:"min(100%, 300px)",aspectRatio:labelW+"/"+labelH}}>{firstPreview&&qrCol?<div className={"flex h-full flex-col "+(textAlign==="center"?"items-center":textAlign==="right"?"items-end":"items-start")+" text-"+textAlign}>{previewQr?<img src={previewQr} alt="QR Vorschau" style={{width:Math.min(220,Math.max(50,qrSize*2.2)),marginTop:Math.max(0,qrTop/2+qrYOffset/2),transform:`translateX(${qrXOffset*2}px)`}} className="h-auto object-contain"/>:<div className="text-sm text-slate-400">QR wird geladen …</div>}<div className="w-full break-words text-sm" style={{marginTop:textGap*2.834645669/2,fontSize:titleSize,fontWeight:boldTitle?700:400,transform:`translate(${textXOffset*2}px, ${-textYOffset*2}px)`}}>{String(firstPreview[qrCol]??"")}</div>{showName&&nameCol&&<div className="w-full break-words text-slate-600" style={{fontSize:detailSize}}>{String(firstPreview[nameCol]??"")}</div>}{showLocation&&locationCol&&<div className="w-full break-words text-slate-500" style={{fontSize:detailSize}}>{String(firstPreview[locationCol]??"")}</div>}</div>:<div className="flex h-full items-center justify-center text-sm text-slate-400">Excel-Datei hochladen</div>}</div></div>
 {preview.length>0&&<div className="mt-5"><div className="mb-2 text-sm font-semibold">Weitere Datensätze</div><div className="space-y-2">{preview.slice(1).map((row,i)=><div key={i} className="rounded-xl border p-3 text-sm"><div className="font-bold">{String(row[qrCol]??"")}</div>{showName&&nameCol&&<div className="text-slate-600">{String(row[nameCol]??"")}</div>}{showLocation&&locationCol&&<div className="text-slate-500">{String(row[locationCol]??"")}</div>}</div>)}</div></div>}
 <div className="mt-6 rounded-xl bg-slate-50 p-4 text-xs text-slate-500">Die Excel-Daten werden vollständig im Browser verarbeitet. Es ist kein Backend und keine Datenbank nötig.</div></aside></div></main>
}