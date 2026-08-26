const input=document.getElementById("input"),stats=document.getElementById("stats"),clear=document.getElementById("clear"),generate=document.getElementById("generate"),status=document.getElementById("status"),tool=document.getElementById("document"),format=document.getElementById("format"),help=document.getElementById("inputHelp");
let mode="заявка",blobUrl=null;

const fields=[
["driver","Водитель"],["car","Автомобиль"],["plate","Госномер"],["passportSeries","Серия паспорта"],
["passportNumber","Номер паспорта"],["issuedBy","Кем выдан"],["passportDate","Дата выдачи паспорта"],
["phone","Телефон"],["recipient","Получатель"],["dt","ДТ"],["do","ДО"],["jdn","ЖДН / CMR"],
["invoice","Инвойс"],["invoiceDate","Дата инвойса"],["ref","REF"],["places","Количество мест"],
["weight","Вес брутто"]
];

function clean(s){return String(s||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim()}
function dateList(t){return [...t.matchAll(/\b\d{2}\.\d{2}\.\d{4}\b/g)].map(m=>m[0])}
function phone(t){let m=t.match(/(?:\+7|8)\s*\(?\d{3}\)?[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}/);return m?clean(m[0]):""}
function passport(t){let m=t.match(/\b(\d{4})\s*(?:№|N|#)?\s*(\d{6})\b/);return m?{series:m[1],number:m[2]}:{series:"",number:""}}
function plate(t){
 let m=t.match(/\b[A-ZА-Я]\d{3}[A-ZА-Я]{2}\d{2,3}\b/i); return m?m[0]:"";
}
function driver(t){
 let m=t.match(/(?:^|\s)водитель\s+([А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё-]+){1,2})(?=,|\s+(?:паспорт|ПАСПОРТ)|$)/i);
 if(m)return clean(m[1]);
 m=t.match(/(?:ФИО водителя[\/\s]*[^:]*[:\-]?\s*)([А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё-]+){1,2})/i);
 if(m)return clean(m[1]);
 let a=[...t.matchAll(/\b[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}(?:\s+[А-ЯЁ][а-яё]{4,})?\b/g)];
 let three=a.find(x=>x[0].split(/\s+/).length===3); return three?clean(three[0]):(a[0]?clean(a[0][0]):"");
}
function afterLabel(t,label,stops){
 let re=new RegExp(label+"\\s*[:№#-]?\\s*(.+?)(?=\\s+(?:"+stops.join("|")+")\\b|[.;]|$)","i"),m=t.match(re);return m?clean(m[1]):"";
}
function parse(t){
 const x=t.replace(/\r/g," ").replace(/\n/g," "), pp=passport(x), dates=dateList(x);
 const d={driver:driver(x),car:"",plate:plate(x),passportSeries:pp.series,passportNumber:pp.number,issuedBy:"",passportDate:"",
 phone:phone(x),recipient:"",dt:"",do:"",jdn:"",invoice:"",invoiceDate:"",ref:"",places:"",weight:""};
 let m=x.match(/(?:газ\s+)?(?:автомобиль|авто|марка)\s+([A-ZА-ЯЁ][A-ZА-ЯЁ0-9 _-]+?)(?=\s+(?:госномер|г\.н\.з\.|телефон)|,|$)/i);
 if(m)d.car=clean(m[1]);
 m=x.match(/в адрес получателя\s+(.+?)(?=\s+(?:паспорт|ПАСПОРТ|ФИО водителя)|,?\s*ФИО)/i);if(m)d.recipient=clean(m[1]);
 m=x.match(/паспорт.*?\bвыдан\s+(.+?)(?=\s+(?:ГАЗ|АВТОМОБИЛЬ|ГОСНОМЕР|ТЕЛЕФОН|ДО|ЖДН|ИНВОЙС|ДТ)|,?\s*$)/i);if(m)d.issuedBy=clean(m[1].replace(/,\s*$/,""));
 m=x.match(/(?:дата\s+выдачи\s+паспорта|дата\s+выдачи)\s*(?:паспорта)?\s*[:\-]?\s*(\d{2}\.\d{2}\.\d{4})/i);if(m)d.passportDate=m[1];else if(dates.length)d.passportDate=dates[0];
 m=x.match(/\bДТ\s*(?:№|N|#)?\s*([0-9/;,\s]+)/i);if(m)d.dt=clean(m[1]);
 m=x.match(/\bДО\s*(?:№|N|#)?\s*([0-9/;,\s]+)/i);if(m)d.do=clean(m[1]);
 m=x.match(/(?:ЖДН|CMR)\s*(?:№|N|#)?\s*([0-9A-ZА-ЯЁ/-]+)(?:\s+от\s+(\d{2}\.\d{2}\.\d{4}))?/i);if(m)d.jdn=clean(m[1])+(m[2]?" от "+m[2]:"");
 m=x.match(/инвойс\s*(?:№|N|#)?\s*([0-9A-ZА-ЯЁ/-]+)(?:\s+от\s+(\d{2}\.\d{2}\.\d{4}))?/i);if(m){d.invoice=clean(m[1]);d.invoiceDate=m[2]||""}
 m=x.match(/\bREF\s+([A-Z0-9-]+(?:\s*;\s*[A-Z0-9-]+)*)/i);if(m)d.ref=clean(m[1]);
 m=x.match(/количество\s+мест\s+([0-9]+)/i);if(m)d.places=m[1];
 m=x.match(/вес\s+брутто\s+(.+?)(?=\s+КГ\b|$)/i);if(m)d.weight=clean(m[1])+" КГ";
 const now=new Date(),pad=n=>String(n).padStart(2,"0"),today=`${pad(now.getDate())}.${pad(now.getMonth()+1)}.${now.getFullYear()}`;
 const expiry=new Date(now);expiry.setFullYear(expiry.getFullYear()+1);expiry.setMonth(expiry.getMonth()+1);expiry.setDate(expiry.getDate()+1);
 d.today=today;d.expiry=`${pad(expiry.getDate())}.${pad(expiry.getMonth()+1)}.${expiry.getFullYear()}`;
 return d;
}
function render(d){
 const panel=document.getElementById("parsedFields"),grid=document.getElementById("fieldsGrid");panel.hidden=false;
 grid.innerHTML=fields.map(([k,l])=>`<div class="field"><label>${l}</label><input data-key="${k}" value="${String(d[k]||"").replace(/"/g,"&quot;")}" class="${d[k]?"":"missing"}">${d[k]?"":"<small>не найдено — можно оставить пустым</small>"}</div>`).join("");
 grid.querySelectorAll("input").forEach(e=>e.oninput=()=>e.classList.toggle("missing",!e.value.trim()));
 stats.textContent="Данные распознаны — проверьте поля";generate.textContent="Сформировать документ";
}
function values(){let d={};document.querySelectorAll("#fieldsGrid input").forEach(e=>d[e.dataset.key]=e.value.trim());return d}

function setText(nodes,index,value){
 const ts=nodes[index]; if(ts) ts.textContent=value?(" "+value):"";
}
function allT(root){return [...root.querySelectorAll("w:t")]}
async function fillDoc(d){
 const r=await fetch("Заявка и доверенность шаблон.docx",{cache:"no-store"});if(!r.ok)throw Error("Не найден Word-шаблон.");
 const zip=await JSZip.loadAsync(await r.arrayBuffer()),xml=await zip.file("word/document.xml").async("string");
 const parser=new DOMParser(),doc=parser.parseFromString(xml,"application/xml"),ps=[...doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main","p")];
 const t=i=>allT(ps[i]);
 // Application page
 setText(t(2),1,d.car);setText(t(2),4,d.plate);setText(t(3),2,d.recipient);setText(t(4),1,d.driver);
 setText(t(5),1,d.passportSeries);setText(t(5),4,d.passportNumber);setText(t(5),7,d.issuedBy);setText(t(5),9,d.passportDate);
 setText(t(6),3,d.phone);setText(t(19),5,d.dt);setText(t(20),1,d.do);setText(t(22),5,d.jdn);
 setText(t(23),5,d.invoice);setText(t(24),2,d.ref);setText(t(25),3,d.places);setText(t(26),2,d.weight);setText(t(28),1,d.today);
 // Power of attorney table
 const tables=[...doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main","tbl")];
 const rows=tbl=>[...tbl.children].filter(n=>n.localName==="tr"), cells=row=>[...row.children].filter(n=>n.localName==="tc");
 if(tables[0]){
  let r=rows(tables[0])[1],c=cells(r); setCell(c,1,d.today,doc);setCell(c,2,d.expiry,doc);setCell(c,3,"Водитель "+(d.driver||""),doc);
 }
 if(tables[1]){
  let rr=rows(tables[1]);
  let c=cells(rr[2]);setCell(c,4,d.driver,doc);
  c=cells(rr[4]);setCell(c,1,d.passportSeries,doc);setCell(c,3,d.passportNumber,doc);
  c=cells(rr[5]);setCell(c,1,d.issuedBy,doc);
  c=cells(rr[6]);setCell(c,1,d.passportDate,doc);
  c=cells(rr[9]);setCell(c,1,"ДТ "+(d.dt||""),doc);
 }
 if(tables[2]){let r=rows(tables[2])[1],c=cells(r);setCell(c,3,d.places,doc)}
 const out=new XMLSerializer().serializeToString(doc);zip.file("word/document.xml",out);
 const buf=await zip.generateAsync({type:"blob",compression:"DEFLATE"});
 return {blob:buf,filename:`Заявка_и_доверенность_${(d.plate||"готово").replace(/[\/\\:*?"<>|]/g,"")}.docx`};
}
function setCell(cells,index,value,doc){
 if(!cells[index])return;
 const ts=[...cells[index].getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main","t")];
 if(ts.length){ts[ts.length-1].textContent=value||""}
}
async function downloadDoc(){
 const d=values(),r=await fillDoc(d);if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(r.blob);
 status.innerHTML=`<div class="result"><strong>✓ Word-документ готов</strong><a class="download-link" href="${blobUrl}" download="${r.filename}">Скачать ${r.filename}</a></div>`;
}
function update(){
 if(!input.value.trim()){stats.textContent="Данные не введены";generate.disabled=true;return}
 generate.disabled=false;stats.textContent="Текст готов к распознаванию";
}
tool.onchange=()=>{mode=tool.value;if(mode==="проработка"){help.textContent="Вставьте диапазон из Excel.";format.textContent="Excel (.xlsx)";generate.textContent="Сформировать документ"}else{help.textContent="Порядок сведений не важен — Помощник распознает поля по структуре текста.";format.textContent="Word (.docx)";generate.textContent="Распознать данные"}input.value="";document.getElementById("parsedFields").hidden=true;update()};
input.oninput=update;clear.onclick=()=>{input.value="";document.getElementById("parsedFields").hidden=true;status.innerHTML="";update()};
generate.onclick=async()=>{
 generate.disabled=true;
 try{
  if(mode==="проработка"){throw Error("Для этого патча оставьте документ «Заявка + Доверенность». Проработка не изменялась.")}
  const panel=document.getElementById("parsedFields");
  if(panel.hidden){render(parse(input.value));generate.disabled=false;return}
  await downloadDoc();
 }catch(e){status.innerHTML=`<div class="error"><strong>Ошибка</strong><br>${e.message}</div>`;generate.disabled=false}
};
