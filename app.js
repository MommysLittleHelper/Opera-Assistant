const input=document.getElementById("input"),stats=document.getElementById("stats"),clear=document.getElementById("clear"),generate=document.getElementById("generate"),status=document.getElementById("status"),tool=document.getElementById("document"),format=document.getElementById("format"),help=document.getElementById("inputHelp");
let mode="заявка",blobUrl=null;
const fields=[["driver","Водитель"],["car","Автомобиль"],["plate","Госномер"],["passportSeries","Серия паспорта"],["passportNumber","Номер паспорта"],["issuedBy","Кем выдан"],["passportDate","Дата выдачи паспорта"],["phone","Телефон"],["recipient","Получатель"],["dt","ДТ"],["do","ДО"],["jdn","ЖДН / CMR"],["invoice","Инвойс"],["invoiceDate","Дата инвойса"],["ref","REF"],["places","Количество мест"],["weight","Вес брутто"]];
const NS="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
function clean(s){return String(s||"").replace(/\u00a0/g," ").replace(/[ \t]+/g," ").trim()}
function dateList(t){return [...t.matchAll(/\b\d{2}\.\d{2}\.\d{4}\b/g)].map(m=>m[0])}
function phone(t){const m=t.match(/(?:\+7|8)\s*\(?\d{3}\)?[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}/);return m?clean(m[0]):""}
function passport(t){const m=t.match(/\b(\d{4})\s+(\d{6})\b/);return m?{series:m[1],number:m[2]}:{series:"",number:""}}
function parse(t){
 const x=t.replace(/\r/g," ").replace(/\n/g," ").replace(/\s+/g," ").trim(), dates=dateList(x), pp=passport(x);
 const d={driver:"",car:"",plate:"",passportSeries:pp.series,passportNumber:pp.number,issuedBy:"",passportDate:"",phone:phone(x),recipient:"",dt:"",do:"",jdn:"",invoice:"",invoiceDate:"",ref:"",places:"",weight:""};
 // Сначала работаем по ЯВНЫМ подписям. Это основная логика версии 0.4.4.
 let m=x.match(/\bВодитель\s+([А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё-]+){1,2})(?=\s*,|\s+ПАСПОРТ\b)/i); if(m)d.driver=clean(m[1]);
 m=x.match(/\bАВТОМОБИЛЬ\s+(.+?)(?=\s+ГОСНОМЕР\b)/i); if(m)d.car=clean(m[1]);
 m=x.match(/\bГОСНОМЕР\s+([A-ZА-ЯЁ]\d{3}[A-ZА-ЯЁ]{2}\d{2,3})\b/i); if(m)d.plate=m[1];
 m=x.match(/\bПАСПОРТ\s+(\d{4})\s+(\d{6})\b/i); if(m){d.passportSeries=m[1];d.passportNumber=m[2]}
 m=x.match(/\bвыдан\s+(.+?)(?=\s+ГАЗ\s+АВТОМОБИЛЬ\b)/i); if(m)d.issuedBy=clean(m[1]);
 m=x.match(/\bдата\s+выдачи\s+паспорта\s+(\d{2}\.\d{2}\.\d{4})/i); if(m)d.passportDate=m[1]; else if(dates.length)d.passportDate=dates[0];
 m=x.match(/\bв\s+адрес:\s*(.+?)(?=\))/i); if(m)d.recipient=clean(m[1]);
 m=x.match(/\bДО\s+(?:№\s*)?([0-9]+)/i); if(m)d.do=m[1];
 m=x.match(/\bДТ\s+(?:№\s*)?([0-9]+\/[0-9]+\/[0-9]+)/i); if(m)d.dt=m[1];
 m=x.match(/\bЖДН\s+(?:№\s*)?([0-9]+)\s+от\s+(\d{2}\.\d{2}\.\d{4})/i); if(m)d.jdn=`№${m[1]} от ${m[2]}`;
 m=x.match(/\bИНВОЙС\s+(?:№\s*)?([0-9A-ZА-ЯЁ/-]+)\s+от\s+(\d{2}\.\d{2}\.\d{4})/i); if(m){d.invoice=m[1];d.invoiceDate=m[2]}
 m=x.match(/\bREF\s+([^\s,;]+)/i); if(m)d.ref=clean(m[1]);
 m=x.match(/\b(?:количество\s+мест|мест)\s*[:№#-]?\s*([0-9]+)/i); if(m)d.places=m[1];
 m=x.match(/\bвес\s+брутто\s*[:№#-]?\s*([0-9]+(?:[,.][0-9]+)?)\s*(?:КГ|кг)?/i); if(m)d.weight=clean(m[1]);
 const now=new Date(),pad=n=>String(n).padStart(2,"0"),today=`${pad(now.getDate())}.${pad(now.getMonth()+1)}.${now.getFullYear()}`; const expiry=new Date(now); expiry.setFullYear(expiry.getFullYear()+1);expiry.setMonth(expiry.getMonth()+1);expiry.setDate(expiry.getDate()+1); d.today=today;d.expiry=`${pad(expiry.getDate())}.${pad(expiry.getMonth()+1)}.${expiry.getFullYear()}`; return d;
}
function render(d){const panel=document.getElementById("parsedFields"),grid=document.getElementById("fieldsGrid");panel.hidden=false;grid.innerHTML=fields.map(([k,l])=>`<div class="field"><label>${l}</label><input data-key="${k}" value="${String(d[k]||"").replace(/"/g,"&quot;")}" class="${d[k]?"":"missing"}">${d[k]?"":"<small>не найдено — можно оставить пустым</small>"}</div>`).join("");grid.querySelectorAll("input").forEach(e=>e.oninput=()=>e.classList.toggle("missing",!e.value.trim()));stats.textContent="Данные распознаны — проверьте поля";generate.textContent="Сформировать документ"}
function values(){const d={};document.querySelectorAll("#fieldsGrid input").forEach(e=>d[e.dataset.key]=e.value.trim());return d}
function textNodes(root){return Array.from(root.getElementsByTagNameNS(NS,"t"))}
function setTextAt(paragraph,index,value){const ts=textNodes(paragraph);if(ts[index])ts[index].textContent=value||""}
function cellText(cell,value){const ts=textNodes(cell);if(ts.length)ts[ts.length-1].textContent=value||""}
async function fillDoc(d){const r=await fetch("Заявка и доверенность шаблон.docx",{cache:"no-store"});if(!r.ok)throw Error("Не найден Word-шаблон.");const zip=await JSZip.loadAsync(await r.arrayBuffer()),xml=await zip.file("word/document.xml").async("string"),doc=new DOMParser().parseFromString(xml,"application/xml");if(doc.getElementsByTagName("parsererror").length)throw Error("Не удалось прочитать Word-шаблон.");const ps=Array.from(doc.getElementsByTagNameNS(NS,"p")),t=i=>ps[i];
 setTextAt(t(2),1,d.car);setTextAt(t(2),4,d.plate);setTextAt(t(3),2,d.recipient);setTextAt(t(4),1,d.driver);setTextAt(t(5),1,d.passportSeries);setTextAt(t(5),4,d.passportNumber);setTextAt(t(5),7,d.issuedBy);setTextAt(t(5),9,d.passportDate);setTextAt(t(6),3,d.phone);setTextAt(t(19),5,d.dt);setTextAt(t(20),1,d.do);setTextAt(t(22),5,d.jdn);setTextAt(t(23),5,d.invoice);setTextAt(t(24),2,d.ref);setTextAt(t(25),3,d.places);setTextAt(t(26),2,d.weight);setTextAt(t(28),1,d.today);
 const tables=Array.from(doc.getElementsByTagNameNS(NS,"tbl")), rows=tbl=>Array.from(tbl.children).filter(n=>n.localName==="tr"),cells=row=>Array.from(row.children).filter(n=>n.localName==="tc");
 if(tables[0]){const c=cells(rows(tables[0])[1]);cellText(c[1],d.today);cellText(c[2],d.expiry);cellText(c[3],"Водитель "+(d.driver||""))}
 if(tables[1]){const rr=rows(tables[1]);let c=cells(rr[2]);cellText(c[4],d.driver);c=cells(rr[4]);cellText(c[1],d.passportSeries);cellText(c[3],d.passportNumber);c=cells(rr[5]);cellText(c[1],d.issuedBy);c=cells(rr[6]);cellText(c[1],d.passportDate);c=cells(rr[9]);cellText(c[1],"ДТ "+(d.dt||""))}
 const out=new XMLSerializer().serializeToString(doc);zip.file("word/document.xml",out);const buf=await zip.generateAsync({type:"blob",compression:"DEFLATE"});return{blob:buf,filename:`Заявка_и_доверенность_${(d.plate||"готово").replace(/[\/\\:*?"<>|]/g,"")}.docx`}}
async function downloadDoc(){const r=await fillDoc(values());if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(r.blob);status.innerHTML=`<div class="result"><strong>✓ Word-документ готов</strong><a class="download-link" href="${blobUrl}" download="${r.filename}">Скачать ${r.filename}</a></div>`}
function update(){generate.disabled=!input.value.trim();if(input.value.trim())stats.textContent="Текст готов к распознаванию";else stats.textContent="Данные не введены"}
tool.onchange=()=>{mode=tool.value;help.textContent=mode==="проработка"?"Вставьте диапазон из Excel.":"Порядок сведений не важен — Помощник распознает поля по структуре текста.";format.textContent=mode==="проработка"?"Excel (.xlsx)":"Word (.docx)";generate.textContent=mode==="проработка"?"Сформировать документ":"Распознать данные";input.value="";document.getElementById("parsedFields").hidden=true;update()};
input.oninput=update;clear.onclick=()=>{input.value="";document.getElementById("parsedFields").hidden=true;status.innerHTML="";update()};
generate.onclick=async()=>{generate.disabled=true;try{if(mode==="проработка")throw Error("Проработка в этой версии не изменялась. Для теста выберите «Заявка + Доверенность».");if(document.getElementById("parsedFields").hidden){render(parse(input.value));generate.disabled=false}else await downloadDoc()}catch(e){status.innerHTML=`<div class="error"><strong>Ошибка</strong><br>${e.message}</div>`;generate.disabled=false}};
