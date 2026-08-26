const input=document.getElementById("input"),refreshParsed=document.getElementById("refreshParsed"),stats=document.getElementById("stats"),clear=document.getElementById("clear"),generate=document.getElementById("generate"),status=document.getElementById("status"),tool=document.getElementById("document"),format=document.getElementById("format"),help=document.getElementById("inputHelp");
let mode="заявка",blobUrl=null;
const fields=[["driver","Водитель"],["car","Автомобиль"],["plate","Госномер"],["passportSeries","Серия паспорта"],["passportNumber","Номер паспорта"],["issuedBy","Кем выдан"],["passportDate","Дата выдачи паспорта"],["phone","Телефон"],["recipient","Получатель"],["dt","ДТ"],["do","ДО"],["jdn","ЖДН / CMR"],["invoice","Инвойс"],["invoiceDate","Дата инвойса"],["ref","REF"],["places","Количество мест"],["weight","Вес брутто"]];
const NS="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
function clean(s){return String(s||"").replace(/\u00a0/g," ").replace(/[ \t]+/g," ").trim()}
function normalizeText(s){return String(s||"").replace(/\r/g," ").replace(/\n/g," ").replace(/\u00a0/g," ").replace(/\s+/g," ").trim()}
function dateList(t){return [...String(t||"").matchAll(/\b\d{2}\.\d{2}\.\d{4}\b/g)].map(m=>m[0])}
function phone(t){const m=String(t||"").match(/(?:\+7|8)\s*\(?\d{3}\)?[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}/);return m?clean(m[0]):""}
const n=v=>String(v??"").replace(/\uFEFF/g,"").replace(/\u00a0/g," ").trim().toLowerCase().replace(/\s+/g," ");
const parseTable=t=>t.replace(/\r\n/g,"\n").replace(/\r/g,"\n").split("\n").map(x=>x.split("\t").map(y=>y.trim())).filter(r=>r.some(x=>x!==""));
const aliases={"номер ктк":["номер ктк","ктк"],ref:["ref","реф","букинг"],статус:["статус","клиент"],"фактический объем, м3":["фактический объем, м3","фактический объем м3","объем м3","объем, м3"],"вес, кг":["вес, кг","вес кг","вес"],"кол-во мест":["кол-во мест","количество мест","мест"]};
function col(headers,key,dict=aliases){let h=headers.map(n),o=(dict[key]||[key]).map(n),i=h.findIndex(x=>o.includes(x));return i>=0?i:h.findIndex(x=>x&&o.some(y=>x.includes(y)||y.includes(x)))}
function total(r){let s=r.map(n).filter(Boolean).join(" ");return !s||/итого|всего|total/i.test(s)||(!n(r[0])&&!n(r[10]))}

const Proработка={update(){let r=parseTable(input.value);if(!r.length){stats.textContent="Данные не введены";generate.disabled=true;return}let req=["номер ктк","ref","статус","фактический объем, м3","вес, кг","кол-во мест"],miss=req.filter(k=>col(r[0],k)<0),data=r.slice(1).filter(x=>!total(x));stats.textContent=miss.length?`${data.length} строк · не найдены: ${miss.join(", ")}`:`${data.length} записей · ${r[0].length} столбцов · все необходимые поля найдены`;generate.disabled=!!miss.length||!data.length},async build(){let r=parseTable(input.value),req=["номер ктк","ref","статус","фактический объем, м3","вес, кг","кол-во мест"],src={};req.forEach(k=>src[k]=col(r[0],k));if(Object.values(src).some(x=>x<0))throw Error("Не найдены необходимые поля");let data=r.slice(1).filter(x=>!total(x)),res=await fetch("проработка.xlsx",{cache:"no-store"});if(!res.ok)throw Error("Не найден шаблон «проработка.xlsx».");let wb=XLSX.read(new Uint8Array(await res.arrayBuffer()),{type:"array",cellStyles:true}),sh=wb.Sheets[wb.SheetNames[0]],rg=XLSX.utils.decode_range(sh["!ref"]),heads=[];for(let c=rg.s.c;c<=rg.e.c;c++)heads.push(sh[XLSX.utils.encode_cell({r:rg.s.r,c})]?.v??"");let dst={};req.forEach(k=>dst[k]=col(heads,k));if(Object.values(dst).some(x=>x<0))throw Error("В шаблоне не найдена необходимая колонка");data.forEach((row,i)=>req.forEach(k=>{let v=String(row[src[k]]??"").trim(),a=XLSX.utils.encode_cell({r:rg.s.r+1+i,c:dst[k]});sh[a]=/^-?\d+(?:[.,]\d+)?$/.test(v.replace(/\s/g,""))?{t:"n",v:Number(v.replace(/\s/g,"").replace(",","."))}:{t:"s",v}}));let out=XLSX.write(wb,{bookType:"xlsx",type:"array",cellStyles:true}),ktk=String(data[0]?.[src["номер ктк"]]||"готово").replace(/\//g,"").replace(/[\\:*?"<>|]/g,"");return{blob:new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),filename:`Проработка_${ktk||"готово"}.xlsx`,text:`Сформировано записей: ${data.length}`}}};
function parse(t){
  const x=normalizeText(t);
  const low=x.toLocaleLowerCase("ru-RU");
  const d={driver:"",car:"",plate:"",passportSeries:"",passportNumber:"",issuedBy:"",passportDate:"",phone:"",recipient:"",dt:"",do:"",jdn:"",invoice:"",invoiceDate:"",ref:"",places:"",weight:""};

  // Значение после явной метки до следующей известной метки.
  const between=(labels,next=[])=>{
    let best=-1,bestLen=0;
    for(const label of labels){
      const i=low.indexOf(label.toLocaleLowerCase("ru-RU"));
      if(i>=0 && (best<0 || i<best)){best=i;bestLen=label.length;}
    }
    if(best<0)return "";
    let s=best+bestLen;
    while(/[ \t:№#.,-]/.test(x[s]||""))s++;
    let e=x.length;
    for(const n of next){
      const j=low.indexOf(n.toLocaleLowerCase("ru-RU"),s);
      if(j>=0 && j<e)e=j;
    }
    return clean(x.slice(s,e).replace(/[;,]+$/,""));
  };

  // Основные поля — сначала по явным меткам.
  d.driver=between(["водитель","фио водителя"],["паспорт"]).replace(/[;,]+$/g,"").trim();
  d.car=between(["автомобиль"],["госномер","г.н.з.","телефон","тел."]).replace(/^газ\s+/i,"").trim();
  d.plate=between(["госномер","г.н.з."],["телефон","тел.","до","ждн","cmr","инвойс","дт"]);
  d.plate=d.plate.match(/[А-ЯЁA-Z]\s*\d{3}\s*[А-ЯЁA-Z]{2}\s*\d{2,3}/i)?.[0]?.replace(/\s+/g,"")||d.plate.replace(/[^\dA-ZА-ЯЁ]/gi,"");

  const pass=between(["паспорт"],["дата выдачи паспорта","выдан","автомобиль","газ автомобиль","госномер","телефон"]);
  let m=pass.match(/\b(\d{4})\s+(?:№\s*)?(\d{6})\b/);
  if(m){d.passportSeries=m[1];d.passportNumber=m[2]}
  else {
    m=x.match(/\bПАСПОРТ\s*:?\s*(\d{4})\s+(?:№\s*)?(\d{6})\b/i);
    if(m){d.passportSeries=m[1];d.passportNumber=m[2]}
  }
  d.passportDate=between(["дата выдачи паспорта"],["выдан","автомобиль","газ автомобиль","госномер","телефон","тел."]).replace(/[;,]+$/g,"").trim();
  if(!d.passportDate){m=x.match(/дата выдачи паспорта\s*[:№#-]?\s*(\d{2}\.\d{2}\.\d{4})/i);if(m)d.passportDate=m[1]}

  d.issuedBy=between(["выдан"],["дата выдачи паспорта","газ автомобиль","автомобиль","госномер","телефон","тел.","до","ждн","cmr","инвойс","дт"]).replace(/[;,]+$/g,"").trim();
  d.phone=phone(x);

  d.do=between(["до"],["ждн","cmr","инвойс","дт"]);
  m=x.match(/(?:^|\s)ДО\s*(?:№\s*)?([0-9]{4,})\b/i);if(m)d.do=m[1];

  d.jdn=between(["ждн","cmr(ждн)","cmr"],["инвойс","дт"]);
  m=x.match(/(?:ЖДН|CMR(?:\(ЖДН\))?)\s*(?:№\s*)*([0-9]+)\s+от\s+(\d{2}\.\d{2}\.\d{4})/i);
  if(m)d.jdn=`№${m[1]} от ${m[2]}`;
  else {m=x.match(/(?:ЖДН|CMR(?:\(ЖДН\))?)\s*(?:№\s*)*([0-9]+)/i);if(m)d.jdn=`№${m[1]}`}

  // Получатель:
  // Если есть «П/П», это самый надёжный ориентир: после него идёт
  // именно получатель. Берём его до закрывающей скобки/запятой/ИНВОЙСА.
  let recipient="";
  let rm=x.match(/П\/П\s+(.+?)(?=\)|\s+(?:ИНВОЙС\b|ДТ\b|ДО\b|ЖДН\b|CMR\b)|$)/i);
  if(rm) {
    recipient=clean(rm[1]).replace(/[\s),.;:]+$/g,"").trim();
  }

  // Запасной вариант — если П/П отсутствует, разбираем «в адрес: ...».
  if(!recipient){
    rm=x.match(/(?:в\s+адрес\s*:\s*|в\s+адрес\s+получателя\s+)(.+?)(?=\s*(?:\)|,)?\s*(?:ИНВОЙС\b|ДТ\b|ДО\b|ЖДН\b|CMR\b)|$)/i);
    if(rm){
      recipient=clean(rm[1])
        .replace(/^АО\s*[«"“]?ЛОГИСТИКА[-–—\s]+ТЕРМИНАЛ[»"”]?\s*/i,"")
        .replace(/^П\/П\s*/i,"")
        .replace(/[\s),.;:]+$/g,"")
        .trim();
    }
  }

  // Если в результате осталась служебная часть перед ООО/ИП, отбрасываем её.
  recipient=recipient
    .replace(/^АО\s*[«"“]?ЛОГИСТИКА[-–—\s]+ТЕРМИНАЛ[»"”]?\s*/i,"")
    .replace(/^П\/П\s*/i,"")
    .replace(/[\s),.;:]+$/g,"")
    .trim();

  d.recipient=recipient;

  d.invoice=between(["инвойс"],["дт"]);
  m=x.match(/ИНВОЙС\s*(?:№\s*)?([0-9A-ZА-ЯЁ/-]+)(?:\s+от\s+(\d{2}\.\d{2}\.\d{4}))?/i);
  if(m){d.invoice=m[1];d.invoiceDate=m[2]||""}

  m=x.match(/(?:^|\s)ДТ\s*(?:№\s*)?([0-9]{5,}\/[0-9]{5,}\/[0-9]{5,})/i);if(m)d.dt=m[1];
  d.ref=between(["ref"],["количество мест","кол-во мест","место","вес"]);
  m=x.match(/\bREF\s+([^\s,;]+)/i);if(m)d.ref=m[1];

  // Места: распознаём и «Количество мест 1», и «Кол-во мест 1», и «1 место».
  m=x.match(/(?:^|\s)(\d+)\s+мест(?:о|а)?(?=\s|$)/i);
  if(m)d.places=m[1];
  if(!d.places){m=x.match(/(?:количество\s+мест|кол-?во\s+мест)\s*[:№#-]?\s*(\d+(?:[.,]\d+)?)(?=\s|$)/i);if(m)d.places=m[1]}
  if(!d.places){m=x.match(/(?:^|\s)мест(?:о|а)?\s*[:№#-]?\s*(\d+)(?=\s|$)/i);if(m)d.places=m[1]}
  d.places=d.places.replace(",",".").match(/\d+(?:\.\d+)?/)?.[0]||"";

  // Вес: «214 кг», «вес брутто 214 кг КГ», «Вес 214 кг».
  m=x.match(/(?:вес\s*(?:брутто|нетто)?\s*[:№#-]?\s*)(\d+(?:[.,]\d+)?)\s*(?:кг|килограмм(?:а|ов)?)?/i);
  if(m)d.weight=`${m[1].replace(",",".")} кг`;
  if(!d.weight){m=x.match(/(?:^|\s)(\d+(?:[.,]\d+)?)\s*(?:кг|килограмм(?:а|ов)?)(?=\s|$)/i);if(m)d.weight=`${m[1].replace(",",".")} кг`}

  const now=new Date();
  const pad=n=>String(n).padStart(2,"0");
  const today=`${pad(now.getDate())}.${pad(now.getMonth()+1)}.${now.getFullYear()}`;
  // По правилу пользователя: сегодня + 1 год + 1 месяц + 1 день.
  const expiry=new Date(now.getTime());
  expiry.setFullYear(expiry.getFullYear()+1);
  expiry.setMonth(expiry.getMonth()+1);
  expiry.setDate(expiry.getDate()+1);
  d.today=today;
  d.expiry=`${pad(expiry.getDate())}.${pad(expiry.getMonth()+1)}.${expiry.getFullYear()}`;
  return d;
}
function render(d){const panel=document.getElementById("parsedFields"),grid=document.getElementById("fieldsGrid");panel.hidden=false;grid.innerHTML=fields.map(([k,l])=>`<div class="field"><label>${l}</label><input data-key="${k}" value="${String(d[k]||"").replace(/"/g,"&quot;")}" class="${d[k]?"":"missing"}">${d[k]?"":"<small>не найдено — можно оставить пустым</small>"}</div>`).join("");grid.querySelectorAll("input").forEach(e=>e.oninput=()=>e.classList.toggle("missing",!e.value.trim()));if(refreshParsed)refreshParsed.disabled=false;stats.textContent="Данные распознаны — проверьте поля";generate.textContent="Сформировать документ"}
function values(){const d={};document.querySelectorAll("#fieldsGrid input").forEach(e=>d[e.dataset.key]=e.value.trim());const parsed=parse(input.value);d.today=parsed.today;d.expiry=parsed.expiry;return d}
function textNodes(root){return Array.from(root.getElementsByTagNameNS(NS,"t"))}
function setNodeText(root,value){
  const ts=textNodes(root);
  if(!ts.length)return;
  ts[0].textContent=value||"";
  for(let i=1;i<ts.length;i++)ts[i].textContent="";
}
function setTextAt(paragraph,index,value){const ts=textNodes(paragraph);if(ts[index])ts[index].textContent=value||""}
function setAllText(paragraph,value){const ts=textNodes(paragraph);if(!ts.length)return;ts[0].textContent=value||"";for(let i=1;i<ts.length;i++)ts[i].textContent=""}
function cellText(cell,value){setNodeText(cell,value)}
function dateCell(cell,value){const ts=textNodes(cell);if(!ts.length)return;ts[0].textContent=value||"";for(let i=1;i<ts.length;i++)ts[i].textContent=""}
function setParagraphMarker(paragraph,marker,value){
  const ts=textNodes(paragraph);
  if(!ts.length)return false;
  const full=ts.map(n=>n.textContent||"").join("");
  const pos=full.indexOf(marker);
  if(pos<0)return false;

  // Сохраняем маркер и его оформление, заменяем только пустое место после него.
  let seen=false;
  for(const t of ts){
    const val=t.textContent||"";
    if(!seen && val.includes(marker)){
      seen=true;
      const at=val.indexOf(marker)+marker.length;
      t.textContent=val.slice(0,at);
    }else if(seen){
      t.textContent="";
    }
  }
  ts[ts.length-1].textContent=(ts[ts.length-1].textContent||"")+" "+value;
  return true;
}

function setParagraphByText(doc,needle,value){
  const ps=Array.from(doc.getElementsByTagNameNS(NS,"p"));
  const p=ps.find(p=>Array.from(p.getElementsByTagNameNS(NS,"t")).map(n=>n.textContent||"").join("").includes(needle));
  return p ? setParagraphMarker(p,needle,value) : false;
}

async function fillDoc(d){
  const r=await fetch("template.docx",{cache:"no-store"});
  if(!r.ok)throw Error("Не найден Word-шаблон.");
  const zip=await JSZip.loadAsync(await r.arrayBuffer());
  const file=zip.file("word/document.xml");
  if(!file)throw Error("В Word-шаблоне не найден document.xml.");
  const xml=await file.async("string");
  const doc=new DOMParser().parseFromString(xml,"application/xml");
  if(doc.getElementsByTagName("parsererror").length)throw Error("Не удалось прочитать Word-шаблон.");
  const ps=Array.from(doc.getElementsByTagNameNS(NS,"p"));
  const paragraph=i=>ps[i]||null;
  const setP=(i,idx,val)=>{const p=paragraph(i);if(p)setTextAt(p,idx,val)};

  // Заявка: используем текстовые ориентиры из исправленного шаблона.
  setP(2,1,d.car);setP(2,4,d.plate);
  setParagraphByText(doc,"Прибывшее за грузом в адрес получателя",d.recipient);
  setParagraphByText(doc,"ФИО водителя/№ ВУ/паспортные данные",d.driver);
  setP(5,1,d.passportSeries);setP(5,4,d.passportNumber);
  setP(5,7,d.issuedBy);setP(5,9,d.passportDate);
  setP(6,3,d.phone);
  setP(19,5,d.dt);setP(20,1,d.do);setP(22,5,d.jdn);
  setP(23,5,d.invoice);setP(24,2,d.ref);
  setP(25,3,d.places);setP(26,4,d.weight);

  const tables=Array.from(doc.getElementsByTagNameNS(NS,"tbl"));
  const rows=tbl=>tbl?Array.from(tbl.children).filter(n=>n.localName==="tr"):[];
  const cells=row=>row?Array.from(row.children).filter(n=>n.localName==="tc"):[];
  const put=(row,index,val)=>{const c=cells(row);if(c[index])cellText(c[index],val)};

  // Таблица 0 — верхняя таблица доверенности + её блок с датами.
  if(tables[0]){
    const rr=rows(tables[0]);
    if(rr[1]){
      if(cells(rr[1])[1])dateCell(cells(rr[1])[1],d.today);
      if(cells(rr[1])[2])dateCell(cells(rr[1])[2],d.expiry);
      put(rr[1],3,"Водитель "+(d.driver||""));
    }
  }

  // Даты: одна пара дат является источником истины — верхняя таблица.
  // Нижние поля правой страницы получают ровно те же значения:
  // «Дата выдачи» = дата выдачи из верхней таблицы,
  // «Доверенность действительна по» = срок действия из верхней таблицы.
  // Левая страница: дата ставится в существующее поле «Дата:».
  setParagraphByText(doc,"Дата:",d.today);

  // Правая страница: обе нижние даты находятся в таблице 0.
  // Используем те же значения, что и в верхней строке, без поиска
  // по тексту абзацев — это исключает смещение дат.
  if(tables[0]){
    const rr=rows(tables[0]);
    if(rr[14]){
      const c=cells(rr[14]);
      if(c[1])dateCell(c[1],d.today);
    }
    if(rr[15]){
      const c=cells(rr[15]);
      if(c[1])dateCell(c[1],d.expiry);
    }
  }

  if(tables[1]){
    const rr=rows(tables[1]);
    if(rr[2])put(rr[2],4,d.driver);
    if(rr[4]){put(rr[4],1,d.passportSeries);put(rr[4],3,d.passportNumber)}
    if(rr[5])put(rr[5],1,d.issuedBy);
    if(rr[6])put(rr[6],1,d.passportDate);
    if(rr[7])put(rr[7],1,d.recipient);
    if(rr[9])put(rr[9],1,"ДТ "+(d.dt||""));
  }
  if(tables[2]){
    const rr=rows(tables[2]);
    if(rr[1])put(rr[1],3,d.places||"");
  }

  // Удаляем только жёлтую подсветку/заливку, не трогая прочее форматирование.
  for(const el of Array.from(doc.getElementsByTagNameNS(NS,"shd"))){
    const fill=el.getAttributeNS(NS,"fill")||el.getAttribute("w:fill");
    if(String(fill||"").toUpperCase()==="FFFF00" && el.parentNode)el.parentNode.removeChild(el);
  }
  for(const el of Array.from(doc.getElementsByTagNameNS(NS,"highlight"))){
    const val=el.getAttributeNS(NS,"val")||el.getAttribute("w:val");
    if(String(val||"").toLowerCase()==="yellow" && el.parentNode)el.parentNode.removeChild(el);
  }

  const out=new XMLSerializer().serializeToString(doc);
  zip.file("word/document.xml",out);
  const buf=await zip.generateAsync({type:"blob",compression:"DEFLATE"});
  return{blob:buf,filename:`Заявка_и_доверенность_${(d.plate||"готово").replace(/[\/\\:*?"<>|]/g,"")}.docx`};
}
async function downloadDoc(){const r=await fillDoc(values());if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(r.blob);status.innerHTML=`<div class="result"><strong>✓ Word-документ готов</strong><a class="download-link" href="${blobUrl}" download="${r.filename}">Скачать ${r.filename}</a></div>`}
function update(){generate.disabled=!input.value.trim();if(input.value.trim())stats.textContent="Текст готов к распознаванию";else stats.textContent="Данные не введены"}
tool.onchange=()=>{mode=tool.value;help.textContent=mode==="проработка"?"Вставьте диапазон из Excel.":"Порядок сведений не важен — Помощник распознает поля по структуре текста.";format.textContent=mode==="проработка"?"Excel (.xlsx)":"Word (.docx)";generate.textContent=mode==="проработка"?"Сформировать документ":"Распознать данные";input.value="";document.getElementById("parsedFields").hidden=true;if(mode==="проработка")Proработка.update();else update()};
input.oninput=()=>{if(mode==="проработка")Proработка.update();else{update();if(!document.getElementById("parsedFields").hidden)stats.textContent="Текст изменён — нажмите «Обновить распознанные данные»"}};
if(refreshParsed)refreshParsed.onclick=()=>{if(!input.value.trim())return;status.innerHTML="";render(parse(input.value));};
clear.onclick=()=>{input.value="";document.getElementById("parsedFields").hidden=true;status.innerHTML="";if(mode==="проработка")Proработка.update();else update()};
generate.onclick=async()=>{generate.disabled=true;try{if(mode==="проработка"){const r=await Proработка.build();if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(r.blob);status.innerHTML=`<div class="result"><strong>✓ Excel готов</strong><span>${r.text}</span><a class="download-link" href="${blobUrl}" download="${r.filename}">Скачать ${r.filename}</a></div>`;generate.disabled=false;return}if(document.getElementById("parsedFields").hidden){render(parse(input.value));generate.disabled=false}else await downloadDoc()}catch(e){status.innerHTML=`<div class="error"><strong>Ошибка</strong><br>${e.message}</div>`;generate.disabled=false}};
