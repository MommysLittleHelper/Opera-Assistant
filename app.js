const input=document.getElementById("input"),stats=document.getElementById("stats"),clear=document.getElementById("clear"),generate=document.getElementById("generate"),status=document.getElementById("status"),tool=document.getElementById("document"),format=document.getElementById("format"),help=document.getElementById("inputHelp");
let mode="заявка",blobUrl=null;
const fields=[["driver","Водитель"],["car","Автомобиль"],["plate","Госномер"],["passportSeries","Серия паспорта"],["passportNumber","Номер паспорта"],["issuedBy","Кем выдан"],["passportDate","Дата выдачи паспорта"],["phone","Телефон"],["recipient","Получатель"],["dt","ДТ"],["do","ДО"],["jdn","ЖДН / CMR"],["invoice","Инвойс"],["invoiceDate","Дата инвойса"],["ref","REF"],["places","Количество мест"],["weight","Вес брутто"]];
const NS="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
function clean(s){return String(s||"").replace(/\u00a0/g," ").replace(/[ \t]+/g," ").trim()}
function dateList(t){return [...t.matchAll(/\b\d{2}\.\d{2}\.\d{4}\b/g)].map(m=>m[0])}
function phone(t){const m=t.match(/(?:\+7|8)\s*\(?\d{3}\)?[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}/);return m?clean(m[0]):""}
function passport(t){const m=t.match(/\b(\d{4})\s+(\d{6})\b/);return m?{series:m[1],number:m[2]}:{series:"",number:""}}
function parse(t){
  const x=String(t||'').replace(/\r/g,' ').replace(/\n/g,' ').replace(/\s+/g,' ').trim();
  const d={driver:'',car:'',plate:'',passportSeries:'',passportNumber:'',issuedBy:'',passportDate:'',phone:'',recipient:'',dt:'',do:'',jdn:'',invoice:'',invoiceDate:'',ref:'',places:'',weight:''};

  const between=(labels,next=[])=>{
    const low=x.toLocaleLowerCase('ru-RU');
    let best=-1,bestLen=0;
    for(const label of labels){
      const i=low.indexOf(label.toLocaleLowerCase('ru-RU'));
      if(i>=0 && (best<0 || i<best)){best=i;bestLen=label.length}
    }
    if(best<0)return '';
    let s=best+bestLen;
    while(/[ \t:№#.,-]/.test(x[s]||''))s++;
    let e=x.length;
    for(const n of next){
      const j=low.indexOf(n.toLocaleLowerCase('ru-RU'),s);
      if(j>=0 && j<e)e=j;
    }
    return x.slice(s,e).trim().replace(/[;,]+$/,'').trim();
  };

  d.driver=between(['водитель','ФИО водителя'],['паспорт']);
  d.car=between(['автомобиль'],['госномер','г.н.з.','телефон']).replace(/^ГАЗ\s+/i,'').trim();
  d.plate=between(['госномер','г.н.з.'],['телефон','ДО','ЖДН','CMR','ИНВОЙС','ДТ']).replace(/[^\dA-ZА-ЯЁ]/gi,'');
  const pass=between(['паспорт'],['дата выдачи паспорта','выдан']);
  let m=pass.match(/(\d{4})\s+(?:№\s*)?(\d{6})/);
  if(m){d.passportSeries=m[1];d.passportNumber=m[2]}
  d.passportDate=between(['дата выдачи паспорта'],['выдан','автомобиль','газ автомобиль','госномер','телефон']);
  d.issuedBy=between(['выдан'],['дата выдачи паспорта','газ автомобиль','автомобиль','госномер','телефон','ДО','ЖДН','CMR','ИНВОЙС','ДТ']).replace(/[,\s]+$/,'');
  d.phone=between(['телефон','тел.'],['ДО','ЖДН','CMR','ИНВОЙС','ДТ']);
  d.do=between(['ДО'],['ЖДН','CMR','ИНВОЙС','ДТ']);
  d.jdn=between(['ЖДН','CMR(ЖДН)'],['ИНВОЙС','ДТ']);
  m=d.jdn.match(/№?\s*(\d+)\s+от\s+(\d{2}\.\d{2}\.\d{4})/i);
  if(m)d.jdn=`№${m[1]} от ${m[2]}`; else {m=d.jdn.match(/(\d+)/);if(m)d.jdn=`№${m[1]}`}
  const addr=between(['в адрес:'],['инвойс','ДТ']);
  d.recipient=addr.replace(/\s*П\/П\s+.+$/i,'').trim();
  d.invoice=between(['инвойс'],['ДТ']);
  m=d.invoice.match(/№?\s*([0-9A-ZА-ЯЁ/-]+)(?:\s+от\s+(\d{2}\.\d{2}\.\d{4}))?/i);
  if(m){d.invoice=m[1];d.invoiceDate=m[2]||''}
  d.dt=between(['ДТ'],[]);
  m=d.dt.match(/\d{5,}\/\d{5,}\/\d{5,}/); if(m)d.dt=m[0];
  d.ref=between(['REF'],['количество мест','кол-во мест','количество','вес','место']);
  // Поддерживаем все три реальные формы: «Количество мест 1», «Кол-во мест 1», «1 место».
  d.places=between(['количество мест','кол-во мест','количество мест:'],['вес','реф','ref']);
  if(!d.places){
    m=x.match(/\b(\d+)\s+мест(?:о|а)?\b/i); if(m)d.places=m[1];
  }
  d.weight=between(['вес брутто','вес'],['реф','ref','количество мест','кол-во мест']);
  if(!d.weight){
    m=x.match(/\b(\d+(?:[.,]\d+)?)\s*(?:кг|килограмм(?:а|ов)?)\b/i); if(m)d.weight=m[1]+' кг';
  }
  d.places=(d.places.match(/\d+/)||[''])[0];
  d.weight=d.weight.replace(/\s+/g,' ').trim();
  const now=new Date(),pad=n=>String(n).padStart(2,'0');
  const today=`${pad(now.getDate())}.${pad(now.getMonth()+1)}.${now.getFullYear()}`;
  const expiry=new Date(now);
  expiry.setFullYear(expiry.getFullYear()+1);
  expiry.setMonth(expiry.getMonth()+1);
  expiry.setDate(expiry.getDate()+1);
  d.today=today; d.expiry=`${pad(expiry.getDate())}.${pad(expiry.getMonth()+1)}.${expiry.getFullYear()}`;
  return d;
}
function render(d){const panel=document.getElementById("parsedFields"),grid=document.getElementById("fieldsGrid");panel.hidden=false;grid.innerHTML=fields.map(([k,l])=>`<div class="field"><label>${l}</label><input data-key="${k}" value="${String(d[k]||"").replace(/"/g,"&quot;")}" class="${d[k]?"":"missing"}">${d[k]?"":"<small>не найдено — можно оставить пустым</small>"}</div>`).join("");grid.querySelectorAll("input").forEach(e=>e.oninput=()=>e.classList.toggle("missing",!e.value.trim()));stats.textContent="Данные распознаны — проверьте поля";generate.textContent="Сформировать документ"}
function values(){const d={};document.querySelectorAll("#fieldsGrid input").forEach(e=>d[e.dataset.key]=e.value.trim());return d}
function textNodes(root){return Array.from(root.getElementsByTagNameNS(NS,"t"))}
function setTextAt(paragraph,index,value){const ts=textNodes(paragraph);if(ts[index])ts[index].textContent=value||""}
function cellText(cell,value){const ts=textNodes(cell);if(ts.length)ts[ts.length-1].textContent=value||""}
async function fillDoc(d){
  const r=await fetch("Заявка и доверенность шаблон.docx",{cache:"no-store"});
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

  // Заявка. Проверяем существование абзаца перед обращением к нему.
  setP(2,1,d.car);setP(2,4,d.plate);
  setP(3,2,d.recipient);setP(4,1,d.driver);
  setP(5,1,d.passportSeries);setP(5,4,d.passportNumber);
  setP(5,7,d.issuedBy);setP(5,9,d.passportDate);
  setP(6,3,d.phone);
  setP(19,5,d.dt);setP(20,1,d.do);setP(22,5,d.jdn);
  setP(23,5,d.invoice);setP(24,2,d.ref);
  setP(25,3,d.places);setP(26,4,d.weight);setP(28,1,d.today);

  const tables=Array.from(doc.getElementsByTagNameNS(NS,"tbl"));
  const rows=tbl=>tbl?Array.from(tbl.children).filter(n=>n.localName==="tr"):[];
  const cells=row=>row?Array.from(row.children).filter(n=>n.localName==="tc"):[];
  const put=(row,index,val)=>{const c=cells(row);if(c[index])cellText(c[index],val)};

  // Таблица 0: верхняя часть доверенности.
  if(tables[0]){
    const rr=rows(tables[0]);
    if(rr[1]){
      put(rr[1],1,d.today);
      put(rr[1],2,d.expiry);
      put(rr[1],3,"Водитель "+(d.driver||""));
    }
    // Две даты также есть в нижней части этого же документа.
    if(rr[14])put(rr[14],1,d.today);
    if(rr[15])put(rr[15],1,d.expiry);
  }

  // Таблица 1: реквизиты доверенности.
  if(tables[1]){
    const rr=rows(tables[1]);
    if(rr[2])put(rr[2],4,d.driver);
    if(rr[4]){put(rr[4],1,d.passportSeries);put(rr[4],3,d.passportNumber)}
    if(rr[5])put(rr[5],1,d.issuedBy);
    if(rr[6])put(rr[6],1,d.passportDate);
    if(rr[7])put(rr[7],1,d.recipient);
    if(rr[9])put(rr[9],2,"ДТ "+(d.dt||""));
  }

  // Таблица 2: количество мест.
  if(tables[2]){
    const rr=rows(tables[2]);
    if(rr[1])put(rr[1],3,d.places||"");
  }

  // Убираем жёлтую разметку шаблона после заполнения.
  const marks=Array.from(doc.getElementsByTagNameNS(NS,"shd"));
  marks.forEach(el=>{
    const fill=el.getAttributeNS(NS,"fill")||el.getAttribute("w:fill");
    if(String(fill||"").toUpperCase()==="FFFF00" && el.parentNode)el.parentNode.removeChild(el);
  });
  const highlights=Array.from(doc.getElementsByTagNameNS(NS,"highlight"));
  highlights.forEach(el=>{
    const val=el.getAttributeNS(NS,"val")||el.getAttribute("w:val");
    if(String(val||"").toLowerCase()==="yellow" && el.parentNode)el.parentNode.removeChild(el);
  });

  const out=new XMLSerializer().serializeToString(doc);
  zip.file("word/document.xml",out);
  const buf=await zip.generateAsync({type:"blob",compression:"DEFLATE"});
  return{blob:buf,filename:`Заявка_и_доверенность_${(d.plate||"готово").replace(/[\/\\:*?"<>|]/g,"")}.docx`};
}
async function downloadDoc(){const r=await fillDoc(values());if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(r.blob);status.innerHTML=`<div class="result"><strong>✓ Word-документ готов</strong><a class="download-link" href="${blobUrl}" download="${r.filename}">Скачать ${r.filename}</a></div>`}
function update(){generate.disabled=!input.value.trim();if(input.value.trim())stats.textContent="Текст готов к распознаванию";else stats.textContent="Данные не введены"}
tool.onchange=()=>{mode=tool.value;help.textContent=mode==="проработка"?"Вставьте диапазон из Excel.":"Порядок сведений не важен — Помощник распознает поля по структуре текста.";format.textContent=mode==="проработка"?"Excel (.xlsx)":"Word (.docx)";generate.textContent=mode==="проработка"?"Сформировать документ":"Распознать данные";input.value="";document.getElementById("parsedFields").hidden=true;update()};
input.oninput=update;clear.onclick=()=>{input.value="";document.getElementById("parsedFields").hidden=true;status.innerHTML="";update()};
generate.onclick=async()=>{generate.disabled=true;try{if(mode==="проработка")throw Error("Проработка в этой версии не изменялась. Для теста выберите «Заявка + Доверенность».");if(document.getElementById("parsedFields").hidden){render(parse(input.value));generate.disabled=false}else await downloadDoc()}catch(e){status.innerHTML=`<div class="error"><strong>Ошибка</strong><br>${e.message}</div>`;generate.disabled=false}};
