const App=(()=> {
const input=document.getElementById("input"),stats=document.getElementById("stats"),
clear=document.getElementById("clear"),generate=document.getElementById("generate"),
status=document.getElementById("status"),tool=document.getElementById("document"),
help=document.getElementById("inputHelp"),format=document.querySelector(".format");
let mode=tool.value||"proработка",url=null;

const norm=v=>String(v??"").replace(/\uFEFF/g,"").replace(/\u00a0/g," ").trim().toLowerCase().replace(/\s+/g," ");
const parseTable=t=>t.replace(/\r\n/g,"\n").replace(/\r/g,"\n").split("\n")
 .map(x=>x.split("\t").map(y=>y.trim())).filter(r=>r.some(x=>x!==""));
const aliases={
 "номер ктк":["номер ктк","ктк"],ref:["ref","реф","букинг"],статус:["статус","клиент"],
 "фактический объем, м3":["фактический объем, м3","фактический объем м3","объем м3","объем, м3"],
 "вес, кг":["вес, кг","вес кг","вес"],"кол-во мест":["кол-во мест","количество мест","мест"]
};
function col(headers,key,dict=aliases){
 let h=headers.map(norm),o=(dict[key]||[key]).map(norm),i=h.findIndex(x=>o.includes(x));
 return i>=0?i:h.findIndex(x=>x&&o.some(y=>x.includes(y)||y.includes(x)));
}
function total(r){let s=r.map(norm).filter(Boolean).join(" ");return !s||/итого|всего|total/i.test(s)||(!norm(r[0])&&!norm(r[10]));}

const Proработка={
 update(){
  let r=parseTable(input.value);
  if(!r.length){stats.textContent="Данные не введены";generate.disabled=true;return}
  let req=["номер ктк","ref","статус","фактический объем, м3","вес, кг","кол-во мест"],
      miss=req.filter(k=>col(r[0],k)<0),data=r.slice(1).filter(x=>!total(x));
  stats.textContent=miss.length?`${data.length} строк · не найдены: ${miss.join(", ")}`:
    `${data.length} записей · ${r[0].length} столбцов · все необходимые поля найдены`;
  generate.disabled=!!miss.length||!data.length;
 },
 async build(){
  let r=parseTable(input.value),req=["номер ктк","ref","статус","фактический объем, м3","вес, кг","кол-во мест"],src={};
  req.forEach(k=>src[k]=col(r[0],k));
  if(Object.values(src).some(x=>x<0))throw Error("Не найдены необходимые поля");
  let data=r.slice(1).filter(x=>!total(x)),res=await fetch("проработка.xlsx",{cache:"no-store"});
  if(!res.ok)throw Error("Не найден шаблон «проработка.xlsx».");
  let wb=XLSX.read(new Uint8Array(await res.arrayBuffer()),{type:"array",cellStyles:true}),
      sh=wb.Sheets[wb.SheetNames[0]],rg=XLSX.utils.decode_range(sh["!ref"]),heads=[];
  for(let c=rg.s.c;c<=rg.e.c;c++)heads.push(sh[XLSX.utils.encode_cell({r:rg.s.r,c})]?.v??"");
  let dst={};req.forEach(k=>dst[k]=col(heads,k));
  if(Object.values(dst).some(x=>x<0))throw Error("В шаблоне не найдена необходимая колонка");
  data.forEach((row,i)=>req.forEach(k=>{
   let v=String(row[src[k]]??"").trim(),a=XLSX.utils.encode_cell({r:rg.s.r+1+i,c:dst[k]});
   sh[a]=/^-?\d+(?:[.,]\d+)?$/.test(v.replace(/\s/g,""))?
     {t:"n",v:Number(v.replace(/\s/g,"").replace(",","."))}:{t:"s",v};
  }));
  let out=XLSX.write(wb,{bookType:"xlsx",type:"array",cellStyles:true}),
      ktk=String(data[0]?.[src["номер ктк"]]||"готово").replace(/\//g,"").replace(/[\\:*?"<>|]/g,"");
  return{blob:new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),
    filename:`Проработка_${ktk||"готово"}.xlsx`,text:`Сформировано записей: ${data.length}`};
 }
};

const Request={
 fields:[
  ["driver","Водитель"],["car","Автомобиль"],["plate","Госномер"],["passport","Паспорт"],
  ["issuedBy","Кем выдан"],["passportDate","Дата выдачи паспорта"],["phone","Телефон"],
  ["dt","ДТ"],["do","ДО"],["invoice","Инвойс"]
 ],
 clean(v){return String(v||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim()},
 findPhone(t){const m=t.match(/(?:\+7|8)\s*[\s(]?\d{3}[)\s-]?\s*\d{3}[\s-]?\d{2}[\s-]?\d{2}/);return m?this.clean(m[0]):""},
 findPassport(t){const m=t.match(/\b\d{4}\s*(?:№|N|#)?\s*\d{6}\b/);return m?this.clean(m[0]):""},
 findDates(t){return [...t.matchAll(/\b\d{2}\.\d{2}\.\d{4}\b/g)].map(m=>m[0])},
 findPlate(t){const m=t.match(/\b[АВЕКМНОРСТУХABEKMHOPCTYX]\d{3}[АВЕКМНОРСТУХABEKMHOPCTYX]{2}\d{2,3}\b/i);return m?m[0]:""},
 findFio(t){
  const labeled=t.match(/(?:водитель|фио)\s*[:\-]?\s*([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+){1,2})/);
  if(labeled)return this.clean(labeled[1]);
  const matches=[...t.matchAll(/\b([А-ЯЁ][а-яё]{2,})\s+([А-ЯЁ][а-яё]{2,})(?:\s+([А-ЯЁ][а-яё]{5,}))?\b/g)];
  if(!matches.length)return "";
  const three=matches.find(m=>m[3]);return this.clean(three?three[0]:matches[0][0]);
 },
 extractLabeled(t,label,nextLabels){
  const stop=nextLabels.join("|"),
   re=new RegExp(label+"\\s*[:\\-]?\\s*([^\\n.;]+?)(?=\\s+(?:"+stop+")\\b|[.;]|$)","i"),
   m=t.match(re);return m?this.clean(m[1]):"";
 },
 findDocuments(t){
  const d={dt:"",do:"",invoice:""},
   dt=t.match(/(?:^|[\s.;])ДТ\s*(?:№|N|#)?\s*[:\-]?\s*([0-9/]+)/i),
   dok=t.match(/(?:^|[\s.;])ДО\s*(?:№|N|#)?\s*[:\-]?\s*([0-9A-ZА-ЯЁ\/-]+)/i),
   inv=t.match(/(?:инвойс|invoice)\s*(?:№|N|#)?\s*[:\-]?\s*([0-9A-ZА-ЯЁ\/-]+)/i);
  if(dt)d.dt=dt[1];if(dok)d.do=dok[1];if(inv)d.invoice=inv[1];return d;
 },
 parse(t){
  const x=this.clean(t),phone=this.findPhone(x),passport=this.findPassport(x),
   dates=this.findDates(x),plate=this.findPlate(x),driver=this.findFio(x),docs=this.findDocuments(x);
  let issuedBy=this.extractLabeled(x,"кем выдан",["дата выдачи","телефон","ДТ","ДО","инвойс"]);
  if(!issuedBy){const m=x.match(/паспорт[^.;]*?,?\s*(?:выдан|выданный)\s+(.+?)\s+\d{2}\.\d{2}\.\d{4}/i);if(m)issuedBy=this.clean(m[1]);}
  let passportDate="";
  const labeledDate=x.match(/(?:дата выдачи паспорта|дата выдачи)\s*[:\-]?\s*(\d{2}\.\d{2}\.\d{4})/i);
  if(labeledDate)passportDate=labeledDate[1];else if(dates.length)passportDate=dates[0];
  let car="";
  const carLabeled=x.match(/(?:марка|машина|автомобиль|авто)\s*[:\-]?\s*([A-Za-zА-ЯЁа-яё0-9 ._-]+?)(?=\s+(?:госномер|г\.н\.з\.|г\/н)|[.;]|$)/i);
  if(carLabeled)car=this.clean(carLabeled[1]);
  const now=new Date(),pad=n=>String(n).padStart(2,"0");
  const today=`${pad(now.getDate())}.${pad(now.getMonth()+1)}.${now.getFullYear()}`;
  const expiry=new Date(now);expiry.setFullYear(expiry.getFullYear()+1);expiry.setMonth(expiry.getMonth()+1);expiry.setDate(expiry.getDate()+1);
  return{driver,car,plate,passport,issuedBy,passportDate,phone,dt:docs.dt,do:docs.do,invoice:docs.invoice,today,
    expiry:`${pad(expiry.getDate())}.${pad(expiry.getMonth()+1)}.${expiry.getFullYear()}`};
 },
 render(data){
  const panel=document.getElementById("parsedFields"),grid=document.getElementById("fieldsGrid");panel.hidden=false;
  grid.innerHTML=this.fields.map(([key,label])=>{
   const value=data[key]||"";
   return`<div class="field"><label for="field-${key}">${label}</label>
   <input id="field-${key}" data-key="${key}" value="${String(value).replace(/"/g,"&quot;")}" class="${value?"":"missing"}"></div>`;
  }).join("");
  grid.querySelectorAll("input").forEach(el=>el.addEventListener("input",()=>{
   el.classList.toggle("missing",!el.value.trim());stats.textContent="Проверьте распознанные данные";
  }));
  stats.textContent="Данные распознаны — проверьте поля";generate.textContent="Сформировать документ";
 },
 values(){
  const d={};document.querySelectorAll("#fieldsGrid input[data-key]").forEach(el=>d[el.dataset.key]=el.value.trim());return d;
 },
 async build(d){
  const res=await fetch(encodeURI("Заявка и доверенность шаблон.docx"),{cache:"no-store"});
  if(!res.ok)throw Error("Не найден Word-шаблон «Заявка и доверенность шаблон.docx».");
  return{blob:new Blob([await res.arrayBuffer()],{type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"}),
   filename:`Заявка_и_доверенность_${d.plate||"готово"}.docx`,
   text:"Шаблон подготовлен. Подстановка полей в Word — следующий этап."};
 },
 update(){
  document.getElementById("parsedFields").hidden=true;
  stats.textContent=input.value.trim()?"Текст готов к распознаванию":"Данные не введены";
  generate.disabled=!input.value.trim();generate.textContent="Распознать данные";
 }
};

function modeSet(){
 mode=tool.value;input.value="";status.innerHTML="";
 if(mode==="proработка"){
  help.textContent="В Excel выделите таблицу → Ctrl+C → вставьте ниже.";
  input.placeholder="Вставьте сюда диапазон из Excel (Ctrl+V)";
  format.textContent="Excel (.xlsx)";generate.textContent="Сформировать документ";Proработка.update();
 }else{
  help.textContent="Порядок сведений не важен — Помощник попробует распознать поля автоматически.";
  input.placeholder="Например: Иванов Иван Иванович, Volvo FH, А123АА77, паспорт 4510 123456, 89051234567, ДТ 107..., ДО 000..., инвойс 123...";
  format.textContent="Word (.docx)";generate.textContent="Распознать данные";Request.update();
 }
}

tool.addEventListener("change",modeSet);
input.addEventListener("input",()=>mode==="proработка"?Proработка.update():Request.update());
clear.addEventListener("click",()=>{input.value="";status.innerHTML="";mode==="proработка"?Proработка.update():Request.update();input.focus()});
generate.addEventListener("click",async()=>{
 generate.disabled=true;
 try{
  if(mode==="proработка"){download(await Proработка.build())}
  else{
   const panel=document.getElementById("parsedFields");
   if(panel.hidden){Request.render(Request.parse(input.value));generate.disabled=false;return}
   download(await Request.build(Request.values()));
  }
 }catch(e){status.innerHTML=`<div class="error"><strong>Ошибка</strong><span>${e.message}</span></div>`}
 finally{generate.disabled=false}
});
function download(r){
 if(url)URL.revokeObjectURL(url);url=URL.createObjectURL(r.blob);
 status.innerHTML=`<div class="result"><strong>✓ Документ готов</strong><span>${r.text}</span>
 <a class="download-link" href="${url}" download="${r.filename}">Скачать ${r.filename}</a></div>`;
}
modeSet();
})();