// Модуль «Заявка + Доверенность».
const Заявка = (() => {
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
    const raw=String(t||"").replace(/\r/g,"");
    const lines=raw.split("\n").map(x=>x.trim()).filter(Boolean);
    const all=lines.join(" ").replace(/\s+/g," ").trim();
    const d={};

    // Helpers
    const clean=v=>String(v||"").replace(/\s+/g," ").replace(/^[,;:\s]+|[,;:\s]+$/g,"").trim();
    const set=(k,v)=>{v=clean(v);if(v&&!d[k])d[k]=v};

    // 1) Phone — independent of labels.
    const phone=all.match(/(?:\+7|8)\s*[\(\s-]?\d{3}[\)\s-]?\s*\d{3}\s*[-\s]?\d{2}\s*[-\s]?\d{2}/);
    if(phone) set("phone",phone[0]);

    // 2) Passport: 4 digits + 6 digits. Also accept compact 10 digits.
    //    Date immediately following the passport is treated as issue date.
    const pass=/\b(\d{4})\s*(\d{6})\b/.exec(all);
    if(pass){
      set("passportSeries",pass[1]);
      set("passportNumber",pass[2]);
      const after=all.slice(pass.index+pass[0].length);
      const pd=after.match(/\b(\d{2}\.\d{2}\.\d{4})\b/);
      if(pd)set("passportDate",pd[1]);
    }

    // 3) Russian vehicle registration plate. Support Cyrillic letters
    //    visually used in Russian plates and 3-digit region.
    const plate=/\b([АВЕКМНОРСТУХABEKMHOPCTYX])\s?\d{3}\s?([АВЕКМНОРСТУХABEKMHOPCTYX]{2})\s?\d{2,3}\b/i.exec(all);
    if(plate){
      const normalized=plate[0].replace(/\s+/g,"").toUpperCase();
      set("plate",normalized);
    }

    // 4) FIO without the word "водитель".
    //    Prefer a 3-word Cyrillic sequence; fall back to 2 words.
    //    Exclude obvious organization/technical words.
    const bad=new Set(["ТП","ОТДЕЛА","РОССИИ","САНКТ","ПЕТЕРБУРГУ","ЛО","ГАЗ","GAZ","NEXT","ВОДИТЕЛЬ","ПАСПОРТ"]);
    const fio3=/\b([А-ЯЁ][а-яё-]+)\s+([А-ЯЁ][а-яё-]+)\s+([А-ЯЁ][а-яё-]+)\b/.exec(all);
    if(fio3 && !bad.has(fio3[1].toUpperCase()) && !bad.has(fio3[2].toUpperCase()) && !bad.has(fio3[3].toUpperCase())){
      set("driver",fio3[0]);
    }else{
      const fio2=/\b([А-ЯЁ][а-яё-]+)\s+([А-ЯЁ][а-яё-]+)\b/.exec(all);
      if(fio2 && !bad.has(fio2[1].toUpperCase()) && !bad.has(fio2[2].toUpperCase()))
        set("driver",fio2[0]);
    }

    // 5) Passport issuer: text after passport date and before vehicle/plate.
    if(d.passportDate){
      const di=all.indexOf(d.passportDate);
      if(di>=0){
        const after=all.slice(di+d.passportDate.length);
        const stopCandidates=[
          d.plate ? after.toUpperCase().indexOf(d.plate.toUpperCase()) : -1,
          after.search(/\b(?:ГАЗ|КАМАЗ|МАЗ|MAN|VOLVO|SCANIA|DAF|IVECO|MERCEDES|RENAULT|FORD|FAW|HOWO|SHACMAN)\b/i)
        ].filter(x=>x>=0);
        const stop=stopCandidates.length?Math.min(...stopCandidates):after.length;
        set("passportIssuedBy",after.slice(0,stop));
      }
    }

    // 6) Vehicle model: text around the plate, preferably after a known
    //    manufacturer/vehicle marker, but do not require the marker.
    if(d.plate){
      const pi=all.toUpperCase().indexOf(d.plate.toUpperCase());
      if(pi>=0){
        let before=all.slice(0,pi).trim();
        const marker=before.match(/(?:ГАЗ|GAZ|КАМАЗ|МАЗ|MAN|VOLVO|SCANIA|DAF|IVECO|MERCEDES|RENAULT|FORD|FAW|HOWO|SHACMAN)\b/i);
        if(marker){
          const candidate=before.slice(marker.index).trim();
          set("vehicle",candidate);
        }else if(d.passportDate){
          const afterDate=all.slice(all.indexOf(d.passportDate)+d.passportDate.length,pi).trim();
          set("vehicle",afterDate);
        }
      }
    }

    // 7) Fixed transport block — parse structurally, without requiring
    //    the user to add labels.
    const jdn=all.match(/\bЖДН\s*№?\s*([0-9]+)\s*от\s*(\d{2}\.\d{2}\.\d{4})(?:\s*\(в\s*адрес:\s*(.*?)\))?/i);
    if(jdn){
      set("jdn",jdn[1]);
      set("jdnDate",jdn[2]);
      if(jdn[3]){
        let recipient=clean(jdn[3]);
        recipient=recipient.replace(/\s+П\/П\s+/ig," ").replace(/\bАО\s*["«].*?["»]\s*/ig,"").trim();
        recipient=recipient.replace(/^.*?\bООО\b/i,"ООО").replace(/\s+/g," ");
        set("recipient",recipient);
      }
    }

    const inv=all.match(/\bИНВОЙС\s*№?\s*([A-ZА-Я0-9-]+)\s*от\s*(\d{2}\.\d{2}\.\d{4})/i);
    if(inv){set("invoice",inv[1]);set("invoiceDate",inv[2]);}

    const dt=all.match(/\bДТ\s*([0-9\/]+)/i);
    if(dt)set("dt",dt[1]);

    // 8) Manual-format values are also recognized when pasted, but manual
    //    fields in the UI override them.
    const weight=all.match(/\b(\d+(?:[.,]\d+)?)\s*кг\b/i);
    if(weight)set("weight",weight[1].replace(",",".")+" кг");

    const places=all.match(/\b(?:кол-?во\s+мест|количество\s+мест|мест(?:о|а)?)\s*[:№]?\s*(\d+)\b/i)
      || all.match(/\b(\d+)\s+мест(?:о|а)?\b/i);
    if(places)set("places",places[1]);

    const doNum=all.match(/\bДО\s*№?\s*([0-9]+)\b/i);
    if(doNum)set("do",doNum[1]);

    const ref=all.match(/\bREF\s*[:№]?\s*([A-Z0-9-]+)\b/i);
    if(ref)set("ref",ref[1]);

    return d;
  }

  function render(d){const panel=document.getElementById("parsedFields"),grid=document.getElementById("fieldsGrid");panel.hidden=false;grid.innerHTML=fields.map(([k,l])=>`<div class="field"><label>${l}</label><input data-key="${k}" value="${String(d[k]||"").replace(/"/g,"&quot;")}" class="${d[k]?"":"missing"}">${d[k]?"":"<small>не найдено — можно оставить пустым</small>"}</div>`).join("");grid.querySelectorAll("input").forEach(e=>e.oninput=()=>e.classList.toggle("missing",!e.value.trim()));if(refreshParsed)refreshParsed.disabled=false;stats.textContent="Данные распознаны — проверьте поля";generate.textContent="Сформировать документ"}
  function values(){
    const get=id=>{const e=document.getElementById(id);return e?e.value.trim():""};
    return {
      driver:get("appDriver"),car:get("appCar"),plate:get("appPlate"),phone:get("appPhone"),
      passportSeries:get("appPassportSeries"),passportNumber:get("appPassportNumber"),
      issuedBy:get("appIssuedBy"),passportDate:get("appPassportDate"),recipient:get("appRecipient"),
      jdn:get("appJdn"),invoice:get("appInvoice"),dt:get("appDt"),do:get("appDo"),ref:get("appRef"),
      places:get("appPlaces"),weight:get("appWeight"),today:get("appToday"),expiry:get("appExpiry")
    };
  }
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
  function parseSources(driverText, transportText, manual={}) {
    const d = parse(`${driverText || ""}\n${transportText || ""}`);
    if (manual.places) d.places=String(manual.places).trim();
    if (manual.weight) d.weight=String(manual.weight).trim().replace(/\s*кг\s*$/i,"")+" кг";
    if (manual.do) d.do=String(manual.do).trim();
    if (manual.ref) d.ref=String(manual.ref).trim();
    return d;
  }
  return {parse,parseSources,render,values,fillDoc};
})();
