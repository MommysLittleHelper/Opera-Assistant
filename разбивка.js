// Модуль «Разбивка».
const Разбивка=(()=>{
  const norm=v=>String(v??"").replace(/\uFEFF/g,"").replace(/\u00a0/g," ").trim().toLowerCase().replace(/\s+/g," ");
  const rows=t=>String(t||"").replace(/\r\n/g,"\n").replace(/\r/g,"\n").split("\n").map(x=>x.split("\t").map(y=>y.trim())).filter(r=>r.some(Boolean));
  const findCol=(heads,names)=>{
    const hs=heads.map(norm), ns=names.map(norm);
    let i=hs.findIndex(h=>ns.includes(h));
    if(i<0)i=hs.findIndex(h=>h&&ns.some(n=>h.includes(n)||n.includes(h)));
    return i;
  };
  const num=v=>{
    const x=String(v??"").replace(/\s/g,"").replace(/,/g,".");
    const n=Number(x);
    return Number.isFinite(n)?n:null;
  };
  let parsed=null;

  function parse(){
    const table=rows(document.getElementById("splitTable")?.value);
    const sums=String(document.getElementById("splitSums")?.value||"").split(/\r?\n/).map(x=>x.trim()).filter(Boolean).map(num).filter(x=>x!==null);
    if(table.length<2)throw Error("Вставьте таблицу с заголовками и данными.");
    if(!sums.length)throw Error("Введите хотя бы одну сумму.");
    const headers=table[0];
    const k=findCol(headers,["номер ктк","ктк"]);
    const vol=findCol(headers,["коммерческий объем","коммерческий объём"]);
    const legend=findCol(headers,["номер заказ легенда","номер заказа легенда"]);
    if(k<0||vol<0||legend<0)throw Error("Не найдены обязательные столбцы: «номер ктк», «Коммерческий объем», «номер заказ легенда».");
    const data=table.slice(1).filter(r=>r.some(Boolean)).map(r=>({k:cleanK(r[k]),vol:num(r[vol])||0,legend:String(r[legend]??"").trim()}));
    if(!data.length)throw Error("В таблице нет строк данных.");
    return {data,sums,conditions:[...new Set(data.map(x=>x.k).filter(Boolean))].sort()};
  }
  function cleanK(v){return String(v??"").replace(/\s+/g," ").trim();}
  function update(){
    const msg=document.getElementById("splitMessage"), box=document.getElementById("splitConditions");
    try{
      parsed=parse();
      box.innerHTML="";
      parsed.sums.forEach((s,i)=>{
        const row=document.createElement("div");row.className="split-row";
        const label=document.createElement("label");label.textContent=s.toLocaleString("ru-RU",{maximumFractionDigits:2});
        const sel=document.createElement("select");sel.dataset.index=i;
        ["все","сифы",...parsed.conditions].forEach(v=>{const o=document.createElement("option");o.value=v;o.textContent=v;sel.appendChild(o)});
        row.append(label,sel);box.appendChild(row);
      });
      box.hidden=false;msg.hidden=false;
      stats.textContent=`${parsed.data.length} строк · ${parsed.sums.length} сумм`;
      generate.disabled=false;
    }catch(e){
      parsed=null;box.hidden=true;msg.hidden=true;stats.textContent=e.message;generate.disabled=true;
    }
  }
  function round2(x){return Math.round((x+Number.EPSILON)*100)/100;}
  function resultFor(sum,condition,data){
    if(condition!=="все" && condition!=="сифы"){
      return [{k:condition,v:round2(sum)}];
    }
    let usable=data;
    if(condition==="сифы") usable=data.filter(r=>!r.legend);
    const denominator=condition==="все"
      ? (data[data.length-1]?.vol||0)
      : usable.reduce((a,r)=>a+r.vol,0);
    if(!denominator)throw Error(`Невозможно рассчитать сумму ${sum}: база коммерческого объёма равна 0.`);
    const source=condition==="все" ? data.slice(0,-1) : usable;
    return source.map(r=>({k:r.k,v:round2(sum/denominator*r.vol)}));
  }
  async function build(){
    parsed=parse();
    const selects=[...document.querySelectorAll("#splitConditions select")];
    if(selects.length!==parsed.sums.length)throw Error("Выберите условие для каждой суммы.");
    const groups=parsed.sums.map((sum,i)=>resultFor(sum,selects[i].value,parsed.data));
    const res=await fetch("разбивка пустая.xlsx",{cache:"no-store"});
    if(!res.ok)throw Error("Не найден шаблон «разбивка пустая.xlsx».");
    const wb=XLSX.read(new Uint8Array(await res.arrayBuffer()),{type:"array",cellStyles:true});
    const sh=wb.Sheets[wb.SheetNames[0]];
    let row=0;
    groups.forEach((group,i)=>{
      const sum=parsed.sums[i], condition=selects[i].value;
      sh[XLSX.utils.encode_cell({r:row,c:0})]={t:"s",v:String(sum).replace(".",",")+(condition==="все"||condition==="сифы"?"":` – ${condition}`)};
      sh[XLSX.utils.encode_cell({r:row+1,c:0})]={t:"s",v:"номер ктк"};
      sh[XLSX.utils.encode_cell({r:row+1,c:1})]={t:"s",v:"Разбивка"};
      group.forEach((x,j)=>{
        sh[XLSX.utils.encode_cell({r:row+2+j,c:0})]={t:"s",v:x.k};
        sh[XLSX.utils.encode_cell({r:row+2+j,c:1})]={t:"n",v:x.v};
      });
      const totalRow=row+2+group.length;
      sh[XLSX.utils.encode_cell({r:totalRow,c:1})]={t:"n",v:round2(group.reduce((a,x)=>a+x.v,0))};
      row=totalRow+3;
    });
    sh["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:Math.max(row-1,0),c:1}});
    const out=XLSX.write(wb,{bookType:"xlsx",type:"array",cellStyles:true});
    return {blob:new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),filename:"Разбивка.xlsx",text:`Сформировано блоков: ${groups.length}`};
  }
  return {update,build};
})();
