// Модуль «Разбивка».
const Разбивка=(()=>{
  const $=id=>document.getElementById(id);
  const norm=v=>String(v??"").replace(/\uFEFF/g,"").replace(/\u00A0/g," ").trim().toLowerCase().replace(/\s+/g," ");
  const clean=v=>String(v??"").replace(/\uFEFF/g,"").replace(/\u00A0/g," ").replace(/\s+/g," ").trim();
  const rows=t=>String(t||"").replace(/\r\n/g,"\n").replace(/\r/g,"\n")
    .split("\n").map(x=>x.split("\t").map(clean)).filter(r=>r.some(Boolean));
  const findCol=(heads,names)=>{
    const hs=heads.map(norm), ns=names.map(norm);
    let i=hs.findIndex(h=>ns.includes(h));
    if(i<0)i=hs.findIndex(h=>h && ns.some(n=>h.includes(n)||n.includes(h)));
    return i;
  };
  const num=v=>{
    let x=String(v??"").replace(/\u00A0/g," ").replace(/\s/g,"").replace(/,/g,".");
    x=x.replace(/[^\d.+-]/g,"");
    const n=Number(x);
    return Number.isFinite(n)?n:null;
  };
  const money=v=>{
    const n=num(v);
    return n===null?null:n;
  };
  let parsed=null;

  function parse(){
    const table=rows($("splitTable")?.value);
    const sumLines=String($("splitSums")?.value||"").replace(/\r\n/g,"\n").replace(/\r/g,"\n")
      .split("\n").map(clean).filter(Boolean);
    if(table.length<2)throw Error("Вставьте таблицу с заголовками и данными.");
    if(!sumLines.length)throw Error("Введите хотя бы одну сумму.");

    const headers=table[0];
    const k=findCol(headers,["номер ктк"]);
    const vol=findCol(headers,["коммерческий объем","коммерческий объём"]);
    const legend=findCol(headers,["номер заказ легенда","номер заказа легенда"]);
    if(k<0)throw Error("Не найден столбец «номер ктк».");
    if(vol<0)throw Error("Не найден столбец «Коммерческий объем».");
    if(legend<0)throw Error("Не найден столбец «номер заказ легенда».");

    const data=table.slice(1).map(r=>({
      k:clean(r[k]??""),
      vol:num(r[vol])??0,
      legend:clean(r[legend]??"")
    })).filter(r=>r.k || r.vol || r.legend);

    if(!data.length)throw Error("В таблице нет строк данных.");

    const sums=[];
    for(const line of sumLines){
      const n=money(line);
      if(n===null)throw Error(`Не удалось распознать сумму: «${line}».`);
      sums.push(n);
    }
    return {data,sums,conditions:[...new Set(data.map(x=>x.k).filter(Boolean))]};
  }

  function update(){
    const msg=$("splitMessage"), box=$("splitConditions");
    try{
      parsed=parse();
      box.innerHTML="";
      parsed.sums.forEach((s,i)=>{
        const row=document.createElement("div");
        row.className="split-row";
        const label=document.createElement("label");
        label.textContent=s.toLocaleString("ru-RU",{maximumFractionDigits:2});
        const sel=document.createElement("select");
        sel.dataset.index=String(i);
        ["все","сифы",...parsed.conditions].forEach(v=>{
          const o=document.createElement("option");
          o.value=v;o.textContent=v;sel.appendChild(o);
        });
        row.append(label,sel);
        box.appendChild(row);
      });
      box.hidden=false;
      msg.hidden=false;
      stats.textContent=`Расчёты сделаны, выберите условия показа`;
      generate.disabled=false;
    }catch(e){
      parsed=null;
      box.innerHTML="";
      box.hidden=true;
      msg.hidden=true;
      stats.textContent=e.message;
      generate.disabled=true;
    }
  }

  function round2(x){
    return Math.round((x+Number.EPSILON)*100)/100;
  }

  function resultFor(sum,condition,data){
    if(condition!=="все" && condition!=="сифы"){
      return [{k:condition,v:round2(sum)}];
    }

    const usable=condition==="сифы" ? data.filter(r=>!r.legend) : data;
    if(!usable.length)throw Error(`Для суммы ${sum} не осталось строк для расчёта.`);

    let denominator;
    let source;
    if(condition==="все"){
      const last=data[data.length-1];
      denominator=last?.vol??0;
      source=data.slice(0,-1);
    }else{
      denominator=usable.reduce((a,r)=>a+r.vol,0);
      source=usable;
    }

    if(!(denominator>0))throw Error(`Невозможно рассчитать сумму ${sum}: база коммерческого объёма равна 0.`);
    return source.map(r=>({k:r.k,v:round2(sum/denominator*r.vol)}));
  }

  async function build(){
    const current=parse();
    const selects=[...document.querySelectorAll("#splitConditions select")];
    if(selects.length!==current.sums.length){
      throw Error("Сначала выберите условие для каждой суммы.");
    }

    const groups=current.sums.map((sum,i)=>{
      const condition=String(selects[i].value||"").trim();
      if(!condition)throw Error(`Не выбрано условие для суммы ${sum}.`);
      return resultFor(sum,condition,current.data);
    });

    const templateUrl="разбивка пустая.xlsx";
    const res=await fetch(templateUrl,{cache:"no-store"});
    if(!res.ok)throw Error("Не найден шаблон «разбивка пустая.xlsx».");
    const wb=XLSX.read(new Uint8Array(await res.arrayBuffer()),{type:"array",cellStyles:true});
    const sh=wb.Sheets[wb.SheetNames[0]];
    let row=0;

    groups.forEach((group,i)=>{
      const sum=current.sums[i],condition=selects[i].value;
      const put=(r,c,v,t)=>{sh[XLSX.utils.encode_cell({r,c})]={t,v};};

      put(row,0,String(sum).replace(".",",")+" — "+condition,"s");
      put(row+1,0,"номер ктк","s");
      put(row+1,1,"Разбивка","s");

      group.forEach((x,j)=>{
        put(row+2+j,0,x.k,"s");
        put(row+2+j,1,x.v,"n");
      });

      const total=round2(group.reduce((a,x)=>a+x.v,0));
      put(row+2+group.length,0,"Итого","s");
      put(row+2+group.length,1,total,"n");
      row=row+2+group.length+3;
    });

    sh["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:Math.max(row-1,0),c:1}});
    const out=XLSX.write(wb,{bookType:"xlsx",type:"array",cellStyles:true});
    return {
      blob:new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),
      filename:"Разбивка.xlsx",
      text:`Сформировано блоков: ${groups.length}`
    };
  }

  return {update,build};
})();
