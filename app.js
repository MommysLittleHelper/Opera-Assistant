const input=document.getElementById("input"),stats=document.getElementById("stats"),clear=document.getElementById("clear"),generate=document.getElementById("generate"),status=document.getElementById("status"),tool=document.getElementById("document"),format=document.getElementById("format"),help=document.getElementById("inputHelp"),dataCard=document.getElementById("dataCard"),simpleFields=document.getElementById("simpleFields"),containerNumber=document.getElementById("containerNumber"),jdnNumber=document.getElementById("jdnNumber"),applicationForm=document.getElementById("applicationForm");
const appIds=["appDriver","appCar","appPlate","appPhone","appPassportSeries","appPassportNumber","appIssuedBy","appPassportDate","appRecipient","appJdn","appInvoice","appDt","appDo","appRef","appPlaces","appWeight","appToday","appExpiry"];
const appEls=Object.fromEntries(appIds.map(id=>[id,document.getElementById(id)]));
let mode="",blobUrl=null;

function applicationData(){return Заявка.values()}
function hasApplicationData(){return appIds.some(id=>String(appEls[id]?.value||"").trim())}
function updateApplicationState(){
  const has=hasApplicationData();
  stats.textContent=has?"Данные готовы к формированию":"Данные не введены";
  generate.disabled=!has;
}
async function downloadDoc(){
  const r=await Заявка.fillDoc(applicationData());
  if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(r.blob);
  status.innerHTML=`<div class="result"><strong>✓ Word-документ готов</strong><a class="download-link" href="${blobUrl}" download="${r.filename}">Скачать ${r.filename}</a></div>`;
}
async function downloadLetter(){
  const r=await ПисьмоПеревод.build(containerNumber.value,jdnNumber.value);
  if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(r.blob);
  status.innerHTML=`<div class="result"><strong>✓ Word-документ готов</strong><a class="download-link" href="${blobUrl}" download="${r.filename}">Скачать ${r.filename}</a></div>`;
}
function todayString(){const d=new Date();return `${String(d.getDate()).padStart(2,"0")}.${String(d.getMonth()+1).padStart(2,"0")}.${d.getFullYear()}`}
function syncTool(){
  mode=String(tool.value||"");
  dataCard.classList.toggle("data-card-hidden",!mode);
  applicationForm.hidden=mode!=="заявка";
  input.hidden=mode!=="проработка";
  simpleFields.hidden=mode!=="письмо";
  status.innerHTML="";
  if(mode==="заявка"){
    help.textContent="Заполните поля вручную. Поля подписаны — распознавание текста больше не требуется.";
    format.textContent="Word (.docx)";generate.textContent="Сформировать документ";
    if(appEls.appToday&&!appEls.appToday.value)appEls.appToday.value=todayString();
    updateApplicationState();
  }else if(mode==="проработка"){
    help.textContent="Вставьте диапазон из Excel.";format.textContent="Excel (.xlsx)";generate.textContent="Сформировать документ";Proработка.update();
  }else if(mode==="письмо"){
    help.textContent="Укажите номер контейнера и номер ЖДН. Дата будет проставлена автоматически.";format.textContent="Word (.docx)";generate.textContent="Сформировать документ";generate.disabled=!containerNumber.value.trim()||!jdnNumber.value.trim();
  }else{
    help.textContent="Сначала выберите документ.";format.textContent="—";generate.textContent="Сформировать документ";generate.disabled=true;
  }
}
tool.onchange=syncTool;
input.oninput=()=>{if(mode==="проработка")Proработка.update()};
appIds.forEach(id=>appEls[id]?.addEventListener("input",()=>{if(mode==="заявка")updateApplicationState()}));
containerNumber.oninput=()=>{if(mode==="письмо")generate.disabled=!containerNumber.value.trim()||!jdnNumber.value.trim()};
jdnNumber.oninput=()=>{if(mode==="письмо")generate.disabled=!containerNumber.value.trim()||!jdnNumber.value.trim()};
clear.onclick=()=>{
  if(mode==="заявка"){
    appIds.forEach(id=>{if(appEls[id])appEls[id].value=""});
    if(appEls.appToday)appEls.appToday.value=todayString();
    status.innerHTML="";updateApplicationState();
  }else{
    input.value="";containerNumber.value="";jdnNumber.value="";status.innerHTML="";
    if(mode==="проработка")Proработка.update();else if(mode==="письмо")generate.disabled=true;
  }
};
generate.onclick=async()=>{
  generate.disabled=true;
  try{
    if(mode==="проработка"){
      const r=await Proработка.build();if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(r.blob);status.innerHTML=`<div class="result"><strong>✓ Excel готов</strong><span>${r.text}</span><a class="download-link" href="${blobUrl}" download="${r.filename}">Скачать ${r.filename}</a></div>`;generate.disabled=false;return
    }
    if(mode==="письмо"){await downloadLetter();generate.disabled=false;return}
    if(mode==="заявка"){await downloadDoc();generate.disabled=false}
  }catch(e){status.innerHTML=`<div class="error"><strong>Ошибка</strong><br>${e.message}</div>`;generate.disabled=false}
};
syncTool();
