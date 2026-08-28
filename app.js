const input=document.getElementById("input"),refreshParsed=document.getElementById("refreshParsed"),stats=document.getElementById("stats"),clear=document.getElementById("clear"),generate=document.getElementById("generate"),status=document.getElementById("status"),tool=document.getElementById("document"),format=document.getElementById("format"),help=document.getElementById("inputHelp"),dataCard=document.getElementById("dataCard"),simpleFields=document.getElementById("simpleFields"),containerNumber=document.getElementById("containerNumber"),jdnNumber=document.getElementById("jdnNumber");
const sourceDriver=document.getElementById("sourceDriver"),sourceTransport=document.getElementById("sourceTransport"),manualPlaces=document.getElementById("manualPlaces"),manualWeight=document.getElementById("manualWeight"),manualDo=document.getElementById("manualDo"),manualRef=document.getElementById("manualRef"),applicationSources=document.getElementById("applicationSources");
let mode="",blobUrl=null;

function applicationData(){
  return Заявка.parseSources(sourceDriver?.value||"",sourceTransport?.value||"",{
    places:manualPlaces?.value||"",weight:manualWeight?.value||"",do:manualDo?.value||"",ref:manualRef?.value||""
  });
}
function hasApplicationData(){return [sourceDriver,sourceTransport,manualPlaces,manualWeight,manualDo,manualRef].some(e=>e&&String(e.value||"").trim())}
function updateApplicationState(){
  const has=hasApplicationData();
  stats.textContent=has?"Данные готовы к проверке":"Данные не введены";
  generate.disabled=!has;
}
async function downloadDoc(){
  const r=await Заявка.fillDoc(Заявка.values());
  if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(r.blob);
  status.innerHTML=`<div class="result"><strong>✓ Word-документ готов</strong><a class="download-link" href="${blobUrl}" download="${r.filename}">Скачать ${r.filename}</a></div>`;
}
async function downloadLetter(){
  const r=await ПисьмоПеревод.build(containerNumber.value,jdnNumber.value);
  if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(r.blob);
  status.innerHTML=`<div class="result"><strong>✓ Word-документ готов</strong><a class="download-link" href="${blobUrl}" download="${r.filename}">Скачать ${r.filename}</a></div>`;
}
function syncTool(){
  mode=String(tool.value||"");
  dataCard.classList.toggle("data-card-hidden",!mode);
  if(applicationSources)applicationSources.hidden=mode!=="заявка";
  input.hidden=mode!=="проработка";
  simpleFields.hidden=mode!=="письмо";
  document.getElementById("parsedFields").hidden=true;
  status.innerHTML="";
  if(mode==="проработка"){help.textContent="Вставьте диапазон из Excel.";format.textContent="Excel (.xlsx)";generate.textContent="Сформировать документ";Proработка.update()}
  else if(mode==="письмо"){help.textContent="Укажите номер контейнера и номер ЖДН. Дата будет проставлена автоматически.";format.textContent="Word (.docx)";generate.textContent="Сформировать документ";generate.disabled=!containerNumber.value.trim()||!jdnNumber.value.trim()}
  else if(mode==="заявка"){help.textContent="Вставьте данные по трём источникам. Первые два блока можно вставлять как есть, последние данные вводятся вручную.";format.textContent="Word (.docx)";generate.textContent="Проверить данные";updateApplicationState()}
  else{help.textContent="Сначала выберите документ.";format.textContent="—";generate.textContent="Сформировать документ";generate.disabled=true}
}
tool.onchange=syncTool;
input.oninput=()=>{if(mode==="проработка")Proработка.update()};
[sourceDriver,sourceTransport,manualPlaces,manualWeight,manualDo,manualRef].forEach(e=>e&&e.addEventListener("input",()=>{if(mode==="заявка"){if(!document.getElementById("parsedFields").hidden)stats.textContent="Данные изменены — нажмите «Обновить распознанные данные»";else updateApplicationState()}}));
containerNumber.oninput=()=>{if(mode==="письмо")generate.disabled=!containerNumber.value.trim()||!jdnNumber.value.trim()};
jdnNumber.oninput=()=>{if(mode==="письмо")generate.disabled=!containerNumber.value.trim()||!jdnNumber.value.trim()};
refreshParsed.onclick=()=>{if(mode==="заявка"&&hasApplicationData()){status.innerHTML="";Заявка.render(applicationData())}};
clear.onclick=()=>{
  if(mode==="заявка"){[sourceDriver,sourceTransport,manualPlaces,manualWeight,manualDo,manualRef].forEach(e=>e.value="");document.getElementById("parsedFields").hidden=true;status.innerHTML="";updateApplicationState()}
  else{input.value="";containerNumber.value="";jdnNumber.value="";document.getElementById("parsedFields").hidden=true;status.innerHTML="";if(mode==="проработка")Proработка.update();else if(mode==="письмо")generate.disabled=true}
};
generate.onclick=async()=>{
  generate.disabled=true;
  try{
    if(mode==="проработка"){const r=await Proработка.build();if(blobUrl)URL.revokeObjectURL(blobUrl);blobUrl=URL.createObjectURL(r.blob);status.innerHTML=`<div class="result"><strong>✓ Excel готов</strong><span>${r.text}</span><a class="download-link" href="${blobUrl}" download="${r.filename}">Скачать ${r.filename}</a></div>`;generate.disabled=false;return}
    if(mode==="письмо"){await downloadLetter();generate.disabled=false;return}
    if(mode==="заявка"){if(document.getElementById("parsedFields").hidden){Заявка.render(applicationData());generate.disabled=false}else{await downloadDoc();generate.disabled=false}}
  }catch(e){status.innerHTML=`<div class="error"><strong>Ошибка</strong><br>${e.message}</div>`;generate.disabled=false}
};
syncTool();
