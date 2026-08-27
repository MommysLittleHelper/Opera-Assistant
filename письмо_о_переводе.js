// Модуль «Письмо о переводе».
const ПисьмоПеревод = {
  async build(container, jdn) {
    const clean = value => String(value ?? "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
    container = clean(container);
    jdn = clean(jdn);
    if (!container) throw Error("Укажите номер контейнера.");
    if (!jdn) throw Error("Укажите номер ЖДН.");

    const r = await fetch("письмо_о_переводе.docx", {cache:"no-store"});
    if (!r.ok) throw Error("Не найден шаблон «письмо_о_переводе.docx».");
    const zip = await JSZip.loadAsync(await r.arrayBuffer());
    const file = zip.file("word/document.xml");
    if (!file) throw Error("В шаблоне письма не найден document.xml.");

    const xml = await file.async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw Error("Не удалось прочитать шаблон письма.");

    const NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const ps = Array.from(doc.getElementsByTagNameNS(NS, "p"));
    const text = p => Array.from(p.getElementsByTagNameNS(NS, "t")).map(n => n.textContent || "").join("");

    const date = new Date();
    const months = ["января","февраля","марта","апреля","мая","июня","июля","августа","сентября","октября","ноября","декабря"];
    const today = `«${String(date.getDate()).padStart(2,"0")}» ${months[date.getMonth()]} ${date.getFullYear()}г.`;

    function replaceParagraphText(p, predicate, replacement) {
      const ts = Array.from(p.getElementsByTagNameNS(NS, "t"));
      if (!ts.length) return false;
      const full = ts.map(n => n.textContent || "").join("");
      if (!predicate(full)) return false;
      ts[0].textContent = replacement;
      for (let i = 1; i < ts.length; i++) ts[i].textContent = "";
      return true;
    }

    // Date paragraph: template contains «от               г.».
    const dateP = ps.find(p => /от\s+.*г\./i.test(text(p)) && text(p).length < 40);
    if (!dateP) throw Error("В шаблоне не найдено поле даты.");
    replaceParagraphText(dateP, s => /от\s+.*г\./i.test(s), `от ${today}`);

    // Main paragraph with two yellow placeholders.
    const bodyP = ps.find(p => text(p).includes("груз/ контейнер №") && text(p).includes("накладной СМГС"));
    if (!bodyP) throw Error("В шаблоне не найдено поле контейнера и ЖДН.");

    const ts = Array.from(bodyP.getElementsByTagNameNS(NS, "t"));
    let containerSet = false, jdnSet = false;
    for (const t of ts) {
      const v = t.textContent || "";
      if (!containerSet && /^\s*$/.test(v) && t.parentNode?.getElementsByTagNameNS(NS,"shd").length) {
        t.textContent = container;
        containerSet = true;
        continue;
      }
      if (containerSet && !jdnSet && /^\s*$/.test(v) && t.parentNode?.getElementsByTagNameNS(NS,"shd").length) {
        t.textContent = jdn;
        jdnSet = true;
      }
    }

    // Fallback for templates where the placeholder is split differently.
    if (!containerSet || !jdnSet) {
      const full = ts.map(n => n.textContent || "").join("");
      const ci = full.indexOf("груз/ контейнер №");
      const ji = full.indexOf("накладной СМГС");
      if (ci >= 0 && ji > ci) {
        let seenContainer=false, seenJdn=false;
        for (const t of ts) {
          const v=t.textContent||"";
          if (!seenContainer && v.includes("груз/ контейнер №")) {
            t.textContent=v.replace("груз/ контейнер №","груз/ контейнер № "+container);
            seenContainer=true;
          }
          if (seenContainer && !seenJdn && v.includes("накладной СМГС")) {
            t.textContent=v.replace("накладной СМГС","накладной СМГС");
            seenJdn=true;
          }
        }
      }
    }

    // Remove yellow highlighting from filled template.
    for (const el of Array.from(doc.getElementsByTagNameNS(NS, "shd"))) {
      const fill = el.getAttributeNS(NS, "fill") || el.getAttribute("w:fill");
      if (String(fill || "").toUpperCase() === "FFFF00" && el.parentNode) el.parentNode.removeChild(el);
    }
    for (const el of Array.from(doc.getElementsByTagNameNS(NS, "highlight"))) {
      const val = el.getAttributeNS(NS, "val") || el.getAttribute("w:val");
      if (String(val || "").toLowerCase() === "yellow" && el.parentNode) el.parentNode.removeChild(el);
    }

    const out = new XMLSerializer().serializeToString(doc);
    zip.file("word/document.xml", out);
    const blob = await zip.generateAsync({type:"blob", compression:"DEFLATE"});
    const safe = container.replace(/[\\/:*?"<>|]/g,"");
    return {
      blob,
      filename: `Письмо_о_переводе_${safe || "готово"}.docx`
    };
  }
};
