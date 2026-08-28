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

    // Main paragraph: fill the TWO independent yellow placeholders.
    // In the Word template the container placeholder is split across two
    // yellow runs, while the JDN placeholder is a separate yellow run.
    const bodyP = ps.find(p => {
      const s = text(p);
      return s.includes("груз/ контейнер №") && s.includes("накладной СМГС) №");
    });
    if (!bodyP) throw Error("В шаблоне не найдено поле контейнера и ЖДН.");

    const getRuns = p => Array.from(p.getElementsByTagNameNS(NS, "r"));
    const runText = r => Array.from(r.getElementsByTagNameNS(NS, "t"))
      .map(n => n.textContent || "").join("");

    const isYellowRun = r => {
      const shds = Array.from(r.getElementsByTagNameNS(NS, "shd"));
      return shds.some(el => {
        const fill = el.getAttributeNS(NS, "fill") || el.getAttribute("w:fill");
        return String(fill || "").toUpperCase() === "FFFF00";
      });
    };

    const runs = getRuns(bodyP);
    let cursor = 0;
    const runRanges = runs.map(r => {
      const value = runText(r);
      const item = {run:r, text:value, start:cursor, end:cursor + value.length};
      cursor += value.length;
      return item;
    });

    const full = runRanges.map(x => x.text).join("");
    const containerAnchor = "груз/ контейнер №";
    const jdnAnchor = "накладной СМГС) №";
    const ci = full.indexOf(containerAnchor);
    const ji = full.indexOf(jdnAnchor);

    if (ci < 0 || ji < 0 || ji <= ci) {
      throw Error("Не удалось определить места контейнера и ЖДН в шаблоне.");
    }

    // Put text into every <w:t> of the selected run and clear all
    // additional yellow runs belonging to the same placeholder.
    const putInRun = (r, value) => {
      const ts = Array.from(r.getElementsByTagNameNS(NS, "t"));
      if (!ts.length) return;
      ts[0].textContent = value;
      for (let i = 1; i < ts.length; i++) ts[i].textContent = "";
    };

    const clearRun = r => {
      const ts = Array.from(r.getElementsByTagNameNS(NS, "t"));
      for (const t of ts) t.textContent = "";
    };

    // Container placeholder ends immediately before the comma.
    const containerStart = ci + containerAnchor.length;
    const commaPos = full.indexOf(",", containerStart);
    const containerEnd = commaPos >= 0 && commaPos < ji ? commaPos : ji;

    const containerRuns = runRanges
      .filter(x =>
        x.start >= containerStart &&
        x.start < containerEnd &&
        isYellowRun(x.run)
      );

    if (!containerRuns.length) {
      throw Error("В шаблоне не найдено жёлтое поле номера контейнера.");
    }

    putInRun(containerRuns[0].run, container);
    for (let i = 1; i < containerRuns.length; i++) clearRun(containerRuns[i].run);

    // JDN placeholder is the yellow run after the explicit JDN marker,
    // ending at the following comma.
    const jdnStart = ji + jdnAnchor.length;
    const jdnComma = full.indexOf(",", jdnStart);
    const jdnEnd = jdnComma >= 0 ? jdnComma : full.length;

    const jdnRuns = runRanges
      .filter(x =>
        x.start >= jdnStart &&
        x.start < jdnEnd &&
        isYellowRun(x.run)
      );

    if (!jdnRuns.length) {
      throw Error("В шаблоне не найдено жёлтое поле номера ЖДН.");
    }

    putInRun(jdnRuns[0].run, jdn);
    for (let i = 1; i < jdnRuns.length; i++) clearRun(jdnRuns[i].run);

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
