const input = document.getElementById("input");
const stats = document.getElementById("stats");
const clearBtn = document.getElementById("clear");
const generateBtn = document.getElementById("generate");
const status = document.getElementById("status");

const aliases = {
  "номер ктк": ["номер ктк", "ктк"],
  "ref": ["ref", "реф", "букинг"],
  "статус": ["статус", "клиент", "получатель"],
  "фактический объем, м3": [
    "фактический объем, м3",
    "фактический объем м3",
    "объем, м3",
    "объем м3"
  ],
  "вес, кг": ["вес, кг", "вес кг", "вес"],
  "кол-во мест": ["кол-во мест", "количество мест", "мест"]
};

const targetNames = {
  "номер ктк": ["номер ктк", "ктк"],
  "ref": ["букинг", "ref", "реф"],
  "статус": ["клиент", "статус"],
  "фактический объем, м3": ["объем, м3", "объем м3", "фактический объем, м3"],
  "вес, кг": ["вес, кг", "вес кг", "вес"],
  "кол-во мест": ["кол-во мест", "количество мест", "мест"]
};

function norm(value) {
  return String(value ?? "")
    .replace(/\uFEFF/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/\r?\n/g, " ")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function parsePastedTable(text) {
  return text
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map(line => line.split("\t").map(cell => cell.trim()))
    .filter(row => row.some(cell => cell !== ""));
}

function findColumn(headers, name, dictionary = aliases) {
  const normalizedHeaders = headers.map(norm);
  const options = (dictionary[name] || [name]).map(norm);

  let index = normalizedHeaders.findIndex(h => options.includes(h));
  if (index !== -1) return index;

  index = normalizedHeaders.findIndex(h =>
    h && options.some(option => h.includes(option) || option.includes(h))
  );

  return index;
}

function detectColumns(headers) {
  const required = [
    "номер ктк",
    "ref",
    "статус",
    "фактический объем, м3",
    "вес, кг",
    "кол-во мест"
  ];

  const result = {};
  const missing = [];

  for (const name of required) {
    const index = findColumn(headers, name);
    if (index === -1) missing.push(name);
    else result[name] = index;
  }

  return { result, missing };
}

/*
 * Последняя строка в исходной таблице — итоговая:
 * 63,202 | 63,202 | 15135,54 | 0 | 463 ...
 * Её не нужно переносить как отдельную запись.
 */
function isTotalRow(row) {
  const values = row.map(norm).filter(Boolean);
  if (!values.length) return true;

  const joined = values.join(" ");
  const hasTotalWord = /итого|всего|total/i.test(joined);

  // В нашей таблице итоговая строка не содержит КТК/REF,
  // но содержит итоговые числа.
  const first = norm(row[0]);
  const ref = norm(row[10]);

  return hasTotalWord || (!first && !ref);
}

function getDataRows(rows) {
  return rows.slice(1).filter(row =>
    row.some(value => String(value ?? "").trim() !== "") &&
    !isTotalRow(row)
  );
}

function updateState() {
  const rows = parsePastedTable(input.value);

  if (!rows.length) {
    stats.textContent = "Данные не введены";
    generateBtn.disabled = true;
    return;
  }

  const headers = rows[0];
  const dataRows = getDataRows(rows);
  const { missing } = detectColumns(headers);

  if (missing.length) {
    stats.textContent =
      `${dataRows.length} строк · не найдены: ${missing.join(", ")}`;
    generateBtn.disabled = true;
  } else {
    stats.textContent =
      `${dataRows.length} записей · ${headers.length} столбцов · все необходимые поля найдены`;
    generateBtn.disabled = dataRows.length === 0;
  }

  status.innerHTML = "";
}

input.addEventListener("input", updateState);

clearBtn.addEventListener("click", () => {
  input.value = "";
  updateState();
  input.focus();
});

async function loadTemplate() {
  // В текущей структуре GitHub файл «проработка.xlsx» лежит
  // в КОРНЕ репозитория. Сначала пробуем этот путь.
  const paths = [
    "проработка.xlsx",
    "templates/проработка.xlsx"
  ];

  for (const path of paths) {
    const response = await fetch(path, { cache: "no-store" });
    if (response.ok) {
      return new Uint8Array(await response.arrayBuffer());
    }
  }

  throw new Error(
    "Не найден шаблон «проработка.xlsx». Проверьте, что файл находится в корне репозитория."
  );
}

function writeCell(sheet, row, col, value) {
  const address = XLSX.utils.encode_cell({ r: row, c: col });

  if (value === null || value === undefined || value === "") {
    sheet[address] = { t: "s", v: "" };
    return;
  }

  const text = String(value).replace(/\u00A0/g, " ").trim();

  // Числа с русской десятичной запятой превращаем в числа Excel.
  const numeric = text.replace(/\s/g, "").replace(",", ".");
  if (numeric !== "" && /^-?\d+(?:\.\d+)?$/.test(numeric)) {
    sheet[address] = { t: "n", v: Number(numeric) };
  } else {
    sheet[address] = { t: "s", v: text };
  }
}

function clearTargetRows(sheet, firstRow, lastRow, columns) {
  for (let r = firstRow; r <= lastRow; r++) {
    for (const c of columns) {
      const address = XLSX.utils.encode_cell({ r, c });
      if (sheet[address]) {
        sheet[address].v = "";
        sheet[address].w = "";
      }
    }
  }
}

async function buildDocument() {
  const rows = parsePastedTable(input.value);

  if (!rows.length) {
    throw new Error("Сначала вставьте данные из Excel.");
  }

  const headers = rows[0];
  const detected = detectColumns(headers);

  if (detected.missing.length) {
    throw new Error(
      `Не найдены необходимые поля: ${detected.missing.join(", ")}`
    );
  }

  const data = getDataRows(rows);

  if (!data.length) {
    throw new Error("Не найдено ни одной записи для формирования документа.");
  }

  const templateBytes = await loadTemplate();

  const workbook = XLSX.read(templateBytes, {
    type: "array",
    cellStyles: true
  });

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet || !sheet["!ref"]) {
    throw new Error("В шаблоне отсутствует рабочий лист с данными.");
  }

  const range = XLSX.utils.decode_range(sheet["!ref"]);

  // Заголовки шаблона ищем по всей первой строке,
  // поэтому положение столбцов в шаблоне не имеет значения.
  const templateHeaders = [];

  for (let c = range.s.c; c <= range.e.c; c++) {
    const address = XLSX.utils.encode_cell({ r: range.s.r, c });
    templateHeaders.push(sheet[address]?.v ?? "");
  }

  const targetColumns = {};

  for (const sourceName of Object.keys(targetNames)) {
    const index = findColumn(templateHeaders, sourceName, targetNames);

    if (index === -1) {
      throw new Error(
        `В шаблоне не найдена колонка для поля «${sourceName}».`
      );
    }

    targetColumns[sourceName] = index;
  }

  const firstDataRow = range.s.r + 1;

  // Очищаем старые значения в используемых столбцах шаблона.
  clearTargetRows(
    sheet,
    firstDataRow,
    range.e.r,
    Object.values(targetColumns)
  );

  // Записываем только реальные строки из вставленной таблицы.
  data.forEach((row, index) => {
    const targetRow = firstDataRow + index;

    for (const sourceName of Object.keys(targetNames)) {
      const sourceColumn = detected.result[sourceName];
      const targetColumn = targetColumns[sourceName];

      writeCell(
        sheet,
        targetRow,
        targetColumn,
        row[sourceColumn] ?? ""
      );
    }
  });

  // Расширяем диапазон листа, если строк стало больше, чем было в шаблоне.
  const lastDataRow = firstDataRow + data.length - 1;

  if (lastDataRow > range.e.r) {
    sheet["!ref"] = XLSX.utils.encode_range({
      s: range.s,
      e: { r: lastDataRow, c: range.e.c }
    });
  }

  const output = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
    cellStyles: true
  });

  return {
    blob: new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }),
    count: data.length
  };
}

let generatedUrl = null;

function showDownload(blob, count) {
  if (generatedUrl) {
    URL.revokeObjectURL(generatedUrl);
  }

  generatedUrl = URL.createObjectURL(blob);

  status.innerHTML = `
    <div class="result">
      <strong>✓ Документ готов</strong>
      <span>Сформировано записей: ${count}</span>
      <a class="download-link"
         href="${generatedUrl}"
         download="Проработка_готово.xlsx">
        Скачать Проработка_готово.xlsx
      </a>
    </div>
  `;
}

generateBtn.addEventListener("click", async () => {
  generateBtn.disabled = true;
  generateBtn.textContent = "Формируем…";
  status.textContent = "";

  try {
    const result = await buildDocument();
    showDownload(result.blob, result.count);
  } catch (error) {
    console.error(error);

    status.innerHTML = `
      <div class="error">
        <strong>Не удалось сформировать документ</strong>
        <span>${error.message}</span>
      </div>
    `;
  } finally {
    generateBtn.textContent = "Сформировать документ";
    updateState();
  }
});
