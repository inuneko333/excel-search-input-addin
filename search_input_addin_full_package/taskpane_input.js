
const INPUT_ADDIN_SETTINGS_KEY = "searchInputAddinSettings_v1";

function loadSettings() {
  try {
    const raw = window.localStorage.getItem(INPUT_ADDIN_SETTINGS_KEY);
    if (!raw) {
      return {
        searchColumn: "A",
        skipRows: 0,
        form1: { column: "N", fixed: "", hidden: false },
        form2: { column: "J", fixed: "", hidden: false },
        code12: { column: "K", hidden: false },
      };
    }
    const obj = JSON.parse(raw);
    return {
      searchColumn: obj.searchColumn || "A",
      skipRows: Number(obj.skipRows) || 0,
      form1: {
        column: (obj.form1 && obj.form1.column) || "N",
        fixed: (obj.form1 && obj.form1.fixed) || "",
        hidden: !!(obj.form1 && obj.form1.hidden),
      },
      form2: {
        column: (obj.form2 && obj.form2.column) || "J",
        fixed: (obj.form2 && obj.form2.fixed) || "",
        hidden: !!(obj.form2 && obj.form2.hidden),
      },
      code12: {
        column: (obj.code12 && obj.code12.column) || "K",
        hidden: !!(obj.code12 && obj.code12.hidden),
      },
    };
  } catch {
    return {
      searchColumn: "A",
      skipRows: 0,
      form1: { column: "N", fixed: "", hidden: false },
      form2: { column: "J", fixed: "", hidden: false },
      code12: { column: "K", hidden: false },
    };
  }
}

function saveSettings(settings) {
  try {
    window.localStorage.setItem(INPUT_ADDIN_SETTINGS_KEY, JSON.stringify(settings));
  } catch {
    // noop
  }
}

function columnLetterToIndex(letter) {
  if (!letter) return 0;
  const s = letter.toUpperCase().trim();
  let col = 0;
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code < 65 || code > 90) {
      return 0;
    }
    col = col * 26 + (code - 64);
  }
  return col - 1;
}

function initInputAddinUi() {
  const searchInput = document.getElementById("searchTerm");
  const searchButton = document.getElementById("searchButton");
  const statusMessage = document.getElementById("statusMessage");
  const currentRowLabel = document.getElementById("currentRowLabel");

  const form1Input = document.getElementById("form1Input");
  const form1WriteButton = document.getElementById("form1Write");
  const form1Note = document.getElementById("form1Note");
  const form1Body = document.getElementById("form1Body");
  const toggleForm1 = document.getElementById("toggleForm1");

  const form2Input = document.getElementById("form2Input");
  const form2WriteButton = document.getElementById("form2Write");
  const form2Note = document.getElementById("form2Note");
  const form2Body = document.getElementById("form2Body");
  const toggleForm2 = document.getElementById("toggleForm2");

  const code12Input = document.getElementById("code12Input");
  const code12WriteButton = document.getElementById("code12Write");
  const code12Counter = document.getElementById("code12Counter");
  const code12Body = document.getElementById("code12Body");
  const toggleCode12 = document.getElementById("toggleCode12");

  const toggleSettingsButton = document.getElementById("toggleSettings");
  const settingsPanel = document.getElementById("settingsPanel");

  const searchColumnInput = document.getElementById("searchColumnInput");
  const skipRowsInput = document.getElementById("skipRowsInput");
  const form1ColumnInput = document.getElementById("form1ColumnInput");
  const form1FixedInput = document.getElementById("form1FixedInput");
  const form2ColumnInput = document.getElementById("form2ColumnInput");
  const form2FixedInput = document.getElementById("form2FixedInput");
  const code12ColumnInput = document.getElementById("code12ColumnInput");

  let settings = loadSettings();

  searchColumnInput.value = settings.searchColumn;
  skipRowsInput.value = settings.skipRows;
  form1ColumnInput.value = settings.form1.column;
  form1FixedInput.value = settings.form1.fixed;
  form2ColumnInput.value = settings.form2.column;
  form2FixedInput.value = settings.form2.fixed;
  code12ColumnInput.value = settings.code12.column;

  function applyFormVisibility() {
    if (settings.form1.hidden) {
      form1Body.classList.add("hidden");
      toggleForm1.textContent = "▶ 表示";
    } else {
      form1Body.classList.remove("hidden");
      toggleForm1.textContent = "▼ 非表示";
    }

    if (settings.form2.hidden) {
      form2Body.classList.add("hidden");
      toggleForm2.textContent = "▶ 表示";
    } else {
      form2Body.classList.remove("hidden");
      toggleForm2.textContent = "▼ 非表示";
    }

    if (settings.code12.hidden) {
      code12Body.classList.add("hidden");
      toggleCode12.textContent = "▶ 表示";
    } else {
      code12Body.classList.remove("hidden");
      toggleCode12.textContent = "▼ 非表示";
    }
  }

  function applyFixedNotes() {
    form1Note.textContent = settings.form1.fixed
      ? `固定値: 「${settings.form1.fixed}」が書き込まれます`
      : "入力した値が書き込まれます";
    form2Note.textContent = settings.form2.fixed
      ? `固定値: 「${settings.form2.fixed}」が書き込まれます`
      : "入力した値が書き込まれます";
  }

  applyFormVisibility();
  applyFixedNotes();

  function updateSettingsFromInputs() {
    settings = {
      searchColumn: (searchColumnInput.value || "A").trim(),
      skipRows: Number(skipRowsInput.value) || 0,
      form1: {
        column: (form1ColumnInput.value || "N").trim(),
        fixed: form1FixedInput.value || "",
        hidden: settings.form1.hidden,
      },
      form2: {
        column: (form2ColumnInput.value || "J").trim(),
        fixed: form2FixedInput.value || "",
        hidden: settings.form2.hidden,
      },
      code12: {
        column: (code12ColumnInput.value || "K").trim(),
        hidden: settings.code12.hidden,
      },
    };
    saveSettings(settings);
    applyFixedNotes();
  }

  let settingsVisible = false;
  toggleSettingsButton.addEventListener("click", () => {
    settingsVisible = !settingsVisible;
    if (settingsVisible) {
      settingsPanel.classList.remove("hidden");
      toggleSettingsButton.textContent = "▼ 詳細設定を隠す";
    } else {
      settingsPanel.classList.add("hidden");
      toggleSettingsButton.textContent = "▶ 詳細設定を表示";
    }
  });

  toggleForm1.addEventListener("click", () => {
    settings.form1.hidden = !settings.form1.hidden;
    saveSettings(settings);
    applyFormVisibility();
  });

  toggleForm2.addEventListener("click", () => {
    settings.form2.hidden = !settings.form2.hidden;
    saveSettings(settings);
    applyFormVisibility();
  });

  toggleCode12.addEventListener("click", () => {
    settings.code12.hidden = !settings.code12.hidden;
    saveSettings(settings);
    applyFormVisibility();
  });

  function updateCode12Status() {
    const len = code12Input.value.length;
    code12Counter.textContent = `${len} / 12`;
    if (len === 12) {
      code12Counter.classList.remove("code12-counter-error");
    } else {
      code12Counter.classList.add("code12-counter-error");
    }
  }
  code12Input.addEventListener("input", updateCode12Status);
  updateCode12Status();

  searchInput.addEventListener("focus", (ev) => {
    ev.target.select();
  });

  async function runSearch() {
    const termRaw = searchInput.value.trim();
    if (!termRaw) {
      statusMessage.textContent = "検索値を入力してください。";
      return;
    }

    updateSettingsFromInputs();

    const searchValue = termRaw;
    const columnLetter = settings.searchColumn;
    const skipRows = settings.skipRows;

    statusMessage.textContent = "検索中…";
    currentRowLabel.textContent = "–";

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
        await context.sync();

        const values = usedRange.values;
        const rowCount = usedRange.rowCount;
        const colCount = usedRange.columnCount;
        const startRowIndex = usedRange.rowIndex;
        const startColIndex = usedRange.columnIndex;

        const targetColIndex = columnLetterToIndex(columnLetter);
        const colOffset = targetColIndex - startColIndex;

        let firstHitRow = null;

        if (colOffset >= 0 && colOffset < colCount) {
          for (let r = 0; r < rowCount; r++) {
            const actualRowNumber = startRowIndex + r + 1;
            if (actualRowNumber <= skipRows) continue;

            const cellValue = values[r][colOffset];
            if (cellValue === null || cellValue === undefined) continue;

            const cellText = String(cellValue).trim();
            if (cellText === searchValue) {
              firstHitRow = actualRowNumber;
              break;
            }
          }
        }

        if (!firstHitRow) {
          statusMessage.textContent = "ヒットなし";
          currentRowLabel.textContent = "–";
          return;
        }

        const upperColumnLetter = (columnLetter || "A").toUpperCase().trim() || "A";
        const address = `${upperColumnLetter}${firstHitRow}`;
        const range = sheet.getRange(address);
        range.select();

        currentRowLabel.textContent = `行 ${firstHitRow}`;
        statusMessage.textContent = "";
        await context.sync();
      });
    } catch (error) {
      console.error(error);
      statusMessage.textContent = "検索中にエラーが発生しました。";
    }
  }

  async function writeToActiveRow(columnLetter, value) {
    if (!columnLetter) return;
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const activeCell = sheet.getActiveCell();
        activeCell.load("rowIndex");
        await context.sync();

        const rowNumber = activeCell.rowIndex + 1;
        const upperColumnLetter = columnLetter.toUpperCase().trim();
        const address = `${upperColumnLetter}${rowNumber}`;
        const target = sheet.getRange(address);
        target.values = [[value]];
        await context.sync();

        currentRowLabel.textContent = `行 ${rowNumber}`;
      });
    } catch (error) {
      console.error(error);
      statusMessage.textContent = "書き込み中にエラーが発生しました。";
    }
  }

  form1WriteButton.addEventListener("click", () => {
    updateSettingsFromInputs();
    const val = settings.form1.fixed || form1Input.value;
    if (val === "") {
      statusMessage.textContent = "フォーム1の値が空です。";
      return;
    }
    writeToActiveRow(settings.form1.column, val);
  });

  form2WriteButton.addEventListener("click", () => {
    updateSettingsFromInputs();
    const val = settings.form2.fixed || form2Input.value;
    if (val === "") {
      statusMessage.textContent = "フォーム2の値が空です。";
      return;
    }
    writeToActiveRow(settings.form2.column, val);
  });

  code12WriteButton.addEventListener("click", () => {
    updateSettingsFromInputs();
    const val = code12Input.value.trim();
    if (val.length !== 12) {
      statusMessage.textContent = "12桁コードが12文字ではありません。";
      return;
    }
    writeToActiveRow(settings.code12.column, val);
  });

  searchButton.addEventListener("click", () => {
    runSearch();
  });
  searchInput.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      runSearch();
    }
  });

  searchInput.focus();
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initInputAddinUi();
  }
});
