const KEY="search_input_addin_v3";
const MAX_FORMS=10;

function isValidCol(s){
  const t=(s||"").toUpperCase().trim();
  return /^[A-Z]{1,3}$/.test(t);
}

function loadState(){
  try{
    const raw=localStorage.getItem(KEY);
    if(!raw) return null;
    return JSON.parse(raw);
  }catch{return null;}
}

function saveState(st){
  localStorage.setItem(KEY, JSON.stringify(st));
}

function defaultForm(i){
  return {
    id: (crypto.randomUUID?crypto.randomUUID():String(Date.now()+i)),
    title: `フォーム${i+1}`,
    col: "N",
    useFixed: false,
    fixed: "",
    useLen12: false
  };
}

function normalizeState(st){
  const base = st && typeof st==="object" ? st : {};
  const forms = Array.isArray(base.forms) ? base.forms : [defaultForm(0), defaultForm(1)];
  return {
    searchCol: isValidCol(base.searchCol) ? base.searchCol.toUpperCase().trim() : "A",
    skipRows: Number(base.skipRows)||0,
    pageSize: Math.max(1, Number(base.pageSize)||18),
    forms: forms.slice(0,MAX_FORMS).map((f,idx)=>({
      id: f.id || (crypto.randomUUID?crypto.randomUUID():String(Date.now()+idx)),
      title: (f.title||`フォーム${idx+1}`).toString(),
      col: isValidCol(f.col) ? f.col.toUpperCase().trim() : "N",
      useFixed: !!f.useFixed,
      fixed: (f.fixed||"").toString(),
      useLen12: !!f.useLen12
    }))
  };
}

let runtime = { lastHitRow: null };

async function writeCell(col, row, value){
  await Excel.run(async (ctx)=>{
    const sh = ctx.workbook.worksheets.getActiveWorksheet();
    sh.getRange(`${col}${row}`).values = [[value]];
    await ctx.sync();
  });
}

function colToIndex(s){
  const t=(s||"A").toUpperCase().trim();
  let n=0;
  for(const ch of t){
    const c=ch.charCodeAt(0);
    if(c<65||c>90) return 0;
    n = n*26 + (c-64);
  }
  return n-1;
}

function calcPage(hitRow, skipRows, pageSize){
  const logical = Math.max(1, hitRow - skipRows);
  return Math.ceil(logical / pageSize);
}

Office.onReady(()=>{
  const $ = (id)=>document.getElementById(id);

  const ui = {
    searchTerm: $("searchTerm"),
    searchBtn: $("searchBtn"),
    status: $("status"),
    targetRow: $("targetRow"),
    pageNum: $("pageNum"),
    pageMeta: $("pageMeta"),
    forms: $("forms"),
    searchCol: $("searchCol"),
    skipRows: $("skipRows"),
    pageSize: $("pageSize"),
    addForm: $("addForm"),
    removeForm: $("removeForm"),
  };

  let st = normalizeState(loadState());
  saveState(st);

  ui.searchCol.value = st.searchCol;
  ui.skipRows.value = st.skipRows;
  ui.pageSize.value = st.pageSize;

  function persistFromSettings(){
    st.searchCol = isValidCol(ui.searchCol.value) ? ui.searchCol.value.toUpperCase().trim() : "A";
    st.skipRows = Number(ui.skipRows.value)||0;
    st.pageSize = Math.max(1, Number(ui.pageSize.value)||18);
    saveState(st);
  }

  function setStatus(msg){ ui.status.textContent = msg || ""; }

  function renderForms(){
    ui.forms.innerHTML = "";
    st.forms.forEach((f, idx)=>{
      const wrap = document.createElement("div");
      wrap.className = "formCard";
      wrap.dataset.formId = f.id;

      const title = document.createElement("div");
      title.className = "formTitle";
      title.textContent = f.title;
      wrap.appendChild(title);

      const row = document.createElement("div");
      row.className = "row";

      const input = document.createElement("input");
      input.className = "input";
      input.id = `formInput_${f.id}`;
      input.placeholder = f.useLen12 ? "12桁" : "入力";
      if(f.useLen12){ input.inputMode = "numeric"; }
      input.tabIndex = 0;

      const writeBtn = document.createElement("button");
      writeBtn.className = "btn";
      writeBtn.type = "button";
      writeBtn.textContent = "書き込み";
      writeBtn.tabIndex = -1; // tabは入力だけに集中

      row.appendChild(input);
      row.appendChild(writeBtn);
      wrap.appendChild(row);

      const metaRow = document.createElement("div");
      metaRow.className = "row";
      metaRow.style.justifyContent = "space-between";
      metaRow.style.marginTop = "8px";

      const left = document.createElement("div");
      left.className = "small";
      left.textContent = `列: ${f.col}`;

      const right = document.createElement("div");
      right.className = "small";

      const counter = document.createElement("span");
      counter.className = "counter";
      counter.textContent = f.useLen12 ? "0 / 12" : "";
      right.appendChild(counter);

      metaRow.appendChild(left);
      metaRow.appendChild(right);
      wrap.appendChild(metaRow);

      const hr = document.createElement("hr");
      hr.className = "sep";
      wrap.appendChild(hr);

      const colLabel = document.createElement("label");
      colLabel.className = "label";
      colLabel.textContent = "書き込み列";

      const colInput = document.createElement("input");
      colInput.className = "input";
      colInput.value = f.col;
      colInput.placeholder = "N";
      colInput.tabIndex = -1;

      const cbFixed = document.createElement("label");
      cbFixed.className = "checkbox";
      const cbFixedInput = document.createElement("input");
      cbFixedInput.type = "checkbox";
      cbFixedInput.checked = f.useFixed;
      cbFixedInput.tabIndex = -1;
      const cbFixedText = document.createElement("span");
      cbFixedText.textContent = "固定値を使う";
      cbFixed.appendChild(cbFixedInput);
      cbFixed.appendChild(cbFixedText);

      const fixedLabel = document.createElement("label");
      fixedLabel.className = "label";
      fixedLabel.textContent = "固定値";

      const fixedInput = document.createElement("input");
      fixedInput.className = "input";
      fixedInput.value = f.fixed;
      fixedInput.placeholder = "例: 2026/02/10";
      fixedInput.tabIndex = -1;

      const cbLen = document.createElement("label");
      cbLen.className = "checkbox";
      const cbLenInput = document.createElement("input");
      cbLenInput.type = "checkbox";
      cbLenInput.checked = f.useLen12;
      cbLenInput.tabIndex = -1;
      const cbLenText = document.createElement("span");
      cbLenText.textContent = "12桁チェック（12文字でOK）";
      cbLen.appendChild(cbLenInput);
      cbLen.appendChild(cbLenText);

      fixedInput.disabled = !cbFixedInput.checked;
      fixedInput.classList.toggle("muted", fixedInput.disabled);

      wrap.appendChild(colLabel);
      wrap.appendChild(colInput);
      wrap.appendChild(cbFixed);
      wrap.appendChild(fixedLabel);
      wrap.appendChild(fixedInput);
      wrap.appendChild(cbLen);

      function updateCounter(){
        if(!cbLenInput.checked){ counter.textContent=""; counter.classList.remove("bad"); return; }
        const len = input.value.length;
        counter.textContent = `${len} / 12`;
        counter.classList.toggle("bad", len!==12);
      }
      input.addEventListener("input", updateCounter);
      updateCounter();

      colInput.addEventListener("change", ()=>{
        const v = colInput.value.toUpperCase().trim();
        if(isValidCol(v)){
          f.col = v;
          left.textContent = `列: ${f.col}`;
          saveState(st);
        } else {
          colInput.value = f.col;
        }
      });
      cbFixedInput.addEventListener("change", ()=>{
        f.useFixed = cbFixedInput.checked;
        fixedInput.disabled = !f.useFixed;
        fixedInput.classList.toggle("muted", fixedInput.disabled);
        saveState(st);
      });
      fixedInput.addEventListener("change", ()=>{
        f.fixed = fixedInput.value || "";
        saveState(st);
      });
      cbLenInput.addEventListener("change", ()=>{
        f.useLen12 = cbLenInput.checked;
        input.placeholder = f.useLen12 ? "12桁" : "入力";
        updateCounter();
        saveState(st);
      });

      async function doWrite(){
        persistFromSettings();
        if(!runtime.lastHitRow){ setStatus("先に検索して行を確定してください"); return; }
        const rowNum = runtime.lastHitRow;
        const col = f.col;
        if(!isValidCol(col)){ setStatus(`${f.title}: 列指定が不正`); return; }

        const value = f.useFixed ? (f.fixed || "") : (input.value || "");
        if(value === ""){ setStatus(`${f.title}: 値が空`); return; }
        if(f.useLen12 && String(value).length !== 12){ setStatus(`${f.title}: 12桁ではありません`); return; }

        try{
          await writeCell(col, rowNum, value);
          setStatus("");
          ui.targetRow.textContent = `行 ${rowNum}`;
          const nextIdx = (idx + 1) % st.forms.length;
          const next = document.getElementById(`formInput_${st.forms[nextIdx].id}`);
          if(next) next.focus();
        }catch(e){
          console.error(e);
          setStatus("書き込みに失敗（共有/保護/権限を確認）");
        }
      }

      writeBtn.addEventListener("click", doWrite);
      input.addEventListener("keydown", (ev)=>{
        if(ev.key === "Enter"){ ev.preventDefault(); doWrite(); }
      });

      ui.forms.appendChild(wrap);
    });
  }

  renderForms();

  ui.addForm.addEventListener("click", ()=>{
    if(st.forms.length >= MAX_FORMS) return;
    st.forms.push(defaultForm(st.forms.length));
    saveState(st);
    renderForms();
  });

  ui.removeForm.addEventListener("click", ()=>{
    if(st.forms.length <= 1) return;
    st.forms.pop();
    saveState(st);
    renderForms();
  });

  [ui.searchCol, ui.skipRows, ui.pageSize].forEach(el=>el.addEventListener("change", persistFromSettings));

  async function runSearch(){
    persistFromSettings();
    const term = ui.searchTerm.value.trim();
    if(!term){ setStatus("検索値を入力してください"); return; }

    setStatus("検索中…");
    ui.targetRow.textContent = "–";
    ui.pageNum.textContent = "–";
    ui.pageMeta.textContent = "";

    try{
      await Excel.run(async (ctx)=>{
        const sh = ctx.workbook.worksheets.getActiveWorksheet();
        const used = sh.getUsedRange();
        used.load(["values","rowCount","columnCount","rowIndex","columnIndex"]);
        await ctx.sync();

        const targetCol = colToIndex(st.searchCol);
        const offset = targetCol - used.columnIndex;

        let firstHitRow = null;
        if(offset >= 0 && offset < used.columnCount){
          for(let r=0; r<used.rowCount; r++){
            const rowNum = used.rowIndex + r + 1;
            if(rowNum <= st.skipRows) continue;
            const v = used.values[r][offset];
            if(v == null) continue;
            if(String(v).trim() === term){ firstHitRow = rowNum; break; }
          }
        }

        if(!firstHitRow){
          runtime.lastHitRow = null;
          setStatus("ヒットなし");
          ui.pageMeta.textContent = "—";
          return;
        }

        runtime.lastHitRow = firstHitRow;
        const col = st.searchCol.toUpperCase().trim();
        sh.getRange(`${col}${firstHitRow}`).select();
        await ctx.sync();

        const page = calcPage(firstHitRow, st.skipRows, st.pageSize);
        ui.pageNum.textContent = String(page);
        ui.pageMeta.textContent = `先頭ヒット: 行 ${firstHitRow}`;
        ui.targetRow.textContent = `行 ${firstHitRow}`;
        setStatus("");

        const firstInput = document.getElementById(`formInput_${st.forms[0].id}`);
        if(firstInput) firstInput.focus();
      });
    }catch(e){
      console.error(e);
      setStatus("検索に失敗（保護/共有/権限を確認）");
    }
  }

  ui.searchBtn.addEventListener("click", runSearch);
  ui.searchTerm.addEventListener("focus", (e)=>e.target.select());
  ui.searchTerm.addEventListener("keydown", (ev)=>{
    if(ev.key === "Enter"){ ev.preventDefault(); runSearch(); }
  });

  ui.searchTerm.focus();
});
