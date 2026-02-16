const KEY="input_addin_v7";
const MAX_FORMS=10;

function isValidCol(s){ return /^[A-Z]{1,3}$/.test((s||"").toUpperCase().trim()); }
function loadState(){ try{ const r=localStorage.getItem(KEY); return r?JSON.parse(r):null; }catch{return null;} }
function saveState(st){ localStorage.setItem(KEY, JSON.stringify(st)); }

function defaultForm(i){
  return { id:(crypto.randomUUID?crypto.randomUUID():String(Date.now()+i)), title:`フォーム${i+1}`, col:"C", useFixed:false, fixed:"", useLen12:false, cfgOpen:false };
}
function normalizeState(st){
  const base=st&&typeof st==="object"?st:{};
  const forms=Array.isArray(base.forms)?base.forms:[defaultForm(0), defaultForm(1)];
  return {
    searchCol: isValidCol(base.searchCol)?base.searchCol.toUpperCase().trim():"A",
    skipRows: Number(base.skipRows)||0,
    pageSize: Math.max(1, Number(base.pageSize)||18),
    forms: forms.slice(0,MAX_FORMS).map((f,idx)=>({
      id: f.id || (crypto.randomUUID?crypto.randomUUID():String(Date.now()+idx)),
      title: (f.title||`フォーム${idx+1}`).toString(),
      col: isValidCol(f.col)?f.col.toUpperCase().trim():"C",
      useFixed: !!f.useFixed,
      fixed: (f.fixed||"").toString(),
      useLen12: !!f.useLen12,
      cfgOpen: !!f.cfgOpen
    }))
  };
}

let runtime={ lastHitRow:null };

function colToIndex(s){
  const t=(s||"A").toUpperCase().trim(); let n=0;
  for(const ch of t){ const c=ch.charCodeAt(0); if(c<65||c>90) return 0; n=n*26+(c-64); }
  return n-1;
}
function calcPage(hitRow, skipRows, pageSize){
  const logical=Math.max(1, hitRow - skipRows);
  return Math.ceil(logical / pageSize);
}
async function writeCell(col,row,value){
  await Excel.run(async ctx=>{
    const sh=ctx.workbook.worksheets.getActiveWorksheet();
    sh.getRange(`${col}${row}`).values=[[value]];
    await ctx.sync();
  });
}

function digitsOnly(s){ return String(s||"").replace(/\D/g,""); }
function format12WithSpaces(raw){
  const d=digitsOnly(raw).slice(0,12);
  return d.replace(/(\d{4})(?=\d)/g, "$1 ").trim();
}
function focusAndSelect(el){
  if(!el) return;
  el.focus();
  try{ el.select(); }catch{}
}

Office.onReady(()=>{
  const $=id=>document.getElementById(id);
  const ui={
    term:$("searchTerm"), btn:$("searchBtn"), status:$("status"),
    pageNum:$("pageNum"), pageMeta:$("pageMeta"),
    forms:$("forms"),
    searchCol:$("searchCol"), skipRows:$("skipRows"), pageSize:$("pageSize"),
    addForm:$("addForm"), removeForm:$("removeForm"),
  };

  let st=normalizeState(loadState()); saveState(st);

  ui.searchCol.value=st.searchCol; ui.skipRows.value=st.skipRows; ui.pageSize.value=st.pageSize;

  function setStatus(msg){ ui.status.textContent = msg || ""; }
  function persistSettings(){
    st.searchCol = isValidCol(ui.searchCol.value)?ui.searchCol.value.toUpperCase().trim():"A";
    st.skipRows = Number(ui.skipRows.value)||0;
    st.pageSize = Math.max(1, Number(ui.pageSize.value)||18);
    saveState(st);
  }

  function makeChip(text, cls){
    const s=document.createElement("span");
    s.className="chip"+(cls?(" "+cls):"");
    s.textContent=text;
    return s;
  }

  function renderForms(){
    ui.forms.innerHTML="";
    st.forms.forEach((f, idx)=>{
      const card=document.createElement("div");
      card.className="formCard";
      card.dataset.id=f.id;

      const top=document.createElement("div"); top.className="formTop";
      const titleWrap=document.createElement("div"); titleWrap.className="formTitle";
      titleWrap.textContent=f.title;

      const chips=document.createElement("div"); chips.className="chips";
      const rebuildChips=()=>{
        chips.innerHTML="";
        chips.appendChild(makeChip(`列 ${f.col}`));
        if(f.useLen12) chips.appendChild(makeChip("12桁"));
        if(f.useFixed) chips.appendChild(makeChip("固定", "fixed"));
      };
      rebuildChips();

      const left=document.createElement("div");
      left.style.display="flex";
      left.style.alignItems="center";
      left.style.gap="8px";
      left.appendChild(titleWrap);
      left.appendChild(chips);

      const cfgBtn=document.createElement("button");
      cfgBtn.className="formCfgBtn";
      cfgBtn.type="button";
      cfgBtn.textContent=f.cfgOpen?"⚙ 設定":"⚙";
      cfgBtn.tabIndex=-1;

      top.appendChild(left);
      top.appendChild(cfgBtn);
      card.appendChild(top);

      const row=document.createElement("div"); row.className="formRow";

      const input=document.createElement("input");
      input.className="input";
      input.id=`in_${f.id}`;
      input.tabIndex=0;

      const writeBtn=document.createElement("button");
      writeBtn.className="btn tiny writeBtn";
      writeBtn.type="button";
      writeBtn.innerHTML='<span>✎</span>';
      writeBtn.title="書き込み";
      writeBtn.tabIndex=-1;

      row.appendChild(input);
      row.appendChild(writeBtn);
      card.appendChild(row);

      const under=document.createElement("div"); under.className="underRow";
      const counter=document.createElement("span"); counter.className="counter small";
      under.appendChild(counter);
      card.appendChild(under);

      function applyFixedUI(){
        if(f.useFixed){
          input.value = f.fixed || "";
          input.readOnly = true;
          input.classList.add("fixed");
          input.placeholder = "";
        }else{
          input.readOnly = false;
          input.classList.remove("fixed");
          input.placeholder = f.useLen12 ? "0000 0000 0000" : "入力";
          if(f.useLen12 && input.value){
            input.value = format12WithSpaces(input.value);
          }
        }
      }

      function updateCounter(){
        if(!f.useLen12){ counter.textContent=""; counter.classList.remove("bad"); return; }
        const raw = digitsOnly(input.value);
        counter.textContent=`${raw.length}/12`;
        counter.classList.toggle("bad", raw.length!==12);
      }

      input.addEventListener("focus", ()=>{
        setTimeout(()=>{ try{ input.select(); }catch{} }, 0);
      });

      input.addEventListener("input", ()=>{
        if(f.useFixed) return;
        if(f.useLen12){
          input.value = format12WithSpaces(input.value);
        }
        updateCounter();
      });
      input.addEventListener("keydown", (ev)=>{
        if(ev.key==="Enter"){ ev.preventDefault(); doWrite(); }
      });

      const cfg=document.createElement("div");
      cfg.className="cfg"+(f.cfgOpen?" open":"");
      cfg.innerHTML = `
        <div class="cfgLine">
          <span class="small">列</span>
          <input class="miniInput" data-k="col" value="${f.col}" />
          <label class="checkbox"><input type="checkbox" data-k="useFixed" ${f.useFixed?"checked":""}/> <span class="small">固定</span></label>
          <label class="checkbox"><input type="checkbox" data-k="useLen12" ${f.useLen12?"checked":""}/> <span class="small">12桁</span></label>
        </div>
        <div class="cfgLine" style="margin-top:8px">
          <span class="small">固定値</span>
          <input class="miniInput" data-k="fixed" value="${(f.fixed||"").replace(/"/g,'&quot;')}" />
          <span class="small">（固定ONで使用）</span>
        </div>
      `;
      cfg.querySelectorAll("input").forEach(el=>el.tabIndex=-1);
      card.appendChild(cfg);

      cfgBtn.addEventListener("click", ()=>{
        f.cfgOpen=!f.cfgOpen;
        cfg.classList.toggle("open", f.cfgOpen);
        cfgBtn.textContent=f.cfgOpen?"⚙ 設定":"⚙";
        saveState(st);
      });

      const colEl=cfg.querySelector('input[data-k="col"]');
      const fixedEl=cfg.querySelector('input[data-k="fixed"]');
      const useFixedEl=cfg.querySelector('input[data-k="useFixed"]');
      const useLenEl=cfg.querySelector('input[data-k="useLen12"]');

      colEl.addEventListener("change", ()=>{
        const v=colEl.value.toUpperCase().trim();
        if(isValidCol(v)){ f.col=v; saveState(st); rebuildChips(); }
        else colEl.value=f.col;
      });
      fixedEl.addEventListener("change", ()=>{
        f.fixed=fixedEl.value||"";
        saveState(st);
        applyFixedUI();
      });
      useFixedEl.addEventListener("change", ()=>{
        f.useFixed=useFixedEl.checked;
        saveState(st);
        rebuildChips();
        applyFixedUI();
        updateCounter();
      });
      useLenEl.addEventListener("change", ()=>{
        f.useLen12=useLenEl.checked;
        saveState(st);
        rebuildChips();
        applyFixedUI();
        updateCounter();
      });

      async function doWrite(){
        persistSettings();
        if(!runtime.lastHitRow){ setStatus("先に検索して行を確定してください"); return; }
        const rowNum=runtime.lastHitRow;
        if(!isValidCol(f.col)){ setStatus(`${f.title}: 列指定が不正`); return; }

        let value = f.useFixed ? (f.fixed||"") : (input.value||"");
        if(value===""){ setStatus(`${f.title}: 値が空`); return; }

        if(f.useLen12){
          const raw=digitsOnly(value);
          if(raw.length!==12){ setStatus(`${f.title}: 12桁ではありません`); return; }
          value = raw;
        }

        try{
          await writeCell(f.col, rowNum, value);
          setStatus("");
          const nextIdx=(idx+1)%st.forms.length;
          const next=document.getElementById(`in_${st.forms[nextIdx].id}`);
          focusAndSelect(next);
        }catch(e){
          console.error(e);
          setStatus("書き込みに失敗（共有/保護/権限を確認）");
        }
      }

      writeBtn.addEventListener("click", doWrite);

      applyFixedUI();
      updateCounter();

      ui.forms.appendChild(card);
    });
  }

  renderForms();

  ui.addForm.addEventListener("click", ()=>{
    if(st.forms.length>=MAX_FORMS) return;
    st.forms.push(defaultForm(st.forms.length));
    saveState(st); renderForms();
  });
  ui.removeForm.addEventListener("click", ()=>{
    if(st.forms.length<=1) return;
    st.forms.pop();
    saveState(st); renderForms();
  });

  [ui.searchCol, ui.skipRows, ui.pageSize].forEach(el=>el.addEventListener("change", persistSettings));

  async function runSearch(){
    persistSettings();
    const term=ui.term.value.trim();
    if(!term){ setStatus("検索値を入力してください"); return; }

    setStatus("検索中…");
    ui.pageNum.textContent="–";
    ui.pageMeta.textContent="…";

    try{
      await Excel.run(async ctx=>{
        const sh=ctx.workbook.worksheets.getActiveWorksheet();
        const used=sh.getUsedRange();
        used.load(["values","rowCount","columnCount","rowIndex","columnIndex"]);
        await ctx.sync();

        const targetCol=colToIndex(st.searchCol);
        const offset=targetCol - used.columnIndex;
        let firstHit=null;

        if(offset>=0 && offset<used.columnCount){
          for(let r=0;r<used.rowCount;r++){
            const rowNum=used.rowIndex + r + 1;
            if(rowNum<=st.skipRows) continue;
            const v=used.values[r][offset];
            if(v==null) continue;
            if(String(v).trim()===term){ firstHit=rowNum; break; }
          }
        }

        if(!firstHit){
          runtime.lastHitRow=null;
          setStatus("ヒットなし");
          ui.pageMeta.textContent="—";
          return;
        }

        runtime.lastHitRow=firstHit;
        sh.getRange(`${st.searchCol}${firstHit}`).select();
        await ctx.sync();

        const page=calcPage(firstHit, st.skipRows, st.pageSize);
        ui.pageNum.textContent=String(page);
        ui.pageMeta.textContent=`行${firstHit}`;
        setStatus("");

        const firstInput=document.getElementById(`in_${st.forms[0].id}`);
        focusAndSelect(firstInput);
      });
    }catch(e){
      console.error(e);
      setStatus("検索に失敗（保護/共有/権限を確認）");
    }
  }

  document.addEventListener("keydown", (e)=>{
    if(e.ctrlKey && (e.key==="F2" || e.code==="F2")){
      e.preventDefault();
      focusAndSelect(ui.term);
    }
  });

  ui.btn.addEventListener("click", runSearch);
  ui.term.addEventListener("focus", (e)=>setTimeout(()=>{ try{ e.target.select(); }catch{} },0));
  ui.term.addEventListener("keydown", (ev)=>{ if(ev.key==="Enter"){ ev.preventDefault(); runSearch(); } });

  ui.term.focus();
});
