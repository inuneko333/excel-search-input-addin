const KEY="input_addin_v4";
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

Office.onReady(()=>{
  const $=id=>document.getElementById(id);
  const ui={
    term:$('searchTerm'), btn:$('searchBtn'), status:$('status'), targetRow:$('targetRow'),
    pageNum:$('pageNum'), pageMeta:$('pageMeta'),
    forms:$('forms'),
    searchCol:$('searchCol'), skipRows:$('skipRows'), pageSize:$('pageSize'),
    addForm:$('addForm'), removeForm:$('removeForm'),
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

  function renderForms(){
    ui.forms.innerHTML="";
    st.forms.forEach((f, idx)=>{
      const card=document.createElement('div');
      card.className='formCard';
      card.dataset.id=f.id;

      const top=document.createElement('div'); top.className='formTop';
      const title=document.createElement('div'); title.className='formTitle'; title.textContent=f.title;

      const cfgBtn=document.createElement('button');
      cfgBtn.className='formCfgBtn'; cfgBtn.type='button'; cfgBtn.textContent=f.cfgOpen? '⚙ 設定':'⚙';
      cfgBtn.tabIndex=-1;

      top.appendChild(title); top.appendChild(cfgBtn); card.appendChild(top);

      const row=document.createElement('div'); row.className='formRow';
      const input=document.createElement('input');
      input.className='input'; input.id=`in_${f.id}`; input.placeholder=f.useLen12? '12桁':'入力';
      if(f.useLen12) input.inputMode='numeric';
      input.tabIndex=0;

      const btn=document.createElement('button');
      btn.className='btn tiny writeBtn'; btn.type='button'; btn.textContent='書';
      btn.title='書き込み'; btn.tabIndex=-1;

      row.appendChild(input); row.appendChild(btn); card.appendChild(row);

      const meta=document.createElement('div');
      meta.className='row'; meta.style.justifyContent='space-between'; meta.style.marginTop='6px';
      const left=document.createElement('div'); left.className='small'; left.textContent=`列:${f.col}`;
      const right=document.createElement('div'); right.className='small';
      const counter=document.createElement('span'); counter.className='counter';
      right.appendChild(counter); meta.appendChild(left); meta.appendChild(right); card.appendChild(meta);

      const cfg=document.createElement('div');
      cfg.className='cfg'+(f.cfgOpen? ' open':'');
      cfg.innerHTML = `
        <div class="cfgGrid">
          <div class="cfgLine">
            <span class="small mini">列</span>
            <input class="miniInput" data-k="col" value="${f.col}" />
            <label class="checkbox"><input type="checkbox" data-k="useFixed" ${f.useFixed? 'checked':''}/> <span class="small">固定</span></label>
            <label class="checkbox"><input type="checkbox" data-k="useLen12" ${f.useLen12? 'checked':''}/> <span class="small">12桁</span></label>
          </div>
          <div class="cfgLine">
            <span class="small mini">固定値</span>
            <input class="miniInput" data-k="fixed" value="${(f.fixed||'').replace(/"/g,'&quot;')}" />
            <span class="small">(固定ONの時だけ)</span>
          </div>
        </div>
      `;
      cfg.querySelectorAll('input').forEach(el=>el.tabIndex=-1);
      card.appendChild(cfg);

      function updateCounter(){
        if(!f.useLen12){ counter.textContent=''; counter.classList.remove('bad'); return; }
        const len=input.value.length;
        counter.textContent=`${len}/12`;
        counter.classList.toggle('bad', len!==12);
      }
      input.addEventListener('input', updateCounter); updateCounter();

      cfgBtn.addEventListener('click', ()=>{
        f.cfgOpen=!f.cfgOpen;
        cfg.classList.toggle('open', f.cfgOpen);
        cfgBtn.textContent=f.cfgOpen? '⚙ 設定':'⚙';
        saveState(st);
      });

      const colEl=cfg.querySelector('input[data-k="col"]');
      const fixedEl=cfg.querySelector('input[data-k="fixed"]');
      const useFixedEl=cfg.querySelector('input[data-k="useFixed"]');
      const useLenEl=cfg.querySelector('input[data-k="useLen12"]');

      function refreshCfg(){
        left.textContent=`列:${f.col}`;
        input.placeholder=f.useLen12? '12桁':'入力';
        if(f.useLen12){ input.inputMode='numeric'; } else { input.removeAttribute('inputmode'); }
        updateCounter();
      }

      colEl.addEventListener('change', ()=>{
        const v=colEl.value.toUpperCase().trim();
        if(isValidCol(v)){ f.col=v; saveState(st); refreshCfg(); }
        else colEl.value=f.col;
      });
      fixedEl.addEventListener('change', ()=>{ f.fixed=fixedEl.value||''; saveState(st); });
      useFixedEl.addEventListener('change', ()=>{ f.useFixed=useFixedEl.checked; saveState(st); });
      useLenEl.addEventListener('change', ()=>{ f.useLen12=useLenEl.checked; saveState(st); refreshCfg(); });

      async function doWrite(){
        persistSettings();
        if(!runtime.lastHitRow){ setStatus('先に検索して行を確定してください'); return; }
        const rowNum=runtime.lastHitRow;
        if(!isValidCol(f.col)){ setStatus(`${f.title}: 列指定が不正`); return; }
        const value = f.useFixed ? (f.fixed||'') : (input.value||'');
        if(value===''){ setStatus(`${f.title}: 値が空`); return; }
        if(f.useLen12 && String(value).length!==12){ setStatus(`${f.title}: 12桁ではありません`); return; }
        try{
          await writeCell(f.col, rowNum, value);
          setStatus('');
          ui.targetRow.textContent=`行 ${rowNum}`;
          const nextIdx=(idx+1)%st.forms.length;
          const next=document.getElementById(`in_${st.forms[nextIdx].id}`);
          if(next) next.focus();
        }catch(e){
          console.error(e);
          setStatus('書き込みに失敗（共有/保護/権限を確認）');
        }
      }

      btn.addEventListener('click', doWrite);
      input.addEventListener('keydown', (ev)=>{
        if(ev.key==='Enter'){ ev.preventDefault(); doWrite(); }
      });

      ui.forms.appendChild(card);
    });
  }

  renderForms();

  ui.addForm.addEventListener('click', ()=>{
    if(st.forms.length>=MAX_FORMS) return;
    st.forms.push(defaultForm(st.forms.length));
    saveState(st); renderForms();
  });
  ui.removeForm.addEventListener('click', ()=>{
    if(st.forms.length<=1) return;
    st.forms.pop();
    saveState(st); renderForms();
  });

  [ui.searchCol, ui.skipRows, ui.pageSize].forEach(el=>el.addEventListener('change', persistSettings));

  async function runSearch(){
    persistSettings();
    const term=ui.term.value.trim();
    if(!term){ setStatus('検索値を入力してください'); return; }

    setStatus('検索中…');
    ui.targetRow.textContent='–';
    ui.pageNum.textContent='–';
    ui.pageMeta.textContent='…';

    try{
      await Excel.run(async ctx=>{
        const sh=ctx.workbook.worksheets.getActiveWorksheet();
        const used=sh.getUsedRange();
        used.load(['values','rowCount','columnCount','rowIndex','columnIndex']);
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
          setStatus('ヒットなし');
          ui.pageMeta.textContent='—';
          return;
        }

        runtime.lastHitRow=firstHit;
        sh.getRange(`${st.searchCol}${firstHit}`).select();
        await ctx.sync();

        const page=calcPage(firstHit, st.skipRows, st.pageSize);
        ui.pageNum.textContent=String(page);
        ui.pageMeta.textContent=`行${firstHit}`;
        ui.targetRow.textContent=`行 ${firstHit}`;
        setStatus('');

        const firstInput=document.getElementById(`in_${st.forms[0].id}`);
        if(firstInput) firstInput.focus();
      });
    }catch(e){
      console.error(e);
      setStatus('検索に失敗（保護/共有/権限を確認）');
    }
  }

  function focusSearch(){ ui.term.focus(); ui.term.select(); }

  document.addEventListener('keydown', (e)=>{
    if(e.ctrlKey && e.shiftKey && (e.key==='F' || e.key==='f')){ e.preventDefault(); focusSearch(); }
    if(e.key==='/' && !(e.target && (e.target.tagName==='INPUT' || e.target.tagName==='TEXTAREA'))){
      e.preventDefault(); focusSearch();
    }
  });

  ui.btn.addEventListener('click', runSearch);
  ui.term.addEventListener('focus', (e)=>e.target.select());
  ui.term.addEventListener('keydown', (ev)=>{ if(ev.key==='Enter'){ ev.preventDefault(); runSearch(); } });

  ui.term.focus();
});
