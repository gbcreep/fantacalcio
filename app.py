<!DOCTYPE html>
<html lang="it">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Fantacalcio — Gestione Asta (Upload Excel)</title>
  <!-- SheetJS per leggere/scrivere Excel (CDN + fallback) -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script>if(typeof XLSX==='undefined'){document.write('<script src="https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js"><\/script>')}</script>
  <style>
    :root{ --bg:#f7f7f9; --card:#ffffff; --ink:#1b1f23; --muted:#6b7280; --brand:#104C97; --ok:#159947; --warn:#CE3262; --line:#e5e7eb; }
    *{box-sizing:border-box}
    body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.4 system-ui,-apple-system,Segoe UI,Roboto,Ubuntu,Cantarell,Noto Sans,Arial}
    .wrap{max-width:1200px;margin:24px auto;padding:0 16px}
    h1{font-size:24px;margin:0 0 8px}
    .sub{color:var(--muted);margin-bottom:12px}
    .row{display:flex;gap:16px;flex-wrap:wrap}
    .card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px;box-shadow:0 1px 2px rgba(0,0,0,.04)}
    .card h2{font-size:18px;margin:0 0 12px;color:var(--brand)}
    label{font-weight:600}
    input[type="text"], input[type="number"], textarea, select{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:10px;background:#fff}
    textarea{min-height:120px;resize:vertical}
    .btn{appearance:none;border:none;border-radius:12px;background:var(--brand);color:#fff;padding:10px 14px;font-weight:700;cursor:pointer}
    .btn.secondary{background:#3e2a18}
    .btn.ghost{background:#fff;color:var(--brand);border:1px solid var(--brand)}
    .btn:disabled{opacity:.5;cursor:not-allowed}
    .pill{display:inline-flex;align-items:center;gap:8px;border:1px solid var(--line);border-radius:999px;padding:6px 10px;background:#fff}
    .grid{overflow:auto;border:1px solid var(--line);border-radius:12px}
    table{border-collapse:separate;border-spacing:0;width:100%;min-width:800px}
    th,td{padding:8px 10px;border-bottom:1px solid var(--line);text-align:left}
    thead th{position:sticky;top:0;background:#fafafa;z-index:1}
    tr:nth-child(even){background:#fcfcfd}
    .tag{display:inline-block;font-size:12px;border-radius:999px;padding:2px 8px;border:1px solid var(--line);color:#111}
    .role-P{background:#e7f3ff}
    .role-D{background:#eaf7ef}
    .role-C{background:#fff6e6}
    .role-A{background:#fdeaf1}
    .warn{background:#fff8f8}
    .ok{color:var(--ok)}
    .bad{color:var(--warn)}
    .cols{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
    @media (max-width:900px){.cols{grid-template-columns:1fr}}
    .kpi{display:flex;gap:8px;flex-wrap:wrap}
    .kpi .pill{background:#f9fafb}
    .flex{display:flex;gap:12px;align-items:center;flex-wrap:wrap}
    .sep{height:1px;background:var(--line);margin:12px 0}
    .muted{color:var(--muted)}
    .right{margin-left:auto}
    .mini{font-size:12px}
    .sticky-cta{position:sticky;bottom:8px;display:flex;gap:8px;justify-content:flex-end;margin-top:8px}
    .tabs{display:flex;gap:8px;margin:6px 0 14px}
    .tab{padding:8px 12px;border:1px solid var(--line);border-radius:999px;background:#fff;cursor:pointer}
    .tab.active{background:var(--brand);color:#fff;border-color:var(--brand)}
    .comp-grid{display:grid;grid-template-columns:repeat(3,minmax(260px,1fr));gap:16px}
    @media (max-width:1100px){.comp-grid{grid-template-columns:repeat(2,minmax(260px,1fr))}}
    @media (max-width:740px){.comp-grid{grid-template-columns:1fr}}
    .comp-card{border:1px solid var(--line);border-radius:14px;background:#fff;padding:12px}
    .comp-card h3{margin:0 0 8px;color:var(--brand)}
    .comp-card table{min-width:unset}
    .totline{display:flex;justify-content:space-between;border-top:1px solid var(--line);padding-top:6px;margin-top:6px}
    .topbar{position:sticky;top:0;z-index:5;background:linear-gradient(180deg,#ffffff 0%,rgba(255,255,255,0.92) 100%);border:1px solid var(--line);border-radius:12px;padding:10px;margin:10px 0;box-shadow:0 4px 12px rgba(0,0,0,.04)}
  </style>
  <!-- ExcelJS fallback -->
  <script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
</head>
<body>
  <div class="wrap">
    <h1>Fantacalcio — Gestione Asta</h1>
    <div class="sub">Carica il <strong>listone Excel</strong>, definisci <strong>squadre</strong> e <strong>budget</strong>, assegna i giocatori inserendo <em>FantaSquadra</em> e <em>Costo</em>, quindi esporta un Excel finale. Limiti rosa: <strong>3P / 8D / 8C / 6A</strong>.</div>

    <div class="tabs">
      <button class="tab active" data-tab="asta">Asta & Roster</button>
      <button class="tab" data-tab="composizione">Composizione Squadre</button>
    </div>

    <section id="tab-asta">
      <div class="card topbar">
        <div class="kpi" id="summaryKPITop"></div>
      </div>

      <div class="row">
        <div class="card" style="flex:1 1 360px;min-width:320px">
          <h2>1) Carica listone</h2>
          <div class="cols">
            <div>
              <label for="file">File Excel (.xlsx)</label>
              <input type="file" id="file" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" />
            </div>
            <div>
              <label for="sheet">Foglio</label>
              <select id="sheet"></select>
            </div>
          </div>
          <div class="sep"></div>
          <div class="mini muted">Colonne richieste: <code>Nome</code>, <code>Ruolo</code> (accetta anche <code>R.</code>). Opzionali: <code>QUOT.</code> o <code>Quota</code>, <code>FantaSquadra</code>, <code>Costo</code>.</div>
          <div id="loadError" class="mini" style="color:#CE3262;display:none;margin-top:6px"></div>
          <div id="loadInfo" class="mini" style="color:#159947;display:none;margin-top:6px"></div>
          <label class="mini" style="display:block;margin-top:8px"><input type="checkbox" id="debugToggle"/> Mostra debug import</label>
          <pre id="debugBox" style="display:none;background:#0b1020;color:#d7e3ff;padding:8px;border-radius:8px;max-height:180px;overflow:auto"></pre>
        </div>

        <div class="card" style="flex:1 1 360px;min-width:320px">
          <h2>2) Squadre & Budget</h2>
          <div class="cols">
            <div>
              <label for="numTeams">Numero squadre</label>
              <input type="number" id="numTeams" min="2" max="20" value="9" />
            </div>
            <div>
              <label for="defaultBudget">Budget per squadra</label>
              <input type="number" id="defaultBudget" min="1" value="700" />
            </div>
          </div>
          <div style="margin-top:8px">
            <label for="teams">Elenco squadre (una per riga). Budget per riga con ";" es: <em>Atletico 'na vorta;700</em></label>
            <textarea id="teams" placeholder="Catania FC;700
Real SaraZozza;700
Hawk Tuah;700
Ciampino with love;700
Plusvalencia;700
Atletico 'na vorta;700
Abate Borisov;700
Team Alessio;700
Rockers Cave;700"></textarea>
          </div>
          <div class="flex" style="margin-top:8px">
            <div class="mini muted">Aggiorna e vedi i contatori in alto.</div>
            <div class="right"><button class="btn ghost" id="applyTeams">Applica squadre</button></div>
          </div>
        </div>
      </div>

      <div class="row">
        <div class="card" style="flex:1 1 100%">
          <h2>3) Roster (assegna FantaSquadra & Costo)</h2>
          <div class="flex" style="gap:8px;margin-bottom:10px">
            <input type="text" id="search" placeholder="Cerca giocatore…" style="max-width:280px" />
            <select id="roleFilter" style="max-width:160px"><option value="">Tutti i ruoli</option><option value="P">Portieri</option><option value="D">Difensori</option><option value="C">Centrocampisti</option><option value="A">Attaccanti</option></select>
            <label class="flex" style="gap:8px"><input type="checkbox" id="onlyAvailable" /> <span class="mini">Solo disponibili (non assegnati)</span></label>
            <div class="right"><button class="btn secondary" id="downloadXLSX" disabled>Esporta Excel</button></div>
          </div>

          <div class="grid">
            <table id="tbl">
              <thead>
                <tr>
                  <th>#</th>
                  <th>Nome</th>
                  <th>Ruolo</th>
                  <th>Quota</th>
                  <th>FantaSquadra</th>
                  <th>Costo</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>

          <div class="sep"></div>
          <div class="grid">
            <table id="board"><thead></thead><tbody></tbody></table>
          </div>
        </div>
      </div>
    </section>

    <section id="tab-composizione" style="display:none">
      <div class="card">
        <h2>Composizione Squadre (tabella)</h2>
        <div class="comp-grid" id="compGrid"></div>
      </div>
    </section>
  </div>

<script>
(function(){
  let workbook, data = [], teams = [], budgets = {}, limits = {P:3,D:8,C:8,A:6};

  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));
  const fileInput = $('#file');
  const sheetSel = $('#sheet');
  const numTeams = $('#numTeams');
  const defaultBudget = $('#defaultBudget');
  const teamsTA = $('#teams');
  const applyBtn = $('#applyTeams');
  const tblBody = $('#tbl tbody');
  const board = $('#board');
  const dlBtn = $('#downloadXLSX');
  const kpiTop = $('#summaryKPITop');
  const search = $('#search');
  const roleFilter = $('#roleFilter');
  const onlyAvail = $('#onlyAvailable');
  const compGrid = $('#compGrid');

  // Tabs
  $$('.tab').forEach(btn=>btn.addEventListener('click',()=>{
    $$('.tab').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    const tab = btn.dataset.tab;
    $('#tab-asta').style.display = (tab==='asta')?'block':'none';
    $('#tab-composizione').style.display = (tab==='composizione')?'block':'none';
    if(tab==='composizione') renderComposition();
  }));

  // DataList squadre per input
  const dataList = document.createElement('datalist');
  dataList.id = 'teams-list';
  document.body.appendChild(dataList);

  function parseTeamsText(){
    const lines = teamsTA.value.split(/
+/).map(s=>s.trim()).filter(Boolean);
    const n = parseInt(numTeams.value||'0',10);
    teams = lines.slice(0, n).map(l=>{
      const [name] = l.split(';');
      const nm = (name||'').trim();
      return nm?nm:null;
    }).filter(Boolean);

    budgets = {}; lines.slice(0,n).forEach(l=>{
      const [name,b] = l.split(';');
      const nm=(name||'').trim(); const bud=parseInt((b||'').trim(),10);
      if(nm){ budgets[nm] = Number.isFinite(bud)?bud:parseInt(defaultBudget.value,10); }
    });

    dataList.innerHTML=''; teams.forEach(t=>{const o=document.createElement('option');o.value=t;dataList.appendChild(o);});
  }

  applyBtn.addEventListener('click',()=>{ parseTeamsText(); renderAll(); });

  fileInput.addEventListener('change', async (e)=>{
    const f = e.target.files[0]; if(!f) return;
    const elErr = document.getElementById('loadError');
    const elInfo = document.getElementById('loadInfo');
    elErr.style.display='none'; elInfo.style.display='none';
    sheetSel.innerHTML = '<option>Caricamento…</option>';

    async function tryXLSX(){
      if(typeof XLSX==='undefined') throw new Error('Libreria XLSX non disponibile');
      const ab = await f.arrayBuffer();
      const wb = XLSX.read(ab, {type:'array'});
      return {names: wb.SheetNames, getRows: (nm)=> XLSX.utils.sheet_to_json(wb.Sheets[nm], {header:1, defval:''})};
    }

    async function tryExcelJS(){
      if(typeof ExcelJS==='undefined') throw new Error('ExcelJS non disponibile');
      const ab = await f.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(ab);
      const names = wb.worksheets.map(ws=> ws.name);
      const getRows = (nm)=>{
        const ws = wb.getWorksheet(nm); if(!ws) return [];
        const rows=[]; ws.eachRow({includeEmpty:true}, (row)=>{ rows.push(row.values.slice(1)); });
        return rows;
      };
      return {names, getRows};
    }

    try{
      let parser;
      try{ parser = await tryXLSX(); }
      catch(xerr){ console.warn('XLSX fallito, provo ExcelJS', xerr); parser = await tryExcelJS(); }

      sheetSel.innerHTML = '';
      parser.names.forEach(nm=>{ const opt=document.createElement('option'); opt.value=nm; opt.textContent=nm; sheetSel.appendChild(opt); });
      const prefer = parser.names.find(n=> n.toLowerCase().includes('lista') || n.toLowerCase().includes('calciatori'));
      sheetSel.value = prefer || parser.names[0];

      // memorizza parser per loadSheet
      window.__parser = parser;
      loadSheet();
    }catch(err){
      console.error(err);
      sheetSel.innerHTML = '';
      elErr.textContent = 'Errore caricamento: ' + err.message; elErr.style.display='block';
    }
  });
  sheetSel.addEventListener('change', ()=>{ try{ loadSheet(); document.getElementById('loadError').style.display='none'; } catch(err){ console.error(err); const el=document.getElementById('loadError'); el.textContent='Errore lettura foglio: '+err.message; el.style.display='block'; }} ); $('#loadError').style.display='none'; } catch(err){ console.error(err); const el=$('#loadError'); el.textContent='Errore lettura foglio: '+err.message; el.style.display='block'; }} );

  function normalizeHeader(s){
    const t = (s||'').toString().trim().toLowerCase();
    if(t==='nome') return 'Nome';
    if(t==='ruolo'||t==='r.'||t==='r') return 'Ruolo';
    if(t.startsWith('quot')) return 'Quota';
    if(t==='fantasquadra') return 'FantaSquadra';
    if(t==='costo') return 'Costo';
    return null;
  }

  function loadSheet(){
    try{
      const parser = window.__parser; if(!parser) throw new Error('Parser non pronto');
      const rowsAll = parser.getRows(sheetSel.value);
      if(!rowsAll||!rowsAll.length) throw new Error('Foglio vuoto o non leggibile');

      // Trova automaticamente la riga intestazioni (cerca Nome + Ruolo/R.) nelle prime 50 righe
      let headerIdx = -1;
      for(let i=0;i<Math.min(50, rowsAll.length); i++){
        const r = rowsAll[i].map(x=> (x||'').toString().trim().toLowerCase());
        if(r.some(x=>x==='nome') && r.some(x=> x==='ruolo' || x==='r.' || x==='r')){ headerIdx = i; break; }
      }
      if(headerIdx===-1) headerIdx = 0; // fallback

      const headers = rowsAll[headerIdx];
      const rows = rowsAll.slice(headerIdx+1);

      const map = {}; headers.forEach((h,i)=>{const t=(h||'').toString().trim().toLowerCase();
        if(t==='nome') map['Nome']=i;
        if(t==='ruolo'||t==='r.'||t==='r') map['Ruolo']=i;
        if(t==='quot.'||t==='quot'||t==='quota'||t.startsWith('quot')) map['Quota']=i;
        if(t==='fantasquadra'||t==='fanta squadra') map['FantaSquadra']=i;
        if(t==='costo'||t==='prezzo') map['Costo']=i;
      });
      if(map['Nome']===undefined || map['Ruolo']===undefined){
        throw new Error('Colonne obbligatorie mancanti. Servono almeno "Nome" e "Ruolo/R."');
      }

      data = rows.map((r,idx)=>({
        idx: idx+1,
        Nome: (r[map['Nome']]||'').toString().trim(),
        Ruolo: (r[map['Ruolo']]||'').toString().trim(),
        Quota: Number((r[map['Quota']]||'').toString().replace(',', '.')) || 0,
        FantaSquadra: (r[map['FantaSquadra']]||'').toString().trim(),
        Costo: (r[map['Costo']]!==undefined && r[map['Costo']]!=='' ? Number((r[map['Costo']]).toString().replace(',', '.'))||0 : '')
      })).filter(x=>x.Nome && ['P','D','C','A'].includes(x.Ruolo));

      data.sort((a,b)=> ({P:0,D:1,C:2,A:3}[a.Ruolo]-({P:0,D:1,C:2,A:3}[b.Ruolo]) || a.Nome.localeCompare(b.Nome));

      if(!teamsTA.value.trim()){
        teamsTA.value=[
          'Catania FC;700','Real SaraZozza;700','Hawk Tuah;700','Ciampino with love;700',
          'Plusvalencia;700','Atletico \'na vorta;700','Abate Borisov;700','Team Alessio;700','Rockers Cave;700'
        ].join('
');
      }
      parseTeamsText();
      renderAll();
      dlBtn.disabled=false;
      $('#loadError').style.display='none';

      // Debug info
      const info = `Foglio: ${sheetSel.value} — Righe totali: ${rowsAll.length}
Riga intestazioni rilevata: ${headerIdx+1}
Colonne mappate: ${Object.keys(map).join(', ')}
Giocatori letti: ${data.length}`;
      const roleSet = Array.from(new Set(rows.map(r=> (r[map['Ruolo']]||'').toString().trim()))).slice(0,10);
      document.getElementById('loadInfo').textContent = info + `
Valori Ruolo (esempio): ${roleSet.join(', ')}`;
      document.getElementById('loadInfo').style.display='block';
      const dbg = document.getElementById('debugBox');
      const dbgToggle = document.getElementById('debugToggle');
      dbg.textContent = JSON.stringify({headers: headers, sample: rows.slice(0,5)}, null, 2);
      dbg.style.display = dbgToggle.checked ? 'block' : 'none';
      dbgToggle.onchange = ()=>{ dbg.style.display = dbgToggle.checked ? 'block' : 'none'; };

    }catch(err){
      console.error(err);
      const el=$('#loadError'); el.textContent='Errore lettura dati: '+err.message; el.style.display='block';
      data=[]; renderAll();
    }
  }

  function computeTotals(){(){
    const perTeam = {}; teams.forEach(t=> perTeam[t] = {spesaTot:0, count:{P:0,D:0,C:0,A:0}, byRole:{P:[],D:[],C:[],A:[]}});
    data.forEach(r=>{
      const t=r.FantaSquadra; if(!t||!perTeam[t]) return;
      perTeam[t].spesaTot += (r.Costo===''?0:Number(r.Costo)||0);
      perTeam[t].count[r.Ruolo]=(perTeam[t].count[r.Ruolo]||0)+1;
      perTeam[t].byRole[r.Ruolo].push(r.Nome);
    });
    return {perTeam};
  }

  function renderKPIs(){
    const totals = computeTotals();
    kpiTop.innerHTML='';
    teams.forEach(team=>{
      const v = totals.perTeam[team]||{spesaTot:0,count:{}};
      const left = (budgets[team]||0) - (v.spesaTot||0);
      const pill = document.createElement('div');
      pill.className='pill';
      const c=v.count||{};
      pill.innerHTML=`<strong>${team}</strong> · credito: <span class="${left<0?'bad':'ok'}">${left}</span> · P:${c.P||0}/3 D:${c.D||0}/8 C:${c.C||0}/8 A:${c.A||0}/6`;
      kpiTop.appendChild(pill);
    });
  }

  [search, roleFilter, onlyAvail].forEach(el=> el.addEventListener('input', renderTable));

  function renderTable(){
    const q = (search.value||'').toLowerCase();
    const rf = roleFilter.value||'';
    const only = onlyAvail.checked;

    const rows = data.filter(r=>{
      if(q && !r.Nome.toLowerCase().includes(q)) return false;
      if(rf && r.Ruolo!==rf) return false;
      if(only && r.FantaSquadra) return false;
      return true;
    });

    const tbody = tblBody; tbody.innerHTML='';
    rows.forEach((row)=>{
      const tr=document.createElement('tr');
      tr.innerHTML=`
        <td>${row.idx}</td>
        <td>${row.Nome}</td>
        <td><span class="tag role-${row.Ruolo}">${row.Ruolo}</span></td>
        <td>${row.Quota||''}</td>
        <td><input list="teams-list" data-key="FantaSquadra" data-id="${row.idx}" value="${row.FantaSquadra||''}" placeholder="scegli squadra" /></td>
        <td><input type="number" min="0" step="1" data-key="Costo" data-id="${row.idx}" value="${row.Costo!==''?row.Costo:''}" placeholder="0" /></td>`;
      tbody.appendChild(tr);
    });

    $$("input[data-key]").forEach(inp=>{ inp.addEventListener('change', onEditCell); });
  }

  function onEditCell(e){
    const inp=e.target; const key=inp.dataset.key; const id=Number(inp.dataset.id);
    const row=data.find(x=>x.idx===id); if(!row) return;
    let val=inp.value; if(key==='Costo'){ val=(val==='')?'':Math.max(0,parseInt(val,10)||0); }
    row[key]=val;

    const totals = computeTotals();
    const t=row.FantaSquadra; const r=row.Ruolo;
    if(t && totals.perTeam[t]){
      const countR = totals.perTeam[t].count[r]||0; const lim=limits[r];
      if(countR>lim){
        if(key==='FantaSquadra') { row.FantaSquadra=''; inp.value=''; alert(`Limite superato per ${t} nel ruolo ${r} (max ${lim}).`); }
      }
    }

    renderAll();
  }

  function renderBoard(){
    const thead = board.querySelector('thead'); const tbody = board.querySelector('tbody');
    thead.innerHTML=''; tbody.innerHTML='';
    const htr=document.createElement('tr'); htr.innerHTML=['Ruolo',...teams].map(h=>`<th>${h}</th>`).join(''); thead.appendChild(htr);

    const totals=computeTotals(); const order=['P','D','C','A'];
    order.forEach(role=>{
      const maxRows=Math.max(...teams.map(t=> (totals.perTeam[t]?.byRole[role]?.length||0)),1);
      for(let i=0;i<maxRows;i++){
        const tr=document.createElement('tr');
        tr.innerHTML=`<td>${i===0?role:''}</td>`+teams.map(t=>{
          const arr=totals.perTeam[t]?.byRole[role]||[];return `<td>${arr[i]||''}</td>`;}).join('');
        tbody.appendChild(tr);
      }
      const trTot=document.createElement('tr');
      trTot.innerHTML=`<td><strong>Tot ${role}</strong></td>`+teams.map(t=>{
        const c=totals.perTeam[t]?.count[role]||0; const lim=limits[role];
        const cls=c>lim?'bad':(c===lim?'ok':''); return `<td class="${cls}"><strong>${c}/${lim}</strong></td>`;}).join('');
      tbody.appendChild(trTot);
      const trGap=document.createElement('tr'); trGap.innerHTML=`<td colspan="${teams.length+1}"><div class="sep"></div></td>`; tbody.appendChild(trGap);
    });

    const trB=document.createElement('tr');
    trB.innerHTML=`<td><strong>Spesa / Rimanente</strong></td>`+teams.map(t=>{
      const spent=computeTotals().perTeam[t]?.spesaTot||0; const left=(budgets[t]||0)-spent; const cls=left<0?'bad':'ok';
      return `<td><span>${spent}</span> / <span class="${cls}">${left}</span></td>`;}).join('');
    tbody.appendChild(trB);
  }

  function renderComposition(){
    const totals = computeTotals();
    compGrid.innerHTML='';
    teams.forEach(team=>{
      const c = totals.perTeam[team] || {count:{},byRole:{P:[],D:[],C:[],A:[]},spesaTot:0};
      const card=document.createElement('div'); card.className='comp-card';
      card.innerHTML=`<h3>${team}</h3>`;
      const tbl=document.createElement('table');
      const tb=document.createElement('tbody');
      const blocks=[["P",3,'Portieri'],["D",8,'Difensori'],["C",8,'Centrocampisti'],["A",6,'Attaccanti']];
      blocks.forEach(([role,lim,label])=>{
        const names=c.byRole[role]||[];
        const max=Math.max(lim,names.length);
        const head=document.createElement('tr'); head.innerHTML=`<th colspan="2">${label} (${(c.count[role]||0)}/${lim})</th>`; tb.appendChild(head);
        for(let i=0;i<max;i++){
          const tr=document.createElement('tr');
          tr.innerHTML=`<td style="width:24px">${i+1}</td><td>${names[i]||''}</td>`; tb.appendChild(tr);
        }
      });
      tbl.appendChild(tb); card.appendChild(tbl);
      const bud=budgets[team]||0; const spent=c.spesaTot||0; const left=bud-spent; const cls=left<0?'bad':'ok';
      const tot=document.createElement('div'); tot.className='totline'; tot.innerHTML=`<span>Totale: <strong>${spent}</strong></span><span>Resto: <strong class="${cls}">${left}</strong></span>`; card.appendChild(tot);
      compGrid.appendChild(card);
    });
  }

  function renderAll(){
    renderKPIs();
    renderTable();
    renderBoard();
    if($('#tab-composizione').style.display!=='none') renderComposition();
  }

  dlBtn.addEventListener('click', ()=>{
    const wb = XLSX.utils.book_new();
    const cfgRows = [["Squadra","Budget"], ...teams.map(t=>[t, budgets[t]||0]), [], ["Ruolo","Limite"],["P",3],["D",8],["C",8],["A",6]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cfgRows), 'CONFIG');

    const rosterRows = [["Nome","Ruolo","QuotaIniziale","FantaSquadra","Costo"], ...data.map(r=>[r.Nome,r.Ruolo,r.Quota,r.FantaSquadra||'', (r.Costo===''? '': Number(r.Costo)||0)])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rosterRows), 'ROSTER');

    const totals = computeTotals();
    const sumHeader=["Squadra","Budget","P","D","C","A","Totale speso","Resto"];
    const sumRows=[sumHeader];
    teams.forEach(t=>{ const c=totals.perTeam[t]?.count||{P:0,D:0,C:0,A:0}; const spent=totals.perTeam[t]?.spesaTot||0; const bud=budgets[t]||0; sumRows.push([t,bud,c.P||0,c.D||0,c.C||0,c.A||0,spent,bud-spent]);});
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sumRows), 'RIEPILOGO');

    const order=['P','D','C','A'];
    const bRows = [['Ruolo',...teams]];
    order.forEach(role=>{
      const maxRows=Math.max(...teams.map(t=>(totals.perTeam[t]?.byRole[role]?.length||0)),1);
      for(let i=0;i<maxRows;i++){ bRows.push([i===0?role:'', ...teams.map(t=> (totals.perTeam[t]?.byRole[role]?.[i]||''))]); }
      bRows.push([`Tot ${role}`, ...teams.map(t=> (totals.perTeam[t]?.count[role]||0) + '/' + limits[role])]);
      bRows.push([]);
    });
    bRows.push(['Spesa / Rimanente', ...teams.map(t=>{ const spent=totals.perTeam[t]?.spesaTot||0; const bud=budgets[t]||0; return `${spent} / ${bud-spent}`; })]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bRows), 'ASSEGNAZIONI');

    // COMPOSIZIONE
    const compRows = [];
    const chunks = (arr,size)=> arr.reduce((a,_,i)=> (i%size? a[a.length-1].push(arr[i]) : a.push([arr[i]]), a), []);
    chunks(teams,3).forEach(group=>{
      compRows.push(group);
      const blocks=[["P",3],["D",8],["C",8],["A",6]];
      blocks.forEach(([role,lim])=>{
        const rowsNeeded = Math.max(lim, ...group.map(t=> (totals.perTeam[t]?.byRole[role]?.length||0)));
        for(let i=0;i<rowsNeeded;i++) compRows.push(group.map(t=> totals.perTeam[t]?.byRole[role]?.[i]||''));
        compRows.push(group.map(t=> `${(totals.perTeam[t]?.count[role]||0)}/${lim}`));
      });
      compRows.push(group.map(t=>{const spent=totals.perTeam[t]?.spesaTot||0; const bud=budgets[t]||0; return `${spent} / ${bud-spent}`;}));
      compRows.push(['']);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(compRows), 'COMPOSIZIONE');

    XLSX.writeFile(wb, 'fantacalcio_app_export.xlsx');
  });
})();
</script>
</body>
</html>
