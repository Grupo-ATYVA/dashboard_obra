/**
 * IFC 3D VIEWER ENGINE v4 — Dashboard Obra · TEMAZEC / SPE T2 BUENO
 * Leitura em chunks de Blob — nunca carrega o arquivo inteiro na RAM
 * Parser assíncrono com yield · IFC2x3 + IFC4 · Three.js
 */
const IFC = (() => {
  let scene, camera, renderer, raycaster;
  let initialized = false;
  let modelGroup = null;
  let meshMap = {}, allElements = [];
  let selectedIds = new Set(), hiddenIds = new Set(), wireMode = false, shadowMode = true;
  let sidebarOpen = true, viewFilter = 'all', medStorage = {}, execFilter = '', periodMode = false;
  // Filtro "Executado em obra?" do layout BIM — independente do viewFilter
  // classico (Tudo/Medidos/Pendentes/Resumo), aplicado em conjunto com ele.
  // Mesma semantica do filtro "resumo" ja existente (pct>=100 ou em andamento
  // conta como executado).
  function execFilterOk(e){
    if(!execFilter) return true;
    const executed=e.pct>=100||e.emAndamento;
    return execFilter==='yes'?executed:!executed;
  }
  // ── COMPATIBILIZAÇÃO (multi-modelo) ─────────────────────────────────
  // Modelos adicionais (outras disciplinas) carregados por cima do modelo
  // principal só pra detecção de interferência e levantamento de materiais —
  // não entram no fluxo de medição/orçamento, que é exclusivo do principal.
  let compatModels = [], clashes = [], compatSeq = 0;
  const DISC_PALETTE = [0x4ade80,0x60a5fa,0xf5a623,0xc084fc,0xf472b6,0x2dd4bf,0xfacc15,0xfb7185];
  const STORAGE_KEY = 'ifc_med_v1';
  const STEEL_RATES_KEY = 'ifc_steel_rates_v1';
  // Taxas de consumo de aço (kg de aço / m³ de concreto) por tipo de elemento —
  // o IFC do TQS não traz armadura modelada nem quantitativo de aço, então o
  // peso é uma ESTIMATIVA por taxa (prática usual em orçamento preliminar),
  // não uma extração real. Valores default são referências de mercado para
  // edifícios residenciais/comerciais em concreto armado; editáveis na UI.
  const STEEL_RATES_DEFAULT = { IFCCOLUMN:100, IFCBEAM:120, IFCSLAB:75, IFCFOOTING:80, DEFAULT:90 };
  let STEEL_RATES = { ...STEEL_RATES_DEFAULT };
  function loadSteelRates(){
    try{ STEEL_RATES = { ...STEEL_RATES_DEFAULT, ...JSON.parse(localStorage.getItem(STEEL_RATES_KEY)||'{}') }; }
    catch(e){ STEEL_RATES = { ...STEEL_RATES_DEFAULT }; }
  }
  function saveSteelRates(rates){
    STEEL_RATES = { ...STEEL_RATES_DEFAULT, ...rates };
    try{ localStorage.setItem(STEEL_RATES_KEY, JSON.stringify(STEEL_RATES)); }catch(e){}
  }
  function steelRateFor(type){ return STEEL_RATES[type] ?? STEEL_RATES.DEFAULT; }
  function steelWeightOf(el){ return (el.volume||0) * steelRateFor(el.type); }
  loadSteelRates();

  // ── COLOURS ──────────────────────────────────────────────────────
  const COL = { default:0x7a6060,measured:0x4ade80,partial:0xf97316,selected:0xA50E24,andamento:0xf5a623,pending:0xc4c4c4,wall:0x8a7070,slab:0x9a8a7a,beam:0xb06050,column:0xd08060,window:0x6090c0,door:0xa06040,stair:0xc09070,roof:0x806070,mesAtual:0x22d3ee,mesAnterior:0x3b6fd6 };
  // "Mes atual" x "meses anteriores" — usado quando periodMode esta ligado,
  // pra visualizar o que foi medido dentro do mes corrente (ex.: pro RSO
  // semanal) versus o que ja tinha sido medido em meses passados.
  function isMesAtual(ts){
    if(!ts)return false;
    const d=new Date(ts),now=new Date();
    return d.getFullYear()===now.getFullYear()&&d.getMonth()===now.getMonth();
  }
  const typeColor = t => { t=(t||'').toLowerCase(); return t.includes('wall')?COL.wall:t.includes('slab')?COL.slab:t.includes('beam')?COL.beam:t.includes('column')?COL.column:t.includes('window')?COL.window:t.includes('door')?COL.door:t.includes('stair')?COL.stair:t.includes('roof')?COL.roof:COL.default; };
  // Categoria construtiva em portugues, usada no layout BIM (tabela de
  // quantitativos e grafico de volume por piso) — agrupa os mesmos tipos
  // IFC que ja orientam a cor (typeColor acima), mas com rotulo legivel.
  const categoriaConstrutiva = t => {
    t=(t||'').toLowerCase();
    if(t.includes('column')) return 'Pilares';
    if(t.includes('beam')) return 'Vigas';
    if(t.includes('slab')) return 'Lajes';
    if(t.includes('wall')||t.includes('curtainwall')) return 'Paredes';
    if(t.includes('footing')||t.includes('pile')) return 'Fundações';
    if(t.includes('roof')) return 'Coberturas';
    if(t.includes('stair')||t.includes('ramp')) return 'Escadas/Rampas';
    if(t.includes('door')) return 'Portas';
    if(t.includes('window')) return 'Janelas';
    return 'Outros';
  };

  const PRODUCT_TYPES = new Set(['IFCWALL','IFCWALLSTANDARDCASE','IFCSLAB','IFCBEAM','IFCCOLUMN','IFCDOOR','IFCWINDOW','IFCSTAIR','IFCSTAIRFLIGHT','IFCROOF','IFCRAMP','IFCPLATE','IFCMEMBER','IFCCURTAINWALL','IFCFURNISHINGELEMENT','IFCFLOWSEGMENT','IFCFLOWTERMINAL','IFCBUILDINGELEMENTPROXY','IFCPILE','IFCFOOTING','IFCCOVERING']);

  // ── PROGRESS UI ──────────────────────────────────────────────────
  function setStatus(msg, pct) {
    console.log('[IFC]', pct != null ? pct+'%' : '', msg);
    const s = document.getElementById('ifc-status'); if(s) s.textContent = msg;
    const wrap = document.getElementById('ifc-progress-bar-wrap');
    const fill = document.getElementById('ifc-progress-fill');
    const pctEl = document.getElementById('ifc-load-pct');
    const lbl = document.getElementById('ifc-load-label');
    const det = document.getElementById('ifc-load-detail');
    if(pct != null) {
      if(wrap) wrap.style.display = 'block';
      const p = Math.min(Math.round(pct), 100);
      if(fill) fill.style.width = p + '%';
      if(pctEl) pctEl.textContent = p + '%';
      if(lbl) { lbl.textContent = msg.length > 55 ? msg.slice(0,55)+'…' : msg; lbl.style.color = msg.startsWith('❌') ? '#ff4d6d' : 'var(--gold)'; }
      if(det) det.textContent = msg;
      if(p >= 100) setTimeout(() => { if(wrap) wrap.style.display='none'; }, 2800);
    } else {
      if(det) det.textContent = msg;
    }
  }

  // ── CHUNK-BASED FILE READER ──────────────────────────────────────
  // Reads file in 8MB binary chunks, decodes UTF-8, extracts IFC lines
  const READ_CHUNK = 8 * 1024 * 1024; // 8 MB per read

  function readFileInChunks(file, onDone, onError) {
    setStatus('Iniciando leitura: ' + file.name + ' (' + (file.size/1048576).toFixed(0) + ' MB)…', 1);

    const decoder = new TextDecoder('utf-8');
    const entities = new Map();
    let offset = 0;
    let leftover = '';       // incomplete line from previous chunk
    let inData = false;      // are we inside DATA; … ENDSEC; ?
    let depth = 0, inStr = false;
    let stmtBuf = '';
    let totalEntities = 0;

    function readNext() {
      if(offset >= file.size) {
        setStatus('Leitura completa: ' + totalEntities.toLocaleString() + ' entidades', 62);
        setTimeout(() => buildFromEntities(entities, onDone, onError), 20);
        return;
      }
      const end = Math.min(offset + READ_CHUNK, file.size);
      const blob = file.slice(offset, end);
      const pct = Math.round(1 + (offset / file.size) * 58);
      setStatus('Lendo e parseando: ' + (offset/1048576).toFixed(0) + ' / ' + (file.size/1048576).toFixed(0) + ' MB — ' + totalEntities.toLocaleString() + ' entidades…', pct);
      offset = end;

      const fr = new FileReader();
      fr.onload = ev => {
        const chunk = leftover + decoder.decode(ev.target.result, { stream: true });
        // Find last newline to avoid splitting a statement
        const lastNL = chunk.lastIndexOf('\n');
        const toProcess = lastNL >= 0 ? chunk.slice(0, lastNL + 1) : chunk;
        leftover = lastNL >= 0 ? chunk.slice(lastNL + 1) : '';

        // Find DATA section
        if(!inData) {
          const di = toProcess.indexOf('DATA;');
          if(di === -1) { setTimeout(readNext, 0); return; }
          inData = true;
          processText(toProcess.slice(di + 5), entities);
        } else {
          // Check for ENDSEC
          const ei = toProcess.indexOf('ENDSEC;');
          if(ei !== -1) {
            processText(toProcess.slice(0, ei), entities);
            totalEntities = entities.size;
            offset = file.size; // stop reading
          } else {
            processText(toProcess, entities);
          }
        }
        totalEntities = entities.size;
        setTimeout(readNext, 0);
      };
      fr.onerror = () => onError(new Error('Falha ao ler chunk do arquivo.'));
      fr.readAsArrayBuffer(blob);
    }

    // State machine runs over each text chunk
    function processText(text, ents) {
      for(let i = 0; i < text.length; i++) {
        const c = text[i];
        if(inStr) {
          stmtBuf += c;
          if(c === "'") { if(text[i+1] === "'") { stmtBuf += "'"; i++; } else inStr = false; }
          continue;
        }
        if(c === "'") { inStr = true; stmtBuf += c; continue; }
        if(c === '(' || c === '[') { depth++; stmtBuf += c; continue; }
        if(c === ')' || c === ']') { depth--; stmtBuf += c; continue; }
        if(c === ';' && depth === 0) {
          const s = stmtBuf.replace(/\r?\n\s*/g,' ').trim();
          if(s.startsWith('#')) processStmt(s, ents);
          stmtBuf = '';
          continue;
        }
        if(depth === 0 && stmtBuf.length === 0 && (c === '\n'||c === '\r'||c === ' '||c === '\t')) continue;
        stmtBuf += c;
      }
    }

    setTimeout(readNext, 50);
  }

  function processStmt(s, ents) {
    const eq = s.indexOf('='); if(eq === -1) return;
    const id = s.slice(1, eq).trim();
    const rest = s.slice(eq+1).trim();
    const po = rest.indexOf('('); if(po === -1) return;
    const type = rest.slice(0, po).trim().toUpperCase();
    const rawArgs = rest.slice(po+1, rest.lastIndexOf(')'));
    if(id && type) ents.set(id, {type, rawArgs});
  }

  // ── UNIDADES ──────────────────────────────────────────────────────
  const SI_PREFIX_SCALE = {EXA:1e18,PETA:1e15,TERA:1e12,GIGA:1e9,MEGA:1e6,KILO:1e3,HECTO:1e2,DECA:10,DECI:.1,CENTI:.01,MILLI:.001,MICRO:1e-6,NANO:1e-9,PICO:1e-12,FEMTO:1e-15,ATTO:1e-18};
  function detectUnitScale(ents) {
    // Arquivos IFC podem declarar comprimento em cm/mm/etc (ex.: obras de
    // estrutura metalica). O visualizador assume metros em toda a UI (grade,
    // camera, tamanhos de fallback) — sem essa normalizacao, um modelo em cm
    // fica 100x maior que o esperado e desaparece nos limites de precisao da
    // camera quando enquadrado no zoom padrao.
    for(const [,e] of ents) {
      if(e.type!=='IFCSIUNIT') continue;
      const a = sa(e.rawArgs);
      const unitType = (a[1]||'').replace(/\./g,'');
      if(unitType!=='LENGTHUNIT') continue;
      const prefix = (a[2]||'').replace(/\./g,'').replace(/^\$$/,'');
      return SI_PREFIX_SCALE[prefix]||1;
    }
    return 1;
  }

  // ── BUILD FROM ENTITIES ──────────────────────────────────────────
  function buildFromEntities(ents, onDone, onError) {
    try {
      setStatus('Mapeando ' + ents.size.toLocaleString() + ' entidades: pavimentos e contenção…', 64);

      const unitScale = detectUnitScale(ents);
      const storeys = {};
      const containment = {};
      const productIds = [];

      for(const [id, e] of ents) {
        if(e.type === 'IFCBUILDINGSTOREY') {
          const a = sa(e.rawArgs); storeys[id] = sv(a[2])||sv(a[1])||('Pav.'+id);
        }
        if(e.type === 'IFCRELCONTAINEDINSPATIALSTRUCTURE') {
          const a = sa(e.rawArgs); const nm = storeys[ri(a[5])]||'Sem pavimento';
          const re = /#(\d+)/g; let m; while((m=re.exec(a[4]||''))!==null) containment[m[1]]=nm;
        }
        if(PRODUCT_TYPES.has(e.type)) productIds.push(id);
      }

      // Second pass: fix containment (storeys may not have been parsed yet in first pass)
      for(const [,e] of ents) {
        if(e.type === 'IFCRELCONTAINEDINSPATIALSTRUCTURE') {
          const a = sa(e.rawArgs);
          const strId = ri(a[5]);
          if(strId && storeys[strId]) {
            const nm = storeys[strId];
            const re = /#(\d+)/g; let m;
            while((m=re.exec(a[4]||''))!==null) containment[m[1]]=nm;
          }
        }
      }

      setStatus(productIds.length.toLocaleString() + ' elementos — construindo geometria…', 66);

      if(productIds.length === 0) {
        onError(new Error('0 produtos encontrados. Entidades: ' + ents.size + '. O arquivo pode ser MEP/instalações — verifique se é IFC de arquitetura ou estrutura.'));
        return;
      }

      const elements = [];
      const storeyList = Object.values(storeys);
      let idx = 0;

      function buildChunk() {
        const end = Math.min(idx + 150, productIds.length);
        setStatus('Geometria: ' + idx + ' / ' + productIds.length + ' elementos…', Math.round(66+(idx/productIds.length)*26));

        for(let i = idx; i < end; i++) {
          const id = productIds[i], e = ents.get(id);
          const a = sa(e.rawArgs);
          const gid = sv(a[0]), nm = sv(a[2])||sv(a[1])||e.type;
          const lvl = containment[id]||'Sem pavimento';
          let geom = null;

          try {
            const repId = ri(a[6]);
            if(repId && ents.has(repId)) geom = extractGeom(repId, ents);
            if(geom && ri(a[5])) { geom.applyMatrix4(worldMat(ri(a[5]), ents)); }
          } catch(ex) { geom = null; }

          if(!geom) {
            // Sem geometria extraivel (ex.: BREP nao suportado): usa uma caixa
            // generica, mas posicionada com o placement REAL do elemento —
            // sem isso ela cai numa grade arbitraria pequena, longe do
            // predio real, e some da vista (era o caso das rampas).
            const t = e.type.toLowerCase();
            const invScale = 1/(unitScale||1);
            const sz = t.includes('slab')?[5,.25,5]:t.includes('wall')?[3,2.8,.2]:t.includes('column')?[.4,3,.4]:t.includes('beam')?[3,.4,.3]:t.includes('stair')?[2.5,1.5,3]:t.includes('ramp')?[3,.2,4]:t.includes('door')?[.9,2.1,.1]:t.includes('window')?[1.2,1,.1]:[1,1,1];
            geom = new THREE.BoxGeometry(sz[0]*invScale,sz[1]*invScale,sz[2]*invScale);
            const placementId = ri(a[5]);
            if(placementId && ents.has(placementId)) {
              try { geom.applyMatrix4(worldMat(placementId, ents)); } catch(ex) {}
            } else {
              const gi = elements.length, gx=((gi%30)*6.5-97)*invScale, gz=(Math.floor(gi/30)*6.5-97)*invScale, gy=Math.max(storeyList.indexOf(lvl),0)*4*invScale;
              geom.translate(gx,gy,gz);
            }
          }
          elements.push({id,gid,type:e.type,name:nm,objectType:sv(a[4])||'',level:lvl,geom});
        }
        idx = end;
        idx < productIds.length ? setTimeout(buildChunk, 0) : setTimeout(()=>{ try{onDone({elements,storeys,unitScale});}catch(ex){onError(ex);} }, 0);
      }
      setTimeout(buildChunk, 0);
    } catch(ex) { onError(ex); }
  }

  // ── GEOMETRY ──────────────────────────────────────────────────────
  function extractGeom(repId, ents) {
    if(!ents.has(repId)) return null;
    const rv = sa(ents.get(repId).rawArgs), repsRaw = rv[2]||rv[1]||'';
    const repRefs = []; const re2=/#(\d+)/g; let mr;
    while((mr=re2.exec(repsRaw))!==null) repRefs.push(mr[1]);
    for(const rr of repRefs) {
      if(!ents.has(rr)) continue;
      const rv2=sa(ents.get(rr).rawArgs), id2=sv(rv2[1]);
      if(id2&&!['Body','Facetation','Brep','SolidModel','SweptSolid','Clipping','Mapped','MappedRepresentation',''].some(k=>id2.includes(k))) continue;
      const itemsRaw=rv2[3]||rv2[2]||''; const refs=[]; const re3=/#(\d+)/g; let m2;
      while((m2=re3.exec(itemsRaw))!==null) refs.push(m2[1]);
      for(const ir of refs) { const g=geomItem(ir,ents); if(g) return g; }
    }
    return null;
  }

  function geomItem(id, ents) {
    if(!id||!ents.has(id)) return null;
    const e=ents.get(id);
    if(e.type==='IFCEXTRUDEDAREASOLID') return extrudeSolid(id,ents);
    if(e.type==='IFCFACETEDBREP') return facetedBrep(id,ents);
    if(e.type==='IFCBOOLEANCLIPPINGRESULT'||e.type==='IFCBOOLEANRESULT') { const a=sa(e.rawArgs); return geomItem(ri(a[1]),ents); }
    if(e.type==='IFCMAPPEDITEM') {
      const a=sa(e.rawArgs), srcId=ri(a[0]);
      if(srcId&&ents.has(srcId)) {
        const src=ents.get(srcId), sa2=sa(src.rawArgs), repId2=ri(sa2[1]);
        if(repId2&&ents.has(repId2)) {
          const ra=sa(ents.get(repId2).rawArgs), iRaw=ra[3]||ra[2]||''; const re3=/#(\d+)/g; let m;
          while((m=re3.exec(iRaw))!==null){ const g=geomItem(m[1],ents); if(g)return g; }
        }
      }
    }
    return null;
  }

  function extrudeSolid(id,ents) {
    const e=ents.get(id); if(!e||e.type!=='IFCEXTRUDEDAREASOLID') return null;
    const a=sa(e.rawArgs), profId=ri(a[0]), posId=ri(a[1]), depth=nv(a[3]);
    if(!profId||depth<=0||depth>2000) return null;
    const prof=ents.get(profId); if(!prof) return null;
    let shape=null;
    try {
      let sectionDims=null;
      if(prof.type==='IFCRECTANGLEPROFILEDEF'){const pa=sa(prof.rawArgs);const w=nv(pa[3]),h=nv(pa[4]);if(w<=0||h<=0||w>1000||h>1000)return null;shape=new THREE.Shape();shape.moveTo(-w/2,-h/2);shape.lineTo(w/2,-h/2);shape.lineTo(w/2,h/2);shape.lineTo(-w/2,h/2);shape.closePath();sectionDims={w,h};}
      else if(prof.type==='IFCCIRCLEPROFILEDEF'){const pa=sa(prof.rawArgs);const r=nv(pa[3]);if(r<=0||r>1000)return null;shape=new THREE.Shape();shape.absarc(0,0,r,0,Math.PI*2,false);sectionDims={r};}
      else if(prof.type==='IFCISHAPEPROFILEDEF'){const pa=sa(prof.rawArgs);const fw=nv(pa[3]),fd=nv(pa[4]),tw=nv(pa[5]),tf=nv(pa[6]);if(fw<=0||fd<=0)return null;shape=new THREE.Shape();shape.moveTo(-fw/2,-fd/2);shape.lineTo(fw/2,-fd/2);shape.lineTo(fw/2,-fd/2+tf);shape.lineTo(tw/2,-fd/2+tf);shape.lineTo(tw/2,fd/2-tf);shape.lineTo(fw/2,fd/2-tf);shape.lineTo(fw/2,fd/2);shape.lineTo(-fw/2,fd/2);shape.lineTo(-fw/2,fd/2-tf);shape.lineTo(-tw/2,fd/2-tf);shape.lineTo(-tw/2,-fd/2+tf);shape.lineTo(-fw/2,-fd/2+tf);shape.closePath();}
      else if(prof.type==='IFCARBITRARYCLOSEDPROFILEDEF'){const pa=sa(prof.rawArgs);shape=polyShape(ri(pa[2]),ents);}
      else{shape=new THREE.Shape();shape.moveTo(-.5,-.5);shape.lineTo(.5,-.5);shape.lineTo(.5,.5);shape.lineTo(-.5,.5);shape.closePath();}
      if(!shape) return null;
      const geom=new THREE.ExtrudeGeometry(shape,{depth,bevelEnabled:false});
      if(posId&&ents.has(posId)){const pe=ents.get(posId);if(pe.type==='IFCAXIS2PLACEMENT3D')geom.applyMatrix4(pl2mat(axPl3D(posId,ents)));}
      // Perimetro da secao transversal + profundidade de extrusao (comprimento
      // do elemento) — guardados aqui para calcular a area lateral (forma) em
      // buildScene, sem precisar reclassificar triangulos da malha 3D final.
      geom.userData.sectionPerimeter=shapePerimeter(shape);
      geom.userData.extrudeDepth=depth;
      geom.userData.sectionDims=sectionDims;
      return geom;
    }catch(ex){return null;}
  }
  function shapePerimeter(shape){
    try{
      const pts=shape.getPoints(24);
      let per=0;
      for(let i=0;i<pts.length;i++){const p1=pts[i],p2=pts[(i+1)%pts.length];per+=p1.distanceTo(p2);}
      return per;
    }catch(ex){return 0;}
  }

  function polyShape(id,ents){
    if(!id||!ents.has(id))return null;
    const a=sa(ents.get(id).rawArgs),ptRaw=a[0]||'',ptIds=[],re=/#(\d+)/g;let m;
    while((m=re.exec(ptRaw))!==null)ptIds.push(m[1]);
    if(ptIds.length<3)return null;
    const pts=ptIds.map(pid=>{const pa=sa((ents.get(pid)||{rawArgs:''}).rawArgs),inn=(pa[0]||'').replace(/[()]/g,'').split(',');return[parseFloat(inn[0])||0,parseFloat(inn[1])||0];});
    const s=new THREE.Shape();s.moveTo(pts[0][0],pts[0][1]);for(let i=1;i<pts.length;i++)s.lineTo(pts[i][0],pts[i][1]);s.closePath();return s;
  }

  function facetedBrep(id,ents){
    // IFCFACETEDBREP -> IFCCLOSEDSHELL -> IFCFACE[] -> IFCFACEOUTERBOUND ->
    // IFCPOLYLOOP -> pontos. Usado por rampas e outras formas nao extrudadas
    // no arquivo. Triangula cada face em leque (assume face convexa/planar,
    // suficiente para as faces retas tipicas de um BREP arquitetonico).
    try{
      const a=sa(ents.get(id).rawArgs), shellId=ri(a[0]);
      if(!shellId||!ents.has(shellId))return null;
      const facesRaw=(sa(ents.get(shellId).rawArgs)[0]||''), faceIds=[]; let m;
      const reF=/#(\d+)/g; while((m=reF.exec(facesRaw))!==null)faceIds.push(m[1]);
      if(!faceIds.length)return null;
      const positions=[];
      for(const fid of faceIds){
        const face=ents.get(fid); if(!face)continue;
        const boundsRaw=sa(face.rawArgs)[0]||'', boundIds=[]; const reB=/#(\d+)/g; let mb;
        while((mb=reB.exec(boundsRaw))!==null)boundIds.push(mb[1]);
        if(!boundIds.length)continue;
        const boundE=ents.get(boundIds[0]); if(!boundE)continue;
        const loopId=ri(sa(boundE.rawArgs)[0]); if(!loopId||!ents.has(loopId))continue;
        const loop=ents.get(loopId); if(loop.type!=='IFCPOLYLOOP')continue;
        const ptsRaw=sa(loop.rawArgs)[0]||'', ptIds=[]; const reP=/#(\d+)/g; let mp;
        while((mp=reP.exec(ptsRaw))!==null)ptIds.push(mp[1]);
        if(ptIds.length<3)continue;
        const pts=ptIds.map(pid=>cp(pid,ents));
        for(let i=1;i<pts.length-1;i++)positions.push(...pts[0],...pts[i],...pts[i+1]);
      }
      if(!positions.length)return null;
      const geom=new THREE.BufferGeometry();
      geom.setAttribute('position',new THREE.Float32BufferAttribute(positions,3));
      geom.computeVertexNormals();
      return geom;
    }catch(ex){return null;}
  }

  // ── PLACEMENT HELPERS ─────────────────────────────────────────────
  const sa = raw => { const res=[],len=(raw||'').length;let d=0,cur='',inS=false;for(let i=0;i<len;i++){const c=raw[i];if(inS){cur+=c;if(c==="'"&&raw[i+1]!=="'")inS=false;else if(c==="'"&&raw[i+1]==="'"){cur+="'";i++;}continue;}if(c==="'"){inS=true;cur+=c;continue;}if(c==='('||c==='['){d++;cur+=c;}else if(c===')'||c===']'){d--;cur+=c;}else if(c===','&&d===0){res.push(cur.trim());cur='';}else cur+=c;}if(cur.trim())res.push(cur.trim());return res; };
  const ri = v => (v||'').trim().startsWith('#')?(v||'').trim().slice(1):null;
  // Textos IFC com acentos (comum em exports brasileiros/TQS) vem codificados
  // no formato STEP/EXPRESS (ISO 10303-21): \X\HH (1 byte Latin-1) ou
  // \X2\....\X0\ (UCS-2, 4 hex digitos por caractere) — sem decodificar isso
  // "Fundações" chega como "Funda\X\E7\X\F5es" em qualquer campo de texto
  // (Name/Description/ObjectType/Tag).
  function decodeStepStr(s){
    if(!s||s.indexOf('\\X')===-1)return s;
    s=s.replace(/\\X2\\([0-9A-Fa-f]+)\\X0\\/g,(m,hex)=>{
      let out='';
      for(let i=0;i<hex.length;i+=4)out+=String.fromCharCode(parseInt(hex.substr(i,4),16));
      return out;
    });
    s=s.replace(/\\X\\([0-9A-Fa-f]{2})/g,(m,hex)=>String.fromCharCode(parseInt(hex,16)));
    return s;
  }
  const sv = v => decodeStepStr((v||'').replace(/^'|'$/g,'').trim());
  const nv = v => parseFloat(v)||0;

  function cp(id,ents){const e=ents.get(id);if(!e)return[0,0,0];const a=sa(e.rawArgs),inn=(a[0]||'').replace(/[()]/g,'').split(',');return[parseFloat(inn[0])||0,parseFloat(inn[1])||0,parseFloat(inn[2])||0];}
  function dir(id,ents){const e=ents.get(id);if(!e)return[0,0,1];const a=sa(e.rawArgs),inn=(a[0]||'').replace(/[()]/g,'').split(',');return[parseFloat(inn[0])||0,parseFloat(inn[1])||0,parseFloat(inn[2])||0];}
  function axPl3D(id,ents){const e=ents.get(id);if(!e)return{pos:[0,0,0],ax:[0,0,1],ref:[1,0,0]};const a=sa(e.rawArgs);return{pos:ri(a[0])?cp(ri(a[0]),ents):[0,0,0],ax:ri(a[1])?dir(ri(a[1]),ents):[0,0,1],ref:ri(a[2])?dir(ri(a[2]),ents):[1,0,0]};}
  function pl2mat(pl){const[ox,oy,oz]=pl.pos,zv=new THREE.Vector3(...pl.ax).normalize(),xv=new THREE.Vector3(...pl.ref).normalize(),yv=new THREE.Vector3().crossVectors(zv,xv).normalize(),m=new THREE.Matrix4();m.set(xv.x,yv.x,zv.x,ox,xv.y,yv.y,zv.y,oy,xv.z,yv.z,zv.z,oz,0,0,0,1);return m;}
  function worldMat(placId,ents){const stack=[],seen=new Set();let cur=placId;while(cur&&ents.has(cur)&&!seen.has(cur)){seen.add(cur);const e=ents.get(cur);if(e.type==='IFCLOCALPLACEMENT'){const a=sa(e.rawArgs);stack.unshift({r:ri(a[0]),l:ri(a[1])});cur=ri(a[0]);}else break;}let mat=new THREE.Matrix4();for(const s of stack){if(!s.l||!ents.has(s.l))continue;const le=ents.get(s.l);if(le.type==='IFCAXIS2PLACEMENT3D')mat.multiply(pl2mat(axPl3D(s.l,ents)));else if(le.type==='IFCAXIS2PLACEMENT2D'){const a=sa(le.rawArgs);mat.multiply(pl2mat({pos:ri(a[0])?cp(ri(a[0]),ents):[0,0,0],ax:[0,0,1],ref:[1,0,0]}));}}return mat;}

  // ── SCENE ──────────────────────────────────────────────────────────
  function meshVolume(geometry){
    // Soma dos volumes assinados dos tetraedros (origem-triangulo) — funciona
    // para qualquer malha fechada e manifold, como as geradas por
    // ExtrudeGeometry/BoxGeometry/BREP usadas aqui.
    const pos=geometry.attributes.position; if(!pos) return 0;
    const idx=geometry.index;
    const a=new THREE.Vector3(),b=new THREE.Vector3(),c=new THREE.Vector3();
    let vol=0;
    const tri=(i0,i1,i2)=>{a.fromBufferAttribute(pos,i0);b.fromBufferAttribute(pos,i1);c.fromBufferAttribute(pos,i2);vol+=a.dot(b.clone().cross(c))/6;};
    if(idx){for(let i=0;i<idx.count;i+=3)tri(idx.getX(i),idx.getX(i+1),idx.getX(i+2));}
    else{for(let i=0;i<pos.count;i+=3)tri(i,i+1,i+2);}
    return Math.abs(vol);
  }
  function buildScene(data){
    if(modelGroup){scene.remove(modelGroup);modelGroup.traverse(o=>{if(o.geometry)o.geometry.dispose();if(o.material)o.material.dispose();});modelGroup=null;}
    meshMap={};allElements=[];selectedIds=new Set();hiddenIds=new Set();modelGroup=new THREE.Group();modelGroup.scale.setScalar(data.unitScale||1);
    // IFC usa Z como eixo vertical (norma do formato); esta cena usa Y como
    // "para cima" (convencao Three.js/grid/camera). Sem essa rotacao o
    // modelo aparece deitado de lado, e nenhum angulo de orbita corrige isso.
    modelGroup.rotation.x=-Math.PI/2;
    scene.add(modelGroup);
    const unitScale3=Math.pow(data.unitScale||1,3);
    const unitScale2=Math.pow(data.unitScale||1,2);
    data.elements.forEach(el=>{
      const mat=new THREE.MeshLambertMaterial({color:typeColor(el.type),transparent:true,opacity:.88,side:THREE.DoubleSide});
      // Raycast e box-select dependem do boundingSphere estar correto; geometria
      // degenerada (extrusao com perfil vazio, etc.) pode gerar um sphere NaN
      // ou de raio 0, o que faz o elemento nunca ser clicavel — garante um
      // sphere minimo valido nesses casos.
      el.geom.computeBoundingSphere();
      const bs=el.geom.boundingSphere;
      if(!bs||!isFinite(bs.radius)||bs.radius<0.01){
        el.geom.computeBoundingBox();
        const bb=el.geom.boundingBox;
        if(bb&&isFinite(bb.min.x)){const c=new THREE.Vector3();bb.getCenter(c);const r=Math.max(bb.getSize(new THREE.Vector3()).length()/2,0.05);el.geom.boundingSphere=new THREE.Sphere(c,r);}
      }
      const mesh=new THREE.Mesh(el.geom,mat);
      el.volume=meshVolume(el.geom)*unitScale3;
      // Area de forma (lateral) = perimetro da secao x comprimento extrudado —
      // area total das faces laterais do elemento (para pilares, e a area de
      // forma completa; para vigas, inclui a face de topo que normalmente nao
      // e formada, entao e um teto/aproximacao um pouco por cima do real).
      const gu=el.geom.userData||{};
      el.formworkArea=(gu.sectionPerimeter||0)*(gu.extrudeDepth||0)*unitScale2;
      el.steelWeight=steelWeightOf(el);
      const scl=data.unitScale||1;
      el.sectionDims=gu.sectionDims?('w' in gu.sectionDims?{w:gu.sectionDims.w*scl,h:gu.sectionDims.h*scl}:{r:gu.sectionDims.r*scl}):null;
      mesh.userData={gid:el.gid,type:el.type,name:el.name,level:el.level,volume:el.volume,formworkArea:el.formworkArea,steelWeight:el.steelWeight};
      mesh.castShadow=mesh.receiveShadow=shadowMode;
      modelGroup.add(mesh);delete el.geom;
      meshMap[el.gid]={mesh,data:el,pct:medStorage[el.gid]?.pct??-1,orcRef:medStorage[el.gid]?.orcRef??'',emAndamento:!!medStorage[el.gid]?.emAndamento,ts:medStorage[el.gid]?.ts||0};
      allElements.push(el);
    });
    IFC._meshMap=meshMap;window.IFCViz?.onModelLoaded?.(meshMap);window.IFCBim?.onModelLoaded?.();
    setStatus('✅ Carregado: '+data.elements.length+' elementos',100);
    refreshColors();fitCamera();buildTree();buildOrcSel();updateMedCount();updateHiddenCount();
    setTimeout(()=>document.getElementById('ifc-drop-zone')?.classList.add('hidden'),400);
  }

  // ── COLORS / TREE ─────────────────────────────────────────────────
  function refreshColors(){for(const gid in meshMap){const e=meshMap[gid];if(!e.mesh)continue;if(selectedIds.has(gid)){e.mesh.material.color.setHex(COL.selected);e.mesh.material.opacity=1;e.mesh.material.emissive=new THREE.Color(0x3a0808);}else{
    // "Em andamento" e um status manual, independente do % medido (nao entra
    // no calculo de avanco fisico) — mas tem prioridade visual pra ficar
    // facil de achar no 3D.
    const pendente=!e.emAndamento&&e.pct<1;
    const corMedido=periodMode?(isMesAtual(e.ts)?COL.mesAtual:COL.mesAnterior):(e.pct>=100?COL.measured:COL.partial);
    e.mesh.material.color.setHex(e.emAndamento?COL.andamento:e.pct>=1?corMedido:COL.pending);
    // Pendente (0% medido, sem marcacao) fica em cinza claro semitransparente
    // — um "fantasma" do que falta, sem competir visualmente com o que ja
    // esta medido/em andamento (que ficam solidos e chamam mais atencao).
    e.mesh.material.opacity=pendente?.55:.88;
    e.mesh.material.emissive=new THREE.Color(0);}const filterOk=viewFilter==='all'?true:viewFilter==='measured'?e.pct>=1:viewFilter==='resumo'?(e.pct>=100||e.emAndamento):e.pct<1;e.mesh.visible=filterOk&&execFilterOk(e)&&!hiddenIds.has(gid);}window.IFCBim?.onStateChanged?.();}
  function buildTree(){renderTree();const c=document.getElementById('ifc-el-count');if(c)c.textContent=allElements.length+' el.';}
  function renderTree(){
    const tree=document.getElementById('ifc-tree');if(!tree)return;
    const q=(document.getElementById('ifc-search')?.value||'').toLowerCase();
    const filtered=allElements.filter(el=>(viewFilter==='all'||(viewFilter==='measured'?(meshMap[el.gid]?.pct||0)>=1:(meshMap[el.gid]?.pct||0)<1))&&execFilterOk(meshMap[el.gid]||{})&&(!q||el.name.toLowerCase().includes(q)||el.type.toLowerCase().includes(q)||el.level.toLowerCase().includes(q)));
    const byLevel={};filtered.forEach(el=>{const l=el.level||'Sem pavimento';(byLevel[l]||(byLevel[l]=[])).push(el);});
    let html='';
    for(const lvl of Object.keys(byLevel).sort()){
      html+=`<div style="padding:6px 10px 4px;font-size:10px;letter-spacing:1.2px;text-transform:uppercase;color:#6b5555;background:rgba(0,0,0,.2);border-bottom:1px solid #3a2424;position:sticky;top:0">${lvl}</div>`;
      for(const el of byLevel[lvl]){const e=meshMap[el.gid],pct=e?.pct??-1,isHidden=hiddenIds.has(el.gid),cls=(selectedIds.has(el.gid)?'selected':e?.emAndamento?'andamento':pct>=100?'measured':pct>=1?'partial':'')+(isHidden?' ifc-hidden-item':'');const ico=el.type.includes('WALL')?'🧱':el.type.includes('SLAB')?'⬛':el.type.includes('BEAM')?'═':el.type.includes('COLUMN')?'▐':el.type.includes('DOOR')?'🚪':el.type.includes('WINDOW')?'🪟':el.type.includes('STAIR')?'🪜':'📦';const orcBadge=e?.orcRef?`<span class="ifc-orc-badge" title="Vinculado ao orçamento: ${e.orcRef}">${e.orcRef}</span>`:'';const andBadge=e?.emAndamento?`<span class="ifc-andamento-badge" title="Em andamento">🟠</span>`:'';html+=`<div class="ifc-tree-item ${cls}" data-gid="${el.gid}" onclick="IFC.selectEl('${el.gid}',event)"><span class="ifc-dot"></span><span style="font-size:11px">${ico}</span><span class="ifc-nm" title="${el.name}">${el.name}</span>${andBadge}${orcBadge}<span class="ifc-pct">${pct>=0?pct+'%':''}</span><span class="ifc-eye" title="${isHidden?'Reexibir':'Ocultar'}" onclick="event.stopPropagation();IFC.toggleHide('${el.gid}')">${isHidden?'○':'●'}</span></div>`;}
    }
    tree.innerHTML=html||'<div style="padding:20px;text-align:center;color:#6b5555;font-size:12px">Nenhum elemento</div>';
  }
  function getDash(){try{return typeof DASH!=='undefined'?DASH:null;}catch(ex){return null;}}
  function buildOrcSel(){const sel=document.getElementById('ifc-orc-sel');const dash=getDash();if(!sel||!dash)return;const opts=['<option value="">— sem vínculo —</option>'];(dash.orc_tree||[]).forEach(n1=>(n1.children||[]).forEach(n2=>(n2.children||[]).forEach(n3=>{opts.push(`<option value="${n3.cod}">[${n3.cod}] ${(n3.nm||'').slice(0,50)}</option>`);})));sel.innerHTML=opts.join('');}
  function flattenOrcN3(){const map={},dash=getDash();(dash?.orc_tree||[]).forEach(n1=>(n1.children||[]).forEach(n2=>(n2.children||[]).forEach(n3=>{map[n3.cod]=n3;})));return map;}
  const LEVEL_ORC_ALIASES={'SS1':'SS1','SS2':'SS2','PG1':'G1','PG2':'G2','TER':'T','TÉRREO':'T','TERREO':'T'};
  function matchOrcForLevel(lvl,n3List){
    const key=(lvl||'').toString().trim().toUpperCase();
    const alias=LEVEL_ORC_ALIASES[key]||key;
    const esc=alias.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
    let hit=n3List.find(n3=>new RegExp('CONCRETO\\s+'+esc+'(\\s|\\+|$)','i').test(n3.nm||''));
    if(hit)return hit;
    const num=key.match(/(\d+)/);
    if(num){hit=n3List.find(n3=>new RegExp('\\b'+num[1]+'\\s*PAV\\b','i').test(n3.nm||''));if(hit)return hit;}
    return null;
  }
  function autoVincularPavimentos(){
    const dash=getDash();
    if(!dash){setStatus('⚠ Orçamento (API) não carregado — não é possível vincular por pavimento.');return;}
    const n3List=[];(dash.orc_tree||[]).forEach(n1=>(n1.children||[]).forEach(n2=>(n2.children||[]).forEach(n3=>n3List.push(n3))));
    const byLevel={};
    allElements.forEach(el=>{(byLevel[el.level]=byLevel[el.level]||[]).push(el.gid);});
    const matched=[],unmatched=[];let totalLinked=0;
    for(const lvl in byLevel){
      const hit=matchOrcForLevel(lvl,n3List);
      if(!hit){unmatched.push(lvl);continue;}
      byLevel[lvl].forEach(gid=>{
        if(!meshMap[gid])return;
        meshMap[gid].orcRef=hit.cod;
        medStorage[gid]=medStorage[gid]||{pct:meshMap[gid].pct,ts:Date.now()};
        medStorage[gid].orcRef=hit.cod;
        totalLinked++;
      });
      matched.push(lvl+' → ['+hit.cod+']');
    }
    saveStorage();refreshColors();renderTree();refreshOrcPanelIfOpen();
    let msg='🔗 '+totalLinked+' elemento(s) vinculado(s) automaticamente — '+matched.join(', ')+'.';
    if(unmatched.length)msg+=' Não identificados (vincule manualmente): '+unmatched.join(', ')+'.';
    setStatus(msg);
  }
  const brl=v=>'R$ '+(Number(v)||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
  function toggleOrcPanel(){const panel=document.getElementById('ifc-orc-panel');if(!panel)return;const opening=panel.style.display==='none';document.getElementById('ifc-pav-panel')?.style.setProperty('display','none');document.getElementById('ifc-quant-panel')?.style.setProperty('display','none');document.getElementById('ifc-compat-panel')?.style.setProperty('display','none');panel.style.display=opening?'flex':'none';if(opening)renderOrcVinculoTable();}

  // ── PAVIMENTOS (lista/seleção por pavimento) ────────────────────────
  let pavSelected=new Set();
  function pavimentoGroups(){
    const byLevel={};
    allElements.forEach(el=>{
      const lvl=el.level||'Sem pavimento';
      const g=byLevel[lvl]=byLevel[lvl]||{gids:[],vol:0,linked:0,measured:0};
      g.gids.push(el.gid); g.vol+=el.volume||0;
      const e=meshMap[el.gid];
      if(e?.orcRef)g.linked++;
      if(e?.pct>=100)g.measured++;
    });
    return byLevel;
  }
  // ── LAYOUT BIM: agregadores puros (sem UI), usados pelo painel novo ──
  // Agrupa elementos por (ObjectType, Name) distintos — os unicos campos de
  // "propriedade" confiaveis nos exports IFC do TQS usados neste projeto
  // (Material/Tag vem vazios, confirmado no modelo real).
  function propriedadesElementosTable(){
    const groups={};
    allElements.forEach(el=>{
      const mark=el.name||'—', artigo=el.objectType||el.type;
      const key=artigo+'|'+mark;
      const g=groups[key]=groups[key]||{mark,artigo,count:0,volume:0};
      g.count++; g.volume+=el.volume||0;
    });
    return Object.values(groups).sort((a,b)=>a.artigo.localeCompare(b.artigo)||a.mark.localeCompare(b.mark,undefined,{numeric:true}));
  }
  // Pivota volume real (m³) por [piso][categoria construtiva] — alimenta o
  // grafico de barras "Volume por elemento construtivo" do layout BIM.
  function volumePorPavimentoETipo(){
    const byLevel={};
    allElements.forEach(el=>{
      const lvl=el.level||'Sem pavimento', cat=categoriaConstrutiva(el.type);
      const g=byLevel[lvl]=byLevel[lvl]||{};
      g[cat]=(g[cat]||0)+(el.volume||0);
    });
    return byLevel;
  }
  // Agrupa por (categoria construtiva, secao transversal) — ex.: "Pilares
  // 0.20x3.00" — somando contagem/volume/area de forma. Sem colunas de
  // moldes: essa informacao (reaproveitamento de forma) e uma regra de
  // negocio que nao existe nos dados do modelo, nao e derivavel so da
  // geometria.
  function computeQuantPorSecao(){
    const groups={};
    allElements.forEach(el=>{
      const cat=categoriaConstrutiva(el.type);
      const d=el.sectionDims;
      const tipo=d?('w' in d?`${d.w.toFixed(2)}x${d.h.toFixed(2)}`:`Ø${(d.r*2).toFixed(2)}`):'—';
      const key=cat+'|'+tipo;
      const g=groups[key]=groups[key]||{categoria:cat,tipo,count:0,volume:0,formworkArea:0};
      g.count++; g.volume+=el.volume||0; g.formworkArea+=el.formworkArea||0;
    });
    return Object.values(groups).sort((a,b)=>a.categoria.localeCompare(b.categoria)||a.tipo.localeCompare(b.tipo,undefined,{numeric:true}));
  }
  function togglePavPanel(){const panel=document.getElementById('ifc-pav-panel');if(!panel)return;const opening=panel.style.display==='none';document.getElementById('ifc-orc-panel')?.style.setProperty('display','none');document.getElementById('ifc-quant-panel')?.style.setProperty('display','none');document.getElementById('ifc-compat-panel')?.style.setProperty('display','none');panel.style.display=opening?'flex':'none';if(opening)renderPavPanel();}
  function toggleQuantPanel(){const panel=document.getElementById('ifc-quant-panel');if(!panel)return;const opening=panel.style.display==='none';document.getElementById('ifc-orc-panel')?.style.setProperty('display','none');document.getElementById('ifc-pav-panel')?.style.setProperty('display','none');document.getElementById('ifc-compat-panel')?.style.setProperty('display','none');panel.style.display=opening?'flex':'none';if(opening)renderQuantPanel();}
  function renderPavPanel(){
    const body=document.getElementById('ifc-pav-panel-body'); if(!body)return;
    const groups=pavimentoGroups();
    const levels=Object.keys(groups).sort();
    if(!levels.length){body.innerHTML='<div style="padding:16px;text-align:center;color:#6b5555;font-size:12px">Nenhum modelo carregado ainda.</div>';return;}
    const rows=levels.map(lvl=>{
      const g=groups[lvl],enc=encodeURIComponent(lvl),chk=pavSelected.has(lvl)?'checked':'';
      return `<tr class="ifc-pav-row"><td><input type="checkbox" class="ifc-pav-chk" data-lvl="${enc}" ${chk}></td>
        <td class="ifc-pav-nmcell" data-lvl="${enc}" style="cursor:pointer;text-align:left">${lvl}</td>
        <td>${g.gids.length}</td>
        <td>${g.vol.toLocaleString('pt-BR',{maximumFractionDigits:2})} m³</td>
        <td>${g.linked}/${g.gids.length}</td>
        <td>${g.measured}/${g.gids.length}</td>
        <td><span class="ifc-pav-isolate" data-lvl="${enc}" title="Mostrar só este pavimento no 3D">👁 Isolar</span></td></tr>`;
    }).join('');
    body.innerHTML=`<div style="font-size:11px;color:#cfc2c2">${levels.length} pavimento(s) · ${allElements.length} elemento(s) no total — clique no nome para selecionar, marque as caixas para combinar vários</div>
    <table class="ifc-orc-table ifc-pav-table"><thead><tr><th></th><th style="text-align:left">Pavimento</th><th>Elementos</th><th>Volume</th><th>Vinculados</th><th>Medidos</th><th>Visão</th></tr></thead><tbody>${rows}</tbody></table>
    <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap">
      <button class="ifc-btn" id="ifc-pav-sel-btn">Selecionar marcados</button>
      <button class="ifc-btn" id="ifc-pav-show-all-btn">Mostrar todos os pavimentos</button>
    </div>`;
    body.querySelectorAll('.ifc-pav-nmcell').forEach(el=>el.addEventListener('click',()=>selectPavimento(decodeURIComponent(el.dataset.lvl))));
    body.querySelectorAll('.ifc-pav-isolate').forEach(el=>el.addEventListener('click',e=>{e.stopPropagation();isolarPavimento(decodeURIComponent(el.dataset.lvl));}));
    body.querySelectorAll('.ifc-pav-chk').forEach(el=>el.addEventListener('click',e=>{e.stopPropagation();togglePavCheck(decodeURIComponent(el.dataset.lvl));}));
    document.getElementById('ifc-pav-sel-btn')?.addEventListener('click',selectPavimentosMarcados);
    document.getElementById('ifc-pav-show-all-btn')?.addEventListener('click',mostrarTodosPavimentos);
  }
  let currentPavView=null,currentPavGids=null,currentPavType='all';
  const PAV_TYPE_LABELS={all:'Todos',COLUMN:'Pilares',BEAM:'Vigas',SLAB:'Lajes'};
  function updatePavIndicator(){
    const el=document.getElementById('ifc-pav-indicator'); if(!el)return;
    if(!currentPavView){el.style.display='none';return;}
    el.style.display='flex';
    const typeSuffix=currentPavType!=='all'?' · '+PAV_TYPE_LABELS[currentPavType]:'';
    el.querySelector('.ifc-pav-indicator-txt').textContent='🏢 Visualizando: '+currentPavView+typeSuffix;
    el.querySelectorAll('.ifc-pav-type-btn').forEach(b=>b.classList.toggle('act',b.dataset.type===currentPavType));
  }
  // aplica (ou remove) um filtro por tipo — Pilares/Vigas/Lajes — dentro do
  // pavimento atualmente em vista, isolando so esse subconjunto no 3D.
  function filterPavimentoByType(typeKey){
    if(!currentPavGids){setStatus('⚠ Selecione um pavimento primeiro (aba "Pavimentos").');return;}
    currentPavType=typeKey;
    const base=[...currentPavGids];
    const ids=typeKey==='all'?base:base.filter(gid=>(meshMap[gid]?.data?.type||'').includes(typeKey));
    selectedIds=new Set(ids);
    hiddenIds=new Set(allElements.filter(el=>!selectedIds.has(el.gid)).map(el=>el.gid));
    refreshColors();renderTree();showInfoPanelMulti();updateHiddenCount();updatePavIndicator();
    setStatus('🔎 '+PAV_TYPE_LABELS[typeKey]+' em '+currentPavView+': '+ids.length+' elemento(s).');
  }
  function selectPavimento(lvl){
    const g=pavimentoGroups()[lvl]; if(!g)return;
    selectedIds=new Set(g.gids);
    // "Selecionar" um pavimento tambem isola ele no 3D (esconde os demais) —
    // sem isso, elementos de outros andares continuavam visiveis e
    // clicaveis "atras"/perto dos deste pavimento, confundindo o usuario
    // que esperava estar vendo/interagindo só com o pavimento escolhido.
    const keep=new Set(g.gids);
    hiddenIds=new Set(allElements.filter(el=>!keep.has(el.gid)).map(el=>el.gid));
    currentPavView=lvl;currentPavGids=keep;currentPavType='all';
    refreshColors();renderTree();showInfoPanelMulti();updateHiddenCount();updatePavIndicator();
    setStatus('🏢 '+lvl+': '+g.gids.length+' elemento(s) selecionado(s) e isolado(s) no 3D · '+g.vol.toLocaleString('pt-BR',{maximumFractionDigits:2})+' m³. Use "Mostrar todos os pavimentos" para reverter.');
  }
  function togglePavCheck(lvl){if(pavSelected.has(lvl))pavSelected.delete(lvl);else pavSelected.add(lvl);}
  function selectPavimentosMarcados(){
    if(!pavSelected.size){setStatus('⚠ Marque ao menos um pavimento na lista.');return;}
    const groups=pavimentoGroups();const ids=new Set();
    pavSelected.forEach(lvl=>{(groups[lvl]?.gids||[]).forEach(gid=>ids.add(gid));});
    selectedIds=ids;
    hiddenIds=new Set(allElements.filter(el=>!ids.has(el.gid)).map(el=>el.gid));
    currentPavView=[...pavSelected].sort().join(', ');currentPavGids=ids;currentPavType='all';
    refreshColors();renderTree();showInfoPanelMulti();updateHiddenCount();updatePavIndicator();
    setStatus('🏢 '+pavSelected.size+' pavimento(s): '+ids.size+' elemento(s) selecionado(s) e isolado(s) no 3D.');
  }
  function isolarPavimento(lvl){
    const g=pavimentoGroups()[lvl]; if(!g)return;
    const keep=new Set(g.gids);
    hiddenIds=new Set(allElements.filter(el=>!keep.has(el.gid)).map(el=>el.gid));
    currentPavView=lvl;currentPavGids=keep;currentPavType='all';
    refreshColors();renderTree();updateHiddenCount();updatePavIndicator();
    setStatus('👁 Isolado: apenas o pavimento "'+lvl+'" está visível no 3D. Use "Mostrar todos" para reverter.');
  }
  function mostrarTodosPavimentos(){
    hiddenIds=new Set();
    currentPavView=null;currentPavGids=null;currentPavType='all';
    refreshColors();renderTree();updateHiddenCount();updatePavIndicator();
    setStatus('Todos os pavimentos visíveis novamente.');
  }
  function renderOrcVinculoTable(){
    const body=document.getElementById('ifc-orc-panel-body'); if(!body)return;
    const orcMap=flattenOrcN3();
    const groups={};
    Object.values(meshMap).forEach(e=>{
      if(!e.orcRef)return;
      const g=groups[e.orcRef]=groups[e.orcRef]||{count:0,volTotal:0,volMedido:0};
      g.count++; g.volTotal+=e.data.volume||0;
      if(e.pct>=100)g.volMedido+=e.data.volume||0; else if(e.pct>=1)g.volMedido+=(e.data.volume||0)*e.pct/100;
    });
    const codes=Object.keys(groups);
    if(!codes.length){body.innerHTML='<div style="padding:16px;text-align:center;color:#6b5555;font-size:12px">Nenhum elemento vinculado a um item de orçamento ainda.<br>Selecione um elemento, escolha o item em "Vincular ao ORC" e marque como medido.</div>';return;}
    const rows=codes.sort().map(cod=>{
      const orc=orcMap[cod]||{};
      const g=groups[cod];
      const orcado=Number(orc.orcado||0),realizado=Number(orc.realizado||0),aPagar=Number(orc.a_pagar||0);
      const saldo=orcado-realizado-aPagar;
      const achou=orcMap[cod]!=null;
      return `<tr><td>${cod}${achou?'':' <span style="color:#f97316">(não encontrado na API)</span>'}</td><td>${(orc.nm||'—').slice(0,60)}</td><td>${g.count}</td><td>${g.volTotal.toLocaleString('pt-BR',{maximumFractionDigits:2})} m³</td><td>${g.volMedido.toLocaleString('pt-BR',{maximumFractionDigits:2})} m³</td><td>${brl(orcado)}</td><td>${brl(realizado)}</td><td>${brl(aPagar)}</td><td class="${saldo<0?'neg':''}">${brl(saldo)}</td></tr>`;
    }).join('');
    const totalVol=codes.reduce((s,c)=>s+groups[c].volTotal,0), totalVolMed=codes.reduce((s,c)=>s+groups[c].volMedido,0);
    body.innerHTML=`<div style="font-size:11px;color:var(--tx3)">${codes.length} itens de orçamento vinculados · ${totalVol.toLocaleString('pt-BR',{maximumFractionDigits:2})} m³ no total · ${totalVolMed.toLocaleString('pt-BR',{maximumFractionDigits:2})} m³ medidos</div>
    <table class="ifc-orc-table"><thead><tr><th>Código EAP</th><th>Descrição</th><th>Elementos</th><th>Volume total (IFC)</th><th>Volume medido (IFC)</th><th>Orçado</th><th>Realizado</th><th>A Pagar</th><th>Saldo</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  // ── QUANTITATIVOS (volume, forma, aço) ──────────────────────────────
  const QUANT_TYPE_LABEL={IFCCOLUMN:'Pilares',IFCBEAM:'Vigas',IFCSLAB:'Lajes',IFCFOOTING:'Sapatas',IFCSTAIR:'Escadas',IFCRAMP:'Rampas'};
  function quantTypeLabel(t){return QUANT_TYPE_LABEL[t]||(t||'').replace('IFC','');}
  function renderQuantPanel(){
    const body=document.getElementById('ifc-quant-panel-body'); if(!body)return;
    if(!allElements.length){body.innerHTML='<div style="padding:16px;text-align:center;color:#6b5555;font-size:12px">Carregue um modelo IFC primeiro.</div>';return;}
    const groups={};
    allElements.forEach(el=>{
      const g=groups[el.type]=groups[el.type]||{count:0,volume:0,formworkArea:0,steelWeight:0};
      g.count++; g.volume+=el.volume||0; g.formworkArea+=el.formworkArea||0; g.steelWeight+=el.steelWeight||0;
    });
    const types=Object.keys(groups).sort((a,b)=>groups[b].volume-groups[a].volume);
    const totals=types.reduce((t,ty)=>({count:t.count+groups[ty].count,volume:t.volume+groups[ty].volume,formworkArea:t.formworkArea+groups[ty].formworkArea,steelWeight:t.steelWeight+groups[ty].steelWeight}),{count:0,volume:0,formworkArea:0,steelWeight:0});
    const rateInputs=types.map(ty=>{
      const hasRate=ty in QUANT_TYPE_LABEL;
      return hasRate?`<div class="ifc-quant-rate"><span>${quantTypeLabel(ty)}</span><input type="number" min="0" step="1" id="ifc-quant-rate-${ty}" value="${steelRateFor(ty)}"><span>kg/m³</span></div>`:'';
    }).filter(Boolean).join('');
    const rows=types.map(ty=>{
      const g=groups[ty];
      return `<tr><td>${quantTypeLabel(ty)}</td><td>${g.count}</td><td>${g.volume.toLocaleString('pt-BR',{maximumFractionDigits:2})} m³</td><td>${g.formworkArea.toLocaleString('pt-BR',{maximumFractionDigits:2})} m²</td><td>${(g.steelWeight/1000).toLocaleString('pt-BR',{maximumFractionDigits:2})} ton</td></tr>`;
    }).join('');
    body.innerHTML=`
      <div style="font-size:11px;color:var(--tx3);line-height:1.6;margin-bottom:10px">
        <b>Volume de concreto</b> e <b>área de forma</b> são calculados direto da geometria do IFC (perímetro da seção × comprimento extrudado — para vigas inclui a face de topo, que normalmente não é formada).<br>
        <b>Peso de aço</b> é uma <b style="color:#f5a623">estimativa</b> por taxa de consumo (kg de aço / m³ de concreto): o IFC do TQS não traz armadura modelada. Ajuste as taxas abaixo conforme o projeto e clique em Aplicar.
      </div>
      <table class="ifc-orc-table"><thead><tr><th>Tipo</th><th>Elementos</th><th>Volume concreto</th><th>Área de forma</th><th>Peso de aço (est.)</th></tr></thead><tbody>${rows}
        <tr style="font-weight:700;border-top:2px solid var(--bd)"><td>TOTAL</td><td>${totals.count}</td><td>${totals.volume.toLocaleString('pt-BR',{maximumFractionDigits:2})} m³</td><td>${totals.formworkArea.toLocaleString('pt-BR',{maximumFractionDigits:2})} m²</td><td>${(totals.steelWeight/1000).toLocaleString('pt-BR',{maximumFractionDigits:2})} ton</td></tr>
      </tbody></table>
      <div style="margin-top:14px;padding-top:10px;border-top:1px solid var(--bd)">
        <div style="font-size:11px;color:var(--tx3);margin-bottom:8px">Taxas de consumo de aço (editáveis):</div>
        <div class="ifc-quant-rates">${rateInputs}</div>
        <button class="ifc-btn act" style="margin-top:10px" onclick="IFC.applySteelRates()">Aplicar taxas e recalcular</button>
      </div>`;
  }
  function applySteelRates(){
    const rates={};
    Object.keys(QUANT_TYPE_LABEL).forEach(ty=>{
      const inp=document.getElementById('ifc-quant-rate-'+ty);
      if(inp)rates[ty]=Math.max(0,parseFloat(inp.value)||0);
    });
    saveSteelRates(rates);
    allElements.forEach(el=>{el.steelWeight=steelWeightOf(el);if(meshMap[el.gid]?.mesh)meshMap[el.gid].mesh.userData.steelWeight=el.steelWeight;});
    renderQuantPanel();
    showInfoPanelMulti();
    setStatus('⚙️ Taxas de aço atualizadas e quantitativos recalculados.');
  }

  // ── COMPATIBILIZAÇÃO: modelos adicionais + detecção de interferência ──
  // Cada modelo extra é carregado num Group próprio (translúcido, cor por
  // disciplina), somado à cena do modelo principal, sem tocar em
  // meshMap/medição — é só overlay visual + dados pra clash-check/materiais.
  function discColor(i){ return DISC_PALETTE[i % DISC_PALETTE.length]; }

  function loadCompatModel(file, disciplina){
    return new Promise((resolve,reject)=>{
      if(!scene){ reject(new Error('Carregue o modelo principal primeiro (IFC.init precisa da cena 3D).')); return; }
      readFileInChunks(file,
        parsed=>{
          try{
            const id='cm'+(++compatSeq);
            const cor=discColor(compatModels.length);
            const group=new THREE.Group();
            group.scale.setScalar(parsed.unitScale||1);
            group.rotation.x=-Math.PI/2;
            const elements=[];
            parsed.elements.forEach(el=>{
              const mat=new THREE.MeshLambertMaterial({color:cor,transparent:true,opacity:.32,side:THREE.DoubleSide,depthWrite:false});
              const mesh=new THREE.Mesh(el.geom,mat);
              group.add(mesh);
              elements.push({gid:el.gid,type:el.type,name:el.name,level:el.level,mesh,vol:0});
            });
            scene.add(group);
            group.updateMatrixWorld(true);
            const unitScale3=Math.pow(parsed.unitScale||1,3);
            elements.forEach(e=>{
              e.box=new THREE.Box3().setFromObject(e.mesh);
              e.vol=meshVolume(e.mesh.geometry)*unitScale3;
            });
            compatModels.push({id,nome:file.name,disciplina:disciplina||file.name,cor,group,elements,visible:true});
            renderCompatPanel();
            setStatus('✅ Modelo adicional carregado: '+(disciplina||file.name)+' ('+elements.length+' elementos)');
            resolve({id,count:elements.length});
          }catch(ex){ reject(ex); }
        },
        err=>reject(err)
      );
    });
  }

  function removeCompatModel(id){
    const i=compatModels.findIndex(m=>m.id===id); if(i===-1)return;
    const m=compatModels[i];
    scene.remove(m.group);
    m.group.traverse(o=>{if(o.geometry)o.geometry.dispose();if(o.material)o.material.dispose();});
    compatModels.splice(i,1);
    clashes=clashes.filter(c=>c.a.modelId!==id&&c.b.modelId!==id);
    renderCompatPanel();
  }

  function toggleCompatModel(id){
    const m=compatModels.find(m=>m.id===id); if(!m)return;
    m.visible=!m.visible; m.group.visible=m.visible;
    renderCompatPanel();
  }

  function renameCompatModel(id,nome){
    const m=compatModels.find(m=>m.id===id); if(!m)return;
    m.disciplina=nome||m.disciplina;
    renderCompatPanel();
  }

  // Sweep espacial (grid hash) pra não comparar cada elemento de um modelo
  // com TODOS os do outro — pra modelos de milhares de elementos (estrutura
  // inteira x MEP inteiro) a força bruta O(N×M) trava o navegador.
  const CLASH_CELL = 2.0; // metros
  function gridIndex(elements){
    const idx=new Map();
    elements.forEach((el,i)=>{
      const b=el.box;
      const x0=Math.floor(b.min.x/CLASH_CELL),x1=Math.floor(b.max.x/CLASH_CELL);
      const y0=Math.floor(b.min.y/CLASH_CELL),y1=Math.floor(b.max.y/CLASH_CELL);
      const z0=Math.floor(b.min.z/CLASH_CELL),z1=Math.floor(b.max.z/CLASH_CELL);
      for(let x=x0;x<=x1;x++)for(let y=y0;y<=y1;y++)for(let z=z0;z<=z1;z++){
        const k=x+'_'+y+'_'+z;
        if(!idx.has(k))idx.set(k,[]);
        idx.get(k).push(i);
      }
    });
    return idx;
  }
  function candidatesFor(el,idx){
    const b=el.box, set=new Set();
    const x0=Math.floor(b.min.x/CLASH_CELL),x1=Math.floor(b.max.x/CLASH_CELL);
    const y0=Math.floor(b.min.y/CLASH_CELL),y1=Math.floor(b.max.y/CLASH_CELL);
    const z0=Math.floor(b.min.z/CLASH_CELL),z1=Math.floor(b.max.z/CLASH_CELL);
    for(let x=x0;x<=x1;x++)for(let y=y0;y<=y1;y++)for(let z=z0;z<=z1;z++){
      const arr=idx.get(x+'_'+y+'_'+z); if(arr)arr.forEach(i=>set.add(i));
    }
    return set;
  }
  function clashBetween(elsA,elsB,tol){
    const idxB=gridIndex(elsB), out=[];
    elsA.forEach(elA=>{
      candidatesFor(elA,idxB).forEach(j=>{
        const elB=elsB[j];
        if(!elA.box.intersectsBox(elB.box))return;
        const ix=elA.box.clone().intersect(elB.box);
        const sz=ix.getSize(new THREE.Vector3());
        const vol=Math.max(0,sz.x)*Math.max(0,sz.y)*Math.max(0,sz.z);
        if(vol>tol)out.push({a:elA,b:elB,volume:vol,center:ix.getCenter(new THREE.Vector3())});
      });
    });
    return out;
  }
  function primaryGroupForClash(){
    if(!allElements.length)return null;
    const elements=allElements.map(el=>{
      const e=meshMap[el.gid];
      if(e&&e.mesh){ if(!e.box)e.box=new THREE.Box3().setFromObject(e.mesh); }
      return {gid:el.gid,type:el.type,name:el.name,level:el.level,box:e?.box,mesh:e?.mesh,modelId:'__principal__'};
    }).filter(e=>e.box);
    return {id:'__principal__',disciplina:'Modelo principal',elements};
  }
  function runClashDetection(tol){
    tol=Math.max(0,tol||0.005);
    clashes=[];
    const groups=[];
    const principal=primaryGroupForClash();
    if(principal)groups.push(principal);
    compatModels.filter(m=>m.visible).forEach(m=>{
      groups.push({id:m.id,disciplina:m.disciplina,elements:m.elements.map(e=>({...e,modelId:m.id}))});
    });
    if(groups.length<2){
      setStatus('⚠️ Carregue pelo menos 2 modelos (principal + 1 adicional) pra detectar interferência.');
      renderCompatPanel();
      return;
    }
    setStatus('🔍 Detectando interferências entre '+groups.length+' modelos…');
    for(let i=0;i<groups.length;i++)for(let j=i+1;j<groups.length;j++){
      clashBetween(groups[i].elements,groups[j].elements,tol).forEach(c=>{
        clashes.push({
          a:{...c.a,disciplina:groups[i].disciplina},
          b:{...c.b,disciplina:groups[j].disciplina},
          volume:c.volume,center:c.center,
        });
      });
    }
    clashes.sort((x,y)=>y.volume-x.volume);
    setStatus('✅ '+clashes.length+' interferência(s) encontrada(s) (tolerância '+tol+' m³).');
    renderCompatPanel();
  }
  let _clashFlashTimer=null;
  function focusClash(i){
    const c=clashes[i]; if(!c)return;
    _target.copy(c.center); _radius=6; updateCamera();
    [c.a,c.b].forEach(el=>{
      const mesh=el.mesh; if(!mesh)return;
      mesh.material.emissive=new THREE.Color(0xff2020); mesh.material.opacity=1;
    });
    if(_clashFlashTimer)clearTimeout(_clashFlashTimer);
    _clashFlashTimer=setTimeout(()=>{
      [c.a,c.b].forEach(el=>{ if(el.mesh){ el.mesh.material.emissive=new THREE.Color(0); } });
      refreshColors();
    },3500);
  }
  function exportClashesCSV(){
    const rows=[['Disciplina A','Elemento A','Disciplina B','Elemento B','Volume de sobreposição (m³)']];
    clashes.forEach(c=>rows.push([c.a.disciplina,c.a.name||c.a.type,c.b.disciplina,c.b.name||c.b.type,c.volume.toFixed(4)]));
    const csv=rows.map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');
    const a=document.createElement('a');a.href='data:text/csv;charset=utf-8,﻿'+encodeURIComponent(csv);a.download='interferencias_ifc_'+new Date().toISOString().slice(0,10)+'.csv';a.click();
  }

  // ── LEVANTAMENTO DE MATERIAIS (multi-disciplina) ────────────────────
  const MAT_TYPE_LABEL={IFCWALL:'Paredes',IFCWALLSTANDARDCASE:'Paredes',IFCSLAB:'Lajes',IFCBEAM:'Vigas',IFCCOLUMN:'Pilares',IFCDOOR:'Portas',IFCWINDOW:'Janelas',IFCSTAIR:'Escadas',IFCSTAIRFLIGHT:'Escadas',IFCROOF:'Coberturas',IFCRAMP:'Rampas',IFCPLATE:'Chapas',IFCMEMBER:'Perfis',IFCCURTAINWALL:'Fachada cortina',IFCFURNISHINGELEMENT:'Mobiliário',IFCFLOWSEGMENT:'Trechos de tubulação/duto',IFCFLOWTERMINAL:'Terminais (MEP)',IFCBUILDINGELEMENTPROXY:'Elementos genéricos',IFCPILE:'Estacas',IFCFOOTING:'Sapatas/blocos',IFCCOVERING:'Revestimentos'};
  function matTypeLabel(t){return MAT_TYPE_LABEL[t]||(t||'').replace('IFC','');}
  function computeMateriais(){
    const rows={};
    function add(disciplina,type,vol){
      const k=disciplina+'|'+type;
      rows[k]=rows[k]||{disciplina,type,count:0,volume:0};
      rows[k].count++; rows[k].volume+=vol||0;
    }
    allElements.forEach(el=>add('Modelo principal',el.type,meshMap[el.gid]?.data?.volume||el.volume||0));
    compatModels.forEach(m=>m.elements.forEach(el=>add(m.disciplina,el.type,el.vol)));
    return Object.values(rows).sort((a,b)=>a.disciplina.localeCompare(b.disciplina,'pt')||b.volume-a.volume);
  }
  function exportMateriaisCSV(){
    const rows=[['Disciplina','Tipo','Quantidade','Volume (m³)']];
    computeMateriais().forEach(r=>rows.push([r.disciplina,matTypeLabel(r.type),r.count,r.volume.toFixed(4)]));
    const csv=rows.map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');
    const a=document.createElement('a');a.href='data:text/csv;charset=utf-8,﻿'+encodeURIComponent(csv);a.download='materiais_ifc_'+new Date().toISOString().slice(0,10)+'.csv';a.click();
  }

  function renderCompatPanel(){
    const body=document.getElementById('ifc-compat-panel-body'); if(!body)return;
    const modelosRows=compatModels.map(m=>`
      <tr>
        <td><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#${m.cor.toString(16).padStart(6,'0')};margin-right:6px"></span>
          <input class="ifc-compat-disc-input" value="${(m.disciplina||'').replace(/"/g,'&quot;')}" onchange="IFC.renameCompatModel('${m.id}',this.value)"></td>
        <td>${m.nome}</td>
        <td>${m.elements.length}</td>
        <td><button class="ifc-btn" style="padding:2px 8px;font-size:11px" onclick="IFC.toggleCompatModel('${m.id}')">${m.visible?'👁️ Visível':'🚫 Oculto'}</button></td>
        <td><button class="ifc-btn" style="padding:2px 8px;font-size:11px;color:#ff8080" onclick="IFC.removeCompatModel('${m.id}')">Remover</button></td>
      </tr>`).join('');
    const clashRows=clashes.slice(0,300).map((c,i)=>`
      <tr>
        <td>${c.a.disciplina}<br><span style="font-size:10px;color:var(--tx3)">${(c.a.name||c.a.type||'').slice(0,40)}</span></td>
        <td>${c.b.disciplina}<br><span style="font-size:10px;color:var(--tx3)">${(c.b.name||c.b.type||'').slice(0,40)}</span></td>
        <td>${(c.volume*1000).toLocaleString('pt-BR',{maximumFractionDigits:1})} L</td>
        <td><button class="ifc-btn" style="padding:2px 8px;font-size:11px" onclick="IFC.focusClash(${i})">🎯 Ver</button></td>
      </tr>`).join('');
    const materiais=computeMateriais();
    const matRows=materiais.map(r=>`<tr><td>${r.disciplina}</td><td>${matTypeLabel(r.type)}</td><td>${r.count}</td><td>${r.volume.toLocaleString('pt-BR',{maximumFractionDigits:3})} m³</td></tr>`).join('');

    body.innerHTML=`
      <div style="font-size:11px;color:var(--tx3);line-height:1.6;margin-bottom:12px">
        Carregue o modelo de <b>outra disciplina</b> (ex.: elétrico, hidráulico) por cima do modelo principal já aberto no visualizador — cada um fica com uma cor translúcida própria. Depois clique em <b>Detectar interferências</b> pra cruzar a geometria dos modelos e achar sobreposições reais (não meras encostas: ajuste a tolerância se aparecer ruído).
      </div>
      <div style="margin-bottom:6px;font-weight:700;font-size:12px">Modelos adicionais carregados</div>
      <table class="ifc-orc-table"><thead><tr><th>Disciplina</th><th>Arquivo</th><th>Elementos</th><th>Exibição</th><th></th></tr></thead>
        <tbody>${modelosRows||'<tr><td colspan="5" style="text-align:center;color:var(--tx3);padding:10px">Nenhum modelo adicional carregado ainda.</td></tr>'}</tbody></table>
      <div style="display:flex;gap:8px;align-items:center;margin:10px 0 16px">
        <label class="ifc-btn" style="cursor:pointer">📂 Carregar .ifc adicional<input type="file" accept=".ifc" style="display:none" onchange="IFC.pickCompatFile(this.files[0])"></label>
      </div>

      <div style="display:flex;gap:10px;align-items:center;margin-bottom:6px">
        <span style="font-weight:700;font-size:12px">Detecção de interferências</span>
        <label style="font-size:11px;color:var(--tx3)">Tolerância (m³ de sobreposição): <input type="number" id="ifc-clash-tol" value="0.005" min="0" step="0.001" style="width:70px;padding:2px 4px;background:var(--sf2);border:1px solid var(--bd);border-radius:4px;color:var(--tx)"></label>
        <button class="ifc-btn act" onclick="IFC.runClashDetection(parseFloat(document.getElementById('ifc-clash-tol').value)||0.005)">🔍 Detectar</button>
        ${clashes.length?`<button class="ifc-btn" onclick="IFC.exportClashesCSV()">⬇️ CSV</button>`:''}
      </div>
      <table class="ifc-orc-table"><thead><tr><th>Elemento A</th><th>Elemento B</th><th>Sobreposição</th><th></th></tr></thead>
        <tbody>${clashRows||'<tr><td colspan="4" style="text-align:center;color:var(--tx3);padding:10px">Nenhuma interferência detectada ainda.</td></tr>'}</tbody></table>
      ${clashes.length>300?`<div style="font-size:10px;color:var(--tx3);margin-top:4px">Mostrando as 300 maiores de ${clashes.length} interferências.</div>`:''}

      <div style="margin-top:16px;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center">
        <span style="font-weight:700;font-size:12px">Levantamento de materiais (todos os modelos carregados)</span>
        ${materiais.length?`<button class="ifc-btn" onclick="IFC.exportMateriaisCSV()">⬇️ CSV</button>`:''}
      </div>
      <table class="ifc-orc-table"><thead><tr><th>Disciplina</th><th>Tipo</th><th>Quantidade</th><th>Volume</th></tr></thead>
        <tbody>${matRows||'<tr><td colspan="4" style="text-align:center;color:var(--tx3);padding:10px">Carregue um modelo pra ver o levantamento.</td></tr>'}</tbody></table>
    `;
  }
  function toggleCompatPanel(){
    const panel=document.getElementById('ifc-compat-panel'); if(!panel)return;
    const opening=panel.style.display==='none';
    document.getElementById('ifc-orc-panel')?.style.setProperty('display','none');
    document.getElementById('ifc-pav-panel')?.style.setProperty('display','none');
    document.getElementById('ifc-quant-panel')?.style.setProperty('display','none');
    panel.style.display=opening?'flex':'none';
    if(opening)renderCompatPanel();
  }
  function pickCompatFile(file){
    if(!file)return;
    const disciplina=prompt('Nome da disciplina desse modelo (ex.: Elétrico, Hidrossanitário):', file.name.replace(/\.ifc$/i,''));
    if(disciplina===null)return;
    loadCompatModel(file,disciplina).catch(ex=>{ alert('Erro ao carregar modelo: '+ex.message); console.error(ex); });
  }
  async function carregarModeloAutoDoc(modelId,nomeArquivo,disciplina){
    setStatus('☁️ Baixando '+nomeArquivo+' do visualizador…');
    try{
      const r=await fetch('/api/ifc/modelos/'+encodeURIComponent(modelId)+'/arquivo');
      if(!r.ok)throw new Error('Modelo ainda não disponível localmente.');
      const blob=await r.blob();
      const file=new File([blob],nomeArquivo,{type:'application/octet-stream'});
      await loadCompatModel(file,disciplina||nomeArquivo);
    }catch(ex){
      alert('Erro ao carregar modelo do AutoDoc: '+ex.message);
      console.error(ex);
    }
  }

  // ── SELECT / MEDIR ────────────────────────────────────────────────
  function selectEl(gid,evt){const additive=evt&&(evt.ctrlKey||evt.metaKey||evt.shiftKey);if(additive){if(selectedIds.has(gid))selectedIds.delete(gid);else selectedIds.add(gid);}else if(selectedIds.size===1&&selectedIds.has(gid)){selectedIds=new Set();}else{selectedIds=new Set([gid]);}refreshColors();renderTree();showInfoPanelMulti();document.querySelector(`[data-gid="${gid}"]`)?.scrollIntoView({block:'nearest'});}
  function toggleHide(gid){if(hiddenIds.has(gid))hiddenIds.delete(gid);else hiddenIds.add(gid);refreshColors();renderTree();updateHiddenCount();}
  function hideSelected(){if(!selectedIds.size)return;selectedIds.forEach(gid=>hiddenIds.add(gid));const n=selectedIds.size;selectedIds=new Set();refreshColors();renderTree();updateHiddenCount();showInfoPanelMulti();setStatus(n+' elemento(s) ocultado(s).');}
  function unhideAll(){if(!hiddenIds.size)return;hiddenIds=new Set();currentPavView=null;refreshColors();renderTree();updateHiddenCount();updatePavIndicator();setStatus('Todos os elementos ocultos foram reexibidos.');}
  function updateHiddenCount(){const btn=document.getElementById('ifc-unhide-btn');if(btn)btn.textContent='Reexibir ocultos'+(hiddenIds.size?' ('+hiddenIds.size+')':'');}
  function showInfoPanelMulti(){
    const ids=[...selectedIds];
    const title=document.getElementById('ifc-info-title');
    if(ids.length===0){document.getElementById('ifc-info-panel')?.classList.remove('open');return;}
    if(ids.length===1){if(title)title.textContent='Elemento Selecionado';showInfoPanel(ids[0]);return;}
    if(title)title.textContent=ids.length+' elementos selecionados';
    const panel=document.getElementById('ifc-info-panel'),body=document.getElementById('ifc-info-body'),pctIn=document.getElementById('ifc-pct-input'),sel=document.getElementById('ifc-orc-sel');
    if(!panel||!body)return;
    let totalVol=0,totalForma=0,totalAco=0,measured=0,andamento=0;const refs=new Set();
    ids.forEach(gid=>{const e=meshMap[gid];if(!e)return;totalVol+=e.data.volume||0;totalForma+=e.data.formworkArea||0;totalAco+=e.data.steelWeight||0;if(e.pct>=100)measured++;if(e.emAndamento)andamento++;refs.add(e.orcRef||'');});
    body.innerHTML=[['Selecionados',ids.length],['Volume de concreto',totalVol.toLocaleString('pt-BR',{maximumFractionDigits:3})+' m³'],['Área de forma',totalForma.toLocaleString('pt-BR',{maximumFractionDigits:2})+' m²'],['Peso de aço (est.)',totalAco.toLocaleString('pt-BR',{maximumFractionDigits:1})+' kg'],['Já medidos',measured+' de '+ids.length],['Em andamento',andamento?`<span style="color:#f5a623">🟠 ${andamento} de ${ids.length}</span>`:'0 de '+ids.length]].map(([l,v])=>`<div class="ifc-info-row"><span class="ifc-info-lbl">${l}</span><span class="ifc-info-val">${v}</span></div>`).join('');
    if(pctIn)pctIn.value=100;
    if(sel)sel.value=refs.size===1?[...refs][0]:'';
    panel.classList.add('open');
  }
  function showInfoPanel(gid){const panel=document.getElementById('ifc-info-panel'),body=document.getElementById('ifc-info-body'),pctIn=document.getElementById('ifc-pct-input'),sel=document.getElementById('ifc-orc-sel');if(!panel||!body)return;const e=meshMap[gid];if(!e){panel.classList.remove('open');return;}const d=e.data;body.innerHTML=[['Nome',d.name],['Tipo',d.type.replace('IFC','')],['Nível',d.level],['Volume de concreto',(d.volume||0).toLocaleString('pt-BR',{maximumFractionDigits:3})+' m³'],['Área de forma',(d.formworkArea||0).toLocaleString('pt-BR',{maximumFractionDigits:2})+' m²'],['Peso de aço (est.)',(d.steelWeight||0).toLocaleString('pt-BR',{maximumFractionDigits:1})+' kg'],['GlobalId',d.gid.slice(0,16)+'…'],['Medido',e.pct>=0?e.pct+'%':'Não medido'],['Em andamento',e.emAndamento?'<span style="color:#f5a623">🟠 Sim</span>':'Não']].map(([l,v])=>`<div class="ifc-info-row"><span class="ifc-info-lbl">${l}</span><span class="ifc-info-val">${v}</span></div>`).join('');if(pctIn)pctIn.value=e.pct>=0?e.pct:100;if(sel)sel.value=e.orcRef||'';panel.classList.add('open');}
  function refreshOrcPanelIfOpen(){const panel=document.getElementById('ifc-orc-panel');if(panel&&panel.style.display!=='none')renderOrcVinculoTable();}
  function medir(){if(!selectedIds.size)return;const pct=Math.min(Math.max(parseInt(document.getElementById('ifc-pct-input')?.value)||100,0),100),orcRef=document.getElementById('ifc-orc-sel')?.value||'';let last='';const ts=Date.now();selectedIds.forEach(gid=>{if(!meshMap[gid])return;meshMap[gid].pct=pct;meshMap[gid].orcRef=orcRef;meshMap[gid].ts=ts;medStorage[gid]={pct,orcRef,emAndamento:meshMap[gid].emAndamento,ts};last=meshMap[gid].data.name;});saveStorage();refreshColors();renderTree();updateMedCount();showInfoPanelMulti();refreshOrcPanelIfOpen();setStatus(selectedIds.size>1?('✅ Medido: '+selectedIds.size+' elementos — '+pct+'%'):('✅ Medido: '+last+' — '+pct+'%'));}
  function desmarcar(){if(!selectedIds.size)return;let last='';selectedIds.forEach(gid=>{if(!meshMap[gid])return;meshMap[gid].pct=-1;meshMap[gid].orcRef='';if(meshMap[gid].emAndamento)medStorage[gid]={pct:-1,orcRef:'',emAndamento:true,ts:Date.now()};else delete medStorage[gid];meshMap[gid].ts=medStorage[gid]?.ts||0;last=meshMap[gid].data.name;});saveStorage();refreshColors();renderTree();updateMedCount();showInfoPanelMulti();refreshOrcPanelIfOpen();setStatus(selectedIds.size>1?('↩ Desmarcado: '+selectedIds.size+' elementos'):('↩ Desmarcado: '+last));}
  // "Em andamento" e independente do % medido/avanco fisico — so uma marcacao
  // manual de execucao em campo, com cor propria no 3D e percentual proprio.
  function toggleEmAndamento(){
    if(!selectedIds.size){setStatus('⚠ Selecione ao menos um elemento.');return;}
    const allOn=[...selectedIds].every(gid=>meshMap[gid]?.emAndamento);
    const novo=!allOn;
    selectedIds.forEach(gid=>{
      const e=meshMap[gid]; if(!e)return;
      e.emAndamento=novo;
      medStorage[gid]=medStorage[gid]||{pct:e.pct,orcRef:e.orcRef,ts:Date.now()};
      medStorage[gid].emAndamento=novo;
      if(!novo && medStorage[gid].pct<0 && !medStorage[gid].orcRef)delete medStorage[gid];
    });
    saveStorage();refreshColors();renderTree();updateMedCount();showInfoPanelMulti();
    setStatus((novo?'🟠 Marcado como em andamento: ':'⚪ Removido de em andamento: ')+selectedIds.size+' elemento(s).');
  }
  function medStats(){
    const entries=Object.values(meshMap),total=entries.length;
    const measured=entries.filter(e=>e.pct>=100).length,partial=entries.filter(e=>e.pct>=1&&e.pct<100).length;
    const pct=total?((measured+partial*.5)/total*100):0;
    let volTotal=0,volMedido=0,volAndamento=0,countAndamento=0,formaTotal=0,acoTotal=0;
    entries.forEach(e=>{const v=e.data.volume||0;volTotal+=v;if(e.pct>=100)volMedido+=v;else if(e.pct>=1)volMedido+=v*e.pct/100;if(e.emAndamento){volAndamento+=v;countAndamento++;}formaTotal+=e.data.formworkArea||0;acoTotal+=e.data.steelWeight||0;});
    const pctAndamento=volTotal?volAndamento/volTotal*100:0;
    return {total,measured,partial,pendentes:total-measured-partial,pct,volTotal,volMedido,volAndamento,countAndamento,pctAndamento,formaTotal,acoTotal};
  }
  function updateMedCount(){
    const s=medStats(); if(!s.total)return;
    const el=document.getElementById('ifc-med-count'),st=document.getElementById('ifc-med-stats');
    if(el)el.style.display='block';
    // "Em andamento" e uma marcacao manual separada — nao entra no calculo de
    // avanco fisico (que segue vindo so do % medido), por isso aparece numa
    // linha propria com seu proprio percentual.
    const andamentoLine=s.countAndamento?`<div style="margin-top:6px;font-size:11px;color:${'#f5a623'}">🟠 ${s.countAndamento} elemento(s) em andamento — ${s.pctAndamento.toFixed(1)}% do volume (separado do avanço físico)</div>`:'';
    if(st)st.innerHTML=`<span style="color:#4ade80">✅ ${s.measured}</span> medidos · <span style="color:#f97316">● ${s.partial}</span> parciais · <span style="color:#6b5555">${s.pendentes}</span> pendentes<br><div style="margin-top:4px;height:4px;background:#221818;border-radius:2px"><div style="width:${s.pct.toFixed(1)}%;height:100%;background:linear-gradient(90deg,#4ade80,#A50E24);border-radius:2px"></div></div><div style="text-align:right;font-size:10px;color:#6b5555;margin-top:2px">${s.pct.toFixed(1)}% avanço físico</div><div style="margin-top:6px;font-size:11px">Volume medido (concreto): <b>${s.volMedido.toLocaleString('pt-BR',{maximumFractionDigits:2})} m³</b><br><span style="color:#6b5555;font-size:10px">de ${s.volTotal.toLocaleString('pt-BR',{maximumFractionDigits:2})} m³ no modelo</span></div><div style="margin-top:6px;font-size:11px">Modelo inteiro: <b>${s.formaTotal.toLocaleString('pt-BR',{maximumFractionDigits:1})} m²</b> de forma · <b>${(s.acoTotal/1000).toLocaleString('pt-BR',{maximumFractionDigits:1})} ton</b> de aço (est.) <span style="color:#6b5555;font-size:10px">— ver 📐 Quantitativos</span></div>${andamentoLine}`;
    const dateIn=document.getElementById('ifc-med-save-date');
    if(dateIn&&!dateIn.value)dateIn.value=new Date().toISOString().slice(0,10);
  }
  async function salvarMedicao(){
    const s=medStats();
    const dateIn=document.getElementById('ifc-med-save-date');
    const dataRef=dateIn?.value||new Date().toISOString().slice(0,10);
    const items={}; for(const gid in meshMap){const e=meshMap[gid]; if(e.pct>=0)items[gid]={pct:e.pct,orcRef:e.orcRef||''};}
    const modelName=document.getElementById('ifc-file-name')?.textContent||'';
    try{
      const resp=await fetch('/api/ifc_medicao/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({
        data:dataRef,model_name:modelName,total_elementos:s.total,medidos:s.measured,parciais:s.partial,
        pct_avanco:Number(s.pct.toFixed(2)),volume_total:Number(s.volTotal.toFixed(3)),volume_medido:Number(s.volMedido.toFixed(3)),items
      })});
      const data=await resp.json();
      if(!resp.ok||data.ok===false)throw new Error(data.error||'Falha ao salvar.');
      setStatus('💾 Medição salva para '+dataRef+'.');
      const hist=document.getElementById('ifc-med-hist-list');
      if(hist&&hist.style.display!=='none')loadMedicaoHistorico();
    }catch(ex){setStatus('❌ Erro ao salvar medição: '+ex.message);}
  }
  async function toggleHistorico(){
    const list=document.getElementById('ifc-med-hist-list'),toggle=document.getElementById('ifc-med-hist-toggle');
    if(!list)return;
    const opening=list.style.display==='none';
    list.style.display=opening?'block':'none';
    if(toggle)toggle.textContent=(opening?'▴':'▾')+' Histórico de medições';
    if(opening)await loadMedicaoHistorico();
  }
  async function loadMedicaoHistorico(){
    const list=document.getElementById('ifc-med-hist-list'); if(!list)return;
    list.innerHTML='<div style="color:#6b5555;font-size:10px;padding:4px 0">Carregando…</div>';
    try{
      const resp=await fetch('/api/ifc_medicao/saves',{cache:'no-store'});
      const saves=await resp.json();
      if(!Array.isArray(saves)||!saves.length){list.innerHTML='<div style="color:#6b5555;font-size:10px;padding:4px 0">Nenhuma medição salva ainda.</div>';return;}
      list.innerHTML=saves.map(sv=>{
        const dt=(sv.data||'').split('-').reverse().join('/');
        return `<div class="ifc-med-hist-item"><div><b>${dt}</b> · ${(sv.pct_avanco||0).toLocaleString('pt-BR',{maximumFractionDigits:1})}%<br><span style="color:#6b5555">${(sv.volume_medido||0).toLocaleString('pt-BR',{maximumFractionDigits:1})} m³ · ${sv.medidos||0}/${sv.total_elementos||0} itens</span></div><span class="ifc-med-hist-del" title="Excluir" onclick="IFC.excluirMedicaoSalva('${sv.id}')">✕</span></div>`;
      }).join('');
    }catch(ex){list.innerHTML='<div style="color:#ff4d6d;font-size:10px;padding:4px 0">Erro ao carregar histórico.</div>';}
  }
  async function excluirMedicaoSalva(id){
    if(!confirm('Excluir esta medição salva?'))return;
    try{
      await fetch('/api/ifc_medicao/saves/'+id,{method:'DELETE'});
      loadMedicaoHistorico();
    }catch(ex){}
  }

  // ── STORAGE ──────────────────────────────────────────────────────
  const saveStorage=()=>{try{localStorage.setItem(STORAGE_KEY,JSON.stringify(medStorage));}catch(e){}};
  const loadStorage=()=>{try{medStorage=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');}catch(e){medStorage={};}};

  // ── CAMERA ──────────────────────────────────────────────────────────
  let _leftDown=false,_midDown=false,_rightDown=false,_touchDrag=false,_boxSelecting=false,_dragStartX=0,_dragStartY=0,_lastX=0,_lastY=0,_theta=45,_phi=60,_radius=80,_target=new THREE.Vector3(),_leftMode='select';
  const BOX_DRAG_THRESHOLD=4;
  // Controles: botao esquerdo segue o "modo" ativo (barra de ferramentas —
  // padrao Selecionar: clique seleciona, arrastar abre caixa de selecao,
  // igual ao que ja funcionava bem). Botao direito arrastando SEMPRE gira a
  // camera; botao do meio arrastando SEMPRE move a vista — atalhos fixos,
  // disponiveis em qualquer modo, sem precisar trocar de ferramenta.
  function panBy(dx,dy){const r=new THREE.Vector3().crossVectors(camera.position.clone().sub(_target),camera.up).normalize();_target.addScaledVector(r,-dx*.05*_radius/40);_target.y+=dy*.05*_radius/40;}
  function orbitBy(dx,dy){_theta-=dx*.4;_phi=Math.max(2,Math.min(178,_phi-dy*.4));}
  function updateBoxOverlay(wrap,e){const rect=wrap.getBoundingClientRect(),b=document.getElementById('ifc-select-box');if(!b)return;const x1=Math.min(_dragStartX,e.clientX)-rect.left,x2=Math.max(_dragStartX,e.clientX)-rect.left,y1=Math.min(_dragStartY,e.clientY)-rect.top,y2=Math.max(_dragStartY,e.clientY)-rect.top;b.style.left=x1+'px';b.style.top=y1+'px';b.style.width=(x2-x1)+'px';b.style.height=(y2-y1)+'px';b.style.display='block';}
  let _lastClickX=null,_lastClickY=null,_lastHitIds=[],_lastHitIdx=0;
  const CLICK_TOLERANCE_OFFSETS=[[0,0],[3,0],[-3,0],[0,3],[0,-3],[2,2],[-2,2],[2,-2],[-2,-2],[5,0],[-5,0],[0,5],[0,-5]];
  function raycastAtPixel(rect,x,y){
    raycaster.setFromCamera({x:((x-rect.left)/rect.width)*2-1,y:-((y-rect.top)/rect.height)*2+1},camera);
    // O Raycaster do three.js (r128) NAO pula objetos com visible=false — ele
    // raycasta tudo, mesmo o que esta oculto na tela. Sem este filtro
    // explicito, elementos escondidos (Isolar pavimento, filtro
    // Medidos/Pendentes, Ocultar selecao) continuavam sendo selecionaveis
    // "atraves" do que estava realmente visivel.
    return raycaster.intersectObjects(Object.values(meshMap).map(x=>x.mesh).filter(m=>m&&m.visible),false);
  }
  function doRaycastSelect(canvas,e){
    const rect=canvas.getBoundingClientRect();
    if(e.clientX<rect.left||e.clientX>rect.right||e.clientY<rect.top||e.clientY>rect.bottom)return;
    // Elementos com poucos pixels na tela (ferragens, pecas pequenas,
    // sobretudo em pavimentos com muita coisa junta como o SS1) podem escapar
    // do pixel exato do clique. Tenta o ponto exato primeiro; se nao acertar
    // nada, varre uma pequena vizinhanca ao redor antes de desistir.
    let hits=[];
    for(const [dx,dy] of CLICK_TOLERANCE_OFFSETS){
      hits=raycastAtPixel(rect,e.clientX+dx,e.clientY+dy);
      if(hits.length)break;
    }
    if(!hits.length){
      _lastHitIds=[];
      if(!(e.ctrlKey||e.metaKey||e.shiftKey)){selectedIds=new Set();refreshColors();renderTree();document.getElementById('ifc-info-panel')?.classList.remove('open');}
      return;
    }
    // Elementos podem se sobrepor no mesmo pixel (ex.: pilar bem proximo a
    // uma laje) — o hit mais proximo da camera nem sempre e o que o usuario
    // quer. Clicar de novo no mesmo ponto avanca para o proximo elemento
    // sob o cursor, em vez de travar sempre no mais proximo.
    const ids=hits.map(h=>h.object.userData.gid);
    const samePoint=_lastClickX!==null&&Math.hypot(e.clientX-_lastClickX,e.clientY-_lastClickY)<4&&ids.length===_lastHitIds.length&&ids.every((id,i)=>id===_lastHitIds[i]);
    const idx=samePoint?(_lastHitIdx+1)%ids.length:0;
    _lastClickX=e.clientX;_lastClickY=e.clientY;_lastHitIds=ids;_lastHitIdx=idx;
    selectEl(ids[idx],e);
    if(ids.length>1)setStatus((idx+1)+'/'+ids.length+' elementos sobrepostos neste ponto — clique de novo no mesmo lugar para alternar.');
  }
  function finishBoxSelect(canvas,e){
    const rect=canvas.getBoundingClientRect();
    const x1=Math.min(_dragStartX,e.clientX),x2=Math.max(_dragStartX,e.clientX),y1=Math.min(_dragStartY,e.clientY),y2=Math.max(_dragStartY,e.clientY);
    const additive=e.ctrlKey||e.metaKey||e.shiftKey;
    const picked=[],corner=new THREE.Vector3();
    for(const gid in meshMap){
      const mesh=meshMap[gid].mesh; if(!mesh||!mesh.visible)continue;
      if(!mesh.geometry.boundingBox)mesh.geometry.computeBoundingBox();
      const bb=mesh.geometry.boundingBox; if(!bb)continue;
      // testa sobreposicao usando os 8 cantos do bounding box projetados na
      // tela — usar so o centro (como antes) perdia elementos grandes
      // (lajes, vigas longas) cujo centro cai fora da caixa arrastada
      // mesmo com parte do elemento visivelmente dentro.
      let minX=Infinity,minY=Infinity,maxX=-Infinity,maxY=-Infinity,anyInFront=false;
      for(let i=0;i<8;i++){
        corner.set(i&1?bb.max.x:bb.min.x, i&2?bb.max.y:bb.min.y, i&4?bb.max.z:bb.min.z);
        corner.applyMatrix4(mesh.matrixWorld).project(camera);
        if(corner.z>=1)continue;
        anyInFront=true;
        const sx=rect.left+(corner.x*.5+.5)*rect.width, sy=rect.top+(-corner.y*.5+.5)*rect.height;
        if(sx<minX)minX=sx; if(sx>maxX)maxX=sx; if(sy<minY)minY=sy; if(sy>maxY)maxY=sy;
      }
      if(anyInFront && maxX>=x1 && minX<=x2 && maxY>=y1 && minY<=y2)picked.push(gid);
    }
    if(!additive)selectedIds=new Set();
    picked.forEach(gid=>selectedIds.add(gid));
    refreshColors();renderTree();showInfoPanelMulti();
    setStatus(picked.length+' elemento(s) selecionado(s) pela caixa.');
  }
  function setLeftMode(mode){_leftMode=mode;document.querySelectorAll('.ifc-mode-btn').forEach(b=>b.classList.toggle('act',b.dataset.mode===mode));}
  function setupControls(canvas){
    const wrap=canvas.parentElement;
    canvas.addEventListener('mousedown',e=>{
      _lastX=e.clientX;_lastY=e.clientY;
      if(e.button===0){_leftDown=true;_dragStartX=e.clientX;_dragStartY=e.clientY;_boxSelecting=false;}
      else if(e.button===1){_midDown=true;e.preventDefault();}
      else if(e.button===2){_rightDown=true;}
    });
    window.addEventListener('mouseup',e=>{
      if(_leftDown&&_leftMode==='select'){
        // _boxSelecting fica true assim que o mouse passa do threshold em
        // QUALQUER ponto do arrasto e nao volta a false — um clique com leve
        // tremor da mao (comum) podia disparar isso e cair na selecao por
        // caixa, que nao tem a logica de "clicar de novo desmarca". Decide
        // pelo deslocamento TOTAL entre mousedown e mouseup: se no final o
        // ponto voltou perto de onde comecou, trata como clique normal.
        const totalDrag=Math.hypot(e.clientX-_dragStartX,e.clientY-_dragStartY);
        if(_boxSelecting&&totalDrag>BOX_DRAG_THRESHOLD)finishBoxSelect(canvas,e);
        else doRaycastSelect(canvas,e);
      }
      _leftDown=false;_midDown=false;_rightDown=false;_boxSelecting=false;
      const b=document.getElementById('ifc-select-box');if(b)b.style.display='none';
    });
    canvas.addEventListener('mousemove',e=>{
      const dx=e.clientX-_lastX,dy=e.clientY-_lastY;
      if(_rightDown){orbitBy(dx,dy);updateCamera();}
      else if(_midDown){panBy(dx,dy);updateCamera();}
      else if(_leftDown){
        if(_leftMode==='orbit'){orbitBy(dx,dy);updateCamera();}
        else if(_leftMode==='pan'){panBy(dx,dy);updateCamera();}
        else{
          const ddx=e.clientX-_dragStartX,ddy=e.clientY-_dragStartY;
          if(!_boxSelecting&&Math.hypot(ddx,ddy)>BOX_DRAG_THRESHOLD)_boxSelecting=true;
          if(_boxSelecting)updateBoxOverlay(wrap,e);
        }
      }
      _lastX=e.clientX;_lastY=e.clientY;
    });
    canvas.addEventListener('wheel',e=>{_radius=Math.max(1,Math.min(5000,_radius*(1+e.deltaY*.001)));updateCamera();});
    canvas.addEventListener('contextmenu',e=>e.preventDefault());
    canvas.addEventListener('touchstart',e=>{if(e.touches.length===1){_touchDrag=true;_lastX=e.touches[0].clientX;_lastY=e.touches[0].clientY;}});
    canvas.addEventListener('touchend',()=>_touchDrag=false);
    canvas.addEventListener('touchmove',e=>{if(!_touchDrag||e.touches.length!==1)return;const t=e.touches[0],dx=t.clientX-_lastX,dy=t.clientY-_lastY;_lastX=t.clientX;_lastY=t.clientY;orbitBy(dx,dy);updateCamera();});
  }
  function updateCamera(){const th=_theta*Math.PI/180,ph=_phi*Math.PI/180;camera.position.set(_target.x+_radius*Math.sin(ph)*Math.sin(th),_target.y+_radius*Math.cos(ph),_target.z+_radius*Math.sin(ph)*Math.cos(th));camera.lookAt(_target);}
  function fitCamera(){if(!modelGroup)return;const box=new THREE.Box3().setFromObject(modelGroup);if(box.isEmpty())return;const center=new THREE.Vector3(),size=new THREE.Vector3();box.getCenter(center);box.getSize(size);_target.copy(center);_radius=Math.max(size.x,size.y,size.z)*1.8||80;_theta=45;_phi=55;
    // Near/far fixos (.01–20000) quebram a precisao do depth buffer em modelos
    // grandes (coordenadas de projeto reais, ex. dezenas de metros em cm) —
    // ajusta a escala da camera ao tamanho real do modelo carregado.
    if(camera){camera.near=Math.max(_radius/5000,0.01);camera.far=Math.max(_radius*20,1000);camera.updateProjectionMatrix();}
    updateCamera();}
  function onResize(){const wrap=document.getElementById('ifc-canvas-wrap');if(!wrap||!renderer||!camera)return;camera.aspect=wrap.clientWidth/wrap.clientHeight;camera.updateProjectionMatrix();renderer.setSize(wrap.clientWidth,wrap.clientHeight);}
  function animate(){requestAnimationFrame(animate);if(renderer&&scene&&camera)renderer.render(scene,camera);}
  function cameraPreset(p){if(p==='front'){_theta=0;_phi=70;}else if(p==='top'){_theta=0;_phi=5;}else if(p==='iso'){_theta=45;_phi=55;}else if(p==='reset'){fitCamera();return;}updateCamera();}

  // ── PUBLIC API ──────────────────────────────────────────────────────
  return {
    _meshMap:{},

    init(){
      if(initialized)return;initialized=true;loadStorage();
      const canvas=document.getElementById('ifc-canvas'),wrap=document.getElementById('ifc-canvas-wrap');
      if(!canvas||!wrap){setStatus('❌ Canvas não encontrado.');return;}
      if(!window.THREE){setStatus('❌ Three.js não carregado — verifique conexão com internet.');return;}
      scene=new THREE.Scene();scene.background=new THREE.Color(0x0a0808);scene.fog=new THREE.Fog(0x0a0808,500,3000);
      camera=new THREE.PerspectiveCamera(45,wrap.clientWidth/wrap.clientHeight,.01,20000);camera.position.set(60,40,60);
      renderer=new THREE.WebGLRenderer({canvas,antialias:true,preserveDrawingBuffer:true});renderer.setPixelRatio(Math.min(devicePixelRatio,2));renderer.setSize(wrap.clientWidth,wrap.clientHeight);renderer.shadowMap.enabled=true;renderer.shadowMap.type=THREE.PCFSoftShadowMap;
      scene.add(new THREE.AmbientLight(0xffeedd,.6));
      const sun=new THREE.DirectionalLight(0xfff5ee,1.2);sun.position.set(50,80,40);sun.castShadow=true;sun.shadow.mapSize.set(2048,2048);sun.shadow.camera.far=5000;scene.add(sun);
      const fillLight=new THREE.DirectionalLight(0xaaccff,.4);fillLight.position.set(-30,20,-20);scene.add(fillLight);
      raycaster=new THREE.Raycaster();
      setupControls(canvas);
      window.addEventListener('resize',onResize);
      animate();updateCamera();
      setStatus('Three.js OK — selecione ou arraste um arquivo .ifc');
    },

    loadFile(file){
      if(!file)return;
      if(!file.name.toLowerCase().endsWith('.ifc')){alert('Selecione um arquivo .ifc');return;}
      const nm=document.getElementById('ifc-file-name');if(nm)nm.textContent=file.name+' ('+(file.size/1048576).toFixed(1)+' MB)';
      document.getElementById('ifc-drop-zone')?.classList.add('hidden');
      setStatus('Iniciando leitura de '+file.name+'…',1);
      readFileInChunks(file,
        parsed=>{ setStatus('Construindo cena 3D…',93); setTimeout(()=>{try{buildScene(parsed);}catch(ex){setStatus('❌ Cena: '+ex.message);console.error(ex);}},50); },
        err=>{ setStatus('❌ '+err.message); console.error('[IFC ERROR]',err); }
      );
    },

    onDragOver(e){e.preventDefault();document.getElementById('ifc-drop-border')?.classList.add('drag-over');},
    onDragLeave(e){document.getElementById('ifc-drop-border')?.classList.remove('drag-over');},
    onDrop(e){e.preventDefault();document.getElementById('ifc-drop-border')?.classList.remove('drag-over');const f=e.dataTransfer?.files?.[0];if(f)this.loadFile(f);},

    selectEl,medir,desmarcar,toggleEmAndamento,toggleHide,hideSelected,unhideAll,setLeftMode,
    toggleOrcPanel,autoVincularPavimentos,togglePavPanel,mostrarTodosPavimentos,filterPavimentoByType,
    toggleQuantPanel,applySteelRates,
    toggleCompatPanel,pickCompatFile,loadCompatModel,removeCompatModel,toggleCompatModel,renameCompatModel,
    runClashDetection,focusClash,exportClashesCSV,exportMateriaisCSV,carregarModeloAutoDoc,
    salvarMedicao,toggleHistorico,excluirMedicaoSalva,
    // Layout BIM (painel alternativo) — reaproveita funcoes internas ja
    // existentes (pavimentos) e expoe os agregadores novos.
    selectPavimento,isolarPavimento,pavimentoGroups,
    propriedadesElementosTable,volumePorPavimentoETipo,computeQuantPorSecao,
    setExecFilter(val){execFilter=val;refreshColors();renderTree();},
    getExecFilter(){return execFilter;},
    filterTree(){renderTree();},
    view(f){viewFilter=f;document.querySelectorAll('.ifc-btn[onclick^="IFC.view"]').forEach(b=>b.classList.toggle('act',b.getAttribute('onclick').includes(`'${f}'`)));refreshColors();renderTree();},
    camera(p){cameraPreset(p);},
    toggleWire(){wireMode=!wireMode;Object.values(meshMap).forEach(e=>{if(e.mesh)e.mesh.material.wireframe=wireMode;});},
    togglePeriodMode(){
      periodMode=!periodMode;
      document.getElementById('ifc-period-btn')?.classList.toggle('act',periodMode);
      document.getElementById('ifc-legend-periodo')?.style.setProperty('display',periodMode?'flex':'none');
      document.getElementById('ifc-legend-status')?.style.setProperty('display',periodMode?'none':'flex');
      refreshColors();
      setStatus(periodMode?'📅 Colorindo por período: medido no mês atual x meses anteriores.':'Cor por status restaurada.');
    },
    toggleSidebar(){sidebarOpen=!sidebarOpen;document.getElementById('ifc-sidebar')?.classList.toggle('collapsed',!sidebarOpen);setTimeout(onResize,300);},
    exportCSV(){const rows=[['GlobalId','Nome','Tipo','Nível','Volume concreto (m³)','Área de forma (m²)','Peso de aço est. (kg)','% Medido','Ref ORC']];for(const gid in meshMap){const e=meshMap[gid],d=e.data;rows.push([gid,d.name,d.type,d.level,(d.volume||0).toFixed(4),(d.formworkArea||0).toFixed(3),(d.steelWeight||0).toFixed(1),e.pct>=0?e.pct:'',e.orcRef||'']);}const csv=rows.map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');const a=document.createElement('a');a.href='data:text/csv;charset=utf-8,﻿'+encodeURIComponent(csv);a.download='medicao_ifc_'+new Date().toISOString().slice(0,10)+'.csv';a.click();},
    exportFoto(){
      if(!renderer||!scene||!camera){setStatus('❌ Carregue um modelo primeiro.');return;}
      const wrap=document.getElementById('ifc-canvas-wrap'); if(!wrap)return;
      const cssW=wrap.clientWidth,cssH=wrap.clientHeight;
      const origPR=renderer.getPixelRatio();
      // Sobe a resolucao do buffer de desenho pra tirar a foto na maior
      // qualidade possivel (o canvas visivel na tela nao muda de tamanho —
      // so o que e capturado fica em resolucao maior), depois volta ao
      // normal pra nao atrapalhar a navegacao.
      const hiRes=Math.min(4,Math.max(origPR*2,2));
      renderer.setPixelRatio(hiRes);renderer.setSize(cssW,cssH,false);
      renderer.render(scene,camera);
      const dataUrl=renderer.domElement.toDataURL('image/jpeg',0.98);
      renderer.setPixelRatio(origPR);renderer.setSize(cssW,cssH,false);renderer.render(scene,camera);
      const a=document.createElement('a');a.href=dataUrl;a.download='modelo_ifc_'+new Date().toISOString().slice(0,19).replace(/[:T]/g,'-')+'.jpg';a.click();
      setStatus('📷 Foto do modelo exportada em alta qualidade.');
    },
    reset(){if(!confirm('Resetar todas as medições?'))return;medStorage={};saveStorage();for(const gid in meshMap){meshMap[gid].pct=-1;meshMap[gid].orcRef='';}refreshColors();renderTree();updateMedCount();setStatus('Medições resetadas.');},
  };
})();


// ═══════════════════════════════════════════════════════════════════════════
// FILTRO DE PERÍODO — PAGAMENTOS  (auto-inject, sem chamada manual)
// Detecta a seção de pagamentos pelo texto "FLUXO DE PAGAMENTOS" ou pelos ids
// dos cards, e injeta o filtro automaticamente na posição certa.
// ═══════════════════════════════════════════════════════════════════════════
const FiltroPeriodo = (() => {
  const MESES_PT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                    'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const MESES_SH = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  const brl = v => 'R$ ' + parseFloat(v||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
  const $ = id => document.getElementById(id);
  let _injected = false;

  // ── Estilos CSS injetados uma única vez ─────────────────────────────
  function injectCSS() {
    if ($('fp-styles')) return;
    const s = document.createElement('style');
    s.id = 'fp-styles';
    s.textContent = `
      #fp-wrap{margin:0 0 18px 0;font-family:inherit}
      #fp-wrap .fp-label{font-size:10px;letter-spacing:.1em;text-transform:uppercase;
        color:var(--muted,#888);margin-bottom:8px;display:flex;align-items:center;gap:8px}
      #fp-badge{display:none;background:#fdeaec;color:#b81c2e;border:1px solid #f5b8bf;
        border-radius:20px;padding:2px 10px;font-size:11px;cursor:pointer;user-select:none}
      #fp-box{border:1px solid var(--border,rgba(255,255,255,.12));border-radius:10px;
        padding:14px 16px;background:var(--card-bg,rgba(255,255,255,.04))}
      #fp-atalhos{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px}
      .fp-at{font-size:12px;background:rgba(181,28,46,.12);color:#c9404f;
        border:1px solid rgba(181,28,46,.3);border-radius:20px;padding:3px 11px;
        cursor:pointer;white-space:nowrap;transition:background .15s,color .15s}
      .fp-at:hover,.fp-at.act{background:#b81c2e;color:#fff;border-color:#b81c2e}
      #fp-custom{display:flex;align-items:flex-end;flex-wrap:wrap;gap:10px}
      .fp-grp{display:flex;flex-direction:column;gap:3px}
      .fp-grp label{font-size:11px;color:var(--muted,#888)}
      .fp-grp .fp-row{display:flex;gap:4px}
      .fp-sel{font-size:13px;padding:7px 10px;border-radius:6px;
        border:1px solid var(--border,rgba(255,255,255,.18));
        background:var(--input-bg,rgba(255,255,255,.08));
        color:var(--text,#e8e8e8);cursor:pointer;outline:none;font-family:inherit}
      .fp-sel:focus{border-color:#b81c2e}
      .fp-arr{font-size:18px;color:var(--muted,#666);padding-bottom:4px;align-self:flex-end}
      .fp-btn-apply{background:#b81c2e;color:#fff;border:none;border-radius:6px;
        padding:8px 20px;font-size:13px;cursor:pointer;font-family:inherit;
        display:flex;align-items:center;gap:6px;align-self:flex-end}
      .fp-btn-apply:hover{background:#9a1626}
      .fp-btn-clear{background:transparent;border:1px solid var(--border,rgba(255,255,255,.2));
        border-radius:6px;padding:8px 14px;font-size:13px;cursor:pointer;
        color:var(--muted,#aaa);font-family:inherit;align-self:flex-end}
      .fp-btn-clear:hover{background:rgba(255,255,255,.06)}
      #fp-resultado{display:none;margin-top:12px}
      .fp-res-label{font-size:10px;letter-spacing:.1em;text-transform:uppercase;
        color:var(--muted,#888);margin-bottom:8px}
      .fp-cards{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:8px}
      .fp-card{flex:1;min-width:150px;border:1px solid var(--border,rgba(255,255,255,.1));
        border-radius:8px;padding:10px 14px;background:var(--card-bg,rgba(255,255,255,.04))}
      .fp-card-lbl{font-size:10px;letter-spacing:.06em;color:var(--muted,#888);margin-bottom:4px;text-transform:uppercase}
      .fp-card-val{font-size:18px;font-weight:600}
      .fp-note{font-size:11px;color:var(--muted,#888)}
      .fp-note a{color:#b81c2e;text-decoration:none}
      .fp-divider{border:none;border-top:1px solid var(--border,rgba(255,255,255,.1));margin:12px 0}
    `;
    document.head.appendChild(s);
  }

  // ── Encontra o container correto da aba pagamentos ──────────────────
  function findContainer() {
    // 1. Pelos ids dos cards conhecidos
    const cardIds = ['pag-entrada','pag-saida','pag-saldo','card-desembolsado','card-apagar','card-apropriado'];
    for (const id of cardIds) {
      const el = $(id);
      if (el) {
        // sobe até um container com pelo menos 300px de largura
        let p = el.parentElement;
        while (p && p !== document.body) {
          if (p.offsetWidth > 300) return p;
          p = p.parentElement;
        }
      }
    }
    // 2. Pelo texto "FLUXO DE PAGAMENTOS"
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = walker.nextNode())) {
      if (/FLUXO DE PAGAMENTOS/i.test(node.textContent)) {
        let p = node.parentElement;
        while (p && p !== document.body) {
          if (p.offsetWidth > 300) return p.parentElement || p;
          p = p.parentElement;
        }
      }
    }
    // 3. Por ids comuns de seção
    const sectionIds = ['pagamentos-content','tab-pagamentos','pag-section','pagamentos'];
    for (const id of sectionIds) {
      const el = $(id); if (el) return el;
    }
    return null;
  }

  // ── Encontra o melhor ponto de inserção (antes do primeiro card/chart) ──
  function findInsertionPoint(container) {
    // Procura o elemento que contém "FLUXO DE PAGAMENTOS"
    const all = container.querySelectorAll('*');
    for (const el of all) {
      if (/FLUXO DE PAGAMENTOS/i.test(el.textContent) && el.children.length === 0) {
        return el.closest('div, section, p') || el.parentElement;
      }
    }
    return null;
  }

  // ── Constrói e injeta o HTML ──────────────────────────────────────
  function inject(container) {
    if (_injected && $('fp-wrap')) return;
    _injected = true;
    injectCSS();

    const wrap = document.createElement('div');
    wrap.id = 'fp-wrap';
    wrap.innerHTML = `
      <div class="fp-label">
        // filtro de período
        <span id="fp-badge" onclick="FiltroPeriodo.limpar()">
          📅 <span id="fp-badge-txt"></span> &times;
        </span>
      </div>
      <div id="fp-box">
        <div id="fp-atalhos"></div>
        <hr class="fp-divider">
        <div style="font-size:11px;color:var(--muted,#888);margin-bottom:10px">Intervalo personalizado</div>
        <div id="fp-custom">
          <div class="fp-grp">
            <label>De (mês)</label>
            <div class="fp-row">
              <select id="fp-de-mes" class="fp-sel"></select>
              <select id="fp-de-ano" class="fp-sel"></select>
            </div>
          </div>
          <div class="fp-arr">→</div>
          <div class="fp-grp">
            <label>Até (mês)</label>
            <div class="fp-row">
              <select id="fp-ate-mes" class="fp-sel"></select>
              <select id="fp-ate-ano" class="fp-sel"></select>
            </div>
          </div>
          <button class="fp-btn-apply" onclick="FiltroPeriodo.aplicar()">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>
            Aplicar
          </button>
          <button class="fp-btn-clear" onclick="FiltroPeriodo.limpar()">Limpar</button>
        </div>
      </div>
      <div id="fp-resultado">
        <div class="fp-res-label">// resultado do período filtrado</div>
        <div class="fp-cards">
          <div class="fp-card"><div class="fp-card-lbl">Entradas (banco)</div><div id="fp-r-entrada" class="fp-card-val" style="color:#1d9e75">—</div></div>
          <div class="fp-card"><div class="fp-card-lbl">Saídas (banco)</div><div id="fp-r-saida" class="fp-card-val" style="color:#b81c2e">—</div></div>
          <div class="fp-card"><div class="fp-card-lbl">Saldo</div><div id="fp-r-saldo" class="fp-card-val">—</div></div>
          <div class="fp-card"><div class="fp-card-lbl">Desembolsado (AP)</div><div id="fp-r-desemb" class="fp-card-val" style="color:#b81c2e">—</div></div>
          <div class="fp-card"><div class="fp-card-lbl">Títulos</div><div id="fp-r-tits" class="fp-card-val">—</div></div>
        </div>
        <div class="fp-note">⚠️ Gráfico e tabelas filtrados para o período.
          <a href="#" onclick="FiltroPeriodo.limpar();return false">Limpar filtro</a>
        </div>
      </div>`;

    // Insere antes do primeiro filho do container
    const ref = findInsertionPoint(container);
    if (ref && ref.parentElement === container) {
      container.insertBefore(wrap, ref);
    } else {
      container.insertBefore(wrap, container.firstChild);
    }

    popularSelects();
    gerarAtalhos();
  }

  // ── Selects ──────────────────────────────────────────────────────
  function popularSelects() {
    const now = new Date(), mm = String(now.getMonth()+1).padStart(2,'0'), aa = now.getFullYear();
    ['fp-de-mes','fp-ate-mes'].forEach(id => {
      const sel = $(id); if (!sel) return; sel.innerHTML = '';
      MESES_PT.forEach((nm, i) => {
        const o = document.createElement('option');
        o.value = String(i+1).padStart(2,'0'); o.textContent = nm;
        if (o.value === mm) o.selected = true;
        sel.appendChild(o);
      });
    });
    ['fp-de-ano','fp-ate-ano'].forEach(id => {
      const sel = $(id); if (!sel) return; sel.innerHTML = '';
      for (let a = 2022; a <= aa; a++) {
        const o = document.createElement('option');
        o.value = a; o.textContent = a;
        if (a === aa) o.selected = true;
        sel.appendChild(o);
      }
    });
  }

  // ── Atalhos rápidos ──────────────────────────────────────────────
  function gerarAtalhos() {
    const wrap = $('fp-atalhos'); if (!wrap) return;
    const now = new Date(), aa = now.getFullYear(), mm = now.getMonth();
    const ats = [];
    for (let i = 0; i < 5; i++) {
      const d = new Date(aa, mm - i, 1);
      const m = String(d.getMonth()+1).padStart(2,'0'), a = d.getFullYear();
      ats.push({lbl: MESES_SH[d.getMonth()]+'/'+String(a).slice(2), de:`${a}-${m}`, ate:`${a}-${m}`});
    }
    ats.push({lbl:`${aa} (jan–${MESES_SH[mm]})`, de:`${aa}-01`, ate:`${aa}-${String(mm+1).padStart(2,'0')}`});
    ats.push({lbl:`${aa-1} completo`, de:`${aa-1}-01`, ate:`${aa-1}-12`});
    wrap.innerHTML = '';
    ats.forEach(at => {
      const btn = document.createElement('span');
      btn.className = 'fp-at'; btn.textContent = at.lbl;
      btn.onclick = () => {
        wrap.querySelectorAll('.fp-at').forEach(b => b.classList.remove('act'));
        btn.classList.add('act');
        aplicarPeriodo(at.de, at.ate);
      };
      wrap.appendChild(btn);
    });
  }

  // ── Aplicar ──────────────────────────────────────────────────────
  function aplicarPeriodo(de, ate) {
    // sincroniza selects
    if (de)  { const [a,m]=de.split('-');  if($('fp-de-mes'))  $('fp-de-mes').value=m;  if($('fp-de-ano'))  $('fp-de-ano').value=a; }
    if (ate) { const [a,m]=ate.split('-'); if($('fp-ate-mes')) $('fp-ate-mes').value=m; if($('fp-ate-ano')) $('fp-ate-ano').value=a; }
    // badge
    const [dm,da]=de.split('-'), [am,aa]=ate.split('-');
    const lbl = de===ate
      ? MESES_SH[parseInt(dm)-1]+'/'+da.slice(2)
      : MESES_SH[parseInt(dm)-1]+'/'+da.slice(2)+' – '+MESES_SH[parseInt(am)-1]+'/'+aa.slice(2);
    if ($('fp-badge-txt')) $('fp-badge-txt').textContent = lbl;
    if ($('fp-badge'))     $('fp-badge').style.display   = 'inline';
    // API
    fetch(`/api/pagamentos_filtro?de=${de}&ate=${ate}`)
      .then(r => r.json())
      .then(res => {
        if (!res.ok) { console.error('[FiltroPeriodo]', res.error); return; }
        const s = res.summary;
        if ($('fp-r-entrada'))  $('fp-r-entrada').textContent  = brl(s.entradas_banco);
        if ($('fp-r-saida'))    $('fp-r-saida').textContent    = brl(s.pagamentos_banco);
        if ($('fp-r-desemb'))   $('fp-r-desemb').textContent   = brl(s.vlr_desembolsado);
        if ($('fp-r-tits'))     $('fp-r-tits').textContent     = s.qt_titulos;
        if ($('fp-r-saldo')) {
          $('fp-r-saldo').textContent = brl(s.saldo_banco);
          $('fp-r-saldo').style.color = s.saldo_banco >= 0 ? '#1d9e75' : '#b81c2e';
        }
        if ($('fp-resultado')) $('fp-resultado').style.display = 'block';
        atualizarGrafico(res.fluxo);
        atualizarCredores(res.by_cred);
        document.dispatchEvent(new CustomEvent('filtro-periodo-aplicado', {detail: res}));
      })
      .catch(e => console.error('[FiltroPeriodo]', e));
  }

  // ── Limpar ──────────────────────────────────────────────────────
  function limpar() {
    if ($('fp-badge'))     $('fp-badge').style.display     = 'none';
    if ($('fp-resultado')) $('fp-resultado').style.display = 'none';
    document.querySelectorAll('#fp-atalhos .fp-at').forEach(b => b.classList.remove('act'));
    if (typeof DASH !== 'undefined') {
      if (DASH.fluxo)   atualizarGrafico(DASH.fluxo);
      if (DASH.by_cred) atualizarCredores(DASH.by_cred);
    }
    document.dispatchEvent(new CustomEvent('filtro-periodo-limpo'));
  }

  // ── Atualizar gráfico Chart.js ──────────────────────────────────────
  function atualizarGrafico(fluxo) {
    let chart = null;
    for (const id of ['fluxo-chart','chart-fluxo','pagamentos-chart','chart-pagamentos','monthly-chart']) {
      const el = $(id);
      if (el && window.Chart && Chart.getChart) { chart = Chart.getChart(el); if (chart) break; }
    }
    if (!chart && window.chartFluxoPag) chart = window.chartFluxoPag;
    if (!chart) {
      // tenta qualquer canvas dentro da seção de pagamentos
      if (window.Chart && Chart.getChart) {
        document.querySelectorAll('canvas').forEach(c => { if (!chart) chart = Chart.getChart(c); });
      }
    }
    if (!chart) return;
    chart.data.labels = fluxo.map(f => f.mes);
    chart.data.datasets[0].data = fluxo.map(f => f.v);
    chart.update();
  }

  // ── Atualizar tabela de credores ──────────────────────────────────────
  function atualizarCredores(creds) {
    const tbody = document.querySelector('#cred-tbody, #tabela-credores tbody, .cred-table tbody, [id*="cred"] tbody');
    if (!tbody) return;
    tbody.innerHTML = '';
    creds.forEach(c => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${c.nmcredor}</td><td>${brl(c.vlr_ap)}</td><td>${brl(c.vlr_pg)}</td>`+
        `<td>${brl(c.vlr_ap-c.vlr_pg)}</td><td>${brl(c.vlrjur)}</td><td>${c.qt}</td>`;
      tbody.appendChild(tr);
    });
  }

  // ── Auto-init via MutationObserver ──────────────────────────────────────
  // Detecta quando a aba de pagamentos é exibida e injeta o filtro.
  function autoInit() {
    function tryInject() {
      if ($('fp-wrap')) return true; // já injetado
      const container = findContainer();
      if (container && container.offsetParent !== null) { // visível
        inject(container);
        return true;
      }
      return false;
    }
    // Tenta imediatamente (caso aba já esteja ativa)
    if (tryInject()) return;
    // MutationObserver para capturar quando a aba for exibida
    const obs = new MutationObserver(() => { if (tryInject()) obs.disconnect(); });
    obs.observe(document.body, {childList: true, subtree: true, attributes: true, attributeFilter: ['style','class','hidden']});
    // Fallback: polling por 10 s
    let attempts = 0;
    const poll = setInterval(() => {
      if (tryInject() || ++attempts > 100) clearInterval(poll);
    }, 100);
  }

  document.addEventListener('DOMContentLoaded', autoInit);
  if (document.readyState !== 'loading') autoInit();

  // ── API pública ──────────────────────────────────────────────────────
  return {
    init(containerId) {
      const c = containerId ? $(containerId) : findContainer();
      if (c) inject(c);
      else console.warn('[FiltroPeriodo] container não encontrado');
    },
    aplicar() {
      const de  = `${$('fp-de-ano').value}-${$('fp-de-mes').value}`;
      const ate = `${$('fp-ate-ano').value}-${$('fp-ate-mes').value}`;
      aplicarPeriodo(de, ate);
    },
    limpar,
  };
})();
// IFC e declarado com "const" no escopo global deste script classico —
// isso o torna acessivel como identificador global (IFC), mas NAO como
// window.IFC. Outros scripts (ifcviz.js, o script inline de index.html)
// dependem de checar window.IFC/window.IFCBim de forma defensiva (protege
// contra ordem de carregamento), entao precisa ser exposto explicitamente.
window.IFC = IFC;
