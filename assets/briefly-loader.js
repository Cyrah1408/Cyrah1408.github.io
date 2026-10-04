/* Briefly data loader.
   Reads data/briefly.xlsx (no third-party library: a small .xlsx reader using the browser's
   built-in DecompressionStream), turns its sheets into the edition objects the app expects,
   then starts the app. Every few minutes it checks data/version.json (a tiny file) and only
   downloads the spreadsheet again when that changes.
   The sheet/column schema must match tools/briefly_data.py. */
(function(){
  "use strict";
  const XLSX_URL = "data/briefly.xlsx", VERSION_URL = "data/version.json";
  const POLL_MS = 5 * 60 * 1000;   // how often to look for a new version
  const SCHEMA = 1;

  /* ---------- minimal .xlsx reader ---------- */
  async function inflateRaw(bytes){
    const ds = new DecompressionStream("deflate-raw");
    const out = new Response(new Blob([bytes]).stream().pipeThrough(ds));
    return new Uint8Array(await out.arrayBuffer());
  }
  async function unzip(buf){
    const u8 = new Uint8Array(buf), dv = new DataView(buf), files = {};
    let eocd = -1;
    for(let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--){ if(dv.getUint32(i, true) === 0x06054b50){ eocd = i; break; } }
    if(eocd < 0) throw new Error("not a valid .xlsx file");
    const n = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const dec = new TextDecoder();
    for(let k = 0; k < n; k++){
      if(dv.getUint32(p, true) !== 0x02014b50) throw new Error("bad zip directory");
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nl = dv.getUint16(p + 28, true), xl = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true), name = dec.decode(u8.subarray(p + 46, p + 46 + nl));
      files[name] = {method, csize, off};
      p += 46 + nl + xl + cl;
    }
    return {
      async text(name){
        const f = files[name]; if(!f) return null;
        const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
        const raw = u8.subarray(start, start + f.csize);
        const data = f.method === 0 ? raw : f.method === 8 ? await inflateRaw(raw) : null;
        if(!data) throw new Error("unsupported compression in " + name);
        return dec.decode(data);
      }
    };
  }
  const xml = t => new DOMParser().parseFromString(t, "application/xml");
  const tags = (node, name) => Array.from(node.getElementsByTagNameNS("*", name));
  const colIdx = ref => { let n = 0; for(const ch of ref.replace(/\d+/g, "")) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };

  async function readWorkbook(buf){
    const z = await unzip(buf);
    const wb = xml(await z.text("xl/workbook.xml")), rels = xml(await z.text("xl/_rels/workbook.xml.rels"));
    const target = {};
    tags(rels, "Relationship").forEach(r => { const t = r.getAttribute("Target"); target[r.getAttribute("Id")] = t.startsWith("/") ? t.slice(1) : "xl/" + t; });
    const sstText = await z.text("xl/sharedStrings.xml");
    const sst = sstText ? tags(xml(sstText), "si").map(si => tags(si, "t").map(t => t.textContent).join("")) : [];
    const sheets = {};
    for(const sh of tags(wb, "sheet")){
      const rid = sh.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || sh.getAttribute("r:id");
      const doc = xml(await z.text(target[rid]));
      const rows = tags(doc, "row").map(row => {
        const out = [];
        tags(row, "c").forEach(c => {
          const t = c.getAttribute("t"), v = tags(c, "v")[0];
          let val = null;
          if(t === "s") val = sst[+v.textContent];
          else if(t === "inlineStr") val = tags(c, "t").map(x => x.textContent).join("");
          else if(t === "str") val = v ? v.textContent : "";
          else if(t === "b") val = v ? v.textContent === "1" : null;
          else if(v) val = Number(v.textContent);
          out[colIdx(c.getAttribute("r"))] = val;
        });
        return out;
      });
      const head = (rows[0] || []).map(h => String(h ?? ""));
      sheets[sh.getAttribute("name")] = rows.slice(1)
        .filter(r => r.some(v => v !== null && v !== undefined && v !== ""))
        .map(r => Object.fromEntries(head.map((h, i) => [h, r[i] ?? null])));
    }
    return sheets;
  }

  /* ---------- sheets -> editions (mirror of from_rows() in tools/briefly_data.py) ---------- */
  const s = v => v === null || v === undefined ? "" : String(v);
  const num = v => v === null || v === undefined || v === "" ? null : Number(v);
  const split = v => s(v).trim() ? s(v).split("|").map(x => x.trim()) : [];
  function by(rows, key){
    const out = {};
    (rows || []).forEach(r => (out[s(r[key])] = out[s(r[key])] || []).push(r));
    Object.values(out).forEach(v => v.sort((a, b) => (num(a.seq) || 0) - (num(b.seq) || 0)));
    return out;
  }
  function buildEditions(S){
    const need = ["editions", "stories", "sources", "pairs", "opinions", "games", "connections", "timeline", "duel", "audio"];
    need.forEach(k => { if(!S[k]) throw new Error("sheet '" + k + "' is missing"); });
    const stories = by(S.stories, "edition_date"), sources = by(S.sources, "item_id"), pairs = by(S.pairs, "parent_id");
    const opinions = by(S.opinions, "edition_date"), games = by(S.games, "edition_date"), conns = by(S.connections, "edition_date");
    const tls = by(S.timeline, "edition_date"), duels = by(S.duel, "edition_date"), audio = by(S.audio, "edition_date");
    const srcs = id => (sources[id] || []).map(x => [s(x.outlet), s(x.url)]);
    const prs = id => (pairs[id] || []).map(x => [s(x.term), s(x.definition)]);
    const eds = S.editions.map(r => {
      const d = s(r.date), e = {date:d, no:num(r.no), label:s(r.label)};
      const au = {};
      (audio[d] || []).forEach(a => { const k = s(a.track); (au[k] = au[k] || {lines:[]}).lines.push([s(a.speaker), s(a.text), s(a.tone)]); });
      if(Object.keys(au).length) e.audio = au;
      e.bulletin = {greeting:s(r.greeting), paras:["para_1", "para_2", "para_3"].map(k => s(r[k])).filter(Boolean)};
      const sec = {brief:[], curated:[], desk:[]};
      (stories[d] || []).forEach(x => {
        const id = s(x.id), kind = s(x.section);
        let it;
        if(kind === "brief") it = {id, theme:s(x.theme), stat:{big:s(x.stat_big), label:s(x.stat_label)}, head:s(x.head), summary:s(x.summary), why:s(x.why), cls:s(x.cls), words:prs(id), sources:srcs(id)};
        else if(kind === "curated") it = {id, theme:s(x.theme), minutes:num(x.minutes), head:s(x.head), dek:s(x.dek),
          body:[1, 2, 3, 4].map(i => s(x["body_" + i])).filter(Boolean), game:{type:s(x.game_type), word:s(x.game_word), hint:s(x.game_hint)},
          note:s(x.note), checks:split(x.checks), think:split(x.think), sources:srcs(id)};
        else { it = {id, theme:s(x.theme), head:s(x.head), summary:s(x.summary), why:s(x.why)};
          if(s(x.stat_big) || s(x.stat_label)) it.stat = {big:s(x.stat_big), label:s(x.stat_label)};
          it.sources = srcs(id); }
        (sec[kind] = sec[kind] || []).push(it);
      });
      e.brief = sec.brief;
      e.games = (games[d] || []).map(g => { const gi = {id:s(g.id), type:s(g.type)};
        if(s(g.word) || s(g.hint)){ gi.word = s(g.word); gi.hint = s(g.hint); }
        if((pairs[gi.id] || []).length) gi.pairs = prs(gi.id);
        return gi; });
      e.connections = {groups:(conns[d] || []).map(c => ({name:s(c.name), words:split(c.words)}))};
      e.hive = {letters:s(r.hive_letters), centre:s(r.hive_centre), words:split(r.hive_words)};
      e.timeline = {prompt:s(r.timeline_prompt), events:(tls[d] || []).map(t => ({t:s(t.event), y:num(t.year)}))};
      e.duel = (duels[d] || []).map(x => ({a:[s(x.a_label), num(x.a_value)], b:[s(x.b_label), num(x.b_value)], note:s(x.note)}));
      e.debate = {motion:s(r.debate_motion), for:["debate_for_1", "debate_for_2"].map(k => s(r[k])).filter(Boolean),
        against:["debate_against_1", "debate_against_2"].map(k => s(r[k])).filter(Boolean), opinion:s(r.debate_opinion)};
      e.opinions = (opinions[d] || []).map(o => ({id:s(o.id), theme:s(o.theme), outlet:s(o.outlet), author:s(o.author), role:s(o.role),
        date:s(o.date), head:s(o.head), url:s(o.url), gist:s(o.gist), other:s(o.other), think:split(o.think)}));
      e.curated = sec.curated;
      if(sec.desk.length) e.desk = sec.desk;
      return e;
    });
    eds.sort((a, b) => a.date < b.date ? 1 : a.date > b.date ? -1 : 0);
    if(!eds.length) throw new Error("the spreadsheet has no editions");
    return {editions:eds};
  }

  /* ---------- boot + polling ---------- */
  async function getVersion(){
    const r = await fetch(VERSION_URL + "?t=" + Date.now(), {cache:"no-store"});
    if(!r.ok) throw new Error("version.json " + r.status);
    return r.json();
  }
  async function getData(version){
    const r = await fetch(XLSX_URL + "?v=" + encodeURIComponent(version), {cache:"no-cache"});
    if(!r.ok) throw new Error("briefly.xlsx " + r.status);
    return buildEditions(await readWorkbook(await r.arrayBuffer()));
  }
  function startApp(){
    const code = document.getElementById("briefly-app").textContent;
    const el = document.createElement("script");
    el.textContent = code;
    document.body.appendChild(el);
  }
  function status(msg, retry){
    const app = document.getElementById("app");
    app.innerHTML = `<p class="empty" role="status">${msg}</p>` + (retry ? `<p style="margin-top:12px"><button class="btn sm" onclick="location.reload()">Try again</button></p>` : "");
  }
  function notice(){
    if(document.getElementById("briefly-update")) return;
    const b = document.createElement("div");
    b.id = "briefly-update"; b.setAttribute("role", "status");
    b.style.cssText = "position:fixed;left:50%;transform:translateX(-50%);top:calc(12px + env(safe-area-inset-top,0px));z-index:50;background:var(--ink);color:var(--oat);border-radius:999px;padding:8px 10px 8px 18px;display:flex;gap:12px;align-items:center;font-size:14.5px;font-weight:600;box-shadow:var(--sh-lg);max-width:calc(100% - 24px)";
    b.innerHTML = `<span>A new Briefly update is ready.</span><button class="btn alt sm" type="button">Load it</button>`;
    b.querySelector("button").onclick = () => location.reload();
    document.body.appendChild(b);
  }

  let loaded = null;
  async function poll(){
    try{
      const v = await getVersion();
      if(loaded && v.version !== loaded){
        loaded = v.version;
        if(document.hidden) location.reload();   // nobody is looking: refresh quietly
        else notice();                           // someone is reading: let them choose
      }
    }catch(e){ /* offline or a brief GitHub hiccup: try again next time */ }
  }

  window.BrieflyLoader = {readWorkbook, buildEditions, poll: () => poll()};   // exposed for testing
  boot();
  async function boot(){
    if(!document.getElementById("briefly-app")) return;       // test harness
    status("Loading today’s edition…");
    try{
      if(typeof DecompressionStream === "undefined") throw new Error("This browser is too old to open Briefly. Please update it.");
      let v;
      try{ v = await getVersion(); }catch(e){ v = {version:String(Date.now())}; }
      if(v.schema && v.schema > SCHEMA) throw new Error("schema " + v.schema + " is newer than this page");
      window.BRIEFLY_DATA = await getData(v.version);
      loaded = v.version;
      startApp();
      setInterval(poll, POLL_MS);
      document.addEventListener("visibilitychange", () => { if(!document.hidden) poll(); });
    }catch(e){
      console.error(e);
      status("Sorry, today’s edition could not be loaded. " + (e && e.message && /browser/.test(e.message) ? e.message : "Please check your connection and try again."), true);
    }
  }
})();
