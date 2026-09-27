/*
 * 開発用のにせ Apps Script（ブラウザだけで動かすためのもの）
 * SpreadsheetApp などを最小限まねして、データはブラウザの localStorage に保存します。
 * tools/preview.mjs が作る dev/preview.html だけで使います。Apps Script には入れません。
 */
(function(){
  const KEY = "salon-mock-gas";
  let db = {};
  try { db = JSON.parse(localStorage.getItem(KEY) || "{}"); } catch(e){}
  const save = () => { try { localStorage.setItem(KEY, JSON.stringify(db, (k,v) => v)); } catch(e){} };
  const revive = v => (typeof v === "string" && /^__date:/.test(v)) ? new Date(Number(v.slice(7))) : v;
  // 本物のスプレッドシートと同じく、先頭の ' は「文字として保存する」印なので取り除いて保存する
  const pack = v => (v instanceof Date) ? "__date:" + v.getTime() : (typeof v === "string" && v[0] === "'") ? v.slice(1) : v;

  function colIndex(letters){ let n = 0; for (const ch of letters) n = n*26 + (ch.charCodeAt(0)-64); return n; }

  function makeSheet(name){
    const rows = () => db[name];
    const sheet = {
      getName: () => name,
      getLastRow: () => rows().length,
      setFrozenRows(){},
      appendRow(arr){ rows().push(arr.map(pack)); save(); return sheet; },
      getRange(r, c, nr = 1, nc = 1){
        if (typeof r === "string"){ // "B:C" のような列指定だけ対応
          const [a, b] = r.split(":"); c = colIndex(a); nc = colIndex(b) - c + 1; r = 1; nr = Math.max(1, rows().length);
        }
        return makeRange(rows, r, c, nr, nc);
      },
    };
    return sheet;
  }
  function makeRange(rows, r, c, nr, nc){
    const range = {
      getRow: () => r,
      getValues(){
        const out = [];
        for (let i=0;i<nr;i++){ const row = rows()[r-1+i] || []; out.push(Array.from({length:nc}, (_,j) => revive(row[c-1+j] ?? ""))); }
        return out;
      },
      getValue(){ return range.getValues()[0][0]; },
      setValues(vals){
        vals.forEach((v,i) => { const ri = r-1+i; while (rows().length <= ri) rows().push([]); v.forEach((x,j) => rows()[ri][c-1+j] = pack(x)); });
        save(); return range;
      },
      setValue(v){ return range.setValues([[v]]); },
      clearContent(){ for (let i=0;i<nr;i++){ const row = rows()[r-1+i]; if (row) for (let j=0;j<nc;j++) row[c-1+j] = ""; } save(); return range; },
      setNumberFormat(){ return range; },
      setFontWeight(){ return range; },
      createTextFinder(text){
        return { matchEntireCell(){ return this; }, findNext(){
          const vals = range.getValues();
          for (let i=0;i<vals.length;i++) for (let j=0;j<vals[i].length;j++)
            if (String(vals[i][j]) === String(text)) return makeRange(rows, r+i, c+j, 1, 1);
          return null;
        }};
      },
    };
    return range;
  }

  const ss = {
    getSheetByName: n => db[n] ? makeSheet(n) : null,
    insertSheet: n => { db[n] = db[n] || []; save(); return makeSheet(n); },
  };
  window.SpreadsheetApp = { getActive: () => ss, flush(){} };
  window.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock(){} }) };
  const props = () => (db.__props = db.__props || {});
  window.PropertiesService = { getScriptProperties: () => ({
    getProperty: k => props()[k] ?? null,
    setProperty: (k, v) => { props()[k] = String(v); save(); },
  })};
  window.Session = { getScriptTimeZone: () => "Asia/Tokyo" };
  const p2 = n => String(n).padStart(2, "0");
  window.Utilities = {
    getUuid: () => crypto.randomUUID(),
    formatDate(d, tz, fmt){
      return fmt.replace("yyyy", d.getFullYear()).replace("MM", p2(d.getMonth()+1)).replace("dd", p2(d.getDate()))
                .replace("HH", p2(d.getHours())).replace("mm", p2(d.getMinutes()));
    },
  };
  window.HtmlService = {};

  // google.script.run のまね：少し遅らせて非同期に呼び、結果は JSON にして返す（本物も Date は返せない）
  function runner(ok, ng){
    return new Proxy({}, { get(_, prop){
      if (prop === "withSuccessHandler") return fn => runner(fn, ng);
      if (prop === "withFailureHandler") return fn => runner(ok, fn);
      return (...args) => setTimeout(() => {
        try {
          const fn = window.__gas && window.__gas[prop];
          if (!fn) throw new Error("サーバー関数 " + String(prop) + " がありません");
          const res = fn(...JSON.parse(JSON.stringify(args)));
          ok && ok(res === undefined ? undefined : JSON.parse(JSON.stringify(res)));
        } catch(e){ console.error(e); ng && ng(e); }
      }, 250);
    }});
  }
  window.google = { script: { get run(){ return runner(null, null); } } };
  window.__resetMock = () => { localStorage.removeItem(KEY); location.reload(); };
  console.info("にせ Apps Script で動いています。データを消すにはコンソールで __resetMock() を実行してください。");
})();
