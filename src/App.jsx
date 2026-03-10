import { useState, useRef, useEffect } from "react";
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip, BarChart, Bar, XAxis, YAxis, CartesianGrid, Legend } from "recharts";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────
// 定数・純粋関数（コンポーネント外）
// ─────────────────────────────────────────────
const BUDGET_COLORS = ["#b8872a","#3a72a8","#4a8c62","#7a50a8","#a05a3a","#8a8030"];
const INITIAL_YEARS = {
  "2025": {
    budgets: [
      { key:"個人研究費",    label:"個人研究費",    amount:400000,  color:"#b8872a", note:"通常枠" },
      { key:"学長配分研究費", label:"学長配分研究費", amount:600000,  color:"#3a72a8", note:"特別配分" },
      { key:"学科配分研究費", label:"学科配分研究費", amount:1000000, color:"#4a8c62", note:"2025年度許可" },
    ]
  }
};
const CATEGORIES = ["旅費","消耗品費","備品費","図書費","印刷費","謝金","その他"];
const CAT_COLORS = {
  "旅費":"#b8872a","消耗品費":"#4a8c62","備品費":"#3a72a8",
  "図書費":"#a05a3a","印刷費":"#7a50a8","謝金":"#8a8030","その他":"#707070"
};
const T = {
  bg:"#f5f0e8", card:"#ffffff", cardAlt:"#fdf9f3",
  border:"#ddd6c8", borderLight:"#ede7dc",
  text:"#2a2520", textMid:"#6a6050", textLight:"#9a9080",
  accent:"#b8872a", accentLight:"#f0e4cc",
  success:"#2d6e48", successBg:"#e8f4ee",
  danger:"#8c2a2a", dangerBg:"#faeaea",
  warn:"#7a5a00", warnBg:"#fff8e8",
};

const fmt = n => `¥${Number(n).toLocaleString()}`;
const pct = (a,b) => b > 0 ? ((a/b)*100).toFixed(1) : "0.0";
const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2)}`;

// 伝票部門名称 → 予算枠キーの確定マッピング
const resolveBudget = (raw, yearBudgets) => {
  const s = String(raw || "").trim();
  // 登録済み予算枠の label または key に部分一致するか確認
  const found = yearBudgets.find(b =>
    s.includes(b.key) || s.includes(b.label) ||
    b.key.includes(s) || b.label.includes(s)
  );
  if (found) return { key: found.key, confidence: "確定", raw: s };
  // キーワードでフォールバック
  let key = "";
  if (s.includes("学科配分"))      key = "学科配分研究費";
  else if (s.includes("学長配分")) key = "学長配分研究費";
  else if (s.length > 0)           key = "個人研究費";
  const fallback = yearBudgets.find(b => b.key === key);
  if (fallback) return { key: fallback.key, confidence: s ? "未知" : "不明", raw: s };
  return { key: yearBudgets[0]?.key || "", confidence: s ? "未知" : "不明", raw: s };
};

// 品目名からカテゴリを推定
const guessCategory = (desc) => {
  const d = String(desc || "").toLowerCase();
  if (/旅費|出張|交通|新幹線|飛行機|ホテル|宿泊|バス|タクシー/.test(d)) return "旅費";
  if (/書籍|図書|文献|雑誌|journal/.test(d))                           return "図書費";
  if (/印刷|コピー|製本/.test(d))                                       return "印刷費";
  if (/謝金|謝礼|講師料/.test(d))                                       return "謝金";
  if (/機材|機器|器具|カメラ|パソコン|pc|mac|scanner|スキャナ/.test(d)) return "備品費";
  if (/消耗|文具|素材|材料|用紙|インク|トナー/.test(d))                 return "消耗品費";
  return "その他";
};

// ヘッダー配列からキーワードに一致する列インデックスを返す
const findColIdx = (headers, ...keywords) => {
  const normalize = s => String(s).replace(/[\s　・\-_]/g, "").toLowerCase();
  return headers.findIndex(h => keywords.some(w => normalize(h).includes(normalize(w))));
};

// ExcelシリアルナンバーをYYYY-MM-DDに変換
const excelDateToStr = (val) => {
  if (typeof val === "number") {
    try {
      const d = XLSX.SSF.parse_date_code(val);
      if (d) return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
    } catch {}
  }
  return String(val || "").trim();
};

// ─────────────────────────────────────────────
// CSV処理
// ─────────────────────────────────────────────
const parseCSV = (text, yearBudgets, currentYear, fileName) => {
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (lines.length < 2) return [];

  // 区切り文字の自動検出（タブ or カンマ）
  const delimiter = lines[0].includes("\t") ? "\t" : ",";
  const splitLine = line => {
    if (delimiter === ",") {
      const result = [];
      let current = "";
      let inQuote = false;
      for (const ch of line) {
        if (ch === '"') { inQuote = !inQuote; }
        else if (ch === "," && !inQuote) { result.push(current.trim()); current = ""; }
        else { current += ch; }
      }
      result.push(current.trim());
      return result;
    }
    return line.split("\t").map(s => s.trim());
  };

  const headers = splitLine(lines[0]).map(h => h.replace(/"/g,"").trim());
  const di = findColIdx(headers,"日付","年月日","date","日");
  const mi = findColIdx(headers,"摘要","内容","品目","商品","件名","明細","説明","memo","name");
  const ai = findColIdx(headers,"金額","支出","出金","請求","amount","合計","税込");
  const bi = findColIdx(headers,"伝票部門名称","部門名称","部門名","部門","区分");

  return lines.slice(1).flatMap((line) => {
    const c = splitLine(line).map(s => s.replace(/"/g,"").trim());
    if (c.length < 2) return [];
    let amount = 0;
    if (ai >= 0) {
      amount = Math.abs(parseInt((c[ai]||"0").replace(/[¥,\s円]/g,""),10)) || 0;
    } else {
      for (const cell of c) {
        const n = Math.abs(parseInt(cell.replace(/[¥,\s円]/g,""),10)) || 0;
        if (n > 0) { amount = n; break; }
      }
    }
    if (!amount) return [];
    const rawBudget = bi >= 0 ? (c[bi]||"") : "";
    const resolved = resolveBudget(rawBudget, yearBudgets);
    const description = mi >= 0 ? (c[mi]||"") : c.filter(v=>v&&isNaN(v.replace(/[¥,円]/g,""))).join(" ");
    return [{
      _id:uid(), _source:fileName,
      _budget_raw:resolved.raw, _budget_confidence:resolved.confidence,
      year:currentYear,
      date: di >= 0 ? c[di] : "",
      description, amount,
      category: guessCategory(description),
      budget: resolved.key,
    }];
  });
};

// ─────────────────────────────────────────────
// Excel処理
// ─────────────────────────────────────────────
const parseExcel = (buf, yearBudgets, currentYear, fileName) => {
  const wb = XLSX.read(buf, { type:"array", cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const jsonRows = XLSX.utils.sheet_to_json(ws, { defval:"", raw:true });
  if (!jsonRows.length) return [];
  const keys = Object.keys(jsonRows[0]);
  const dk = keys.find(k => /日付|年月日|date/i.test(k));
  const mk = keys.find(k => /摘要|内容|品目|商品|memo|name|件名/i.test(k));
  const ak = keys.find(k => /金額|支出|出金|請求|amount/i.test(k));
  const bk = keys.find(k => /伝票部門名称|部門名称|部門名|区分/i.test(k));
  return jsonRows.map((row) => {
    const amount = Math.abs(parseInt(String(row[ak] || "0").replace(/[¥,\s円]/g,""), 10)) || 0;
    if (!amount) return null;
    const rawBudget = bk ? String(row[bk] || "").trim() : "";
    const resolved = resolveBudget(rawBudget, yearBudgets);
    const description = mk ? String(row[mk] || "").trim() : "";
    return {
      _id: uid(), _source: fileName,
      _budget_raw: resolved.raw, _budget_confidence: resolved.confidence,
      year: currentYear, date: dk ? excelDateToStr(row[dk]) : "",
      description, amount,
      category: guessCategory(description),
      budget: resolved.key,
    };
  }).filter(Boolean);
};

// ─────────────────────────────────────────────
// 信頼度バッジ
// ─────────────────────────────────────────────
const ConfidenceBadge = ({ confidence, raw }) => {
  const map = {
    "確定": { col:T.success, bg:T.successBg, icon:"✓", text: raw ? `「${raw}」から確定` : "確定" },
    "未知": { col:T.warn,    bg:T.warnBg,    icon:"△", text: `「${raw}」未知の部門名` },
    "不明": { col:T.danger,  bg:T.dangerBg,  icon:"⚠", text: "部門名なし・要選択" },
  };
  const s = map[confidence] || map["不明"];
  return (
    <div style={{fontSize:10,color:s.col,background:s.bg,border:`1px solid ${s.col}33`,
      borderRadius:4,padding:"2px 6px",marginBottom:4,display:"inline-block",whiteSpace:"nowrap"}}>
      {s.icon} {s.text}
    </div>
  );
};

// ─────────────────────────────────────────────
// メインコンポーネント
// ─────────────────────────────────────────────
export default function App() {
  // ── state（localStorageから復元） ──────────────
  const [fiscalYears, setFiscalYears] = useState(() => {
    try { return JSON.parse(localStorage.getItem("research_fiscalYears")) || INITIAL_YEARS; }
    catch { return INITIAL_YEARS; }
  });
  const [expenses, setExpenses] = useState(() => {
    try { return JSON.parse(localStorage.getItem("research_expenses")) || []; }
    catch { return []; }
  });
  const [apiKey, setApiKey] = useState(() =>
    localStorage.getItem("research_anthropic_key") || ""
  );
  const [currentYear, setCurrentYear]     = useState("2025");
  const [activeTab, setActiveTab]         = useState("overview");
  const [notification, setNotification]   = useState(null);
  const [isDragging, setIsDragging]       = useState(false);
  const [isProcessing, setIsProcessing]   = useState(false);
  const [processMsg, setProcessMsg]       = useState("");
  const [reviewItems, setReviewItems]     = useState(null);
  const [showSettings, setShowSettings]   = useState(false);
  const [settingsDraft, setSettingsDraft] = useState(null);
  const fileInputRef = useRef();

  // ── localStorage への自動保存 ──────────────────
  useEffect(() => {
    localStorage.setItem("research_fiscalYears", JSON.stringify(fiscalYears));
  }, [fiscalYears]);
  useEffect(() => {
    localStorage.setItem("research_expenses", JSON.stringify(expenses));
  }, [expenses]);

  const saveApiKey = (key) => {
    setApiKey(key);
    localStorage.setItem("research_anthropic_key", key);
  };

  // ── 通知 ───────────────────────────────────────
  const showNotif = (msg, type="success") => {
    setNotification({msg,type});
    setTimeout(() => setNotification(null), 3500);
  };

  // ── 現在年度のデータ ────────────────────────────
  const yearConfig  = fiscalYears[currentYear] || { budgets:[] };
  const yearBudgets = yearConfig.budgets;
  const yearExp     = expenses.filter(e => e.year === currentYear);
  const totalBudget = yearBudgets.reduce((s,b) => s+b.amount, 0);
  const totalSpent  = yearExp.reduce((s,e) => s+e.amount, 0);

  const budgetStats = yearBudgets.map(b => {
    const spent = yearExp.filter(e=>e.budget===b.key).reduce((s,e)=>s+e.amount,0);
    return { ...b, spent, remaining: b.amount-spent, pctVal: pct(spent,b.amount) };
  });

  const catData = CATEGORIES.map(cat => ({
    name: cat,
    value: yearExp.filter(e=>e.category===cat).reduce((s,e)=>s+e.amount,0),
    color: CAT_COLORS[cat]
  })).filter(c => c.value > 0);

  const monthlyData = (() => {
    const m = {};
    yearExp.forEach(e => {
      const k = e.date ? e.date.slice(0,7) : "不明";
      if (!m[k]) m[k] = { month: k.replace("-","/") };
      yearBudgets.forEach(b => { m[k][b.key] = (m[k][b.key]||0) + (e.budget===b.key ? e.amount : 0); });
    });
    return Object.values(m).sort((a,b) => a.month.localeCompare(b.month));
  })();

  // ── PDF処理（Claude API） ──────────────────────
  const parsePDF = async (file) => {
    if (!apiKey) throw new Error("APIキーが設定されていません。「⚙ 予算設定」からAnthropicのAPIキーを登録してください。");
    const base64 = await new Promise((res,rej) => {
      const r = new FileReader();
      r.onload = () => res(r.result.split(",")[1]);
      r.onerror = rej;
      r.readAsDataURL(file);
    });
    const budgetKeys = yearBudgets.map(b=>b.label).join("・");
    const prompt = `日本の大学研究費精算伝票・購入要望書から支出項目を全件抽出してください。JSON配列のみ返してください（説明文・コードブロック不要）。
フィールド: date(YYYY-MM-DD), description(品目・摘要), amount(正の整数・税込), category(旅費/消耗品費/備品費/図書費/印刷費/謝金/その他), budget_raw(伝票の部門名・区分名をそのままコピー)
予算枠候補: ${budgetKeys}
データなし→[]`;
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method:"POST",
      headers:{
        "Content-Type":"application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 8000,
        messages:[{ role:"user", content:[
          { type:"document", source:{ type:"base64", media_type:"application/pdf", data:base64 }},
          { type:"text", text: prompt }
        ]}]
      })
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(`API Error ${res.status}: ${err?.error?.message || res.statusText}`);
    }
    const data = await res.json();
    const raw = (data.content||[]).filter(b=>b.type==="text").map(b=>b.text).join("")
      .replace(/```[a-z]*\n?|```/g,"").trim();
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) throw new Error("API応答が不正です");
    return parsed.map(r => {
      const resolved = resolveBudget(r.budget_raw || "", yearBudgets);
      return {
        _id: uid(), _source: file.name,
        _budget_raw: resolved.raw, _budget_confidence: resolved.confidence,
        year: currentYear,
        date: r.date || "", description: r.description || "",
        amount: Math.abs(Number(r.amount)) || 0,
        category: CATEGORIES.includes(r.category) ? r.category : "その他",
        budget: resolved.key,
      };
    }).filter(r => r.amount > 0);
  };

  // ── ファイル処理 ────────────────────────────────
  const processFiles = async (files) => {
    if (yearBudgets.length === 0) {
      showNotif("先に「⚙ 予算設定」から予算枠を登録してください", "error");
      return;
    }
    setIsProcessing(true);
    const all = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      setProcessMsg(`(${i+1}/${files.length}) "${file.name}" を処理中...`);
      try {
        const ext = file.name.split(".").pop().toLowerCase();
        let items = [];
        if (ext === "pdf") {
          items = await parsePDF(file);
        } else if (ext === "csv" || ext === "tsv") {
          const buf = await file.arrayBuffer();
          let text;
          try { text = new TextDecoder("utf-8", { fatal: true }).decode(buf); }
          catch { text = new TextDecoder("shift-jis").decode(buf); }
          items = parseCSV(text, yearBudgets, currentYear, file.name);
        } else if (["xlsx","xls"].includes(ext)) {
          const buf = await file.arrayBuffer();
          items = parseExcel(buf, yearBudgets, currentYear, file.name);
        } else {
          showNotif(`"${file.name}" は非対応形式です（PDF・CSV・Excel のみ）`, "error");
          continue;
        }
        if (items.length === 0) {
          showNotif(`"${file.name}" から支出データを検出できませんでした`, "error");
        } else {
          all.push(...items);
        }
      } catch (err) {
        showNotif(`"${file.name}" エラー: ${err.message}`, "error");
      }
    }
    setIsProcessing(false);
    if (all.length > 0) setReviewItems(all);
  };

  const updateReview = (id, field, val) =>
    setReviewItems(prev => prev.map(r => r._id===id ? {...r, [field]: field==="amount" ? Number(val)||0 : val} : r));
  const deleteReview = id =>
    setReviewItems(prev => { const n = prev.filter(r=>r._id!==id); return n.length ? n : null; });
  const confirmImport = () => {
    const valid = reviewItems.filter(r => r.amount > 0);
    setExpenses(prev => [...prev, ...valid.map(({_id,_source,_budget_raw,_budget_confidence,...rest}) => ({...rest, id:uid()}))]);
    showNotif(`${valid.length}件を登録しました`);
    setReviewItems(null);
  };

  // ── 年度操作 ────────────────────────────────────
  const addYear = () => {
    const next = String(Math.max(...Object.keys(fiscalYears).map(Number)) + 1);
    if (fiscalYears[next]) return;
    setFiscalYears(prev => ({...prev, [next]:{ budgets:[] }}));
    setCurrentYear(next);
    setSettingsDraft({ budgets:[], apiKeyDraft: apiKey });
    setShowSettings(true);
    showNotif(`${next}年度を追加しました。予算枠を設定してください。`);
  };
  const openSettings = () => {
    setSettingsDraft({...JSON.parse(JSON.stringify(fiscalYears[currentYear])), apiKeyDraft: apiKey});
    setShowSettings(true);
  };
  const saveSettings = () => {
    const { apiKeyDraft, ...yearData } = settingsDraft;
    setFiscalYears(prev => ({...prev, [currentYear]: yearData}));
    if (apiKeyDraft !== undefined) saveApiKey(apiKeyDraft);
    setShowSettings(false);
    showNotif("設定を保存しました");
  };
  const draftUpdate = (i, field, val) => {
    const b = [...settingsDraft.budgets];
    b[i] = {...b[i], [field]: field==="amount" ? Number(val)||0 : val};
    setSettingsDraft({...settingsDraft, budgets:b});
  };
  const draftAdd = () => {
    const used = settingsDraft.budgets.map(b=>b.color);
    const color = BUDGET_COLORS.find(c=>!used.includes(c)) || BUDGET_COLORS[0];
    setSettingsDraft({...settingsDraft, budgets:[...settingsDraft.budgets,
      {key:`budget_${uid()}`, label:"新しい予算枠", amount:0, color, note:""}
    ]});
  };
  const draftDelete = i => {
    const b = [...settingsDraft.budgets]; b.splice(i,1);
    setSettingsDraft({...settingsDraft, budgets:b});
  };

  // ── スタイル ────────────────────────────────────
  const cardStyle = { background:T.card, border:`1px solid ${T.border}`, borderRadius:12, padding:24, boxShadow:"0 2px 8px rgba(0,0,0,0.05)" };
  const selStyle  = { background:T.bg, border:`1px solid ${T.border}`, borderRadius:6, padding:"5px 8px", color:T.text, fontSize:12, fontFamily:"Georgia,serif" };
  const inpStyle  = { background:T.bg, border:`1px solid ${T.border}`, borderRadius:6, padding:"6px 10px", color:T.text, fontSize:13, width:"100%", boxSizing:"border-box", fontFamily:"Georgia,serif" };
  const btnStyle  = (bg, col) => ({ background:bg, color:col, border:"none", borderRadius:8, padding:"9px 20px", fontSize:13, cursor:"pointer", fontWeight:600, fontFamily:"Georgia,serif" });

  // ── レンダリング ───────────────────────────────
  return (
    <div style={{fontFamily:"Georgia,serif", background:T.bg, minHeight:"100vh", color:T.text}}>

      {/* 通知トースト */}
      {notification && (
        <div style={{position:"fixed",top:24,right:24,zIndex:300,
          background:notification.type==="error"?T.dangerBg:T.successBg,
          border:`1px solid ${notification.type==="error"?"#d08080":"#6aaa88"}`,
          color:notification.type==="error"?T.danger:T.success,
          padding:"12px 20px",borderRadius:8,fontSize:14,boxShadow:"0 4px 20px rgba(0,0,0,0.15)",maxWidth:320}}>
          {notification.msg}
        </div>
      )}

      {/* ── 年度設定モーダル ── */}
      {showSettings && settingsDraft && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:24,overflowY:"auto"}}>
          <div style={{background:T.card,borderRadius:16,width:"100%",maxWidth:560,boxShadow:"0 20px 60px rgba(0,0,0,0.25)"}}>
            <div style={{padding:"20px 26px",borderBottom:`1px solid ${T.border}`}}>
              <div style={{fontSize:11,color:T.accent,letterSpacing:"0.2em",marginBottom:4}}>設定</div>
              <div style={{fontSize:18,fontWeight:700}}>{currentYear}年度　予算枠とAPIキー</div>
            </div>
            <div style={{padding:"20px 26px",maxHeight:"60vh",overflowY:"auto"}}>

              {/* APIキー設定 */}
              <div style={{marginBottom:20,padding:14,background:T.warnBg,border:`1px solid ${T.warn}33`,borderRadius:10}}>
                <div style={{fontSize:11,color:T.warn,fontWeight:700,marginBottom:6}}>🔑 Anthropic APIキー（PDF解析に必要）</div>
                <input
                  type="password"
                  placeholder="sk-ant-api03-..."
                  value={settingsDraft.apiKeyDraft || ""}
                  onChange={e => setSettingsDraft({...settingsDraft, apiKeyDraft:e.target.value})}
                  style={{...inpStyle, fontFamily:"monospace", fontSize:12}}
                />
                <div style={{fontSize:10,color:T.textLight,marginTop:4}}>
                  Anthropic Console（console.anthropic.com）で発行できます。ブラウザのlocalStorageに保存されます。
                </div>
              </div>

              {/* 予算枠リスト */}
              <div style={{fontSize:12,color:T.textLight,marginBottom:10,fontWeight:700}}>予算枠</div>
              {settingsDraft.budgets.map((b,i) => (
                <div key={i} style={{marginBottom:16,paddingBottom:16,borderBottom:`1px solid ${T.borderLight}`}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 130px 36px",gap:10,marginBottom:8}}>
                    <div>
                      <div style={{fontSize:10,color:T.textLight,marginBottom:3}}>予算枠名（伝票の部門名と一致させる）</div>
                      <input value={b.label} onChange={e=>draftUpdate(i,"label",e.target.value)} style={inpStyle}/>
                    </div>
                    <div>
                      <div style={{fontSize:10,color:T.textLight,marginBottom:3}}>上限金額（円）</div>
                      <input type="number" value={b.amount} onChange={e=>draftUpdate(i,"amount",e.target.value)} style={{...inpStyle,textAlign:"right"}}/>
                    </div>
                    <div style={{paddingTop:18}}>
                      <button onClick={()=>draftDelete(i)} style={{...btnStyle(T.dangerBg,T.danger),padding:"6px 8px",fontWeight:400}}>✕</button>
                    </div>
                  </div>
                  <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                    <span style={{fontSize:10,color:T.textLight}}>色：</span>
                    {BUDGET_COLORS.map(c => (
                      <div key={c} onClick={()=>draftUpdate(i,"color",c)}
                        style={{width:20,height:20,borderRadius:4,background:c,cursor:"pointer",
                          border:b.color===c?`3px solid ${T.text}`:"2px solid transparent"}}/>
                    ))}
                    <input value={b.note} onChange={e=>draftUpdate(i,"note",e.target.value)} placeholder="備考（任意）"
                      style={{...inpStyle,width:160,fontSize:11,marginLeft:"auto"}}/>
                  </div>
                </div>
              ))}
              <button onClick={draftAdd}
                style={{width:"100%",background:T.cardAlt,border:`1px dashed ${T.border}`,borderRadius:8,padding:10,fontSize:13,color:T.textMid,cursor:"pointer"}}>
                ＋ 予算枠を追加
              </button>
            </div>
            <div style={{padding:"16px 26px",borderTop:`1px solid ${T.border}`,display:"flex",justifyContent:"flex-end",gap:10}}>
              <button onClick={()=>setShowSettings(false)} style={btnStyle(T.cardAlt,T.textMid)}>キャンセル</button>
              <button onClick={saveSettings} style={btnStyle(T.success,"#fff")}>保存する</button>
            </div>
          </div>
        </div>
      )}

      {/* ── レビューモーダル ── */}
      {reviewItems && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:150,display:"flex",alignItems:"flex-start",justifyContent:"center",padding:"28px 16px",overflowY:"auto"}}>
          <div style={{background:T.card,borderRadius:16,width:"100%",maxWidth:1000,boxShadow:"0 20px 60px rgba(0,0,0,0.25)"}}>
            <div style={{padding:"20px 26px",borderBottom:`1px solid ${T.border}`}}>
              <div style={{fontSize:11,color:T.success,letterSpacing:"0.2em",marginBottom:4}}>AI自動抽出結果 — {currentYear}年度</div>
              <div style={{fontSize:18,fontWeight:700}}>内容を確認・編集してから登録してください</div>
              <div style={{display:"flex",gap:10,marginTop:10,flexWrap:"wrap"}}>
                {[["確定",T.success,T.successBg],["未知",T.warn,T.warnBg],["不明",T.danger,T.dangerBg]].map(([conf,col,bg]) => {
                  const count = reviewItems.filter(r=>r._budget_confidence===conf).length;
                  if (!count) return null;
                  return <div key={conf} style={{fontSize:11,color:col,background:bg,border:`1px solid ${col}33`,borderRadius:6,padding:"3px 12px"}}>{conf}：<strong>{count}件</strong></div>;
                })}
                <div style={{fontSize:11,color:T.textMid,marginLeft:"auto"}}>合計 <strong>{fmt(reviewItems.reduce((s,r)=>s+r.amount,0))}</strong>（{reviewItems.length}件）</div>
              </div>
            </div>
            <div style={{overflowX:"auto",maxHeight:"55vh",overflowY:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead style={{position:"sticky",top:0,zIndex:1}}>
                  <tr style={{background:T.cardAlt}}>
                    {["ファイル","日付","内容","金額（円）","カテゴリ","予算枠",""].map((h,i)=>(
                      <th key={i} style={{padding:"10px 10px",fontSize:11,color:T.textLight,fontWeight:600,textAlign:"left",borderBottom:`1px solid ${T.border}`,whiteSpace:"nowrap"}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {reviewItems.map(r => {
                    const bInfo = yearBudgets.find(b=>b.key===r.budget);
                    return (
                      <tr key={r._id} style={{borderBottom:`1px solid ${T.borderLight}`}}>
                        <td style={{padding:"7px 10px",maxWidth:90}}>
                          <span style={{fontSize:10,color:T.textLight,display:"block",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r._source}>{r._source}</span>
                        </td>
                        <td style={{padding:"7px 10px"}}>
                          <input type="date" value={r.date} onChange={e=>updateReview(r._id,"date",e.target.value)}
                            style={{...inpStyle,width:130,fontSize:12}}/>
                        </td>
                        <td style={{padding:"7px 10px"}}>
                          <input value={r.description} onChange={e=>updateReview(r._id,"description",e.target.value)}
                            style={{...inpStyle,minWidth:160}}/>
                        </td>
                        <td style={{padding:"7px 10px"}}>
                          <input type="number" value={r.amount} onChange={e=>updateReview(r._id,"amount",e.target.value)}
                            style={{...inpStyle,width:110,textAlign:"right"}}/>
                        </td>
                        <td style={{padding:"7px 10px"}}>
                          <select value={r.category} onChange={e=>updateReview(r._id,"category",e.target.value)} style={selStyle}>
                            {CATEGORIES.map(c=><option key={c}>{c}</option>)}
                          </select>
                        </td>
                        <td style={{padding:"7px 10px",minWidth:190}}>
                          <ConfidenceBadge confidence={r._budget_confidence} raw={r._budget_raw}/>
                          <select value={r.budget} onChange={e=>updateReview(r._id,"budget",e.target.value)}
                            style={{...selStyle, width:"100%", fontWeight:600, color:bInfo?.color||T.text,
                              borderColor: r._budget_confidence==="確定" ? `${T.success}66` : r._budget_confidence==="不明" ? `${T.danger}66` : T.border}}>
                            {yearBudgets.map(b=><option key={b.key} value={b.key}>{b.label}</option>)}
                          </select>
                        </td>
                        <td style={{padding:"7px 10px"}}>
                          <button onClick={()=>deleteReview(r._id)}
                            style={{...btnStyle(T.dangerBg,T.danger),padding:"3px 8px",fontSize:11,fontWeight:400}}>削除</button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div style={{padding:"16px 26px",borderTop:`1px solid ${T.border}`,display:"flex",justifyContent:"flex-end",gap:10}}>
              <button onClick={()=>setReviewItems(null)} style={btnStyle(T.cardAlt,T.textMid)}>キャンセル</button>
              <button onClick={confirmImport} style={btnStyle(T.success,"#fff")}>✓ この内容で登録する</button>
            </div>
          </div>
        </div>
      )}

      {/* ── メイン画面 ── */}
      <div style={{maxWidth:980,margin:"0 auto",padding:"28px 24px"}}>

        {/* ヘッダー */}
        <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",flexWrap:"wrap",gap:12,marginBottom:20}}>
          <div>
            <div style={{fontSize:11,letterSpacing:"0.25em",color:T.accent,textTransform:"uppercase",marginBottom:4}}>新見公立大学 健康保育学科</div>
            <h1 style={{margin:0,fontSize:24,fontWeight:700,color:T.text}}>研究費ダッシュボード</h1>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:10,padding:"8px 14px",display:"flex",alignItems:"center",gap:8,boxShadow:"0 2px 6px rgba(0,0,0,0.05)"}}>
              <span style={{fontSize:11,color:T.textLight}}>年度</span>
              <select value={currentYear} onChange={e=>setCurrentYear(e.target.value)}
                style={{background:"none",border:"none",fontSize:18,fontWeight:700,color:T.text,cursor:"pointer",fontFamily:"Georgia,serif"}}>
                {Object.keys(fiscalYears).sort().map(y=><option key={y} value={y}>{y}年度</option>)}
              </select>
            </div>
            <button onClick={addYear}      style={btnStyle(T.card,T.textMid)}>＋ 年度追加</button>
            <button onClick={openSettings} style={btnStyle(T.accentLight,T.accent)}>⚙ 予算設定</button>
          </div>
        </div>

        {/* APIキー未設定バナー */}
        {!apiKey && (
          <div style={{background:T.warnBg,border:`1px solid ${T.warn}44`,borderRadius:10,padding:"12px 16px",marginBottom:16,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
            <span style={{fontSize:13,color:T.warn}}>🔑 PDFの自動解析を使うにはAnthropicのAPIキーが必要です</span>
            <button onClick={openSettings} style={{...btnStyle(T.warn,"#fff"),padding:"6px 14px",fontSize:12}}>設定する</button>
          </div>
        )}

        {/* 予算カード */}
        <div style={{display:"grid",gridTemplateColumns:`repeat(${Math.max(Math.min(yearBudgets.length,3),1)},1fr)`,gap:14,marginBottom:14}}>
          {yearBudgets.length === 0 ? (
            <div style={{...cardStyle,textAlign:"center",padding:24,color:T.textLight}}>
              <div style={{marginBottom:6,fontSize:20}}>⚙</div>
              <div style={{fontSize:13}}>「予算設定」から今年度の予算枠を追加してください</div>
            </div>
          ) : budgetStats.map(b => {
            const pctNum = Number(b.pctVal);
            const barColor = pctNum >= 90 ? T.danger : pctNum >= 70 ? "#b8600a" : b.color;
            return (
              <div key={b.key} style={{background:T.card,border:`2px solid ${b.color}33`,borderRadius:12,padding:18,boxShadow:"0 2px 8px rgba(0,0,0,0.05)"}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
                  <div>
                    <div style={{fontSize:12,fontWeight:700,color:b.color}}>{b.label}</div>
                    <div style={{fontSize:10,color:T.textLight}}>{b.note}　上限 {fmt(b.amount)}</div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontSize:10,color:T.textLight}}>残高</div>
                    <div style={{fontSize:15,fontWeight:700,color:b.remaining < 0 ? T.danger : b.remaining < b.amount*0.2 ? "#b8600a" : T.success}}>
                      {fmt(b.remaining)}
                    </div>
                  </div>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
                  <span style={{fontSize:12,color:T.textMid}}>使用 <strong style={{color:b.color}}>{fmt(b.spent)}</strong></span>
                  <span style={{fontSize:12,color:barColor,fontWeight:700}}>{b.pctVal}%</span>
                </div>
                <div style={{height:7,background:T.borderLight,borderRadius:4}}>
                  <div style={{height:"100%",width:`${Math.min(pctNum,100)}%`,background:barColor,borderRadius:4,transition:"width 0.4s"}}/>
                </div>
              </div>
            );
          })}
        </div>

        {/* 合算バー */}
        {yearBudgets.length > 0 && (
          <div style={{...cardStyle,padding:"12px 20px",marginBottom:20}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
              <span style={{fontSize:12,color:T.textMid}}>合計　{fmt(totalSpent)} / {fmt(totalBudget)}</span>
              <span style={{fontSize:13,fontWeight:700}}>{pct(totalSpent,totalBudget)}%　残 {fmt(totalBudget-totalSpent)}</span>
            </div>
            <div style={{height:9,background:T.borderLight,borderRadius:5,overflow:"hidden",display:"flex"}}>
              {budgetStats.map(b=><div key={b.key} style={{height:"100%",width:`${Math.max(0,Number(pct(b.spent,totalBudget)))}%`,background:b.color}}/>)}
            </div>
            <div style={{display:"flex",gap:16,marginTop:5,flexWrap:"wrap"}}>
              {budgetStats.map(b=>(
                <div key={b.key} style={{display:"flex",alignItems:"center",gap:5}}>
                  <div style={{width:10,height:10,borderRadius:2,background:b.color}}/>
                  <span style={{fontSize:11,color:T.textMid}}>{b.label}：{fmt(b.spent)}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ドロップゾーン */}
        <div
          onDragOver={e=>{e.preventDefault();setIsDragging(true);}}
          onDragLeave={()=>setIsDragging(false)}
          onDrop={e=>{e.preventDefault();setIsDragging(false);const files=Array.from(e.dataTransfer.files);if(files.length)processFiles(files);}}
          onClick={()=>!isProcessing && fileInputRef.current?.click()}
          style={{border:`2px dashed ${isDragging?T.success:isProcessing?T.accent:T.border}`,borderRadius:14,padding:"22px 20px",textAlign:"center",
            background:isDragging?T.successBg:isProcessing?T.accentLight:T.cardAlt,
            cursor:isProcessing?"default":"pointer",marginBottom:16,transition:"all 0.2s"}}>
          <input ref={fileInputRef} type="file" accept=".csv,.tsv,.xlsx,.xls,.pdf"
            onChange={e=>{const files=Array.from(e.target.files);if(files.length)processFiles(files);e.target.value="";}}
            multiple style={{display:"none"}}/>
          {isProcessing ? (
            <div>
              <div style={{fontSize:26,marginBottom:8}}>🤖</div>
              <div style={{fontSize:14,color:T.accent,fontWeight:600}}>{processMsg}</div>
              <div style={{fontSize:12,color:T.textLight,marginTop:3}}>しばらくお待ちください...</div>
            </div>
          ) : (
            <div>
              <div style={{fontSize:28,marginBottom:6}}>📂</div>
              <div style={{fontSize:14,color:T.textMid,fontWeight:600,marginBottom:3}}>ファイルをドラッグ＆ドロップ（複数可）</div>
              <div style={{fontSize:12,color:T.textLight}}>クリックしてファイル選択　／　PDF・CSV・TSV・Excel 対応</div>
              <div style={{fontSize:11,color:T.success,marginTop:6}}>✦ Spendia CSV・購入要望書PDF・Excelに対応　✦ 伝票部門名称から予算枠を自動判別</div>
            </div>
          )}
        </div>

        {/* タブ */}
        <div style={{display:"flex",gap:4,marginBottom:18,borderBottom:`1px solid ${T.border}`}}>
          {[["overview","概要"],["expenses","支出一覧"],["chart","グラフ"]].map(([tab,label])=>(
            <button key={tab} onClick={()=>setActiveTab(tab)}
              style={{background:"none",border:"none",cursor:"pointer",padding:"9px 18px",fontSize:13,fontFamily:"Georgia,serif",
                color:activeTab===tab?T.accent:T.textLight,
                borderBottom:activeTab===tab?`2px solid ${T.accent}`:"2px solid transparent",marginBottom:-1}}>
              {label}
            </button>
          ))}
        </div>

        {/* 概要タブ */}
        {activeTab==="overview" && (
          yearExp.length===0 ? (
            <div style={{...cardStyle,textAlign:"center",padding:48,color:T.textLight}}>
              <div style={{fontSize:32,marginBottom:10}}>☝️</div>
              <div style={{fontSize:15,marginBottom:5}}>上のエリアにファイルをドロップしてください</div>
              <div style={{fontSize:13}}>PDF・CSV・Excelを複数まとめて処理できます</div>
            </div>
          ) : (
            <div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:18}}>
                <div style={cardStyle}>
                  <div style={{fontSize:12,color:T.textLight,letterSpacing:"0.12em",marginBottom:14}}>カテゴリ別支出</div>
                  <ResponsiveContainer width="100%" height={170}>
                    <PieChart>
                      <Pie data={catData} cx="50%" cy="50%" innerRadius={48} outerRadius={78} dataKey="value" paddingAngle={2}>
                        {catData.map((e,i)=><Cell key={i} fill={e.color}/>)}
                      </Pie>
                      <Tooltip formatter={v=>fmt(v)} contentStyle={{background:T.card,border:`1px solid ${T.border}`,borderRadius:8,color:T.text,fontSize:12}}/>
                    </PieChart>
                  </ResponsiveContainer>
                  <div style={{display:"flex",flexDirection:"column",gap:6,marginTop:6}}>
                    {[...catData].sort((a,b)=>b.value-a.value).map(c=>(
                      <div key={c.name} style={{display:"flex",alignItems:"center",gap:8}}>
                        <div style={{width:9,height:9,borderRadius:2,background:c.color,flexShrink:0}}/>
                        <div style={{fontSize:12,color:T.textMid,flex:1}}>{c.name}</div>
                        <div style={{fontSize:13,color:T.text}}>{fmt(c.value)}</div>
                      </div>
                    ))}
                  </div>
                </div>
                <div style={cardStyle}>
                  <div style={{fontSize:12,color:T.textLight,letterSpacing:"0.12em",marginBottom:14}}>予算別消化率</div>
                  {budgetStats.map(b=>{
                    const pctNum = Number(b.pctVal);
                    const barColor = pctNum >= 90 ? T.danger : pctNum >= 70 ? "#b8600a" : b.color;
                    return (
                      <div key={b.key} style={{marginBottom:16}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                          <span style={{fontSize:12,fontWeight:700,color:b.color}}>{b.label}</span>
                          <span style={{fontSize:11,color:T.textMid}}>{b.pctVal}%　{fmt(b.spent)} / {fmt(b.amount)}</span>
                        </div>
                        <div style={{height:9,background:T.borderLight,borderRadius:5}}>
                          <div style={{height:"100%",width:`${Math.min(pctNum,100)}%`,background:barColor,borderRadius:5}}/>
                        </div>
                        {b.remaining < 0 && (
                          <div style={{fontSize:10,color:T.danger,marginTop:2}}>⚠ 予算超過：{fmt(Math.abs(b.remaining))}</div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
              <div style={{...cardStyle,marginTop:18}}>
                <div style={{fontSize:12,color:T.textLight,letterSpacing:"0.12em",marginBottom:14}}>最近の支出</div>
                {[...yearExp].sort((a,b)=>b.date.localeCompare(a.date)).slice(0,8).map(exp=>{
                  const bInfo=yearBudgets.find(b=>b.key===exp.budget);
                  return(
                    <div key={exp.id} style={{display:"flex",alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${T.borderLight}`,gap:10,flexWrap:"wrap"}}>
                      <div style={{fontSize:11,color:T.textLight,width:84,flexShrink:0}}>{exp.date}</div>
                      <span style={{fontSize:10,padding:"1px 6px",borderRadius:4,background:(bInfo?.color||"#888")+"22",color:bInfo?.color||"#888",flexShrink:0,fontWeight:700}}>{exp.budget}</span>
                      <span style={{fontSize:10,padding:"1px 6px",borderRadius:4,background:CAT_COLORS[exp.category]+"22",color:CAT_COLORS[exp.category],flexShrink:0}}>{exp.category}</span>
                      <div style={{flex:1,fontSize:13,color:T.textMid,minWidth:120}}>{exp.description}</div>
                      <div style={{fontSize:13,color:T.text,fontWeight:600}}>{fmt(exp.amount)}</div>
                    </div>
                  );
                })}
              </div>
            </div>
          )
        )}

        {/* 支出一覧タブ */}
        {activeTab==="expenses" && (
          yearExp.length===0 ? (
            <div style={{...cardStyle,textAlign:"center",padding:48,color:T.textLight}}>
              <div style={{fontSize:32,marginBottom:10}}>📋</div>
              <div style={{fontSize:15}}>データがまだありません</div>
            </div>
          ) : (
            <div style={{...cardStyle,overflow:"hidden",padding:0}}>
              <div style={{padding:"10px 16px",borderBottom:`1px solid ${T.border}`,background:T.cardAlt,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{fontSize:12,color:T.textLight}}>{yearExp.length}件　合計 {fmt(totalSpent)}</span>
                <button onClick={()=>{if(window.confirm(`${currentYear}年度の支出データを全件削除しますか？`))setExpenses(prev=>prev.filter(e=>e.year!==currentYear));}}
                  style={{...btnStyle(T.dangerBg,T.danger),padding:"4px 10px",fontSize:11,fontWeight:400}}>全削除</button>
              </div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead>
                    <tr style={{background:T.cardAlt}}>
                      {["日付","予算枠","カテゴリ","内容","金額",""].map((h,i)=>(
                        <th key={i} style={{padding:"10px 12px",fontSize:11,color:T.textLight,fontWeight:600,textAlign:"left",borderBottom:`1px solid ${T.border}`}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {[...yearExp].sort((a,b)=>b.date.localeCompare(a.date)).map(exp=>{
                      const bInfo=yearBudgets.find(b=>b.key===exp.budget);
                      return(
                        <tr key={exp.id} style={{borderBottom:`1px solid ${T.borderLight}`}}>
                          <td style={{padding:"10px 12px",fontSize:12,color:T.textLight,whiteSpace:"nowrap"}}>{exp.date}</td>
                          <td style={{padding:"10px 12px"}}>
                            <span style={{fontSize:10,padding:"2px 6px",borderRadius:4,background:(bInfo?.color||"#888")+"22",color:bInfo?.color||"#888",fontWeight:700,whiteSpace:"nowrap"}}>{exp.budget}</span>
                          </td>
                          <td style={{padding:"10px 12px"}}>
                            <span style={{fontSize:10,padding:"2px 6px",borderRadius:4,background:CAT_COLORS[exp.category]+"22",color:CAT_COLORS[exp.category],whiteSpace:"nowrap"}}>{exp.category}</span>
                          </td>
                          <td style={{padding:"10px 12px",fontSize:13,color:T.textMid}}>{exp.description}</td>
                          <td style={{padding:"10px 12px",fontSize:14,color:T.text,fontWeight:600,whiteSpace:"nowrap"}}>{fmt(exp.amount)}</td>
                          <td style={{padding:"10px 12px"}}>
                            <button onClick={()=>setExpenses(prev=>prev.filter(e=>e.id!==exp.id))}
                              style={{...btnStyle(T.dangerBg,T.danger),padding:"3px 8px",fontSize:11,fontWeight:400}}>削除</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )
        )}

        {/* グラフタブ */}
        {activeTab==="chart" && (
          yearExp.length===0 ? (
            <div style={{...cardStyle,textAlign:"center",padding:48,color:T.textLight}}>
              <div style={{fontSize:32,marginBottom:10}}>📊</div>
              <div style={{fontSize:15}}>データをインポートするとグラフが表示されます</div>
            </div>
          ) : (
            <div style={{display:"flex",flexDirection:"column",gap:18}}>
              <div style={cardStyle}>
                <div style={{fontSize:12,color:T.textLight,letterSpacing:"0.12em",marginBottom:18}}>月別支出推移（予算枠別）</div>
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={monthlyData} margin={{top:5,right:10,left:10,bottom:5}}>
                    <CartesianGrid strokeDasharray="3 3" stroke={T.borderLight}/>
                    <XAxis dataKey="month" tick={{fill:T.textLight,fontSize:12}}/>
                    <YAxis tick={{fill:T.textLight,fontSize:11}} tickFormatter={v=>`¥${(v/1000).toFixed(0)}k`}/>
                    <Tooltip formatter={v=>fmt(v)} contentStyle={{background:T.card,border:`1px solid ${T.border}`,borderRadius:8,color:T.text,fontSize:12}}/>
                    <Legend formatter={v=><span style={{fontSize:12,color:T.textMid}}>{v}</span>}/>
                    {yearBudgets.map(b=><Bar key={b.key} dataKey={b.key} name={b.label} fill={b.color} radius={[3,3,0,0]} stackId="a"/>)}
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div style={cardStyle}>
                <div style={{fontSize:12,color:T.textLight,letterSpacing:"0.12em",marginBottom:18}}>カテゴリ別支出分布</div>
                <div style={{display:"flex",alignItems:"center",gap:28,flexWrap:"wrap"}}>
                  <ResponsiveContainer width={200} height={200}>
                    <PieChart>
                      <Pie data={catData} cx="50%" cy="50%" outerRadius={90} dataKey="value" paddingAngle={2}>
                        {catData.map((e,i)=><Cell key={i} fill={e.color}/>)}
                      </Pie>
                      <Tooltip formatter={v=>fmt(v)} contentStyle={{background:T.card,border:`1px solid ${T.border}`,borderRadius:8,color:T.text,fontSize:12}}/>
                    </PieChart>
                  </ResponsiveContainer>
                  <div style={{flex:1,minWidth:200}}>
                    {[...catData].sort((a,b)=>b.value-a.value).map(cat=>(
                      <div key={cat.name} style={{marginBottom:12}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                          <span style={{fontSize:13,color:T.textMid}}>{cat.name}</span>
                          <span style={{fontSize:13,color:T.text}}>{fmt(cat.value)}</span>
                        </div>
                        <div style={{height:6,background:T.borderLight,borderRadius:3}}>
                          <div style={{height:"100%",width:`${totalSpent>0?(cat.value/totalSpent*100).toFixed(1):0}%`,background:cat.color,borderRadius:3}}/>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          )
        )}

        <div style={{marginTop:28,textAlign:"center",fontSize:11,color:T.textLight,letterSpacing:"0.1em"}}>
          新見公立大学 健康保育学科 ／ {currentYear}年度研究費管理
        </div>
      </div>
    </div>
  );
}
