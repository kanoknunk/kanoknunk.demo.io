const { useState, useEffect, useMemo } = React;

/* ═══════════════════════════════════════════════
   CONSTANTS
═══════════════════════════════════════════════ */
const TEAL = "#0F6E56";
const SEV = {
  normal: { label:"ปกติ",       bg:"#ECFDF5", border:"#6EE7B7", color:"#065F46", dot:"#10B981", emoji:"🟢" },
  watch:  { label:"เฝ้าระวัง", bg:"#FFFBEB", border:"#FCD34D", color:"#92400E", dot:"#F59E0B", emoji:"🟡" },
  urgent: { label:"วิกฤต",     bg:"#FFF1F2", border:"#FECDD3", color:"#991B1B", dot:"#EF4444", emoji:"🔴" },
};
const SHIFTS = [
  { k:"morning",   l:"เช้า", t:"06–14" },
  { k:"afternoon", l:"บ่าย", t:"14–22" },
  { k:"night",     l:"ดึก",  t:"22–06" },
];
const CATS = ["ป้องกัน","ดูแล","การแพทย์","ทั่วไป"];
const STOCK0 = [
  {id:"s1", name:"ถุงมือยาง S",    cat:"ป้องกัน",  unit:"คู่",   qty:80, max:100},
  {id:"s2", name:"ถุงมือยาง M",    cat:"ป้องกัน",  unit:"คู่",   qty:14, max:100},
  {id:"s3", name:"ถุงมือยาง L",    cat:"ป้องกัน",  unit:"คู่",   qty:60, max:100},
  {id:"s4", name:"ผ้าอ้อม M",      cat:"ดูแล",     unit:"ชิ้น",  qty:120,max:200},
  {id:"s5", name:"ผ้าอ้อม L",      cat:"ดูแล",     unit:"ชิ้น",  qty:22, max:200},
  {id:"s6", name:"แอลกอฮอล์เจล",  cat:"ทั่วไป",   unit:"ขวด",  qty:38, max:50 },
  {id:"s7", name:"หน้ากากอนามัย", cat:"ป้องกัน",  unit:"กล่อง", qty:25, max:30 },
  {id:"s8", name:"ผ้าพันแผล",      cat:"การแพทย์",unit:"ม้วน",  qty:50, max:60 },
  {id:"s9", name:"สำลีก้อน",       cat:"การแพทย์",unit:"ถุง",   qty:30, max:40 },
  {id:"s10",name:"พลาสเตอร์",      cat:"การแพทย์",unit:"กล่อง", qty:8,  max:30 },
];
const BEDS = Array.from({length:40},(_,i)=>i+1);

/* ═══════════════════════════════════════════════
   DATA SERVICE  — localStorage + Google Sheets bridge
═══════════════════════════════════════════════ */
const DS = {
  _url: "",
  init(url){ this._url = url||""; },
  async _lg(k,fb){ try{ const r=await window.storage.get(k); return r?JSON.parse(r.value):fb; }catch{ return fb; } },
  async _ls(k,v){ try{ await window.storage.set(k,JSON.stringify(v)); }catch{} },
  async load(table,fb){
    const local = await this._lg(`cc_${table}`, fb);
    if(this._url){ try{
      const r = await fetch(`${this._url}?action=get&table=${table}`);
      const j = await r.json();
      if(j.ok && j.data) return j.data;
    }catch{} }
    return local;
  },
  async save(table,data){
    await this._ls(`cc_${table}`,data);
    if(this._url){ try{ 
      // ✅ จุดที่แก้ 1: เปลี่ยนเป็น text/plain
      await fetch(this._url,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({action:"set",table,data})}); 
    }catch{} }
  },
  async syncAll(tables){
    if(!this._url) return {ok:false,msg:"ยังไม่ได้ตั้งค่า URL"};
    const data={};
    for(const t of tables) data[t]=await this._lg(`cc_${t}`,null);
    try{
      // ✅ จุดที่แก้ 2: เปลี่ยนเป็น text/plain
      const r=await fetch(this._url,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({action:"syncAll",data})});
      return r.json();
    }catch(e){ return {ok:false,msg:e.message}; }
  },
};

/* Google Apps Script template (readonly string) */
const GAS_CODE = `// ① สร้าง Google Sheet ใหม่ → copy Sheet ID จาก URL
// ② Extensions → Apps Script → วางโค้ดนี้ → แก้ SHEET_ID
// ③ Deploy → New deployment → Web App
//    Execute as: Me | Who has access: Anyone
// ④ Copy Web App URL ไปใส่ในแอพ

const SHEET_ID = "YOUR_SPREADSHEET_ID_HERE";

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.openById(SHEET_ID);
  if (d.action === "set") {
    const sh = ss.getSheetByName(d.table) || ss.insertSheet(d.table);
    sh.clearContents();
    sh.getRange(1,1).setValue("data");
    sh.getRange(2,1).setValue(JSON.stringify(d.data));
    return ok({saved: d.table});
  }
  if (d.action === "syncAll") {
    Object.entries(d.data).forEach(([t, v]) => {
      if (!v) return;
      const sh = ss.getSheetByName(t) || ss.insertSheet(t);
      sh.clearContents();
      sh.getRange(1,1).setValue("data");
      sh.getRange(2,1).setValue(JSON.stringify(v));
    });
    return ok({synced: Object.keys(d.data)});
  }
  return ok(null);
}

function doGet(e) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(e.parameter.table);
  if (!sh) return ok(null);
  const v = sh.getRange(2,1).getValue();
  return ok(v ? JSON.parse(v) : null);
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ok:true, data}))
    .setMimeType(ContentService.MimeType.JSON);
}`;

/* ═══════════════════════════════════════════════
   LINE NOTIFY
═══════════════════════════════════════════════ */
async function line(token, msg){
  if(!token) return false;
  try{
    const r = await fetch("https://corsproxy.io/?https://notify-api.line.me/api/notify",{
      method:"POST",
      headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/x-www-form-urlencoded"},
      body: new URLSearchParams({message:msg}),
    });
    return r.ok;
  }catch{ return false; }
}

/* ═══════════════════════════════════════════════
   UTILS
═══════════════════════════════════════════════ */
const fmtDT = d => new Date(d).toLocaleString("th-TH",{dateStyle:"short",timeStyle:"short"});
const today  = () => new Date().toISOString().slice(0,10);
const uid    = () => Date.now().toString(36)+Math.random().toString(36).slice(2,5);
const isoNow = () => new Date().toISOString();

/* ═══════════════════════════════════════════════
   DESIGN ATOMS
═══════════════════════════════════════════════ */
const IS = { padding:"9px 10px", border:"0.5px solid var(--color-border-secondary)", borderRadius:8, fontSize:14, background:"var(--color-background-secondary)", color:"var(--color-text-primary)", width:"100%", boxSizing:"border-box", fontFamily:"var(--font-sans)" };

const Card  = ({c,style}) => <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"12px 14px",...style}}>{c}</div>;
const Lbl   = ({t,req}) => <label style={{fontSize:12,color:"var(--color-text-secondary)",fontWeight:500,display:"block",marginBottom:4}}>{t}{req&&<span style={{color:"#EF4444",marginLeft:2}}>*</span>}</label>;
const Hr    = () => <div style={{height:"0.5px",background:"var(--color-border-tertiary)",margin:"10px 0"}}/>;
const SLbl  = ({t}) => <div style={{fontSize:11,color:"var(--color-text-secondary)",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:8}}>{t}</div>;

const Badge = ({sev,sm}) => {
  const c=SEV[sev]||SEV.normal;
  return <span style={{display:"inline-flex",alignItems:"center",gap:4,padding:sm?"2px 8px":"4px 11px",borderRadius:20,background:c.bg,border:`1px solid ${c.border}`,color:c.color,fontSize:sm?11:12,fontWeight:700,whiteSpace:"nowrap"}}>
    <span style={{width:6,height:6,borderRadius:3,background:c.dot,flexShrink:0}}/>{c.label}
  </span>;
};

const Tabs = ({tabs,val,set,sm}) => (
  <div style={{display:"flex",gap:4,marginBottom:14,background:"var(--color-background-secondary)",borderRadius:10,padding:4}}>
    {tabs.map(([k,l])=>(
      <button key={k} onClick={()=>set(k)} style={{flex:1,padding:sm?"6px 2px":"7px 0",borderRadius:8,border:"none",background:val===k?"var(--color-background-primary)":"transparent",color:val===k?"var(--color-text-primary)":"var(--color-text-secondary)",fontSize:sm?11:13,fontWeight:val===k?700:400,cursor:"pointer",boxShadow:val===k?"0 0 0 0.5px var(--color-border-tertiary)":undefined,whiteSpace:"nowrap"}}>
        {l}
      </button>
    ))}
  </div>
);

const Bar = ({qty,max,unit,compact}) => {
  const p=qty/max, c=p<0.3?"#EF4444":p<0.55?"#F59E0B":"#10B981";
  return <div>
    <div style={{height:compact?4:5,background:"var(--color-background-secondary)",borderRadius:4,overflow:"hidden"}}>
      <div style={{height:"100%",width:`${Math.min(100,Math.round(p*100))}%`,background:c,borderRadius:4,transition:"width .35s"}}/>
    </div>
    {!compact&&<div style={{display:"flex",justifyContent:"space-between",marginTop:4}}>
      <span style={{fontSize:12,fontWeight:700,color:c}}>{qty} {unit}</span>
      <span style={{fontSize:11,color:"var(--color-text-secondary)"}}>/{max} · {Math.round(p*100)}%{p<0.3?" ⚠":""}</span>
    </div>}
  </div>;
};

/* Manual severity 3-button picker */
const SevPicker = ({val,set}) => (
  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
    {Object.entries(SEV).map(([k,c])=>(
      <button key={k} onClick={()=>set(k)} style={{padding:"11px 0",borderRadius:12,cursor:"pointer",border:`2px solid ${val===k?c.dot:c.border}`,background:val===k?c.bg:"var(--color-background-primary)",transform:val===k?"scale(1.04)":"scale(1)",transition:"all .15s",boxShadow:val===k?`0 0 0 2px ${c.dot}33`:undefined,outline:"none"}}>
        <div style={{fontSize:22,marginBottom:2}}>{c.emoji}</div>
        <div style={{fontSize:13,fontWeight:val===k?700:500,color:val===k?c.color:"var(--color-text-primary)"}}>{c.label}</div>
      </button>
    ))}
  </div>
);

/* Record card */
const RCard = ({r,onDel}) => {
  const c=SEV[r.severity]||SEV.normal, sh=SHIFTS.find(s=>s.k===r.shift);
  return (
    <div style={{background:"var(--color-background-primary)",border:`1px solid ${c.border}`,borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${c.dot}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:13,fontWeight:700}}>เตียง {r.bed} · {r.name}</div>
          {(r.sys||r.pulse)&&<div style={{fontSize:12,color:"var(--color-text-secondary)",marginTop:1}}>
            {r.sys&&r.dia?`BP ${r.sys}/${r.dia} · `:""}
            {r.pulse?`ชีพจร ${r.pulse} bpm`:""}
          </div>}
          {r.notes&&<div style={{fontSize:12,color:"var(--color-text-secondary)",marginTop:2,fontStyle:"italic",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r.notes}</div>}
          <div style={{fontSize:11,color:"var(--color-text-secondary)",marginTop:4,display:"flex",gap:8,flexWrap:"wrap"}}>
            <span>{r.recorder}</span>
            {sh&&<span style={{background:"var(--color-background-secondary)",padding:"1px 5px",borderRadius:4}}>กะ{sh.l}</span>}
            <span style={{marginLeft:"auto"}}>{fmtDT(r.ts)}</span>
          </div>
        </div>
        <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:5,flexShrink:0}}>
          <Badge sev={r.severity} sm/>
          {onDel&&<button onClick={()=>onDel(r.id)} style={{padding:"3px 8px",background:"#FFF1F2",border:"1px solid #FECDD3",borderRadius:6,color:"#991B1B",fontSize:11,cursor:"pointer"}}>ลบ</button>}
        </div>
      </div>
    </div>
  );
};

/* ═══════════════════════════════════════════════
   DASHBOARD
═══════════════════════════════════════════════ */
function Dashboard({records,patients,stock,setPage,setSelBed}){
  const occupied=Object.values(patients).filter(p=>p?.name).length;
  const lb={};
  records.forEach(r=>{ const b=+r.bed; if(!lb[b]||r.ts>lb[b].ts) lb[b]=r; });
  const urg=Object.values(lb).filter(r=>r.severity==="urgent");
  const wat=Object.values(lb).filter(r=>r.severity==="watch");
  const low=stock.filter(s=>s.qty/s.max<0.3);
  const tRecs=records.filter(r=>r.ts.slice(0,10)===today());
  const recent=[...records].sort((a,b)=>b.ts.localeCompare(a.ts)).slice(0,6);

  const bars=Array.from({length:7},(_,i)=>{
    const d=new Date(); d.setDate(d.getDate()-(6-i));
    const key=d.toISOString().slice(0,10);
    const rs=records.filter(r=>r.ts.slice(0,10)===key);
    return{lbl:d.getDate(),u:rs.filter(r=>r.severity==="urgent").length,w:rs.filter(r=>r.severity==="watch").length,n:rs.filter(r=>r.severity==="normal").length,t:rs.length};
  });
  const mx=Math.max(1,...bars.map(b=>b.t));

  return (
    <div style={{padding:"14px 14px 0"}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
        {[{n:occupied,l:"เตียงที่ใช้งาน",s:`จาก 40 เตียง`,c:TEAL},{n:urg.length,l:"วิกฤต",s:"ต้องดูแลด่วน",c:"#DC2626"},{n:wat.length,l:"เฝ้าระวัง",s:"ต้องสังเกตอาการ",c:"#D97706"},{n:tRecs.length,l:"บันทึกวันนี้",s:"รายการ",c:"#7C3AED"}].map((x,i)=>(
          <div key={i} style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"12px 14px"}}>
            <div style={{fontSize:30,fontWeight:800,color:x.c,lineHeight:1}}>{x.n}</div>
            <div style={{fontSize:13,fontWeight:600,marginTop:4}}>{x.l}</div>
            <div style={{fontSize:11,color:"var(--color-text-secondary)",marginTop:1}}>{x.s}</div>
          </div>
        ))}
      </div>

      {/* 7-day chart */}
      <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"12px 14px",marginBottom:12}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <span style={{fontSize:13,fontWeight:600}}>บันทึก 7 วันที่ผ่านมา</span>
          <div style={{display:"flex",gap:8,fontSize:10}}>
            {[["#EF4444","วิกฤต"],["#F59E0B","เฝ้าระวัง"],["#10B981","ปกติ"]].map(([c,l])=>(
              <span key={l} style={{display:"flex",alignItems:"center",gap:3,color:"var(--color-text-secondary)"}}>
                <span style={{width:7,height:7,borderRadius:3,background:c,display:"inline-block"}}/>{l}
              </span>
            ))}
          </div>
        </div>
        <div style={{display:"flex",gap:4,height:56,alignItems:"flex-end"}}>
          {bars.map((b,i)=>(
            <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"stretch",gap:1,height:"100%",justifyContent:"flex-end"}}>
              {b.u>0&&<div style={{background:"#EF4444",borderRadius:"2px 2px 0 0",height:`${Math.round(b.u/mx*100)}%`,minHeight:3}}/>}
              {b.w>0&&<div style={{background:"#F59E0B",height:`${Math.round(b.w/mx*100)}%`,minHeight:3}}/>}
              {b.n>0&&<div style={{background:"#10B981",borderRadius:b.u===0&&b.w===0?"2px 2px 0 0":undefined,height:`${Math.round(b.n/mx*100)}%`,minHeight:3}}/>}
            </div>
          ))}
        </div>
        <div style={{display:"flex",gap:4,marginTop:4}}>
          {bars.map((b,i)=><div key={i} style={{flex:1,textAlign:"center",fontSize:10,color:"var(--color-text-secondary)"}}>{b.lbl}</div>)}
        </div>
      </div>

      {urg.length>0&&(
        <div style={{marginBottom:12}}>
          <SLbl t={`🔴 วิกฤต — ${urg.length} เตียง`}/>
          {urg.map(r=>(
            <div key={r.id} onClick={()=>{setSelBed(+r.bed);setPage("beds");}} style={{background:"#FFF1F2",border:"1px solid #FECDD3",borderRadius:10,padding:"10px 12px",borderLeft:"3px solid #EF4444",cursor:"pointer",marginBottom:6,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><div style={{fontWeight:700,fontSize:13,color:"#991B1B"}}>เตียง {r.bed} · {r.name}</div><div style={{fontSize:11,color:"#B91C1C",marginTop:1}}>{r.sys?`BP ${r.sys}/${r.dia} · `:""}ชีพจร {r.pulse} · {fmtDT(r.ts)}</div></div>
              <span style={{fontSize:11,color:"#991B1B",flexShrink:0}}>ดูเตียง →</span>
            </div>
          ))}
        </div>
      )}

      {low.length>0&&(
        <div style={{marginBottom:12}}>
          <SLbl t={`📦 Stock วิกฤต — ${low.length} รายการ`}/>
          {low.map(s=>(
            <div key={s.id} onClick={()=>setPage("stock")} style={{background:"#FFF1F2",border:"1px solid #FECDD3",borderRadius:10,padding:"10px 12px",cursor:"pointer",marginBottom:6}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                <span style={{fontSize:13,fontWeight:600,color:"#991B1B"}}>{s.name}</span>
                <span style={{fontSize:11,fontWeight:700,color:"#EF4444"}}>{Math.round(s.qty/s.max*100)}%</span>
              </div>
              <Bar qty={s.qty} max={s.max} unit={s.unit} compact/>
            </div>
          ))}
        </div>
      )}

      <SLbl t="บันทึกล่าสุด"/>
      {recent.length===0?<div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"20px",textAlign:"center",color:"var(--color-text-secondary)",fontSize:13}}>ยังไม่มีบันทึก</div>
        :<div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:14}}>{recent.map(r=><RCard key={r.id} r={r}/>)}</div>}
    </div>
  );
}

/* ═══════════════════════════════════════════════
   RECORD PAGE  — manual severity tab picker
═══════════════════════════════════════════════ */
function RecordPage({records,patients,lineToken,showToast,updateRecords}){
  const F0={bed:"",name:"",severity:"normal",sys:"",dia:"",pulse:"",notes:"",recorder:"",shift:"morning"};
  const[tab,setTab]=useState("form");
  const[f,setF]=useState(F0);
  const[filter,setFilter]=useState("");
  const[sevF,setSevF]=useState("all");
  const[dateF,setDateF]=useState(today());

  const upd=(k,v)=>setF(p=>{ const n={...p,[k]:v}; if(k==="bed"&&patients[v]?.name) n.name=patients[v].name; return n; });

  const submit=async()=>{
    if(!f.bed||!f.name||!f.recorder){ showToast("กรุณากรอกข้อมูลให้ครบ","error"); return; }
    const rec={id:uid(),ts:isoNow(),...f};
    await updateRecords([rec,...records]);
    if(lineToken&&f.severity!=="normal"){
      const sh=SHIFTS.find(s=>s.k===f.shift)?.l||"";
      const c=SEV[f.severity];
      line(lineToken,`${c.emoji} ${c.label}: เตียง ${f.bed} (${f.name})\n${f.sys?`BP ${f.sys}/${f.dia} · `:""}${f.pulse?`ชีพจร ${f.pulse} bpm\n`:""}กะ: ${sh} | โดย: ${f.recorder}${f.notes?`\nหมายเหตุ: ${f.notes}`:""}\n${new Date().toLocaleString("th-TH")}`);
    }
    showToast("บันทึกสำเร็จ ✓");
    setF({...F0,recorder:f.recorder,shift:f.shift});
    setTab("list");
  };

  const filtered=useMemo(()=>records.filter(r=>{
    if(sevF!=="all"&&r.severity!==sevF) return false;
    if(dateF&&r.ts.slice(0,10)!==dateF) return false;
    const q=filter.toLowerCase();
    return !q||r.name?.toLowerCase().includes(q)||String(r.bed).includes(q)||r.recorder?.toLowerCase().includes(q);
  }).sort((a,b)=>b.ts.localeCompare(a.ts)),[records,sevF,dateF,filter]);

  return (
    <div style={{padding:"14px 14px 0"}}>
      <Tabs tabs={[["form","บันทึกใหม่"],["list","รายการ"]]} val={tab} set={setTab}/>

      {tab==="form"&&(
        <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"14px",marginBottom:14}}>
          {/* Shift */}
          <div style={{display:"flex",gap:5,marginBottom:12}}>
            {SHIFTS.map(s=>(
              <button key={s.k} onClick={()=>upd("shift",s.k)}
                style={{flex:1,padding:"8px 0",borderRadius:9,border:`1.5px solid ${f.shift===s.k?TEAL:"var(--color-border-tertiary)"}`,background:f.shift===s.k?"#E1F5EE":"var(--color-background-secondary)",color:f.shift===s.k?TEAL:"var(--color-text-secondary)",fontSize:13,fontWeight:f.shift===s.k?700:400,cursor:"pointer"}}>
                {s.l}<div style={{fontSize:9,opacity:.7}}>{s.t}</div>
              </button>
            ))}
          </div>
          <Hr/>

          {/* Bed + name */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
            <div><Lbl t="เตียง" req/>
              <select value={f.bed} onChange={e=>upd("bed",e.target.value)} style={IS}>
                <option value="">-- เลือก --</option>
                {BEDS.map(n=><option key={n} value={n}>{n}{patients[n]?.name?` · ${patients[n].name.split(" ")[0]}`:""}</option>)}
              </select>
            </div>
            <div><Lbl t="ชื่อผู้ป่วย" req/><input value={f.name} onChange={e=>upd("name",e.target.value)} placeholder="ชื่อ-นามสกุล" style={IS}/></div>
          </div>

          {/* ── SEVERITY PICKER ── */}
          <div style={{marginBottom:12}}>
            <Lbl t="ระดับอาการ"/>
            <SevPicker val={f.severity} set={v=>upd("severity",v)}/>
          </div>

          {/* Vitals optional */}
          <div style={{marginBottom:8}}>
            <Lbl t="ค่าวัด (ไม่บังคับ)"/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
              {[["sys","SBP","ความดันบน"],["dia","DBP","ความดันล่าง"],["pulse","bpm","ชีพจร"]].map(([k,ph,sub])=>(
                <div key={k}>
                  <input type="number" value={f[k]} onChange={e=>upd(k,e.target.value)} placeholder={ph} style={{...IS,textAlign:"center"}}/>
                  <div style={{fontSize:9,color:"var(--color-text-secondary)",textAlign:"center",marginTop:2}}>{sub}</div>
                </div>
              ))}
            </div>
          </div>

          <div style={{marginBottom:8}}><Lbl t="อาการ / หมายเหตุ"/><textarea value={f.notes} onChange={e=>upd("notes",e.target.value)} placeholder="พฤติกรรม, อาการผิดปกติ, คำแนะนำ..." style={{...IS,resize:"none",height:58}}/></div>
          <div style={{marginBottom:14}}><Lbl t="ผู้บันทึก" req/><input value={f.recorder} onChange={e=>upd("recorder",e.target.value)} placeholder="ชื่อพยาบาล/NA" style={IS}/></div>

          {/* Preview banner */}
          <div style={{padding:"10px 12px",borderRadius:10,background:SEV[f.severity].bg,border:`1px solid ${SEV[f.severity].border}`,display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontSize:22}}>{SEV[f.severity].emoji}</span>
              <div><div style={{fontSize:11,color:"var(--color-text-secondary)"}}>จะบันทึกเป็น</div><div style={{fontSize:14,fontWeight:700,color:SEV[f.severity].color}}>{SEV[f.severity].label}</div></div>
            </div>
            {lineToken&&f.severity!=="normal"&&<span style={{fontSize:11,color:SEV[f.severity].color,fontWeight:600}}>📲 แจ้ง LINE</span>}
          </div>

          <button onClick={submit} style={{width:"100%",padding:"13px",background:TEAL,color:"#E1F5EE",border:"none",borderRadius:10,fontSize:15,fontWeight:700,cursor:"pointer"}}>บันทึกอาการ</button>
        </div>
      )}

      {tab==="list"&&(
        <>
          <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:10}}>
            <input value={filter} onChange={e=>setFilter(e.target.value)} placeholder="🔍 ค้นหา ชื่อ / เตียง / ผู้บันทึก..." style={IS}/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
              <select value={sevF} onChange={e=>setSevF(e.target.value)} style={IS}>
                <option value="all">ทุกสถานะ</option>
                <option value="urgent">🔴 วิกฤต</option>
                <option value="watch">🟡 เฝ้าระวัง</option>
                <option value="normal">🟢 ปกติ</option>
              </select>
              <input type="date" value={dateF} onChange={e=>setDateF(e.target.value)} style={IS}/>
            </div>
          </div>
          <SLbl t={`${filtered.length} รายการ`}/>
          {filtered.length===0?<div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"20px",textAlign:"center",color:"var(--color-text-secondary)",fontSize:13}}>ไม่พบรายการ</div>
            :<div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:14}}>{filtered.map(r=><RCard key={r.id} r={r}/>)}</div>}
        </>
      )}
    </div>
  );
}

/* ═══════════════════════════════════════════════
   BED GRID
═══════════════════════════════════════════════ */
function BedGrid({records,patients,stock,bedStock,usage,selBed,setSelBed,showToast,updateBedStock,updateStock,updateUsage}){
  const[sevF,setSevF]=useState("all");
  if(selBed!==null) return <BedDetail bed={selBed} records={records} patients={patients} stock={stock} bedStock={bedStock} usage={usage} onBack={()=>setSelBed(null)} showToast={showToast} updateBedStock={updateBedStock} updateStock={updateStock} updateUsage={updateUsage}/>;

  const lb={};
  records.forEach(r=>{ const b=+r.bed; if(!lb[b]||r.ts>lb[b].ts) lb[b]=r; });

  const vis=BEDS.filter(n=>{
    if(sevF==="empty") return !patients[n]?.name;
    if(sevF==="occupied") return !!patients[n]?.name;
    if(["urgent","watch","normal"].includes(sevF)) return lb[n]?.severity===sevF;
    return true;
  });

  return (
    <div style={{padding:"14px 14px 0"}}>
      {/* Sev summary */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:10}}>
        {Object.entries(SEV).map(([k,c])=>(
          <button key={k} onClick={()=>setSevF(sevF===k?"all":k)}
            style={{background:sevF===k?c.bg:"var(--color-background-primary)",border:`1.5px solid ${sevF===k?c.dot:c.border}`,borderRadius:10,padding:"8px 0",cursor:"pointer",textAlign:"center"}}>
            <div style={{fontSize:22,fontWeight:800,color:c.color}}>{Object.values(lb).filter(r=>r.severity===k).length}</div>
            <div style={{fontSize:11,color:c.color}}>{c.emoji} {c.label}</div>
          </button>
        ))}
      </div>
      <div style={{display:"flex",gap:5,marginBottom:12}}>
        {[["all","ทั้งหมด"],["occupied","มีผู้ป่วย"],["empty","ว่าง"]].map(([k,l])=>(
          <button key={k} onClick={()=>setSevF(k)} style={{flex:1,padding:"7px 0",borderRadius:8,border:`1px solid ${sevF===k?TEAL:"var(--color-border-tertiary)"}`,background:sevF===k?"#E1F5EE":"var(--color-background-secondary)",color:sevF===k?TEAL:"var(--color-text-secondary)",fontSize:12,fontWeight:sevF===k?700:400,cursor:"pointer"}}>{l}</button>
        ))}
      </div>

      <SLbl t={`แผนผัง ${vis.length} เตียง`}/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:6,marginBottom:12}}>
        {vis.map(n=>{
          const pt=patients[n],lat=lb[n],sev=lat?.severity||"normal";
          const hasBS=bedStock[n]&&Object.keys(bedStock[n]).length>0;
          return (
            <button key={n} onClick={()=>setSelBed(n)}
              style={{background:lat?SEV[sev].bg:"var(--color-background-primary)",border:`1.5px solid ${lat?SEV[sev].border:"var(--color-border-tertiary)"}`,borderRadius:10,padding:"8px 4px",cursor:"pointer",textAlign:"center",opacity:!pt?.name?.trim()?0.4:1,position:"relative",outline:"none"}}>
              <div style={{fontSize:16,fontWeight:800,color:lat?SEV[sev].color:TEAL,lineHeight:1}}>{n}</div>
              <div style={{fontSize:9,color:"var(--color-text-secondary)",marginTop:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:"100%"}}>{pt?.name?.split(" ")[0]||"ว่าง"}</div>
              {lat&&<div style={{width:6,height:6,borderRadius:3,background:SEV[sev].dot,margin:"3px auto 0"}}/>}
              {hasBS&&<div style={{position:"absolute",top:3,right:3,width:5,height:5,borderRadius:3,background:"#7C3AED"}}/>}
            </button>
          );
        })}
      </div>

      {/* Legend */}
      <div style={{display:"flex",gap:10,justifyContent:"center",flexWrap:"wrap",marginBottom:14,fontSize:10,color:"var(--color-text-secondary)"}}>
        {[["#EF4444","วิกฤต"],["#F59E0B","เฝ้าระวัง"],["#10B981","ปกติ"],["#7C3AED","มี Stock เตียง"]].map(([c,l])=>(
          <span key={l} style={{display:"flex",alignItems:"center",gap:3}}><span style={{width:7,height:7,borderRadius:3,background:c,flexShrink:0}}/>{l}</span>
        ))}
      </div>
    </div>
  );
}

/* ─── Bed Detail ─── */
function BedDetail({bed,records,patients,stock,bedStock,usage,onBack,showToast,updateBedStock,updateStock,updateUsage}){
  const[tab,setTab]=useState("status");
  const[who,setWho]=useState("");
  const[addF,setAddF]=useState({itemId:"",qty:"1"});
  const[useF,setUseF]=useState({itemId:"",qty:"1",note:""});

  const pt=patients[bed];
  const bedRecs=[...records.filter(r=>+r.bed===bed)].sort((a,b)=>b.ts.localeCompare(a.ts));
  const lat=bedRecs[0];
  const bs=bedStock[bed]||{};
  const bsItems=Object.values(bs);
  const sevCnt={urgent:bedRecs.filter(r=>r.severity==="urgent").length,watch:bedRecs.filter(r=>r.severity==="watch").length,normal:bedRecs.filter(r=>r.severity==="normal").length};
  const bedLog=[...usage.filter(u=>+u.bed===bed)].sort((a,b)=>b.ts.localeCompare(a.ts));

  const requireWho=()=>{ if(!who.trim()){ showToast("ระบุชื่อผู้ดำเนินการก่อน","error"); return false; } return true; };

  const addToBed=async()=>{
    if(!requireWho()||!addF.itemId||+addF.qty<=0){ showToast("กรุณากรอกให้ครบ","error"); return; }
    const ci=stock.find(s=>s.id===addF.itemId); if(!ci) return;
    const q=parseInt(addF.qty);
    if(q>ci.qty){ showToast(`คลังเหลือแค่ ${ci.qty} ${ci.unit}`,"error"); return; }
    await updateStock(stock.map(s=>s.id===addF.itemId?{...s,qty:s.qty-q}:s));
    const cur=bs[addF.itemId]||{itemId:addF.itemId,name:ci.name,unit:ci.unit,qty:0};
    await updateBedStock({...bedStock,[bed]:{...bs,[addF.itemId]:{...cur,qty:cur.qty+q}}});
    await updateUsage([{id:uid(),ts:isoNow(),type:"bedIn",bed,itemId:ci.id,itemName:ci.name,unit:ci.unit,qty:q,recorder:who},...usage]);
    showToast(`เบิก ${ci.name} ×${q} → เตียง ${bed} ✓`);
    setAddF({itemId:"",qty:"1"});
  };

  const useFromBed=async()=>{
    if(!requireWho()||!useF.itemId||+useF.qty<=0){ showToast("กรุณากรอกให้ครบ","error"); return; }
    const item=bs[useF.itemId]; if(!item) return;
    const q=parseInt(useF.qty);
    if(q>item.qty){ showToast(`เหลือแค่ ${item.qty} ${item.unit}`,"error"); return; }
    const newQty=item.qty-q;
    const nb={...bedStock,[bed]:{...bs}};
    if(newQty<=0) delete nb[bed][useF.itemId]; else nb[bed][useF.itemId]={...item,qty:newQty};
    await updateBedStock(nb);
    await updateUsage([{id:uid(),ts:isoNow(),type:"bedUse",bed,itemId:item.itemId,itemName:item.name,unit:item.unit,qty:q,note:useF.note,recorder:who},...usage]);
    showToast(`ใช้ ${item.name} ×${q} ✓`);
    setUseF({itemId:"",qty:"1",note:""});
  };

  const returnToCenter=async(itemId)=>{
    if(!requireWho()) return;
    const item=bs[itemId]; if(!item||item.qty<=0) return;
    await updateStock(stock.map(s=>s.id===itemId?{...s,qty:Math.min(s.max,s.qty+item.qty)}:s));
    const nb={...bedStock,[bed]:{...bs}}; delete nb[bed][itemId];
    await updateBedStock(nb);
    await updateUsage([{id:uid(),ts:isoNow(),type:"bedReturn",bed,itemId,itemName:item.name,unit:item.unit,qty:item.qty,recorder:who},...usage]);
    showToast(`คืน ${item.name} ${item.qty} ${item.unit} กลับคลัง ✓`);
  };

  return (
    <div style={{padding:"14px 14px 0"}}>
      <button onClick={onBack} style={{display:"flex",alignItems:"center",gap:6,background:"none",border:"none",cursor:"pointer",color:"var(--color-text-secondary)",fontSize:13,padding:"0 0 12px",marginLeft:-2}}>← กลับแผนผัง</button>

      {/* Header */}
      <div style={{background:lat?SEV[lat.severity].bg:"var(--color-background-primary)",border:`1.5px solid ${lat?SEV[lat.severity].border:"var(--color-border-tertiary)"}`,borderRadius:12,padding:"14px",marginBottom:12}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div><div style={{fontSize:10,color:"var(--color-text-secondary)",textTransform:"uppercase",letterSpacing:".05em"}}>เตียง</div><div style={{fontSize:48,fontWeight:800,color:TEAL,lineHeight:1}}>{bed}</div></div>
          {pt?<div style={{textAlign:"right"}}><div style={{fontSize:16,fontWeight:700}}>{pt.name}</div>{pt.age&&<div style={{fontSize:12,color:"var(--color-text-secondary)"}}>อายุ {pt.age} ปี</div>}{pt.condition&&<div style={{fontSize:12,color:"var(--color-text-secondary)"}}>{pt.condition}</div>}</div>
            :<div style={{fontSize:12,color:"var(--color-text-secondary)"}}>ไม่มีข้อมูลผู้ป่วย</div>}
        </div>
        {lat&&<><Hr/><div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{fontSize:11,color:"var(--color-text-secondary)"}}>ล่าสุด · {fmtDT(lat.ts)}</div><div style={{fontSize:16,fontWeight:700,marginTop:1}}>{lat.sys&&lat.dia?`BP ${lat.sys}/${lat.dia} · `:""}ชีพจร {lat.pulse||"—"}</div><div style={{fontSize:12,color:"var(--color-text-secondary)"}}>{lat.recorder}</div></div>
          <Badge sev={lat.severity}/>
        </div></>}
      </div>

      <Tabs tabs={[["status","อาการ"],["stock","Stock เตียง"],["log","ประวัติ"]]} val={tab} set={setTab}/>

      {tab==="status"&&(
        <>
          <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"12px 14px",marginBottom:12}}>
            <div style={{fontSize:13,fontWeight:600,marginBottom:10}}>สถิติ ({bedRecs.length} บันทึก)</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
              {Object.entries(SEV).map(([k,c])=>(
                <div key={k} style={{textAlign:"center",padding:"10px 4px",background:c.bg,borderRadius:10}}>
                  <div style={{fontSize:24,fontWeight:800,color:c.color}}>{sevCnt[k]}</div>
                  <div style={{fontSize:10,color:c.color}}>{c.emoji} {c.label}</div>
                </div>
              ))}
            </div>
          </div>
          <SLbl t="ประวัติอาการ"/>
          {bedRecs.length===0?<div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"20px",textAlign:"center",color:"var(--color-text-secondary)",fontSize:13,marginBottom:12}}>ยังไม่มีบันทึก</div>
            :<div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:14}}>{bedRecs.map(r=><RCard key={r.id} r={r}/>)}</div>}
        </>
      )}

      {tab==="stock"&&(
        <>
          {/* Who */}
          <div style={{marginBottom:12}}>
            <Lbl t="ชื่อผู้ดำเนินการ" req/>
            <input value={who} onChange={e=>setWho(e.target.value)} placeholder="ชื่อพยาบาล/NA" style={IS}/>
          </div>

          {/* Current bed stock */}
          <SLbl t={`Stock เตียง ${bed} (${bsItems.length} รายการ)`}/>
          {bsItems.length===0
            ?<div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"16px",textAlign:"center",color:"var(--color-text-secondary)",fontSize:13,marginBottom:12}}>ยังไม่มี Stock ประจำเตียงนี้</div>
            :<div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:12}}>
              {bsItems.map(item=>(
                <div key={item.itemId} style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:10,padding:"10px 12px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}>
                    <div style={{flex:1}}>
                      <div style={{fontSize:13,fontWeight:600}}>{item.name}</div>
                      <div style={{fontSize:24,fontWeight:800,color:item.qty<=2?"#EF4444":TEAL,lineHeight:1.2}}>{item.qty}<span style={{fontSize:12,fontWeight:400,color:"var(--color-text-secondary)",marginLeft:4}}>{item.unit}</span></div>
                    </div>
                    <div style={{display:"flex",flexDirection:"column",gap:5}}>
                      <button onClick={()=>setUseF({itemId:item.itemId,qty:"1",note:""})}
                        style={{padding:"6px 12px",background:"#FFF1F2",border:"1px solid #FECDD3",borderRadius:8,color:"#991B1B",fontSize:12,cursor:"pointer",fontWeight:600}}>ใช้ −</button>
                      <button onClick={()=>returnToCenter(item.itemId)}
                        style={{padding:"6px 12px",background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-secondary)",borderRadius:8,color:"var(--color-text-secondary)",fontSize:12,cursor:"pointer"}}>คืนคลัง</button>
                    </div>
                  </div>
                  {/* inline use form */}
                  {useF.itemId===item.itemId&&(
                    <div style={{marginTop:10,padding:"10px",background:"#FFF8F8",borderRadius:8,border:"1px solid #FECDD3"}}>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:6}}>
                        <div><Lbl t="จำนวนที่ใช้"/><input type="number" min="1" value={useF.qty} onChange={e=>setUseF(p=>({...p,qty:e.target.value}))} style={IS}/></div>
                        <div><Lbl t="หมายเหตุ"/><input value={useF.note} onChange={e=>setUseF(p=>({...p,note:e.target.value}))} placeholder="เหตุที่ใช้..." style={IS}/></div>
                      </div>
                      <div style={{display:"flex",gap:6}}>
                        <button onClick={useFromBed} style={{flex:1,padding:"9px",background:"#991B1B",color:"white",border:"none",borderRadius:8,fontSize:13,fontWeight:700,cursor:"pointer"}}>ยืนยัน</button>
                        <button onClick={()=>setUseF({itemId:"",qty:"1",note:""})} style={{padding:"9px 14px",background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-secondary)",borderRadius:8,fontSize:13,cursor:"pointer"}}>ยกเลิก</button>
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>}

          {/* Withdraw from center to bed */}
          <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"14px",marginBottom:14}}>
            <div style={{fontSize:13,fontWeight:600,color:TEAL,marginBottom:10}}>เบิกจากคลังเข้าเตียง</div>
            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:8,marginBottom:8}}>
              <div><Lbl t="รายการ"/>
                <select value={addF.itemId} onChange={e=>setAddF(p=>({...p,itemId:e.target.value}))} style={IS}>
                  <option value="">-- เลือก --</option>
                  {stock.map(s=><option key={s.id} value={s.id}>{s.name} (คลัง:{s.qty}){s.qty/s.max<0.3?" ⚠":""}</option>)}
                </select>
              </div>
              <div><Lbl t="จำนวน"/><input type="number" min="1" value={addF.qty} onChange={e=>setAddF(p=>({...p,qty:e.target.value}))} style={IS}/></div>
            </div>
            {addF.itemId&&(()=>{ const s=stock.find(x=>x.id===addF.itemId); return s&&<div style={{marginBottom:8}}><Bar qty={s.qty} max={s.max} unit={s.unit}/></div>; })()}
            <button onClick={addToBed} style={{width:"100%",padding:"11px",background:TEAL,color:"#E1F5EE",border:"none",borderRadius:8,fontSize:14,fontWeight:700,cursor:"pointer"}}>เบิกเข้าเตียง +</button>
          </div>
        </>
      )}

      {tab==="log"&&(
        <div style={{marginBottom:14}}>
          <SLbl t={`ประวัติ Stock เตียง ${bed} (${bedLog.length} รายการ)`}/>
          {bedLog.length===0?<div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"20px",textAlign:"center",color:"var(--color-text-secondary)",fontSize:13}}>ยังไม่มีประวัติ</div>
            :<div style={{display:"flex",flexDirection:"column",gap:6}}>
              {bedLog.map(u=>{
                const isIn=u.type==="bedIn",isRet=u.type==="bedReturn";
                return (
                  <div key={u.id} style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${isIn?"#10B981":isRet?"#7C3AED":"#EF4444"}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <div>
                        <div style={{fontSize:13,fontWeight:600}}>{u.itemName}</div>
                        <div style={{fontSize:12,color:"var(--color-text-secondary)"}}>{isIn?"เบิกเข้า":isRet?"คืนคลัง":"ใช้แล้ว"} · {u.recorder}</div>
                        {u.note&&<div style={{fontSize:11,color:"var(--color-text-secondary)",fontStyle:"italic"}}>{u.note}</div>}
                        <div style={{fontSize:11,color:"var(--color-text-secondary)"}}>{fmtDT(u.ts)}</div>
                      </div>
                      <div style={{fontSize:18,fontWeight:800,color:isIn?"#10B981":isRet?"#7C3AED":"#EF4444"}}>
                        {isIn?"+":"-"}{u.qty}<span style={{fontSize:11,fontWeight:400,color:"var(--color-text-secondary)"}}> {u.unit}</span>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>}
        </div>
      )}
    </div>
  );
}

/* ═══════════════════════════════════════════════
   STOCK PAGE (center-wide)
═══════════════════════════════════════════════ */
function StockPage({stock,usage,bedStock,patients,lineToken,showToast,updateStock,updateUsage}){
  const[tab,setTab]=useState("overview");
  const[catF,setCatF]=useState("all");
  const[wf,setWf]=useState({bed:"",itemId:"",qty:"1",recorder:""});
  const[rf,setRf]=useState({itemId:"",qty:"",note:""});

  const use7={};
  usage.filter(u=>new Date(u.ts)>new Date(Date.now()-7*864e5)&&u.type==="bedIn")
    .forEach(u=>{ use7[u.itemId]=(use7[u.itemId]||0)+u.qty; });

  /* item → beds that hold it */
  const inBeds=useMemo(()=>{
    const m={};
    Object.entries(bedStock).forEach(([bed,items])=>{
      Object.values(items).forEach(it=>{
        if(!m[it.itemId]) m[it.itemId]=[];
        m[it.itemId].push({bed:+bed,qty:it.qty,unit:it.unit});
      });
    });
    return m;
  },[bedStock]);

  const vis=catF==="all"?stock:stock.filter(s=>s.cat===catF);

  const doWithdraw=async()=>{
    const{bed,itemId,qty,recorder}=wf;
    if(!bed||!itemId||!qty||!recorder){ showToast("กรุณากรอกให้ครบ","error"); return; }
    const item=stock.find(s=>s.id===itemId); if(!item) return;
    const q=parseInt(qty);
    if(q>item.qty){ showToast("Stock ไม่เพียงพอ","error"); return; }
    const ns=stock.map(s=>s.id===itemId?{...s,qty:s.qty-q}:s);
    await updateStock(ns);
    await updateUsage([{id:uid(),ts:isoNow(),type:"withdraw",bed,itemId,itemName:item.name,unit:item.unit,qty:q,recorder},...usage]);
    const upd=ns.find(s=>s.id===itemId);
    if(upd.qty/upd.max<0.3&&lineToken) line(lineToken,`⚠️ Stock ต่ำ: ${item.name}\nเหลือ ${upd.qty}/${upd.max} ${item.unit} (${Math.round(upd.qty/upd.max*100)}%)`);
    showToast(`เบิก ${item.name} ×${q} ✓`);
    setWf({bed:"",itemId:"",qty:"1",recorder:""});
  };

  const doReplenish=async()=>{
    if(!rf.itemId||!rf.qty||+rf.qty<=0){ showToast("กรุณาระบุจำนวน","error"); return; }
    const item=stock.find(s=>s.id===rf.itemId); if(!item) return;
    await updateStock(stock.map(s=>s.id===rf.itemId?{...s,qty:Math.min(s.max,s.qty+parseInt(rf.qty))}:s));
    showToast(`เพิ่ม ${item.name} +${rf.qty} ✓`);
    setRf({itemId:"",qty:"",note:""});
  };

  return (
    <div style={{padding:"14px 14px 0"}}>
      <Tabs tabs={[["overview","ภาพรวม"],["beds","แยกเตียง"],["withdraw","เบิก"],["replenish","รับของ"]]} val={tab} set={setTab}/>

      {tab==="overview"&&(
        <>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:12}}>
            {[{n:stock.length,l:"รายการ",c:"var(--color-text-primary)"},{n:stock.filter(s=>s.qty/s.max>=0.5).length,l:"ปกติ",c:"#10B981"},{n:stock.filter(s=>s.qty/s.max<0.3).length,l:"ต่ำ ⚠",c:"#EF4444"}].map((x,i)=>(
              <div key={i} style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"10px",textAlign:"center"}}>
                <div style={{fontSize:24,fontWeight:800,color:x.c}}>{x.n}</div>
                <div style={{fontSize:11,color:"var(--color-text-secondary)",marginTop:1}}>{x.l}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:5,marginBottom:10,overflowX:"auto",paddingBottom:2}}>
            {["all",...CATS].map(c=>(
              <button key={c} onClick={()=>setCatF(c)}
                style={{flexShrink:0,padding:"6px 12px",borderRadius:20,border:`1px solid ${catF===c?TEAL:"var(--color-border-tertiary)"}`,background:catF===c?"#E1F5EE":"var(--color-background-secondary)",color:catF===c?TEAL:"var(--color-text-secondary)",fontSize:12,fontWeight:catF===c?700:400,cursor:"pointer"}}>
                {c==="all"?"ทั้งหมด":c}
              </button>
            ))}
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:14}}>
            {vis.map(s=>{
              const r7=use7[s.id]||0, dl=r7>0?Math.floor(s.qty/(r7/7)):null;
              const ib=inBeds[s.id];
              return (
                <div key={s.id} style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"12px 14px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
                    <div>
                      <div style={{fontSize:14,fontWeight:600}}>{s.name}</div>
                      <div style={{fontSize:11,color:"var(--color-text-secondary)"}}>{s.cat}
                        {dl!==null&&<span style={{marginLeft:6,color:dl<5?"#EF4444":"var(--color-text-secondary)"}}>~{dl} วัน</span>}
                      </div>
                    </div>
                    {r7>0&&<span style={{fontSize:11,background:"var(--color-background-secondary)",padding:"2px 7px",borderRadius:6,color:"var(--color-text-secondary)"}}>−{r7}/{s.unit}/7วัน</span>}
                  </div>
                  <Bar qty={s.qty} max={s.max} unit={s.unit}/>
                  {ib&&<div style={{marginTop:6,fontSize:11,color:"#7C3AED"}}>
                    📦 {ib.map(b=>`เตียง${b.bed}(${b.qty}${b.unit})`).join(" · ")}
                  </div>}
                </div>
              );
            })}
          </div>
        </>
      )}

      {tab==="beds"&&(
        <div style={{marginBottom:14}}>
          <SLbl t="Stock แยกตามเตียง"/>
          {BEDS.filter(n=>bedStock[n]&&Object.keys(bedStock[n]).length>0).length===0
            ?<div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"20px",textAlign:"center",color:"var(--color-text-secondary)",fontSize:13}}>ยังไม่มีการแจก Stock ประจำเตียง</div>
            :BEDS.filter(n=>bedStock[n]&&Object.keys(bedStock[n]).length>0).map(n=>{
              const items=Object.values(bedStock[n]), pt=patients[n];
              return (
                <div key={n} style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"12px 14px",marginBottom:8}}>
                  <div style={{fontSize:13,fontWeight:700,color:TEAL,marginBottom:8}}>เตียง {n}{pt?.name?` · ${pt.name.split(" ")[0]}`:""}</div>
                  {items.map(item=>(
                    <div key={item.itemId} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
                      <span style={{fontSize:13}}>{item.name}</span>
                      <span style={{fontSize:16,fontWeight:700,color:item.qty<=2?"#EF4444":TEAL}}>{item.qty} <span style={{fontSize:11,fontWeight:400,color:"var(--color-text-secondary)"}}>{item.unit}</span></span>
                    </div>
                  ))}
                </div>
              );
            })}
        </div>
      )}

      {tab==="withdraw"&&(
        <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"14px",marginBottom:14}}>
          <div style={{fontSize:13,fontWeight:600,color:TEAL,marginBottom:12}}>เบิกอุปกรณ์จากคลัง</div>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <div><Lbl t="เตียง" req/>
                <select value={wf.bed} onChange={e=>setWf(p=>({...p,bed:e.target.value}))} style={IS}>
                  <option value="">-- เลือก --</option>
                  {BEDS.map(n=><option key={n} value={n}>{n}{patients[n]?.name?` · ${patients[n].name.split(" ")[0]}`:""}</option>)}
                </select>
              </div>
              <div><Lbl t="จำนวน" req/><input type="number" min="1" value={wf.qty} onChange={e=>setWf(p=>({...p,qty:e.target.value}))} style={IS}/></div>
            </div>
            <div><Lbl t="รายการ" req/>
              <select value={wf.itemId} onChange={e=>setWf(p=>({...p,itemId:e.target.value}))} style={IS}>
                <option value="">-- เลือก --</option>
                {stock.map(s=><option key={s.id} value={s.id}>{s.name} · เหลือ {s.qty} {s.unit}{s.qty/s.max<0.3?" ⚠":""}</option>)}
              </select>
            </div>
            {wf.itemId&&(()=>{ const s=stock.find(x=>x.id===wf.itemId); return s&&<Bar qty={s.qty} max={s.max} unit={s.unit}/>; })()}
            <div><Lbl t="ผู้เบิก" req/><input value={wf.recorder} onChange={e=>setWf(p=>({...p,recorder:e.target.value}))} placeholder="ชื่อพยาบาล/NA" style={IS}/></div>
            <button onClick={doWithdraw} style={{width:"100%",padding:"13px",background:TEAL,color:"#E1F5EE",border:"none",borderRadius:10,fontSize:15,fontWeight:700,cursor:"pointer"}}>ยืนยันการเบิก</button>
          </div>
        </div>
      )}

      {tab==="replenish"&&(
        <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"14px",marginBottom:14}}>
          <div style={{fontSize:13,fontWeight:600,color:TEAL,marginBottom:12}}>รับของเข้าคลัง</div>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            <div><Lbl t="รายการ" req/>
              <select value={rf.itemId} onChange={e=>setRf(p=>({...p,itemId:e.target.value}))} style={IS}>
                <option value="">-- เลือก --</option>
                {stock.map(s=><option key={s.id} value={s.id}>{s.name} · {s.qty}/{s.max} {s.unit}</option>)}
              </select>
            </div>
            {rf.itemId&&(()=>{ const s=stock.find(x=>x.id===rf.itemId); return s&&<Bar qty={s.qty} max={s.max} unit={s.unit}/>; })()}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <div><Lbl t="จำนวน" req/><input type="number" min="1" value={rf.qty} onChange={e=>setRf(p=>({...p,qty:e.target.value}))} style={IS}/></div>
              <div><Lbl t="PO / หมายเหตุ"/><input value={rf.note} onChange={e=>setRf(p=>({...p,note:e.target.value}))} placeholder="เลข PO..." style={IS}/></div>
            </div>
            <button onClick={doReplenish} style={{width:"100%",padding:"13px",background:"#065F46",color:"#E1F5EE",border:"none",borderRadius:10,fontSize:15,fontWeight:700,cursor:"pointer"}}>บันทึกรับของ</button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ═══════════════════════════════════════════════
   ADMIN
═══════════════════════════════════════════════ */
function AdminPage({patients,records,stock,lineToken,gsCfg,showToast,updatePatients,updateRecords,updateStock,updateToken,updateGsCfg,syncAll}){
  const[tab,setTab]=useState("patients");
  const[pf,setPf]=useState({bed:"",name:"",age:"",condition:""});
  const[lTok,setLTok]=useState(lineToken);
  const[showTok,setShowTok]=useState(false);
  const[sending,setSending]=useState(false);
  const[ni,setNi]=useState({name:"",cat:"ป้องกัน",unit:"",max:"50",qty:"0"});
  const[editMax,setEditMax]=useState({});
  const[gs,setGs]=useState(gsCfg||{url:"",sheetId:""});
  const[syncing,setSyncing]=useState(false);
  const[syncRes,setSyncRes]=useState(null);
  const[copied,setCopied]=useState(false);

  const selBed=bed=>{ const e=patients[bed]||{}; setPf({bed,name:e.name||"",age:e.age||"",condition:e.condition||""}); };
  const savePt=async()=>{ if(!pf.bed||!pf.name){showToast("ต้องระบุเตียงและชื่อ","error");return;} await updatePatients({...patients,[pf.bed]:{bed:+pf.bed,...pf}}); showToast("บันทึกแล้ว ✓"); setPf({bed:"",name:"",age:"",condition:""}); };
  const delPt=async bed=>{ if(!confirm(`ลบผู้ป่วยเตียง ${bed}?`))return; const n={...patients};delete n[bed]; await updatePatients(n); showToast("ลบแล้ว"); };
  const delRec=async id=>{ await updateRecords(records.filter(r=>r.id!==id)); showToast("ลบบันทึกแล้ว"); };
  const addSI=async()=>{ if(!ni.name||!ni.unit||!ni.max){showToast("กรอกให้ครบ","error");return;} await updateStock([...stock,{id:"s"+uid(),name:ni.name,cat:ni.cat,unit:ni.unit,qty:+ni.qty||0,max:+ni.max||50}]); showToast("เพิ่มแล้ว ✓"); setNi({name:"",cat:"ป้องกัน",unit:"",max:"50",qty:"0"}); };
  const delSI=async id=>{ if(!confirm("ลบรายการนี้?"))return; await updateStock(stock.filter(s=>s.id!==id)); showToast("ลบแล้ว"); };
  const saveMax=async(id,m)=>{ await updateStock(stock.map(s=>s.id===id?{...s,max:+m}:s)); setEditMax(p=>({...p,[id]:undefined})); showToast("แก้ไขแล้ว ✓"); };
  const saveLTok=async()=>{ await updateToken(lTok); showToast("บันทึก Token ✓"); };
  const testLine=async()=>{ setSending(true); const ok=await line(lTok,`🏥 ทดสอบ CareCenter Pro\nระบบแจ้งเตือนทำงานปกติ ✓\n${new Date().toLocaleString("th-TH")}`); setSending(false); showToast(ok?"ส่งสำเร็จ ✓":"ส่งไม่สำเร็จ",ok?"success":"error"); };
  const saveGs=async()=>{ await updateGsCfg(gs); DS.init(gs.url); showToast("บันทึก config ✓"); };
  const doSync=async()=>{ setSyncing(true); setSyncRes(null); const r=await syncAll(); setSyncing(false); setSyncRes(r); };
  const copyGas=()=>{ navigator.clipboard.writeText(GAS_CODE).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);}); };

  const ptList=Object.values(patients).filter(p=>p?.name).sort((a,b)=>a.bed-b.bed);

  return (
    <div style={{padding:"14px 14px 0"}}>
      <Tabs tabs={[["patients","ผู้ป่วย"],["records","บันทึก"],["stock","Stock"],["line","LINE"],["sheets","Sheets"]]} val={tab} set={setTab} sm/>

      {tab==="patients"&&(<>
        <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"14px",marginBottom:12}}>
          <div style={{fontSize:13,fontWeight:700,color:TEAL,marginBottom:12}}>ลงทะเบียน / แก้ไขผู้ป่วย</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
            <div><Lbl t="เตียง" req/>
              <select value={pf.bed} onChange={e=>selBed(e.target.value)} style={IS}>
                <option value="">-- เลือก --</option>
                {BEDS.map(n=><option key={n} value={n}>{n}{patients[n]?.name?` ✓ ${patients[n].name.split(" ")[0]}`:""}</option>)}
              </select>
            </div>
            <div><Lbl t="อายุ"/><input type="number" value={pf.age} onChange={e=>setPf(p=>({...p,age:e.target.value}))} placeholder="ปี" style={IS}/></div>
          </div>
          <div style={{marginBottom:8}}><Lbl t="ชื่อ-นามสกุล" req/><input value={pf.name} onChange={e=>setPf(p=>({...p,name:e.target.value}))} placeholder="ชื่อ นามสกุล" style={IS}/></div>
          <div style={{marginBottom:12}}><Lbl t="โรคประจำตัว"/><input value={pf.condition} onChange={e=>setPf(p=>({...p,condition:e.target.value}))} placeholder="Hypertension, DM..." style={IS}/></div>
          <button onClick={savePt} style={{width:"100%",padding:"12px",background:TEAL,color:"#E1F5EE",border:"none",borderRadius:10,fontSize:14,fontWeight:700,cursor:"pointer"}}>บันทึก</button>
        </div>
        <SLbl t={`${ptList.length} ผู้ป่วยในระบบ`}/>
        {ptList.length===0?<div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"20px",textAlign:"center",color:"var(--color-text-secondary)",fontSize:13,marginBottom:14}}>ยังไม่มีผู้ป่วย</div>
          :<div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:14}}>{ptList.map(p=>(
            <div key={p.bed} style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:10,padding:"10px 12px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><div style={{fontSize:13,fontWeight:600}}>เตียง {p.bed} · {p.name}</div><div style={{fontSize:12,color:"var(--color-text-secondary)"}}>{p.age?`${p.age} ปี`:""}{p.condition?` · ${p.condition}`:""}</div></div>
              <div style={{display:"flex",gap:5}}>
                <button onClick={()=>selBed(String(p.bed))} style={{padding:"5px 10px",background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-secondary)",borderRadius:6,fontSize:12,cursor:"pointer"}}>แก้ไข</button>
                <button onClick={()=>delPt(p.bed)} style={{padding:"5px 10px",background:"#FFF1F2",border:"1px solid #FECDD3",borderRadius:6,color:"#991B1B",fontSize:12,cursor:"pointer"}}>ลบ</button>
              </div>
            </div>
          ))}</div>}
      </>)}

      {tab==="records"&&(
        <div style={{marginBottom:14}}>
          <div style={{background:"#FFFBEB",border:"1px solid #FCD34D",borderRadius:10,padding:"10px 12px",fontSize:12,color:"#92400E",marginBottom:10}}>⚠️ การลบไม่สามารถกู้คืนได้</div>
          <SLbl t={`${records.length} บันทึก`}/>
          {records.length===0?<div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"20px",textAlign:"center",color:"var(--color-text-secondary)",fontSize:13}}>ไม่มีบันทึก</div>
            :<div style={{display:"flex",flexDirection:"column",gap:6}}>{[...records].sort((a,b)=>b.ts.localeCompare(a.ts)).map(r=><RCard key={r.id} r={r} onDel={delRec}/>)}</div>}
        </div>
      )}

      {tab==="stock"&&(<>
        <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"14px",marginBottom:12}}>
          <div style={{fontSize:13,fontWeight:700,color:TEAL,marginBottom:12}}>เพิ่มรายการ Stock</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
            <div><Lbl t="ชื่อรายการ" req/><input value={ni.name} onChange={e=>setNi(p=>({...p,name:e.target.value}))} placeholder="ชื่ออุปกรณ์" style={IS}/></div>
            <div><Lbl t="หน่วย" req/><input value={ni.unit} onChange={e=>setNi(p=>({...p,unit:e.target.value}))} placeholder="ชิ้น/กล่อง" style={IS}/></div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
            <div><Lbl t="หมวด"/><select value={ni.cat} onChange={e=>setNi(p=>({...p,cat:e.target.value}))} style={IS}>{CATS.map(c=><option key={c}>{c}</option>)}</select></div>
            <div><Lbl t="เริ่มต้น"/><input type="number" value={ni.qty} onChange={e=>setNi(p=>({...p,qty:e.target.value}))} style={IS}/></div>
            <div><Lbl t="สูงสุด"/><input type="number" value={ni.max} onChange={e=>setNi(p=>({...p,max:e.target.value}))} style={IS}/></div>
          </div>
          <button onClick={addSI} style={{width:"100%",padding:"11px",background:TEAL,color:"#E1F5EE",border:"none",borderRadius:8,fontSize:14,fontWeight:700,cursor:"pointer"}}>เพิ่มรายการ</button>
        </div>
        <SLbl t={`${stock.length} รายการ`}/>
        <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:14}}>
          {stock.map(s=>(
            <div key={s.id} style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:10,padding:"10px 12px",display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
              <div style={{flex:1}}>
                <div style={{fontSize:13,fontWeight:600}}>{s.name} <span style={{fontSize:11,fontWeight:400,color:"var(--color-text-secondary)",background:"var(--color-background-secondary)",padding:"1px 5px",borderRadius:4}}>{s.cat}</span></div>
                <div style={{marginTop:6}}><Bar qty={s.qty} max={s.max} unit={s.unit}/></div>
                <div style={{display:"flex",alignItems:"center",gap:6,marginTop:6}}>
                  <span style={{fontSize:11,color:"var(--color-text-secondary)"}}>สูงสุด:</span>
                  <input type="number" value={editMax[s.id]!==undefined?editMax[s.id]:s.max} onChange={e=>setEditMax(p=>({...p,[s.id]:e.target.value}))} style={{...IS,width:60,padding:"3px 6px",fontSize:12}}/>
                  {editMax[s.id]!==undefined&&<button onClick={()=>saveMax(s.id,editMax[s.id])} style={{padding:"3px 8px",background:TEAL,color:"white",borderRadius:5,border:"none",fontSize:11,cursor:"pointer"}}>บันทึก</button>}
                </div>
              </div>
              <button onClick={()=>delSI(s.id)} style={{padding:"5px 10px",background:"#FFF1F2",border:"1px solid #FECDD3",borderRadius:6,color:"#991B1B",fontSize:11,cursor:"pointer",flexShrink:0}}>ลบ</button>
            </div>
          ))}
        </div>
      </>)}

      {tab==="line"&&(
        <div style={{marginBottom:14}}>
          <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"14px",marginBottom:10}}>
            <div style={{fontSize:14,fontWeight:700,color:TEAL,marginBottom:10}}>LINE Notify Token</div>
            <div style={{fontSize:12,color:"var(--color-text-secondary)",marginBottom:12,lineHeight:1.9}}>
              <b>วิธีรับ token:</b><br/>
              1. notify-bot.line.me → My page<br/>
              2. Generate token → เลือก Group ศูนย์<br/><br/>
              <b>แจ้งอัตโนมัติเมื่อ:</b><br/>
              🔴 วิกฤต · 🟡 เฝ้าระวัง · 📦 Stock &lt;30%
            </div>
            <Hr/>
            <div style={{position:"relative",marginBottom:8}}>
              <Lbl t="Access Token"/>
              <input type={showTok?"text":"password"} value={lTok} onChange={e=>setLTok(e.target.value)} placeholder="วาง token ที่นี่..." style={IS}/>
              <button onClick={()=>setShowTok(!showTok)} style={{position:"absolute",right:10,top:28,background:"none",border:"none",cursor:"pointer",fontSize:12,color:"var(--color-text-secondary)"}}>{showTok?"ซ่อน":"แสดง"}</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <button onClick={saveLTok} style={{padding:"11px",background:TEAL,color:"#E1F5EE",border:"none",borderRadius:8,fontSize:13,fontWeight:700,cursor:"pointer"}}>บันทึก</button>
              <button onClick={testLine} disabled={!lTok||sending} style={{padding:"11px",background:lTok?"#F0FDF4":"var(--color-background-secondary)",color:lTok?"#065F46":"var(--color-text-secondary)",border:`1px solid ${lTok?"#6EE7B7":"var(--color-border-tertiary)"}`,borderRadius:8,fontSize:13,fontWeight:600,cursor:lTok?"pointer":"not-allowed"}}>
                {sending?"กำลังส่ง...":"ทดสอบ"}
              </button>
            </div>
          </div>
        </div>
      )}

      {tab==="sheets"&&(
        <div style={{marginBottom:14}}>
          {/* Status */}
          <div style={{padding:"10px 12px",borderRadius:10,background:gs.url?"#ECFDF5":"var(--color-background-secondary)",border:`1px solid ${gs.url?"#6EE7B7":"var(--color-border-tertiary)"}`,marginBottom:12,display:"flex",gap:10,alignItems:"center"}}>
            <span style={{fontSize:22}}>{gs.url?"🟢":"⚪"}</span>
            <div>
              <div style={{fontSize:13,fontWeight:700,color:gs.url?"#065F46":"var(--color-text-secondary)"}}>{gs.url?"เชื่อมต่อ Google Sheets แล้ว":"ยังไม่ได้เชื่อมต่อ"}</div>
              <div style={{fontSize:11,color:"var(--color-text-secondary)"}}>ข้อมูลบันทึก localStorage · Sync ตามต้องการ</div>
            </div>
          </div>

          {/* Config */}
          <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"14px",marginBottom:12}}>
            <div style={{fontSize:13,fontWeight:700,color:TEAL,marginBottom:10}}>ตั้งค่า Google Sheets</div>
            <div style={{marginBottom:8}}><Lbl t="Apps Script Web App URL"/>
              <input value={gs.url} onChange={e=>setGs(p=>({...p,url:e.target.value}))} placeholder="https://script.google.com/macros/s/..." style={IS}/>
            </div>
            <div style={{marginBottom:12}}><Lbl t="Google Sheet ID"/>
              <input value={gs.sheetId} onChange={e=>setGs(p=>({...p,sheetId:e.target.value}))} placeholder="Sheet ID จาก URL" style={IS}/>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <button onClick={saveGs} style={{padding:"11px",background:TEAL,color:"#E1F5EE",border:"none",borderRadius:8,fontSize:13,fontWeight:700,cursor:"pointer"}}>บันทึก Config</button>
              <button onClick={doSync} disabled={!gs.url||syncing} style={{padding:"11px",background:gs.url?"#E1F5EE":"var(--color-background-secondary)",color:gs.url?TEAL:"var(--color-text-secondary)",border:`1px solid ${gs.url?"#6EE7B7":"var(--color-border-tertiary)"}`,borderRadius:8,fontSize:13,fontWeight:600,cursor:gs.url?"pointer":"not-allowed"}}>
                {syncing?"กำลัง Sync...":"Sync ทันที ↑"}
              </button>
            </div>
            {syncRes&&<div style={{marginTop:10,padding:"8px 10px",borderRadius:8,background:syncRes.ok?"#ECFDF5":"#FFF1F2",border:`1px solid ${syncRes.ok?"#6EE7B7":"#FECDD3"}`,fontSize:12,color:syncRes.ok?"#065F46":"#991B1B"}}>
              {syncRes.ok?"✓ Sync สำเร็จแล้ว":`✗ ${syncRes.msg}`}
            </div>}
          </div>

          {/* GAS Template */}
          <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"14px"}}>
            <div style={{fontSize:13,fontWeight:700,marginBottom:4}}>Apps Script Template</div>
            <div style={{fontSize:12,color:"var(--color-text-secondary)",lineHeight:1.8,marginBottom:10}}>
              <b>ขั้นตอน:</b><br/>
              1. สร้าง Google Spreadsheet ใหม่<br/>
              2. Extensions → Apps Script<br/>
              3. วางโค้ด → แก้ SHEET_ID<br/>
              4. Deploy → Web App (Anyone)<br/>
              5. Copy URL → วางด้านบน
            </div>
            <button onClick={copyGas} style={{width:"100%",padding:"10px",background:copied?"#ECFDF5":"var(--color-background-secondary)",color:copied?"#065F46":"var(--color-text-primary)",border:`1px solid ${copied?"#6EE7B7":"var(--color-border-secondary)"}`,borderRadius:8,fontSize:13,fontWeight:600,cursor:"pointer",marginBottom:8}}>
              {copied?"✓ คัดลอกแล้ว":"📋 คัดลอก Apps Script Code"}
            </button>
            <div style={{background:"var(--color-background-secondary)",borderRadius:8,padding:10,fontSize:10,color:"var(--color-text-secondary)",fontFamily:"monospace",maxHeight:140,overflowY:"auto",lineHeight:1.6,whiteSpace:"pre-wrap"}}>{GAS_CODE.slice(0,500)}…</div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ═══════════════════════════════════════════════
   APP SHELL
═══════════════════════════════════════════════ */
function App(){
  const[page,setPage]=useState("dashboard");
  const[records,setRecords]=useState([]);
  const[patients,setPatients]=useState({});
  const[stock,setStock]=useState(STOCK0);
  const[usage,setUsage]=useState([]);
  const[bedStock,setBedStock]=useState({});
  const[lineToken,setLineToken]=useState("");
  const[gsCfg,setGsCfg]=useState({url:"",sheetId:""});
  const[selBed,setSelBed]=useState(null);
  const[ready,setReady]=useState(false);
  const[toast,setToast]=useState(null);

  useEffect(()=>{
    (async()=>{
      const[r,p,s,u,bs,lt,gc]=await Promise.all([
        DS.load("records",[]),DS.load("patients",{}),DS.load("stock",STOCK0),
        DS.load("usage",[]),DS.load("bedstock",{}),DS.load("linetoken",""),DS.load("gsconfig",{url:"",sheetId:""}),
      ]);
      setRecords(r);setPatients(p);setStock(s);setUsage(u);setBedStock(bs);setLineToken(lt||"");setGsCfg(gc||{url:"",sheetId:""});
      if(gc?.url) DS.init(gc.url);
      setReady(true);
    })();
  },[]);

  const showToast=(msg,type="success")=>{ setToast({msg,type}); setTimeout(()=>setToast(null),3200); };
  const upRec =async v=>{ setRecords(v);  await DS.save("records",v);   };
  const upPts =async v=>{ setPatients(v); await DS.save("patients",v);  };
  const upStk =async v=>{ setStock(v);    await DS.save("stock",v);     };
  const upUsg =async v=>{ setUsage(v);    await DS.save("usage",v);     };
  const upBS  =async v=>{ setBedStock(v); await DS.save("bedstock",v);  };
  const upLT  =async v=>{ setLineToken(v);await DS.save("linetoken",v); };
  const upGS  =async v=>{ setGsCfg(v);   await DS.save("gsconfig",v);  DS.init(v.url); };
  const syncAll=async()=>DS.syncAll(["records","patients","stock","usage","bedstock"]);

  const urgCnt=useMemo(()=>{
    const lb={};
    records.forEach(r=>{ const b=+r.bed; if(!lb[b]||r.ts>lb[b].ts) lb[b]=r; });
    return Object.values(lb).filter(r=>r.severity==="urgent").length;
  },[records]);
  const lowCnt=stock.filter(s=>s.qty/s.max<0.3).length;

  const NAV=[
    {k:"dashboard",l:"หน้าหลัก",ic:<svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/></svg>},
    {k:"record",   l:"บันทึก",   ic:<svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 5v14M5 12h14"/></svg>},
    {k:"beds",     l:"เตียง",    ic:<svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><path d="M2 4v16M2 8h18a2 2 0 0 1 2 2v10"/><path d="M2 17h20"/><path d="M6 8v9"/></svg>},
    {k:"stock",    l:"Stock",    ic:<svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/></svg>},
    {k:"admin",    l:"แอดมิน",   ic:<svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="8" r="4"/><path d="M20 21a8 8 0 1 0-16 0"/></svg>},
  ];

  if(!ready) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"var(--color-background-tertiary)"}}>
      <div style={{textAlign:"center"}}>
        <div style={{width:64,height:64,borderRadius:18,background:TEAL,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 14px",fontSize:32}}>🏥</div>
        <div style={{fontSize:18,fontWeight:700}}>CareCenter Pro</div>
        <div style={{fontSize:13,color:"var(--color-text-secondary)",marginTop:5}}>กำลังโหลดข้อมูล...</div>
      </div>
    </div>
  );

  const shared={records,patients,stock,usage,bedStock,lineToken,gsCfg,showToast,updateRecords:upRec,updatePatients:upPts,updateStock:upStk,updateUsage:upUsg,updateBedStock:upBS,updateToken:upLT,updateGsCfg:upGS,syncAll,setPage,selBed,setSelBed};

  return (
    <div style={{fontFamily:"var(--font-sans)",background:"var(--color-background-tertiary)",minHeight:"100vh",maxWidth:480,margin:"0 auto",display:"flex",flexDirection:"column"}}>
      {/* Header */}
      <div style={{background:TEAL,color:"#E1F5EE",padding:"10px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0,position:"sticky",top:0,zIndex:10}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:34,height:34,borderRadius:10,background:"rgba(255,255,255,0.15)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>🏥</div>
          <div>
            <div style={{fontSize:10,opacity:.65,letterSpacing:".08em",textTransform:"uppercase",lineHeight:1}}>CareCenter Pro</div>
            <div style={{fontSize:15,fontWeight:700,lineHeight:1.3}}>{NAV.find(n=>n.k===page)?.l}</div>
          </div>
        </div>
        <div style={{display:"flex",gap:5,alignItems:"center"}}>
          {urgCnt>0&&<div style={{background:"#EF4444",color:"white",borderRadius:20,padding:"3px 9px",fontSize:12,fontWeight:700}}>🔴 {urgCnt}</div>}
          {lowCnt>0&&<div style={{background:"#D97706",color:"white",borderRadius:20,padding:"3px 8px",fontSize:12,fontWeight:700}}>📦 {lowCnt}</div>}
          {gsCfg?.url&&<div style={{background:"rgba(255,255,255,0.18)",borderRadius:20,padding:"3px 8px",fontSize:11,fontWeight:600}}>Sheets ✓</div>}
        </div>
      </div>

      {/* Content */}
      <div style={{flex:1,overflowY:"auto",paddingBottom:72}}>
        {page==="dashboard"&&<Dashboard   records={records} patients={patients} stock={stock} setPage={setPage} setSelBed={setSelBed}/>}
        {page==="record"   &&<RecordPage  records={records} patients={patients} lineToken={lineToken} showToast={showToast} updateRecords={upRec}/>}
        {page==="beds"     &&<BedGrid     records={records} patients={patients} stock={stock} bedStock={bedStock} usage={usage} selBed={selBed} setSelBed={setSelBed} showToast={showToast} updateBedStock={upBS} updateStock={upStk} updateUsage={upUsg}/>}
        {page==="stock"    &&<StockPage   stock={stock} usage={usage} bedStock={bedStock} patients={patients} lineToken={lineToken} showToast={showToast} updateStock={upStk} updateUsage={upUsg}/>}
        {page==="admin"    &&<AdminPage   patients={patients} records={records} stock={stock} lineToken={lineToken} gsCfg={gsCfg} showToast={showToast} updatePatients={upPts} updateRecords={upRec} updateStock={upStk} updateToken={upLT} updateGsCfg={upGS} syncAll={syncAll}/>}
      </div>

      {/* Bottom nav */}
      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:480,background:"var(--color-background-primary)",borderTop:"0.5px solid var(--color-border-tertiary)",display:"flex",zIndex:20}}>
        {NAV.map(n=>(
          <button key={n.k} onClick={()=>{ setPage(n.k); if(n.k!=="beds") setSelBed(null); }}
            style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",padding:"8px 0 6px",gap:3,background:"none",border:"none",cursor:"pointer",color:page===n.k?TEAL:"var(--color-text-secondary)",position:"relative",transition:"color .15s",outline:"none"}}>
            {n.ic}
            <span style={{fontSize:10,fontWeight:page===n.k?700:400}}>{n.l}</span>
            {n.k==="dashboard"&&urgCnt>0&&<span style={{position:"absolute",top:5,right:"calc(50% - 14px)",width:7,height:7,borderRadius:4,background:"#EF4444"}}/>}
          </button>
        ))}
      </div>

      {/* Toast */}
      {toast&&(
        <div style={{position:"fixed",bottom:82,left:"50%",transform:"translateX(-50%)",background:toast.type==="error"?"#991B1B":"#065F46",color:"white",padding:"11px 22px",borderRadius:24,fontSize:13,fontWeight:600,zIndex:100,whiteSpace:"nowrap",maxWidth:"92%",textAlign:"center",boxShadow:"0 4px 18px rgba(0,0,0,.2)"}}>
          {toast.msg}
        </div>
      )}
    </div>
     
  );
}// ═══════════════════════════════════════════════
// RENDER TO DOM
// ═══════════════════════════════════════════════
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
