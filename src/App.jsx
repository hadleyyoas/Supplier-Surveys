import React, { useState, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Palette ──────────────────────────────────────────────────────────────────
const C = {
  navy:"#0F2340",navyMid:"#1A3A5C",navyLight:"#2A5280",
  sky:"#3B8FD4",skyLight:"#EBF4FC",
  mint:"#2DBD8E",mintLight:"#E6F8F2",
  amber:"#F59E0B",amberLight:"#FEF3C7",
  rose:"#E74C3C",roseLight:"#FDEDEC",
  purple:"#7C3AED",purpleLight:"#F3EEFF",
  slate:"#64748B",slateLight:"#F1F5F9",
  white:"#FFFFFF",border:"#CBD5E1",text:"#1E293B",textMuted:"#64748B",
};

const SENDER_EMAIL = "hadley.yoas@kellyservices.com";

// ─── localStorage persistence ────────────────────────────────────────────────
const LS_KEY = "supplierrate_v1";

function lsLoad(){
  try{
    const raw = localStorage.getItem(LS_KEY);
    return raw ? JSON.parse(raw) : {projects:[],projectData:{}};
  }catch{ return {projects:[],projectData:{}}; }
}

function lsSave(data){
  try{ localStorage.setItem(LS_KEY,JSON.stringify(data)); }catch(e){ console.warn("localStorage save failed",e); }
}

// ─── Utils ────────────────────────────────────────────────────────────────────
function avg(arr){return arr.length?Math.round(arr.reduce((a,b)=>a+b,0)/arr.length):0;}
function stdDev(arr){
  if(arr.length<2)return 0;
  const m=arr.reduce((a,b)=>a+b,0)/arr.length;
  return Math.sqrt(arr.reduce((s,v)=>s+(v-m)**2,0)/(arr.length-1));
}
function isOutlier(val,arr){
  if(arr.length<3)return false;
  const m=arr.reduce((a,b)=>a+b,0)/arr.length;
  const sd=stdDev(arr);
  return sd>0&&Math.abs(val-m)>1.5*sd;
}

// Normalize strings for fuzzy matching: lowercase, trim, collapse spaces
function norm(s){ return String(s||"").toLowerCase().replace(/\s+/g," ").trim(); }

// Fuzzy location match: "Bristol, UK" matches "Bristol" etc.
function locMatch(a,b){
  const na=norm(a), nb=norm(b);
  if(na===nb) return true;
  // Check if either contains the other (handles "Bristol" vs "Bristol, UK")
  return na.includes(nb)||nb.includes(na);
}

// Fuzzy title match: strip embedded level suffixes for comparison
function stripLevel(t){ return norm(t).replace(/\s*[-–]\s*(level\s*\d+|\d+\s*entry|\d+\s*intermediate|\d+\s*senior|\d+\s*expert|\d+\s*principal)/i,"").trim(); }

// Completeness: what % of expected role×location combos did a supplier fill in?
// Uses fuzzy matching so imported data aligns with job list even with minor title/level/location differences.
function completeness(supplierName, responses, jobs, locations){
  const expected = jobs.length * locations.length;
  if(expected===0) return {pct:0,filled:0,expected:0,missing:[]};

  const supRecs = responses.filter(r=>norm(r.supplier)===norm(supplierName));

  const missing = [];
  jobs.forEach(j=>{
    locations.forEach(l=>{
      // A response matches this job×location if:
      // 1. Title matches (exact normalized, OR stripped of embedded level, OR fullTitle matches)
      // 2. Level matches (exact normalized, OR both empty/missing, OR level embedded in response title)
      // 3. Location matches (fuzzy)
      const hit = supRecs.some(r=>{
        const rTitle=norm(r.title); const rLevel=norm(r.level); const rLoc=norm(r.location);
        const jTitle=norm(j.title); const jLevel=norm(j.level);
        const jFullTitle=norm(j.fullTitle||"");

        // Title match: direct, stripped, or fullTitle
        const titleOk = rTitle===jTitle
          || stripLevel(rTitle)===stripLevel(jTitle)
          || (jFullTitle&&(rTitle===jFullTitle||stripLevel(rTitle)===stripLevel(jFullTitle)))
          || jTitle.includes(rTitle)||rTitle.includes(jTitle);

        // Level match: exact, or both blank, or level embedded in the other's title
        const levelOk = rLevel===jLevel
          || (!rLevel&&!jLevel)
          || norm(r.title+" "+r.level).includes(jLevel)
          || norm(j.title+" "+j.level).includes(rLevel)
          || (!rLevel) // supplier left level blank — count it against any level variant of that title
          || (!jLevel);

        // Location match: fuzzy
        const locationOk = locMatch(r.location, l);

        return titleOk && levelOk && locationOk;
      });
      if(!hit) missing.push(`${j.title}${j.level?" – "+j.level:""} – ${l}`);
    });
  });
  const filledCount = expected - missing.length;
  return {pct:Math.round((filledCount/expected)*100), filled:filledCount, expected, missing};
}

// ─── Shared UI ────────────────────────────────────────────────────────────────
function StatusBadge({status}){
  const map={
    responded:{bg:C.mintLight,color:C.mint,label:"Responded"},
    sent:{bg:C.skyLight,color:C.sky,label:"Sent"},
    follow_up:{bg:C.amberLight,color:C.amber,label:"Follow-up"},
    not_sent:{bg:C.slateLight,color:C.slate,label:"Not Sent"},
  };
  const s=map[status]||map.not_sent;
  return <span style={{background:s.bg,color:s.color,padding:"2px 10px",borderRadius:99,fontSize:12,fontWeight:600}}>{s.label}</span>;
}

function CompletenessBar({pct,filled,expected}){
  const color = pct===100?C.mint:pct>=60?C.amber:C.rose;
  return(
    <div style={{display:"flex",alignItems:"center",gap:8,minWidth:140}}>
      <div style={{flex:1,height:7,background:C.slateLight,borderRadius:99,overflow:"hidden"}}>
        <div style={{height:"100%",width:`${pct}%`,background:color,borderRadius:99,transition:"width .3s"}}/>
      </div>
      <span style={{fontSize:11,fontWeight:700,color,whiteSpace:"nowrap"}}>{filled}/{expected}</span>
    </div>
  );
}

function Card({children,style={}}){
  return <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:"20px 24px",...style}}>{children}</div>;
}
function Btn({children,onClick,variant="primary",size="md",style={},disabled=false}){
  const sizes={sm:{padding:"5px 12px",fontSize:12},md:{padding:"8px 18px",fontSize:14},lg:{padding:"11px 24px",fontSize:15}};
  const variants={
    primary:{background:C.navy,color:C.white},
    sky:{background:C.sky,color:C.white},
    mint:{background:C.mint,color:C.white},
    ghost:{background:"transparent",color:C.navy,border:`1px solid ${C.border}`},
    danger:{background:C.rose,color:C.white},
    amber:{background:C.amber,color:C.white},
    purple:{background:C.purple,color:C.white},
  };
  return(
    <button onClick={disabled?undefined:onClick} style={{
      border:"none",borderRadius:8,cursor:disabled?"not-allowed":"pointer",fontWeight:600,
      transition:"opacity .15s",opacity:disabled?.5:1,display:"inline-flex",alignItems:"center",gap:6,
      ...sizes[size],...variants[variant],...style,
    }}>{children}</button>
  );
}
function Input({label,value,onChange,placeholder,type="text",style={}}){
  return(
    <div style={{display:"flex",flexDirection:"column",gap:4,...style}}>
      {label&&<label style={{fontSize:12,fontWeight:600,color:C.textMuted}}>{label}</label>}
      <input type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder}
        style={{border:`1px solid ${C.border}`,borderRadius:7,padding:"7px 11px",fontSize:14,color:C.text,outline:"none",background:C.white}}/>
    </div>
  );
}
function Select({label,value,onChange,options,style={}}){
  return(
    <div style={{display:"flex",flexDirection:"column",gap:4,...style}}>
      {label&&<label style={{fontSize:12,fontWeight:600,color:C.textMuted}}>{label}</label>}
      <select value={value} onChange={e=>onChange(e.target.value)}
        style={{border:`1px solid ${C.border}`,borderRadius:7,padding:"7px 11px",fontSize:14,color:C.text,background:C.white,outline:"none"}}>
        {options.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
    </div>
  );
}
function DropZone({onFile,accept=".xlsx,.xls,.csv",label,sublabel}){
  const [dragging,setDragging]=useState(false);
  const ref=useRef();
  const handle=f=>{if(f)onFile(f);};
  return(
    <div onClick={()=>ref.current.click()}
      onDragOver={e=>{e.preventDefault();setDragging(true);}}
      onDragLeave={()=>setDragging(false)}
      onDrop={e=>{e.preventDefault();setDragging(false);handle(e.dataTransfer.files[0]);}}
      style={{border:`2px dashed ${dragging?C.sky:C.border}`,borderRadius:10,padding:"24px 20px",textAlign:"center",cursor:"pointer",background:dragging?C.skyLight:"#FAFBFC",transition:"all .2s"}}>
      <input ref={ref} type="file" accept={accept} style={{display:"none"}} onChange={e=>handle(e.target.files[0])}/>
      <div style={{fontSize:26,marginBottom:6}}>📂</div>
      <div style={{fontWeight:600,color:C.navy,fontSize:14}}>{label}</div>
      {sublabel&&<div style={{color:C.textMuted,fontSize:12,marginTop:3}}>{sublabel}</div>}
    </div>
  );
}
function Toast({msg}){
  if(!msg)return null;
  const ok=msg.startsWith("✅");
  const info=msg.startsWith("📨");
  const color=ok?C.mint:info?C.sky:C.amber;
  const bg=ok?C.mintLight:info?C.skyLight:C.amberLight;
  return <div style={{padding:"9px 14px",borderRadius:8,fontSize:13,marginTop:10,background:bg,color}}>{msg}</div>;
}

// ─── TEMPLATE BUILDER ────────────────────────────────────────────────────────
function TemplateBuilder({jobs,setJobs,locations,setLocations,clientName}){
  const [newTitle,setNewTitle]=useState("");
  const [newLevel,setNewLevel]=useState("Mid");
  const [newLoc,setNewLoc]=useState("");
  const [importMsg,setImportMsg]=useState("");
  function buildTemplate(cn){
    const c=cn||"[CLIENT NAME]";
    return `Dear Supplier Partner,

${c} is partnering with KellyOCG to conduct a Market Rate Analysis across multiple geographies, and your participation has been requested.

Your response will contribute to a broader benchmarking initiative designed to reflect current local market conditions and competitive pay practices. KellyOCG is managing this effort on behalf of ${c}.

Please follow these guidelines when completing the attached country-specific survey template:

  • Complete all green shaded fields at the top of the worksheet:
      • Supplier Name
      • Point of Contact (Name)
      • Email Address
  • Provide responses only for roles and locations your organization actively supports
  • Submit one competitive hourly pay and/or bill rate for a highly qualified candidate (no ranges)
  • If you are unfamiliar with a job title or geography, leave the row blank
  • Enter all rates in local currency
  • Do not modify job titles, structure, or formatting

Please return your completed survey(s) to:
Ratecards@kellyocg.com

Deadline
All responses must be submitted by:
[DEADLINE DATE] (Close of Business EST)

Questions
If you have any questions or need assistance, please contact ratecards@kellyocg.com or your ${c} program representative.

Thank you for your time and partnership. Your input is critical to ensuring ${c} maintains a competitive and market-aligned contingent workforce program.

Kind regards,
KellyOCG MRA Team`;
  }
  const [preview,setPreview]=useState(()=>buildTemplate(clientName));

  // Re-initialize template when client name changes (only if still showing default)
  const prevClientRef = React.useRef(clientName);
  if(prevClientRef.current!==clientName){
    prevClientRef.current=clientName;
    // Only auto-update if the current text still contains the old client name placeholder
    // so manual edits are preserved
  }

  const levels=["1 Entry","2 Intermediate","3 Senior","4 Expert","5 Principal"];

  function addJob(){
    if(!newTitle.trim())return;
    setJobs(p=>[...p,{id:Date.now(),title:newTitle.trim(),level:newLevel}]);
    setNewTitle("");
  }
  function removeJob(id){setJobs(p=>p.filter(j=>j.id!==id));}
  function addLoc(){
    if(!newLoc.trim())return;
    setLocations(p=>[...p,newLoc.trim()]);
    setNewLoc("");
  }
  function removeLoc(l){setLocations(p=>p.filter(x=>x!==l));}

  function handleTemplateFile(file){
    setImportMsg("");
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        // Scan up to row 20 to skip instruction blocks at top of KellyOCG template
        let titleCol=-1,locCol=-1,levelCol=-1,headerRow=-1;
        for(let r=0;r<Math.min(20,rows.length);r++){
          const row=rows[r].map(v=>String(v).toLowerCase().trim());
          const ti=row.findIndex(c=>c.includes("job title")||c.includes("title")||c.includes("role"));
          if(ti>=0){
            titleCol=ti;
            locCol=row.findIndex(c=>c.includes("location")||c.includes("city")||c.includes("region"));
            levelCol=row.findIndex(c=>c.includes("level")&&!c.includes("title"));
            headerRow=r;
            break;
          }
        }
        if(titleCol<0){setImportMsg("⚠️ Couldn't find a Job Title column.");return;}
        const newJobs=[];
        const newLocs=new Set();
        const seenKeys=new Set();
        for(let r=headerRow+1;r<rows.length;r++){
          const row=rows[r];
          const rawTitle=String(row[titleCol]||"").trim();
          if(!rawTitle)continue;
          // Parse level from embedded title e.g. "Commercial Officer - Level 4"
          // fullTitle = original string as-is (used in Excel output)
          let title=rawTitle;
          let level="";
          const lvlMatch=rawTitle.match(/\s*[-–]\s*(level\s*\d+|[A-Za-z]+\s*\d*)$/i);
          if(lvlMatch){
            level=lvlMatch[1].trim();
            title=rawTitle.slice(0,rawTitle.length-lvlMatch[0].length).trim();
          } else if(levelCol>=0&&row[levelCol]){
            level=String(row[levelCol]).trim();
          }
          const key=rawTitle+"||"+level;
          if(!seenKeys.has(key)){
            seenKeys.add(key);
            newJobs.push({id:Date.now()+r+Math.random(),title,level,fullTitle:rawTitle});
          }
          if(locCol>=0&&row[locCol]) newLocs.add(String(row[locCol]).trim());
        }
        if(!newJobs.length){setImportMsg("⚠️ No role rows found.");return;}
        setJobs(newJobs);
        if(newLocs.size)setLocations([...newLocs]);
        setImportMsg("✅ Imported "+newJobs.length+" unique roles"+(newLocs.size?" and "+newLocs.size+" locations":""));
      }catch(err){setImportMsg("⚠️ Error reading file: "+err.message);}
    };
    reader.readAsBinaryString(file);
  }

  function resetTemplate(){
    setPreview(buildTemplate(clientName));
  }

  const [downloading,setDownloading]=useState(false);
  const [dlMsg,setDlMsg]=useState("");

  async function downloadExcelTemplate(){
    if(!jobs.length||!locations.length){
      setDlMsg("⚠️ Add at least one job title and one location first.");
      return;
    }
    setDownloading(true);
    setDlMsg("");
    try{
      const res=await fetch("/api/generate-template",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({jobs,locations,clientName:"[CLIENT NAME]",deadline:"[DEADLINE DATE]"}),
      });
      if(!res.ok){
        const err=await res.json().catch(()=>({error:"Unknown error"}));
        throw new Error(err.error||"Server error");
      }
      const blob=await res.blob();
      const url=URL.createObjectURL(blob);
      const a=document.createElement("a");
      a.href=url;
      a.download="KellyOCG_MRA_Rate_Survey_Template.xlsx";
      a.click();
      URL.revokeObjectURL(url);
      setDlMsg("✅ Formatted template downloaded successfully");
    }catch(err){
      setDlMsg("⚠️ Error: "+err.message+". Make sure the app is deployed on Vercel.");
    }
    setDownloading(false);
  }

  const grouped=jobs.reduce((acc,j)=>{if(!acc[j.title])acc[j.title]=[];acc[j.title].push(j);return acc;},{});

  return(
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <Card>
          <div style={{fontWeight:700,fontSize:15,color:C.navy,marginBottom:12}}>📥 Import Template File</div>
          <DropZone onFile={handleTemplateFile} label="Drop your template Excel/CSV here" sublabel="Columns: Job Title, Level (optional), Location (optional)"/>
          <Toast msg={importMsg}/>
        </Card>
        <Card>
          <div style={{fontWeight:700,fontSize:15,color:C.navy,marginBottom:14}}>📋 Job Titles & Levels</div>
          <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
            <Input value={newTitle} onChange={setNewTitle} placeholder="Job Title" style={{flex:2}}/>
            <Select value={newLevel} onChange={setNewLevel} options={levels.map(l=>({value:l,label:l}))} style={{flex:1}}/>
            <Btn onClick={addJob} variant="sky" size="sm" style={{alignSelf:"flex-end"}}>+ Add</Btn>
          </div>
          <div style={{maxHeight:220,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
            {Object.entries(grouped).map(([title,entries])=>(
              <div key={title}>
                <div style={{fontSize:11,fontWeight:700,color:C.textMuted,textTransform:"uppercase",letterSpacing:1,margin:"8px 0 4px"}}>{title}</div>
                {entries.map(j=>(
                  <div key={j.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:C.slateLight,borderRadius:7,padding:"6px 10px"}}>
                    <span style={{fontSize:13}}>{j.level}</span>
                    <button onClick={()=>removeJob(j.id)} style={{background:"none",border:"none",color:C.rose,cursor:"pointer",fontSize:16}}>×</button>
                  </div>
                ))}
              </div>
            ))}
          </div>
        </Card>
        <Card>
          <div style={{fontWeight:700,fontSize:15,color:C.navy,marginBottom:14}}>📍 Locations</div>
          <div style={{display:"flex",gap:8,marginBottom:12}}>
            <Input value={newLoc} onChange={setNewLoc} placeholder="e.g. Chicago, IL" style={{flex:1}}/>
            <Btn onClick={addLoc} variant="sky" size="sm" style={{alignSelf:"flex-end"}}>+ Add</Btn>
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
            {locations.map(loc=>(
              <span key={loc} style={{background:C.skyLight,color:C.sky,borderRadius:99,padding:"4px 10px",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:5}}>
                {loc}
                <button onClick={()=>removeLoc(loc)} style={{background:"none",border:"none",color:C.sky,cursor:"pointer",fontSize:14}}>×</button>
              </span>
            ))}
          </div>
        </Card>
      </div>

      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <Card style={{flex:1}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div>
              <div style={{fontWeight:700,fontSize:15,color:C.navy}}>✉️ Outreach Email Template</div>
              <div style={{fontSize:12,color:C.textMuted,marginTop:2}}>Edit directly — this is the body used when drafting supplier emails</div>
            </div>
            <Btn onClick={resetTemplate} variant="ghost" size="sm">Reset to Default</Btn>
          </div>
          <textarea value={preview} onChange={e=>setPreview(e.target.value)}
            style={{width:"100%",minHeight:320,border:`1px solid ${C.border}`,borderRadius:8,padding:12,fontSize:13,color:C.text,lineHeight:1.6,resize:"vertical",fontFamily:"inherit",boxSizing:"border-box"}}/>
          <div style={{display:"flex",gap:8,marginTop:10}}>
            <Btn size="sm" variant="mint" onClick={()=>navigator.clipboard?.writeText(preview)}>Copy to Clipboard</Btn>
          </div>
        </Card>
        <Card>
          <div style={{fontWeight:700,fontSize:15,color:C.navy,marginBottom:4}}>📊 Survey Grid Preview</div>
          <div style={{fontSize:12,color:C.textMuted,marginBottom:12}}>{jobs.length} roles · {locations.length} locations → {jobs.length*locations.length} total rows</div>
          <div style={{overflowX:"auto",maxHeight:220,overflowY:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead>
                <tr style={{background:C.navy}}>
                  <th style={{padding:"6px 10px",color:C.white,textAlign:"left",position:"sticky",top:0}}>Role</th>
                  <th style={{padding:"6px 10px",color:C.white,textAlign:"left",position:"sticky",top:0}}>Level</th>
                  <th style={{padding:"6px 10px",color:C.white,textAlign:"left",position:"sticky",top:0}}>Location</th>
                </tr>
              </thead>
              <tbody>
                {locations.slice(0,3).map(loc=>
                  jobs.slice(0,4).map((j,i)=>(
                    <tr key={loc+j.id} style={{background:i%2===0?C.white:C.slateLight,borderBottom:`1px solid ${C.border}`}}>
                      <td style={{padding:"5px 10px"}}>{j.title}</td>
                      <td style={{padding:"5px 10px",color:C.textMuted}}>{j.level}</td>
                      <td style={{padding:"5px 10px",color:C.sky}}>{loc}</td>
                    </tr>
                  ))
                )}
                {(jobs.length>4||locations.length>3)&&(
                  <tr><td colSpan={3} style={{padding:"6px 10px",color:C.textMuted,fontSize:11,fontStyle:"italic"}}>
                    +{Math.max(0,jobs.length*locations.length-12)} more rows in the downloaded file…
                  </td></tr>
                )}
              </tbody>
            </table>
          </div>
        </Card>
        <Card style={{border:`2px solid ${C.mint}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
            <div>
              <div style={{fontWeight:700,fontSize:15,color:C.navy}}>⬇️ Download Survey Template</div>
              <div style={{fontSize:12,color:C.textMuted,marginTop:2}}>Generates a formatted Excel file ready to send to suppliers</div>
            </div>
            <Btn onClick={downloadExcelTemplate} variant="mint" disabled={downloading||!jobs.length||!locations.length}>
              {downloading?"Generating…":"Download Excel"}
            </Btn>
          </div>
          {dlMsg&&<Toast msg={dlMsg}/>}
          <div style={{marginTop:10,display:"flex",flexWrap:"wrap",gap:6}}>
            <span style={{fontSize:11,color:C.textMuted}}>Will include:</span>
            {["KellyOCG header","Supplier info fields","Instructions","Green input cells for rates","All your roles × locations"].map(f=>(
              <span key={f} style={{background:C.mintLight,color:C.mint,borderRadius:99,padding:"2px 8px",fontSize:11,fontWeight:600}}>{f}</span>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

// ─── SUPPLIER TRACKER ─────────────────────────────────────────────────────────
function SupplierTracker({suppliers,setSuppliers,responses,jobs,locations}){
  const [newName,setNewName]=useState("");
  const [newEmail,setNewEmail]=useState("");
  const [filter,setFilter]=useState("all");
  const [importMsg,setImportMsg]=useState("");
  const [expandedCompletion,setExpandedCompletion]=useState(null);
  const [editingNote,setEditingNote]=useState(null);
  const [noteText,setNoteText]=useState("");

  function saveNote(id){
    setSuppliers(p=>p.map(s=>s.id===id?{...s,notes:noteText}:s));
    setEditingNote(null);setNoteText("");
  }

  function addSupplier(){
    if(!newName.trim())return;
    setSuppliers(p=>[...p,{id:Date.now(),name:newName.trim(),contact:newEmail.trim(),pocName:"",country:"",category:"",status:"not_sent",notes:"",sentAt:null,respondedAt:null}]);
    setNewName("");setNewEmail("");
  }
  function updateStatus(id,status){
    const now=new Date().toISOString().split("T")[0];
    setSuppliers(p=>p.map(s=>s.id===id?{
      ...s,status,
      sentAt:(status==="sent"&&!s.sentAt)?now:s.sentAt,
      respondedAt:status==="responded"?now:s.respondedAt,
    }:s));
  }

  function handleSupplierFile(file){
    setImportMsg("");
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"binary"});

        // Prefer "Consolidated Supplier List" sheet, fall back to first sheet
        const sheetName = wb.SheetNames.find(n=>
          n.toLowerCase().includes("consolidated")||n.toLowerCase().includes("supplier list")
        ) || wb.SheetNames[0];
        const ws=wb.Sheets[sheetName];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});

        // Scan up to row 10 to find header (skips KellyOCG banner rows)
        let cols={name:-1,email:-1,contact:-1,country:-1,category:-1};
        let hr=-1;
        for(let r=0;r<Math.min(10,rows.length);r++){
          const row=rows[r].map(v=>String(v).toLowerCase().trim());
          // Look for "supplier" column
          const ni=row.findIndex(c=>c==="supplier"||c.includes("supplier name")||c.includes("vendor")||c.includes("company"));
          if(ni>=0){
            cols.name=ni;
            cols.email=row.findIndex(c=>c.includes("e-mail")||c.includes("email")||c.includes("poc e")||c.includes("mail"));
            cols.contact=row.findIndex(c=>c.includes("point of contact")||c.includes("poc name")||c.includes("contact name")||c.includes("contact (poc)"));
            cols.country=row.findIndex(c=>c.includes("country"));
            cols.category=row.findIndex(c=>c.includes("category")||c.includes("collar")||c.includes("type"));
            hr=r; break;
          }
        }
        if(cols.name<0){setImportMsg("⚠️ Couldn't find a Supplier column. Make sure you're using the Consolidated Supplier List sheet.");return;}

        const imported=[];
        const seenNames=new Set();
        for(let r=hr+1;r<rows.length;r++){
          const row=rows[r];
          const name=String(row[cols.name]||"").trim();
          if(!name||seenNames.has(name.toLowerCase()))continue;
          seenNames.add(name.toLowerCase());
          const email=cols.email>=0?String(row[cols.email]||"").trim():"";
          const pocName=cols.contact>=0?String(row[cols.contact]||"").trim():"";
          const country=cols.country>=0?String(row[cols.country]||"").trim():"";
          const category=cols.category>=0?String(row[cols.category]||"").trim():"";
          imported.push({
            id:Date.now()+r+Math.random(),
            name, contact:email, pocName, country, category,
            status:"not_sent", notes:"", sentAt:null, respondedAt:null
          });
        }
        if(!imported.length){setImportMsg("⚠️ No supplier rows found.");return;}
        setSuppliers(prev=>{
          const existing=new Set(prev.map(s=>s.name.toLowerCase()));
          const fresh=imported.filter(s=>!existing.has(s.name.toLowerCase()));
          setImportMsg("✅ Added "+fresh.length+" suppliers from '"+sheetName+"' ("+(imported.length-fresh.length)+" duplicates skipped)");
          return [...prev,...fresh];
        });
      }catch(err){setImportMsg("⚠️ Error reading file: "+err.message);}
    };
    reader.readAsBinaryString(file);
  }

  const counts={
    all:suppliers.length,
    responded:suppliers.filter(s=>s.status==="responded").length,
    sent:suppliers.filter(s=>s.status==="sent").length,
    follow_up:suppliers.filter(s=>s.status==="follow_up").length,
    not_sent:suppliers.filter(s=>s.status==="not_sent").length,
  };
  const filtered=filter==="all"?suppliers:suppliers.filter(s=>s.status===filter);
  const pct=Math.round((counts.responded/Math.max(1,suppliers.length))*100);

  // Overall completeness across responded suppliers
  const respondedSuppliers=suppliers.filter(s=>s.status==="responded");
  const fullyComplete=respondedSuppliers.filter(s=>completeness(s.name,responses,jobs,locations).pct===100).length;
  const partialCount=respondedSuppliers.length-fullyComplete;

  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      {/* Stats */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12}}>
        {[
          ["Total",counts.all,C.navy],
          ["Responded",counts.responded,C.mint],
          ["Awaiting",counts.sent+counts.follow_up,C.amber],
          ["Not Sent",counts.not_sent,C.slate],
          ["Fully Complete",fullyComplete,C.purple],
        ].map(([l,v,col])=>(
          <Card key={l} style={{textAlign:"center",padding:"14px 10px"}}>
            <div style={{fontSize:26,fontWeight:800,color:col}}>{v}</div>
            <div style={{fontSize:11,color:C.textMuted,marginTop:2}}>{l}</div>
          </Card>
        ))}
      </div>

      {/* Progress */}
      <Card style={{padding:"14px 20px"}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
          <span style={{fontSize:13,fontWeight:600,color:C.navy}}>Response Progress</span>
          <div style={{display:"flex",gap:12,alignItems:"center"}}>
            {partialCount>0&&<span style={{fontSize:12,color:C.amber,fontWeight:600}}>⚠️ {partialCount} partial response{partialCount!==1?"s":""}</span>}
            <span style={{fontSize:13,fontWeight:700,color:C.mint}}>{pct}%</span>
          </div>
        </div>
        <div style={{height:10,background:C.slateLight,borderRadius:99,overflow:"hidden"}}>
          <div style={{height:"100%",width:`${pct}%`,background:C.mint,borderRadius:99,transition:"width .4s"}}/>
        </div>
      </Card>

      {/* Import */}
      <Card>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:10}}>📥 Import Supplier List</div>
        <DropZone onFile={handleSupplierFile} label="Drop your supplier list Excel/CSV here" sublabel="Columns: Supplier Name, Email/Contact (optional)"/>
        <Toast msg={importMsg}/>
      </Card>

      {/* Add single */}
      <Card>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:10}}>+ Add Single Supplier</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <Input value={newName} onChange={setNewName} placeholder="Supplier Name" style={{flex:2}}/>
          <Input value={newEmail} onChange={setNewEmail} placeholder="Contact Email" style={{flex:3}}/>
          <Btn onClick={addSupplier} variant="navy" size="sm" style={{alignSelf:"flex-end"}}>Add</Btn>
        </div>
      </Card>

      {/* Supplier list */}
      <Card>
        <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>
          {[["all","All"],["responded","Responded"],["sent","Sent"],["follow_up","Follow-up"],["not_sent","Not Sent"]].map(([val,label])=>(
            <button key={val} onClick={()=>setFilter(val)} style={{
              padding:"5px 14px",borderRadius:99,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,
              background:filter===val?C.navy:C.slateLight,color:filter===val?C.white:C.textMuted,
            }}>{label} ({val==="all"?counts.all:counts[val]??0})</button>
          ))}
        </div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead>
            <tr style={{borderBottom:`2px solid ${C.border}`}}>
              {["Supplier","Country","Category","Contact","Status","Completeness","Notes"].map(h=>(
                <th key={h} style={{padding:"6px 10px",textAlign:"left",color:C.textMuted,fontWeight:600,fontSize:12}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map(s=>{
              const comp=completeness(s.name,responses,jobs,locations);
              const isExpanded=expandedCompletion===s.id;
              return(
                <>
                  <tr key={s.id} style={{borderBottom:isExpanded?"none":`1px solid ${C.border}`,verticalAlign:"top"}}>
                    <td style={{padding:"9px 10px",fontWeight:600,minWidth:160}}>
                      {s.name}
                      {s.pocName&&<div style={{fontSize:10,color:C.textMuted,marginTop:1}}>👤 {s.pocName}</div>}
                      {s.sentAt&&<div style={{fontSize:10,color:C.textMuted,marginTop:1}}>Sent {s.sentAt}</div>}
                      {s.respondedAt&&<div style={{fontSize:10,color:C.mint,marginTop:1}}>Responded {s.respondedAt}</div>}
                    </td>
                    <td style={{padding:"9px 10px",color:C.textMuted,fontSize:12}}>{s.country||"—"}</td>
                    <td style={{padding:"9px 10px",color:C.textMuted,fontSize:12}}>{s.category||"—"}</td>
                    <td style={{padding:"9px 10px",color:C.textMuted,fontSize:12}}>{s.contact||<span style={{color:C.rose,fontSize:11}}>No email</span>}</td>
                    <td style={{padding:"9px 10px"}}>
                      <select
                        value={s.status}
                        onChange={e=>updateStatus(s.id,e.target.value)}
                        style={{
                          border:`1px solid ${C.border}`,borderRadius:7,padding:"5px 8px",fontSize:12,
                          fontWeight:600,cursor:"pointer",background:C.white,color:C.text,outline:"none",
                        }}
                      >
                        <option value="not_sent">Not Sent</option>
                        <option value="sent">Sent</option>
                        <option value="follow_up">Follow-up</option>
                        <option value="responded">Responded</option>
                      </select>
                      {s.sentAt&&<div style={{fontSize:10,color:C.textMuted,marginTop:3}}>Sent {s.sentAt}</div>}
                      {s.respondedAt&&<div style={{fontSize:10,color:C.mint,marginTop:2}}>Responded {s.respondedAt}</div>}
                    </td>
                    <td style={{padding:"9px 10px"}}>
                      {s.status==="responded"?(
                        <div>
                          <CompletenessBar pct={comp.pct} filled={comp.filled} expected={comp.expected}/>
                          {comp.missing.length>0&&(
                            <button onClick={()=>setExpandedCompletion(isExpanded?null:s.id)}
                              style={{background:"none",border:"none",color:C.amber,cursor:"pointer",fontSize:11,padding:"2px 0",fontWeight:600}}>
                              {isExpanded?"▲ Hide":"▼ Show"} {comp.missing.length} missing
                            </button>
                          )}
                        </div>
                      ):<span style={{color:C.textMuted,fontSize:12}}>—</span>}
                    </td>
                    <td style={{padding:"9px 10px",minWidth:180}}>
                      {editingNote===s.id?(
                        <div style={{display:"flex",flexDirection:"column",gap:4}}>
                          <textarea value={noteText} onChange={e=>setNoteText(e.target.value)}
                            placeholder="e.g. Declined to participate — rate confidentiality policy"
                            style={{width:"100%",fontSize:11,border:`1px solid ${C.border}`,borderRadius:6,padding:"5px 7px",resize:"vertical",minHeight:60,fontFamily:"inherit",boxSizing:"border-box"}}/>
                          <div style={{display:"flex",gap:4}}>
                            <Btn size="sm" variant="mint" onClick={()=>saveNote(s.id)}>Save</Btn>
                            <Btn size="sm" variant="ghost" onClick={()=>setEditingNote(null)}>Cancel</Btn>
                          </div>
                        </div>
                      ):(
                        <div onClick={()=>{setEditingNote(s.id);setNoteText(s.notes||"");}}
                          style={{cursor:"pointer",fontSize:11,color:s.notes?C.text:C.textMuted,lineHeight:1.5,
                            padding:"4px 7px",borderRadius:6,border:`1px dashed ${s.notes?C.border:"#E2E8F0"}`,
                            background:s.notes?C.amberLight:"transparent",minHeight:28,
                          }}
                          title="Click to add/edit note">
                          {s.notes||<span style={{fontStyle:"italic"}}>Add note…</span>}
                        </div>
                      )}
                    </td>
                  </tr>
                  {isExpanded&&comp.missing.length>0&&(
                    <tr key={s.id+"_missing"} style={{borderBottom:`1px solid ${C.border}`}}>
                      <td colSpan={7} style={{padding:"0 10px 10px 10px"}}>
                        <div style={{background:C.amberLight,borderRadius:8,padding:"10px 14px"}}>
                          <div style={{fontWeight:600,fontSize:12,color:C.amber,marginBottom:6}}>Missing from {s.name}'s response:</div>
                          <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                            {comp.missing.map(m=>(
                              <span key={m} style={{background:C.white,color:C.amber,border:`1px solid ${C.amber}44`,
                                borderRadius:99,padding:"2px 8px",fontSize:11}}>{m}</span>
                            ))}
                          </div>
                        </div>
                      </td>
                    </tr>
                  )}
                </>
              );
            })}
          </tbody>
        </table>
      </Card>

    </div>
  );
}

// ─── DATA ENTRY ───────────────────────────────────────────────────────────────
function DataEntry({responses,setResponses,suppliers}){
  const [selectedSupplier,setSelectedSupplier]=useState(suppliers[0]?.name||"");
  const [addingNew,setAddingNew]=useState(false);
  const [newSupName,setNewSupName]=useState("");
  const [importMsg,setImportMsg]=useState("");
  const [preview,setPreview]=useState([]);
  const [manual,setManual]=useState({title:"",level:"Mid",location:"",billRate:"",payRate:""});

  const supplierNames=[...new Set([...suppliers.map(s=>s.name),...responses.map(r=>r.supplier)])];

  function handleExcelFile(file){
    setImportMsg("");setPreview([]);
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        let cols={title:-1,level:-1,location:-1,country:-1,bill:-1,pay:-1};let hr=-1;
        // Scan up to row 20 to handle KellyOCG format with instruction rows at top
        // Normalize cell values: collapse newlines and extra spaces for matching
        for(let r=0;r<Math.min(20,rows.length);r++){
          const row=rows[r].map(v=>String(v).toLowerCase().replace(/\n/g," ").replace(/\s+/g," ").trim());
          // Must be a SHORT cell containing "job title" — not an instruction sentence
          // A header cell will be short (< 60 chars) and contain "job title" or similar
          const ti=row.findIndex(c=>{
            if(c.length>80) return false; // skip long instruction sentences
            return c.includes("job title")||(c.includes("title")&&!c.includes("leave blank")&&!c.includes("please")&&!c.includes("does not"))||c==="role"||c==="position";
          });
          if(ti>=0){
            cols.title=ti;
            // Level: short cell containing "level" but not the title col
            cols.level=row.findIndex((c,i)=>i!==ti&&c.length<40&&(c==="level"||c.includes("level"))&&!c.includes("title"));
            // Location: prefer "location" column over "country/region"
            const strictLoc=row.findIndex(c=>c.length<40&&(c==="location"||c.includes("location")||c.includes("city")));
            const looseLoc=row.findIndex(c=>c.length<40&&(c.includes("region")||c.includes("country")));
            cols.location=strictLoc>=0?strictLoc:looseLoc;
            // Also track country column separately for context
            cols.country=row.findIndex(c=>c.length<40&&(c.includes("country")||c.includes("region")));
            // Bill rate: must contain "bill"
            cols.bill=row.findIndex(c=>c.includes("bill"));
            // Pay rate: contains "pay" but not "bill"
            cols.pay=row.findIndex(c=>c.includes("pay")&&!c.includes("bill"));
            if(cols.pay<0) cols.pay=row.findIndex(c=>c.includes("wage")||c.includes("salary")||c.includes("cost rate"));
            hr=r;break;
          }
        }
        if(cols.title<0){setImportMsg("⚠️ Couldn't find a Title/Role column.");return;}
        const parsed=[];
        for(let r=hr+1;r<rows.length;r++){
          const row=rows[r];
          const rawTitle=String(row[cols.title]||"").trim();
          if(!rawTitle)continue;
          // Parse embedded level from title e.g. "Commercial Officer - Level 4"
          let title=rawTitle;
          let level=cols.level>=0?String(row[cols.level]||"").trim():"";
          if(!level){
            const lvlMatch=rawTitle.match(/\s*[-–]\s*(level\s*\d+|[A-Za-z]+\s*\d*)$/i);
            if(lvlMatch){level=lvlMatch[1].trim();title=rawTitle.slice(0,rawTitle.length-lvlMatch[0].length).trim();}
          }
          const location=cols.location>=0?String(row[cols.location]||"").trim():"";
          const payRate=cols.pay>=0?parseFloat(String(row[cols.pay]).replace(/[^0-9.]/g,""))||0:0;
          const billRate=cols.bill>=0?parseFloat(String(row[cols.bill]).replace(/[^0-9.]/g,""))||0:0;
          // Include rows that have a title — rate can be 0 (supplier may fill in later)
          // Use location col if found, otherwise fall back to country col
          const locVal=cols.location>=0?String(row[cols.location]||"").trim():"";
          const countryVal=cols.country>=0&&cols.country!==cols.location?String(row[cols.country]||"").trim():"";
          const finalLocation=locVal||(countryVal?"":location);
          parsed.push({supplier:selectedSupplier,title,level,location:finalLocation||location,billRate,payRate});
        }
        if(!parsed.length){setImportMsg("⚠️ No data rows found.");return;}
        setPreview(parsed);
        setImportMsg(`✅ Found ${parsed.length} records — review and confirm`);
      }catch{setImportMsg("⚠️ Error reading file.");}
    };
    reader.readAsBinaryString(file);
  }

  function confirmImport(){
    setResponses(p=>[...p,...preview]);
    setImportMsg(`✅ Imported ${preview.length} records from ${selectedSupplier}`);
    setPreview([]);
  }
  function addManual(){
    if(!manual.title||(!manual.billRate&&!manual.payRate))return;
    setResponses(p=>[...p,{...manual,supplier:selectedSupplier,billRate:Number(manual.billRate)||0,payRate:Number(manual.payRate)||0}]);
    setManual(p=>({...p,billRate:"",payRate:""}));
  }
  function exportCSV(){
    const headers=["Supplier","Job Title","Level","Location","Bill Rate","Pay Rate"];
    const rows=responses.map(r=>[r.supplier,r.title,r.level,r.location,r.billRate,r.payRate]);
    const csv=[headers,...rows].map(r=>r.map(v=>`"${v}"`).join(",")).join("\n");
    const blob=new Blob([csv],{type:"text/csv"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");a.href=url;a.download="supplier_rates_consolidated.csv";a.click();
    URL.revokeObjectURL(url);
  }

  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Card>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:10}}>Entering data for supplier:</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          {supplierNames.map(s=>(
            <button key={s} onClick={()=>{setSelectedSupplier(s);setAddingNew(false);}} style={{
              padding:"6px 14px",borderRadius:99,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,
              background:selectedSupplier===s&&!addingNew?C.navy:C.slateLight,
              color:selectedSupplier===s&&!addingNew?C.white:C.textMuted,
            }}>{s}</button>
          ))}
          {!addingNew
            ?<Btn size="sm" variant="ghost" onClick={()=>setAddingNew(true)}>+ New</Btn>
            :(
              <div style={{display:"flex",gap:6,alignItems:"center"}}>
                <Input value={newSupName} onChange={setNewSupName} placeholder="Supplier name"/>
                <Btn size="sm" variant="sky" onClick={()=>{if(newSupName.trim()){setSelectedSupplier(newSupName.trim());setAddingNew(false);setNewSupName("");}}}>Set</Btn>
                <Btn size="sm" variant="ghost" onClick={()=>setAddingNew(false)}>Cancel</Btn>
              </div>
            )
          }
        </div>
      </Card>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <Card>
          <div style={{fontWeight:700,fontSize:15,color:C.navy,marginBottom:4}}>📂 Upload Supplier Excel Response</div>
          <div style={{fontSize:12,color:C.textMuted,marginBottom:12}}>Drop the Excel file sent back by the supplier.</div>
          <DropZone onFile={handleExcelFile} label={`Drop ${selectedSupplier}'s response file here`} sublabel="Expected columns: Job Title, Level, Location, Bill Rate, Pay Rate"/>
          <Toast msg={importMsg}/>
          {preview.length>0&&(
            <div style={{marginTop:12}}>
              <div style={{fontWeight:600,fontSize:13,color:C.navy,marginBottom:6}}>Preview ({preview.length} records):</div>
              <div style={{overflowX:"auto",maxHeight:160,overflowY:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                  <thead><tr style={{background:C.slateLight}}>
                    {["Title","Level","Location","Bill $","Pay $"].map(h=><th key={h} style={{padding:"4px 8px",textAlign:"left",color:C.textMuted}}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {preview.slice(0,8).map((r,i)=>(
                      <tr key={i} style={{borderBottom:`1px solid ${C.border}`}}>
                        <td style={{padding:"4px 8px"}}>{r.title}</td>
                        <td style={{padding:"4px 8px",color:C.textMuted}}>{r.level}</td>
                        <td style={{padding:"4px 8px",color:C.textMuted}}>{r.location}</td>
                        <td style={{padding:"4px 8px",color:C.sky,fontWeight:600}}>{r.billRate||"—"}</td>
                        <td style={{padding:"4px 8px",color:C.mint,fontWeight:600}}>{r.payRate||"—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {preview.length>8&&<div style={{fontSize:11,color:C.textMuted,marginTop:4}}>+{preview.length-8} more rows…</div>}
              <Btn onClick={confirmImport} variant="mint" style={{marginTop:10}}>✓ Confirm Import ({preview.length} records)</Btn>
            </div>
          )}
        </Card>

        <Card>
          <div style={{fontWeight:700,fontSize:15,color:C.navy,marginBottom:12}}>✏️ Manual Entry</div>
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <Input label="Job Title" value={manual.title} onChange={v=>setManual(p=>({...p,title:v}))} placeholder="e.g. Software Engineer"/>
              <Input label="Level" value={manual.level} onChange={v=>setManual(p=>({...p,level:v}))} placeholder="e.g. Senior"/>
            </div>
            <Input label="Location" value={manual.location} onChange={v=>setManual(p=>({...p,location:v}))} placeholder="e.g. Chicago, IL"/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <Input label="Bill Rate ($/hr)" type="number" value={manual.billRate} onChange={v=>setManual(p=>({...p,billRate:v}))} placeholder="95"/>
              <Input label="Pay Rate ($/hr)" type="number" value={manual.payRate} onChange={v=>setManual(p=>({...p,payRate:v}))} placeholder="72"/>
            </div>
            <Btn onClick={addManual} variant="sky">Add Record</Btn>
          </div>
        </Card>
      </div>

      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <div style={{fontWeight:700,fontSize:14,color:C.navy}}>All Responses ({responses.length} records)</div>
          <Btn onClick={exportCSV} variant="mint" size="sm">⬇️ Export to CSV</Btn>
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:C.slateLight}}>
              {["Supplier","Role","Level","Location","Bill $/hr","Pay $/hr",""].map(h=>(
                <th key={h} style={{padding:"7px 10px",textAlign:"left",color:C.textMuted,fontWeight:600}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {[...responses].reverse().slice(0,12).map((r,i)=>(
                <tr key={i} style={{borderBottom:`1px solid ${C.border}`}}>
                  <td style={{padding:"7px 10px",fontWeight:600,color:C.navy}}>{r.supplier}</td>
                  <td style={{padding:"7px 10px"}}>{r.title}</td>
                  <td style={{padding:"7px 10px",color:C.textMuted}}>{r.level}</td>
                  <td style={{padding:"7px 10px",color:C.textMuted}}>{r.location}</td>
                  <td style={{padding:"7px 10px",color:C.sky,fontWeight:700}}>${r.billRate}</td>
                  <td style={{padding:"7px 10px",color:C.mint,fontWeight:700}}>${r.payRate}</td>
                  <td style={{padding:"7px 10px"}}>
                    <button onClick={()=>setResponses(p=>{const rev=[...p].reverse();rev.splice(i,1);return rev.reverse();})}
                      style={{background:"none",border:"none",color:C.rose,cursor:"pointer"}}>×</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ─── ANALYTICS ────────────────────────────────────────────────────────────────
function Analytics({responses,suppliers,jobs,locations}){
  const [filterLoc,setFilterLoc]=useState("All");
  const [excludeOutliers,setExcludeOutliers]=useState(false);
  const [showOutliers,setShowOutliers]=useState(false);

  const locs=["All",...new Set(responses.map(r=>r.location))];
  const locFiltered=filterLoc==="All"?responses:responses.filter(r=>r.location===filterLoc);

  const roleKeys=[...new Set(locFiltered.map(r=>`${r.title}|||${r.level}`))];
  const outlierDetails=[];
  const outlierKeySet=new Set();
  roleKeys.forEach(key=>{
    const [title,level]=key.split("|||");
    const recs=locFiltered.filter(r=>r.title===title&&r.level===level);
    const bills=recs.map(r=>r.billRate).filter(Boolean);
    const pays=recs.map(r=>r.payRate).filter(Boolean);
    recs.forEach(r=>{
      const bo=isOutlier(r.billRate,bills);
      const po=isOutlier(r.payRate,pays);
      if(bo||po){
        outlierKeySet.add(`${r.supplier}|${r.title}|${r.level}|${r.location}`);
        outlierDetails.push({...r,billOutlier:bo,payOutlier:po,avgBill:avg(bills),avgPay:avg(pays)});
      }
    });
  });

  const workingData=excludeOutliers
    ?locFiltered.filter(r=>!outlierKeySet.has(`${r.supplier}|${r.title}|${r.level}|${r.location}`))
    :locFiltered;

  // Use title+level as unique key but display title only (level shown separately)
  const roleKeys2=[...new Set(workingData.map(r=>r.title+"|||"+r.level))];
  const summaries=roleKeys2.map(key=>{
    const [title,level]=key.split("|||");
    const role=title+(level?" – "+level:"");
    const recs=workingData.filter(r=>r.title===title&&r.level===level);
    const bills=recs.map(r=>r.billRate).filter(Boolean);
    const pays=recs.map(r=>r.payRate).filter(Boolean);
    return{role,title,level,responses:recs.length,
      billAvg:avg(bills),billMin:bills.length?Math.min(...bills):0,billMax:bills.length?Math.max(...bills):0,
      payAvg:avg(pays),payMin:pays.length?Math.min(...pays):0,payMax:pays.length?Math.max(...pays):0,
      spread:bills.length?Math.max(...bills)-Math.min(...bills):0,
    };
  }).sort((a,b)=>b.billAvg-a.billAvg);

  function exportCSV(){
    const esc=v=>`"${String(v??'').replace(/"/g,'""')}"`;
    const row=(...vals)=>vals.map(esc).join(",")+",,,";
    const blank=()=>"";
    const lines=[];
    const now=new Date().toLocaleDateString("en-US",{year:"numeric",month:"long",day:"numeric"});

    // ── SECTION 1: Summary Header ──
    lines.push(esc("KellyOCG MRA Rate Survey — Analytics Export")+",,,,,,,");
    lines.push(esc("Exported: "+now)+",,,,,,,");
    lines.push(esc("Location Filter: "+filterLoc)+",,,,,,,");
    lines.push(esc("Outliers: "+(excludeOutliers?"Excluded":"Included"))+",,,,,,,");
    lines.push(blank());

    // ── SECTION 2: Rate Summary by Role ──
    lines.push(esc("RATE SUMMARY BY ROLE")+",,,,,,,");
    lines.push([
      esc("Job Title"),esc("Level"),esc("# Responses"),
      esc("Bill Rate Avg"),esc("Bill Rate Min"),esc("Bill Rate Max"),
      esc("Pay Rate Avg"),esc("Pay Rate Min"),esc("Pay Rate Max"),
      esc("Spread"),
    ].join(","));
    summaries.forEach(s=>{
      lines.push([
        esc(s.title||s.role),esc(s.level||""),esc(s.responses),
        esc("$"+s.billAvg),esc("$"+s.billMin),esc("$"+s.billMax),
        esc("$"+s.payAvg),esc("$"+s.payMin),esc("$"+s.payMax),
        esc("$"+s.spread),
      ].join(","));
    });
    lines.push(blank());

    // ── SECTION 3: Supplier Comparison ──
    lines.push(esc("SUPPLIER COMPARISON")+",,,,,,,");
    lines.push([
      esc("Supplier"),esc("Avg Bill Rate"),esc("Avg Pay Rate"),
      esc("Avg Markup %"),esc("# Records"),esc("Outlier Count"),
    ].join(","));
    const supplierNames=[...new Set(responses.map(r=>r.supplier))];
    supplierNames.forEach(sup=>{
      const recs=locFiltered.filter(r=>r.supplier===sup);
      const ab=avg(recs.map(r=>r.billRate));
      const ap=avg(recs.map(r=>r.payRate));
      const markup=ap>0?Math.round(((ab-ap)/ap)*100):0;
      const oc=outlierDetails.filter(o=>o.supplier===sup).length;
      lines.push([
        esc(sup),esc("$"+ab+"/hr"),esc("$"+ap+"/hr"),
        esc(markup+"%"),esc(recs.length),esc(oc),
      ].join(","));
    });
    lines.push(blank());

    // ── SECTION 4: Outlier Detail ──
    if(outlierDetails.length>0){
      lines.push(esc("OUTLIER DETAIL (±1.5 SD from group mean)")+",,,,,,,");
      lines.push([
        esc("Supplier"),esc("Job Title"),esc("Level"),esc("Location"),
        esc("Bill Rate"),esc("Bill Rate Group Avg"),esc("Bill Outlier"),
        esc("Pay Rate"),esc("Pay Rate Group Avg"),esc("Pay Outlier"),
      ].join(","));
      outlierDetails.forEach(o=>{
        lines.push([
          esc(o.supplier),esc(o.title),esc(o.level),esc(o.location),
          esc("$"+o.billRate),esc("$"+o.avgBill),esc(o.billOutlier?"YES":""),
          esc("$"+o.payRate),esc("$"+o.avgPay),esc(o.payOutlier?"YES":""),
        ].join(","));
      });
      lines.push(blank());
    }

    // ── SECTION 5: Raw Response Data ──
    lines.push(esc("RAW RESPONSE DATA")+",,,,,,,");
    lines.push([
      esc("Supplier"),esc("Job Title"),esc("Level"),esc("Location"),
      esc("Bill Rate ($/hr)"),esc("Pay Rate ($/hr)"),
      esc("Bill Outlier"),esc("Pay Outlier"),
    ].join(","));
    locFiltered.forEach(r=>{
      const k=r.supplier+"|"+r.title+"|"+r.level+"|"+r.location;
      const od=outlierDetails.find(o=>o.supplier+"|"+o.title+"|"+o.level+"|"+o.location===k);
      lines.push([
        esc(r.supplier),esc(r.title),esc(r.level),esc(r.location),
        esc(r.billRate),esc(r.payRate),
        esc(od?.billOutlier?"Yes":"No"),esc(od?.payOutlier?"Yes":"No"),
      ].join(","));
    });

    const csv=lines.join("\n");
    const blob=new Blob([csv],{type:"text/csv"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;
    a.download="KellyOCG_MRA_Rate_Survey_Analytics_"+now.replace(/[, ]+/g,"_")+".csv";
    a.click();
    URL.revokeObjectURL(url);
  }



  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Card style={{padding:"12px 20px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
          <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
            <span style={{fontSize:12,fontWeight:600,color:C.textMuted,marginRight:4}}>Location:</span>
            {locs.map(loc=>(
              <button key={loc} onClick={()=>setFilterLoc(loc)} style={{
                padding:"4px 12px",borderRadius:99,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,
                background:filterLoc===loc?C.navy:C.slateLight,color:filterLoc===loc?C.white:C.textMuted,
              }}>{loc}</button>
            ))}
          </div>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            {outlierDetails.length>0&&(
              <button onClick={()=>setShowOutliers(!showOutliers)} style={{
                fontSize:12,fontWeight:600,color:C.amber,background:C.amberLight,
                border:"none",borderRadius:99,padding:"5px 12px",cursor:"pointer",
              }}>⚠️ {outlierDetails.length} outlier{outlierDetails.length!==1?"s":""} {showOutliers?"▲":"▼"}</button>
            )}
            <label style={{display:"flex",alignItems:"center",gap:7,cursor:"pointer",fontSize:13,fontWeight:600,color:C.text,userSelect:"none"}}>
              <div onClick={()=>setExcludeOutliers(!excludeOutliers)} style={{
                width:36,height:20,borderRadius:99,background:excludeOutliers?C.mint:C.border,
                position:"relative",transition:"background .2s",cursor:"pointer",flexShrink:0,
              }}>
                <div style={{width:16,height:16,borderRadius:"50%",background:C.white,position:"absolute",top:2,left:excludeOutliers?18:2,transition:"left .2s",boxShadow:"0 1px 3px rgba(0,0,0,.2)"}}/>
              </div>
              Exclude outliers
            </label>
            <Btn onClick={exportCSV} variant="ghost" size="sm">⬇️ Export CSV</Btn>
          </div>
        </div>
      </Card>

      {showOutliers&&outlierDetails.length>0&&(
        <Card style={{border:`2px solid ${C.amber}`}}>
          <div style={{fontWeight:700,fontSize:15,color:C.navy,marginBottom:12}}>⚠️ Outlier Detail — Values ±1.5 SD from group mean</div>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr style={{background:C.amberLight}}>
                {["Supplier","Role","Level","Location","Bill Rate","Group Avg","Pay Rate","Group Avg","Flags"].map(h=>(
                  <th key={h} style={{padding:"6px 10px",textAlign:"left",color:"#92400E",fontWeight:700}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {outlierDetails.map((o,i)=>(
                  <tr key={i} style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?C.white:"#FFFBEB"}}>
                    <td style={{padding:"7px 10px",fontWeight:700,color:C.navy}}>{o.supplier}</td>
                    <td style={{padding:"7px 10px"}}>{o.title}</td>
                    <td style={{padding:"7px 10px",color:C.textMuted}}>{o.level}</td>
                    <td style={{padding:"7px 10px",color:C.textMuted}}>{o.location}</td>
                    <td style={{padding:"7px 10px",fontWeight:700,color:o.billOutlier?C.rose:C.text}}>${o.billRate}</td>
                    <td style={{padding:"7px 10px",color:C.textMuted,fontSize:11}}>${o.avgBill} avg</td>
                    <td style={{padding:"7px 10px",fontWeight:700,color:o.payOutlier?C.rose:C.text}}>${o.payRate}</td>
                    <td style={{padding:"7px 10px",color:C.textMuted,fontSize:11}}>${o.avgPay} avg</td>
                    <td style={{padding:"7px 10px"}}>
                      <div style={{display:"flex",gap:4}}>
                        {o.billOutlier&&<span style={{background:C.roseLight,color:C.rose,padding:"1px 7px",borderRadius:99,fontSize:11,fontWeight:700}}>Bill</span>}
                        {o.payOutlier&&<span style={{background:C.amberLight,color:C.amber,padding:"1px 7px",borderRadius:99,fontSize:11,fontWeight:700}}>Pay</span>}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      <Card>
        <div style={{fontWeight:700,fontSize:15,color:C.navy,marginBottom:14}}>
          📊 Rate Summary by Role {excludeOutliers&&<span style={{fontSize:12,color:C.mint,fontWeight:600,marginLeft:6}}>(outliers excluded)</span>}
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:C.navy}}>
                <th style={{padding:"8px 12px",color:C.white,textAlign:"left"}}>Role</th>
                <th style={{padding:"8px 8px",color:C.white,textAlign:"center"}}>#</th>
                <th style={{padding:"8px 8px",color:"#93C5FD",textAlign:"center"}} colSpan={3}>Bill Rate ($/hr)</th>
                <th style={{padding:"8px 8px",color:"#6EE7B7",textAlign:"center"}} colSpan={3}>Pay Rate ($/hr)</th>
                <th style={{padding:"8px 8px",color:C.white,textAlign:"center"}}>Spread</th>
              </tr>
              <tr style={{background:C.navyLight}}>
                <th style={{padding:"5px 12px"}}></th><th></th>
                {["Avg","Min","Max","Avg","Min","Max"].map((h,i)=>(
                  <th key={i} style={{padding:"5px 8px",color:"#93C5FD",textAlign:"center",fontSize:11,fontWeight:600}}>{h}</th>
                ))}
                <th></th>
              </tr>
            </thead>
            <tbody>
              {summaries.map((s,i)=>(
                <tr key={s.role} style={{background:i%2===0?C.white:C.slateLight,borderBottom:`1px solid ${C.border}`}}>
                  <td style={{padding:"8px 12px",fontWeight:600}}>{s.title||s.role}</td>
                  <td style={{padding:"8px 8px",textAlign:"center",color:C.textMuted}}>{s.responses}</td>
                  <td style={{padding:"8px 8px",textAlign:"center",fontWeight:700,color:C.sky}}>${s.billAvg}</td>
                  <td style={{padding:"8px 8px",textAlign:"center",color:C.textMuted}}>${s.billMin}</td>
                  <td style={{padding:"8px 8px",textAlign:"center",color:C.textMuted}}>${s.billMax}</td>
                  <td style={{padding:"8px 8px",textAlign:"center",fontWeight:700,color:C.mint}}>${s.payAvg}</td>
                  <td style={{padding:"8px 8px",textAlign:"center",color:C.textMuted}}>${s.payMin}</td>
                  <td style={{padding:"8px 8px",textAlign:"center",color:C.textMuted}}>${s.payMax}</td>
                  <td style={{padding:"8px 8px",textAlign:"center"}}>
                    <span style={{
                      background:s.spread>30?C.roseLight:s.spread>15?C.amberLight:C.mintLight,
                      color:s.spread>30?C.rose:s.spread>15?C.amber:C.mint,
                      padding:"2px 8px",borderRadius:99,fontSize:12,fontWeight:700,
                    }}>${s.spread}</span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <Card>
        <div style={{fontWeight:700,fontSize:15,color:C.navy,marginBottom:14}}>🏢 Supplier Comparison</div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:C.slateLight}}>
              {["Supplier","Avg Bill","Avg Pay","Markup","Records","Completeness","Outliers"].map(h=>(
                <th key={h} style={{padding:"7px 12px",textAlign:"left",color:C.textMuted,fontWeight:600}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {[...new Set(responses.map(r=>r.supplier))].map((sup,i)=>{
                const recs=locFiltered.filter(r=>r.supplier===sup);
                const ab=avg(recs.map(r=>r.billRate));
                const ap=avg(recs.map(r=>r.payRate));
                const markup=ap>0?Math.round(((ab-ap)/ap)*100):0;
                const oc=outlierDetails.filter(o=>o.supplier===sup).length;
                const comp=completeness(sup,responses,jobs,locations);
                return(
                  <tr key={sup} style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?C.white:C.slateLight}}>
                    <td style={{padding:"8px 12px",fontWeight:600,color:C.navy}}>{sup}</td>
                    <td style={{padding:"8px 12px",color:C.sky,fontWeight:700}}>${ab}/hr</td>
                    <td style={{padding:"8px 12px",color:C.mint,fontWeight:700}}>${ap}/hr</td>
                    <td style={{padding:"8px 12px"}}>
                      <span style={{background:markup>35?C.roseLight:markup>25?C.amberLight:C.mintLight,color:markup>35?C.rose:markup>25?C.amber:C.mint,padding:"2px 8px",borderRadius:99,fontWeight:700}}>{markup}%</span>
                    </td>
                    <td style={{padding:"8px 12px",color:C.textMuted}}>{recs.length}</td>
                    <td style={{padding:"8px 12px"}}>
                      <CompletenessBar pct={comp.pct} filled={comp.filled} expected={comp.expected}/>
                    </td>
                    <td style={{padding:"8px 12px"}}>
                      {oc>0
                        ?<span style={{background:C.roseLight,color:C.rose,padding:"2px 8px",borderRadius:99,fontSize:11,fontWeight:700}}>⚠️ {oc}</span>
                        :<span style={{color:C.mint,fontSize:11}}>✓ None</span>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ─── PROJECT SELECTOR ────────────────────────────────────────────────────────
function ProjectSelector({projects,activeId,onSelect,onCreate,onDelete,loading}){
  const [newName,setNewName]=useState("");
  const [newClient,setNewClient]=useState("");
  const [creating,setCreating]=useState(false);

  function handleCreate(){
    if(!newName.trim())return;
    onCreate(newName.trim(),newClient.trim());
    setNewName("");setNewClient("");setCreating(false);
  }

  return(
    <div style={{minHeight:"100vh",background:"#F0F4F9",fontFamily:"'Inter',system-ui,sans-serif",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24}}>
      <div style={{width:"100%",maxWidth:600}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{fontSize:32,marginBottom:8}}>⚡</div>
          <div style={{fontSize:24,fontWeight:800,color:C.navy,letterSpacing:-0.5}}>SupplierRate</div>
          <div style={{fontSize:14,color:C.textMuted,marginTop:4}}>Rate Survey Management · Kelly Services MRA</div>
        </div>

        <Card>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
            <div style={{fontWeight:700,fontSize:16,color:C.navy}}>Your Survey Projects</div>
            <Btn variant="sky" size="sm" onClick={()=>setCreating(!creating)}>+ New Project</Btn>
          </div>

          {creating&&(
            <div style={{background:C.skyLight,borderRadius:10,padding:16,marginBottom:16,display:"flex",flexDirection:"column",gap:10}}>
              <Input label="Project Name" value={newName} onChange={setNewName} placeholder="e.g. Intel Q3 2026 Rate Survey"/>
              <Input label="Client Name (optional)" value={newClient} onChange={setNewClient} placeholder="e.g. Intel"/>
              <div style={{display:"flex",gap:8}}>
                <Btn variant="sky" onClick={handleCreate} disabled={!newName.trim()}>Create Project</Btn>
                <Btn variant="ghost" size="sm" onClick={()=>setCreating(false)}>Cancel</Btn>
              </div>
            </div>
          )}

          {loading?(
            <div style={{textAlign:"center",padding:"32px 0",color:C.textMuted,fontSize:14}}>Loading projects…</div>
          ):projects.length===0?(
            <div style={{textAlign:"center",padding:"32px 0",color:C.textMuted,fontSize:14}}>
              <div style={{fontSize:32,marginBottom:8}}>📁</div>
              No projects yet — create your first one above
            </div>
          ):(
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {projects.map(p=>(
                <div key={p.id} style={{
                  display:"flex",alignItems:"center",justifyContent:"space-between",
                  background:activeId===p.id?C.navy:C.slateLight,
                  borderRadius:9,padding:"12px 16px",cursor:"pointer",
                  transition:"background .15s",
                }} onClick={()=>onSelect(p.id)}>
                  <div>
                    <div style={{fontWeight:700,fontSize:14,color:activeId===p.id?C.white:C.text}}>{p.name}</div>
                    <div style={{fontSize:11,color:activeId===p.id?"#93C5FD":C.textMuted,marginTop:2}}>
                      {p.client_name&&`Client: ${p.client_name} · `}Created {new Date(p.created_at).toLocaleDateString()}
                    </div>
                  </div>
                  <div style={{display:"flex",gap:6,alignItems:"center"}}>
                    <Btn size="sm" variant={activeId===p.id?"ghost":"sky"} onClick={e=>{e.stopPropagation();onSelect(p.id);}}>
                      {activeId===p.id?"✓ Active":"Open"}
                    </Btn>
                    <button onClick={e=>{e.stopPropagation();if(window.confirm("Delete this project and all its data? This cannot be undone."))onDelete(p.id);}}
                      style={{background:"none",border:"none",color:activeId===p.id?"#FC8181":C.rose,cursor:"pointer",fontSize:16,padding:"2px 6px"}}>×</button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </Card>
      </div>
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
const TABS=[
  {id:"template",label:"Template Builder",icon:"📋"},
  {id:"tracker", label:"Supplier Tracker", icon:"📬"},
  {id:"data",    label:"Data Entry",       icon:"✏️"},
  {id:"analytics",label:"Analytics",      icon:"📊"},
];

export default function App(){
  const [tab,setTab]=useState("template");
  const [showProjects,setShowProjects]=useState(false);

  // ── Bootstrap from localStorage ──
  const [store,setStoreRaw]=useState(()=>lsLoad());

  function setStore(fn){
    setStoreRaw(prev=>{
      const next=typeof fn==="function"?fn(prev):fn;
      lsSave(next);
      return next;
    });
  }

  const projects=store.projects||[];
  const [activeProjectId,setActiveProjectId]=useState(()=>{
    const s=lsLoad();
    return s.projects&&s.projects.length>0?s.projects[0].id:null;
  });

  const activeProject=projects.find(p=>p.id===activeProjectId)||null;
  const pd=store.projectData||{};
  const pdata=pd[activeProjectId]||{jobs:[],locations:[],suppliers:[],responses:[]};

  const jobs=pdata.jobs||[];
  const locations=pdata.locations||[];
  const suppliers=pdata.suppliers||[];
  const responses=pdata.responses||[];

  function updateProjectData(pid,fn){
    setStore(prev=>{
      const existing=prev.projectData||{};
      const current=existing[pid]||{jobs:[],locations:[],suppliers:[],responses:[]};
      const next=typeof fn==="function"?fn(current):fn;
      return {...prev,projectData:{...existing,[pid]:next}};
    });
  }

  function setJobs(fn){
    if(!activeProjectId)return;
    updateProjectData(activeProjectId,cur=>({...cur,jobs:typeof fn==="function"?fn(cur.jobs||[]):fn}));
  }
  function setLocations(fn){
    if(!activeProjectId)return;
    updateProjectData(activeProjectId,cur=>({...cur,locations:typeof fn==="function"?fn(cur.locations||[]):fn}));
  }
  function setSuppliers(fn){
    if(!activeProjectId)return;
    updateProjectData(activeProjectId,cur=>({...cur,suppliers:typeof fn==="function"?fn(cur.suppliers||[]):fn}));
  }
  function setResponses(fn){
    if(!activeProjectId)return;
    updateProjectData(activeProjectId,cur=>({...cur,responses:typeof fn==="function"?fn(cur.responses||[]):fn}));
  }

  function handleSelectProject(pid){
    setActiveProjectId(pid);
    setShowProjects(false);
    setTab("template");
  }

  function handleCreateProject(name,clientName){
    const newP={id:Date.now().toString(),name,client_name:clientName||null,created_at:new Date().toISOString()};
    setStore(prev=>({
      ...prev,
      projects:[newP,...(prev.projects||[])],
      projectData:{...(prev.projectData||{}),[newP.id]:{jobs:[],locations:[],suppliers:[],responses:[]}},
    }));
    setActiveProjectId(newP.id);
    setShowProjects(false);
    setTab("template");
  }

  function handleDeleteProject(pid){
    setStore(prev=>{
      const projects=(prev.projects||[]).filter(p=>p.id!==pid);
      const projectData={...(prev.projectData||{})};
      delete projectData[pid];
      return {...prev,projects,projectData};
    });
    if(activeProjectId===pid){
      const remaining=(store.projects||[]).filter(p=>p.id!==pid);
      setActiveProjectId(remaining.length>0?remaining[0].id:null);
    }
  }

  const respondedCount=suppliers.filter(s=>s.status==="responded").length;
  const partialCount=suppliers.filter(s=>s.status==="responded").filter(s=>{
    const c=completeness(s.name,responses,jobs,locations);
    return c.pct<100&&c.pct>0;
  }).length;

  if(!activeProjectId||showProjects){
    return <ProjectSelector
      projects={projects} activeId={activeProjectId}
      onSelect={handleSelectProject} onCreate={handleCreateProject}
      onDelete={handleDeleteProject} loading={false}/>;
  }

  return(
    <div style={{minHeight:"100vh",background:"#F0F4F9",fontFamily:"'Inter',system-ui,sans-serif",color:C.text}}>
      <div style={{background:C.navy,padding:"0 32px"}}>
        <div style={{maxWidth:1160,margin:"0 auto"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"16px 0 0"}}>
            <div style={{display:"flex",alignItems:"center",gap:12}}>
              <button onClick={()=>setShowProjects(true)} style={{background:"rgba(255,255,255,.1)",border:"none",borderRadius:7,padding:"6px 10px",cursor:"pointer",color:"#93C5FD",fontSize:12}}>← Projects</button>
              <div>
                <div style={{color:C.white,fontSize:18,fontWeight:800,letterSpacing:-0.5}}>⚡ {activeProject?.name||"SupplierRate"}</div>
                <div style={{color:"#93C5FD",fontSize:11,marginTop:1}}>
                  {activeProject?.client_name?"Client: "+activeProject.client_name+" · ":""}Kelly Services MRA
                </div>
              </div>
            </div>
            <div style={{display:"flex",gap:16,alignItems:"center"}}>
              {partialCount>0&&<span style={{fontSize:12,color:C.amber,fontWeight:600}}>⚠️ {partialCount} partial</span>}
              <span style={{fontSize:12,color:"#93C5FD"}}>{respondedCount}/{suppliers.length} responded · {responses.length} records</span>
              <span style={{fontSize:11,color:"#6EE7B7",background:"rgba(45,189,142,.15)",padding:"3px 10px",borderRadius:99}}>💾 Auto-saved</span>
            </div>
          </div>
          <div style={{display:"flex",gap:2,marginTop:16}}>
            {TABS.map(t=>(
              <button key={t.id} onClick={()=>setTab(t.id)} style={{
                padding:"10px 20px",border:"none",cursor:"pointer",fontSize:13,fontWeight:600,
                borderRadius:"8px 8px 0 0",transition:"background .15s",
                background:tab===t.id?"#F0F4F9":"transparent",
                color:tab===t.id?C.navy:"#93C5FD",
              }}>{t.icon} {t.label}</button>
            ))}
          </div>
        </div>
      </div>
      <div style={{maxWidth:1160,margin:"0 auto",padding:"24px 32px 48px"}}>
        {tab==="template"&&<TemplateBuilder jobs={jobs} setJobs={setJobs} locations={locations} setLocations={setLocations} clientName={activeProject?.client_name||""}/>}
        {tab==="tracker"&&<SupplierTracker suppliers={suppliers} setSuppliers={setSuppliers} responses={responses} jobs={jobs} locations={locations}/>}
        {tab==="data"&&<DataEntry responses={responses} setResponses={setResponses} suppliers={suppliers}/>}
        {tab==="analytics"&&<Analytics responses={responses} suppliers={suppliers} jobs={jobs} locations={locations}/>}
      </div>
    </div>
  );
}
