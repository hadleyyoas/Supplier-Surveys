import { useState, useRef } from "react";
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

// ─── Sample data ──────────────────────────────────────────────────────────────
const INIT_JOBS=[
  {id:1,title:"Software Engineer",level:"Junior"},
  {id:2,title:"Software Engineer",level:"Mid"},
  {id:3,title:"Software Engineer",level:"Senior"},
  {id:4,title:"Project Manager",level:"Mid"},
  {id:5,title:"Project Manager",level:"Senior"},
  {id:6,title:"Business Analyst",level:"Junior"},
  {id:7,title:"Business Analyst",level:"Senior"},
];
const INIT_LOCS=["Chicago, IL","New York, NY","Dallas, TX","Remote"];
const INIT_SUPPLIERS=[
  {id:1,name:"Apex Staffing",  contact:"sarah@apexstaff.com", status:"responded",sentAt:"2025-06-01",respondedAt:"2025-06-03"},
  {id:2,name:"TalentBridge",   contact:"mark@talentbridge.com",status:"sent",    sentAt:"2025-06-01",respondedAt:null},
  {id:3,name:"ProSource Group",contact:"linda@prosource.com",  status:"responded",sentAt:"2025-06-01",respondedAt:"2025-06-04"},
  {id:4,name:"EliteForce",     contact:"james@eliteforce.com", status:"not_sent", sentAt:null,respondedAt:null},
  {id:5,name:"PeakHire",       contact:"anna@peakhire.com",    status:"sent",    sentAt:"2025-06-02",respondedAt:null},
  {id:6,name:"CoreStaff",      contact:"tom@corestaff.com",    status:"follow_up",sentAt:"2025-06-01",respondedAt:null},
];
const INIT_RESPONSES=[
  {supplier:"Apex Staffing",  title:"Software Engineer",level:"Mid",   location:"Chicago, IL",  billRate:95, payRate:72},
  {supplier:"Apex Staffing",  title:"Software Engineer",level:"Senior",location:"Chicago, IL",  billRate:130,payRate:98},
  {supplier:"Apex Staffing",  title:"Project Manager",  level:"Senior",location:"New York, NY", billRate:145,payRate:110},
  {supplier:"Apex Staffing",  title:"Business Analyst", level:"Junior",location:"Remote",       billRate:65, payRate:50},
  {supplier:"ProSource Group",title:"Software Engineer",level:"Mid",   location:"Chicago, IL",  billRate:88, payRate:68},
  {supplier:"ProSource Group",title:"Software Engineer",level:"Senior",location:"Chicago, IL",  billRate:125,payRate:95},
  {supplier:"ProSource Group",title:"Project Manager",  level:"Senior",location:"New York, NY", billRate:138,payRate:105},
  {supplier:"ProSource Group",title:"Business Analyst", level:"Junior",location:"Remote",       billRate:70, payRate:52},
  {supplier:"TalentBridge",   title:"Software Engineer",level:"Senior",location:"Chicago, IL",  billRate:165,payRate:118},
  {supplier:"TalentBridge",   title:"Project Manager",  level:"Senior",location:"New York, NY", billRate:195,payRate:145},
];

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

// Completeness: what % of expected role×location combos did a supplier fill in?
function completeness(supplierName, responses, jobs, locations){
  const expected = jobs.length * locations.length;
  if(expected===0) return {pct:0,filled:0,expected:0,missing:[]};
  const filled = new Set(
    responses.filter(r=>r.supplier===supplierName)
      .map(r=>`${r.title}|${r.level}|${r.location}`)
  );
  const missing = [];
  jobs.forEach(j=>{
    locations.forEach(l=>{
      const key=`${j.title}|${j.level}|${l}`;
      if(!filled.has(key)) missing.push(`${j.title} ${j.level} – ${l}`);
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

// ─── AI helpers ───────────────────────────────────────────────────────────────
async function callClaude(prompt, maxTokens=1000){
  const res = await fetch("https://api.anthropic.com/v1/messages",{
    method:"POST",headers:{"Content-Type":"application/json"},
    body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:maxTokens,messages:[{role:"user",content:prompt}]}),
  });
  const data = await res.json();
  return data.content?.find(b=>b.type==="text")?.text||"";
}

// Microsoft Graph API helper — sends via Outlook on behalf of SENDER_EMAIL
async function sendOutlookEmail(toEmail, subject, body){
  // Uses the Microsoft 365 MCP connector which handles auth automatically
  const res = await fetch("https://microsoft365.mcp.claude.com/mcp", {
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body: JSON.stringify({
      tool: "send_email",
      input: {
        to: toEmail,
        subject: subject,
        body: body,
        from: SENDER_EMAIL,
        bodyType: "text",
      }
    }),
  });
  if(!res.ok) throw new Error(`Send failed: ${res.status}`);
  return true;
}

// ─── TEMPLATE BUILDER ────────────────────────────────────────────────────────
function TemplateBuilder({jobs,setJobs,locations,setLocations}){
  const [newTitle,setNewTitle]=useState("");
  const [newLevel,setNewLevel]=useState("Mid");
  const [newLoc,setNewLoc]=useState("");
  const [generating,setGenerating]=useState(false);
  const [preview,setPreview]=useState("");
  const [importMsg,setImportMsg]=useState("");

  const levels=["Junior","Mid","Senior","Lead","Principal","Manager","Director"];

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
        let tc=-1,lc=-1,loc=-1,hr=-1;
        for(let r=0;r<Math.min(6,rows.length);r++){
          const row=rows[r].map(v=>String(v).toLowerCase().trim());
          const ti=row.findIndex(c=>c.includes("title")||c.includes("role")||c.includes("job"));
          if(ti>=0){tc=ti;lc=row.findIndex(c=>c.includes("level")||c.includes("grade"));loc=row.findIndex(c=>c.includes("location")||c.includes("city"));hr=r;break;}
        }
        if(tc<0){setImportMsg("⚠️ Couldn't find a Title/Role column.");return;}
        const newJobs=[];const newLocs=new Set();
        for(let r=hr+1;r<rows.length;r++){
          const row=rows[r];
          const title=String(row[tc]||"").trim();
          if(!title)continue;
          const level=lc>=0?String(row[lc]||"").trim():"Mid";
          newJobs.push({id:Date.now()+r,title,level});
          if(loc>=0&&row[loc])newLocs.add(String(row[loc]).trim());
        }
        if(!newJobs.length){setImportMsg("⚠️ No role rows found.");return;}
        setJobs(newJobs);
        if(newLocs.size)setLocations([...newLocs]);
        setImportMsg(`✅ Imported ${newJobs.length} roles${newLocs.size?` and ${newLocs.size} locations`:""}`);
      }catch{setImportMsg("⚠️ Error reading file.");}
    };
    reader.readAsBinaryString(file);
  }

  async function generateTemplate(){
    setGenerating(true);setPreview("");
    try{
      const text = await callClaude(`Write a professional supplier rate survey email.
Roles (title + level): ${jobs.map(j=>`${j.title} - ${j.level}`).join(", ")}.
Locations: ${locations.join(", ")}.
Ask them to fill in hourly bill and pay rates per role per location.
Concise, friendly. Placeholders: [SUPPLIER NAME], [YOUR NAME].`);
      setPreview(text);
    }catch{setPreview("Error generating. Please try again.");}
    setGenerating(false);
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
            <div style={{fontWeight:700,fontSize:15,color:C.navy}}>✉️ Email Template Preview</div>
            <Btn onClick={generateTemplate} variant="mint" size="sm" disabled={generating}>{generating?"Generating…":"Generate with AI"}</Btn>
          </div>
          {preview?(
            <textarea value={preview} onChange={e=>setPreview(e.target.value)}
              style={{width:"100%",minHeight:300,border:`1px solid ${C.border}`,borderRadius:8,padding:12,fontSize:13,color:C.text,lineHeight:1.6,resize:"vertical",fontFamily:"inherit",boxSizing:"border-box"}}/>
          ):(
            <div style={{minHeight:300,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",color:C.textMuted,fontSize:14,gap:8,border:`2px dashed ${C.border}`,borderRadius:8}}>
              <span style={{fontSize:32}}>✨</span>
              <span>Click "Generate with AI" to draft your survey email</span>
              <span style={{fontSize:12}}>{jobs.length} roles · {locations.length} locations configured</span>
            </div>
          )}
        </Card>
        <Card>
          <div style={{fontWeight:700,fontSize:15,color:C.navy,marginBottom:12}}>📊 Survey Grid Preview</div>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead>
                <tr style={{background:C.navy}}>
                  <th style={{padding:"6px 10px",color:C.white,textAlign:"left"}}>Role</th>
                  {locations.slice(0,3).map(l=><th key={l} style={{padding:"6px 8px",color:C.white,textAlign:"center"}} colSpan={2}>{l}</th>)}
                </tr>
                <tr style={{background:C.navyLight}}>
                  <th></th>
                  {locations.slice(0,3).map(l=>(
                    <>
                      <th key={l+"b"} style={{padding:"4px 8px",color:"#93C5FD",textAlign:"center",fontSize:11}}>Bill $</th>
                      <th key={l+"p"} style={{padding:"4px 8px",color:"#93C5FD",textAlign:"center",fontSize:11}}>Pay $</th>
                    </>
                  ))}
                </tr>
              </thead>
              <tbody>
                {jobs.slice(0,5).map((j,i)=>(
                  <tr key={j.id} style={{background:i%2===0?C.white:C.slateLight}}>
                    <td style={{padding:"5px 10px",whiteSpace:"nowrap"}}>{j.title} – {j.level}</td>
                    {locations.slice(0,3).map(l=>(
                      <>
                        <td key={l+"b"} style={{padding:"5px 8px",textAlign:"center",color:C.textMuted}}>—</td>
                        <td key={l+"p"} style={{padding:"5px 8px",textAlign:"center",color:C.textMuted}}>—</td>
                      </>
                    ))}
                  </tr>
                ))}
                {jobs.length>5&&<tr><td colSpan={1+locations.slice(0,3).length*2} style={{padding:"5px 10px",color:C.textMuted,fontSize:12}}>+{jobs.length-5} more roles…</td></tr>}
              </tbody>
            </table>
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
  const [draftInfo,setDraftInfo]=useState(null);
  const [draft,setDraft]=useState("");
  const [draftSubject,setDraftSubject]=useState("");
  const [drafting,setDrafting]=useState(false);
  const [sending,setSending]=useState(false);
  const [sendStatus,setSendStatus]=useState({});
  const [importMsg,setImportMsg]=useState("");
  const [expandedCompletion,setExpandedCompletion]=useState(null);

  function addSupplier(){
    if(!newName.trim())return;
    setSuppliers(p=>[...p,{id:Date.now(),name:newName.trim(),contact:newEmail.trim(),status:"not_sent",sentAt:null,respondedAt:null}]);
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
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        let nc=-1,ec=-1,hr=-1;
        for(let r=0;r<Math.min(5,rows.length);r++){
          const row=rows[r].map(v=>String(v).toLowerCase().trim());
          const ni=row.findIndex(c=>c.includes("name")||c.includes("supplier")||c.includes("company")||c.includes("vendor"));
          if(ni>=0){nc=ni;ec=row.findIndex(c=>c.includes("email")||c.includes("contact")||c.includes("mail"));hr=r;break;}
        }
        if(nc<0){setImportMsg("⚠️ Couldn't find a Name/Supplier column.");return;}
        const imported=[];
        for(let r=hr+1;r<rows.length;r++){
          const name=String(rows[r][nc]||"").trim();
          if(!name)continue;
          const contact=ec>=0?String(rows[r][ec]||"").trim():"";
          imported.push({id:Date.now()+r,name,contact,status:"not_sent",sentAt:null,respondedAt:null});
        }
        if(!imported.length){setImportMsg("⚠️ No supplier rows found.");return;}
        setSuppliers(prev=>{
          const existing=new Set(prev.map(s=>s.name.toLowerCase()));
          const fresh=imported.filter(s=>!existing.has(s.name.toLowerCase()));
          setImportMsg(`✅ Added ${fresh.length} suppliers (${imported.length-fresh.length} duplicates skipped)`);
          return [...prev,...fresh];
        });
      }catch{setImportMsg("⚠️ Error reading file.");}
    };
    reader.readAsBinaryString(file);
  }

  async function draftEmail(supplier,type){
    setDraftInfo({supplier,type});setDrafting(true);setDraft("");setDraftSubject("");
    const roleList = jobs.map(j=>`${j.title} - ${j.level}`).join(", ");
    const locList = locations.join(", ");
    const prompts={
      initial:`Write a professional initial outreach email to staffing supplier ${supplier.name} asking them to complete a rate survey for hourly bill and pay rates.
Roles to survey: ${roleList}.
Locations: ${locList}.
Keep it warm and concise. The sender is Hadley Yoas from Kelly Services.
Return in this exact format:
SUBJECT: [subject line here]
BODY:
[email body here]`,
      followup:`Write a brief friendly follow-up email to ${supplier.name} who hasn't yet responded to our rate survey. Not pushy. Sender is Hadley Yoas from Kelly Services.
Return:
SUBJECT: [subject line]
BODY:
[body]`,
      final:`Write a final notice email to ${supplier.name} — rate survey closes in 48 hours. Polite urgency. Sender is Hadley Yoas from Kelly Services.
Return:
SUBJECT: [subject line]
BODY:
[body]`,
    };
    try{
      const text = await callClaude(prompts[type], 800);
      const subjectMatch = text.match(/SUBJECT:\s*(.+)/);
      const bodyMatch = text.match(/BODY:\s*([\s\S]+)/);
      setDraftSubject(subjectMatch?subjectMatch[1].trim():"Rate Survey Request");
      setDraft(bodyMatch?bodyMatch[1].trim():text);
    }catch{setDraft("Error generating email.");}
    setDrafting(false);
  }

  async function sendEmail(supplier){
    if(!supplier.contact){
      setSendStatus(p=>({...p,[supplier.id]:"⚠️ No email address on file for this supplier."}));
      return;
    }
    setSending(true);
    setSendStatus(p=>({...p,[supplier.id]:"📨 Sending…"}));
    try{
      await sendOutlookEmail(supplier.contact, draftSubject, draft);
      setSendStatus(p=>({...p,[supplier.id]:`✅ Sent to ${supplier.contact}`}));
      updateStatus(supplier.id, draftInfo.type==="initial"?"sent":draftInfo.type==="followup"?"follow_up":supplier.status);
      setTimeout(()=>setDraftInfo(null),1500);
    }catch(err){
      setSendStatus(p=>({...p,[supplier.id]:`⚠️ Send failed — check your Microsoft 365 connection. (${err.message})`}));
    }
    setSending(false);
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

      {/* Outlook sender badge */}
      <Card style={{padding:"12px 20px",background:C.purpleLight,border:`1px solid #DDD6FE`}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:20}}>📧</span>
          <div>
            <div style={{fontWeight:700,fontSize:13,color:C.purple}}>Outlook Connected</div>
            <div style={{fontSize:12,color:C.slate}}>Emails will be sent from <strong>{SENDER_EMAIL}</strong> via your Microsoft 365 account</div>
          </div>
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
              {["Supplier","Contact","Status","Completeness","Actions"].map(h=>(
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
                    <td style={{padding:"9px 10px",fontWeight:600}}>
                      {s.name}
                      {s.sentAt&&<div style={{fontSize:10,color:C.textMuted,marginTop:2}}>Sent {s.sentAt}</div>}
                      {s.respondedAt&&<div style={{fontSize:10,color:C.mint,marginTop:1}}>Responded {s.respondedAt}</div>}
                    </td>
                    <td style={{padding:"9px 10px",color:C.textMuted,fontSize:12}}>{s.contact||<span style={{color:C.rose,fontSize:11}}>No email</span>}</td>
                    <td style={{padding:"9px 10px"}}><StatusBadge status={s.status}/></td>
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
                    <td style={{padding:"9px 10px"}}>
                      <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                        {s.status==="not_sent"&&<Btn size="sm" variant="sky" onClick={()=>draftEmail(s,"initial")}>Draft & Send</Btn>}
                        {s.status==="sent"&&<Btn size="sm" variant="amber" onClick={()=>draftEmail(s,"followup")}>Follow-up</Btn>}
                        {s.status==="follow_up"&&<Btn size="sm" variant="danger" onClick={()=>draftEmail(s,"final")}>Final Notice</Btn>}
                        {s.status!=="responded"&&<Btn size="sm" variant="mint" onClick={()=>updateStatus(s.id,"responded")}>✓ Mark Responded</Btn>}
                      </div>
                    </td>
                  </tr>
                  {isExpanded&&comp.missing.length>0&&(
                    <tr key={s.id+"_missing"} style={{borderBottom:`1px solid ${C.border}`}}>
                      <td colSpan={5} style={{padding:"0 10px 10px 10px"}}>
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

      {/* Email draft panel */}
      {draftInfo&&(
        <Card style={{border:`2px solid ${C.purple}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div>
              <div style={{fontWeight:700,color:C.navy,fontSize:15}}>
                ✉️ {draftInfo.type==="initial"?"Initial Outreach":draftInfo.type==="followup"?"Follow-up":"Final Notice"} → {draftInfo.supplier.name}
              </div>
              <div style={{fontSize:12,color:C.textMuted,marginTop:2}}>Sending from {SENDER_EMAIL}</div>
            </div>
            <Btn size="sm" variant="ghost" onClick={()=>setDraftInfo(null)}>Dismiss</Btn>
          </div>

          {drafting?(
            <div style={{color:C.textMuted,fontSize:13,padding:"24px 0",textAlign:"center"}}>Generating email draft…</div>
          ):(
            <>
              <Input label="Subject" value={draftSubject} onChange={setDraftSubject} style={{marginBottom:10}}/>
              <div style={{fontSize:12,fontWeight:600,color:C.textMuted,marginBottom:4}}>Body</div>
              <textarea value={draft} onChange={e=>setDraft(e.target.value)}
                style={{width:"100%",minHeight:220,border:`1px solid ${C.border}`,borderRadius:8,padding:12,fontSize:13,lineHeight:1.6,color:C.text,resize:"vertical",fontFamily:"inherit",boxSizing:"border-box"}}/>
              <Toast msg={sendStatus[draftInfo.supplier.id]}/>
              <div style={{display:"flex",gap:8,marginTop:12,alignItems:"center"}}>
                <Btn variant="purple" onClick={()=>sendEmail(draftInfo.supplier)} disabled={sending||!draft.trim()}>
                  {sending?"Sending…":"📧 Send via Outlook"}
                </Btn>
                <Btn size="sm" variant="ghost" onClick={()=>navigator.clipboard?.writeText(draft)}>Copy to Clipboard</Btn>
                <Btn size="sm" variant="ghost" onClick={()=>setDraftInfo(null)}>Close</Btn>
                {!draftInfo.supplier.contact&&(
                  <span style={{fontSize:12,color:C.rose,marginLeft:4}}>⚠️ No email address — update the supplier contact before sending.</span>
                )}
              </div>
            </>
          )}
        </Card>
      )}
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
        let cols={title:-1,level:-1,location:-1,bill:-1,pay:-1};let hr=-1;
        for(let r=0;r<Math.min(8,rows.length);r++){
          const row=rows[r].map(v=>String(v).toLowerCase().trim());
          const ti=row.findIndex(c=>c.includes("title")||c.includes("role")||c.includes("job")||c.includes("position"));
          if(ti>=0){
            cols.title=ti;
            cols.level=row.findIndex(c=>c.includes("level")||c.includes("grade")||c.includes("tier")||c.includes("category"));
            cols.location=row.findIndex(c=>c.includes("location")||c.includes("city")||c.includes("region")||c.includes("market"));
            cols.bill=row.findIndex(c=>c.includes("bill")||c.includes("markup")||c.includes("charge")||c.includes("client"));
            cols.pay=row.findIndex(c=>c.includes("pay")||c.includes("cost")||c.includes("wage")||c.includes("salary")||c.includes("worker"));
            hr=r;break;
          }
        }
        if(cols.title<0){setImportMsg("⚠️ Couldn't find a Title/Role column.");return;}
        const parsed=[];
        for(let r=hr+1;r<rows.length;r++){
          const row=rows[r];
          const title=String(row[cols.title]||"").trim();
          if(!title)continue;
          parsed.push({
            supplier:selectedSupplier,title,
            level:cols.level>=0?String(row[cols.level]||"").trim():"",
            location:cols.location>=0?String(row[cols.location]||"").trim():"",
            billRate:cols.bill>=0?parseFloat(String(row[cols.bill]).replace(/[^0-9.]/g,""))||0:0,
            payRate:cols.pay>=0?parseFloat(String(row[cols.pay]).replace(/[^0-9.]/g,""))||0:0,
          });
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
  const [aiInsight,setAiInsight]=useState("");
  const [loading,setLoading]=useState(false);
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

  const roles=[...new Set(workingData.map(r=>`${r.title} – ${r.level}`))];
  const summaries=roles.map(role=>{
    const [title,level]=role.split(" – ");
    const recs=workingData.filter(r=>r.title===title&&r.level===level);
    const bills=recs.map(r=>r.billRate).filter(Boolean);
    const pays=recs.map(r=>r.payRate).filter(Boolean);
    return{role,responses:recs.length,
      billAvg:avg(bills),billMin:bills.length?Math.min(...bills):0,billMax:bills.length?Math.max(...bills):0,
      payAvg:avg(pays),payMin:pays.length?Math.min(...pays):0,payMax:pays.length?Math.max(...pays):0,
      spread:bills.length?Math.max(...bills)-Math.min(...bills):0,
    };
  }).sort((a,b)=>b.billAvg-a.billAvg);

  function exportCSV(){
    const headers=["Supplier","Job Title","Level","Location","Bill Rate","Pay Rate","Bill Outlier","Pay Outlier"];
    const rows=locFiltered.map(r=>{
      const k=`${r.supplier}|${r.title}|${r.level}|${r.location}`;
      const od=outlierDetails.find(o=>`${o.supplier}|${o.title}|${o.level}|${o.location}`===k);
      return[r.supplier,r.title,r.level,r.location,r.billRate,r.payRate,od?.billOutlier?"Yes":"No",od?.payOutlier?"Yes":"No"];
    });
    const csv=[headers,...rows].map(r=>r.map(v=>`"${v}"`).join(",")).join("\n");
    const blob=new Blob([csv],{type:"text/csv"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");a.href=url;a.download="supplier_rates_analysis.csv";a.click();
    URL.revokeObjectURL(url);
  }

  async function getInsights(){
    setLoading(true);setAiInsight("");
    const outlierSummary=outlierDetails.map(o=>
      `${o.supplier}: ${o.title} ${o.level}${o.location?` (${o.location})`:""}` +
      (o.billOutlier?` bill=$${o.billRate} vs avg $${o.avgBill}`:"") +
      (o.payOutlier?` pay=$${o.payRate} vs avg $${o.avgPay}`:"")
    ).join("; ");

    // Build completeness summary for context
    const compSummary=suppliers.filter(s=>s.status==="responded").map(s=>{
      const c=completeness(s.name,responses,jobs,locations);
      return `${s.name}: ${c.pct}% complete (${c.filled}/${c.expected} roles)${c.missing.length?`, missing: ${c.missing.slice(0,3).join(", ")}${c.missing.length>3?"…":""}` : ""}`;
    }).join("; ");

    try{
      const text = await callClaude(`Analyze this supplier rate survey data. Write 4-5 concise insights as plain paragraphs — no bullet points, no headers, no bold.

Rate summaries by role: ${JSON.stringify(summaries)}
Outlier details: ${outlierSummary||"None detected"}
Response completeness: ${compSummary||"N/A"}

Focus on:
1. Named outliers — supplier, role, exact rate vs group average
2. Roles with widest spreads, which suppliers are at the extremes
3. Any suppliers with incomplete responses and what they're missing
4. Negotiation leverage opportunities
5. Any notable patterns

Use specific dollar amounts and supplier names throughout.`, 1200);
      setAiInsight(text);
    }catch{setAiInsight("Could not load insights.");}
    setLoading(false);
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
                  <td style={{padding:"8px 12px",fontWeight:600}}>{s.role}</td>
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

      <Card style={{border:`2px solid ${C.skyLight}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <div style={{fontWeight:700,fontSize:15,color:C.navy}}>✨ AI Rate Insights</div>
          <Btn onClick={getInsights} variant="sky" size="sm" disabled={loading}>{loading?"Analyzing…":"Analyze Data"}</Btn>
        </div>
        {aiInsight?(
          <div style={{fontSize:13,color:C.text,lineHeight:1.8,whiteSpace:"pre-wrap"}}>{aiInsight}</div>
        ):(
          <div style={{color:C.textMuted,fontSize:13,lineHeight:1.6}}>
            Click "Analyze Data" for AI insights — named outliers with supplier and dollar values, completeness gaps, spreads, and negotiation opportunities.
            {outlierDetails.length>0&&<span style={{color:C.amber}}> {outlierDetails.length} outlier(s) flagged.</span>}
          </div>
        )}
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

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
const TABS=[
  {id:"template",label:"Template Builder",icon:"📋"},
  {id:"tracker", label:"Supplier Tracker", icon:"📬"},
  {id:"data",    label:"Data Entry",       icon:"✏️"},
  {id:"analytics",label:"Analytics",      icon:"📊"},
];

export default function App(){
  const [tab,setTab]=useState("template");
  const [jobs,setJobs]=useState(INIT_JOBS);
  const [locations,setLocations]=useState(INIT_LOCS);
  const [suppliers,setSuppliers]=useState(INIT_SUPPLIERS);
  const [responses,setResponses]=useState(INIT_RESPONSES);

  const respondedCount=suppliers.filter(s=>s.status==="responded").length;
  const partialCount=suppliers.filter(s=>s.status==="responded").filter(s=>{
    const c=completeness(s.name,responses,jobs,locations);
    return c.pct<100&&c.pct>0;
  }).length;

  return(
    <div style={{minHeight:"100vh",background:"#F0F4F9",fontFamily:"'Inter',system-ui,sans-serif",color:C.text}}>
      <div style={{background:C.navy,padding:"0 32px"}}>
        <div style={{maxWidth:1160,margin:"0 auto"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"16px 0 0"}}>
            <div>
              <div style={{color:C.white,fontSize:20,fontWeight:800,letterSpacing:-0.5}}>⚡ SupplierRate</div>
              <div style={{color:"#93C5FD",fontSize:12,marginTop:1}}>Rate Survey Management · Kelly Services MRA</div>
            </div>
            <div style={{display:"flex",gap:16,alignItems:"center"}}>
              {partialCount>0&&<span style={{fontSize:12,color:C.amber,fontWeight:600}}>⚠️ {partialCount} partial response{partialCount!==1?"s":""}</span>}
              <span style={{fontSize:12,color:"#93C5FD"}}>{respondedCount}/{suppliers.length} responded · {responses.length} records</span>
              <span style={{fontSize:11,color:"#6EE7B7",background:"rgba(45,189,142,.15)",padding:"3px 10px",borderRadius:99}}>📧 Outlook Connected</span>
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
        {tab==="template"&&<TemplateBuilder jobs={jobs} setJobs={setJobs} locations={locations} setLocations={setLocations}/>}
        {tab==="tracker"&&<SupplierTracker suppliers={suppliers} setSuppliers={setSuppliers} responses={responses} jobs={jobs} locations={locations}/>}
        {tab==="data"&&<DataEntry responses={responses} setResponses={setResponses} suppliers={suppliers}/>}
        {tab==="analytics"&&<Analytics responses={responses} suppliers={suppliers} jobs={jobs} locations={locations}/>}
      </div>
    </div>
  );
}
