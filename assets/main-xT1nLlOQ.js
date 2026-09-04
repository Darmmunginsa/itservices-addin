import{P as Fe}from"./PublicClientApplication-DLKYUtZW.js";import{a as te,d as ze,p as qe,L as Se,S as Ce}from"./analyzer-CWd3MChg.js";const Ue=/^\s*(Sent|To|Date|Cc|Subject|ส่ง|ถึง|วันที่|สำเนา|เรื่อง)\s*:/i;function Ve(e,t){const n=e[t];if(/^\s*-{2,}\s*(Original Message|Forwarded message|ข้อความต้นฉบับ)\s*-{2,}/i.test(n)||/^\s*_{5,}\s*$/.test(n)||/^\s*(On|เมื่อ)\b.{10,200}(wrote|เขียนว่า)\s*:\s*$/i.test(n))return!0;if(/^\s*(From|จาก)\s*:\s*\S/i.test(n)){for(let a=t+1;a<=t+3&&a<e.length;a++)if(Ue.test(e[a]))return!0;return!1}return!1}function Ge(e){const t=(e??"").replace(/\r\n/g,`
`),n=t.split(`
`);let a=-1;for(let s=0;s<n.length;s++){if(s>0&&Ve(n,s)){a=s;break}if(s>0&&/^\s*>/.test(n[s])&&/^\s*>/.test(n[s+1]??"")){a=s;break}}if(a<0)return{visible:t.trim(),quoted:""};const o=n.slice(0,a).join(`
`).trim(),r=n.slice(a).join(`
`).trim();return o?{visible:o,quoted:r}:{visible:t.trim(),quoted:""}}const Ke=e=>Ge(e).visible,xe=e=>(e??"").trim().toLowerCase();function We(e){const t=xe(e.actorEmail),n=new Set,a=s=>{const c=[];for(const p of s){const v=(p??"").trim(),l=xe(v);!l||!l.includes("@")||l===t||n.has(l)||(n.add(l),c.push(v))}return c};let o=a([e.assignedEmail]);const r=a([e.requesterEmail,...e.watchers??[]]);return o.length===0?(o=r.slice(0,1),{to:o,cc:r.slice(1)}):{to:o,cc:r}}const Je=e=>{if(!e||e<=0)return"";if(e<24)return`${e} ชั่วโมง`;const t=e/24;return Number.isInteger(t)?`${t} วัน`:`${e} ชั่วโมง`};function Ye(e){const t=(e.baseUrl??"").replace(/\/+$/,"");return{incident_title:e.title??"",severity:e.severity??"",status:e.status??"",description:e.description??"",resolution:e.resolution??"",incident_date:(e.incidentDate??"").slice(0,10),sla_hours:Je(e.slaHours),project_name:e.projectName??"",assigned_name:e.assignedName||e.assignedEmail||"-",link:e.projectId?`${t}/#/projects/${e.projectId}`:t}}function Ze(e){const{to:t,cc:n}=We(e);return{to:t,cc:n,vars:Ye(e)}}const Qe=[{hours:1,labelTh:"1 ชั่วโมง"},{hours:2,labelTh:"2 ชั่วโมง"},{hours:4,labelTh:"4 ชั่วโมง"},{hours:8,labelTh:"8 ชั่วโมง (1 วันทำการ)"},{hours:24,labelTh:"24 ชั่วโมง"},{hours:48,labelTh:"2 วัน"},{hours:72,labelTh:"3 วัน"},{hours:168,labelTh:"7 วัน"}],ne={Critical:1,High:4,Medium:24,Low:72};function Be(e,t=new Date){const n=typeof e=="number"&&Number.isFinite(e)&&e>0?e:null;return n?new Date(t.getTime()+n*36e5).toISOString():null}function ve(e){const t=Be(e);return t?new Date(t).toLocaleString("th-TH",{dateStyle:"short",timeStyle:"short"}):""}const je="HD_PhishingReports";let f;const m={mail:null,analysis:null,loading:!1,reporting:!1,reported:!1,showHeaders:!1,kasmTemplate:"",analysedItemId:"",safeDomains:[],safeDomainIds:{},savingDomain:""};function Xe(e){f=e}const F=e=>(e??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"),W=()=>{var e,t;return(t=(e=Office.context)==null?void 0:e.mailbox)==null?void 0:t.item},$e=e=>new Promise(t=>{const n=W();if(!(n!=null&&n.body)){t("");return}n.body.getAsync(e,a=>t(a.status===Office.AsyncResultStatus.Succeeded?a.value??"":""))});function et(e){return e?e.split(",").map(t=>{const n=t.trim(),a=n.match(/^(.*?)\s*<([^>]+)>$/);return a?{name:a[1].replace(/^"|"$/g,"").trim(),email:a[2].trim()}:{name:"",email:n.replace(/[<>]/g,"").trim()}}).filter(t=>t.email.includes("@")):[]}function tt(){return new Promise(e=>{let t=!1;try{t=Office.context.requirements.isSetSupported("Mailbox","1.8")}catch{t=!1}const n=W();if(!t||typeof(n==null?void 0:n.getAllInternetHeadersAsync)!="function"){e({});return}try{n.getAllInternetHeadersAsync(a=>e(a.status===Office.AsyncResultStatus.Succeeded?qe(a.value??""):{}))}catch{e({})}})}async function nt(){try{const e=W();if(!(e!=null&&e.itemId))return{};const t=Office.context.mailbox.convertToRestId(e.itemId,Office.MailboxEnums.RestVersion.v2_0),n=await f.getGraphToken(),a=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${t}?$select=internetMessageHeaders`,{headers:{Authorization:`Bearer ${n}`}});if(!a.ok)return{};const o=await a.json(),r={};for(const s of o.internetMessageHeaders??[])r[s.name]=r[s.name]?`${r[s.name]}
${s.value}`:s.value;return r}catch{return{}}}async function pe(e,t){const n=await f.getToken(),a=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('${e}')/items?${t}`,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});return a.ok?(await a.json()).value:[]}async function at(){try{return(await pe("HD_AgentProfiles","$select=Title,EmailText&$top=500")).filter(t=>t.EmailText).map(t=>({name:t.Title,email:t.EmailText}))}catch{return[]}}async function st(){var e;try{return(((e=(await pe("HD_Options","$select=Title&$filter=Category eq 'KasmConfig'&$top=1"))[0])==null?void 0:e.Title)??"").trim()}catch{return""}}const Pe="SafeDomain";async function fe(){try{const e=await pe("HD_Options",`$select=Id,Title&$filter=Category eq '${Pe}'&$top=500`),t={},n=[];for(const a of e){const o=(a.Title??"").trim().toLowerCase();o&&(t[o]=a.Id,n.push(o))}m.safeDomains=n,m.safeDomainIds=t}catch{}}async function ot(e){const t=e.trim().toLowerCase();if(!(!t||m.savingDomain)){if(!f.canWhitelist()){f.toast("ต้องเป็น Agent ขึ้นไปจึงจะยืนยันโดเมนได้","error");return}m.savingDomain=t,f.rerender();try{const n=await f.getToken(),a=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items`,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:t,Category:Pe})});if(!a.ok)throw new Error(String(a.status));await fe(),await ee(!0),f.toast(`ยืนยันแล้วว่า ${t} ปลอดภัย`)}catch{f.toast("บันทึกไม่สำเร็จ","error")}finally{m.savingDomain="",f.rerender()}}}async function it(e){const t=e.trim().toLowerCase(),n=m.safeDomainIds[t];if(!(!n||m.savingDomain)){if(!f.canWhitelist()){f.toast("ต้องเป็น Agent ขึ้นไปจึงจะถอนโดเมนได้","error");return}m.savingDomain=t,f.rerender();try{const a=await f.getToken(),o=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items(${n})`,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","IF-MATCH":"*","X-HTTP-Method":"DELETE"}});if(!o.ok)throw new Error(String(o.status));await fe(),await ee(!0),f.toast(`ถอน ${t} ออกจากรายการปลอดภัยแล้ว`)}catch{f.toast("ถอนไม่สำเร็จ","error")}finally{m.savingDomain="",f.rerender()}}}async function ee(e=!1){var p,v;const t=W(),n=(t==null?void 0:t.itemId)??"";if(!e&&m.analysedItemId===n&&m.analysis)return;m.loading=!0,m.reported=!1,m.analysedItemId=n,f.rerender();const[a,o]=await Promise.all([$e(Office.CoercionType.Html),$e(Office.CoercionType.Text)]),r={fromName:((p=t==null?void 0:t.from)==null?void 0:p.displayName)??"",fromEmail:((v=t==null?void 0:t.from)==null?void 0:v.emailAddress)??"",replyTo:[],subject:(t==null?void 0:t.subject)??"",bodyHtml:a,bodyText:o,attachments:((t==null?void 0:t.attachments)??[]).map(l=>({name:l.name,size:l.size??0,isInline:!!l.isInline})),headers:{},internalDomains:f.internalDomains,internalPeople:[],safeDomains:m.safeDomains},s=l=>{var b;return{...r,headers:l,replyTo:et(((b=Object.entries(l).find(([j])=>j.toLowerCase()==="reply-to"))==null?void 0:b[1])??"")}},c=await tt();if(m.mail=s(c),m.analysis=te(m.mail),m.loading=!1,f.rerender(),f.account()){const[l,b]=await Promise.all([at(),Object.keys(c).length?Promise.resolve(c):nt()]);m.mail={...s(b),internalPeople:l},m.analysis=te(m.mail),m.kasmTemplate||(m.kasmTemplate=await st()),m.safeDomains.length||(await fe(),m.safeDomains.length&&(m.mail={...m.mail,safeDomains:m.safeDomains},m.analysis=te(m.mail))),f.rerender()}}function De(){const e=m.mail,t=m.analysis;return!e||!t?"":[`ผู้ส่ง: ${e.fromName} <${e.fromEmail}>`,`หัวข้อ: ${e.subject}`,e.replyTo.length?`Reply-To: ${e.replyTo.map(n=>n.email).join(", ")}`:"",`คะแนนความเสี่ยง: ${t.score} (${Se[t.level].label})`,"","สิ่งที่ตรวจพบ:",...t.findings.map(n=>`- [${Ce[n.severity].label}] (${n.category}) ${n.title} — ${n.detail.replace(/\n/g," ")}`),"",t.links.length?"ลิงก์ในอีเมล:":"",...t.links.map(n=>`- ${n.href}${n.flags.length?`  ! ${n.flags.join(" / ")}`:""}`)].filter(n=>n!=="").join(`
`)}async function rt(e){try{const t=W();if(!(t!=null&&t.itemId))return!1;const n=Office.context.mailbox.convertToRestId(t.itemId,Office.MailboxEnums.RestVersion.v2_0),a=await f.getGraphToken(),o=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${n}/$value`,{headers:{Authorization:`Bearer ${a}`}});if(!o.ok)return!1;const r=await o.arrayBuffer(),s=(t.subject||"phishing").replace(/[\\/:*?"<>|#%&{}~]/g,"_").replace(/^_+/,"").slice(0,80).trim()||"phishing",c=await f.getToken();return(await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('${je}')/items(${e})/AttachmentFiles/add(FileName='${encodeURIComponent(s+".eml")}')`,{method:"POST",headers:{Authorization:`Bearer ${c}`,Accept:"application/json;odata=nometadata"},body:r})).ok}catch{return!1}}async function we(e){const t=await f.getToken(),n=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('${je}')/items`,{method:"POST",headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(e)});if(!n.ok)throw new Error(`SharePoint ${n.status}: ${await n.text()}`);return(await n.json()).Id}async function ct(){if(!(!m.mail||!m.analysis||m.reporting)){if(!f.account()){f.toast("กรุณาเข้าสู่ระบบก่อนรายงาน","error");return}m.reporting=!0,f.rerender();try{const e=m.mail,t=m.analysis,n=f.account(),a={Title:(e.subject||"(ไม่มีหัวข้อ)").slice(0,255),SenderName:e.fromName.slice(0,255),SenderEmail:e.fromEmail.slice(0,255),SenderDomain:ze(e.fromEmail),RiskScore:t.score,RiskLevel:t.level,Findings:De(),LinkCount:t.links.length,SuspiciousLinks:t.links.filter(c=>c.flags.length).map(c=>c.href).join(`
`).slice(0,4e3),ReportedBy:(n==null?void 0:n.name)??"",ReportedEmail:(n==null?void 0:n.username)??"",Status:"New"};let o,r=!1;try{o=await we(a)}catch(c){o=await we({Title:a.Title,Findings:a.Findings}).catch(()=>{throw c}),r=!0}const s=await rt(o);m.reported=!0,f.toast(r?"ส่งรายงานแล้ว แต่บันทึกได้บางคอลัมน์ — ตรวจชื่อคอลัมน์ใน HD_PhishingReports":s?"ส่งรายงานพร้อมอีเมลต้นฉบับแล้ว":"ส่งรายงานแล้ว (แนบ .eml ไม่ได้)",r?"info":"success")}catch(e){f.toast(`ส่งรายงานไม่สำเร็จ: ${e instanceof Error?e.message:String(e)}`,"error")}finally{m.reporting=!1,f.rerender()}}}const lt=()=>m.reported?"✓ รายงานแล้ว":m.reporting?"กำลังส่ง…":"🚩 รายงานอีเมลนี้ให้ IT";function dt(e){var t;try{const n=(t=Office.context)==null?void 0:t.ui;if(typeof(n==null?void 0:n.openBrowserWindow)=="function"){n.openBrowserWindow(e);return}}catch{}window.open(e,"_blank","noopener,noreferrer")||f.toast("เปิดหน้าต่างไม่ได้ (ถูกบล็อก)","info")}async function mt(e){var t;try{if((t=navigator.clipboard)!=null&&t.writeText)return await navigator.clipboard.writeText(e),!0}catch{}try{const n=document.createElement("textarea");n.value=e,n.setAttribute("readonly",""),n.style.position="fixed",n.style.opacity="0",document.body.appendChild(n),n.focus(),n.select();const a=document.execCommand("copy");return n.remove(),a}catch{return!1}}function ut(e){if(m.loading&&!m.analysis)return`<div class="py-10 text-center text-slate-500 text-sm">
      <div class="w-7 h-7 border-[3px] border-slate-200 border-t-blue-600 rounded-full animate-spin mx-auto mb-2"></div>
      กำลังตรวจอีเมล…</div>`;const t=m.analysis,n=m.mail;if(!t||!n)return'<p class="text-sm text-slate-400 text-center py-8">เปิดอีเมลเพื่อเริ่มตรวจ</p>';const a=Se[t.level],o=[...t.links.filter(s=>s.flags.length),...t.links.filter(s=>!s.flags.length&&!s.trusted),...t.links.filter(s=>s.trusted)],r=f.canWhitelist();return`
    <div class="rounded-xl border-2 ${a.cls} p-3">
      <div class="flex items-center gap-2">
        <span class="text-2xl leading-none">${a.icon}</span>
        <div class="min-w-0 flex-1">
          <div class="font-bold text-sm">${F(a.label)}</div>
          <div class="text-xs opacity-80">คะแนน ${t.score} · พบสัญญาณ ${t.findings.filter(s=>s.severity!=="info").length} ข้อ</div>
        </div>
        <button id="phish-recheck" class="text-[10px] px-2 py-1 rounded-md bg-white/70 hover:bg-white text-slate-700 font-medium">ตรวจใหม่</button>
      </div>
    </div>

    ${e?"":`<div class="bg-blue-50 border border-blue-200 rounded-xl p-2.5 text-[11px] text-blue-800">
      เข้าสู่ระบบเพื่อตรวจการปลอมเป็นพนักงาน และรายงานเข้า Helpdesk ได้
    </div>`}

    ${t.findings.length===0?'<p class="text-xs text-slate-400 text-center py-3">ไม่พบสัญญาณผิดปกติจากการตรวจอัตโนมัติ</p>':t.findings.map(s=>{const c=Ce[s.severity];return`<div class="bg-white rounded-xl border border-slate-200 p-2.5">
          <div class="flex items-start gap-2">
            <span class="text-[9px] font-bold px-1.5 py-0.5 rounded-full ${c.cls} flex-shrink-0 mt-0.5">${c.label}</span>
            <div class="min-w-0 flex-1">
              <div class="text-xs font-semibold text-slate-800">${F(s.title)}</div>
              <div class="text-[11px] text-slate-500 whitespace-pre-line break-all">${F(s.detail)}</div>
              <div class="text-[9px] text-slate-400 mt-0.5">${F(s.category)}</div>
            </div>
          </div>
        </div>`}).join("")}

    ${o.length?`
    <div class="bg-white rounded-xl border border-slate-200 p-3">
      <div class="text-xs font-semibold text-slate-700 mb-2">ลิงก์ในอีเมล (${o.length})</div>
      <div class="space-y-2">
        ${o.map((s,c)=>{const p=m.savingDomain===(s.host?s.host.toLowerCase():"");return`
          <div class="rounded-lg border ${s.trusted?"border-emerald-200 bg-emerald-50/50":s.flags.length?"border-red-200 bg-red-50/50":"border-slate-100"} p-2">
            <div class="text-[11px] font-medium ${s.trusted?"text-emerald-800":s.flags.length?"text-red-700":"text-slate-700"} break-all">
              ${s.trusted?"✔ ":""}${F(s.host||s.href)}
            </div>
            ${s.trusted?`<div class="text-[10px] text-emerald-700">ทีมตรวจแล้วว่าปลอดภัย${(s.suppressed??[]).length?` · ระงับการเตือน ${s.suppressed.length} ข้อ`:""}</div>`:""}
            ${s.text&&s.text!==s.href?`<div class="text-[10px] text-slate-500 break-all">แสดงว่า: "${F(s.text.slice(0,70))}"</div>`:""}
            <div class="text-[10px] text-slate-400 break-all mt-0.5">${F(s.href.slice(0,150))}</div>
            ${s.flags.map(l=>`<div class="text-[10px] text-red-600 mt-0.5">! ${F(l)}</div>`).join("")}
            <div class="flex flex-wrap gap-1 mt-1.5">
              <button data-kasm="${c}" class="text-[10px] font-semibold px-2 py-1 rounded-md bg-slate-800 hover:bg-slate-900 text-white">
                เปิดใน Kasm
              </button>
              ${r&&s.host?s.trusted?`<button data-untrust="${F(s.host)}" ${p?"disabled":""}
                     class="text-[10px] font-semibold px-2 py-1 rounded-md border border-slate-300 text-slate-600 hover:bg-slate-100 disabled:opacity-50">
                     ${p?"...":"ถอนออกจากรายการปลอดภัย"}</button>`:`<button data-trust="${F(s.host)}" ${p?"disabled":""}
                     class="text-[10px] font-semibold px-2 py-1 rounded-md bg-emerald-600 hover:bg-emerald-700 text-white disabled:opacity-50">
                     ${p?"...":"✔ ตรวจแล้ว ปลอดภัย"}</button>`:""}
            </div>
          </div>`}).join("")}
      </div>
      <p class="text-[10px] text-slate-400 mt-2">อย่าคลิกลิงก์จากอีเมลที่ไม่มั่นใจโดยตรง</p>
    </div>`:""}

    <div class="flex gap-2">
      <button id="phish-copy" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">คัดลอกผลตรวจ</button>
      <button id="phish-headers" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">${m.showHeaders?"ซ่อน header":"ดู header"}</button>
    </div>
    ${m.showHeaders?`<pre class="bg-slate-900 text-slate-100 text-[9px] p-2 rounded-lg overflow-x-auto whitespace-pre-wrap break-all max-h-56 overflow-y-auto">${F(Object.entries(n.headers).map(([s,c])=>`${s}: ${c}`).join(`
`)||"อ่าน header ไม่ได้")}</pre>`:""}
    <p class="text-[10px] text-slate-400 text-center">วิเคราะห์ในเครื่องทั้งหมด — ไม่ส่งเนื้อหาอีเมลออกนอกองค์กร</p>
  `}function pt(){var n,a,o,r;(n=document.getElementById("phish-recheck"))==null||n.addEventListener("click",()=>{ee(!0)}),(a=document.getElementById("phish-headers"))==null||a.addEventListener("click",()=>{m.showHeaders=!m.showHeaders,f.rerender()}),(o=document.getElementById("phish-copy"))==null||o.addEventListener("click",async()=>{const s=await mt(De());f.toast(s?"คัดลอกผลตรวจแล้ว":"คัดลอกไม่ได้",s?"success":"error")}),document.querySelectorAll("[data-trust]").forEach(s=>s.addEventListener("click",()=>ot(s.dataset.trust??""))),document.querySelectorAll("[data-untrust]").forEach(s=>s.addEventListener("click",()=>it(s.dataset.untrust??"")));const e=((r=m.analysis)==null?void 0:r.links)??[],t=[...e.filter(s=>s.flags.length),...e.filter(s=>!s.flags.length&&!s.trusted),...e.filter(s=>s.trusted)];document.querySelectorAll("[data-kasm]").forEach(s=>{s.addEventListener("click",()=>{const c=t[Number(s.dataset.kasm)];if(!c)return;if(!m.kasmTemplate){f.toast("ยังไม่ได้ตั้งค่า Kasm ใน HD_Options (Category=KasmConfig)","info");return}const p=m.kasmTemplate;dt(p.includes("{url}")?p.replace("{url}",encodeURIComponent(c.href)):p+encodeURIComponent(c.href))})})}const ft="0bab07cf-65e6-487c-89af-c917fc1a5a13",ht="d569b991-89fc-4a62-9df5-eb361abcef40",M="https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd",X="https://rpaexpert.sharepoint.com/.default",ae=["https://graph.microsoft.com/Calendars.ReadWrite","https://graph.microsoft.com/Mail.Send","https://graph.microsoft.com/Mail.Read"],L=new Fe({auth:{clientId:ft,authority:`https://login.microsoftonline.com/${ht}`,redirectUri:window.location.origin.includes("localhost")?"http://localhost:3000/":"https://darmmunginsa.github.io/itservices-addin/",navigateToLoginRequestUrl:!1},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:!1}}),gt=window.location.origin.includes("localhost")?`${window.location.origin}/auth.html`:"https://darmmunginsa.github.io/itservices-addin/auth.html";function he(){var t,n;const e=(n=(t=Office.context)==null?void 0:t.diagnostics)==null?void 0:n.platform;return e===Office.PlatformType.iOS||e===Office.PlatformType.Android}function ge(){return new Promise((e,t)=>{Office.context.ui.displayDialogAsync(gt,{height:60,width:30,promptBeforeOpen:!1},n=>{if(n.status!==Office.AsyncResultStatus.Succeeded){t(new Error("เปิดหน้าเข้าสู่ระบบไม่ได้"));return}const a=n.value;a.addEventHandler(Office.EventType.DialogMessageReceived,o=>{a.close();const r=o.message;if(!r){t(new Error("auth message error"));return}try{const s=JSON.parse(r);s.ok?e():t(new Error(s.error||"auth failed"))}catch{t(new Error("auth message error"))}}),a.addEventHandler(Office.EventType.DialogEventReceived,()=>t(new Error("ปิดหน้าเข้าสู่ระบบก่อนเสร็จ")))})})}const i={account:null,tab:"ticket",emailSubject:"",emailBodyPreview:"",emailBodyReply:"",emailSenderName:"",emailSenderEmail:"",loading:!1,projects:[],agents:[],myRole:"",emailAttachments:[],signatureContact:null,droppedFiles:[],tickets:[],contactEmails:[],emailCc:[]};async function H(){const e=L.getAllAccounts();if(e.length===0)throw new Error("Not signed in");const t={scopes:[X],account:e[0]};try{return(await L.acquireTokenSilent(t)).accessToken}catch{if(he()){await ge();const n=L.getAllAccounts()[0];if(!n)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await L.acquireTokenSilent({scopes:[X],account:n})).accessToken}return(await L.acquireTokenPopup(t)).accessToken}}async function K(e=!1){const t=L.getAllAccounts();if(t.length===0)throw new Error("Not signed in");const n={scopes:ae,account:t[0],forceRefresh:e};try{return(await L.acquireTokenSilent(n)).accessToken}catch{if(he()){await ge();const o=L.getAllAccounts()[0];if(!o)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await L.acquireTokenSilent({scopes:ae,account:o})).accessToken}return(await L.acquireTokenPopup({scopes:ae,account:t[0]})).accessToken}}async function bt(e){const t=await K(),n={subject:e.subject,start:{dateTime:e.start,timeZone:"Asia/Bangkok"},end:{dateTime:e.end,timeZone:"Asia/Bangkok"},body:e.body?{contentType:"HTML",content:e.body.replace(/\n/g,"<br>")}:void 0,attendees:e.attendees.filter(Boolean).map(o=>({emailAddress:{address:o},type:"required"})),isOnlineMeeting:e.isOnlineMeeting,onlineMeetingProvider:e.isOnlineMeeting?"teamsForBusiness":void 0},a=await fetch("https://graph.microsoft.com/v1.0/me/events",{method:"POST",headers:{Authorization:`Bearer ${t}`,"Content-Type":"application/json"},body:JSON.stringify(n)});if(!a.ok)throw new Error(`Calendar error ${a.status}: ${await a.text()}`)}async function Oe(){try{const e=await H(),t=`${M}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$orderby=Title asc&$top=500`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const a=await n.json();i.projects=a.value.map(o=>({id:o.Id,Title:o.Title}))}}catch{}}async function _e(){var e,t;try{const n=await H(),a=`${M}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText,Role&$orderby=Title asc`,o=await fetch(a,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});if(o.ok){const r=await o.json();i.agents=r.value.map(c=>({email:c.EmailText,name:c.Title}));const s=(((e=i.account)==null?void 0:e.username)??"").toLowerCase();i.myRole=((t=r.value.find(c=>(c.EmailText??"").toLowerCase()===s))==null?void 0:t.Role)??""}}catch{}}async function Le(){try{const e=await H(),t=`${M}/_api/web/lists/getbytitle('HD_Tickets')/items?$select=Id,Title,TicketNumber,Status&$filter=Status ne 'Closed'&$orderby=Modified desc&$top=200`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const a=await n.json();i.tickets=a.value.map(o=>({id:o.Id,Title:o.Title,TicketNumber:o.TicketNumber,Status:o.Status}))}}catch{}}async function Re(){try{const e=await H(),t=`${M}/_api/web/lists/getbytitle('HD_Contracts')/items?$select=CustomerEmail&$top=2000`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const a=await n.json();i.contactEmails=a.value.map(o=>(o.CustomerEmail||"").trim().toLowerCase()).filter(Boolean)}}catch{}}async function Te(){const e=document.getElementById("btn-login-main"),t=document.getElementById("btn-login");e&&(e.disabled=!0,e.textContent="กำลังเข้าสู่ระบบ…"),t&&(t.disabled=!0);try{if(he()){if(await ge(),i.account=L.getAllAccounts()[0]??null,!i.account)throw new Error("เข้าสู่ระบบไม่สำเร็จ")}else{const n=await L.loginPopup({scopes:[X]});i.account=n.account}await Promise.all([Oe(),_e(),Le(),Re()]),z()}catch{e&&(e.disabled=!1,e.textContent="เข้าสู่ระบบ"),t&&(t.disabled=!1)}}async function yt(){i.account&&await L.logoutPopup({account:i.account}),i.account=null,z()}async function q(e,t){const n=await H(),a=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items`,o=await fetch(a,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(t)});if(!o.ok){const s=await o.text();throw new Error(`SharePoint error ${o.status}: ${s}`)}return(await o.json()).Id}let J=null;const se="support@itservices.co.th",Me="engineer@itservices.co.th";async function He(){if(J)return J;try{const e=await H(),t=`${M}/_api/web/lists/getbytitle('HD_EmailTemplates')/items?$select=EventKey,Subject,Body,IsEnabled&$top=50`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});return n.ok?(J=(await n.json()).value,J):[]}catch{return[]}}async function xt(){var e,t;try{const n=await H(),a=`${M}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title,Category&$filter=Category eq 'EmailConfig'&$top=1`,o=await fetch(a,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});return o.ok&&((t=(e=(await o.json()).value[0])==null?void 0:e.Title)==null?void 0:t.trim())||se}catch{return se}}function re(e,t){return e.replace(/\{\{(\w+)\}\}/g,(n,a)=>t[a]??`{{${a}}}`)}async function ke(e,t,n,a=[]){try{const r=(await He()).find(d=>d.EventKey===e&&d.IsEnabled);if(!r)return;const s=re(r.Subject||"",t),c=re(r.Body||"",t);if(!s||!c)return;const p=d=>d.trim().toLowerCase(),v=[...new Map(n.filter(Boolean).map(d=>[p(d),d])).values()];if(v.length===0)return;const l=new Set(v.map(p)),b=e==="ticket_created"?[...a,Me]:a,j=[...new Map(b.filter(Boolean).map(d=>[p(d),d])).values()].filter(d=>!l.has(p(d))),_=await xt(),N=await K(),T={subject:s,body:{contentType:"HTML",content:c},toRecipients:v.map(d=>({emailAddress:{address:d}}))};j.length&&(T.ccRecipients=j.map(d=>({emailAddress:{address:d}}))),_&&(T.from={emailAddress:{address:_}}),await fetch("https://graph.microsoft.com/v1.0/me/sendMail",{method:"POST",headers:{Authorization:`Bearer ${N}`,"Content-Type":"application/json"},body:JSON.stringify({message:T,saveToSentItems:!0})})}catch{}}async function vt(e,t=[]){try{const n=Office.context.mailbox.item;if(!(n!=null&&n.itemId))return!1;const a=Office.context.mailbox.convertToRestId(n.itemId,Office.MailboxEnums.RestVersion.v2_0),r={Authorization:`Bearer ${await K()}`,"Content-Type":"application/json"},s=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${a}/createReplyAll`,{method:"POST",headers:r});if(!s.ok)return!1;const c=await s.json(),p=T=>T.trim().toLowerCase(),v=c.ccRecipients??[],l=new Set(v.map(T=>p(T.emailAddress.address))),b=[...new Set(t.filter(Boolean).map(T=>T.trim()))].filter(T=>!l.has(p(T))).map(T=>({emailAddress:{address:T}})),j={body:{contentType:"HTML",content:e}};return b.length&&(j.ccRecipients=[...v,...b]),(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${c.id}`,{method:"PATCH",headers:r,body:JSON.stringify(j)})).ok?(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${c.id}/send`,{method:"POST",headers:r})).ok:!1}catch{return!1}}async function $t(e,t){const a=(await He()).find(r=>r.EventKey===e&&r.IsEnabled);return a&&re(a.Body||"",t)||null}async function Y(e){var r;const t=s=>s.trim().toLowerCase(),n=t(((r=i.account)==null?void 0:r.username)??""),a=new Set,o=e.recipients.filter(Boolean).filter(s=>{const c=t(s);return!c||c===n||a.has(c)?!1:(a.add(c),!0)});if(o.length!==0)try{const s=await H(),c=`${M}/_api/web/lists/getbytitle('HD_Notifications')/items`;await Promise.all(o.map(p=>fetch(c,{method:"POST",headers:{Authorization:`Bearer ${s}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:e.title.slice(0,255),RecipientEmail:p,EventType:e.eventType,Message:e.message,LinkPath:e.linkPath,IsRead:!1})})))}catch{}}async function ce(e,t){const n=document.querySelectorAll(".email-att-cb:checked");if(n.length===0)return;const a=await H(),o=new Set,r=s=>{if(!o.has(s.toLowerCase()))return o.add(s.toLowerCase()),s;const c=s.lastIndexOf("."),p=c>0?s.slice(0,c):s,v=c>0?s.slice(c):"";for(let l=2;;l++){const b=p+"-"+l+v;if(!o.has(b.toLowerCase()))return o.add(b.toLowerCase()),b}};for(const s of Array.from(n)){const c=s.dataset.attId,p=r(s.dataset.attName),v=await new Promise((d,h)=>{Office.context.mailbox.item.getAttachmentContentAsync(c,{},$=>{$.status===Office.AsyncResultStatus.Succeeded?d($):h(new Error($.error.message))})}),{content:l,format:b}=v.value;let j;if(b===Office.MailboxEnums.AttachmentContentFormat.Base64){const d=atob(l);j=new Uint8Array(d.length);for(let h=0;h<d.length;h++)j[h]=d.charCodeAt(h)}else if(b===Office.MailboxEnums.AttachmentContentFormat.Eml||b===Office.MailboxEnums.AttachmentContentFormat.ICalendar)j=new TextEncoder().encode(l);else continue;const _=encodeURIComponent(p),N=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${_}')`;if(!(await fetch(N,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:j.buffer})).ok)throw new Error(`Upload ${p} failed`)}}async function wt(e){const t=`https://graph.microsoft.com/v1.0/me/messages/${e}/$value`;let n=await K(),a=await fetch(t,{headers:{Authorization:`Bearer ${n}`}});if((a.status===401||a.status===403)&&(n=await K(!0),a=await fetch(t,{headers:{Authorization:`Bearer ${n}`}})),!a.ok)throw new Error(`Graph ${a.status}`);return a.arrayBuffer()}async function Tt(e){const t=await new Promise((a,o)=>{Office.context.mailbox.getCallbackTokenAsync({isRest:!0},r=>{r.status===Office.AsyncResultStatus.Succeeded?a(r.value):o(new Error("callback token failed"))})}),n=await fetch(`${Office.context.mailbox.restUrl}/v2.0/me/messages/${e}/$value`,{headers:{Authorization:`Bearer ${t}`}});if(!n.ok)throw new Error(`REST ${n.status}`);return n.arrayBuffer()}async function le(e,t){const n=document.getElementById("f-attach-eml");if(!(n!=null&&n.checked))return;const a=Office.context.mailbox.item;if(!a)return;const o=Office.context.mailbox.convertToRestId(a.itemId,Office.MailboxEnums.RestVersion.v2_0);let r,s="",c="";try{r=await wt(o)}catch(j){s=j instanceof Error?j.message:String(j);try{r=await Tt(o)}catch(_){c=_ instanceof Error?_.message:String(_),console.error("[eml] graph:",s,"| callback:",c),B(`ดึง .eml ไม่ได้ (Graph: ${s} / REST: ${c}) — ไฟล์อื่นบันทึกแล้ว`,"error");return}}const p=(a.subject||"email").replace(/[\\/:*?"<>|#%&{}~]/g,"_").slice(0,100).trim()||"email",v=await H(),l=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${encodeURIComponent(p+".eml")}')`;(await fetch(l,{method:"POST",headers:{Authorization:`Bearer ${v}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok||B("บันทึกไฟล์ .eml ไม่สำเร็จ","error")}async function de(e,t,n){const a=await H();for(const o of n){const r=await o.arrayBuffer(),s=encodeURIComponent(o.name),c=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${s}')`;if(!(await fetch(c,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok)throw new Error(`Upload ${o.name} failed`)}}function B(e,t="success"){const n=document.getElementById("toast-container");if(!n)return;const a=t==="success"?"bg-green-500":t==="error"?"bg-red-500":"bg-slate-700",o=t==="success"?"✅":t==="error"?"❌":"ℹ️",r=document.createElement("div");r.className=`toast pointer-events-auto ${a} text-white text-sm font-medium px-4 py-3 rounded-lg shadow-lg max-w-xs mx-2`,r.textContent=`${o} ${e}`,n.appendChild(r),setTimeout(()=>r.remove(),4e3)}function kt(e){const t=e.split(`
`).map(l=>l.trim()).filter(Boolean),n=/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/,a=/(\+?[\d\s()\-]{7,}(?:ext\.?\s*\d+)?)/i,o=/\b(co\.?,?\s*ltd\.?|co\.?,?\s*limited|corporation|corp\.?|บริษัท|จำกัด|holding|group|inc\.?|llc)\b/i;let r="",s="",c="";const p=[];for(const l of t)if(!/^[-_=*]{2,}$/.test(l)&&!/^(best regards|regards|sincerely|ขอแสดงความนับถือ|ด้วยความนับถือ|from|sent|thanks|thank you)/i.test(l)){if(!r){const b=l.match(n);if(b){r=b[0];continue}}if(!s){const b=l.match(a);if(b&&b[0].replace(/\D/g,"").length>=7){s=b[0].trim();continue}}if(!c&&o.test(l)){c=l;continue}l.length>=2&&l.length<=50&&!/\d{4,}/.test(l)&&p.push(l)}const v=p.find(l=>!n.test(l)&&!o.test(l))??"";return!r&&!v?null:{name:v,company:c,email:r,phone:s}}async function Et(){const e=i.signatureContact;if(!e)return;const t=(i.emailSenderEmail||"").toLowerCase();if(t&&i.contactEmails.includes(t)){B("ลูกค้านี้มีในระบบแล้ว","success"),i.signatureContact=null,z();return}const n=document.getElementById("btn-import-customer");n&&(n.disabled=!0,n.textContent="กำลังบันทึก…");try{await q("HD_Contracts",{Title:i.emailSenderName||e.name,CustomerEmail:i.emailSenderEmail,Phone:e.phone||void 0,Company:e.company||void 0,Status:"Active"}),t&&i.contactEmails.push(t),B("เพิ่มลูกค้าสำเร็จ!"),i.signatureContact=null,z()}catch(a){const o=a instanceof Error?a.message:String(a);B(`เกิดข้อผิดพลาด: ${o}`,"error"),n&&(n.disabled=!1,n.textContent="เพิ่มเป็นลูกค้า")}}function me(){return new Date().toISOString().split("T")[0]}function It(){const e=new Date;return`HD-${`${e.getFullYear()}${String(e.getMonth()+1).padStart(2,"0")}${String(e.getDate()).padStart(2,"0")}`}-${Math.floor(Math.random()*900+100)}`}function Ee(){var e;return i.droppedFiles.length>0||document.querySelectorAll(".email-att-cb:checked").length>0||(((e=document.getElementById("f-attach-eml"))==null?void 0:e.checked)??!1)}async function Z(e,t){i.droppedFiles.length>0&&await de(e,t,i.droppedFiles),await ce(e,t),await le(e,t)}let oe=!1;async function At(){var t,n,a,o,r,s,c,p,v,l,b,j,_,N,T;if(!i.account){B("กรุณาเข้าสู่ระบบก่อน","error");return}if(oe)return;oe=!0;const e=document.getElementById("submit-btn");e&&(e.disabled=!0,e.textContent="กำลังบันทึก…");try{if(i.tab==="phish")await ct();else if(i.tab==="ticket"){const d=document.getElementById("f-title").value.trim(),h=document.getElementById("f-description").value.trim(),$=document.getElementById("f-priority").value,k=document.getElementById("f-customer-email").value.trim(),x=((t=document.getElementById("f-cc-enable"))==null?void 0:t.checked)??!0?(((n=document.getElementById("f-cc"))==null?void 0:n.value)||"").split(/[,;\s]+/).map(A=>A.trim()).filter(Boolean):[],S=document.getElementById("f-assigned-email").value,C=i.agents.find(A=>A.email===S),D=It(),P=await q("HD_Tickets",{Title:d,TicketNumber:D,Description:h,Priority:$,CustomerEmail:k,CustomerName:i.emailSenderName||k,Status:"Open",AssignedEmail:S||void 0,AssignedToName:(C==null?void 0:C.name)??((a=i.account)==null?void 0:a.name)??"",ProjectID:parseInt(((o=document.getElementById("f-project"))==null?void 0:o.value)||"0")||null});if(Ee()){const A=await q("HD_TicketComments",{Title:"📎 ไฟล์แนบจากอีเมล",TicketID:P,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await Z("HD_TicketComments",A)}i.droppedFiles=[];const R={ticket_number:D,ticket_title:d,priority:$,category:"-",description:(h||"-").replace(/\n/g,"<br>"),customer_name:i.emailSenderName||k,assigned_name:(C==null?void 0:C.name)??((r=i.account)==null?void 0:r.name)??"-",link:"https://itservices.co.th/helpdesk/"},u=[S,i.account.username,...x,Me].filter(Boolean);let g=!1;const O=await $t("ticket_created",R);if(O){const A=`<div style="border-left:4px solid #2563eb;background:#eff6ff;padding:10px 14px;margin:0 0 14px;font-family:Segoe UI,sans-serif">
             <div style="font-size:15px;font-weight:700;color:#1e40af">Ticket No. ${D}</div>
             <div style="font-size:12px;color:#475569;margin-top:2px">กรุณาตอบกลับในอีเมลฉบับนี้เพื่อให้ข้อมูลอยู่ใน Ticket เดียวกัน</div>
           </div>`;g=await vt(A+O,u)}g||await ke("ticket_created",R,[k],u),B(g?"สร้าง Ticket และตอบกลับในเธรดเดิมแล้ว!":"สร้าง Ticket สำเร็จ!")}else if(i.tab==="task"){const d=document.getElementById("f-title").value.trim(),h=parseInt(((s=document.getElementById("f-project"))==null?void 0:s.value)||"0"),$=document.getElementById("f-due-date").value,k=document.getElementById("f-note").value.trim(),I=document.getElementById("f-assigned-email").value,x=i.agents.find(D=>D.email===I);if(!h){B("กรุณาเลือก Project","error");return}const S=await q("PM_Tasks",{Title:d,DueDate:$||null,TaskNote:k,AssignedTo:(x==null?void 0:x.name)??i.account.name??i.account.username,AssignedEmail:I,IsCompleted:!1,IsAcknowledged:!1,ProjectID:h});if(i.droppedFiles.length>0&&await de("PM_Tasks",S,i.droppedFiles),await ce("PM_Tasks",S),await le("PM_Tasks",S),i.droppedFiles=[],await Y({recipients:[I],title:`📋 ได้รับมอบหมาย Task: ${d}`,message:k||($?`กำหนดส่ง ${$}`:"มี Task ใหม่"),linkPath:h?`/projects/${h}`:"/my-work",eventType:"task_assigned"}),((c=document.getElementById("f-teams"))==null?void 0:c.checked)&&$){const D=Array.from(document.querySelectorAll(".att-internal:checked")).map(g=>g.value),P=(((p=document.getElementById("f-ext-att"))==null?void 0:p.value)||"").split(/[,;\s]+/).map(g=>g.trim()).filter(Boolean),R=`${$}T09:00:00`,u=`${$}T10:00:00`;try{await bt({subject:d,start:R,end:u,body:k,attendees:[...D,...P],isOnlineMeeting:!0}),B("สร้าง Task + นัดประชุม Teams สำเร็จ!")}catch(g){B("สร้าง Task แล้ว แต่สร้างนัดประชุมไม่สำเร็จ: "+(g instanceof Error?g.message:""),"error")}}else B("สร้าง Task สำเร็จ!")}else if(i.tab==="incident"){const d=document.getElementById("f-title").value.trim(),h=parseInt(((v=document.getElementById("f-project"))==null?void 0:v.value)||"0"),$=document.getElementById("f-description").value.trim(),k=document.getElementById("f-severity").value,I=document.getElementById("f-assigned-email").value,x=i.agents.find(u=>u.email===I),S=document.getElementById("f-status").value,C=document.getElementById("f-incident-date").value,D=document.getElementById("f-resolution").value.trim();if(!h){B("กรุณาเลือก Project","error");return}const P=parseInt(((l=document.getElementById("f-sla"))==null?void 0:l.value)||"0")||null,R=await q("PM_Incidents",{Title:d,Description:$||void 0,Severity:k,Status:S,AssignedTo:(x==null?void 0:x.name)??i.account.name??i.account.username,AssignedEmail:I,ProjectID:h,IncidentDate:C||me(),Resolution:D||void 0,SLAHours:P,SLADue:Be(P),...S==="Resolved"?{ResolvedDate:new Date().toISOString()}:{}});i.droppedFiles.length>0&&await de("PM_Incidents",R,i.droppedFiles),await ce("PM_Incidents",R),await le("PM_Incidents",R),i.droppedFiles=[],await Y({recipients:[I],title:`🚨 ได้รับมอบหมาย Incident: ${d}`,message:`ความรุนแรง ${k}${$?" — "+$.slice(0,120):""}`,linkPath:h?`/projects/${h}`:"/my-work",eventType:"incident_created"});{const u=i.projects.find(O=>O.id===h),g=Ze({title:d,severity:k,status:S,description:$,incidentDate:C||me(),slaHours:P,projectName:u==null?void 0:u.Title,projectId:h,assignedName:x==null?void 0:x.name,assignedEmail:I,requesterEmail:i.account.username,actorEmail:i.account.username,baseUrl:"https://itservices.co.th/helpdesk/"});g.to.length>0&&await ke("incident_created",g.vars,g.to,g.cc)}B("สร้าง Incident สำเร็จ!")}else if(i.tab==="comment"){const d=parseInt(((b=document.getElementById("f-ticket"))==null?void 0:b.value)||"0"),h=document.getElementById("f-comment").value.trim(),$=document.getElementById("f-comment-type").value;if(!d){B("กรุณาเลือก Ticket","error");return}if(!h){B("กรุณาพิมพ์ Comment","error");return}const k=await q("HD_TicketComments",{Title:h.slice(0,100),TicketID:d,CommentText:h,CommentType:$,CommentDate:new Date().toISOString()});await Z("HD_TicketComments",k),i.droppedFiles=[];try{const I=await H(),x=`${M}/_api/web/lists/getbytitle('HD_Tickets')/items(${d})?$select=TicketNumber,Title,AssignedEmail,Author/EMail&$expand=Author`,S=await fetch(x,{headers:{Authorization:`Bearer ${I}`,Accept:"application/json;odata=nometadata"}});if(S.ok){const C=await S.json(),D=i.account.username.toLowerCase(),P=[...new Set([C.AssignedEmail,(j=C.Author)==null?void 0:j.EMail].filter(Boolean))].filter(R=>R.toLowerCase()!==D);P.length&&await Y({recipients:P,title:`💬 ${((_=i.account)==null?void 0:_.name)??"มีคน"} คอมเมนต์ใน ${C.TicketNumber||"#"+d}`,message:h.slice(0,200),linkPath:`/tickets/${d}`,eventType:"comment_added"})}}catch{}B("เพิ่ม Comment สำเร็จ!")}else if(i.tab==="project"){const d=document.getElementById("f-title").value.trim(),h=document.getElementById("f-company").value.trim(),$=document.getElementById("f-group").value,k=document.getElementById("f-status").value,I=document.getElementById("f-start").value,x=document.getElementById("f-end").value,S=document.getElementById("f-description").value.trim();if(!d){B("กรุณาใส่ชื่อโครงการ","error");return}const C=await q("PM_Projects",{Title:d,Company:h||void 0,ProjectGroup:$,Progress:0,StartDate:I||void 0,EndDate:x||null,Status:k,CreatedByEmail:i.account.username,Comment:S||void 0});if(Ee()){const D=await q("PM_Comments",{Title:"📎 ไฟล์แนบจากอีเมล",ProjectID:C,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await Z("PM_Comments",D)}i.droppedFiles=[],B("สร้างโครงการสำเร็จ!")}else if(i.tab==="projcomment"){const d=parseInt(((N=document.getElementById("f-project"))==null?void 0:N.value)||"0"),h=document.getElementById("f-comment").value.trim(),$=document.getElementById("f-comment-type").value;if(!d){B("กรุณาเลือกโครงการ","error");return}if(!h){B("กรุณาพิมพ์ Comment","error");return}const k=await q("PM_Comments",{Title:h.slice(0,100),ProjectID:d,CommentText:h,CommentType:$,CommentDate:new Date().toISOString(),ParentID:0});await Z("PM_Comments",k),i.droppedFiles=[];try{const I=await H(),x=`${M}/_api/web/lists/getbytitle('PM_Projects')/items(${d})?$select=Title,CreatedByEmail`,S=await fetch(x,{headers:{Authorization:`Bearer ${I}`,Accept:"application/json;odata=nometadata"}});if(S.ok){const C=await S.json(),D=i.account.username.toLowerCase();C.CreatedByEmail&&C.CreatedByEmail.toLowerCase()!==D&&await Y({recipients:[C.CreatedByEmail],title:`💬 ${((T=i.account)==null?void 0:T.name)??"มีคน"} คอมเมนต์ในโครงการ ${C.Title??""}`,message:h.slice(0,200),linkPath:`/projects/${d}?tab=comments`,eventType:"comment_added"})}}catch{}B("เพิ่ม Comment สำเร็จ!")}}catch(d){const h=d instanceof Error?d.message:String(d);B(`เกิดข้อผิดพลาด: ${h}`,"error")}finally{oe=!1,e&&(e.disabled=!1,e.textContent="บันทึก")}}const St={ticket:{label:"Ticket",icon:"🎫"},comment:{label:"Comment",icon:"💬"},project:{label:"Project",icon:"📁"},task:{label:"Task",icon:"✅"},incident:{label:"Incident",icon:"🚨"},projcomment:{label:"Comment",icon:"💬"},phish:{label:"PhishGuard",icon:"🛡️"}},Ct=[{title:"🎫 Helpdesk",tabs:["ticket","comment"]},{title:"📁 Project",tabs:["project","task","incident","projcomment"]},{title:"🛡️ Security",tabs:["phish"]}],Ne=["f-title","f-description","f-priority","f-customer-email","f-cc","f-assigned-email","f-project","f-due-date","f-note","f-severity","f-status","f-sla","f-incident-date","f-resolution","f-ticket","f-comment","f-comment-type","f-company","f-group","f-start","f-end","f-ext-att","f-attach-eml"];let V={};function Bt(){for(const t of Ne){const n=document.getElementById(t);n&&(V[t]=n.value)}const e=document.getElementById("f-teams");e&&(V["f-teams"]=e.checked)}function jt(){for(const t of Ne){const n=document.getElementById(t);n&&V[t]!==void 0&&V[t]!==""&&(n.value=V[t])}const e=document.getElementById("f-teams");if(e&&V["f-teams"]!==void 0){e.checked=V["f-teams"];const t=document.getElementById("teams-fields");t&&(t.style.display=e.checked?"block":"none")}}function z(){var x,S,C,D,P,R;const e=document.getElementById("app");if(!e)return;Bt();const{account:t,tab:n,emailSubject:a,emailSenderName:o,emailSenderEmail:r,emailBodyPreview:s}=i,c=i.emailBodyReply||s,p=t!==null,v=`
    <div class="bg-gradient-to-r from-blue-700 to-blue-600 text-white px-4 py-2.5 flex items-center gap-2.5 shadow flex-shrink-0">
      <div class="w-7 h-7 bg-white/95 rounded-lg flex items-center justify-center flex-shrink-0">
        <span class="text-blue-700 font-extrabold text-xs">iT</span>
      </div>
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm leading-tight truncate">iT Services Helpdesk</div>
        ${p?`<div class="text-[10px] text-blue-100 truncate">${y((t==null?void 0:t.name)??(t==null?void 0:t.username)??"")}</div>`:""}
      </div>
      ${p?`<button id="btn-logout" title="ออกจากระบบ" class="p-1.5 rounded-lg hover:bg-white/15 transition flex-shrink-0">
             <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
           </button>`:`<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded-lg transition flex-shrink-0">
             เข้าสู่ระบบ
           </button>`}
    </div>
  `;if(!p){e.innerHTML=`
      ${v}
      <div class="flex flex-col items-center justify-center px-6 py-16 gap-4 text-center">
        <div class="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center text-3xl">🔐</div>
        <h2 class="text-base font-semibold text-slate-700">เข้าสู่ระบบด้วย Microsoft</h2>
        <p class="text-xs text-slate-500">เพื่อสร้าง Ticket / Task / Incident จาก Email นี้</p>
        <button id="btn-login-main"
          class="mt-2 w-full bg-blue-700 hover:bg-blue-600 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          เข้าสู่ระบบ
        </button>
      </div>
    `,(x=document.getElementById("btn-login"))==null||x.addEventListener("click",Te),(S=document.getElementById("btn-login-main"))==null||S.addEventListener("click",Te);return}const l=a?`<div class="mx-3 mt-3 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5 text-xs text-slate-700 space-y-1">
         <div class="font-semibold text-blue-800 truncate" title="${y(a)}">📧 ${y(a)}</div>
         ${o?`<div class="text-slate-500">จาก: <span class="font-medium text-slate-700">${y(o)}</span></div>`:""}
         ${r&&r!==o?`<div class="text-slate-400 truncate">${y(r)}</div>`:""}
       </div>`:`<div class="mx-3 mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-700">
         ⚠️ ไม่พบข้อมูล Email (โหมดทดสอบ)
       </div>`,b=i.signatureContact,j=!!r&&i.contactEmails.includes(r.toLowerCase()),_=b?`<div class="mx-3 mt-3 bg-orange-50 border border-orange-200 rounded-lg px-3 py-2.5 text-xs text-slate-700">
         <div class="flex items-center justify-between mb-2">
           <span class="font-semibold text-orange-700">👤 ข้อมูลผู้ส่ง (จาก Signature)</span>
         </div>
         <div class="space-y-0.5 mb-2.5">
           ${o?`<div><span class="text-slate-400">ชื่อ:</span> <span class="font-medium">${y(o)}</span></div>`:""}
           ${b.company?`<div><span class="text-slate-400">บริษัท:</span> ${y(b.company)}</div>`:""}
           ${r?`<div><span class="text-slate-400">Email:</span> ${y(r)}</div>`:""}
           ${b.phone?`<div><span class="text-slate-400">โทร:</span> ${y(b.phone)}</div>`:""}
         </div>
         ${j?'<div class="w-full bg-green-100 text-green-700 text-xs font-semibold py-1.5 rounded-md text-center">✓ ลูกค้านี้มีในระบบแล้ว</div>':`<button id="btn-import-customer"
                class="w-full bg-orange-500 hover:bg-orange-400 text-white text-xs font-semibold py-1.5 rounded-md transition">
                + เพิ่มเป็นลูกค้า
              </button>`}
       </div>`:"",N=`
    <div class="mx-3 mt-3 space-y-2">
      ${Ct.map(u=>`
        <div>
          <div class="text-[10px] font-semibold text-slate-400 uppercase tracking-wide mb-1 px-0.5">${u.title}</div>
          <div class="grid grid-cols-4 gap-1">
            ${u.tabs.map(g=>{const O=St[g];return`<button data-tab="${g}"
                class="tab-btn flex flex-col items-center gap-1 py-2 rounded-lg transition ${n===g?"bg-blue-700 text-white shadow":"text-slate-500 hover:bg-slate-100"}">
                <span class="text-base leading-none">${O.icon}</span>
                <span class="text-[9px] font-medium leading-none">${O.label}</span>
              </button>`}).join("")}
          </div>
        </div>
      `).join("")}
    </div>
  `;let T="";n==="phish"?T=ut(!!t):n==="ticket"?T=`
      ${E("Title / หัวข้อ",`<input id="f-title" type="text"
        class="${w}"
        value="${y(a)}" />`)}
      ${E("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${w} resize-none">${y(s)}</textarea>`)}
      ${E("Priority",`<select id="f-priority" class="${w}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
      ${E("Customer Email",`<input id="f-customer-email" type="email"
        class="${w}"
        value="${y(r)}" />`)}
      ${E("CC — ให้ผู้ที่อยู่ในเมลนี้รับรู้",`
        <label class="flex items-center gap-2 text-xs text-slate-600 mb-1.5 cursor-pointer">
          <input id="f-cc-enable" type="checkbox" ${i.emailCc.length?"checked":""} /> แนบผู้รับในเมลนี้เป็น CC อัตโนมัติ
        </label>
        <input id="f-cc" type="text" class="${w}" value="${y(i.emailCc.join(", "))}" placeholder="someone@company.com, boss@company.com" />`)}
      ${E("Assign ให้ Agent",ie(t.username))}
      ${E("โครงการ (ไม่บังคับ)",Q(!0))}
      ${G()}
    `:n==="task"?T=`
      ${E("ชื่อ Task *",`<input id="f-title" type="text" required
        class="${w}" value="${y(a)}" />`)}
      ${E("Project *",Q())}
      ${E("Assign ให้",ie(t.username))}
      ${E("Due Date",`<input id="f-due-date" type="date" class="${w}" />`)}
      ${E("Task Note",`<textarea id="f-note" rows="4"
        class="${w} resize-y">${y(s)}</textarea>`)}
      <label class="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer mb-1">
        <input id="f-teams" type="checkbox" class="rounded" onchange="document.getElementById('teams-fields').style.display=this.checked?'block':'none'" />
        💻 เพิ่มการประชุมออนไลน์ (Teams) — ใช้เวลา 09:00–10:00 ของวัน Due Date
      </label>
      <div id="teams-fields" style="display:none" class="space-y-2 mb-2">
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">ผู้เข้าร่วม Internal</label>
          <div class="max-h-28 overflow-y-auto border border-slate-200 rounded-md p-1.5 space-y-0.5">
            ${i.agents.map(u=>`<label class="flex items-center gap-2 text-xs text-slate-700 px-1 py-0.5 hover:bg-slate-50 rounded cursor-pointer">
              <input type="checkbox" class="att-internal" value="${y(u.email)}" /> ${y(u.name)}
            </label>`).join("")}
          </div>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">Email ภายนอก (คั่นด้วย ,)</label>
          <input id="f-ext-att" type="text" class="${w}" placeholder="someone@company.com, ..." />
        </div>
      </div>
      ${G()}
    `:n==="incident"?T=`
      ${E("ชื่อ Incident *",`<input id="f-title" type="text" required
        class="${w}" value="${y(a)}" />`)}
      ${E("Project *",Q())}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">ความรุนแรง</label>
          <select id="f-severity" class="${w}">
            <option value="Low">Low</option>
            <option value="Medium" selected>Medium</option>
            <option value="High">High</option>
            <option value="Critical">Critical</option>
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${w}">
            <option value="Open" selected>Open</option>
            <option value="In Progress">In Progress</option>
            <option value="Resolved">Resolved</option>
          </select>
        </div>
      </div>
      ${E("SLA — ต้องแก้ให้จบภายใน",`
        <select id="f-sla" class="${w}">
          <option value="">ไม่กำหนด SLA</option>
          ${Qe.map(u=>`<option value="${u.hours}" ${u.hours===ne.Medium?"selected":""}>${u.labelTh}</option>`).join("")}
        </select>
        <p id="f-sla-hint" class="text-[11px] text-slate-400 mt-1">นับจากตอนนี้ · ครบกำหนด ${ve(ne.Medium)}</p>`)}
      ${E("Assign ให้ Agent",ie(t.username))}
      ${E("วันที่เกิด Incident",`<input id="f-incident-date" type="date" class="${w}" value="${me()}" />`)}
      ${E("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${w} resize-y">${y(s)}</textarea>`)}
      ${E("วิธีแก้ไข (ถ้ามี)",`<textarea id="f-resolution" rows="2"
        class="${w} resize-y" placeholder="อธิบายวิธีแก้ไขปัญหา..."></textarea>`)}
      ${G()}
    `:n==="comment"?T=`
      ${E("เลือก Ticket *",`<select id="f-ticket" class="${w}">
        <option value="">-- เลือก Ticket ที่จะเพิ่ม Comment --</option>
        ${i.tickets.map(u=>`<option value="${u.id}">${y(u.TicketNumber||"#"+u.id)} · ${y(u.Title)}</option>`).join("")}
      </select>`)}
      ${E("ประเภท",`<select id="f-comment-type" class="${w}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${E("Comment *",`<textarea id="f-comment" rows="5"
        class="${w} resize-y" placeholder="พิมพ์ comment...">${y(c)}</textarea>`)}
      ${G()}
    `:n==="project"?T=`
      ${E("ชื่อโครงการ *",`<input id="f-title" type="text" required
        class="${w}" value="${y(a)}" />`)}
      ${E("บริษัท / ลูกค้า",`<input id="f-company" type="text" class="${w}" value="${y(((C=i.signatureContact)==null?void 0:C.company)??"")}" />`)}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">กลุ่มโครงการ</label>
          <select id="f-group" class="${w}">
            ${["Internal","External","R&D","Maintenance","อื่นๆ"].map(u=>`<option>${u}</option>`).join("")}
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${w}">
            ${["Planning","Active","On Hold","Completed","Cancelled"].map(u=>`<option>${u}</option>`).join("")}
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันที่เริ่ม</label>
          <input id="f-start" type="date" class="${w}" /></div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันสิ้นสุด</label>
          <input id="f-end" type="date" class="${w}" /></div>
      </div>
      ${E("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${w} resize-y">${y(s)}</textarea>`)}
      ${G()}
    `:n==="projcomment"&&(T=`
      ${E("เลือกโครงการ *",Q())}
      ${E("ประเภท",`<select id="f-comment-type" class="${w}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${E("Comment *",`<textarea id="f-comment" rows="5"
        class="${w} resize-y" placeholder="พิมพ์ comment...">${y(c)}</textarea>`)}
      ${G()}
    `);const d=n==="phish"?lt():n==="comment"||n==="projcomment"?"เพิ่ม Comment":n==="project"?"สร้างโครงการ":n==="incident"?"แจ้ง Incident":n==="task"?"สร้าง Task":"สร้าง Ticket";e.innerHTML=`
    <div class="flex flex-col h-screen bg-slate-50">
      ${v}
      <div class="flex-1 overflow-y-auto">
        ${l}
        ${_}
        ${N}
        <div class="mx-3 mt-3 space-y-3 pb-3">
          ${T}
        </div>
      </div>
      <div class="border-t border-slate-200 bg-white px-3 py-2.5 flex-shrink-0">
        <button id="submit-btn"
          class="w-full bg-blue-700 hover:bg-blue-600 active:bg-blue-800 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          ${d}
        </button>
      </div>
    </div>
  `,(D=document.getElementById("btn-logout"))==null||D.addEventListener("click",yt),(P=document.getElementById("submit-btn"))==null||P.addEventListener("click",At),(R=document.getElementById("btn-import-customer"))==null||R.addEventListener("click",Et),n==="phish"&&pt();const h=document.getElementById("f-severity"),$=document.getElementById("f-sla");if(h&&$){let u=!1;const g=document.getElementById("f-sla-hint"),O=()=>{if(!g)return;const A=parseInt($.value||"0")||null;g.textContent=A?`นับจากตอนนี้ · ครบกำหนด ${ve(A)}`:"ไม่กำหนด SLA — เคสนี้จะวัดไม่ได้ในรายงาน"};$.addEventListener("change",()=>{u=!0,O()}),h.addEventListener("change",()=>{if(u)return;const A=ne[h.value];A&&($.value=String(A),O())})}document.querySelectorAll(".tab-btn").forEach(u=>{u.addEventListener("click",()=>{const g=u.dataset.tab;g&&g!==i.tab&&(i.tab=g,z(),g==="phish"&&ee())})});const k=document.getElementById("drop-zone"),I=document.getElementById("f-files");k&&I&&(I.addEventListener("change",()=>{I.files&&ue(Array.from(I.files)),I.value=""}),k.addEventListener("dragover",u=>{u.preventDefault(),k.classList.add("border-blue-500","bg-blue-50")}),k.addEventListener("dragleave",()=>{k.classList.remove("border-blue-500","bg-blue-50")}),k.addEventListener("drop",u=>{var O;u.preventDefault(),k.classList.remove("border-blue-500","bg-blue-50");const g=Array.from(((O=u.dataTransfer)==null?void 0:O.files)??[]);g.length&&ue(g)})),document.querySelectorAll(".remove-dropped").forEach(u=>{u.addEventListener("click",()=>{const g=parseInt(u.dataset.remove??"-1");g>=0&&(i.droppedFiles.splice(g,1),z())})}),jt()}function ue(e){i.droppedFiles.push(...e),z()}document.addEventListener("paste",e=>{var a;if(!i.account)return;const t=Array.from(((a=e.clipboardData)==null?void 0:a.items)??[]),n=[];for(const o of t)if(o.kind==="file"){const r=o.getAsFile();if(r){const s=r.name&&r.name!=="image.png"?r.name:`screenshot-${new Date().toISOString().replace(/[:.]/g,"-").slice(0,19)}.png`;n.push(new File([r],s,{type:r.type}))}}n.length&&(e.preventDefault(),ue(n),B(`แนบไฟล์แล้ว: ${n.map(o=>o.name).join(", ")}`))});const w="w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white";function Ie(e){return e<1024?`${e} B`:e<1024*1024?`${(e/1024).toFixed(0)} KB`:`${(e/1024/1024).toFixed(1)} MB`}function G(){const e=i.emailAttachments,t=i.droppedFiles,n=c=>`
    <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
      <input type="checkbox" class="email-att-cb" data-att-id="${y(c.id)}" data-att-name="${y(c.name)}" data-att-item="${c.isItem?"1":"0"}" ${c.defaultOn?"checked":""} />
      <span class="flex-1 truncate">${c.isItem?"📧 ":c.isInline?"🖼️ ":""}${y(c.name)}</span>
      <span class="text-slate-400 flex-shrink-0">${Ie(c.size)}</span>
    </label>`,a=e.filter(c=>!c.isInline),o=e.filter(c=>c.isInline),r=e.length>0?`<div class="mb-2 space-y-1">
        ${a.length?`<p class="text-xs text-slate-500">📎 ไฟล์แนบจาก Email:</p>
        ${a.map(n).join("")}`:""}
        ${o.length?`<p class="text-xs text-slate-500 ${a.length?"pt-1":""}">🖼️ รูปในเนื้อเมล
          <span class="text-slate-400">(รูปเล็กมักเป็นโลโก้ในลายเซ็น — ติ๊กเพิ่มได้)</span></p>
        ${o.map(n).join("")}`:""}
      </div>`:"",s=t.length>0?`<div class="mt-2 space-y-1">
        ${t.map((c,p)=>`<div class="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 rounded px-2 py-1">
            <span class="text-base">${c.type.startsWith("image/")?"🖼️":"📄"}</span>
            <span class="flex-1 truncate">${y(c.name)}</span>
            <span class="text-slate-400">${Ie(c.size)}</span>
            <button type="button" data-remove="${p}"
              class="remove-dropped text-red-400 hover:text-red-600 font-bold leading-none">✕</button>
          </div>`).join("")}
      </div>`:"";return`<div class="space-y-1">
    <label class="block text-xs font-medium text-slate-600">ไฟล์แนบ</label>
    <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
      <input type="checkbox" id="f-attach-eml" />
      <span class="flex-1">📧 แนบอีเมลต้นฉบับ (.eml)</span>
    </label>
    ${r}
    <div id="drop-zone"
      class="relative border-2 border-dashed border-slate-300 rounded-lg p-4 text-center text-xs text-slate-400
             hover:border-blue-400 hover:bg-blue-50 transition cursor-pointer select-none">
      <div class="pointer-events-none">
        <div class="text-2xl mb-1">📂</div>
        <div>ลากไฟล์มาวาง หรือ <span class="text-blue-600 font-medium">คลิกเลือก</span></div>
        <div class="mt-0.5 text-slate-300">หรือกด <kbd class="bg-slate-100 text-slate-500 px-1 rounded">Ctrl+V</kbd> วางจาก clipboard ได้เลย</div>
      </div>
      <input id="f-files" type="file" multiple
        class="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
    </div>
    ${s}
  </div>`}function ie(e){var t;return`<select id="f-assigned-email" class="${w}">
    <option value="${y(e)}">${y(((t=i.account)==null?void 0:t.name)??e)} (ฉัน)</option>
    ${i.agents.filter(n=>n.email!==e).map(n=>`<option value="${y(n.email)}">${y(n.name)}</option>`).join("")}
  </select>`}function Q(e=!1){return i.projects.length===0?e?'<div class="text-xs text-slate-400">ไม่พบ Project ที่ Active</div>':'<div class="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-2">⚠️ ไม่พบ Project ที่ Active</div>':`<select id="f-project" class="${w}">
    <option value="">${e?"-- ไม่ผูกกับโครงการ --":"-- เลือก Project --"}</option>
    ${i.projects.map(t=>`<option value="${t.id}">${y(t.Title)}</option>`).join("")}
  </select>`}function E(e,t){return`
    <div class="space-y-1">
      <label class="block text-xs font-medium text-slate-600">${e}</label>
      ${t}
    </div>
  `}function y(e){return e.replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}async function Pt(){Xe({sharepointUrl:M,internalDomains:["itservices.co.th","rpaexpert.com","rpaexpert.onmicrosoft.com"],getToken:H,getGraphToken:()=>K(),account:()=>i.account?{name:i.account.name,username:i.account.username}:null,toast:(t,n)=>B(t,n??"success"),rerender:z,canWhitelist:()=>["Agent","Supervisor","Boss","Admin"].includes(i.myRole)}),await L.initialize(),await L.handleRedirectPromise();const e=L.getAllAccounts();if(e.length>0){i.account=e[0];try{await L.acquireTokenSilent({scopes:[X],account:e[0]}),await Promise.all([Oe(),_e(),Le(),Re()])}catch{i.account=null}}typeof Office<"u"?Office.onReady(t=>{var n;if(t.host===Office.HostType.Outlook){const a=Office.context.mailbox.item;if(a){i.emailSubject=a.subject??"";const o=a.from;o&&(i.emailSenderName=o.displayName??"",i.emailSenderEmail=o.emailAddress??"");const r=(((n=i.account)==null?void 0:n.username)??"").toLowerCase(),s=((o==null?void 0:o.emailAddress)??"").toLowerCase(),c=[...a.to??[],...a.cc??[]].map(l=>l.emailAddress).filter(Boolean);i.emailCc=[...new Set(c.map(l=>l.toLowerCase()))].filter(l=>l!==r&&l!==s);const p=a.attachments??[],v=20*1024;i.emailAttachments=p.filter(l=>l.attachmentType===Office.MailboxEnums.AttachmentType.File||l.attachmentType===Office.MailboxEnums.AttachmentType.Item).map(l=>({id:l.id,name:l.attachmentType===Office.MailboxEnums.AttachmentType.Item?`${(l.name||"email").replace(/\.eml$/i,"")}.eml`:l.name,size:l.size,isItem:l.attachmentType===Office.MailboxEnums.AttachmentType.Item,isInline:!!l.isInline,defaultOn:!l.isInline||l.size>=v})),a.body.getAsync(Office.CoercionType.Html,{asyncContext:{}},l=>{if(l.status===Office.AsyncResultStatus.Succeeded){let b=function(P,R=!1){if(P.nodeType===3){const A=P.textContent??"";return R&&A.trim()===""?"":A}const u=P,g=(u.tagName??"").toLowerCase();if(N.includes(g))return"";if(g==="br")return" ";if(g==="tr"){const A=[];for(let U=0;U<u.childNodes.length;U++){const be=u.childNodes[U],ye=(be.tagName??"").toLowerCase();(ye==="td"||ye==="th")&&A.push((be.textContent??"").replace(/\s+/g," ").trim())}return A.length?A.join("	")+`
`:""}if(d.includes(g)){let A="";for(let U=0;U<u.childNodes.length;U++)A+=b(u.childNodes[U],!0);return A}let O="";for(let A=0;A<u.childNodes.length;A++)O+=b(u.childNodes[A],!1);return T.includes(g)&&(O=`
`+O.trim()+`
`),O};const j=l.value,_=new DOMParser().parseFromString(j,"text/html"),N=["style","script","head","img","meta","link","noscript"],T=["p","div","li","h1","h2","h3","h4","h5","h6","blockquote"],d=["table","thead","tbody","tfoot"],k=b(_.body??_.documentElement).replace(/[ \t]{2,}/g," ").replace(/\n[ \t]+/g,`
`).replace(/\n{3,}/g,`

`).trim().split(`
`),I=[];let x="";for(const P of k)P.trim()===""?x&&(I.push(x.trim()),x=""):P.includes("	")?(x&&(I.push(x.trim()),x=""),I.push(P)):x=x?x+" "+P.trim():P.trim();x&&I.push(x.trim());const S=I.join(`
`),C=/\n([-_]{3,}|From:\s|Best regards|Regards,|ขอแสดงความนับถือ|Sent:\s)/i,D=S.search(C);i.signatureContact=D>80?kt(S.slice(D).trim()):null,i.emailBodyPreview=S.trim().slice(0,2e4),i.emailBodyReply=Ke(i.emailBodyPreview)}z()});return}}Ae(),z()}):(Ae(),z())}function Ae(){i.emailSubject="[DEV] Test Email Subject",i.emailSenderName="Test Sender",i.emailSenderEmail="test@example.com",i.emailBodyPreview="This is a placeholder email body for development mode.",i.emailBodyReply=i.emailBodyPreview}Pt().catch(e=>{console.error("Init error:",e);const t=document.getElementById("app");t&&(t.innerHTML=`<div class="p-4 text-red-600 text-sm">เกิดข้อผิดพลาด: ${String(e)}</div>`)});
