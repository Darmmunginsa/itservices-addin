import{P as Je}from"./PublicClientApplication-DLKYUtZW.js";import{a as ae,d as Ke,p as Ye,L as je,S as Pe}from"./analyzer-CWd3MChg.js";const Ze=/^\s*(Sent|To|Date|Cc|Subject|ส่ง|ถึง|วันที่|สำเนา|เรื่อง)\s*:/i;function Qe(e,t){const n=e[t];if(/^\s*-{2,}\s*(Original Message|Forwarded message|ข้อความต้นฉบับ)\s*-{2,}/i.test(n)||/^\s*_{5,}\s*$/.test(n)||/^\s*(On|เมื่อ)\b.{10,200}(wrote|เขียนว่า)\s*:\s*$/i.test(n))return!0;if(/^\s*(From|จาก)\s*:\s*\S/i.test(n)){for(let a=t+1;a<=t+3&&a<e.length;a++)if(Ze.test(e[a]))return!0;return!1}return!1}function Xe(e){const t=(e??"").replace(/\r\n/g,`
`),n=t.split(`
`);let a=-1;for(let s=0;s<n.length;s++){if(s>0&&Qe(n,s)){a=s;break}if(s>0&&/^\s*>/.test(n[s])&&/^\s*>/.test(n[s+1]??"")){a=s;break}}if(a<0)return{visible:t.trim(),quoted:""};const o=n.slice(0,a).join(`
`).trim(),r=n.slice(a).join(`
`).trim();return o?{visible:o,quoted:r}:{visible:t.trim(),quoted:""}}const et=e=>Xe(e).visible,xe=e=>(typeof e=="string"?e:"").trim().toLowerCase();function _e(e,t){const n=xe(t);return e.find(a=>xe(a.EventKey)===n&&tt(a.IsEnabled))}function tt(e){if(e==null)return!0;if(typeof e=="boolean")return e;const t=String(e).trim().toLowerCase();return!["false","no","0","ไม่","ปิด",""].includes(t)}const nt=e=>({__html:e??""}),De=e=>!!e&&typeof e=="object"&&typeof e.__html=="string";function Oe(e){return(e??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}const le=e=>nt(Oe(e??"").replace(/\r?\n/g,"<br>"));function Le(e,t){const n=[];return{text:(e??"").replace(/\{\{(\w+)\}\}/g,(o,r)=>{const s=t[r];return s===void 0?(n.push(r),""):De(s)?s.__html:Oe(s)}),missing:[...new Set(n)]}}function at(e,t){return(e??"").replace(/\{\{(\w+)\}\}/g,(n,a)=>{const o=t[a];return o===void 0?"":(De(o)?o.__html.replace(/<[^>]*>/g,""):o).replace(/\s+/g," ").trim()})}const ve=e=>(e??"").trim().toLowerCase();function st(e){const t=ve(e.actorEmail),n=new Set,a=s=>{const c=[];for(const m of s){const x=(m??"").trim(),l=ve(x);!l||!l.includes("@")||l===t||n.has(l)||(n.add(l),c.push(x))}return c};let o=a([e.assignedEmail]);const r=a([e.requesterEmail,...e.watchers??[]]);return o.length===0?(o=r.slice(0,1),{to:o,cc:r.slice(1)}):{to:o,cc:r}}const ot=e=>{if(!e||e<=0)return"ไม่ได้กำหนด";if(e<24)return`${e} ชั่วโมง`;const t=e/24;return Number.isInteger(t)?`${t} วัน`:`${e} ชั่วโมง`};function it(e){const t=(e.baseUrl??"").replace(/\/+$/,"");return{incident_title:e.title??"",severity:e.severity??"",status:e.status??"",incident_status:e.status??"",description:le(e.description),resolution:le(e.resolution),incident_date:(e.incidentDate??"").slice(0,10),sla_hours:ot(e.slaHours),project_name:e.projectName??"",assigned_name:e.assignedName||e.assignedEmail||"-",link:e.projectId?`${t}/#/projects/${e.projectId}`:t}}function rt(e){const{to:t,cc:n}=st(e);return{to:t,cc:n,vars:it(e)}}const ct=[{hours:1,labelTh:"1 ชั่วโมง"},{hours:2,labelTh:"2 ชั่วโมง"},{hours:4,labelTh:"4 ชั่วโมง"},{hours:8,labelTh:"8 ชั่วโมง (1 วันทำการ)"},{hours:24,labelTh:"24 ชั่วโมง"},{hours:48,labelTh:"2 วัน"},{hours:72,labelTh:"3 วัน"},{hours:168,labelTh:"7 วัน"}],se={Critical:1,High:4,Medium:24,Low:72};function Re(e,t,n=new Date){const a=typeof e=="number"&&Number.isFinite(e)&&e>0?e:null;if(!a)return null;const o=n,r=isNaN(o.getTime())?n:o;return new Date(r.getTime()+a*36e5).toISOString()}function we(e){const t=Re(e);return t?new Date(t).toLocaleString("th-TH",{dateStyle:"short",timeStyle:"short"}):""}const Me="HD_PhishingReports";let p;const d={mail:null,analysis:null,loading:!1,reporting:!1,reported:!1,showHeaders:!1,kasmTemplate:"",analysedItemId:"",safeDomains:[],safeDomainIds:{},savingDomain:""};function lt(e){p=e}const z=e=>(e??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"),Y=()=>{var e,t;return(t=(e=Office.context)==null?void 0:e.mailbox)==null?void 0:t.item},$e=e=>new Promise(t=>{const n=Y();if(!(n!=null&&n.body)){t("");return}n.body.getAsync(e,a=>t(a.status===Office.AsyncResultStatus.Succeeded?a.value??"":""))});function dt(e){return e?e.split(",").map(t=>{const n=t.trim(),a=n.match(/^(.*?)\s*<([^>]+)>$/);return a?{name:a[1].replace(/^"|"$/g,"").trim(),email:a[2].trim()}:{name:"",email:n.replace(/[<>]/g,"").trim()}}).filter(t=>t.email.includes("@")):[]}function mt(){return new Promise(e=>{let t=!1;try{t=Office.context.requirements.isSetSupported("Mailbox","1.8")}catch{t=!1}const n=Y();if(!t||typeof(n==null?void 0:n.getAllInternetHeadersAsync)!="function"){e({});return}try{n.getAllInternetHeadersAsync(a=>e(a.status===Office.AsyncResultStatus.Succeeded?Ye(a.value??""):{}))}catch{e({})}})}async function ut(){try{const e=Y();if(!(e!=null&&e.itemId))return{};const t=Office.context.mailbox.convertToRestId(e.itemId,Office.MailboxEnums.RestVersion.v2_0),n=await p.getGraphToken(),a=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${t}?$select=internetMessageHeaders`,{headers:{Authorization:`Bearer ${n}`}});if(!a.ok)return{};const o=await a.json(),r={};for(const s of o.internetMessageHeaders??[])r[s.name]=r[s.name]?`${r[s.name]}
${s.value}`:s.value;return r}catch{return{}}}async function he(e,t){const n=await p.getToken(),a=await fetch(`${p.sharepointUrl}/_api/web/lists/getbytitle('${e}')/items?${t}`,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});return a.ok?(await a.json()).value:[]}async function pt(){try{return(await he("HD_AgentProfiles","$select=Title,EmailText&$top=500")).filter(t=>t.EmailText).map(t=>({name:t.Title,email:t.EmailText}))}catch{return[]}}async function ft(){var e;try{return(((e=(await he("HD_Options","$select=Title&$filter=Category eq 'KasmConfig'&$top=1"))[0])==null?void 0:e.Title)??"").trim()}catch{return""}}const He="SafeDomain";async function ge(){try{const e=await he("HD_Options",`$select=Id,Title&$filter=Category eq '${He}'&$top=500`),t={},n=[];for(const a of e){const o=(a.Title??"").trim().toLowerCase();o&&(t[o]=a.Id,n.push(o))}d.safeDomains=n,d.safeDomainIds=t}catch{}}async function ht(e){const t=e.trim().toLowerCase();if(!(!t||d.savingDomain)){if(!p.canWhitelist()){p.toast("ต้องเป็น Agent ขึ้นไปจึงจะยืนยันโดเมนได้","error");return}d.savingDomain=t,p.rerender();try{const n=await p.getToken(),a=await fetch(`${p.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items`,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:t,Category:He})});if(!a.ok)throw new Error(String(a.status));await ge(),await ne(!0),p.toast(`ยืนยันแล้วว่า ${t} ปลอดภัย`)}catch{p.toast("บันทึกไม่สำเร็จ","error")}finally{d.savingDomain="",p.rerender()}}}async function gt(e){const t=e.trim().toLowerCase(),n=d.safeDomainIds[t];if(!(!n||d.savingDomain)){if(!p.canWhitelist()){p.toast("ต้องเป็น Agent ขึ้นไปจึงจะถอนโดเมนได้","error");return}d.savingDomain=t,p.rerender();try{const a=await p.getToken(),o=await fetch(`${p.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items(${n})`,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","IF-MATCH":"*","X-HTTP-Method":"DELETE"}});if(!o.ok)throw new Error(String(o.status));await ge(),await ne(!0),p.toast(`ถอน ${t} ออกจากรายการปลอดภัยแล้ว`)}catch{p.toast("ถอนไม่สำเร็จ","error")}finally{d.savingDomain="",p.rerender()}}}async function ne(e=!1){var m,x;const t=Y(),n=(t==null?void 0:t.itemId)??"";if(!e&&d.analysedItemId===n&&d.analysis)return;d.loading=!0,d.reported=!1,d.analysedItemId=n,p.rerender();const[a,o]=await Promise.all([$e(Office.CoercionType.Html),$e(Office.CoercionType.Text)]),r={fromName:((m=t==null?void 0:t.from)==null?void 0:m.displayName)??"",fromEmail:((x=t==null?void 0:t.from)==null?void 0:x.emailAddress)??"",replyTo:[],subject:(t==null?void 0:t.subject)??"",bodyHtml:a,bodyText:o,attachments:((t==null?void 0:t.attachments)??[]).map(l=>({name:l.name,size:l.size??0,isInline:!!l.isInline})),headers:{},internalDomains:p.internalDomains,internalPeople:[],safeDomains:d.safeDomains},s=l=>{var f;return{...r,headers:l,replyTo:dt(((f=Object.entries(l).find(([j])=>j.toLowerCase()==="reply-to"))==null?void 0:f[1])??"")}},c=await mt();if(d.mail=s(c),d.analysis=ae(d.mail),d.loading=!1,p.rerender(),p.account()){const[l,f]=await Promise.all([pt(),Object.keys(c).length?Promise.resolve(c):ut()]);d.mail={...s(f),internalPeople:l},d.analysis=ae(d.mail),d.kasmTemplate||(d.kasmTemplate=await ft()),d.safeDomains.length||(await ge(),d.safeDomains.length&&(d.mail={...d.mail,safeDomains:d.safeDomains},d.analysis=ae(d.mail))),p.rerender()}}function Ne(){const e=d.mail,t=d.analysis;return!e||!t?"":[`ผู้ส่ง: ${e.fromName} <${e.fromEmail}>`,`หัวข้อ: ${e.subject}`,e.replyTo.length?`Reply-To: ${e.replyTo.map(n=>n.email).join(", ")}`:"",`คะแนนความเสี่ยง: ${t.score} (${je[t.level].label})`,"","สิ่งที่ตรวจพบ:",...t.findings.map(n=>`- [${Pe[n.severity].label}] (${n.category}) ${n.title} — ${n.detail.replace(/\n/g," ")}`),"",t.links.length?"ลิงก์ในอีเมล:":"",...t.links.map(n=>`- ${n.href}${n.flags.length?`  ! ${n.flags.join(" / ")}`:""}`)].filter(n=>n!=="").join(`
`)}async function bt(e){try{const t=Y();if(!(t!=null&&t.itemId))return!1;const n=Office.context.mailbox.convertToRestId(t.itemId,Office.MailboxEnums.RestVersion.v2_0),a=await p.getGraphToken(),o=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${n}/$value`,{headers:{Authorization:`Bearer ${a}`}});if(!o.ok)return!1;const r=await o.arrayBuffer(),s=(t.subject||"phishing").replace(/[\\/:*?"<>|#%&{}~]/g,"_").replace(/^_+/,"").slice(0,80).trim()||"phishing",c=await p.getToken();return(await fetch(`${p.sharepointUrl}/_api/web/lists/getbytitle('${Me}')/items(${e})/AttachmentFiles/add(FileName='${encodeURIComponent(s+".eml")}')`,{method:"POST",headers:{Authorization:`Bearer ${c}`,Accept:"application/json;odata=nometadata"},body:r})).ok}catch{return!1}}async function Te(e){const t=await p.getToken(),n=await fetch(`${p.sharepointUrl}/_api/web/lists/getbytitle('${Me}')/items`,{method:"POST",headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(e)});if(!n.ok)throw new Error(`SharePoint ${n.status}: ${await n.text()}`);return(await n.json()).Id}async function yt(){if(!(!d.mail||!d.analysis||d.reporting)){if(!p.account()){p.toast("กรุณาเข้าสู่ระบบก่อนรายงาน","error");return}d.reporting=!0,p.rerender();try{const e=d.mail,t=d.analysis,n=p.account(),a={Title:(e.subject||"(ไม่มีหัวข้อ)").slice(0,255),SenderName:e.fromName.slice(0,255),SenderEmail:e.fromEmail.slice(0,255),SenderDomain:Ke(e.fromEmail),RiskScore:t.score,RiskLevel:t.level,Findings:Ne(),LinkCount:t.links.length,SuspiciousLinks:t.links.filter(c=>c.flags.length).map(c=>c.href).join(`
`).slice(0,4e3),ReportedBy:(n==null?void 0:n.name)??"",ReportedEmail:(n==null?void 0:n.username)??"",Status:"New"};let o,r=!1;try{o=await Te(a)}catch(c){o=await Te({Title:a.Title,Findings:a.Findings}).catch(()=>{throw c}),r=!0}const s=await bt(o);d.reported=!0,p.toast(r?"ส่งรายงานแล้ว แต่บันทึกได้บางคอลัมน์ — ตรวจชื่อคอลัมน์ใน HD_PhishingReports":s?"ส่งรายงานพร้อมอีเมลต้นฉบับแล้ว":"ส่งรายงานแล้ว (แนบ .eml ไม่ได้)",r?"info":"success")}catch(e){p.toast(`ส่งรายงานไม่สำเร็จ: ${e instanceof Error?e.message:String(e)}`,"error")}finally{d.reporting=!1,p.rerender()}}}const xt=()=>d.reported?"✓ รายงานแล้ว":d.reporting?"กำลังส่ง…":"🚩 รายงานอีเมลนี้ให้ IT";function vt(e){var t;try{const n=(t=Office.context)==null?void 0:t.ui;if(typeof(n==null?void 0:n.openBrowserWindow)=="function"){n.openBrowserWindow(e);return}}catch{}window.open(e,"_blank","noopener,noreferrer")||p.toast("เปิดหน้าต่างไม่ได้ (ถูกบล็อก)","info")}async function wt(e){var t;try{if((t=navigator.clipboard)!=null&&t.writeText)return await navigator.clipboard.writeText(e),!0}catch{}try{const n=document.createElement("textarea");n.value=e,n.setAttribute("readonly",""),n.style.position="fixed",n.style.opacity="0",document.body.appendChild(n),n.focus(),n.select();const a=document.execCommand("copy");return n.remove(),a}catch{return!1}}function $t(e){if(d.loading&&!d.analysis)return`<div class="py-10 text-center text-slate-500 text-sm">
      <div class="w-7 h-7 border-[3px] border-slate-200 border-t-blue-600 rounded-full animate-spin mx-auto mb-2"></div>
      กำลังตรวจอีเมล…</div>`;const t=d.analysis,n=d.mail;if(!t||!n)return'<p class="text-sm text-slate-400 text-center py-8">เปิดอีเมลเพื่อเริ่มตรวจ</p>';const a=je[t.level],o=[...t.links.filter(s=>s.flags.length),...t.links.filter(s=>!s.flags.length&&!s.trusted),...t.links.filter(s=>s.trusted)],r=p.canWhitelist();return`
    <div class="rounded-xl border-2 ${a.cls} p-3">
      <div class="flex items-center gap-2">
        <span class="text-2xl leading-none">${a.icon}</span>
        <div class="min-w-0 flex-1">
          <div class="font-bold text-sm">${z(a.label)}</div>
          <div class="text-xs opacity-80">คะแนน ${t.score} · พบสัญญาณ ${t.findings.filter(s=>s.severity!=="info").length} ข้อ</div>
        </div>
        <button id="phish-recheck" class="text-[10px] px-2 py-1 rounded-md bg-white/70 hover:bg-white text-slate-700 font-medium">ตรวจใหม่</button>
      </div>
    </div>

    ${e?"":`<div class="bg-blue-50 border border-blue-200 rounded-xl p-2.5 text-[11px] text-blue-800">
      เข้าสู่ระบบเพื่อตรวจการปลอมเป็นพนักงาน และรายงานเข้า Helpdesk ได้
    </div>`}

    ${t.findings.length===0?'<p class="text-xs text-slate-400 text-center py-3">ไม่พบสัญญาณผิดปกติจากการตรวจอัตโนมัติ</p>':t.findings.map(s=>{const c=Pe[s.severity];return`<div class="bg-white rounded-xl border border-slate-200 p-2.5">
          <div class="flex items-start gap-2">
            <span class="text-[9px] font-bold px-1.5 py-0.5 rounded-full ${c.cls} flex-shrink-0 mt-0.5">${c.label}</span>
            <div class="min-w-0 flex-1">
              <div class="text-xs font-semibold text-slate-800">${z(s.title)}</div>
              <div class="text-[11px] text-slate-500 whitespace-pre-line break-all">${z(s.detail)}</div>
              <div class="text-[9px] text-slate-400 mt-0.5">${z(s.category)}</div>
            </div>
          </div>
        </div>`}).join("")}

    ${o.length?`
    <div class="bg-white rounded-xl border border-slate-200 p-3">
      <div class="text-xs font-semibold text-slate-700 mb-2">ลิงก์ในอีเมล (${o.length})</div>
      <div class="space-y-2">
        ${o.map((s,c)=>{const m=d.savingDomain===(s.host?s.host.toLowerCase():"");return`
          <div class="rounded-lg border ${s.trusted?"border-emerald-200 bg-emerald-50/50":s.flags.length?"border-red-200 bg-red-50/50":"border-slate-100"} p-2">
            <div class="text-[11px] font-medium ${s.trusted?"text-emerald-800":s.flags.length?"text-red-700":"text-slate-700"} break-all">
              ${s.trusted?"✔ ":""}${z(s.host||s.href)}
            </div>
            ${s.trusted?`<div class="text-[10px] text-emerald-700">ทีมตรวจแล้วว่าปลอดภัย${(s.suppressed??[]).length?` · ระงับการเตือน ${s.suppressed.length} ข้อ`:""}</div>`:""}
            ${s.text&&s.text!==s.href?`<div class="text-[10px] text-slate-500 break-all">แสดงว่า: "${z(s.text.slice(0,70))}"</div>`:""}
            <div class="text-[10px] text-slate-400 break-all mt-0.5">${z(s.href.slice(0,150))}</div>
            ${s.flags.map(l=>`<div class="text-[10px] text-red-600 mt-0.5">! ${z(l)}</div>`).join("")}
            <div class="flex flex-wrap gap-1 mt-1.5">
              <button data-kasm="${c}" class="text-[10px] font-semibold px-2 py-1 rounded-md bg-slate-800 hover:bg-slate-900 text-white">
                เปิดใน Kasm
              </button>
              ${r&&s.host?s.trusted?`<button data-untrust="${z(s.host)}" ${m?"disabled":""}
                     class="text-[10px] font-semibold px-2 py-1 rounded-md border border-slate-300 text-slate-600 hover:bg-slate-100 disabled:opacity-50">
                     ${m?"...":"ถอนออกจากรายการปลอดภัย"}</button>`:`<button data-trust="${z(s.host)}" ${m?"disabled":""}
                     class="text-[10px] font-semibold px-2 py-1 rounded-md bg-emerald-600 hover:bg-emerald-700 text-white disabled:opacity-50">
                     ${m?"...":"✔ ตรวจแล้ว ปลอดภัย"}</button>`:""}
            </div>
          </div>`}).join("")}
      </div>
      <p class="text-[10px] text-slate-400 mt-2">อย่าคลิกลิงก์จากอีเมลที่ไม่มั่นใจโดยตรง</p>
    </div>`:""}

    <div class="flex gap-2">
      <button id="phish-copy" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">คัดลอกผลตรวจ</button>
      <button id="phish-headers" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">${d.showHeaders?"ซ่อน header":"ดู header"}</button>
    </div>
    ${d.showHeaders?`<pre class="bg-slate-900 text-slate-100 text-[9px] p-2 rounded-lg overflow-x-auto whitespace-pre-wrap break-all max-h-56 overflow-y-auto">${z(Object.entries(n.headers).map(([s,c])=>`${s}: ${c}`).join(`
`)||"อ่าน header ไม่ได้")}</pre>`:""}
    <p class="text-[10px] text-slate-400 text-center">วิเคราะห์ในเครื่องทั้งหมด — ไม่ส่งเนื้อหาอีเมลออกนอกองค์กร</p>
  `}function Tt(){var n,a,o,r;(n=document.getElementById("phish-recheck"))==null||n.addEventListener("click",()=>{ne(!0)}),(a=document.getElementById("phish-headers"))==null||a.addEventListener("click",()=>{d.showHeaders=!d.showHeaders,p.rerender()}),(o=document.getElementById("phish-copy"))==null||o.addEventListener("click",async()=>{const s=await wt(Ne());p.toast(s?"คัดลอกผลตรวจแล้ว":"คัดลอกไม่ได้",s?"success":"error")}),document.querySelectorAll("[data-trust]").forEach(s=>s.addEventListener("click",()=>ht(s.dataset.trust??""))),document.querySelectorAll("[data-untrust]").forEach(s=>s.addEventListener("click",()=>gt(s.dataset.untrust??"")));const e=((r=d.analysis)==null?void 0:r.links)??[],t=[...e.filter(s=>s.flags.length),...e.filter(s=>!s.flags.length&&!s.trusted),...e.filter(s=>s.trusted)];document.querySelectorAll("[data-kasm]").forEach(s=>{s.addEventListener("click",()=>{const c=t[Number(s.dataset.kasm)];if(!c)return;if(!d.kasmTemplate){p.toast("ยังไม่ได้ตั้งค่า Kasm ใน HD_Options (Category=KasmConfig)","info");return}const m=d.kasmTemplate;vt(m.includes("{url}")?m.replace("{url}",encodeURIComponent(c.href)):m+encodeURIComponent(c.href))})})}const kt="0bab07cf-65e6-487c-89af-c917fc1a5a13",Et="d569b991-89fc-4a62-9df5-eb361abcef40",M="https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd",te="https://rpaexpert.sharepoint.com/.default",oe=["https://graph.microsoft.com/Calendars.ReadWrite","https://graph.microsoft.com/Mail.Send","https://graph.microsoft.com/Mail.Read"],L=new Je({auth:{clientId:kt,authority:`https://login.microsoftonline.com/${Et}`,redirectUri:window.location.origin.includes("localhost")?"http://localhost:3000/":"https://darmmunginsa.github.io/itservices-addin/",navigateToLoginRequestUrl:!1},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:!1}}),It=window.location.origin.includes("localhost")?`${window.location.origin}/auth.html`:"https://darmmunginsa.github.io/itservices-addin/auth.html";function be(){var t,n;const e=(n=(t=Office.context)==null?void 0:t.diagnostics)==null?void 0:n.platform;return e===Office.PlatformType.iOS||e===Office.PlatformType.Android}function ye(){return new Promise((e,t)=>{Office.context.ui.displayDialogAsync(It,{height:60,width:30,promptBeforeOpen:!1},n=>{if(n.status!==Office.AsyncResultStatus.Succeeded){t(new Error("เปิดหน้าเข้าสู่ระบบไม่ได้"));return}const a=n.value;a.addEventHandler(Office.EventType.DialogMessageReceived,o=>{a.close();const r=o.message;if(!r){t(new Error("auth message error"));return}try{const s=JSON.parse(r);s.ok?e():t(new Error(s.error||"auth failed"))}catch{t(new Error("auth message error"))}}),a.addEventHandler(Office.EventType.DialogEventReceived,()=>t(new Error("ปิดหน้าเข้าสู่ระบบก่อนเสร็จ")))})})}const i={account:null,tab:"ticket",emailSubject:"",emailBodyPreview:"",emailBodyReply:"",emailSenderName:"",emailSenderEmail:"",loading:!1,projects:[],agents:[],myRole:"",emailAttachments:[],signatureContact:null,droppedFiles:[],tickets:[],contactEmails:[],emailCc:[]};async function H(){const e=L.getAllAccounts();if(e.length===0)throw new Error("Not signed in");const t={scopes:[te],account:e[0]};try{return(await L.acquireTokenSilent(t)).accessToken}catch{if(be()){await ye();const n=L.getAllAccounts()[0];if(!n)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await L.acquireTokenSilent({scopes:[te],account:n})).accessToken}return(await L.acquireTokenPopup(t)).accessToken}}async function K(e=!1){const t=L.getAllAccounts();if(t.length===0)throw new Error("Not signed in");const n={scopes:oe,account:t[0],forceRefresh:e};try{return(await L.acquireTokenSilent(n)).accessToken}catch{if(be()){await ye();const o=L.getAllAccounts()[0];if(!o)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await L.acquireTokenSilent({scopes:oe,account:o})).accessToken}return(await L.acquireTokenPopup({scopes:oe,account:t[0]})).accessToken}}async function At(e){const t=await K(),n={subject:e.subject,start:{dateTime:e.start,timeZone:"Asia/Bangkok"},end:{dateTime:e.end,timeZone:"Asia/Bangkok"},body:e.body?{contentType:"HTML",content:e.body.replace(/\n/g,"<br>")}:void 0,attendees:e.attendees.filter(Boolean).map(o=>({emailAddress:{address:o},type:"required"})),isOnlineMeeting:e.isOnlineMeeting,onlineMeetingProvider:e.isOnlineMeeting?"teamsForBusiness":void 0},a=await fetch("https://graph.microsoft.com/v1.0/me/events",{method:"POST",headers:{Authorization:`Bearer ${t}`,"Content-Type":"application/json"},body:JSON.stringify(n)});if(!a.ok)throw new Error(`Calendar error ${a.status}: ${await a.text()}`)}async function Fe(){try{const e=await H(),t=`${M}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$orderby=Title asc&$top=500`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const a=await n.json();i.projects=a.value.map(o=>({id:o.Id,Title:o.Title}))}}catch{}}async function ze(){var e,t;try{const n=await H(),a=`${M}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText,Role&$orderby=Title asc`,o=await fetch(a,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});if(o.ok){const r=await o.json();i.agents=r.value.map(c=>({email:c.EmailText,name:c.Title}));const s=(((e=i.account)==null?void 0:e.username)??"").toLowerCase();i.myRole=((t=r.value.find(c=>(c.EmailText??"").toLowerCase()===s))==null?void 0:t.Role)??""}}catch{}}async function qe(){try{const e=await H(),t=`${M}/_api/web/lists/getbytitle('HD_Tickets')/items?$select=Id,Title,TicketNumber,Status&$filter=Status ne 'Closed'&$orderby=Modified desc&$top=200`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const a=await n.json();i.tickets=a.value.map(o=>({id:o.Id,Title:o.Title,TicketNumber:o.TicketNumber,Status:o.Status}))}}catch{}}async function Ue(){try{const e=await H(),t=`${M}/_api/web/lists/getbytitle('HD_Contracts')/items?$select=CustomerEmail&$top=2000`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const a=await n.json();i.contactEmails=a.value.map(o=>(o.CustomerEmail||"").trim().toLowerCase()).filter(Boolean)}}catch{}}async function ke(){const e=document.getElementById("btn-login-main"),t=document.getElementById("btn-login");e&&(e.disabled=!0,e.textContent="กำลังเข้าสู่ระบบ…"),t&&(t.disabled=!0);try{if(be()){if(await ye(),i.account=L.getAllAccounts()[0]??null,!i.account)throw new Error("เข้าสู่ระบบไม่สำเร็จ")}else{const n=await L.loginPopup({scopes:[te]});i.account=n.account}await Promise.all([Fe(),ze(),qe(),Ue()]),q()}catch{e&&(e.disabled=!1,e.textContent="เข้าสู่ระบบ"),t&&(t.disabled=!1)}}async function St(){i.account&&await L.logoutPopup({account:i.account}),i.account=null,q()}function Ee(e,t){const n=(e??"").trim().toLowerCase();return n?{IsAcknowledged:n===(t??"").trim().toLowerCase()}:{IsAcknowledged:!1}}async function V(e,t){const n=await H(),a=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items`,o=await fetch(a,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(t)});if(!o.ok){const s=await o.text(),c=["IsAcknowledged","AcknowledgedBy","AcknowledgedDate"];if(c.some(m=>m in t)){const m={...t};for(const l of c)delete m[l];const x=await fetch(a,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(m)});if(x.ok)return console.warn(`[SP] ${e} ไม่มีคอลัมน์ IsAcknowledged — สร้างแล้วแต่ไม่ได้เข้ากล่องรอรับงาน`),(await x.json()).Id}throw new Error(`SharePoint error ${o.status}: ${s}`)}return(await o.json()).Id}let Z=null,Ie=0;const Ct=5*60*1e3,ie="support@itservices.co.th",Ve="engineer@itservices.co.th";async function Ge(){if(Z&&Date.now()-Ie<Ct)return Z;try{const e=await H(),t=`${M}/_api/web/lists/getbytitle('HD_EmailTemplates')/items?$select=EventKey,Subject,Body,IsEnabled&$top=50`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});return n.ok?(Z=(await n.json()).value,Ie=Date.now(),Z):[]}catch{return[]}}async function Bt(){var e,t;try{const n=await H(),a=`${M}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title,Category&$filter=Category eq 'EmailConfig'&$top=1`,o=await fetch(a,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});return o.ok&&((t=(e=(await o.json()).value[0])==null?void 0:e.Title)==null?void 0:t.trim())||ie}catch{return ie}}async function Ae(e,t,n,a=[]){try{const o=await Ge(),r=_e(o,e);if(!r)return console.warn(`[mail] ไม่พบ template "${e}" ที่เปิดใช้`),!1;const s=at(r.Subject||"",t),c=Le(r.Body||"",t).text;if(!s||!c)return console.warn(`[mail] template "${e}" Subject/Body ว่าง`),!1;const m=I=>I.trim().toLowerCase(),x=[...new Map(n.filter(Boolean).map(I=>[m(I),I])).values()];if(x.length===0)return!1;const l=new Set(x.map(m)),f=e==="ticket_created"?[...a,Ve]:a,j=[...new Map(f.filter(Boolean).map(I=>[m(I),I])).values()].filter(I=>!l.has(m(I))),O=await Bt(),N=await K(),E={subject:s,body:{contentType:"HTML",content:c},toRecipients:x.map(I=>({emailAddress:{address:I}}))};j.length&&(E.ccRecipients=j.map(I=>({emailAddress:{address:I}}))),O&&(E.from={emailAddress:{address:O}});const R=await fetch("https://graph.microsoft.com/v1.0/me/sendMail",{method:"POST",headers:{Authorization:`Bearer ${N}`,"Content-Type":"application/json"},body:JSON.stringify({message:E,saveToSentItems:!0})});return R.ok?!0:(console.warn(`[mail] sendMail ${R.status}`,await R.text().catch(()=>"")),!1)}catch(o){return console.warn("[mail] ส่งไม่สำเร็จ",o),!1}}async function jt(e,t=[]){try{const n=Office.context.mailbox.item;if(!(n!=null&&n.itemId))return!1;const a=Office.context.mailbox.convertToRestId(n.itemId,Office.MailboxEnums.RestVersion.v2_0),r={Authorization:`Bearer ${await K()}`,"Content-Type":"application/json"},s=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${a}/createReplyAll`,{method:"POST",headers:r});if(!s.ok)return!1;const c=await s.json(),m=E=>E.trim().toLowerCase(),x=c.ccRecipients??[],l=new Set(x.map(E=>m(E.emailAddress.address))),f=[...new Set(t.filter(Boolean).map(E=>E.trim()))].filter(E=>!l.has(m(E))).map(E=>({emailAddress:{address:E}})),j={body:{contentType:"HTML",content:e}};return f.length&&(j.ccRecipients=[...x,...f]),(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${c.id}`,{method:"PATCH",headers:r,body:JSON.stringify(j)})).ok?(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${c.id}/send`,{method:"POST",headers:r})).ok:!1}catch{return!1}}async function Pt(e,t){const n=await Ge(),a=_e(n,e);return a&&Le(a.Body||"",t).text||null}async function Q(e){var r;const t=s=>s.trim().toLowerCase(),n=t(((r=i.account)==null?void 0:r.username)??""),a=new Set,o=e.recipients.filter(Boolean).filter(s=>{const c=t(s);return!c||c===n||a.has(c)?!1:(a.add(c),!0)});if(o.length!==0)try{const s=await H(),c=`${M}/_api/web/lists/getbytitle('HD_Notifications')/items`;await Promise.all(o.map(m=>fetch(c,{method:"POST",headers:{Authorization:`Bearer ${s}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:e.title.slice(0,255),RecipientEmail:m,EventType:e.eventType,Message:e.message,LinkPath:e.linkPath,IsRead:!1})})))}catch{}}async function de(e,t){const n=document.querySelectorAll(".email-att-cb:checked");if(n.length===0)return;const a=await H(),o=new Set,r=s=>{if(!o.has(s.toLowerCase()))return o.add(s.toLowerCase()),s;const c=s.lastIndexOf("."),m=c>0?s.slice(0,c):s,x=c>0?s.slice(c):"";for(let l=2;;l++){const f=m+"-"+l+x;if(!o.has(f.toLowerCase()))return o.add(f.toLowerCase()),f}};for(const s of Array.from(n)){const c=s.dataset.attId,m=r(s.dataset.attName),x=await new Promise((R,I)=>{Office.context.mailbox.item.getAttachmentContentAsync(c,{},h=>{h.status===Office.AsyncResultStatus.Succeeded?R(h):I(new Error(h.error.message))})}),{content:l,format:f}=x.value;let j;if(f===Office.MailboxEnums.AttachmentContentFormat.Base64){const R=atob(l);j=new Uint8Array(R.length);for(let I=0;I<R.length;I++)j[I]=R.charCodeAt(I)}else if(f===Office.MailboxEnums.AttachmentContentFormat.Eml||f===Office.MailboxEnums.AttachmentContentFormat.ICalendar)j=new TextEncoder().encode(l);else continue;const O=encodeURIComponent(m),N=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${O}')`;if(!(await fetch(N,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:j.buffer})).ok)throw new Error(`Upload ${m} failed`)}}async function _t(e){const t=`https://graph.microsoft.com/v1.0/me/messages/${e}/$value`;let n=await K(),a=await fetch(t,{headers:{Authorization:`Bearer ${n}`}});if((a.status===401||a.status===403)&&(n=await K(!0),a=await fetch(t,{headers:{Authorization:`Bearer ${n}`}})),!a.ok)throw new Error(`Graph ${a.status}`);return a.arrayBuffer()}async function Dt(e){const t=await new Promise((a,o)=>{Office.context.mailbox.getCallbackTokenAsync({isRest:!0},r=>{r.status===Office.AsyncResultStatus.Succeeded?a(r.value):o(new Error("callback token failed"))})}),n=await fetch(`${Office.context.mailbox.restUrl}/v2.0/me/messages/${e}/$value`,{headers:{Authorization:`Bearer ${t}`}});if(!n.ok)throw new Error(`REST ${n.status}`);return n.arrayBuffer()}async function me(e,t){const n=document.getElementById("f-attach-eml");if(!(n!=null&&n.checked))return;const a=Office.context.mailbox.item;if(!a)return;const o=Office.context.mailbox.convertToRestId(a.itemId,Office.MailboxEnums.RestVersion.v2_0);let r,s="",c="";try{r=await _t(o)}catch(j){s=j instanceof Error?j.message:String(j);try{r=await Dt(o)}catch(O){c=O instanceof Error?O.message:String(O),console.error("[eml] graph:",s,"| callback:",c),C(`ดึง .eml ไม่ได้ (Graph: ${s} / REST: ${c}) — ไฟล์อื่นบันทึกแล้ว`,"error");return}}const m=(a.subject||"email").replace(/[\\/:*?"<>|#%&{}~]/g,"_").slice(0,100).trim()||"email",x=await H(),l=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${encodeURIComponent(m+".eml")}')`;(await fetch(l,{method:"POST",headers:{Authorization:`Bearer ${x}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok||C("บันทึกไฟล์ .eml ไม่สำเร็จ","error")}async function ue(e,t,n){const a=await H();for(const o of n){const r=await o.arrayBuffer(),s=encodeURIComponent(o.name),c=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${s}')`;if(!(await fetch(c,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok)throw new Error(`Upload ${o.name} failed`)}}function C(e,t="success"){const n=document.getElementById("toast-container");if(!n)return;const a=t==="success"?"bg-green-500":t==="error"?"bg-red-500":"bg-slate-700",o=t==="success"?"✅":t==="error"?"❌":"ℹ️",r=document.createElement("div");r.className=`toast pointer-events-auto ${a} text-white text-sm font-medium px-4 py-3 rounded-lg shadow-lg max-w-xs mx-2`,r.textContent=`${o} ${e}`,n.appendChild(r),setTimeout(()=>r.remove(),4e3)}function Ot(e){const t=e.split(`
`).map(l=>l.trim()).filter(Boolean),n=/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/,a=/(\+?[\d\s()\-]{7,}(?:ext\.?\s*\d+)?)/i,o=/\b(co\.?,?\s*ltd\.?|co\.?,?\s*limited|corporation|corp\.?|บริษัท|จำกัด|holding|group|inc\.?|llc)\b/i;let r="",s="",c="";const m=[];for(const l of t)if(!/^[-_=*]{2,}$/.test(l)&&!/^(best regards|regards|sincerely|ขอแสดงความนับถือ|ด้วยความนับถือ|from|sent|thanks|thank you)/i.test(l)){if(!r){const f=l.match(n);if(f){r=f[0];continue}}if(!s){const f=l.match(a);if(f&&f[0].replace(/\D/g,"").length>=7){s=f[0].trim();continue}}if(!c&&o.test(l)){c=l;continue}l.length>=2&&l.length<=50&&!/\d{4,}/.test(l)&&m.push(l)}const x=m.find(l=>!n.test(l)&&!o.test(l))??"";return!r&&!x?null:{name:x,company:c,email:r,phone:s}}async function Lt(){const e=i.signatureContact;if(!e)return;const t=(i.emailSenderEmail||"").toLowerCase();if(t&&i.contactEmails.includes(t)){C("ลูกค้านี้มีในระบบแล้ว","success"),i.signatureContact=null,q();return}const n=document.getElementById("btn-import-customer");n&&(n.disabled=!0,n.textContent="กำลังบันทึก…");try{await V("HD_Contracts",{Title:i.emailSenderName||e.name,CustomerEmail:i.emailSenderEmail,Phone:e.phone||void 0,Company:e.company||void 0,Status:"Active"}),t&&i.contactEmails.push(t),C("เพิ่มลูกค้าสำเร็จ!"),i.signatureContact=null,q()}catch(a){const o=a instanceof Error?a.message:String(a);C(`เกิดข้อผิดพลาด: ${o}`,"error"),n&&(n.disabled=!1,n.textContent="เพิ่มเป็นลูกค้า")}}function pe(){return new Date().toISOString().split("T")[0]}function Rt(){const e=new Date;return`HD-${`${e.getFullYear()}${String(e.getMonth()+1).padStart(2,"0")}${String(e.getDate()).padStart(2,"0")}`}-${Math.floor(Math.random()*900+100)}`}function Se(){var e;return i.droppedFiles.length>0||document.querySelectorAll(".email-att-cb:checked").length>0||(((e=document.getElementById("f-attach-eml"))==null?void 0:e.checked)??!1)}async function X(e,t){i.droppedFiles.length>0&&await ue(e,t,i.droppedFiles),await de(e,t),await me(e,t)}let re=!1;async function Mt(){var t,n,a,o,r,s,c,m,x,l,f,j,O,N,E,R,I;if(!i.account){C("กรุณาเข้าสู่ระบบก่อน","error");return}if(re)return;re=!0;const e=document.getElementById("submit-btn");e&&(e.disabled=!0,e.textContent="กำลังบันทึก…");try{if(i.tab==="phish")await yt();else if(i.tab==="ticket"){const h=document.getElementById("f-title").value.trim(),g=document.getElementById("f-description").value.trim(),T=document.getElementById("f-priority").value,$=document.getElementById("f-customer-email").value.trim(),_=((t=document.getElementById("f-cc-enable"))==null?void 0:t.checked)??!0?(((n=document.getElementById("f-cc"))==null?void 0:n.value)||"").split(/[,;\s]+/).map(U=>U.trim()).filter(Boolean):[],S=document.getElementById("f-assigned-email").value,v=i.agents.find(U=>U.email===S),D=Rt(),u=await V("HD_Tickets",{Title:h,TicketNumber:D,Description:g,Priority:T,CustomerEmail:$,CustomerName:i.emailSenderName||$,Status:"Open",AssignedEmail:S||void 0,AssignedToName:(v==null?void 0:v.name)??((a=i.account)==null?void 0:a.name)??"",...Ee(S,(o=i.account)==null?void 0:o.username),ProjectID:parseInt(((r=document.getElementById("f-project"))==null?void 0:r.value)||"0")||null});if(Se()){const U=await V("HD_TicketComments",{Title:"📎 ไฟล์แนบจากอีเมล",TicketID:u,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await X("HD_TicketComments",U)}i.droppedFiles=[];const w={ticket_number:D,ticket_title:h,priority:T,category:"-",description:le(g||"-"),customer_name:i.emailSenderName||$,assigned_name:(v==null?void 0:v.name)??((s=i.account)==null?void 0:s.name)??"-",link:"https://itservices.co.th/helpdesk/"},B=[S,i.account.username,..._,Ve].filter(Boolean);let b=!1;const F=await Pt("ticket_created",w);if(F){const U=`<div style="border-left:4px solid #2563eb;background:#eff6ff;padding:10px 14px;margin:0 0 14px;font-family:Segoe UI,sans-serif">
             <div style="font-size:15px;font-weight:700;color:#1e40af">Ticket No. ${D}</div>
             <div style="font-size:12px;color:#475569;margin-top:2px">กรุณาตอบกลับในอีเมลฉบับนี้เพื่อให้ข้อมูลอยู่ใน Ticket เดียวกัน</div>
           </div>`;b=await jt(U+F,B)}let G=b;b||(G=await Ae("ticket_created",w,[$],B)),G?C(b?"สร้าง Ticket และตอบกลับในเธรดเดิมแล้ว!":"สร้าง Ticket สำเร็จ!"):C(`สร้าง Ticket ${D} แล้ว แต่ส่งเมลถึงลูกค้าไม่สำเร็จ — ตรวจ template ticket_created ที่หน้า Diagnostic`,"error")}else if(i.tab==="task"){const h=document.getElementById("f-title").value.trim(),g=parseInt(((c=document.getElementById("f-project"))==null?void 0:c.value)||"0"),T=document.getElementById("f-due-date").value,$=document.getElementById("f-note").value.trim(),P=document.getElementById("f-assigned-email").value,_=i.agents.find(D=>D.email===P);if(!g){C("กรุณาเลือก Project","error");return}const S=await V("PM_Tasks",{Title:h,DueDate:T||null,TaskNote:$,AssignedTo:(_==null?void 0:_.name)??i.account.name??i.account.username,AssignedEmail:P,IsCompleted:!1,IsAcknowledged:!1,ProjectID:g});if(i.droppedFiles.length>0&&await ue("PM_Tasks",S,i.droppedFiles),await de("PM_Tasks",S),await me("PM_Tasks",S),i.droppedFiles=[],await Q({recipients:[P],title:`📋 ได้รับมอบหมาย Task: ${h}`,message:$||(T?`กำหนดส่ง ${T}`:"มี Task ใหม่"),linkPath:g?`/projects/${g}`:"/my-work",eventType:"task_assigned"}),((m=document.getElementById("f-teams"))==null?void 0:m.checked)&&T){const D=Array.from(document.querySelectorAll(".att-internal:checked")).map(b=>b.value),u=(((x=document.getElementById("f-ext-att"))==null?void 0:x.value)||"").split(/[,;\s]+/).map(b=>b.trim()).filter(Boolean),w=`${T}T09:00:00`,B=`${T}T10:00:00`;try{await At({subject:h,start:w,end:B,body:$,attendees:[...D,...u],isOnlineMeeting:!0}),C("สร้าง Task + นัดประชุม Teams สำเร็จ!")}catch(b){C("สร้าง Task แล้ว แต่สร้างนัดประชุมไม่สำเร็จ: "+(b instanceof Error?b.message:""),"error")}}else C("สร้าง Task สำเร็จ!")}else if(i.tab==="incident"){const h=document.getElementById("f-title").value.trim(),g=parseInt(((l=document.getElementById("f-project"))==null?void 0:l.value)||"0"),T=document.getElementById("f-description").value.trim(),$=document.getElementById("f-severity").value,P=document.getElementById("f-assigned-email").value,_=i.agents.find(B=>B.email===P),S=document.getElementById("f-status").value,v=document.getElementById("f-incident-date").value,D=document.getElementById("f-resolution").value.trim();if(!g){C("กรุณาเลือก Project","error");return}const u=parseInt(((f=document.getElementById("f-sla"))==null?void 0:f.value)||"0")||null,w=await V("PM_Incidents",{Title:h,Description:T||void 0,Severity:$,Status:S,AssignedTo:(_==null?void 0:_.name)??i.account.name??i.account.username,AssignedEmail:P,...Ee(P,(j=i.account)==null?void 0:j.username),ProjectID:g,IncidentDate:v||pe(),Resolution:D||void 0,SLAHours:u,SLADue:Re(u),...S==="Resolved"?{ResolvedDate:new Date().toISOString()}:{}});i.droppedFiles.length>0&&await ue("PM_Incidents",w,i.droppedFiles),await de("PM_Incidents",w),await me("PM_Incidents",w),i.droppedFiles=[],await Q({recipients:[P],title:`🚨 ได้รับมอบหมาย Incident: ${h}`,message:`ความรุนแรง ${$}${T?" — "+T.slice(0,120):""}`,linkPath:g?`/projects/${g}`:"/my-work",eventType:"incident_created"});{const B=i.projects.find(G=>G.id===g),b=rt({title:h,severity:$,status:S,description:T,incidentDate:v||pe(),slaHours:u,projectName:B==null?void 0:B.Title,projectId:g,assignedName:_==null?void 0:_.name,assignedEmail:P,requesterEmail:i.account.username,actorEmail:i.account.username,baseUrl:"https://itservices.co.th/helpdesk/"});(b.to.length===0?!0:await Ae("incident_created",b.vars,b.to,b.cc))?C("สร้าง Incident สำเร็จ!"):C("สร้าง Incident แล้ว แต่ส่งเมลแจ้งไม่สำเร็จ — ผู้รับผิดชอบยังไม่รู้เรื่อง (ตรวจ template incident_created)","error")}}else if(i.tab==="comment"){const h=parseInt(((O=document.getElementById("f-ticket"))==null?void 0:O.value)||"0"),g=document.getElementById("f-comment").value.trim(),T=document.getElementById("f-comment-type").value;if(!h){C("กรุณาเลือก Ticket","error");return}if(!g){C("กรุณาพิมพ์ Comment","error");return}const $=await V("HD_TicketComments",{Title:g.slice(0,100),TicketID:h,CommentText:g,CommentType:T,CommentDate:new Date().toISOString()});await X("HD_TicketComments",$),i.droppedFiles=[];try{const P=await H(),_=`${M}/_api/web/lists/getbytitle('HD_Tickets')/items(${h})?$select=TicketNumber,Title,AssignedEmail,Author/EMail&$expand=Author`,S=await fetch(_,{headers:{Authorization:`Bearer ${P}`,Accept:"application/json;odata=nometadata"}});if(S.ok){const v=await S.json(),D=i.account.username.toLowerCase(),u=[...new Set([v.AssignedEmail,(N=v.Author)==null?void 0:N.EMail].filter(Boolean))].filter(w=>w.toLowerCase()!==D);u.length&&await Q({recipients:u,title:`💬 ${((E=i.account)==null?void 0:E.name)??"มีคน"} คอมเมนต์ใน ${v.TicketNumber||"#"+h}`,message:g.slice(0,200),linkPath:`/tickets/${h}`,eventType:"comment_added"})}}catch{}C("เพิ่ม Comment สำเร็จ!")}else if(i.tab==="project"){const h=document.getElementById("f-title").value.trim(),g=document.getElementById("f-company").value.trim(),T=document.getElementById("f-group").value,$=document.getElementById("f-status").value,P=document.getElementById("f-start").value,_=document.getElementById("f-end").value,S=document.getElementById("f-description").value.trim();if(!h){C("กรุณาใส่ชื่อโครงการ","error");return}const v=await V("PM_Projects",{Title:h,Company:g||void 0,ProjectGroup:T,Progress:0,StartDate:P||void 0,EndDate:_||null,Status:$,CreatedByEmail:i.account.username,Comment:S||void 0});if(Se()){const D=await V("PM_Comments",{Title:"📎 ไฟล์แนบจากอีเมล",ProjectID:v,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await X("PM_Comments",D)}i.droppedFiles=[],C("สร้างโครงการสำเร็จ!")}else if(i.tab==="projcomment"){const h=parseInt(((R=document.getElementById("f-project"))==null?void 0:R.value)||"0"),g=document.getElementById("f-comment").value.trim(),T=document.getElementById("f-comment-type").value;if(!h){C("กรุณาเลือกโครงการ","error");return}if(!g){C("กรุณาพิมพ์ Comment","error");return}const $=await V("PM_Comments",{Title:g.slice(0,100),ProjectID:h,CommentText:g,CommentType:T,CommentDate:new Date().toISOString(),ParentID:0});await X("PM_Comments",$),i.droppedFiles=[];try{const P=await H(),_=`${M}/_api/web/lists/getbytitle('PM_Projects')/items(${h})?$select=Title,CreatedByEmail`,S=await fetch(_,{headers:{Authorization:`Bearer ${P}`,Accept:"application/json;odata=nometadata"}});if(S.ok){const v=await S.json(),D=i.account.username.toLowerCase();v.CreatedByEmail&&v.CreatedByEmail.toLowerCase()!==D&&await Q({recipients:[v.CreatedByEmail],title:`💬 ${((I=i.account)==null?void 0:I.name)??"มีคน"} คอมเมนต์ในโครงการ ${v.Title??""}`,message:g.slice(0,200),linkPath:`/projects/${h}?tab=comments`,eventType:"comment_added"})}}catch{}C("เพิ่ม Comment สำเร็จ!")}}catch(h){const g=h instanceof Error?h.message:String(h);C(`เกิดข้อผิดพลาด: ${g}`,"error")}finally{re=!1,e&&(e.disabled=!1,e.textContent="บันทึก")}}const Ht={ticket:{label:"Ticket",icon:"🎫"},comment:{label:"Comment",icon:"💬"},project:{label:"Project",icon:"📁"},task:{label:"Task",icon:"✅"},incident:{label:"Incident",icon:"🚨"},projcomment:{label:"Comment",icon:"💬"},phish:{label:"PhishGuard",icon:"🛡️"}},Nt=[{title:"🎫 Helpdesk",tabs:["ticket","comment"]},{title:"📁 Project",tabs:["project","task","incident","projcomment"]},{title:"🛡️ Security",tabs:["phish"]}],We=["f-title","f-description","f-priority","f-customer-email","f-cc","f-assigned-email","f-project","f-due-date","f-note","f-severity","f-status","f-sla","f-incident-date","f-resolution","f-ticket","f-comment","f-comment-type","f-company","f-group","f-start","f-end","f-ext-att","f-attach-eml"];let W={};function Ft(){for(const t of We){const n=document.getElementById(t);n&&(W[t]=n.value)}const e=document.getElementById("f-teams");e&&(W["f-teams"]=e.checked)}function zt(){for(const t of We){const n=document.getElementById(t);n&&W[t]!==void 0&&W[t]!==""&&(n.value=W[t])}const e=document.getElementById("f-teams");if(e&&W["f-teams"]!==void 0){e.checked=W["f-teams"];const t=document.getElementById("teams-fields");t&&(t.style.display=e.checked?"block":"none")}}function q(){var $,P,_,S,v,D;const e=document.getElementById("app");if(!e)return;Ft();const{account:t,tab:n,emailSubject:a,emailSenderName:o,emailSenderEmail:r,emailBodyPreview:s}=i,c=i.emailBodyReply||s,m=t!==null,x=`
    <div class="bg-gradient-to-r from-blue-700 to-blue-600 text-white px-4 py-2.5 flex items-center gap-2.5 shadow flex-shrink-0">
      <div class="w-7 h-7 bg-white/95 rounded-lg flex items-center justify-center flex-shrink-0">
        <span class="text-blue-700 font-extrabold text-xs">iT</span>
      </div>
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm leading-tight truncate">iT Services Helpdesk</div>
        ${m?`<div class="text-[10px] text-blue-100 truncate">${y((t==null?void 0:t.name)??(t==null?void 0:t.username)??"")}</div>`:""}
      </div>
      ${m?`<button id="btn-logout" title="ออกจากระบบ" class="p-1.5 rounded-lg hover:bg-white/15 transition flex-shrink-0">
             <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
           </button>`:`<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded-lg transition flex-shrink-0">
             เข้าสู่ระบบ
           </button>`}
    </div>
  `;if(!m){e.innerHTML=`
      ${x}
      <div class="flex flex-col items-center justify-center px-6 py-16 gap-4 text-center">
        <div class="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center text-3xl">🔐</div>
        <h2 class="text-base font-semibold text-slate-700">เข้าสู่ระบบด้วย Microsoft</h2>
        <p class="text-xs text-slate-500">เพื่อสร้าง Ticket / Task / Incident จาก Email นี้</p>
        <button id="btn-login-main"
          class="mt-2 w-full bg-blue-700 hover:bg-blue-600 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          เข้าสู่ระบบ
        </button>
      </div>
    `,($=document.getElementById("btn-login"))==null||$.addEventListener("click",ke),(P=document.getElementById("btn-login-main"))==null||P.addEventListener("click",ke);return}const l=a?`<div class="mx-3 mt-3 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5 text-xs text-slate-700 space-y-1">
         <div class="font-semibold text-blue-800 truncate" title="${y(a)}">📧 ${y(a)}</div>
         ${o?`<div class="text-slate-500">จาก: <span class="font-medium text-slate-700">${y(o)}</span></div>`:""}
         ${r&&r!==o?`<div class="text-slate-400 truncate">${y(r)}</div>`:""}
       </div>`:`<div class="mx-3 mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-700">
         ⚠️ ไม่พบข้อมูล Email (โหมดทดสอบ)
       </div>`,f=i.signatureContact,j=!!r&&i.contactEmails.includes(r.toLowerCase()),O=f?`<div class="mx-3 mt-3 bg-orange-50 border border-orange-200 rounded-lg px-3 py-2.5 text-xs text-slate-700">
         <div class="flex items-center justify-between mb-2">
           <span class="font-semibold text-orange-700">👤 ข้อมูลผู้ส่ง (จาก Signature)</span>
         </div>
         <div class="space-y-0.5 mb-2.5">
           ${o?`<div><span class="text-slate-400">ชื่อ:</span> <span class="font-medium">${y(o)}</span></div>`:""}
           ${f.company?`<div><span class="text-slate-400">บริษัท:</span> ${y(f.company)}</div>`:""}
           ${r?`<div><span class="text-slate-400">Email:</span> ${y(r)}</div>`:""}
           ${f.phone?`<div><span class="text-slate-400">โทร:</span> ${y(f.phone)}</div>`:""}
         </div>
         ${j?'<div class="w-full bg-green-100 text-green-700 text-xs font-semibold py-1.5 rounded-md text-center">✓ ลูกค้านี้มีในระบบแล้ว</div>':`<button id="btn-import-customer"
                class="w-full bg-orange-500 hover:bg-orange-400 text-white text-xs font-semibold py-1.5 rounded-md transition">
                + เพิ่มเป็นลูกค้า
              </button>`}
       </div>`:"",N=`
    <div class="mx-3 mt-3 space-y-2">
      ${Nt.map(u=>`
        <div>
          <div class="text-[10px] font-semibold text-slate-400 uppercase tracking-wide mb-1 px-0.5">${u.title}</div>
          <div class="grid grid-cols-4 gap-1">
            ${u.tabs.map(w=>{const B=Ht[w];return`<button data-tab="${w}"
                class="tab-btn flex flex-col items-center gap-1 py-2 rounded-lg transition ${n===w?"bg-blue-700 text-white shadow":"text-slate-500 hover:bg-slate-100"}">
                <span class="text-base leading-none">${B.icon}</span>
                <span class="text-[9px] font-medium leading-none">${B.label}</span>
              </button>`}).join("")}
          </div>
        </div>
      `).join("")}
    </div>
  `;let E="";n==="phish"?E=$t(!!t):n==="ticket"?E=`
      ${A("Title / หัวข้อ",`<input id="f-title" type="text"
        class="${k}"
        value="${y(a)}" />`)}
      ${A("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${k} resize-none">${y(s)}</textarea>`)}
      ${A("Priority",`<select id="f-priority" class="${k}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
      ${A("Customer Email",`<input id="f-customer-email" type="email"
        class="${k}"
        value="${y(r)}" />`)}
      ${A("CC — ให้ผู้ที่อยู่ในเมลนี้รับรู้",`
        <label class="flex items-center gap-2 text-xs text-slate-600 mb-1.5 cursor-pointer">
          <input id="f-cc-enable" type="checkbox" ${i.emailCc.length?"checked":""} /> แนบผู้รับในเมลนี้เป็น CC อัตโนมัติ
        </label>
        <input id="f-cc" type="text" class="${k}" value="${y(i.emailCc.join(", "))}" placeholder="someone@company.com, boss@company.com" />`)}
      ${A("Assign ให้ Agent",ce(t.username))}
      ${A("โครงการ (ไม่บังคับ)",ee(!0))}
      ${J()}
    `:n==="task"?E=`
      ${A("ชื่อ Task *",`<input id="f-title" type="text" required
        class="${k}" value="${y(a)}" />`)}
      ${A("Project *",ee())}
      ${A("Assign ให้",ce(t.username))}
      ${A("Due Date",`<input id="f-due-date" type="date" class="${k}" />`)}
      ${A("Task Note",`<textarea id="f-note" rows="4"
        class="${k} resize-y">${y(s)}</textarea>`)}
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
          <input id="f-ext-att" type="text" class="${k}" placeholder="someone@company.com, ..." />
        </div>
      </div>
      ${J()}
    `:n==="incident"?E=`
      ${A("ชื่อ Incident *",`<input id="f-title" type="text" required
        class="${k}" value="${y(a)}" />`)}
      ${A("Project *",ee())}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">ความรุนแรง</label>
          <select id="f-severity" class="${k}">
            <option value="Low">Low</option>
            <option value="Medium" selected>Medium</option>
            <option value="High">High</option>
            <option value="Critical">Critical</option>
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${k}">
            <option value="Open" selected>Open</option>
            <option value="In Progress">In Progress</option>
            <option value="Resolved">Resolved</option>
          </select>
        </div>
      </div>
      ${A("SLA — ต้องแก้ให้จบภายใน",`
        <select id="f-sla" class="${k}">
          <option value="">ไม่กำหนด SLA</option>
          ${ct.map(u=>`<option value="${u.hours}" ${u.hours===se.Medium?"selected":""}>${u.labelTh}</option>`).join("")}
        </select>
        <p id="f-sla-hint" class="text-[11px] text-slate-400 mt-1">นับจากตอนนี้ · ครบกำหนด ${we(se.Medium)}</p>`)}
      ${A("Assign ให้ Agent",ce(t.username))}
      ${A("วันที่เกิด Incident",`<input id="f-incident-date" type="date" class="${k}" value="${pe()}" />`)}
      ${A("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${k} resize-y">${y(s)}</textarea>`)}
      ${A("วิธีแก้ไข (ถ้ามี)",`<textarea id="f-resolution" rows="2"
        class="${k} resize-y" placeholder="อธิบายวิธีแก้ไขปัญหา..."></textarea>`)}
      ${J()}
    `:n==="comment"?E=`
      ${A("เลือก Ticket *",`<select id="f-ticket" class="${k}">
        <option value="">-- เลือก Ticket ที่จะเพิ่ม Comment --</option>
        ${i.tickets.map(u=>`<option value="${u.id}">${y(u.TicketNumber||"#"+u.id)} · ${y(u.Title)}</option>`).join("")}
      </select>`)}
      ${A("ประเภท",`<select id="f-comment-type" class="${k}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${A("Comment *",`<textarea id="f-comment" rows="5"
        class="${k} resize-y" placeholder="พิมพ์ comment...">${y(c)}</textarea>`)}
      ${J()}
    `:n==="project"?E=`
      ${A("ชื่อโครงการ *",`<input id="f-title" type="text" required
        class="${k}" value="${y(a)}" />`)}
      ${A("บริษัท / ลูกค้า",`<input id="f-company" type="text" class="${k}" value="${y(((_=i.signatureContact)==null?void 0:_.company)??"")}" />`)}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">กลุ่มโครงการ</label>
          <select id="f-group" class="${k}">
            ${["Internal","External","R&D","Maintenance","อื่นๆ"].map(u=>`<option>${u}</option>`).join("")}
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${k}">
            ${["Planning","Active","On Hold","Completed","Cancelled"].map(u=>`<option>${u}</option>`).join("")}
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันที่เริ่ม</label>
          <input id="f-start" type="date" class="${k}" /></div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันสิ้นสุด</label>
          <input id="f-end" type="date" class="${k}" /></div>
      </div>
      ${A("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${k} resize-y">${y(s)}</textarea>`)}
      ${J()}
    `:n==="projcomment"&&(E=`
      ${A("เลือกโครงการ *",ee())}
      ${A("ประเภท",`<select id="f-comment-type" class="${k}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${A("Comment *",`<textarea id="f-comment" rows="5"
        class="${k} resize-y" placeholder="พิมพ์ comment...">${y(c)}</textarea>`)}
      ${J()}
    `);const R=n==="phish"?xt():n==="comment"||n==="projcomment"?"เพิ่ม Comment":n==="project"?"สร้างโครงการ":n==="incident"?"แจ้ง Incident":n==="task"?"สร้าง Task":"สร้าง Ticket";e.innerHTML=`
    <div class="flex flex-col h-screen bg-slate-50">
      ${x}
      <div class="flex-1 overflow-y-auto">
        ${l}
        ${O}
        ${N}
        <div class="mx-3 mt-3 space-y-3 pb-3">
          ${E}
        </div>
      </div>
      <div class="border-t border-slate-200 bg-white px-3 py-2.5 flex-shrink-0">
        <button id="submit-btn"
          class="w-full bg-blue-700 hover:bg-blue-600 active:bg-blue-800 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          ${R}
        </button>
      </div>
    </div>
  `,(S=document.getElementById("btn-logout"))==null||S.addEventListener("click",St),(v=document.getElementById("submit-btn"))==null||v.addEventListener("click",Mt),(D=document.getElementById("btn-import-customer"))==null||D.addEventListener("click",Lt),n==="phish"&&Tt();const I=document.getElementById("f-severity"),h=document.getElementById("f-sla");if(I&&h){let u=!1;const w=document.getElementById("f-sla-hint"),B=()=>{if(!w)return;const b=parseInt(h.value||"0")||null;w.textContent=b?`นับจากตอนนี้ · ครบกำหนด ${we(b)}`:"ไม่กำหนด SLA — เคสนี้จะวัดไม่ได้ในรายงาน"};h.addEventListener("change",()=>{u=!0,B()}),I.addEventListener("change",()=>{if(u)return;const b=se[I.value];b&&(h.value=String(b),B())})}document.querySelectorAll(".tab-btn").forEach(u=>{u.addEventListener("click",()=>{const w=u.dataset.tab;w&&w!==i.tab&&(i.tab=w,q(),w==="phish"&&ne())})});const g=document.getElementById("drop-zone"),T=document.getElementById("f-files");g&&T&&(T.addEventListener("change",()=>{T.files&&fe(Array.from(T.files)),T.value=""}),g.addEventListener("dragover",u=>{u.preventDefault(),g.classList.add("border-blue-500","bg-blue-50")}),g.addEventListener("dragleave",()=>{g.classList.remove("border-blue-500","bg-blue-50")}),g.addEventListener("drop",u=>{var B;u.preventDefault(),g.classList.remove("border-blue-500","bg-blue-50");const w=Array.from(((B=u.dataTransfer)==null?void 0:B.files)??[]);w.length&&fe(w)})),document.querySelectorAll(".remove-dropped").forEach(u=>{u.addEventListener("click",()=>{const w=parseInt(u.dataset.remove??"-1");w>=0&&(i.droppedFiles.splice(w,1),q())})}),zt()}function fe(e){i.droppedFiles.push(...e),q()}document.addEventListener("paste",e=>{var a;if(!i.account)return;const t=Array.from(((a=e.clipboardData)==null?void 0:a.items)??[]),n=[];for(const o of t)if(o.kind==="file"){const r=o.getAsFile();if(r){const s=r.name&&r.name!=="image.png"?r.name:`screenshot-${new Date().toISOString().replace(/[:.]/g,"-").slice(0,19)}.png`;n.push(new File([r],s,{type:r.type}))}}n.length&&(e.preventDefault(),fe(n),C(`แนบไฟล์แล้ว: ${n.map(o=>o.name).join(", ")}`))});const k="w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white";function Ce(e){return e<1024?`${e} B`:e<1024*1024?`${(e/1024).toFixed(0)} KB`:`${(e/1024/1024).toFixed(1)} MB`}function J(){const e=i.emailAttachments,t=i.droppedFiles,n=c=>`
    <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
      <input type="checkbox" class="email-att-cb" data-att-id="${y(c.id)}" data-att-name="${y(c.name)}" data-att-item="${c.isItem?"1":"0"}" ${c.defaultOn?"checked":""} />
      <span class="flex-1 truncate">${c.isItem?"📧 ":c.isInline?"🖼️ ":""}${y(c.name)}</span>
      <span class="text-slate-400 flex-shrink-0">${Ce(c.size)}</span>
    </label>`,a=e.filter(c=>!c.isInline),o=e.filter(c=>c.isInline),r=e.length>0?`<div class="mb-2 space-y-1">
        ${a.length?`<p class="text-xs text-slate-500">📎 ไฟล์แนบจาก Email:</p>
        ${a.map(n).join("")}`:""}
        ${o.length?`<p class="text-xs text-slate-500 ${a.length?"pt-1":""}">🖼️ รูปในเนื้อเมล
          <span class="text-slate-400">(รูปเล็กมักเป็นโลโก้ในลายเซ็น — ติ๊กเพิ่มได้)</span></p>
        ${o.map(n).join("")}`:""}
      </div>`:"",s=t.length>0?`<div class="mt-2 space-y-1">
        ${t.map((c,m)=>`<div class="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 rounded px-2 py-1">
            <span class="text-base">${c.type.startsWith("image/")?"🖼️":"📄"}</span>
            <span class="flex-1 truncate">${y(c.name)}</span>
            <span class="text-slate-400">${Ce(c.size)}</span>
            <button type="button" data-remove="${m}"
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
  </div>`}function ce(e){var t;return`<select id="f-assigned-email" class="${k}">
    <option value="${y(e)}">${y(((t=i.account)==null?void 0:t.name)??e)} (ฉัน)</option>
    ${i.agents.filter(n=>n.email!==e).map(n=>`<option value="${y(n.email)}">${y(n.name)}</option>`).join("")}
  </select>`}function ee(e=!1){return i.projects.length===0?e?'<div class="text-xs text-slate-400">ไม่พบ Project ที่ Active</div>':'<div class="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-2">⚠️ ไม่พบ Project ที่ Active</div>':`<select id="f-project" class="${k}">
    <option value="">${e?"-- ไม่ผูกกับโครงการ --":"-- เลือก Project --"}</option>
    ${i.projects.map(t=>`<option value="${t.id}">${y(t.Title)}</option>`).join("")}
  </select>`}function A(e,t){return`
    <div class="space-y-1">
      <label class="block text-xs font-medium text-slate-600">${e}</label>
      ${t}
    </div>
  `}function y(e){return e.replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}async function qt(){lt({sharepointUrl:M,internalDomains:["itservices.co.th","rpaexpert.com","rpaexpert.onmicrosoft.com"],getToken:H,getGraphToken:()=>K(),account:()=>i.account?{name:i.account.name,username:i.account.username}:null,toast:(t,n)=>C(t,n??"success"),rerender:q,canWhitelist:()=>["Agent","Supervisor","Boss","Admin"].includes(i.myRole)}),await L.initialize(),await L.handleRedirectPromise();const e=L.getAllAccounts();if(e.length>0){i.account=e[0];try{await L.acquireTokenSilent({scopes:[te],account:e[0]}),await Promise.all([Fe(),ze(),qe(),Ue()])}catch{i.account=null}}typeof Office<"u"?Office.onReady(t=>{var n;if(t.host===Office.HostType.Outlook){const a=Office.context.mailbox.item;if(a){i.emailSubject=a.subject??"";const o=a.from;o&&(i.emailSenderName=o.displayName??"",i.emailSenderEmail=o.emailAddress??"");const r=(((n=i.account)==null?void 0:n.username)??"").toLowerCase(),s=((o==null?void 0:o.emailAddress)??"").toLowerCase(),c=[...a.to??[],...a.cc??[]].map(l=>l.emailAddress).filter(Boolean);i.emailCc=[...new Set(c.map(l=>l.toLowerCase()))].filter(l=>l!==r&&l!==s);const m=a.attachments??[],x=20*1024;i.emailAttachments=m.filter(l=>l.attachmentType===Office.MailboxEnums.AttachmentType.File||l.attachmentType===Office.MailboxEnums.AttachmentType.Item).map(l=>({id:l.id,name:l.attachmentType===Office.MailboxEnums.AttachmentType.Item?`${(l.name||"email").replace(/\.eml$/i,"")}.eml`:l.name,size:l.size,isItem:l.attachmentType===Office.MailboxEnums.AttachmentType.Item,isInline:!!l.isInline,defaultOn:!l.isInline||l.size>=x})),a.body.getAsync(Office.CoercionType.Html,{asyncContext:{}},l=>{if(l.status===Office.AsyncResultStatus.Succeeded){let f=function(v,D=!1){if(v.nodeType===3){const b=v.textContent??"";return D&&b.trim()===""?"":b}const u=v,w=(u.tagName??"").toLowerCase();if(N.includes(w))return"";if(w==="br")return" ";if(w==="tr"){const b=[];for(let F=0;F<u.childNodes.length;F++){const G=u.childNodes[F],U=(G.tagName??"").toLowerCase();(U==="td"||U==="th")&&b.push((G.textContent??"").replace(/\s+/g," ").trim())}return b.length?b.join("	")+`
`:""}if(R.includes(w)){let b="";for(let F=0;F<u.childNodes.length;F++)b+=f(u.childNodes[F],!0);return b}let B="";for(let b=0;b<u.childNodes.length;b++)B+=f(u.childNodes[b],!1);return E.includes(w)&&(B=`
`+B.trim()+`
`),B};const j=l.value,O=new DOMParser().parseFromString(j,"text/html"),N=["style","script","head","img","meta","link","noscript"],E=["p","div","li","h1","h2","h3","h4","h5","h6","blockquote"],R=["table","thead","tbody","tfoot"],g=f(O.body??O.documentElement).replace(/[ \t]{2,}/g," ").replace(/\n[ \t]+/g,`
`).replace(/\n{3,}/g,`

`).trim().split(`
`),T=[];let $="";for(const v of g)v.trim()===""?$&&(T.push($.trim()),$=""):v.includes("	")?($&&(T.push($.trim()),$=""),T.push(v)):$=$?$+" "+v.trim():v.trim();$&&T.push($.trim());const P=T.join(`
`),_=/\n([-_]{3,}|From:\s|Best regards|Regards,|ขอแสดงความนับถือ|Sent:\s)/i,S=P.search(_);i.signatureContact=S>80?Ot(P.slice(S).trim()):null,i.emailBodyPreview=P.trim().slice(0,2e4),i.emailBodyReply=et(i.emailBodyPreview)}q()});return}}Be(),q()}):(Be(),q())}function Be(){i.emailSubject="[DEV] Test Email Subject",i.emailSenderName="Test Sender",i.emailSenderEmail="test@example.com",i.emailBodyPreview="This is a placeholder email body for development mode.",i.emailBodyReply=i.emailBodyPreview}qt().catch(e=>{console.error("Init error:",e);const t=document.getElementById("app");t&&(t.innerHTML=`<div class="p-4 text-red-600 text-sm">เกิดข้อผิดพลาด: ${String(e)}</div>`)});
