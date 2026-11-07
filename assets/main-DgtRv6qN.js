import{P as Ke}from"./PublicClientApplication-DLKYUtZW.js";import{a as ae,d as Ye,p as Ze,L as Pe,S as _e}from"./analyzer-CWd3MChg.js";const Qe=/^\s*(Sent|To|Date|Cc|Subject|ส่ง|ถึง|วันที่|สำเนา|เรื่อง)\s*:/i;function Xe(e,t){const n=e[t];if(/^\s*-{2,}\s*(Original Message|Forwarded message|ข้อความต้นฉบับ)\s*-{2,}/i.test(n)||/^\s*_{5,}\s*$/.test(n)||/^\s*(On|เมื่อ)\b.{10,200}(wrote|เขียนว่า)\s*:\s*$/i.test(n))return!0;if(/^\s*(From|จาก)\s*:\s*\S/i.test(n)){for(let a=t+1;a<=t+3&&a<e.length;a++)if(Qe.test(e[a]))return!0;return!1}return!1}function et(e){const t=(e??"").replace(/\r\n/g,`
`),n=t.split(`
`);let a=-1;for(let s=0;s<n.length;s++){if(s>0&&Xe(n,s)){a=s;break}if(s>0&&/^\s*>/.test(n[s])&&/^\s*>/.test(n[s+1]??"")){a=s;break}}if(a<0)return{visible:t.trim(),quoted:""};const o=n.slice(0,a).join(`
`).trim(),r=n.slice(a).join(`
`).trim();return o?{visible:o,quoted:r}:{visible:t.trim(),quoted:""}}const tt=e=>et(e).visible,ve=e=>(typeof e=="string"?e:"").trim().toLowerCase();function De(e,t){const n=ve(t);return e.find(a=>ve(a.EventKey)===n&&nt(a.IsEnabled))}function nt(e){if(e==null)return!0;if(typeof e=="boolean")return e;const t=String(e).trim().toLowerCase();return!["false","no","0","ไม่","ปิด",""].includes(t)}const at=e=>({__html:e??""}),Oe=e=>!!e&&typeof e=="object"&&typeof e.__html=="string";function Le(e){return(e??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}const le=e=>at(Le(e??"").replace(/\r?\n/g,"<br>"));function Re(e,t){const n=[];return{text:(e??"").replace(/\{\{(\w+)\}\}/g,(o,r)=>{const s=t[r];return s===void 0?(n.push(r),""):Oe(s)?s.__html:Le(s)}),missing:[...new Set(n)]}}function st(e,t){return(e??"").replace(/\{\{(\w+)\}\}/g,(n,a)=>{const o=t[a];return o===void 0?"":(Oe(o)?o.__html.replace(/<[^>]*>/g,""):o).replace(/\s+/g," ").trim()})}function ot(e){const t=o=>({emailAddress:{address:o}}),n={subject:e.subject,body:{contentType:"HTML",content:e.html},toRecipients:e.to.filter(Boolean).map(t)},a=(e.cc??[]).filter(Boolean);return a.length&&(n.ccRecipients=a.map(t)),e.from&&(e.sendAs==="shared"?n.from=t(e.from):n.replyTo=[t(e.from)]),n}function it(e,t){return e===403&&/ErrorSendAsDenied/i.test(t)}const we=e=>(e??"").trim().toLowerCase();function rt(e){const t=we(e.actorEmail),n=new Set,a=s=>{const c=[];for(const m of s){const y=(m??"").trim(),l=we(y);!l||!l.includes("@")||l===t||n.has(l)||(n.add(l),c.push(y))}return c};let o=a([e.assignedEmail]);const r=a([e.requesterEmail,...e.watchers??[]]);return o.length===0?(o=r.slice(0,1),{to:o,cc:r.slice(1)}):{to:o,cc:r}}const ct=e=>{if(!e||e<=0)return"ไม่ได้กำหนด";if(e<24)return`${e} ชั่วโมง`;const t=e/24;return Number.isInteger(t)?`${t} วัน`:`${e} ชั่วโมง`};function lt(e){const t=(e.baseUrl??"").replace(/\/+$/,"");return{incident_title:e.title??"",severity:e.severity??"",status:e.status??"",incident_status:e.status??"",description:le(e.description),resolution:le(e.resolution),incident_date:(e.incidentDate??"").slice(0,10),sla_hours:ct(e.slaHours),project_name:e.projectName??"",assigned_name:e.assignedName||e.assignedEmail||"-",link:e.projectId?`${t}/#/projects/${e.projectId}`:t}}function dt(e){const{to:t,cc:n}=rt(e);return{to:t,cc:n,vars:lt(e)}}const mt=[{hours:1,labelTh:"1 ชั่วโมง"},{hours:2,labelTh:"2 ชั่วโมง"},{hours:4,labelTh:"4 ชั่วโมง"},{hours:8,labelTh:"8 ชั่วโมง (1 วันทำการ)"},{hours:24,labelTh:"24 ชั่วโมง"},{hours:48,labelTh:"2 วัน"},{hours:72,labelTh:"3 วัน"},{hours:168,labelTh:"7 วัน"}],se={Critical:1,High:4,Medium:24,Low:72};function Me(e,t,n=new Date){const a=typeof e=="number"&&Number.isFinite(e)&&e>0?e:null;if(!a)return null;const o=n,r=isNaN(o.getTime())?n:o;return new Date(r.getTime()+a*36e5).toISOString()}function $e(e){const t=Me(e);return t?new Date(t).toLocaleString("th-TH",{dateStyle:"short",timeStyle:"short"}):""}const He="HD_PhishingReports";let p;const d={mail:null,analysis:null,loading:!1,reporting:!1,reported:!1,showHeaders:!1,kasmTemplate:"",analysedItemId:"",safeDomains:[],safeDomainIds:{},savingDomain:""};function ut(e){p=e}const q=e=>(e??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"),Y=()=>{var e,t;return(t=(e=Office.context)==null?void 0:e.mailbox)==null?void 0:t.item},ke=e=>new Promise(t=>{const n=Y();if(!(n!=null&&n.body)){t("");return}n.body.getAsync(e,a=>t(a.status===Office.AsyncResultStatus.Succeeded?a.value??"":""))});function pt(e){return e?e.split(",").map(t=>{const n=t.trim(),a=n.match(/^(.*?)\s*<([^>]+)>$/);return a?{name:a[1].replace(/^"|"$/g,"").trim(),email:a[2].trim()}:{name:"",email:n.replace(/[<>]/g,"").trim()}}).filter(t=>t.email.includes("@")):[]}function ft(){return new Promise(e=>{let t=!1;try{t=Office.context.requirements.isSetSupported("Mailbox","1.8")}catch{t=!1}const n=Y();if(!t||typeof(n==null?void 0:n.getAllInternetHeadersAsync)!="function"){e({});return}try{n.getAllInternetHeadersAsync(a=>e(a.status===Office.AsyncResultStatus.Succeeded?Ze(a.value??""):{}))}catch{e({})}})}async function ht(){try{const e=Y();if(!(e!=null&&e.itemId))return{};const t=Office.context.mailbox.convertToRestId(e.itemId,Office.MailboxEnums.RestVersion.v2_0),n=await p.getGraphToken(),a=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${t}?$select=internetMessageHeaders`,{headers:{Authorization:`Bearer ${n}`}});if(!a.ok)return{};const o=await a.json(),r={};for(const s of o.internetMessageHeaders??[])r[s.name]=r[s.name]?`${r[s.name]}
${s.value}`:s.value;return r}catch{return{}}}async function he(e,t){const n=await p.getToken(),a=await fetch(`${p.sharepointUrl}/_api/web/lists/getbytitle('${e}')/items?${t}`,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});return a.ok?(await a.json()).value:[]}async function gt(){try{return(await he("HD_AgentProfiles","$select=Title,EmailText&$top=500")).filter(t=>t.EmailText).map(t=>({name:t.Title,email:t.EmailText}))}catch{return[]}}async function bt(){var e;try{return(((e=(await he("HD_Options","$select=Title&$filter=Category eq 'KasmConfig'&$top=1"))[0])==null?void 0:e.Title)??"").trim()}catch{return""}}const Ne="SafeDomain";async function ge(){try{const e=await he("HD_Options",`$select=Id,Title&$filter=Category eq '${Ne}'&$top=500`),t={},n=[];for(const a of e){const o=(a.Title??"").trim().toLowerCase();o&&(t[o]=a.Id,n.push(o))}d.safeDomains=n,d.safeDomainIds=t}catch{}}async function yt(e){const t=e.trim().toLowerCase();if(!(!t||d.savingDomain)){if(!p.canWhitelist()){p.toast("ต้องเป็น Agent ขึ้นไปจึงจะยืนยันโดเมนได้","error");return}d.savingDomain=t,p.rerender();try{const n=await p.getToken(),a=await fetch(`${p.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items`,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:t,Category:Ne})});if(!a.ok)throw new Error(String(a.status));await ge(),await ne(!0),p.toast(`ยืนยันแล้วว่า ${t} ปลอดภัย`)}catch{p.toast("บันทึกไม่สำเร็จ","error")}finally{d.savingDomain="",p.rerender()}}}async function xt(e){const t=e.trim().toLowerCase(),n=d.safeDomainIds[t];if(!(!n||d.savingDomain)){if(!p.canWhitelist()){p.toast("ต้องเป็น Agent ขึ้นไปจึงจะถอนโดเมนได้","error");return}d.savingDomain=t,p.rerender();try{const a=await p.getToken(),o=await fetch(`${p.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items(${n})`,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","IF-MATCH":"*","X-HTTP-Method":"DELETE"}});if(!o.ok)throw new Error(String(o.status));await ge(),await ne(!0),p.toast(`ถอน ${t} ออกจากรายการปลอดภัยแล้ว`)}catch{p.toast("ถอนไม่สำเร็จ","error")}finally{d.savingDomain="",p.rerender()}}}async function ne(e=!1){var m,y;const t=Y(),n=(t==null?void 0:t.itemId)??"";if(!e&&d.analysedItemId===n&&d.analysis)return;d.loading=!0,d.reported=!1,d.analysedItemId=n,p.rerender();const[a,o]=await Promise.all([ke(Office.CoercionType.Html),ke(Office.CoercionType.Text)]),r={fromName:((m=t==null?void 0:t.from)==null?void 0:m.displayName)??"",fromEmail:((y=t==null?void 0:t.from)==null?void 0:y.emailAddress)??"",replyTo:[],subject:(t==null?void 0:t.subject)??"",bodyHtml:a,bodyText:o,attachments:((t==null?void 0:t.attachments)??[]).map(l=>({name:l.name,size:l.size??0,isInline:!!l.isInline})),headers:{},internalDomains:p.internalDomains,internalPeople:[],safeDomains:d.safeDomains},s=l=>{var g;return{...r,headers:l,replyTo:pt(((g=Object.entries(l).find(([_])=>_.toLowerCase()==="reply-to"))==null?void 0:g[1])??"")}},c=await ft();if(d.mail=s(c),d.analysis=ae(d.mail),d.loading=!1,p.rerender(),p.account()){const[l,g]=await Promise.all([gt(),Object.keys(c).length?Promise.resolve(c):ht()]);d.mail={...s(g),internalPeople:l},d.analysis=ae(d.mail),d.kasmTemplate||(d.kasmTemplate=await bt()),d.safeDomains.length||(await ge(),d.safeDomains.length&&(d.mail={...d.mail,safeDomains:d.safeDomains},d.analysis=ae(d.mail))),p.rerender()}}function Fe(){const e=d.mail,t=d.analysis;return!e||!t?"":[`ผู้ส่ง: ${e.fromName} <${e.fromEmail}>`,`หัวข้อ: ${e.subject}`,e.replyTo.length?`Reply-To: ${e.replyTo.map(n=>n.email).join(", ")}`:"",`คะแนนความเสี่ยง: ${t.score} (${Pe[t.level].label})`,"","สิ่งที่ตรวจพบ:",...t.findings.map(n=>`- [${_e[n.severity].label}] (${n.category}) ${n.title} — ${n.detail.replace(/\n/g," ")}`),"",t.links.length?"ลิงก์ในอีเมล:":"",...t.links.map(n=>`- ${n.href}${n.flags.length?`  ! ${n.flags.join(" / ")}`:""}`)].filter(n=>n!=="").join(`
`)}async function vt(e){try{const t=Y();if(!(t!=null&&t.itemId))return!1;const n=Office.context.mailbox.convertToRestId(t.itemId,Office.MailboxEnums.RestVersion.v2_0),a=await p.getGraphToken(),o=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${n}/$value`,{headers:{Authorization:`Bearer ${a}`}});if(!o.ok)return!1;const r=await o.arrayBuffer(),s=(t.subject||"phishing").replace(/[\\/:*?"<>|#%&{}~]/g,"_").replace(/^_+/,"").slice(0,80).trim()||"phishing",c=await p.getToken();return(await fetch(`${p.sharepointUrl}/_api/web/lists/getbytitle('${He}')/items(${e})/AttachmentFiles/add(FileName='${encodeURIComponent(s+".eml")}')`,{method:"POST",headers:{Authorization:`Bearer ${c}`,Accept:"application/json;odata=nometadata"},body:r})).ok}catch{return!1}}async function Te(e){const t=await p.getToken(),n=await fetch(`${p.sharepointUrl}/_api/web/lists/getbytitle('${He}')/items`,{method:"POST",headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(e)});if(!n.ok)throw new Error(`SharePoint ${n.status}: ${await n.text()}`);return(await n.json()).Id}async function wt(){if(!(!d.mail||!d.analysis||d.reporting)){if(!p.account()){p.toast("กรุณาเข้าสู่ระบบก่อนรายงาน","error");return}d.reporting=!0,p.rerender();try{const e=d.mail,t=d.analysis,n=p.account(),a={Title:(e.subject||"(ไม่มีหัวข้อ)").slice(0,255),SenderName:e.fromName.slice(0,255),SenderEmail:e.fromEmail.slice(0,255),SenderDomain:Ye(e.fromEmail),RiskScore:t.score,RiskLevel:t.level,Findings:Fe(),LinkCount:t.links.length,SuspiciousLinks:t.links.filter(c=>c.flags.length).map(c=>c.href).join(`
`).slice(0,4e3),ReportedBy:(n==null?void 0:n.name)??"",ReportedEmail:(n==null?void 0:n.username)??"",Status:"New"};let o,r=!1;try{o=await Te(a)}catch(c){o=await Te({Title:a.Title,Findings:a.Findings}).catch(()=>{throw c}),r=!0}const s=await vt(o);d.reported=!0,p.toast(r?"ส่งรายงานแล้ว แต่บันทึกได้บางคอลัมน์ — ตรวจชื่อคอลัมน์ใน HD_PhishingReports":s?"ส่งรายงานพร้อมอีเมลต้นฉบับแล้ว":"ส่งรายงานแล้ว (แนบ .eml ไม่ได้)",r?"info":"success")}catch(e){p.toast(`ส่งรายงานไม่สำเร็จ: ${e instanceof Error?e.message:String(e)}`,"error")}finally{d.reporting=!1,p.rerender()}}}const $t=()=>d.reported?"✓ รายงานแล้ว":d.reporting?"กำลังส่ง…":"🚩 รายงานอีเมลนี้ให้ IT";function kt(e){var t;try{const n=(t=Office.context)==null?void 0:t.ui;if(typeof(n==null?void 0:n.openBrowserWindow)=="function"){n.openBrowserWindow(e);return}}catch{}window.open(e,"_blank","noopener,noreferrer")||p.toast("เปิดหน้าต่างไม่ได้ (ถูกบล็อก)","info")}async function Tt(e){var t;try{if((t=navigator.clipboard)!=null&&t.writeText)return await navigator.clipboard.writeText(e),!0}catch{}try{const n=document.createElement("textarea");n.value=e,n.setAttribute("readonly",""),n.style.position="fixed",n.style.opacity="0",document.body.appendChild(n),n.focus(),n.select();const a=document.execCommand("copy");return n.remove(),a}catch{return!1}}function Et(e){if(d.loading&&!d.analysis)return`<div class="py-10 text-center text-slate-500 text-sm">
      <div class="w-7 h-7 border-[3px] border-slate-200 border-t-blue-600 rounded-full animate-spin mx-auto mb-2"></div>
      กำลังตรวจอีเมล…</div>`;const t=d.analysis,n=d.mail;if(!t||!n)return'<p class="text-sm text-slate-400 text-center py-8">เปิดอีเมลเพื่อเริ่มตรวจ</p>';const a=Pe[t.level],o=[...t.links.filter(s=>s.flags.length),...t.links.filter(s=>!s.flags.length&&!s.trusted),...t.links.filter(s=>s.trusted)],r=p.canWhitelist();return`
    <div class="rounded-xl border-2 ${a.cls} p-3">
      <div class="flex items-center gap-2">
        <span class="text-2xl leading-none">${a.icon}</span>
        <div class="min-w-0 flex-1">
          <div class="font-bold text-sm">${q(a.label)}</div>
          <div class="text-xs opacity-80">คะแนน ${t.score} · พบสัญญาณ ${t.findings.filter(s=>s.severity!=="info").length} ข้อ</div>
        </div>
        <button id="phish-recheck" class="text-[10px] px-2 py-1 rounded-md bg-white/70 hover:bg-white text-slate-700 font-medium">ตรวจใหม่</button>
      </div>
    </div>

    ${e?"":`<div class="bg-blue-50 border border-blue-200 rounded-xl p-2.5 text-[11px] text-blue-800">
      เข้าสู่ระบบเพื่อตรวจการปลอมเป็นพนักงาน และรายงานเข้า Helpdesk ได้
    </div>`}

    ${t.findings.length===0?'<p class="text-xs text-slate-400 text-center py-3">ไม่พบสัญญาณผิดปกติจากการตรวจอัตโนมัติ</p>':t.findings.map(s=>{const c=_e[s.severity];return`<div class="bg-white rounded-xl border border-slate-200 p-2.5">
          <div class="flex items-start gap-2">
            <span class="text-[9px] font-bold px-1.5 py-0.5 rounded-full ${c.cls} flex-shrink-0 mt-0.5">${c.label}</span>
            <div class="min-w-0 flex-1">
              <div class="text-xs font-semibold text-slate-800">${q(s.title)}</div>
              <div class="text-[11px] text-slate-500 whitespace-pre-line break-all">${q(s.detail)}</div>
              <div class="text-[9px] text-slate-400 mt-0.5">${q(s.category)}</div>
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
              ${s.trusted?"✔ ":""}${q(s.host||s.href)}
            </div>
            ${s.trusted?`<div class="text-[10px] text-emerald-700">ทีมตรวจแล้วว่าปลอดภัย${(s.suppressed??[]).length?` · ระงับการเตือน ${s.suppressed.length} ข้อ`:""}</div>`:""}
            ${s.text&&s.text!==s.href?`<div class="text-[10px] text-slate-500 break-all">แสดงว่า: "${q(s.text.slice(0,70))}"</div>`:""}
            <div class="text-[10px] text-slate-400 break-all mt-0.5">${q(s.href.slice(0,150))}</div>
            ${s.flags.map(l=>`<div class="text-[10px] text-red-600 mt-0.5">! ${q(l)}</div>`).join("")}
            <div class="flex flex-wrap gap-1 mt-1.5">
              <button data-kasm="${c}" class="text-[10px] font-semibold px-2 py-1 rounded-md bg-slate-800 hover:bg-slate-900 text-white">
                เปิดใน Kasm
              </button>
              ${r&&s.host?s.trusted?`<button data-untrust="${q(s.host)}" ${m?"disabled":""}
                     class="text-[10px] font-semibold px-2 py-1 rounded-md border border-slate-300 text-slate-600 hover:bg-slate-100 disabled:opacity-50">
                     ${m?"...":"ถอนออกจากรายการปลอดภัย"}</button>`:`<button data-trust="${q(s.host)}" ${m?"disabled":""}
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
    ${d.showHeaders?`<pre class="bg-slate-900 text-slate-100 text-[9px] p-2 rounded-lg overflow-x-auto whitespace-pre-wrap break-all max-h-56 overflow-y-auto">${q(Object.entries(n.headers).map(([s,c])=>`${s}: ${c}`).join(`
`)||"อ่าน header ไม่ได้")}</pre>`:""}
    <p class="text-[10px] text-slate-400 text-center">วิเคราะห์ในเครื่องทั้งหมด — ไม่ส่งเนื้อหาอีเมลออกนอกองค์กร</p>
  `}function It(){var n,a,o,r;(n=document.getElementById("phish-recheck"))==null||n.addEventListener("click",()=>{ne(!0)}),(a=document.getElementById("phish-headers"))==null||a.addEventListener("click",()=>{d.showHeaders=!d.showHeaders,p.rerender()}),(o=document.getElementById("phish-copy"))==null||o.addEventListener("click",async()=>{const s=await Tt(Fe());p.toast(s?"คัดลอกผลตรวจแล้ว":"คัดลอกไม่ได้",s?"success":"error")}),document.querySelectorAll("[data-trust]").forEach(s=>s.addEventListener("click",()=>yt(s.dataset.trust??""))),document.querySelectorAll("[data-untrust]").forEach(s=>s.addEventListener("click",()=>xt(s.dataset.untrust??"")));const e=((r=d.analysis)==null?void 0:r.links)??[],t=[...e.filter(s=>s.flags.length),...e.filter(s=>!s.flags.length&&!s.trusted),...e.filter(s=>s.trusted)];document.querySelectorAll("[data-kasm]").forEach(s=>{s.addEventListener("click",()=>{const c=t[Number(s.dataset.kasm)];if(!c)return;if(!d.kasmTemplate){p.toast("ยังไม่ได้ตั้งค่า Kasm ใน HD_Options (Category=KasmConfig)","info");return}const m=d.kasmTemplate;kt(m.includes("{url}")?m.replace("{url}",encodeURIComponent(c.href)):m+encodeURIComponent(c.href))})})}const At="0bab07cf-65e6-487c-89af-c917fc1a5a13",St="d569b991-89fc-4a62-9df5-eb361abcef40",H="https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd",te="https://rpaexpert.sharepoint.com/.default",oe=["https://graph.microsoft.com/Calendars.ReadWrite","https://graph.microsoft.com/Mail.Send","https://graph.microsoft.com/Mail.Read"],M=new Ke({auth:{clientId:At,authority:`https://login.microsoftonline.com/${St}`,redirectUri:window.location.origin.includes("localhost")?"http://localhost:3000/":"https://darmmunginsa.github.io/itservices-addin/",navigateToLoginRequestUrl:!1},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:!1}}),Ct=window.location.origin.includes("localhost")?`${window.location.origin}/auth.html`:"https://darmmunginsa.github.io/itservices-addin/auth.html";function be(){var t,n;const e=(n=(t=Office.context)==null?void 0:t.diagnostics)==null?void 0:n.platform;return e===Office.PlatformType.iOS||e===Office.PlatformType.Android}function ye(){return new Promise((e,t)=>{Office.context.ui.displayDialogAsync(Ct,{height:60,width:30,promptBeforeOpen:!1},n=>{if(n.status!==Office.AsyncResultStatus.Succeeded){t(new Error("เปิดหน้าเข้าสู่ระบบไม่ได้"));return}const a=n.value;a.addEventHandler(Office.EventType.DialogMessageReceived,o=>{a.close();const r=o.message;if(!r){t(new Error("auth message error"));return}try{const s=JSON.parse(r);s.ok?e():t(new Error(s.error||"auth failed"))}catch{t(new Error("auth message error"))}}),a.addEventHandler(Office.EventType.DialogEventReceived,()=>t(new Error("ปิดหน้าเข้าสู่ระบบก่อนเสร็จ")))})})}const i={account:null,tab:"ticket",emailSubject:"",emailBodyPreview:"",emailBodyReply:"",emailSenderName:"",emailSenderEmail:"",loading:!1,projects:[],agents:[],myRole:"",emailAttachments:[],signatureContact:null,droppedFiles:[],tickets:[],contactEmails:[],emailCc:[]};async function N(){const e=M.getAllAccounts();if(e.length===0)throw new Error("Not signed in");const t={scopes:[te],account:e[0]};try{return(await M.acquireTokenSilent(t)).accessToken}catch{if(be()){await ye();const n=M.getAllAccounts()[0];if(!n)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await M.acquireTokenSilent({scopes:[te],account:n})).accessToken}return(await M.acquireTokenPopup(t)).accessToken}}async function K(e=!1){const t=M.getAllAccounts();if(t.length===0)throw new Error("Not signed in");const n={scopes:oe,account:t[0],forceRefresh:e};try{return(await M.acquireTokenSilent(n)).accessToken}catch{if(be()){await ye();const o=M.getAllAccounts()[0];if(!o)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await M.acquireTokenSilent({scopes:oe,account:o})).accessToken}return(await M.acquireTokenPopup({scopes:oe,account:t[0]})).accessToken}}async function Bt(e){const t=await K(),n={subject:e.subject,start:{dateTime:e.start,timeZone:"Asia/Bangkok"},end:{dateTime:e.end,timeZone:"Asia/Bangkok"},body:e.body?{contentType:"HTML",content:e.body.replace(/\n/g,"<br>")}:void 0,attendees:e.attendees.filter(Boolean).map(o=>({emailAddress:{address:o},type:"required"})),isOnlineMeeting:e.isOnlineMeeting,onlineMeetingProvider:e.isOnlineMeeting?"teamsForBusiness":void 0},a=await fetch("https://graph.microsoft.com/v1.0/me/events",{method:"POST",headers:{Authorization:`Bearer ${t}`,"Content-Type":"application/json"},body:JSON.stringify(n)});if(!a.ok)throw new Error(`Calendar error ${a.status}: ${await a.text()}`)}async function ze(){try{const e=await N(),t=`${H}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$orderby=Title asc&$top=500`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const a=await n.json();i.projects=a.value.map(o=>({id:o.Id,Title:o.Title}))}}catch{}}async function qe(){var e,t;try{const n=await N(),a=`${H}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText,Role&$orderby=Title asc`,o=await fetch(a,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});if(o.ok){const r=await o.json();i.agents=r.value.map(c=>({email:c.EmailText,name:c.Title}));const s=(((e=i.account)==null?void 0:e.username)??"").toLowerCase();i.myRole=((t=r.value.find(c=>(c.EmailText??"").toLowerCase()===s))==null?void 0:t.Role)??""}}catch{}}async function Ue(){try{const e=await N(),t=`${H}/_api/web/lists/getbytitle('HD_Tickets')/items?$select=Id,Title,TicketNumber,Status&$filter=Status ne 'Closed'&$orderby=Modified desc&$top=200`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const a=await n.json();i.tickets=a.value.map(o=>({id:o.Id,Title:o.Title,TicketNumber:o.TicketNumber,Status:o.Status}))}}catch{}}async function Ve(){try{const e=await N(),t=`${H}/_api/web/lists/getbytitle('HD_Contracts')/items?$select=CustomerEmail&$top=2000`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const a=await n.json();i.contactEmails=a.value.map(o=>(o.CustomerEmail||"").trim().toLowerCase()).filter(Boolean)}}catch{}}async function Ee(){const e=document.getElementById("btn-login-main"),t=document.getElementById("btn-login");e&&(e.disabled=!0,e.textContent="กำลังเข้าสู่ระบบ…"),t&&(t.disabled=!0);try{if(be()){if(await ye(),i.account=M.getAllAccounts()[0]??null,!i.account)throw new Error("เข้าสู่ระบบไม่สำเร็จ")}else{const n=await M.loginPopup({scopes:[te]});i.account=n.account}await Promise.all([ze(),qe(),Ue(),Ve()]),U()}catch{e&&(e.disabled=!1,e.textContent="เข้าสู่ระบบ"),t&&(t.disabled=!1)}}async function jt(){i.account&&await M.logoutPopup({account:i.account}),i.account=null,U()}function Ie(e,t){const n=(e??"").trim().toLowerCase();return n?{IsAcknowledged:n===(t??"").trim().toLowerCase()}:{IsAcknowledged:!1}}async function V(e,t){const n=await N(),a=`${H}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items`,o=await fetch(a,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(t)});if(!o.ok){const s=await o.text(),c=["IsAcknowledged","AcknowledgedBy","AcknowledgedDate"];if(c.some(m=>m in t)){const m={...t};for(const l of c)delete m[l];const y=await fetch(a,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(m)});if(y.ok)return console.warn(`[SP] ${e} ไม่มีคอลัมน์ IsAcknowledged — สร้างแล้วแต่ไม่ได้เข้ากล่องรอรับงาน`),(await y.json()).Id}throw new Error(`SharePoint error ${o.status}: ${s}`)}return(await o.json()).Id}let Z=null,Ae=0;const Pt=5*60*1e3,ie="support@itservices.co.th",Ge="engineer@itservices.co.th";async function We(){if(Z&&Date.now()-Ae<Pt)return Z;try{const e=await N(),t=`${H}/_api/web/lists/getbytitle('HD_EmailTemplates')/items?$select=EventKey,Subject,Body,IsEnabled&$top=50`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});return n.ok?(Z=(await n.json()).value,Ae=Date.now(),Z):[]}catch{return[]}}async function _t(){var e,t;try{const n=await N(),a=`${H}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title,Category&$filter=Category eq 'EmailConfig'&$top=1`,o=await fetch(a,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});return o.ok&&((t=(e=(await o.json()).value[0])==null?void 0:e.Title)==null?void 0:t.trim())||ie}catch{return ie}}async function Se(e,t,n,a=[]){try{const o=await We(),r=De(o,e);if(!r)return console.warn(`[mail] ไม่พบ template "${e}" ที่เปิดใช้`),!1;const s=st(r.Subject||"",t),c=Re(r.Body||"",t).text;if(!s||!c)return console.warn(`[mail] template "${e}" Subject/Body ว่าง`),!1;const m=I=>I.trim().toLowerCase(),y=[...new Map(n.filter(Boolean).map(I=>[m(I),I])).values()];if(y.length===0)return!1;const l=new Set(y.map(m)),g=e==="ticket_created"?[...a,Ge]:a,_=[...new Map(g.filter(Boolean).map(I=>[m(I),I])).values()].filter(I=>!l.has(m(I))),D=await _t(),z=await K(),L=I=>fetch("https://graph.microsoft.com/v1.0/me/sendMail",{method:"POST",headers:{Authorization:`Bearer ${z}`,"Content-Type":"application/json"},body:JSON.stringify({message:ot({to:y,cc:_,subject:s,html:c,from:D||void 0,sendAs:I}),saveToSentItems:!0})});let T=await L(D?"shared":"self");if(!T.ok&&D){const I=await T.text().catch(()=>"");if(!it(T.status,I))return console.warn(`[mail] sendMail ${T.status}`,I),!1;console.warn(`[mail] ไม่มีสิทธิ์ Send As ของ ${D} — ส่งจากบัญชีตัวเองแทน`),T=await L("self")}return T.ok?!0:(console.warn(`[mail] sendMail ${T.status}`,await T.text().catch(()=>"")),!1)}catch(o){return console.warn("[mail] ส่งไม่สำเร็จ",o),!1}}async function Dt(e,t=[],n=!0){try{const a=Office.context.mailbox.item;if(!(a!=null&&a.itemId))return!1;const o=Office.context.mailbox.convertToRestId(a.itemId,Office.MailboxEnums.RestVersion.v2_0),s={Authorization:`Bearer ${await K()}`,"Content-Type":"application/json"},c=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${o}/${n?"createReplyAll":"createReply"}`,{method:"POST",headers:s});if(!c.ok)return!1;const m=await c.json(),y=T=>T.trim().toLowerCase(),l=m.ccRecipients??[],g=new Set(l.map(T=>y(T.emailAddress.address))),_=[...new Set(t.filter(Boolean).map(T=>T.trim()))].filter(T=>!g.has(y(T))).map(T=>({emailAddress:{address:T}})),D={body:{contentType:"HTML",content:e}};return _.length&&(D.ccRecipients=[...l,..._]),(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${m.id}`,{method:"PATCH",headers:s,body:JSON.stringify(D)})).ok?(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${m.id}/send`,{method:"POST",headers:s})).ok:!1}catch{return!1}}async function Ot(e,t){const n=await We(),a=De(n,e);return a&&Re(a.Body||"",t).text||null}async function Q(e){var r;const t=s=>s.trim().toLowerCase(),n=t(((r=i.account)==null?void 0:r.username)??""),a=new Set,o=e.recipients.filter(Boolean).filter(s=>{const c=t(s);return!c||c===n||a.has(c)?!1:(a.add(c),!0)});if(o.length!==0)try{const s=await N(),c=`${H}/_api/web/lists/getbytitle('HD_Notifications')/items`;await Promise.all(o.map(m=>fetch(c,{method:"POST",headers:{Authorization:`Bearer ${s}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:e.title.slice(0,255),RecipientEmail:m,EventType:e.eventType,Message:e.message,LinkPath:e.linkPath,IsRead:!1})})))}catch{}}async function de(e,t){const n=document.querySelectorAll(".email-att-cb:checked");if(n.length===0)return;const a=await N(),o=new Set,r=s=>{if(!o.has(s.toLowerCase()))return o.add(s.toLowerCase()),s;const c=s.lastIndexOf("."),m=c>0?s.slice(0,c):s,y=c>0?s.slice(c):"";for(let l=2;;l++){const g=m+"-"+l+y;if(!o.has(g.toLowerCase()))return o.add(g.toLowerCase()),g}};for(const s of Array.from(n)){const c=s.dataset.attId,m=r(s.dataset.attName),y=await new Promise((T,I)=>{Office.context.mailbox.item.getAttachmentContentAsync(c,{},h=>{h.status===Office.AsyncResultStatus.Succeeded?T(h):I(new Error(h.error.message))})}),{content:l,format:g}=y.value;let _;if(g===Office.MailboxEnums.AttachmentContentFormat.Base64){const T=atob(l);_=new Uint8Array(T.length);for(let I=0;I<T.length;I++)_[I]=T.charCodeAt(I)}else if(g===Office.MailboxEnums.AttachmentContentFormat.Eml||g===Office.MailboxEnums.AttachmentContentFormat.ICalendar)_=new TextEncoder().encode(l);else continue;const D=encodeURIComponent(m),z=`${H}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${D}')`;if(!(await fetch(z,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:_.buffer})).ok)throw new Error(`Upload ${m} failed`)}}async function Lt(e){const t=`https://graph.microsoft.com/v1.0/me/messages/${e}/$value`;let n=await K(),a=await fetch(t,{headers:{Authorization:`Bearer ${n}`}});if((a.status===401||a.status===403)&&(n=await K(!0),a=await fetch(t,{headers:{Authorization:`Bearer ${n}`}})),!a.ok)throw new Error(`Graph ${a.status}`);return a.arrayBuffer()}async function Rt(e){const t=await new Promise((a,o)=>{Office.context.mailbox.getCallbackTokenAsync({isRest:!0},r=>{r.status===Office.AsyncResultStatus.Succeeded?a(r.value):o(new Error("callback token failed"))})}),n=await fetch(`${Office.context.mailbox.restUrl}/v2.0/me/messages/${e}/$value`,{headers:{Authorization:`Bearer ${t}`}});if(!n.ok)throw new Error(`REST ${n.status}`);return n.arrayBuffer()}async function me(e,t){const n=document.getElementById("f-attach-eml");if(!(n!=null&&n.checked))return;const a=Office.context.mailbox.item;if(!a)return;const o=Office.context.mailbox.convertToRestId(a.itemId,Office.MailboxEnums.RestVersion.v2_0);let r,s="",c="";try{r=await Lt(o)}catch(_){s=_ instanceof Error?_.message:String(_);try{r=await Rt(o)}catch(D){c=D instanceof Error?D.message:String(D),console.error("[eml] graph:",s,"| callback:",c),S(`ดึง .eml ไม่ได้ (Graph: ${s} / REST: ${c}) — ไฟล์อื่นบันทึกแล้ว`,"error");return}}const m=(a.subject||"email").replace(/[\\/:*?"<>|#%&{}~]/g,"_").slice(0,100).trim()||"email",y=await N(),l=`${H}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${encodeURIComponent(m+".eml")}')`;(await fetch(l,{method:"POST",headers:{Authorization:`Bearer ${y}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok||S("บันทึกไฟล์ .eml ไม่สำเร็จ","error")}async function ue(e,t,n){const a=await N();for(const o of n){const r=await o.arrayBuffer(),s=encodeURIComponent(o.name),c=`${H}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${s}')`;if(!(await fetch(c,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok)throw new Error(`Upload ${o.name} failed`)}}function S(e,t="success"){const n=document.getElementById("toast-container");if(!n)return;const a=t==="success"?"bg-green-500":t==="error"?"bg-red-500":"bg-slate-700",o=t==="success"?"✅":t==="error"?"❌":"ℹ️",r=document.createElement("div");r.className=`toast pointer-events-auto ${a} text-white text-sm font-medium px-4 py-3 rounded-lg shadow-lg max-w-xs mx-2`,r.textContent=`${o} ${e}`,n.appendChild(r),setTimeout(()=>r.remove(),4e3)}function Mt(e){const t=e.split(`
`).map(l=>l.trim()).filter(Boolean),n=/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/,a=/(\+?[\d\s()\-]{7,}(?:ext\.?\s*\d+)?)/i,o=/\b(co\.?,?\s*ltd\.?|co\.?,?\s*limited|corporation|corp\.?|บริษัท|จำกัด|holding|group|inc\.?|llc)\b/i;let r="",s="",c="";const m=[];for(const l of t)if(!/^[-_=*]{2,}$/.test(l)&&!/^(best regards|regards|sincerely|ขอแสดงความนับถือ|ด้วยความนับถือ|from|sent|thanks|thank you)/i.test(l)){if(!r){const g=l.match(n);if(g){r=g[0];continue}}if(!s){const g=l.match(a);if(g&&g[0].replace(/\D/g,"").length>=7){s=g[0].trim();continue}}if(!c&&o.test(l)){c=l;continue}l.length>=2&&l.length<=50&&!/\d{4,}/.test(l)&&m.push(l)}const y=m.find(l=>!n.test(l)&&!o.test(l))??"";return!r&&!y?null:{name:y,company:c,email:r,phone:s}}async function Ht(){const e=i.signatureContact;if(!e)return;const t=(i.emailSenderEmail||"").toLowerCase();if(t&&i.contactEmails.includes(t)){S("ลูกค้านี้มีในระบบแล้ว","success"),i.signatureContact=null,U();return}const n=document.getElementById("btn-import-customer");n&&(n.disabled=!0,n.textContent="กำลังบันทึก…");try{await V("HD_Contracts",{Title:i.emailSenderName||e.name,CustomerEmail:i.emailSenderEmail,Phone:e.phone||void 0,Company:e.company||void 0,Status:"Active"}),t&&i.contactEmails.push(t),S("เพิ่มลูกค้าสำเร็จ!"),i.signatureContact=null,U()}catch(a){const o=a instanceof Error?a.message:String(a);S(`เกิดข้อผิดพลาด: ${o}`,"error"),n&&(n.disabled=!1,n.textContent="เพิ่มเป็นลูกค้า")}}function pe(){return new Date().toISOString().split("T")[0]}function Nt(){const e=new Date;return`HD-${`${e.getFullYear()}${String(e.getMonth()+1).padStart(2,"0")}${String(e.getDate()).padStart(2,"0")}`}-${Math.floor(Math.random()*900+100)}`}function Ce(){var e;return i.droppedFiles.length>0||document.querySelectorAll(".email-att-cb:checked").length>0||(((e=document.getElementById("f-attach-eml"))==null?void 0:e.checked)??!1)}async function X(e,t){i.droppedFiles.length>0&&await ue(e,t,i.droppedFiles),await de(e,t),await me(e,t)}let re=!1;async function Ft(){var t,n,a,o,r,s,c,m,y,l,g,_,D,z,L,T,I;if(!i.account){S("กรุณาเข้าสู่ระบบก่อน","error");return}if(re)return;re=!0;const e=document.getElementById("submit-btn");e&&(e.disabled=!0,e.textContent="กำลังบันทึก…");try{if(i.tab==="phish")await wt();else if(i.tab==="ticket"){const h=document.getElementById("f-title").value.trim(),f=document.getElementById("f-description").value.trim(),k=document.getElementById("f-priority").value,$=document.getElementById("f-customer-email").value.trim(),B=((t=document.getElementById("f-cc-enable"))==null?void 0:t.checked)??!0,P=B?(((n=document.getElementById("f-cc"))==null?void 0:n.value)||"").split(/[,;\s]+/).map(F=>F.trim()).filter(Boolean):[],C=document.getElementById("f-assigned-email").value,x=i.agents.find(F=>F.email===C),O=Nt(),u=await V("HD_Tickets",{Title:h,TicketNumber:O,Description:f,Priority:k,CustomerEmail:$,CustomerName:i.emailSenderName||$,Status:"Open",AssignedEmail:C||void 0,AssignedToName:(x==null?void 0:x.name)??((a=i.account)==null?void 0:a.name)??"",...Ie(C,(o=i.account)==null?void 0:o.username),ProjectID:parseInt(((r=document.getElementById("f-project"))==null?void 0:r.value)||"0")||null});if(Ce()){const F=await V("HD_TicketComments",{Title:"📎 ไฟล์แนบจากอีเมล",TicketID:u,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await X("HD_TicketComments",F)}i.droppedFiles=[];const v={ticket_number:O,ticket_title:h,priority:k,category:"-",description:le(f||"-"),customer_name:i.emailSenderName||$,assigned_name:(x==null?void 0:x.name)??((s=i.account)==null?void 0:s.name)??"-",link:"https://itservices.co.th/helpdesk/"},j=[C,i.account.username,...P,Ge].filter(Boolean);let w=!1;const R=await Ot("ticket_created",v);if(R){const F=`<div style="border-left:4px solid #2563eb;background:#eff6ff;padding:10px 14px;margin:0 0 14px;font-family:Segoe UI,sans-serif">
             <div style="font-size:15px;font-weight:700;color:#1e40af">Ticket No. ${O}</div>
             <div style="font-size:12px;color:#475569;margin-top:2px">กรุณาตอบกลับในอีเมลฉบับนี้เพื่อให้ข้อมูลอยู่ใน Ticket เดียวกัน</div>
           </div>`;w=await Dt(F+R,j,B)}let W=w;w||(W=await Se("ticket_created",v,[$],j)),W?S(w?"สร้าง Ticket และตอบกลับในเธรดเดิมแล้ว!":"สร้าง Ticket สำเร็จ!"):S(`สร้าง Ticket ${O} แล้ว แต่ส่งเมลถึงลูกค้าไม่สำเร็จ — ตรวจ template ticket_created ที่หน้า Diagnostic`,"error")}else if(i.tab==="task"){const h=document.getElementById("f-title").value.trim(),f=parseInt(((c=document.getElementById("f-project"))==null?void 0:c.value)||"0"),k=document.getElementById("f-due-date").value,$=document.getElementById("f-note").value.trim(),B=document.getElementById("f-assigned-email").value,P=i.agents.find(O=>O.email===B);if(!f){S("กรุณาเลือก Project","error");return}const C=await V("PM_Tasks",{Title:h,DueDate:k||null,TaskNote:$,AssignedTo:(P==null?void 0:P.name)??i.account.name??i.account.username,AssignedEmail:B,IsCompleted:!1,IsAcknowledged:!1,ProjectID:f});if(i.droppedFiles.length>0&&await ue("PM_Tasks",C,i.droppedFiles),await de("PM_Tasks",C),await me("PM_Tasks",C),i.droppedFiles=[],await Q({recipients:[B],title:`📋 ได้รับมอบหมาย Task: ${h}`,message:$||(k?`กำหนดส่ง ${k}`:"มี Task ใหม่"),linkPath:f?`/projects/${f}`:"/my-work",eventType:"task_assigned"}),((m=document.getElementById("f-teams"))==null?void 0:m.checked)&&k){const O=Array.from(document.querySelectorAll(".att-internal:checked")).map(w=>w.value),u=(((y=document.getElementById("f-ext-att"))==null?void 0:y.value)||"").split(/[,;\s]+/).map(w=>w.trim()).filter(Boolean),v=`${k}T09:00:00`,j=`${k}T10:00:00`;try{await Bt({subject:h,start:v,end:j,body:$,attendees:[...O,...u],isOnlineMeeting:!0}),S("สร้าง Task + นัดประชุม Teams สำเร็จ!")}catch(w){S("สร้าง Task แล้ว แต่สร้างนัดประชุมไม่สำเร็จ: "+(w instanceof Error?w.message:""),"error")}}else S("สร้าง Task สำเร็จ!")}else if(i.tab==="incident"){const h=document.getElementById("f-title").value.trim(),f=parseInt(((l=document.getElementById("f-project"))==null?void 0:l.value)||"0"),k=document.getElementById("f-description").value.trim(),$=document.getElementById("f-severity").value,B=document.getElementById("f-assigned-email").value,P=i.agents.find(j=>j.email===B),C=document.getElementById("f-status").value,x=document.getElementById("f-incident-date").value,O=document.getElementById("f-resolution").value.trim();if(!f){S("กรุณาเลือก Project","error");return}const u=parseInt(((g=document.getElementById("f-sla"))==null?void 0:g.value)||"0")||null,v=await V("PM_Incidents",{Title:h,Description:k||void 0,Severity:$,Status:C,AssignedTo:(P==null?void 0:P.name)??i.account.name??i.account.username,AssignedEmail:B,...Ie(B,(_=i.account)==null?void 0:_.username),ProjectID:f,IncidentDate:x||pe(),Resolution:O||void 0,SLAHours:u,SLADue:Me(u),...C==="Resolved"?{ResolvedDate:new Date().toISOString()}:{}});i.droppedFiles.length>0&&await ue("PM_Incidents",v,i.droppedFiles),await de("PM_Incidents",v),await me("PM_Incidents",v),i.droppedFiles=[],await Q({recipients:[B],title:`🚨 ได้รับมอบหมาย Incident: ${h}`,message:`ความรุนแรง ${$}${k?" — "+k.slice(0,120):""}`,linkPath:f?`/projects/${f}`:"/my-work",eventType:"incident_created"});{const j=i.projects.find(F=>F.id===f);let w;try{const F=await N(),xe=await fetch(`${H}/_api/web/lists/getbytitle('PM_Projects')/items(${f})?$select=CreatedByEmail`,{headers:{Authorization:`Bearer ${F}`,Accept:"application/json;odata=nometadata"}});xe.ok&&(w=(await xe.json()).CreatedByEmail)}catch{}const R=dt({title:h,severity:$,status:C,description:k,incidentDate:x||pe(),slaHours:u,projectName:j==null?void 0:j.Title,projectId:f,assignedName:P==null?void 0:P.name,assignedEmail:B,requesterEmail:i.account.username,watchers:[w,i.emailSenderEmail],actorEmail:i.account.username,baseUrl:"https://itservices.co.th/helpdesk/"});if(R.to.length===0){S("สร้าง Incident สำเร็จ — ไม่ได้ส่งเมล เพราะไม่มีผู้รับนอกจากตัวคุณเอง (มอบหมายให้ตัวเอง และไม่มีผู้เกี่ยวข้องอื่น)");return}await Se("incident_created",R.vars,R.to,R.cc)?S(`สร้าง Incident และส่งเมลแจ้ง ${[...R.to,...R.cc].length} คนแล้ว!`):S("สร้าง Incident แล้ว แต่ส่งเมลแจ้งไม่สำเร็จ — ผู้รับผิดชอบยังไม่รู้เรื่อง (ตรวจ template incident_created)","error")}}else if(i.tab==="comment"){const h=parseInt(((D=document.getElementById("f-ticket"))==null?void 0:D.value)||"0"),f=document.getElementById("f-comment").value.trim(),k=document.getElementById("f-comment-type").value;if(!h){S("กรุณาเลือก Ticket","error");return}if(!f){S("กรุณาพิมพ์ Comment","error");return}const $=await V("HD_TicketComments",{Title:f.slice(0,100),TicketID:h,CommentText:f,CommentType:k,CommentDate:new Date().toISOString()});await X("HD_TicketComments",$),i.droppedFiles=[];try{const B=await N(),P=`${H}/_api/web/lists/getbytitle('HD_Tickets')/items(${h})?$select=TicketNumber,Title,AssignedEmail,Author/EMail&$expand=Author`,C=await fetch(P,{headers:{Authorization:`Bearer ${B}`,Accept:"application/json;odata=nometadata"}});if(C.ok){const x=await C.json(),O=i.account.username.toLowerCase(),u=[...new Set([x.AssignedEmail,(z=x.Author)==null?void 0:z.EMail].filter(Boolean))].filter(v=>v.toLowerCase()!==O);u.length&&await Q({recipients:u,title:`💬 ${((L=i.account)==null?void 0:L.name)??"มีคน"} คอมเมนต์ใน ${x.TicketNumber||"#"+h}`,message:f.slice(0,200),linkPath:`/tickets/${h}`,eventType:"comment_added"})}}catch{}S("เพิ่ม Comment สำเร็จ!")}else if(i.tab==="project"){const h=document.getElementById("f-title").value.trim(),f=document.getElementById("f-company").value.trim(),k=document.getElementById("f-group").value,$=document.getElementById("f-status").value,B=document.getElementById("f-start").value,P=document.getElementById("f-end").value,C=document.getElementById("f-description").value.trim();if(!h){S("กรุณาใส่ชื่อโครงการ","error");return}const x=await V("PM_Projects",{Title:h,Company:f||void 0,ProjectGroup:k,Progress:0,StartDate:B||void 0,EndDate:P||null,Status:$,CreatedByEmail:i.account.username,Comment:C||void 0});if(Ce()){const O=await V("PM_Comments",{Title:"📎 ไฟล์แนบจากอีเมล",ProjectID:x,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await X("PM_Comments",O)}i.droppedFiles=[],S("สร้างโครงการสำเร็จ!")}else if(i.tab==="projcomment"){const h=parseInt(((T=document.getElementById("f-project"))==null?void 0:T.value)||"0"),f=document.getElementById("f-comment").value.trim(),k=document.getElementById("f-comment-type").value;if(!h){S("กรุณาเลือกโครงการ","error");return}if(!f){S("กรุณาพิมพ์ Comment","error");return}const $=await V("PM_Comments",{Title:f.slice(0,100),ProjectID:h,CommentText:f,CommentType:k,CommentDate:new Date().toISOString(),ParentID:0});await X("PM_Comments",$),i.droppedFiles=[];try{const B=await N(),P=`${H}/_api/web/lists/getbytitle('PM_Projects')/items(${h})?$select=Title,CreatedByEmail`,C=await fetch(P,{headers:{Authorization:`Bearer ${B}`,Accept:"application/json;odata=nometadata"}});if(C.ok){const x=await C.json(),O=i.account.username.toLowerCase();x.CreatedByEmail&&x.CreatedByEmail.toLowerCase()!==O&&await Q({recipients:[x.CreatedByEmail],title:`💬 ${((I=i.account)==null?void 0:I.name)??"มีคน"} คอมเมนต์ในโครงการ ${x.Title??""}`,message:f.slice(0,200),linkPath:`/projects/${h}?tab=comments`,eventType:"comment_added"})}}catch{}S("เพิ่ม Comment สำเร็จ!")}}catch(h){const f=h instanceof Error?h.message:String(h);S(`เกิดข้อผิดพลาด: ${f}`,"error")}finally{re=!1,e&&(e.disabled=!1,e.textContent="บันทึก")}}const zt={ticket:{label:"Ticket",icon:"🎫"},comment:{label:"Comment",icon:"💬"},project:{label:"Project",icon:"📁"},task:{label:"Task",icon:"✅"},incident:{label:"Incident",icon:"🚨"},projcomment:{label:"Comment",icon:"💬"},phish:{label:"PhishGuard",icon:"🛡️"}},qt=[{title:"🎫 Helpdesk",tabs:["ticket","comment"]},{title:"📁 Project",tabs:["project","task","incident","projcomment"]},{title:"🛡️ Security",tabs:["phish"]}],Je=["f-title","f-description","f-priority","f-customer-email","f-cc","f-assigned-email","f-project","f-due-date","f-note","f-severity","f-status","f-sla","f-incident-date","f-resolution","f-ticket","f-comment","f-comment-type","f-company","f-group","f-start","f-end","f-ext-att","f-attach-eml"];let G={};function Ut(){for(const t of Je){const n=document.getElementById(t);n&&(G[t]=n.value)}const e=document.getElementById("f-teams");e&&(G["f-teams"]=e.checked)}function Vt(){for(const t of Je){const n=document.getElementById(t);n&&G[t]!==void 0&&G[t]!==""&&(n.value=G[t])}const e=document.getElementById("f-teams");if(e&&G["f-teams"]!==void 0){e.checked=G["f-teams"];const t=document.getElementById("teams-fields");t&&(t.style.display=e.checked?"block":"none")}}function U(){var $,B,P,C,x,O;const e=document.getElementById("app");if(!e)return;Ut();const{account:t,tab:n,emailSubject:a,emailSenderName:o,emailSenderEmail:r,emailBodyPreview:s}=i,c=i.emailBodyReply||s,m=t!==null,y=`
    <div class="bg-gradient-to-r from-blue-700 to-blue-600 text-white px-4 py-2.5 flex items-center gap-2.5 shadow flex-shrink-0">
      <div class="w-7 h-7 bg-white/95 rounded-lg flex items-center justify-center flex-shrink-0">
        <span class="text-blue-700 font-extrabold text-xs">iT</span>
      </div>
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm leading-tight truncate">iT Services Helpdesk</div>
        ${m?`<div class="text-[10px] text-blue-100 truncate">${b((t==null?void 0:t.name)??(t==null?void 0:t.username)??"")}</div>`:""}
      </div>
      ${m?`<button id="btn-logout" title="ออกจากระบบ" class="p-1.5 rounded-lg hover:bg-white/15 transition flex-shrink-0">
             <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
           </button>`:`<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded-lg transition flex-shrink-0">
             เข้าสู่ระบบ
           </button>`}
    </div>
  `;if(!m){e.innerHTML=`
      ${y}
      <div class="flex flex-col items-center justify-center px-6 py-16 gap-4 text-center">
        <div class="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center text-3xl">🔐</div>
        <h2 class="text-base font-semibold text-slate-700">เข้าสู่ระบบด้วย Microsoft</h2>
        <p class="text-xs text-slate-500">เพื่อสร้าง Ticket / Task / Incident จาก Email นี้</p>
        <button id="btn-login-main"
          class="mt-2 w-full bg-blue-700 hover:bg-blue-600 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          เข้าสู่ระบบ
        </button>
      </div>
    `,($=document.getElementById("btn-login"))==null||$.addEventListener("click",Ee),(B=document.getElementById("btn-login-main"))==null||B.addEventListener("click",Ee);return}const l=a?`<div class="mx-3 mt-3 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5 text-xs text-slate-700 space-y-1">
         <div class="font-semibold text-blue-800 truncate" title="${b(a)}">📧 ${b(a)}</div>
         ${o?`<div class="text-slate-500">จาก: <span class="font-medium text-slate-700">${b(o)}</span></div>`:""}
         ${r&&r!==o?`<div class="text-slate-400 truncate">${b(r)}</div>`:""}
       </div>`:`<div class="mx-3 mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-700">
         ⚠️ ไม่พบข้อมูล Email (โหมดทดสอบ)
       </div>`,g=i.signatureContact,_=!!r&&i.contactEmails.includes(r.toLowerCase()),D=g?`<div class="mx-3 mt-3 bg-orange-50 border border-orange-200 rounded-lg px-3 py-2.5 text-xs text-slate-700">
         <div class="flex items-center justify-between mb-2">
           <span class="font-semibold text-orange-700">👤 ข้อมูลผู้ส่ง (จาก Signature)</span>
         </div>
         <div class="space-y-0.5 mb-2.5">
           ${o?`<div><span class="text-slate-400">ชื่อ:</span> <span class="font-medium">${b(o)}</span></div>`:""}
           ${g.company?`<div><span class="text-slate-400">บริษัท:</span> ${b(g.company)}</div>`:""}
           ${r?`<div><span class="text-slate-400">Email:</span> ${b(r)}</div>`:""}
           ${g.phone?`<div><span class="text-slate-400">โทร:</span> ${b(g.phone)}</div>`:""}
         </div>
         ${_?'<div class="w-full bg-green-100 text-green-700 text-xs font-semibold py-1.5 rounded-md text-center">✓ ลูกค้านี้มีในระบบแล้ว</div>':`<button id="btn-import-customer"
                class="w-full bg-orange-500 hover:bg-orange-400 text-white text-xs font-semibold py-1.5 rounded-md transition">
                + เพิ่มเป็นลูกค้า
              </button>`}
       </div>`:"",z=`
    <div class="mx-3 mt-3 space-y-2">
      ${qt.map(u=>`
        <div>
          <div class="text-[10px] font-semibold text-slate-400 uppercase tracking-wide mb-1 px-0.5">${u.title}</div>
          <div class="grid grid-cols-4 gap-1">
            ${u.tabs.map(v=>{const j=zt[v];return`<button data-tab="${v}"
                class="tab-btn flex flex-col items-center gap-1 py-2 rounded-lg transition ${n===v?"bg-blue-700 text-white shadow":"text-slate-500 hover:bg-slate-100"}">
                <span class="text-base leading-none">${j.icon}</span>
                <span class="text-[9px] font-medium leading-none">${j.label}</span>
              </button>`}).join("")}
          </div>
        </div>
      `).join("")}
    </div>
  `;let L="";n==="phish"?L=Et(!!t):n==="ticket"?L=`
      ${A("Title / หัวข้อ",`<input id="f-title" type="text"
        class="${E}"
        value="${b(a)}" />`)}
      ${A("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${E} resize-none">${b(s)}</textarea>`)}
      ${A("Priority",`<select id="f-priority" class="${E}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
      ${A("Customer Email",`<input id="f-customer-email" type="email"
        class="${E}"
        value="${b(r)}" />`)}
      ${A("CC — ให้ผู้ที่อยู่ในเมลนี้รับรู้",`
        <label class="flex items-center gap-2 text-xs text-slate-600 mb-1.5 cursor-pointer">
          <input id="f-cc-enable" type="checkbox" ${i.emailCc.length?"checked":""} /> แนบผู้รับในเมลนี้เป็น CC อัตโนมัติ
        </label>
        <input id="f-cc" type="text" class="${E}" value="${b(i.emailCc.join(", "))}" placeholder="someone@company.com, boss@company.com" />`)}
      ${A("Assign ให้ Agent",ce(t.username))}
      ${A("โครงการ (ไม่บังคับ)",ee(!0))}
      ${J()}
    `:n==="task"?L=`
      ${A("ชื่อ Task *",`<input id="f-title" type="text" required
        class="${E}" value="${b(a)}" />`)}
      ${A("Project *",ee())}
      ${A("Assign ให้",ce(t.username))}
      ${A("Due Date",`<input id="f-due-date" type="date" class="${E}" />`)}
      ${A("Task Note",`<textarea id="f-note" rows="4"
        class="${E} resize-y">${b(s)}</textarea>`)}
      <label class="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer mb-1">
        <input id="f-teams" type="checkbox" class="rounded" onchange="document.getElementById('teams-fields').style.display=this.checked?'block':'none'" />
        💻 เพิ่มการประชุมออนไลน์ (Teams) — ใช้เวลา 09:00–10:00 ของวัน Due Date
      </label>
      <div id="teams-fields" style="display:none" class="space-y-2 mb-2">
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">ผู้เข้าร่วม Internal</label>
          <div class="max-h-28 overflow-y-auto border border-slate-200 rounded-md p-1.5 space-y-0.5">
            ${i.agents.map(u=>`<label class="flex items-center gap-2 text-xs text-slate-700 px-1 py-0.5 hover:bg-slate-50 rounded cursor-pointer">
              <input type="checkbox" class="att-internal" value="${b(u.email)}" /> ${b(u.name)}
            </label>`).join("")}
          </div>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">Email ภายนอก (คั่นด้วย ,)</label>
          <input id="f-ext-att" type="text" class="${E}" placeholder="someone@company.com, ..." />
        </div>
      </div>
      ${J()}
    `:n==="incident"?L=`
      ${A("ชื่อ Incident *",`<input id="f-title" type="text" required
        class="${E}" value="${b(a)}" />`)}
      ${A("Project *",ee())}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">ความรุนแรง</label>
          <select id="f-severity" class="${E}">
            <option value="Low">Low</option>
            <option value="Medium" selected>Medium</option>
            <option value="High">High</option>
            <option value="Critical">Critical</option>
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${E}">
            <option value="Open" selected>Open</option>
            <option value="In Progress">In Progress</option>
            <option value="Resolved">Resolved</option>
          </select>
        </div>
      </div>
      ${A("SLA — ต้องแก้ให้จบภายใน",`
        <select id="f-sla" class="${E}">
          <option value="">ไม่กำหนด SLA</option>
          ${mt.map(u=>`<option value="${u.hours}" ${u.hours===se.Medium?"selected":""}>${u.labelTh}</option>`).join("")}
        </select>
        <p id="f-sla-hint" class="text-[11px] text-slate-400 mt-1">นับจากตอนนี้ · ครบกำหนด ${$e(se.Medium)}</p>`)}
      ${A("Assign ให้ Agent",ce(t.username))}
      ${A("วันที่เกิด Incident",`<input id="f-incident-date" type="date" class="${E}" value="${pe()}" />`)}
      ${A("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${E} resize-y">${b(s)}</textarea>`)}
      ${A("วิธีแก้ไข (ถ้ามี)",`<textarea id="f-resolution" rows="2"
        class="${E} resize-y" placeholder="อธิบายวิธีแก้ไขปัญหา..."></textarea>`)}
      ${J()}
    `:n==="comment"?L=`
      ${A("เลือก Ticket *",`<select id="f-ticket" class="${E}">
        <option value="">-- เลือก Ticket ที่จะเพิ่ม Comment --</option>
        ${i.tickets.map(u=>`<option value="${u.id}">${b(u.TicketNumber||"#"+u.id)} · ${b(u.Title)}</option>`).join("")}
      </select>`)}
      ${A("ประเภท",`<select id="f-comment-type" class="${E}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${A("Comment *",`<textarea id="f-comment" rows="5"
        class="${E} resize-y" placeholder="พิมพ์ comment...">${b(c)}</textarea>`)}
      ${J()}
    `:n==="project"?L=`
      ${A("ชื่อโครงการ *",`<input id="f-title" type="text" required
        class="${E}" value="${b(a)}" />`)}
      ${A("บริษัท / ลูกค้า",`<input id="f-company" type="text" class="${E}" value="${b(((P=i.signatureContact)==null?void 0:P.company)??"")}" />`)}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">กลุ่มโครงการ</label>
          <select id="f-group" class="${E}">
            ${["Internal","External","R&D","Maintenance","อื่นๆ"].map(u=>`<option>${u}</option>`).join("")}
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${E}">
            ${["Planning","Active","On Hold","Completed","Cancelled"].map(u=>`<option>${u}</option>`).join("")}
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันที่เริ่ม</label>
          <input id="f-start" type="date" class="${E}" /></div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันสิ้นสุด</label>
          <input id="f-end" type="date" class="${E}" /></div>
      </div>
      ${A("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${E} resize-y">${b(s)}</textarea>`)}
      ${J()}
    `:n==="projcomment"&&(L=`
      ${A("เลือกโครงการ *",ee())}
      ${A("ประเภท",`<select id="f-comment-type" class="${E}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${A("Comment *",`<textarea id="f-comment" rows="5"
        class="${E} resize-y" placeholder="พิมพ์ comment...">${b(c)}</textarea>`)}
      ${J()}
    `);const T=n==="phish"?$t():n==="comment"||n==="projcomment"?"เพิ่ม Comment":n==="project"?"สร้างโครงการ":n==="incident"?"แจ้ง Incident":n==="task"?"สร้าง Task":"สร้าง Ticket";e.innerHTML=`
    <div class="flex flex-col h-screen bg-slate-50">
      ${y}
      <div class="flex-1 overflow-y-auto">
        ${l}
        ${D}
        ${z}
        <div class="mx-3 mt-3 space-y-3 pb-3">
          ${L}
        </div>
      </div>
      <div class="border-t border-slate-200 bg-white px-3 py-2.5 flex-shrink-0">
        <button id="submit-btn"
          class="w-full bg-blue-700 hover:bg-blue-600 active:bg-blue-800 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          ${T}
        </button>
      </div>
    </div>
  `,(C=document.getElementById("btn-logout"))==null||C.addEventListener("click",jt),(x=document.getElementById("submit-btn"))==null||x.addEventListener("click",Ft),(O=document.getElementById("btn-import-customer"))==null||O.addEventListener("click",Ht),n==="phish"&&It();const I=document.getElementById("f-severity"),h=document.getElementById("f-sla");if(I&&h){let u=!1;const v=document.getElementById("f-sla-hint"),j=()=>{if(!v)return;const w=parseInt(h.value||"0")||null;v.textContent=w?`นับจากตอนนี้ · ครบกำหนด ${$e(w)}`:"ไม่กำหนด SLA — เคสนี้จะวัดไม่ได้ในรายงาน"};h.addEventListener("change",()=>{u=!0,j()}),I.addEventListener("change",()=>{if(u)return;const w=se[I.value];w&&(h.value=String(w),j())})}document.querySelectorAll(".tab-btn").forEach(u=>{u.addEventListener("click",()=>{const v=u.dataset.tab;v&&v!==i.tab&&(i.tab=v,U(),v==="phish"&&ne())})});const f=document.getElementById("drop-zone"),k=document.getElementById("f-files");f&&k&&(k.addEventListener("change",()=>{k.files&&fe(Array.from(k.files)),k.value=""}),f.addEventListener("dragover",u=>{u.preventDefault(),f.classList.add("border-blue-500","bg-blue-50")}),f.addEventListener("dragleave",()=>{f.classList.remove("border-blue-500","bg-blue-50")}),f.addEventListener("drop",u=>{var j;u.preventDefault(),f.classList.remove("border-blue-500","bg-blue-50");const v=Array.from(((j=u.dataTransfer)==null?void 0:j.files)??[]);v.length&&fe(v)})),document.querySelectorAll(".remove-dropped").forEach(u=>{u.addEventListener("click",()=>{const v=parseInt(u.dataset.remove??"-1");v>=0&&(i.droppedFiles.splice(v,1),U())})}),Vt()}function fe(e){i.droppedFiles.push(...e),U()}document.addEventListener("paste",e=>{var a;if(!i.account)return;const t=Array.from(((a=e.clipboardData)==null?void 0:a.items)??[]),n=[];for(const o of t)if(o.kind==="file"){const r=o.getAsFile();if(r){const s=r.name&&r.name!=="image.png"?r.name:`screenshot-${new Date().toISOString().replace(/[:.]/g,"-").slice(0,19)}.png`;n.push(new File([r],s,{type:r.type}))}}n.length&&(e.preventDefault(),fe(n),S(`แนบไฟล์แล้ว: ${n.map(o=>o.name).join(", ")}`))});const E="w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white";function Be(e){return e<1024?`${e} B`:e<1024*1024?`${(e/1024).toFixed(0)} KB`:`${(e/1024/1024).toFixed(1)} MB`}function J(){const e=i.emailAttachments,t=i.droppedFiles,n=c=>`
    <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
      <input type="checkbox" class="email-att-cb" data-att-id="${b(c.id)}" data-att-name="${b(c.name)}" data-att-item="${c.isItem?"1":"0"}" ${c.defaultOn?"checked":""} />
      <span class="flex-1 truncate">${c.isItem?"📧 ":c.isInline?"🖼️ ":""}${b(c.name)}</span>
      <span class="text-slate-400 flex-shrink-0">${Be(c.size)}</span>
    </label>`,a=e.filter(c=>!c.isInline),o=e.filter(c=>c.isInline),r=e.length>0?`<div class="mb-2 space-y-1">
        ${a.length?`<p class="text-xs text-slate-500">📎 ไฟล์แนบจาก Email:</p>
        ${a.map(n).join("")}`:""}
        ${o.length?`<p class="text-xs text-slate-500 ${a.length?"pt-1":""}">🖼️ รูปในเนื้อเมล
          <span class="text-slate-400">(รูปเล็กมักเป็นโลโก้ในลายเซ็น — ติ๊กเพิ่มได้)</span></p>
        ${o.map(n).join("")}`:""}
      </div>`:"",s=t.length>0?`<div class="mt-2 space-y-1">
        ${t.map((c,m)=>`<div class="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 rounded px-2 py-1">
            <span class="text-base">${c.type.startsWith("image/")?"🖼️":"📄"}</span>
            <span class="flex-1 truncate">${b(c.name)}</span>
            <span class="text-slate-400">${Be(c.size)}</span>
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
  </div>`}function ce(e){var t;return`<select id="f-assigned-email" class="${E}">
    <option value="${b(e)}">${b(((t=i.account)==null?void 0:t.name)??e)} (ฉัน)</option>
    ${i.agents.filter(n=>n.email!==e).map(n=>`<option value="${b(n.email)}">${b(n.name)}</option>`).join("")}
  </select>`}function ee(e=!1){return i.projects.length===0?e?'<div class="text-xs text-slate-400">ไม่พบ Project ที่ Active</div>':'<div class="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-2">⚠️ ไม่พบ Project ที่ Active</div>':`<select id="f-project" class="${E}">
    <option value="">${e?"-- ไม่ผูกกับโครงการ --":"-- เลือก Project --"}</option>
    ${i.projects.map(t=>`<option value="${t.id}">${b(t.Title)}</option>`).join("")}
  </select>`}function A(e,t){return`
    <div class="space-y-1">
      <label class="block text-xs font-medium text-slate-600">${e}</label>
      ${t}
    </div>
  `}function b(e){return e.replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}async function Gt(){ut({sharepointUrl:H,internalDomains:["itservices.co.th","rpaexpert.com","rpaexpert.onmicrosoft.com"],getToken:N,getGraphToken:()=>K(),account:()=>i.account?{name:i.account.name,username:i.account.username}:null,toast:(t,n)=>S(t,n??"success"),rerender:U,canWhitelist:()=>["Agent","Supervisor","Boss","Admin"].includes(i.myRole)}),await M.initialize(),await M.handleRedirectPromise();const e=M.getAllAccounts();if(e.length>0){i.account=e[0];try{await M.acquireTokenSilent({scopes:[te],account:e[0]}),await Promise.all([ze(),qe(),Ue(),Ve()])}catch{i.account=null}}typeof Office<"u"?Office.onReady(t=>{var n;if(t.host===Office.HostType.Outlook){const a=Office.context.mailbox.item;if(a){i.emailSubject=a.subject??"";const o=a.from;o&&(i.emailSenderName=o.displayName??"",i.emailSenderEmail=o.emailAddress??"");const r=(((n=i.account)==null?void 0:n.username)??"").toLowerCase(),s=((o==null?void 0:o.emailAddress)??"").toLowerCase(),c=[...a.to??[],...a.cc??[]].map(l=>l.emailAddress).filter(Boolean);i.emailCc=[...new Set(c.map(l=>l.toLowerCase()))].filter(l=>l!==r&&l!==s);const m=a.attachments??[],y=20*1024;i.emailAttachments=m.filter(l=>l.attachmentType===Office.MailboxEnums.AttachmentType.File||l.attachmentType===Office.MailboxEnums.AttachmentType.Item).map(l=>({id:l.id,name:l.attachmentType===Office.MailboxEnums.AttachmentType.Item?`${(l.name||"email").replace(/\.eml$/i,"")}.eml`:l.name,size:l.size,isItem:l.attachmentType===Office.MailboxEnums.AttachmentType.Item,isInline:!!l.isInline,defaultOn:!l.isInline||l.size>=y})),a.body.getAsync(Office.CoercionType.Html,{asyncContext:{}},l=>{if(l.status===Office.AsyncResultStatus.Succeeded){let g=function(x,O=!1){if(x.nodeType===3){const w=x.textContent??"";return O&&w.trim()===""?"":w}const u=x,v=(u.tagName??"").toLowerCase();if(z.includes(v))return"";if(v==="br")return" ";if(v==="tr"){const w=[];for(let R=0;R<u.childNodes.length;R++){const W=u.childNodes[R],F=(W.tagName??"").toLowerCase();(F==="td"||F==="th")&&w.push((W.textContent??"").replace(/\s+/g," ").trim())}return w.length?w.join("	")+`
`:""}if(T.includes(v)){let w="";for(let R=0;R<u.childNodes.length;R++)w+=g(u.childNodes[R],!0);return w}let j="";for(let w=0;w<u.childNodes.length;w++)j+=g(u.childNodes[w],!1);return L.includes(v)&&(j=`
`+j.trim()+`
`),j};const _=l.value,D=new DOMParser().parseFromString(_,"text/html"),z=["style","script","head","img","meta","link","noscript"],L=["p","div","li","h1","h2","h3","h4","h5","h6","blockquote"],T=["table","thead","tbody","tfoot"],f=g(D.body??D.documentElement).replace(/[ \t]{2,}/g," ").replace(/\n[ \t]+/g,`
`).replace(/\n{3,}/g,`

`).trim().split(`
`),k=[];let $="";for(const x of f)x.trim()===""?$&&(k.push($.trim()),$=""):x.includes("	")?($&&(k.push($.trim()),$=""),k.push(x)):$=$?$+" "+x.trim():x.trim();$&&k.push($.trim());const B=k.join(`
`),P=/\n([-_]{3,}|From:\s|Best regards|Regards,|ขอแสดงความนับถือ|Sent:\s)/i,C=B.search(P);i.signatureContact=C>80?Mt(B.slice(C).trim()):null,i.emailBodyPreview=B.trim().slice(0,2e4),i.emailBodyReply=tt(i.emailBodyPreview)}U()});return}}je(),U()}):(je(),U())}function je(){i.emailSubject="[DEV] Test Email Subject",i.emailSenderName="Test Sender",i.emailSenderEmail="test@example.com",i.emailBodyPreview="This is a placeholder email body for development mode.",i.emailBodyReply=i.emailBodyPreview}Gt().catch(e=>{console.error("Init error:",e);const t=document.getElementById("app");t&&(t.innerHTML=`<div class="p-4 text-red-600 text-sm">เกิดข้อผิดพลาด: ${String(e)}</div>`)});
