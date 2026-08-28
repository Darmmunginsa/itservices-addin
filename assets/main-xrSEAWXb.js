import{P as He}from"./PublicClientApplication-DLKYUtZW.js";import{a as te,d as Ne,p as Fe,L as Ee,S as Ie}from"./analyzer-CWd3MChg.js";const ze=/^\s*(Sent|To|Date|Cc|Subject|ส่ง|ถึง|วันที่|สำเนา|เรื่อง)\s*:/i;function qe(t,e){const a=t[e];if(/^\s*-{2,}\s*(Original Message|Forwarded message|ข้อความต้นฉบับ)\s*-{2,}/i.test(a)||/^\s*_{5,}\s*$/.test(a)||/^\s*(On|เมื่อ)\b.{10,200}(wrote|เขียนว่า)\s*:\s*$/i.test(a))return!0;if(/^\s*(From|จาก)\s*:\s*\S/i.test(a)){for(let n=e+1;n<=e+3&&n<t.length;n++)if(ze.test(t[n]))return!0;return!1}return!1}function Ue(t){const e=(t??"").replace(/\r\n/g,`
`),a=e.split(`
`);let n=-1;for(let s=0;s<a.length;s++){if(s>0&&qe(a,s)){n=s;break}if(s>0&&/^\s*>/.test(a[s])&&/^\s*>/.test(a[s+1]??"")){n=s;break}}if(n<0)return{visible:e.trim(),quoted:""};const o=a.slice(0,n).join(`
`).trim(),r=a.slice(n).join(`
`).trim();return o?{visible:o,quoted:r}:{visible:e.trim(),quoted:""}}const Ge=t=>Ue(t).visible,Ve=[{hours:1,labelTh:"1 ชั่วโมง"},{hours:2,labelTh:"2 ชั่วโมง"},{hours:4,labelTh:"4 ชั่วโมง"},{hours:8,labelTh:"8 ชั่วโมง (1 วันทำการ)"},{hours:24,labelTh:"24 ชั่วโมง"},{hours:48,labelTh:"2 วัน"},{hours:72,labelTh:"3 วัน"},{hours:168,labelTh:"7 วัน"}],ae={Critical:1,High:4,Medium:24,Low:72};function Ae(t,e=new Date){const a=typeof t=="number"&&Number.isFinite(t)&&t>0?t:null;return a?new Date(e.getTime()+a*36e5).toISOString():null}function ye(t){const e=Ae(t);return e?new Date(e).toLocaleString("th-TH",{dateStyle:"short",timeStyle:"short"}):""}const Ce="HD_PhishingReports";let f;const d={mail:null,analysis:null,loading:!1,reporting:!1,reported:!1,showHeaders:!1,kasmTemplate:"",analysedItemId:"",safeDomains:[],safeDomainIds:{},savingDomain:""};function Ke(t){f=t}const F=t=>(t??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"),W=()=>{var t,e;return(e=(t=Office.context)==null?void 0:t.mailbox)==null?void 0:e.item},xe=t=>new Promise(e=>{const a=W();if(!(a!=null&&a.body)){e("");return}a.body.getAsync(t,n=>e(n.status===Office.AsyncResultStatus.Succeeded?n.value??"":""))});function We(t){return t?t.split(",").map(e=>{const a=e.trim(),n=a.match(/^(.*?)\s*<([^>]+)>$/);return n?{name:n[1].replace(/^"|"$/g,"").trim(),email:n[2].trim()}:{name:"",email:a.replace(/[<>]/g,"").trim()}}).filter(e=>e.email.includes("@")):[]}function Je(){return new Promise(t=>{let e=!1;try{e=Office.context.requirements.isSetSupported("Mailbox","1.8")}catch{e=!1}const a=W();if(!e||typeof(a==null?void 0:a.getAllInternetHeadersAsync)!="function"){t({});return}try{a.getAllInternetHeadersAsync(n=>t(n.status===Office.AsyncResultStatus.Succeeded?Fe(n.value??""):{}))}catch{t({})}})}async function Ye(){try{const t=W();if(!(t!=null&&t.itemId))return{};const e=Office.context.mailbox.convertToRestId(t.itemId,Office.MailboxEnums.RestVersion.v2_0),a=await f.getGraphToken(),n=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${e}?$select=internetMessageHeaders`,{headers:{Authorization:`Bearer ${a}`}});if(!n.ok)return{};const o=await n.json(),r={};for(const s of o.internetMessageHeaders??[])r[s.name]=r[s.name]?`${r[s.name]}
${s.value}`:s.value;return r}catch{return{}}}async function ue(t,e){const a=await f.getToken(),n=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('${t}')/items?${e}`,{headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata"}});return n.ok?(await n.json()).value:[]}async function Ze(){try{return(await ue("HD_AgentProfiles","$select=Title,EmailText&$top=500")).filter(e=>e.EmailText).map(e=>({name:e.Title,email:e.EmailText}))}catch{return[]}}async function Qe(){var t;try{return(((t=(await ue("HD_Options","$select=Title&$filter=Category eq 'KasmConfig'&$top=1"))[0])==null?void 0:t.Title)??"").trim()}catch{return""}}const Se="SafeDomain";async function pe(){try{const t=await ue("HD_Options",`$select=Id,Title&$filter=Category eq '${Se}'&$top=500`),e={},a=[];for(const n of t){const o=(n.Title??"").trim().toLowerCase();o&&(e[o]=n.Id,a.push(o))}d.safeDomains=a,d.safeDomainIds=e}catch{}}async function Xe(t){const e=t.trim().toLowerCase();if(!(!e||d.savingDomain)){if(!f.canWhitelist()){f.toast("ต้องเป็น Agent ขึ้นไปจึงจะยืนยันโดเมนได้","error");return}d.savingDomain=e,f.rerender();try{const a=await f.getToken(),n=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items`,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:e,Category:Se})});if(!n.ok)throw new Error(String(n.status));await pe(),await ee(!0),f.toast(`ยืนยันแล้วว่า ${e} ปลอดภัย`)}catch{f.toast("บันทึกไม่สำเร็จ","error")}finally{d.savingDomain="",f.rerender()}}}async function et(t){const e=t.trim().toLowerCase(),a=d.safeDomainIds[e];if(!(!a||d.savingDomain)){if(!f.canWhitelist()){f.toast("ต้องเป็น Agent ขึ้นไปจึงจะถอนโดเมนได้","error");return}d.savingDomain=e,f.rerender();try{const n=await f.getToken(),o=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items(${a})`,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","IF-MATCH":"*","X-HTTP-Method":"DELETE"}});if(!o.ok)throw new Error(String(o.status));await pe(),await ee(!0),f.toast(`ถอน ${e} ออกจากรายการปลอดภัยแล้ว`)}catch{f.toast("ถอนไม่สำเร็จ","error")}finally{d.savingDomain="",f.rerender()}}}async function ee(t=!1){var p,w;const e=W(),a=(e==null?void 0:e.itemId)??"";if(!t&&d.analysedItemId===a&&d.analysis)return;d.loading=!0,d.reported=!1,d.analysedItemId=a,f.rerender();const[n,o]=await Promise.all([xe(Office.CoercionType.Html),xe(Office.CoercionType.Text)]),r={fromName:((p=e==null?void 0:e.from)==null?void 0:p.displayName)??"",fromEmail:((w=e==null?void 0:e.from)==null?void 0:w.emailAddress)??"",replyTo:[],subject:(e==null?void 0:e.subject)??"",bodyHtml:n,bodyText:o,attachments:((e==null?void 0:e.attachments)??[]).map(l=>({name:l.name,size:l.size??0,isInline:!!l.isInline})),headers:{},internalDomains:f.internalDomains,internalPeople:[],safeDomains:d.safeDomains},s=l=>{var h;return{...r,headers:l,replyTo:We(((h=Object.entries(l).find(([j])=>j.toLowerCase()==="reply-to"))==null?void 0:h[1])??"")}},c=await Je();if(d.mail=s(c),d.analysis=te(d.mail),d.loading=!1,f.rerender(),f.account()){const[l,h]=await Promise.all([Ze(),Object.keys(c).length?Promise.resolve(c):Ye()]);d.mail={...s(h),internalPeople:l},d.analysis=te(d.mail),d.kasmTemplate||(d.kasmTemplate=await Qe()),d.safeDomains.length||(await pe(),d.safeDomains.length&&(d.mail={...d.mail,safeDomains:d.safeDomains},d.analysis=te(d.mail))),f.rerender()}}function Be(){const t=d.mail,e=d.analysis;return!t||!e?"":[`ผู้ส่ง: ${t.fromName} <${t.fromEmail}>`,`หัวข้อ: ${t.subject}`,t.replyTo.length?`Reply-To: ${t.replyTo.map(a=>a.email).join(", ")}`:"",`คะแนนความเสี่ยง: ${e.score} (${Ee[e.level].label})`,"","สิ่งที่ตรวจพบ:",...e.findings.map(a=>`- [${Ie[a.severity].label}] (${a.category}) ${a.title} — ${a.detail.replace(/\n/g," ")}`),"",e.links.length?"ลิงก์ในอีเมล:":"",...e.links.map(a=>`- ${a.href}${a.flags.length?`  ! ${a.flags.join(" / ")}`:""}`)].filter(a=>a!=="").join(`
`)}async function tt(t){try{const e=W();if(!(e!=null&&e.itemId))return!1;const a=Office.context.mailbox.convertToRestId(e.itemId,Office.MailboxEnums.RestVersion.v2_0),n=await f.getGraphToken(),o=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${a}/$value`,{headers:{Authorization:`Bearer ${n}`}});if(!o.ok)return!1;const r=await o.arrayBuffer(),s=(e.subject||"phishing").replace(/[\\/:*?"<>|#%&{}~]/g,"_").replace(/^_+/,"").slice(0,80).trim()||"phishing",c=await f.getToken();return(await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('${Ce}')/items(${t})/AttachmentFiles/add(FileName='${encodeURIComponent(s+".eml")}')`,{method:"POST",headers:{Authorization:`Bearer ${c}`,Accept:"application/json;odata=nometadata"},body:r})).ok}catch{return!1}}async function ve(t){const e=await f.getToken(),a=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('${Ce}')/items`,{method:"POST",headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(t)});if(!a.ok)throw new Error(`SharePoint ${a.status}: ${await a.text()}`);return(await a.json()).Id}async function at(){if(!(!d.mail||!d.analysis||d.reporting)){if(!f.account()){f.toast("กรุณาเข้าสู่ระบบก่อนรายงาน","error");return}d.reporting=!0,f.rerender();try{const t=d.mail,e=d.analysis,a=f.account(),n={Title:(t.subject||"(ไม่มีหัวข้อ)").slice(0,255),SenderName:t.fromName.slice(0,255),SenderEmail:t.fromEmail.slice(0,255),SenderDomain:Ne(t.fromEmail),RiskScore:e.score,RiskLevel:e.level,Findings:Be(),LinkCount:e.links.length,SuspiciousLinks:e.links.filter(c=>c.flags.length).map(c=>c.href).join(`
`).slice(0,4e3),ReportedBy:(a==null?void 0:a.name)??"",ReportedEmail:(a==null?void 0:a.username)??"",Status:"New"};let o,r=!1;try{o=await ve(n)}catch(c){o=await ve({Title:n.Title,Findings:n.Findings}).catch(()=>{throw c}),r=!0}const s=await tt(o);d.reported=!0,f.toast(r?"ส่งรายงานแล้ว แต่บันทึกได้บางคอลัมน์ — ตรวจชื่อคอลัมน์ใน HD_PhishingReports":s?"ส่งรายงานพร้อมอีเมลต้นฉบับแล้ว":"ส่งรายงานแล้ว (แนบ .eml ไม่ได้)",r?"info":"success")}catch(t){f.toast(`ส่งรายงานไม่สำเร็จ: ${t instanceof Error?t.message:String(t)}`,"error")}finally{d.reporting=!1,f.rerender()}}}const nt=()=>d.reported?"✓ รายงานแล้ว":d.reporting?"กำลังส่ง…":"🚩 รายงานอีเมลนี้ให้ IT";function st(t){var e;try{const a=(e=Office.context)==null?void 0:e.ui;if(typeof(a==null?void 0:a.openBrowserWindow)=="function"){a.openBrowserWindow(t);return}}catch{}window.open(t,"_blank","noopener,noreferrer")||f.toast("เปิดหน้าต่างไม่ได้ (ถูกบล็อก)","info")}async function ot(t){var e;try{if((e=navigator.clipboard)!=null&&e.writeText)return await navigator.clipboard.writeText(t),!0}catch{}try{const a=document.createElement("textarea");a.value=t,a.setAttribute("readonly",""),a.style.position="fixed",a.style.opacity="0",document.body.appendChild(a),a.focus(),a.select();const n=document.execCommand("copy");return a.remove(),n}catch{return!1}}function it(t){if(d.loading&&!d.analysis)return`<div class="py-10 text-center text-slate-500 text-sm">
      <div class="w-7 h-7 border-[3px] border-slate-200 border-t-blue-600 rounded-full animate-spin mx-auto mb-2"></div>
      กำลังตรวจอีเมล…</div>`;const e=d.analysis,a=d.mail;if(!e||!a)return'<p class="text-sm text-slate-400 text-center py-8">เปิดอีเมลเพื่อเริ่มตรวจ</p>';const n=Ee[e.level],o=[...e.links.filter(s=>s.flags.length),...e.links.filter(s=>!s.flags.length&&!s.trusted),...e.links.filter(s=>s.trusted)],r=f.canWhitelist();return`
    <div class="rounded-xl border-2 ${n.cls} p-3">
      <div class="flex items-center gap-2">
        <span class="text-2xl leading-none">${n.icon}</span>
        <div class="min-w-0 flex-1">
          <div class="font-bold text-sm">${F(n.label)}</div>
          <div class="text-xs opacity-80">คะแนน ${e.score} · พบสัญญาณ ${e.findings.filter(s=>s.severity!=="info").length} ข้อ</div>
        </div>
        <button id="phish-recheck" class="text-[10px] px-2 py-1 rounded-md bg-white/70 hover:bg-white text-slate-700 font-medium">ตรวจใหม่</button>
      </div>
    </div>

    ${t?"":`<div class="bg-blue-50 border border-blue-200 rounded-xl p-2.5 text-[11px] text-blue-800">
      เข้าสู่ระบบเพื่อตรวจการปลอมเป็นพนักงาน และรายงานเข้า Helpdesk ได้
    </div>`}

    ${e.findings.length===0?'<p class="text-xs text-slate-400 text-center py-3">ไม่พบสัญญาณผิดปกติจากการตรวจอัตโนมัติ</p>':e.findings.map(s=>{const c=Ie[s.severity];return`<div class="bg-white rounded-xl border border-slate-200 p-2.5">
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
        ${o.map((s,c)=>{const p=d.savingDomain===(s.host?s.host.toLowerCase():"");return`
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
      <button id="phish-headers" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">${d.showHeaders?"ซ่อน header":"ดู header"}</button>
    </div>
    ${d.showHeaders?`<pre class="bg-slate-900 text-slate-100 text-[9px] p-2 rounded-lg overflow-x-auto whitespace-pre-wrap break-all max-h-56 overflow-y-auto">${F(Object.entries(a.headers).map(([s,c])=>`${s}: ${c}`).join(`
`)||"อ่าน header ไม่ได้")}</pre>`:""}
    <p class="text-[10px] text-slate-400 text-center">วิเคราะห์ในเครื่องทั้งหมด — ไม่ส่งเนื้อหาอีเมลออกนอกองค์กร</p>
  `}function rt(){var a,n,o,r;(a=document.getElementById("phish-recheck"))==null||a.addEventListener("click",()=>{ee(!0)}),(n=document.getElementById("phish-headers"))==null||n.addEventListener("click",()=>{d.showHeaders=!d.showHeaders,f.rerender()}),(o=document.getElementById("phish-copy"))==null||o.addEventListener("click",async()=>{const s=await ot(Be());f.toast(s?"คัดลอกผลตรวจแล้ว":"คัดลอกไม่ได้",s?"success":"error")}),document.querySelectorAll("[data-trust]").forEach(s=>s.addEventListener("click",()=>Xe(s.dataset.trust??""))),document.querySelectorAll("[data-untrust]").forEach(s=>s.addEventListener("click",()=>et(s.dataset.untrust??"")));const t=((r=d.analysis)==null?void 0:r.links)??[],e=[...t.filter(s=>s.flags.length),...t.filter(s=>!s.flags.length&&!s.trusted),...t.filter(s=>s.trusted)];document.querySelectorAll("[data-kasm]").forEach(s=>{s.addEventListener("click",()=>{const c=e[Number(s.dataset.kasm)];if(!c)return;if(!d.kasmTemplate){f.toast("ยังไม่ได้ตั้งค่า Kasm ใน HD_Options (Category=KasmConfig)","info");return}const p=d.kasmTemplate;st(p.includes("{url}")?p.replace("{url}",encodeURIComponent(c.href)):p+encodeURIComponent(c.href))})})}const ct="0bab07cf-65e6-487c-89af-c917fc1a5a13",lt="d569b991-89fc-4a62-9df5-eb361abcef40",M="https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd",X="https://rpaexpert.sharepoint.com/.default",ne=["https://graph.microsoft.com/Calendars.ReadWrite","https://graph.microsoft.com/Mail.Send","https://graph.microsoft.com/Mail.Read"],L=new He({auth:{clientId:ct,authority:`https://login.microsoftonline.com/${lt}`,redirectUri:window.location.origin.includes("localhost")?"http://localhost:3000/":"https://darmmunginsa.github.io/itservices-addin/",navigateToLoginRequestUrl:!1},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:!1}}),dt=window.location.origin.includes("localhost")?`${window.location.origin}/auth.html`:"https://darmmunginsa.github.io/itservices-addin/auth.html";function fe(){var e,a;const t=(a=(e=Office.context)==null?void 0:e.diagnostics)==null?void 0:a.platform;return t===Office.PlatformType.iOS||t===Office.PlatformType.Android}function he(){return new Promise((t,e)=>{Office.context.ui.displayDialogAsync(dt,{height:60,width:30,promptBeforeOpen:!1},a=>{if(a.status!==Office.AsyncResultStatus.Succeeded){e(new Error("เปิดหน้าเข้าสู่ระบบไม่ได้"));return}const n=a.value;n.addEventHandler(Office.EventType.DialogMessageReceived,o=>{n.close();const r=o.message;if(!r){e(new Error("auth message error"));return}try{const s=JSON.parse(r);s.ok?t():e(new Error(s.error||"auth failed"))}catch{e(new Error("auth message error"))}}),n.addEventHandler(Office.EventType.DialogEventReceived,()=>e(new Error("ปิดหน้าเข้าสู่ระบบก่อนเสร็จ")))})})}const i={account:null,tab:"ticket",emailSubject:"",emailBodyPreview:"",emailBodyReply:"",emailSenderName:"",emailSenderEmail:"",loading:!1,projects:[],agents:[],myRole:"",emailAttachments:[],signatureContact:null,droppedFiles:[],tickets:[],contactEmails:[],emailCc:[]};async function H(){const t=L.getAllAccounts();if(t.length===0)throw new Error("Not signed in");const e={scopes:[X],account:t[0]};try{return(await L.acquireTokenSilent(e)).accessToken}catch{if(fe()){await he();const a=L.getAllAccounts()[0];if(!a)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await L.acquireTokenSilent({scopes:[X],account:a})).accessToken}return(await L.acquireTokenPopup(e)).accessToken}}async function K(t=!1){const e=L.getAllAccounts();if(e.length===0)throw new Error("Not signed in");const a={scopes:ne,account:e[0],forceRefresh:t};try{return(await L.acquireTokenSilent(a)).accessToken}catch{if(fe()){await he();const o=L.getAllAccounts()[0];if(!o)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await L.acquireTokenSilent({scopes:ne,account:o})).accessToken}return(await L.acquireTokenPopup({scopes:ne,account:e[0]})).accessToken}}async function mt(t){const e=await K(),a={subject:t.subject,start:{dateTime:t.start,timeZone:"Asia/Bangkok"},end:{dateTime:t.end,timeZone:"Asia/Bangkok"},body:t.body?{contentType:"HTML",content:t.body.replace(/\n/g,"<br>")}:void 0,attendees:t.attendees.filter(Boolean).map(o=>({emailAddress:{address:o},type:"required"})),isOnlineMeeting:t.isOnlineMeeting,onlineMeetingProvider:t.isOnlineMeeting?"teamsForBusiness":void 0},n=await fetch("https://graph.microsoft.com/v1.0/me/events",{method:"POST",headers:{Authorization:`Bearer ${e}`,"Content-Type":"application/json"},body:JSON.stringify(a)});if(!n.ok)throw new Error(`Calendar error ${n.status}: ${await n.text()}`)}async function je(){try{const t=await H(),e=`${M}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$orderby=Title asc&$top=500`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();i.projects=n.value.map(o=>({id:o.Id,Title:o.Title}))}}catch{}}async function Pe(){var t,e;try{const a=await H(),n=`${M}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText,Role&$orderby=Title asc`,o=await fetch(n,{headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata"}});if(o.ok){const r=await o.json();i.agents=r.value.map(c=>({email:c.EmailText,name:c.Title}));const s=(((t=i.account)==null?void 0:t.username)??"").toLowerCase();i.myRole=((e=r.value.find(c=>(c.EmailText??"").toLowerCase()===s))==null?void 0:e.Role)??""}}catch{}}async function De(){try{const t=await H(),e=`${M}/_api/web/lists/getbytitle('HD_Tickets')/items?$select=Id,Title,TicketNumber,Status&$filter=Status ne 'Closed'&$orderby=Modified desc&$top=200`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();i.tickets=n.value.map(o=>({id:o.Id,Title:o.Title,TicketNumber:o.TicketNumber,Status:o.Status}))}}catch{}}async function Oe(){try{const t=await H(),e=`${M}/_api/web/lists/getbytitle('HD_Contracts')/items?$select=CustomerEmail&$top=2000`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();i.contactEmails=n.value.map(o=>(o.CustomerEmail||"").trim().toLowerCase()).filter(Boolean)}}catch{}}async function $e(){const t=document.getElementById("btn-login-main"),e=document.getElementById("btn-login");t&&(t.disabled=!0,t.textContent="กำลังเข้าสู่ระบบ…"),e&&(e.disabled=!0);try{if(fe()){if(await he(),i.account=L.getAllAccounts()[0]??null,!i.account)throw new Error("เข้าสู่ระบบไม่สำเร็จ")}else{const a=await L.loginPopup({scopes:[X]});i.account=a.account}await Promise.all([je(),Pe(),De(),Oe()]),z()}catch{t&&(t.disabled=!1,t.textContent="เข้าสู่ระบบ"),e&&(e.disabled=!1)}}async function ut(){i.account&&await L.logoutPopup({account:i.account}),i.account=null,z()}async function q(t,e){const a=await H(),n=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items`,o=await fetch(n,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(e)});if(!o.ok){const s=await o.text();throw new Error(`SharePoint error ${o.status}: ${s}`)}return(await o.json()).Id}let J=null;const se="support@itservices.co.th",Le="engineer@itservices.co.th";async function _e(){if(J)return J;try{const t=await H(),e=`${M}/_api/web/lists/getbytitle('HD_EmailTemplates')/items?$select=EventKey,Subject,Body,IsEnabled&$top=50`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});return a.ok?(J=(await a.json()).value,J):[]}catch{return[]}}async function pt(){var t,e;try{const a=await H(),n=`${M}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title,Category&$filter=Category eq 'EmailConfig'&$top=1`,o=await fetch(n,{headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata"}});return o.ok&&((e=(t=(await o.json()).value[0])==null?void 0:t.Title)==null?void 0:e.trim())||se}catch{return se}}function re(t,e){return t.replace(/\{\{(\w+)\}\}/g,(a,n)=>e[n]??`{{${n}}}`)}async function ft(t,e,a,n=[]){try{const r=(await _e()).find(m=>m.EventKey===t&&m.IsEnabled);if(!r)return;const s=re(r.Subject||"",e),c=re(r.Body||"",e);if(!s||!c)return;const p=m=>m.trim().toLowerCase(),w=[...new Map(a.filter(Boolean).map(m=>[p(m),m])).values()];if(w.length===0)return;const l=new Set(w.map(p)),h=t==="ticket_created"?[...n,Le]:n,j=[...new Map(h.filter(Boolean).map(m=>[p(m),m])).values()].filter(m=>!l.has(p(m))),O=await pt(),N=await K(),T={subject:s,body:{contentType:"HTML",content:c},toRecipients:w.map(m=>({emailAddress:{address:m}}))};j.length&&(T.ccRecipients=j.map(m=>({emailAddress:{address:m}}))),O&&(T.from={emailAddress:{address:O}}),await fetch("https://graph.microsoft.com/v1.0/me/sendMail",{method:"POST",headers:{Authorization:`Bearer ${N}`,"Content-Type":"application/json"},body:JSON.stringify({message:T,saveToSentItems:!0})})}catch{}}async function ht(t,e=[]){try{const a=Office.context.mailbox.item;if(!(a!=null&&a.itemId))return!1;const n=Office.context.mailbox.convertToRestId(a.itemId,Office.MailboxEnums.RestVersion.v2_0),r={Authorization:`Bearer ${await K()}`,"Content-Type":"application/json"},s=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${n}/createReplyAll`,{method:"POST",headers:r});if(!s.ok)return!1;const c=await s.json(),p=T=>T.trim().toLowerCase(),w=c.ccRecipients??[],l=new Set(w.map(T=>p(T.emailAddress.address))),h=[...new Set(e.filter(Boolean).map(T=>T.trim()))].filter(T=>!l.has(p(T))).map(T=>({emailAddress:{address:T}})),j={body:{contentType:"HTML",content:t}};return h.length&&(j.ccRecipients=[...w,...h]),(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${c.id}`,{method:"PATCH",headers:r,body:JSON.stringify(j)})).ok?(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${c.id}/send`,{method:"POST",headers:r})).ok:!1}catch{return!1}}async function gt(t,e){const n=(await _e()).find(r=>r.EventKey===t&&r.IsEnabled);return n&&re(n.Body||"",e)||null}async function Y(t){var r;const e=s=>s.trim().toLowerCase(),a=e(((r=i.account)==null?void 0:r.username)??""),n=new Set,o=t.recipients.filter(Boolean).filter(s=>{const c=e(s);return!c||c===a||n.has(c)?!1:(n.add(c),!0)});if(o.length!==0)try{const s=await H(),c=`${M}/_api/web/lists/getbytitle('HD_Notifications')/items`;await Promise.all(o.map(p=>fetch(c,{method:"POST",headers:{Authorization:`Bearer ${s}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:t.title.slice(0,255),RecipientEmail:p,EventType:t.eventType,Message:t.message,LinkPath:t.linkPath,IsRead:!1})})))}catch{}}async function ce(t,e){const a=document.querySelectorAll(".email-att-cb:checked");if(a.length===0)return;const n=await H(),o=new Set,r=s=>{if(!o.has(s.toLowerCase()))return o.add(s.toLowerCase()),s;const c=s.lastIndexOf("."),p=c>0?s.slice(0,c):s,w=c>0?s.slice(c):"";for(let l=2;;l++){const h=p+"-"+l+w;if(!o.has(h.toLowerCase()))return o.add(h.toLowerCase()),h}};for(const s of Array.from(a)){const c=s.dataset.attId,p=r(s.dataset.attName),w=await new Promise((m,g)=>{Office.context.mailbox.item.getAttachmentContentAsync(c,{},$=>{$.status===Office.AsyncResultStatus.Succeeded?m($):g(new Error($.error.message))})}),{content:l,format:h}=w.value;let j;if(h===Office.MailboxEnums.AttachmentContentFormat.Base64){const m=atob(l);j=new Uint8Array(m.length);for(let g=0;g<m.length;g++)j[g]=m.charCodeAt(g)}else if(h===Office.MailboxEnums.AttachmentContentFormat.Eml||h===Office.MailboxEnums.AttachmentContentFormat.ICalendar)j=new TextEncoder().encode(l);else continue;const O=encodeURIComponent(p),N=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${O}')`;if(!(await fetch(N,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:j.buffer})).ok)throw new Error(`Upload ${p} failed`)}}async function bt(t){const e=`https://graph.microsoft.com/v1.0/me/messages/${t}/$value`;let a=await K(),n=await fetch(e,{headers:{Authorization:`Bearer ${a}`}});if((n.status===401||n.status===403)&&(a=await K(!0),n=await fetch(e,{headers:{Authorization:`Bearer ${a}`}})),!n.ok)throw new Error(`Graph ${n.status}`);return n.arrayBuffer()}async function yt(t){const e=await new Promise((n,o)=>{Office.context.mailbox.getCallbackTokenAsync({isRest:!0},r=>{r.status===Office.AsyncResultStatus.Succeeded?n(r.value):o(new Error("callback token failed"))})}),a=await fetch(`${Office.context.mailbox.restUrl}/v2.0/me/messages/${t}/$value`,{headers:{Authorization:`Bearer ${e}`}});if(!a.ok)throw new Error(`REST ${a.status}`);return a.arrayBuffer()}async function le(t,e){const a=document.getElementById("f-attach-eml");if(!(a!=null&&a.checked))return;const n=Office.context.mailbox.item;if(!n)return;const o=Office.context.mailbox.convertToRestId(n.itemId,Office.MailboxEnums.RestVersion.v2_0);let r,s="",c="";try{r=await bt(o)}catch(j){s=j instanceof Error?j.message:String(j);try{r=await yt(o)}catch(O){c=O instanceof Error?O.message:String(O),console.error("[eml] graph:",s,"| callback:",c),B(`ดึง .eml ไม่ได้ (Graph: ${s} / REST: ${c}) — ไฟล์อื่นบันทึกแล้ว`,"error");return}}const p=(n.subject||"email").replace(/[\\/:*?"<>|#%&{}~]/g,"_").slice(0,100).trim()||"email",w=await H(),l=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${encodeURIComponent(p+".eml")}')`;(await fetch(l,{method:"POST",headers:{Authorization:`Bearer ${w}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok||B("บันทึกไฟล์ .eml ไม่สำเร็จ","error")}async function de(t,e,a){const n=await H();for(const o of a){const r=await o.arrayBuffer(),s=encodeURIComponent(o.name),c=`${M}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${s}')`;if(!(await fetch(c,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok)throw new Error(`Upload ${o.name} failed`)}}function B(t,e="success"){const a=document.getElementById("toast-container");if(!a)return;const n=e==="success"?"bg-green-500":e==="error"?"bg-red-500":"bg-slate-700",o=e==="success"?"✅":e==="error"?"❌":"ℹ️",r=document.createElement("div");r.className=`toast pointer-events-auto ${n} text-white text-sm font-medium px-4 py-3 rounded-lg shadow-lg max-w-xs mx-2`,r.textContent=`${o} ${t}`,a.appendChild(r),setTimeout(()=>r.remove(),4e3)}function xt(t){const e=t.split(`
`).map(l=>l.trim()).filter(Boolean),a=/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/,n=/(\+?[\d\s()\-]{7,}(?:ext\.?\s*\d+)?)/i,o=/\b(co\.?,?\s*ltd\.?|co\.?,?\s*limited|corporation|corp\.?|บริษัท|จำกัด|holding|group|inc\.?|llc)\b/i;let r="",s="",c="";const p=[];for(const l of e)if(!/^[-_=*]{2,}$/.test(l)&&!/^(best regards|regards|sincerely|ขอแสดงความนับถือ|ด้วยความนับถือ|from|sent|thanks|thank you)/i.test(l)){if(!r){const h=l.match(a);if(h){r=h[0];continue}}if(!s){const h=l.match(n);if(h&&h[0].replace(/\D/g,"").length>=7){s=h[0].trim();continue}}if(!c&&o.test(l)){c=l;continue}l.length>=2&&l.length<=50&&!/\d{4,}/.test(l)&&p.push(l)}const w=p.find(l=>!a.test(l)&&!o.test(l))??"";return!r&&!w?null:{name:w,company:c,email:r,phone:s}}async function vt(){const t=i.signatureContact;if(!t)return;const e=(i.emailSenderEmail||"").toLowerCase();if(e&&i.contactEmails.includes(e)){B("ลูกค้านี้มีในระบบแล้ว","success"),i.signatureContact=null,z();return}const a=document.getElementById("btn-import-customer");a&&(a.disabled=!0,a.textContent="กำลังบันทึก…");try{await q("HD_Contracts",{Title:i.emailSenderName||t.name,CustomerEmail:i.emailSenderEmail,Phone:t.phone||void 0,Company:t.company||void 0,Status:"Active"}),e&&i.contactEmails.push(e),B("เพิ่มลูกค้าสำเร็จ!"),i.signatureContact=null,z()}catch(n){const o=n instanceof Error?n.message:String(n);B(`เกิดข้อผิดพลาด: ${o}`,"error"),a&&(a.disabled=!1,a.textContent="เพิ่มเป็นลูกค้า")}}function Re(){return new Date().toISOString().split("T")[0]}function $t(){const t=new Date;return`HD-${`${t.getFullYear()}${String(t.getMonth()+1).padStart(2,"0")}${String(t.getDate()).padStart(2,"0")}`}-${Math.floor(Math.random()*900+100)}`}function we(){var t;return i.droppedFiles.length>0||document.querySelectorAll(".email-att-cb:checked").length>0||(((t=document.getElementById("f-attach-eml"))==null?void 0:t.checked)??!1)}async function Z(t,e){i.droppedFiles.length>0&&await de(t,e,i.droppedFiles),await ce(t,e),await le(t,e)}let oe=!1;async function wt(){var e,a,n,o,r,s,c,p,w,l,h,j,O,N,T;if(!i.account){B("กรุณาเข้าสู่ระบบก่อน","error");return}if(oe)return;oe=!0;const t=document.getElementById("submit-btn");t&&(t.disabled=!0,t.textContent="กำลังบันทึก…");try{if(i.tab==="phish")await at();else if(i.tab==="ticket"){const m=document.getElementById("f-title").value.trim(),g=document.getElementById("f-description").value.trim(),$=document.getElementById("f-priority").value,E=document.getElementById("f-customer-email").value.trim(),x=((e=document.getElementById("f-cc-enable"))==null?void 0:e.checked)??!0?(((a=document.getElementById("f-cc"))==null?void 0:a.value)||"").split(/[,;\s]+/).map(I=>I.trim()).filter(Boolean):[],C=document.getElementById("f-assigned-email").value,S=i.agents.find(I=>I.email===C),D=$t(),P=await q("HD_Tickets",{Title:m,TicketNumber:D,Description:g,Priority:$,CustomerEmail:E,CustomerName:i.emailSenderName||E,Status:"Open",AssignedEmail:C||void 0,AssignedToName:(S==null?void 0:S.name)??((n=i.account)==null?void 0:n.name)??"",ProjectID:parseInt(((o=document.getElementById("f-project"))==null?void 0:o.value)||"0")||null});if(we()){const I=await q("HD_TicketComments",{Title:"📎 ไฟล์แนบจากอีเมล",TicketID:P,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await Z("HD_TicketComments",I)}i.droppedFiles=[];const R={ticket_number:D,ticket_title:m,priority:$,category:"-",description:(g||"-").replace(/\n/g,"<br>"),customer_name:i.emailSenderName||E,assigned_name:(S==null?void 0:S.name)??((r=i.account)==null?void 0:r.name)??"-",link:"https://itservices.co.th/helpdesk/"},u=[C,i.account.username,...x,Le].filter(Boolean);let y=!1;const _=await gt("ticket_created",R);if(_){const I=`<div style="border-left:4px solid #2563eb;background:#eff6ff;padding:10px 14px;margin:0 0 14px;font-family:Segoe UI,sans-serif">
             <div style="font-size:15px;font-weight:700;color:#1e40af">Ticket No. ${D}</div>
             <div style="font-size:12px;color:#475569;margin-top:2px">กรุณาตอบกลับในอีเมลฉบับนี้เพื่อให้ข้อมูลอยู่ใน Ticket เดียวกัน</div>
           </div>`;y=await ht(I+_,u)}y||await ft("ticket_created",R,[E],u),B(y?"สร้าง Ticket และตอบกลับในเธรดเดิมแล้ว!":"สร้าง Ticket สำเร็จ!")}else if(i.tab==="task"){const m=document.getElementById("f-title").value.trim(),g=parseInt(((s=document.getElementById("f-project"))==null?void 0:s.value)||"0"),$=document.getElementById("f-due-date").value,E=document.getElementById("f-note").value.trim(),A=document.getElementById("f-assigned-email").value,x=i.agents.find(D=>D.email===A);if(!g){B("กรุณาเลือก Project","error");return}const C=await q("PM_Tasks",{Title:m,DueDate:$||null,TaskNote:E,AssignedTo:(x==null?void 0:x.name)??i.account.name??i.account.username,AssignedEmail:A,IsCompleted:!1,IsAcknowledged:!1,ProjectID:g});if(i.droppedFiles.length>0&&await de("PM_Tasks",C,i.droppedFiles),await ce("PM_Tasks",C),await le("PM_Tasks",C),i.droppedFiles=[],await Y({recipients:[A],title:`📋 ได้รับมอบหมาย Task: ${m}`,message:E||($?`กำหนดส่ง ${$}`:"มี Task ใหม่"),linkPath:g?`/projects/${g}`:"/my-work",eventType:"task_assigned"}),((c=document.getElementById("f-teams"))==null?void 0:c.checked)&&$){const D=Array.from(document.querySelectorAll(".att-internal:checked")).map(y=>y.value),P=(((p=document.getElementById("f-ext-att"))==null?void 0:p.value)||"").split(/[,;\s]+/).map(y=>y.trim()).filter(Boolean),R=`${$}T09:00:00`,u=`${$}T10:00:00`;try{await mt({subject:m,start:R,end:u,body:E,attendees:[...D,...P],isOnlineMeeting:!0}),B("สร้าง Task + นัดประชุม Teams สำเร็จ!")}catch(y){B("สร้าง Task แล้ว แต่สร้างนัดประชุมไม่สำเร็จ: "+(y instanceof Error?y.message:""),"error")}}else B("สร้าง Task สำเร็จ!")}else if(i.tab==="incident"){const m=document.getElementById("f-title").value.trim(),g=parseInt(((w=document.getElementById("f-project"))==null?void 0:w.value)||"0"),$=document.getElementById("f-description").value.trim(),E=document.getElementById("f-severity").value,A=document.getElementById("f-assigned-email").value,x=i.agents.find(u=>u.email===A),C=document.getElementById("f-status").value,S=document.getElementById("f-incident-date").value,D=document.getElementById("f-resolution").value.trim();if(!g){B("กรุณาเลือก Project","error");return}const P=parseInt(((l=document.getElementById("f-sla"))==null?void 0:l.value)||"0")||null,R=await q("PM_Incidents",{Title:m,Description:$||void 0,Severity:E,Status:C,AssignedTo:(x==null?void 0:x.name)??i.account.name??i.account.username,AssignedEmail:A,ProjectID:g,IncidentDate:S||Re(),Resolution:D||void 0,SLAHours:P,SLADue:Ae(P),...C==="Resolved"?{ResolvedDate:new Date().toISOString()}:{}});i.droppedFiles.length>0&&await de("PM_Incidents",R,i.droppedFiles),await ce("PM_Incidents",R),await le("PM_Incidents",R),i.droppedFiles=[],await Y({recipients:[A],title:`🚨 ได้รับมอบหมาย Incident: ${m}`,message:`ความรุนแรง ${E}${$?" — "+$.slice(0,120):""}`,linkPath:g?`/projects/${g}`:"/my-work",eventType:"incident_created"}),B("สร้าง Incident สำเร็จ!")}else if(i.tab==="comment"){const m=parseInt(((h=document.getElementById("f-ticket"))==null?void 0:h.value)||"0"),g=document.getElementById("f-comment").value.trim(),$=document.getElementById("f-comment-type").value;if(!m){B("กรุณาเลือก Ticket","error");return}if(!g){B("กรุณาพิมพ์ Comment","error");return}const E=await q("HD_TicketComments",{Title:g.slice(0,100),TicketID:m,CommentText:g,CommentType:$,CommentDate:new Date().toISOString()});await Z("HD_TicketComments",E),i.droppedFiles=[];try{const A=await H(),x=`${M}/_api/web/lists/getbytitle('HD_Tickets')/items(${m})?$select=TicketNumber,Title,AssignedEmail,Author/EMail&$expand=Author`,C=await fetch(x,{headers:{Authorization:`Bearer ${A}`,Accept:"application/json;odata=nometadata"}});if(C.ok){const S=await C.json(),D=i.account.username.toLowerCase(),P=[...new Set([S.AssignedEmail,(j=S.Author)==null?void 0:j.EMail].filter(Boolean))].filter(R=>R.toLowerCase()!==D);P.length&&await Y({recipients:P,title:`💬 ${((O=i.account)==null?void 0:O.name)??"มีคน"} คอมเมนต์ใน ${S.TicketNumber||"#"+m}`,message:g.slice(0,200),linkPath:`/tickets/${m}`,eventType:"comment_added"})}}catch{}B("เพิ่ม Comment สำเร็จ!")}else if(i.tab==="project"){const m=document.getElementById("f-title").value.trim(),g=document.getElementById("f-company").value.trim(),$=document.getElementById("f-group").value,E=document.getElementById("f-status").value,A=document.getElementById("f-start").value,x=document.getElementById("f-end").value,C=document.getElementById("f-description").value.trim();if(!m){B("กรุณาใส่ชื่อโครงการ","error");return}const S=await q("PM_Projects",{Title:m,Company:g||void 0,ProjectGroup:$,Progress:0,StartDate:A||void 0,EndDate:x||null,Status:E,CreatedByEmail:i.account.username,Comment:C||void 0});if(we()){const D=await q("PM_Comments",{Title:"📎 ไฟล์แนบจากอีเมล",ProjectID:S,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await Z("PM_Comments",D)}i.droppedFiles=[],B("สร้างโครงการสำเร็จ!")}else if(i.tab==="projcomment"){const m=parseInt(((N=document.getElementById("f-project"))==null?void 0:N.value)||"0"),g=document.getElementById("f-comment").value.trim(),$=document.getElementById("f-comment-type").value;if(!m){B("กรุณาเลือกโครงการ","error");return}if(!g){B("กรุณาพิมพ์ Comment","error");return}const E=await q("PM_Comments",{Title:g.slice(0,100),ProjectID:m,CommentText:g,CommentType:$,CommentDate:new Date().toISOString(),ParentID:0});await Z("PM_Comments",E),i.droppedFiles=[];try{const A=await H(),x=`${M}/_api/web/lists/getbytitle('PM_Projects')/items(${m})?$select=Title,CreatedByEmail`,C=await fetch(x,{headers:{Authorization:`Bearer ${A}`,Accept:"application/json;odata=nometadata"}});if(C.ok){const S=await C.json(),D=i.account.username.toLowerCase();S.CreatedByEmail&&S.CreatedByEmail.toLowerCase()!==D&&await Y({recipients:[S.CreatedByEmail],title:`💬 ${((T=i.account)==null?void 0:T.name)??"มีคน"} คอมเมนต์ในโครงการ ${S.Title??""}`,message:g.slice(0,200),linkPath:`/projects/${m}?tab=comments`,eventType:"comment_added"})}}catch{}B("เพิ่ม Comment สำเร็จ!")}}catch(m){const g=m instanceof Error?m.message:String(m);B(`เกิดข้อผิดพลาด: ${g}`,"error")}finally{oe=!1,t&&(t.disabled=!1,t.textContent="บันทึก")}}const Tt={ticket:{label:"Ticket",icon:"🎫"},comment:{label:"Comment",icon:"💬"},project:{label:"Project",icon:"📁"},task:{label:"Task",icon:"✅"},incident:{label:"Incident",icon:"🚨"},projcomment:{label:"Comment",icon:"💬"},phish:{label:"PhishGuard",icon:"🛡️"}},kt=[{title:"🎫 Helpdesk",tabs:["ticket","comment"]},{title:"📁 Project",tabs:["project","task","incident","projcomment"]},{title:"🛡️ Security",tabs:["phish"]}],Me=["f-title","f-description","f-priority","f-customer-email","f-cc","f-assigned-email","f-project","f-due-date","f-note","f-severity","f-status","f-sla","f-incident-date","f-resolution","f-ticket","f-comment","f-comment-type","f-company","f-group","f-start","f-end","f-ext-att","f-attach-eml"];let G={};function Et(){for(const e of Me){const a=document.getElementById(e);a&&(G[e]=a.value)}const t=document.getElementById("f-teams");t&&(G["f-teams"]=t.checked)}function It(){for(const e of Me){const a=document.getElementById(e);a&&G[e]!==void 0&&G[e]!==""&&(a.value=G[e])}const t=document.getElementById("f-teams");if(t&&G["f-teams"]!==void 0){t.checked=G["f-teams"];const e=document.getElementById("teams-fields");e&&(e.style.display=t.checked?"block":"none")}}function z(){var x,C,S,D,P,R;const t=document.getElementById("app");if(!t)return;Et();const{account:e,tab:a,emailSubject:n,emailSenderName:o,emailSenderEmail:r,emailBodyPreview:s}=i,c=i.emailBodyReply||s,p=e!==null,w=`
    <div class="bg-gradient-to-r from-blue-700 to-blue-600 text-white px-4 py-2.5 flex items-center gap-2.5 shadow flex-shrink-0">
      <div class="w-7 h-7 bg-white/95 rounded-lg flex items-center justify-center flex-shrink-0">
        <span class="text-blue-700 font-extrabold text-xs">iT</span>
      </div>
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm leading-tight truncate">iT Services Helpdesk</div>
        ${p?`<div class="text-[10px] text-blue-100 truncate">${b((e==null?void 0:e.name)??(e==null?void 0:e.username)??"")}</div>`:""}
      </div>
      ${p?`<button id="btn-logout" title="ออกจากระบบ" class="p-1.5 rounded-lg hover:bg-white/15 transition flex-shrink-0">
             <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
           </button>`:`<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded-lg transition flex-shrink-0">
             เข้าสู่ระบบ
           </button>`}
    </div>
  `;if(!p){t.innerHTML=`
      ${w}
      <div class="flex flex-col items-center justify-center px-6 py-16 gap-4 text-center">
        <div class="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center text-3xl">🔐</div>
        <h2 class="text-base font-semibold text-slate-700">เข้าสู่ระบบด้วย Microsoft</h2>
        <p class="text-xs text-slate-500">เพื่อสร้าง Ticket / Task / Incident จาก Email นี้</p>
        <button id="btn-login-main"
          class="mt-2 w-full bg-blue-700 hover:bg-blue-600 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          เข้าสู่ระบบ
        </button>
      </div>
    `,(x=document.getElementById("btn-login"))==null||x.addEventListener("click",$e),(C=document.getElementById("btn-login-main"))==null||C.addEventListener("click",$e);return}const l=n?`<div class="mx-3 mt-3 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5 text-xs text-slate-700 space-y-1">
         <div class="font-semibold text-blue-800 truncate" title="${b(n)}">📧 ${b(n)}</div>
         ${o?`<div class="text-slate-500">จาก: <span class="font-medium text-slate-700">${b(o)}</span></div>`:""}
         ${r&&r!==o?`<div class="text-slate-400 truncate">${b(r)}</div>`:""}
       </div>`:`<div class="mx-3 mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-700">
         ⚠️ ไม่พบข้อมูล Email (โหมดทดสอบ)
       </div>`,h=i.signatureContact,j=!!r&&i.contactEmails.includes(r.toLowerCase()),O=h?`<div class="mx-3 mt-3 bg-orange-50 border border-orange-200 rounded-lg px-3 py-2.5 text-xs text-slate-700">
         <div class="flex items-center justify-between mb-2">
           <span class="font-semibold text-orange-700">👤 ข้อมูลผู้ส่ง (จาก Signature)</span>
         </div>
         <div class="space-y-0.5 mb-2.5">
           ${o?`<div><span class="text-slate-400">ชื่อ:</span> <span class="font-medium">${b(o)}</span></div>`:""}
           ${h.company?`<div><span class="text-slate-400">บริษัท:</span> ${b(h.company)}</div>`:""}
           ${r?`<div><span class="text-slate-400">Email:</span> ${b(r)}</div>`:""}
           ${h.phone?`<div><span class="text-slate-400">โทร:</span> ${b(h.phone)}</div>`:""}
         </div>
         ${j?'<div class="w-full bg-green-100 text-green-700 text-xs font-semibold py-1.5 rounded-md text-center">✓ ลูกค้านี้มีในระบบแล้ว</div>':`<button id="btn-import-customer"
                class="w-full bg-orange-500 hover:bg-orange-400 text-white text-xs font-semibold py-1.5 rounded-md transition">
                + เพิ่มเป็นลูกค้า
              </button>`}
       </div>`:"",N=`
    <div class="mx-3 mt-3 space-y-2">
      ${kt.map(u=>`
        <div>
          <div class="text-[10px] font-semibold text-slate-400 uppercase tracking-wide mb-1 px-0.5">${u.title}</div>
          <div class="grid grid-cols-4 gap-1">
            ${u.tabs.map(y=>{const _=Tt[y];return`<button data-tab="${y}"
                class="tab-btn flex flex-col items-center gap-1 py-2 rounded-lg transition ${a===y?"bg-blue-700 text-white shadow":"text-slate-500 hover:bg-slate-100"}">
                <span class="text-base leading-none">${_.icon}</span>
                <span class="text-[9px] font-medium leading-none">${_.label}</span>
              </button>`}).join("")}
          </div>
        </div>
      `).join("")}
    </div>
  `;let T="";a==="phish"?T=it(!!e):a==="ticket"?T=`
      ${k("Title / หัวข้อ",`<input id="f-title" type="text"
        class="${v}"
        value="${b(n)}" />`)}
      ${k("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${v} resize-none">${b(s)}</textarea>`)}
      ${k("Priority",`<select id="f-priority" class="${v}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
      ${k("Customer Email",`<input id="f-customer-email" type="email"
        class="${v}"
        value="${b(r)}" />`)}
      ${k("CC — ให้ผู้ที่อยู่ในเมลนี้รับรู้",`
        <label class="flex items-center gap-2 text-xs text-slate-600 mb-1.5 cursor-pointer">
          <input id="f-cc-enable" type="checkbox" ${i.emailCc.length?"checked":""} /> แนบผู้รับในเมลนี้เป็น CC อัตโนมัติ
        </label>
        <input id="f-cc" type="text" class="${v}" value="${b(i.emailCc.join(", "))}" placeholder="someone@company.com, boss@company.com" />`)}
      ${k("Assign ให้ Agent",ie(e.username))}
      ${k("โครงการ (ไม่บังคับ)",Q(!0))}
      ${V()}
    `:a==="task"?T=`
      ${k("ชื่อ Task *",`<input id="f-title" type="text" required
        class="${v}" value="${b(n)}" />`)}
      ${k("Project *",Q())}
      ${k("Assign ให้",ie(e.username))}
      ${k("Due Date",`<input id="f-due-date" type="date" class="${v}" />`)}
      ${k("Task Note",`<textarea id="f-note" rows="4"
        class="${v} resize-y">${b(s)}</textarea>`)}
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
          <input id="f-ext-att" type="text" class="${v}" placeholder="someone@company.com, ..." />
        </div>
      </div>
      ${V()}
    `:a==="incident"?T=`
      ${k("ชื่อ Incident *",`<input id="f-title" type="text" required
        class="${v}" value="${b(n)}" />`)}
      ${k("Project *",Q())}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">ความรุนแรง</label>
          <select id="f-severity" class="${v}">
            <option value="Low">Low</option>
            <option value="Medium" selected>Medium</option>
            <option value="High">High</option>
            <option value="Critical">Critical</option>
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${v}">
            <option value="Open" selected>Open</option>
            <option value="In Progress">In Progress</option>
            <option value="Resolved">Resolved</option>
          </select>
        </div>
      </div>
      ${k("SLA — ต้องแก้ให้จบภายใน",`
        <select id="f-sla" class="${v}">
          <option value="">ไม่กำหนด SLA</option>
          ${Ve.map(u=>`<option value="${u.hours}" ${u.hours===ae.Medium?"selected":""}>${u.labelTh}</option>`).join("")}
        </select>
        <p id="f-sla-hint" class="text-[11px] text-slate-400 mt-1">นับจากตอนนี้ · ครบกำหนด ${ye(ae.Medium)}</p>`)}
      ${k("Assign ให้ Agent",ie(e.username))}
      ${k("วันที่เกิด Incident",`<input id="f-incident-date" type="date" class="${v}" value="${Re()}" />`)}
      ${k("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${v} resize-y">${b(s)}</textarea>`)}
      ${k("วิธีแก้ไข (ถ้ามี)",`<textarea id="f-resolution" rows="2"
        class="${v} resize-y" placeholder="อธิบายวิธีแก้ไขปัญหา..."></textarea>`)}
      ${V()}
    `:a==="comment"?T=`
      ${k("เลือก Ticket *",`<select id="f-ticket" class="${v}">
        <option value="">-- เลือก Ticket ที่จะเพิ่ม Comment --</option>
        ${i.tickets.map(u=>`<option value="${u.id}">${b(u.TicketNumber||"#"+u.id)} · ${b(u.Title)}</option>`).join("")}
      </select>`)}
      ${k("ประเภท",`<select id="f-comment-type" class="${v}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${k("Comment *",`<textarea id="f-comment" rows="5"
        class="${v} resize-y" placeholder="พิมพ์ comment...">${b(c)}</textarea>`)}
      ${V()}
    `:a==="project"?T=`
      ${k("ชื่อโครงการ *",`<input id="f-title" type="text" required
        class="${v}" value="${b(n)}" />`)}
      ${k("บริษัท / ลูกค้า",`<input id="f-company" type="text" class="${v}" value="${b(((S=i.signatureContact)==null?void 0:S.company)??"")}" />`)}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">กลุ่มโครงการ</label>
          <select id="f-group" class="${v}">
            ${["Internal","External","R&D","Maintenance","อื่นๆ"].map(u=>`<option>${u}</option>`).join("")}
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${v}">
            ${["Planning","Active","On Hold","Completed","Cancelled"].map(u=>`<option>${u}</option>`).join("")}
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันที่เริ่ม</label>
          <input id="f-start" type="date" class="${v}" /></div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันสิ้นสุด</label>
          <input id="f-end" type="date" class="${v}" /></div>
      </div>
      ${k("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${v} resize-y">${b(s)}</textarea>`)}
      ${V()}
    `:a==="projcomment"&&(T=`
      ${k("เลือกโครงการ *",Q())}
      ${k("ประเภท",`<select id="f-comment-type" class="${v}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${k("Comment *",`<textarea id="f-comment" rows="5"
        class="${v} resize-y" placeholder="พิมพ์ comment...">${b(c)}</textarea>`)}
      ${V()}
    `);const m=a==="phish"?nt():a==="comment"||a==="projcomment"?"เพิ่ม Comment":a==="project"?"สร้างโครงการ":a==="incident"?"แจ้ง Incident":a==="task"?"สร้าง Task":"สร้าง Ticket";t.innerHTML=`
    <div class="flex flex-col h-screen bg-slate-50">
      ${w}
      <div class="flex-1 overflow-y-auto">
        ${l}
        ${O}
        ${N}
        <div class="mx-3 mt-3 space-y-3 pb-3">
          ${T}
        </div>
      </div>
      <div class="border-t border-slate-200 bg-white px-3 py-2.5 flex-shrink-0">
        <button id="submit-btn"
          class="w-full bg-blue-700 hover:bg-blue-600 active:bg-blue-800 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          ${m}
        </button>
      </div>
    </div>
  `,(D=document.getElementById("btn-logout"))==null||D.addEventListener("click",ut),(P=document.getElementById("submit-btn"))==null||P.addEventListener("click",wt),(R=document.getElementById("btn-import-customer"))==null||R.addEventListener("click",vt),a==="phish"&&rt();const g=document.getElementById("f-severity"),$=document.getElementById("f-sla");if(g&&$){let u=!1;const y=document.getElementById("f-sla-hint"),_=()=>{if(!y)return;const I=parseInt($.value||"0")||null;y.textContent=I?`นับจากตอนนี้ · ครบกำหนด ${ye(I)}`:"ไม่กำหนด SLA — เคสนี้จะวัดไม่ได้ในรายงาน"};$.addEventListener("change",()=>{u=!0,_()}),g.addEventListener("change",()=>{if(u)return;const I=ae[g.value];I&&($.value=String(I),_())})}document.querySelectorAll(".tab-btn").forEach(u=>{u.addEventListener("click",()=>{const y=u.dataset.tab;y&&y!==i.tab&&(i.tab=y,z(),y==="phish"&&ee())})});const E=document.getElementById("drop-zone"),A=document.getElementById("f-files");E&&A&&(A.addEventListener("change",()=>{A.files&&me(Array.from(A.files)),A.value=""}),E.addEventListener("dragover",u=>{u.preventDefault(),E.classList.add("border-blue-500","bg-blue-50")}),E.addEventListener("dragleave",()=>{E.classList.remove("border-blue-500","bg-blue-50")}),E.addEventListener("drop",u=>{var _;u.preventDefault(),E.classList.remove("border-blue-500","bg-blue-50");const y=Array.from(((_=u.dataTransfer)==null?void 0:_.files)??[]);y.length&&me(y)})),document.querySelectorAll(".remove-dropped").forEach(u=>{u.addEventListener("click",()=>{const y=parseInt(u.dataset.remove??"-1");y>=0&&(i.droppedFiles.splice(y,1),z())})}),It()}function me(t){i.droppedFiles.push(...t),z()}document.addEventListener("paste",t=>{var n;if(!i.account)return;const e=Array.from(((n=t.clipboardData)==null?void 0:n.items)??[]),a=[];for(const o of e)if(o.kind==="file"){const r=o.getAsFile();if(r){const s=r.name&&r.name!=="image.png"?r.name:`screenshot-${new Date().toISOString().replace(/[:.]/g,"-").slice(0,19)}.png`;a.push(new File([r],s,{type:r.type}))}}a.length&&(t.preventDefault(),me(a),B(`แนบไฟล์แล้ว: ${a.map(o=>o.name).join(", ")}`))});const v="w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white";function Te(t){return t<1024?`${t} B`:t<1024*1024?`${(t/1024).toFixed(0)} KB`:`${(t/1024/1024).toFixed(1)} MB`}function V(){const t=i.emailAttachments,e=i.droppedFiles,a=c=>`
    <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
      <input type="checkbox" class="email-att-cb" data-att-id="${b(c.id)}" data-att-name="${b(c.name)}" data-att-item="${c.isItem?"1":"0"}" ${c.defaultOn?"checked":""} />
      <span class="flex-1 truncate">${c.isItem?"📧 ":c.isInline?"🖼️ ":""}${b(c.name)}</span>
      <span class="text-slate-400 flex-shrink-0">${Te(c.size)}</span>
    </label>`,n=t.filter(c=>!c.isInline),o=t.filter(c=>c.isInline),r=t.length>0?`<div class="mb-2 space-y-1">
        ${n.length?`<p class="text-xs text-slate-500">📎 ไฟล์แนบจาก Email:</p>
        ${n.map(a).join("")}`:""}
        ${o.length?`<p class="text-xs text-slate-500 ${n.length?"pt-1":""}">🖼️ รูปในเนื้อเมล
          <span class="text-slate-400">(รูปเล็กมักเป็นโลโก้ในลายเซ็น — ติ๊กเพิ่มได้)</span></p>
        ${o.map(a).join("")}`:""}
      </div>`:"",s=e.length>0?`<div class="mt-2 space-y-1">
        ${e.map((c,p)=>`<div class="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 rounded px-2 py-1">
            <span class="text-base">${c.type.startsWith("image/")?"🖼️":"📄"}</span>
            <span class="flex-1 truncate">${b(c.name)}</span>
            <span class="text-slate-400">${Te(c.size)}</span>
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
  </div>`}function ie(t){var e;return`<select id="f-assigned-email" class="${v}">
    <option value="${b(t)}">${b(((e=i.account)==null?void 0:e.name)??t)} (ฉัน)</option>
    ${i.agents.filter(a=>a.email!==t).map(a=>`<option value="${b(a.email)}">${b(a.name)}</option>`).join("")}
  </select>`}function Q(t=!1){return i.projects.length===0?t?'<div class="text-xs text-slate-400">ไม่พบ Project ที่ Active</div>':'<div class="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-2">⚠️ ไม่พบ Project ที่ Active</div>':`<select id="f-project" class="${v}">
    <option value="">${t?"-- ไม่ผูกกับโครงการ --":"-- เลือก Project --"}</option>
    ${i.projects.map(e=>`<option value="${e.id}">${b(e.Title)}</option>`).join("")}
  </select>`}function k(t,e){return`
    <div class="space-y-1">
      <label class="block text-xs font-medium text-slate-600">${t}</label>
      ${e}
    </div>
  `}function b(t){return t.replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}async function At(){Ke({sharepointUrl:M,internalDomains:["itservices.co.th","rpaexpert.com","rpaexpert.onmicrosoft.com"],getToken:H,getGraphToken:()=>K(),account:()=>i.account?{name:i.account.name,username:i.account.username}:null,toast:(e,a)=>B(e,a??"success"),rerender:z,canWhitelist:()=>["Agent","Supervisor","Boss","Admin"].includes(i.myRole)}),await L.initialize(),await L.handleRedirectPromise();const t=L.getAllAccounts();if(t.length>0){i.account=t[0];try{await L.acquireTokenSilent({scopes:[X],account:t[0]}),await Promise.all([je(),Pe(),De(),Oe()])}catch{i.account=null}}typeof Office<"u"?Office.onReady(e=>{var a;if(e.host===Office.HostType.Outlook){const n=Office.context.mailbox.item;if(n){i.emailSubject=n.subject??"";const o=n.from;o&&(i.emailSenderName=o.displayName??"",i.emailSenderEmail=o.emailAddress??"");const r=(((a=i.account)==null?void 0:a.username)??"").toLowerCase(),s=((o==null?void 0:o.emailAddress)??"").toLowerCase(),c=[...n.to??[],...n.cc??[]].map(l=>l.emailAddress).filter(Boolean);i.emailCc=[...new Set(c.map(l=>l.toLowerCase()))].filter(l=>l!==r&&l!==s);const p=n.attachments??[],w=20*1024;i.emailAttachments=p.filter(l=>l.attachmentType===Office.MailboxEnums.AttachmentType.File||l.attachmentType===Office.MailboxEnums.AttachmentType.Item).map(l=>({id:l.id,name:l.attachmentType===Office.MailboxEnums.AttachmentType.Item?`${(l.name||"email").replace(/\.eml$/i,"")}.eml`:l.name,size:l.size,isItem:l.attachmentType===Office.MailboxEnums.AttachmentType.Item,isInline:!!l.isInline,defaultOn:!l.isInline||l.size>=w})),n.body.getAsync(Office.CoercionType.Html,{asyncContext:{}},l=>{if(l.status===Office.AsyncResultStatus.Succeeded){let h=function(P,R=!1){if(P.nodeType===3){const I=P.textContent??"";return R&&I.trim()===""?"":I}const u=P,y=(u.tagName??"").toLowerCase();if(N.includes(y))return"";if(y==="br")return" ";if(y==="tr"){const I=[];for(let U=0;U<u.childNodes.length;U++){const ge=u.childNodes[U],be=(ge.tagName??"").toLowerCase();(be==="td"||be==="th")&&I.push((ge.textContent??"").replace(/\s+/g," ").trim())}return I.length?I.join("	")+`
`:""}if(m.includes(y)){let I="";for(let U=0;U<u.childNodes.length;U++)I+=h(u.childNodes[U],!0);return I}let _="";for(let I=0;I<u.childNodes.length;I++)_+=h(u.childNodes[I],!1);return T.includes(y)&&(_=`
`+_.trim()+`
`),_};const j=l.value,O=new DOMParser().parseFromString(j,"text/html"),N=["style","script","head","img","meta","link","noscript"],T=["p","div","li","h1","h2","h3","h4","h5","h6","blockquote"],m=["table","thead","tbody","tfoot"],E=h(O.body??O.documentElement).replace(/[ \t]{2,}/g," ").replace(/\n[ \t]+/g,`
`).replace(/\n{3,}/g,`

`).trim().split(`
`),A=[];let x="";for(const P of E)P.trim()===""?x&&(A.push(x.trim()),x=""):P.includes("	")?(x&&(A.push(x.trim()),x=""),A.push(P)):x=x?x+" "+P.trim():P.trim();x&&A.push(x.trim());const C=A.join(`
`),S=/\n([-_]{3,}|From:\s|Best regards|Regards,|ขอแสดงความนับถือ|Sent:\s)/i,D=C.search(S);i.signatureContact=D>80?xt(C.slice(D).trim()):null,i.emailBodyPreview=C.trim().slice(0,2e4),i.emailBodyReply=Ge(i.emailBodyPreview)}z()});return}}ke(),z()}):(ke(),z())}function ke(){i.emailSubject="[DEV] Test Email Subject",i.emailSenderName="Test Sender",i.emailSenderEmail="test@example.com",i.emailBodyPreview="This is a placeholder email body for development mode.",i.emailBodyReply=i.emailBodyPreview}At().catch(t=>{console.error("Init error:",t);const e=document.getElementById("app");e&&(e.innerHTML=`<div class="p-4 text-red-600 text-sm">เกิดข้อผิดพลาด: ${String(t)}</div>`)});
