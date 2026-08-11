import{P as Me}from"./PublicClientApplication-DLKYUtZW.js";import{a as ee,d as He,p as Ne,L as ke,S as Ee}from"./analyzer-CWd3MChg.js";const Fe=[{hours:1,labelTh:"1 ชั่วโมง"},{hours:2,labelTh:"2 ชั่วโมง"},{hours:4,labelTh:"4 ชั่วโมง"},{hours:8,labelTh:"8 ชั่วโมง (1 วันทำการ)"},{hours:24,labelTh:"24 ชั่วโมง"},{hours:48,labelTh:"2 วัน"},{hours:72,labelTh:"3 วัน"},{hours:168,labelTh:"7 วัน"}],te={Critical:1,High:4,Medium:24,Low:72};function Ie(t,e=new Date){const a=typeof t=="number"&&Number.isFinite(t)&&t>0?t:null;return a?new Date(e.getTime()+a*36e5).toISOString():null}function be(t){const e=Ie(t);return e?new Date(e).toLocaleString("th-TH",{dateStyle:"short",timeStyle:"short"}):""}const Ae="HD_PhishingReports";let f;const l={mail:null,analysis:null,loading:!1,reporting:!1,reported:!1,showHeaders:!1,kasmTemplate:"",analysedItemId:"",safeDomains:[],safeDomainIds:{},savingDomain:""};function ze(t){f=t}const F=t=>(t??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"),K=()=>{var t,e;return(e=(t=Office.context)==null?void 0:t.mailbox)==null?void 0:e.item},ye=t=>new Promise(e=>{const a=K();if(!(a!=null&&a.body)){e("");return}a.body.getAsync(t,n=>e(n.status===Office.AsyncResultStatus.Succeeded?n.value??"":""))});function Ue(t){return t?t.split(",").map(e=>{const a=e.trim(),n=a.match(/^(.*?)\s*<([^>]+)>$/);return n?{name:n[1].replace(/^"|"$/g,"").trim(),email:n[2].trim()}:{name:"",email:a.replace(/[<>]/g,"").trim()}}).filter(e=>e.email.includes("@")):[]}function qe(){return new Promise(t=>{let e=!1;try{e=Office.context.requirements.isSetSupported("Mailbox","1.8")}catch{e=!1}const a=K();if(!e||typeof(a==null?void 0:a.getAllInternetHeadersAsync)!="function"){t({});return}try{a.getAllInternetHeadersAsync(n=>t(n.status===Office.AsyncResultStatus.Succeeded?Ne(n.value??""):{}))}catch{t({})}})}async function Ge(){try{const t=K();if(!(t!=null&&t.itemId))return{};const e=Office.context.mailbox.convertToRestId(t.itemId,Office.MailboxEnums.RestVersion.v2_0),a=await f.getGraphToken(),n=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${e}?$select=internetMessageHeaders`,{headers:{Authorization:`Bearer ${a}`}});if(!n.ok)return{};const o=await n.json(),r={};for(const s of o.internetMessageHeaders??[])r[s.name]=r[s.name]?`${r[s.name]}
${s.value}`:s.value;return r}catch{return{}}}async function me(t,e){const a=await f.getToken(),n=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('${t}')/items?${e}`,{headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata"}});return n.ok?(await n.json()).value:[]}async function Ve(){try{return(await me("HD_AgentProfiles","$select=Title,EmailText&$top=500")).filter(e=>e.EmailText).map(e=>({name:e.Title,email:e.EmailText}))}catch{return[]}}async function Ke(){var t;try{return(((t=(await me("HD_Options","$select=Title&$filter=Category eq 'KasmConfig'&$top=1"))[0])==null?void 0:t.Title)??"").trim()}catch{return""}}const Ce="SafeDomain";async function ue(){try{const t=await me("HD_Options",`$select=Id,Title&$filter=Category eq '${Ce}'&$top=500`),e={},a=[];for(const n of t){const o=(n.Title??"").trim().toLowerCase();o&&(e[o]=n.Id,a.push(o))}l.safeDomains=a,l.safeDomainIds=e}catch{}}async function We(t){const e=t.trim().toLowerCase();if(!(!e||l.savingDomain)){if(!f.canWhitelist()){f.toast("ต้องเป็น Agent ขึ้นไปจึงจะยืนยันโดเมนได้","error");return}l.savingDomain=e,f.rerender();try{const a=await f.getToken(),n=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items`,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:e,Category:Ce})});if(!n.ok)throw new Error(String(n.status));await ue(),await Q(!0),f.toast(`ยืนยันแล้วว่า ${e} ปลอดภัย`)}catch{f.toast("บันทึกไม่สำเร็จ","error")}finally{l.savingDomain="",f.rerender()}}}async function Je(t){const e=t.trim().toLowerCase(),a=l.safeDomainIds[e];if(!(!a||l.savingDomain)){if(!f.canWhitelist()){f.toast("ต้องเป็น Agent ขึ้นไปจึงจะถอนโดเมนได้","error");return}l.savingDomain=e,f.rerender();try{const n=await f.getToken(),o=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items(${a})`,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","IF-MATCH":"*","X-HTTP-Method":"DELETE"}});if(!o.ok)throw new Error(String(o.status));await ue(),await Q(!0),f.toast(`ถอน ${e} ออกจากรายการปลอดภัยแล้ว`)}catch{f.toast("ถอนไม่สำเร็จ","error")}finally{l.savingDomain="",f.rerender()}}}async function Q(t=!1){var h,p;const e=K(),a=(e==null?void 0:e.itemId)??"";if(!t&&l.analysedItemId===a&&l.analysis)return;l.loading=!0,l.reported=!1,l.analysedItemId=a,f.rerender();const[n,o]=await Promise.all([ye(Office.CoercionType.Html),ye(Office.CoercionType.Text)]),r={fromName:((h=e==null?void 0:e.from)==null?void 0:h.displayName)??"",fromEmail:((p=e==null?void 0:e.from)==null?void 0:p.emailAddress)??"",replyTo:[],subject:(e==null?void 0:e.subject)??"",bodyHtml:n,bodyText:o,attachments:((e==null?void 0:e.attachments)??[]).map(u=>({name:u.name,size:u.size??0,isInline:!!u.isInline})),headers:{},internalDomains:f.internalDomains,internalPeople:[],safeDomains:l.safeDomains},s=u=>{var A;return{...r,headers:u,replyTo:Ue(((A=Object.entries(u).find(([O])=>O.toLowerCase()==="reply-to"))==null?void 0:A[1])??"")}},c=await qe();if(l.mail=s(c),l.analysis=ee(l.mail),l.loading=!1,f.rerender(),f.account()){const[u,A]=await Promise.all([Ve(),Object.keys(c).length?Promise.resolve(c):Ge()]);l.mail={...s(A),internalPeople:u},l.analysis=ee(l.mail),l.kasmTemplate||(l.kasmTemplate=await Ke()),l.safeDomains.length||(await ue(),l.safeDomains.length&&(l.mail={...l.mail,safeDomains:l.safeDomains},l.analysis=ee(l.mail))),f.rerender()}}function Se(){const t=l.mail,e=l.analysis;return!t||!e?"":[`ผู้ส่ง: ${t.fromName} <${t.fromEmail}>`,`หัวข้อ: ${t.subject}`,t.replyTo.length?`Reply-To: ${t.replyTo.map(a=>a.email).join(", ")}`:"",`คะแนนความเสี่ยง: ${e.score} (${ke[e.level].label})`,"","สิ่งที่ตรวจพบ:",...e.findings.map(a=>`- [${Ee[a.severity].label}] (${a.category}) ${a.title} — ${a.detail.replace(/\n/g," ")}`),"",e.links.length?"ลิงก์ในอีเมล:":"",...e.links.map(a=>`- ${a.href}${a.flags.length?`  ! ${a.flags.join(" / ")}`:""}`)].filter(a=>a!=="").join(`
`)}async function Ze(t){try{const e=K();if(!(e!=null&&e.itemId))return!1;const a=Office.context.mailbox.convertToRestId(e.itemId,Office.MailboxEnums.RestVersion.v2_0),n=await f.getGraphToken(),o=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${a}/$value`,{headers:{Authorization:`Bearer ${n}`}});if(!o.ok)return!1;const r=await o.arrayBuffer(),s=(e.subject||"phishing").replace(/[\\/:*?"<>|#%&{}~]/g,"_").replace(/^_+/,"").slice(0,80).trim()||"phishing",c=await f.getToken();return(await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('${Ae}')/items(${t})/AttachmentFiles/add(FileName='${encodeURIComponent(s+".eml")}')`,{method:"POST",headers:{Authorization:`Bearer ${c}`,Accept:"application/json;odata=nometadata"},body:r})).ok}catch{return!1}}async function xe(t){const e=await f.getToken(),a=await fetch(`${f.sharepointUrl}/_api/web/lists/getbytitle('${Ae}')/items`,{method:"POST",headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(t)});if(!a.ok)throw new Error(`SharePoint ${a.status}: ${await a.text()}`);return(await a.json()).Id}async function Ye(){if(!(!l.mail||!l.analysis||l.reporting)){if(!f.account()){f.toast("กรุณาเข้าสู่ระบบก่อนรายงาน","error");return}l.reporting=!0,f.rerender();try{const t=l.mail,e=l.analysis,a=f.account(),n={Title:(t.subject||"(ไม่มีหัวข้อ)").slice(0,255),SenderName:t.fromName.slice(0,255),SenderEmail:t.fromEmail.slice(0,255),SenderDomain:He(t.fromEmail),RiskScore:e.score,RiskLevel:e.level,Findings:Se(),LinkCount:e.links.length,SuspiciousLinks:e.links.filter(c=>c.flags.length).map(c=>c.href).join(`
`).slice(0,4e3),ReportedBy:(a==null?void 0:a.name)??"",ReportedEmail:(a==null?void 0:a.username)??"",Status:"New"};let o,r=!1;try{o=await xe(n)}catch(c){o=await xe({Title:n.Title,Findings:n.Findings}).catch(()=>{throw c}),r=!0}const s=await Ze(o);l.reported=!0,f.toast(r?"ส่งรายงานแล้ว แต่บันทึกได้บางคอลัมน์ — ตรวจชื่อคอลัมน์ใน HD_PhishingReports":s?"ส่งรายงานพร้อมอีเมลต้นฉบับแล้ว":"ส่งรายงานแล้ว (แนบ .eml ไม่ได้)",r?"info":"success")}catch(t){f.toast(`ส่งรายงานไม่สำเร็จ: ${t instanceof Error?t.message:String(t)}`,"error")}finally{l.reporting=!1,f.rerender()}}}const Xe=()=>l.reported?"✓ รายงานแล้ว":l.reporting?"กำลังส่ง…":"🚩 รายงานอีเมลนี้ให้ IT";function Qe(t){var e;try{const a=(e=Office.context)==null?void 0:e.ui;if(typeof(a==null?void 0:a.openBrowserWindow)=="function"){a.openBrowserWindow(t);return}}catch{}window.open(t,"_blank","noopener,noreferrer")||f.toast("เปิดหน้าต่างไม่ได้ (ถูกบล็อก)","info")}async function et(t){var e;try{if((e=navigator.clipboard)!=null&&e.writeText)return await navigator.clipboard.writeText(t),!0}catch{}try{const a=document.createElement("textarea");a.value=t,a.setAttribute("readonly",""),a.style.position="fixed",a.style.opacity="0",document.body.appendChild(a),a.focus(),a.select();const n=document.execCommand("copy");return a.remove(),n}catch{return!1}}function tt(t){if(l.loading&&!l.analysis)return`<div class="py-10 text-center text-slate-500 text-sm">
      <div class="w-7 h-7 border-[3px] border-slate-200 border-t-blue-600 rounded-full animate-spin mx-auto mb-2"></div>
      กำลังตรวจอีเมล…</div>`;const e=l.analysis,a=l.mail;if(!e||!a)return'<p class="text-sm text-slate-400 text-center py-8">เปิดอีเมลเพื่อเริ่มตรวจ</p>';const n=ke[e.level],o=[...e.links.filter(s=>s.flags.length),...e.links.filter(s=>!s.flags.length&&!s.trusted),...e.links.filter(s=>s.trusted)],r=f.canWhitelist();return`
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

    ${e.findings.length===0?'<p class="text-xs text-slate-400 text-center py-3">ไม่พบสัญญาณผิดปกติจากการตรวจอัตโนมัติ</p>':e.findings.map(s=>{const c=Ee[s.severity];return`<div class="bg-white rounded-xl border border-slate-200 p-2.5">
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
        ${o.map((s,c)=>{const h=l.savingDomain===(s.host?s.host.toLowerCase():"");return`
          <div class="rounded-lg border ${s.trusted?"border-emerald-200 bg-emerald-50/50":s.flags.length?"border-red-200 bg-red-50/50":"border-slate-100"} p-2">
            <div class="text-[11px] font-medium ${s.trusted?"text-emerald-800":s.flags.length?"text-red-700":"text-slate-700"} break-all">
              ${s.trusted?"✔ ":""}${F(s.host||s.href)}
            </div>
            ${s.trusted?`<div class="text-[10px] text-emerald-700">ทีมตรวจแล้วว่าปลอดภัย${(s.suppressed??[]).length?` · ระงับการเตือน ${s.suppressed.length} ข้อ`:""}</div>`:""}
            ${s.text&&s.text!==s.href?`<div class="text-[10px] text-slate-500 break-all">แสดงว่า: "${F(s.text.slice(0,70))}"</div>`:""}
            <div class="text-[10px] text-slate-400 break-all mt-0.5">${F(s.href.slice(0,150))}</div>
            ${s.flags.map(u=>`<div class="text-[10px] text-red-600 mt-0.5">! ${F(u)}</div>`).join("")}
            <div class="flex flex-wrap gap-1 mt-1.5">
              <button data-kasm="${c}" class="text-[10px] font-semibold px-2 py-1 rounded-md bg-slate-800 hover:bg-slate-900 text-white">
                เปิดใน Kasm
              </button>
              ${r&&s.host?s.trusted?`<button data-untrust="${F(s.host)}" ${h?"disabled":""}
                     class="text-[10px] font-semibold px-2 py-1 rounded-md border border-slate-300 text-slate-600 hover:bg-slate-100 disabled:opacity-50">
                     ${h?"...":"ถอนออกจากรายการปลอดภัย"}</button>`:`<button data-trust="${F(s.host)}" ${h?"disabled":""}
                     class="text-[10px] font-semibold px-2 py-1 rounded-md bg-emerald-600 hover:bg-emerald-700 text-white disabled:opacity-50">
                     ${h?"...":"✔ ตรวจแล้ว ปลอดภัย"}</button>`:""}
            </div>
          </div>`}).join("")}
      </div>
      <p class="text-[10px] text-slate-400 mt-2">อย่าคลิกลิงก์จากอีเมลที่ไม่มั่นใจโดยตรง</p>
    </div>`:""}

    <div class="flex gap-2">
      <button id="phish-copy" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">คัดลอกผลตรวจ</button>
      <button id="phish-headers" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">${l.showHeaders?"ซ่อน header":"ดู header"}</button>
    </div>
    ${l.showHeaders?`<pre class="bg-slate-900 text-slate-100 text-[9px] p-2 rounded-lg overflow-x-auto whitespace-pre-wrap break-all max-h-56 overflow-y-auto">${F(Object.entries(a.headers).map(([s,c])=>`${s}: ${c}`).join(`
`)||"อ่าน header ไม่ได้")}</pre>`:""}
    <p class="text-[10px] text-slate-400 text-center">วิเคราะห์ในเครื่องทั้งหมด — ไม่ส่งเนื้อหาอีเมลออกนอกองค์กร</p>
  `}function at(){var a,n,o,r;(a=document.getElementById("phish-recheck"))==null||a.addEventListener("click",()=>{Q(!0)}),(n=document.getElementById("phish-headers"))==null||n.addEventListener("click",()=>{l.showHeaders=!l.showHeaders,f.rerender()}),(o=document.getElementById("phish-copy"))==null||o.addEventListener("click",async()=>{const s=await et(Se());f.toast(s?"คัดลอกผลตรวจแล้ว":"คัดลอกไม่ได้",s?"success":"error")}),document.querySelectorAll("[data-trust]").forEach(s=>s.addEventListener("click",()=>We(s.dataset.trust??""))),document.querySelectorAll("[data-untrust]").forEach(s=>s.addEventListener("click",()=>Je(s.dataset.untrust??"")));const t=((r=l.analysis)==null?void 0:r.links)??[],e=[...t.filter(s=>s.flags.length),...t.filter(s=>!s.flags.length&&!s.trusted),...t.filter(s=>s.trusted)];document.querySelectorAll("[data-kasm]").forEach(s=>{s.addEventListener("click",()=>{const c=e[Number(s.dataset.kasm)];if(!c)return;if(!l.kasmTemplate){f.toast("ยังไม่ได้ตั้งค่า Kasm ใน HD_Options (Category=KasmConfig)","info");return}const h=l.kasmTemplate;Qe(h.includes("{url}")?h.replace("{url}",encodeURIComponent(c.href)):h+encodeURIComponent(c.href))})})}const nt="0bab07cf-65e6-487c-89af-c917fc1a5a13",st="d569b991-89fc-4a62-9df5-eb361abcef40",H="https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd",X="https://rpaexpert.sharepoint.com/.default",ae=["https://graph.microsoft.com/Calendars.ReadWrite","https://graph.microsoft.com/Mail.Send","https://graph.microsoft.com/Mail.Read"],L=new Me({auth:{clientId:nt,authority:`https://login.microsoftonline.com/${st}`,redirectUri:window.location.origin.includes("localhost")?"http://localhost:3000/":"https://darmmunginsa.github.io/itservices-addin/",navigateToLoginRequestUrl:!1},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:!1}}),ot=window.location.origin.includes("localhost")?`${window.location.origin}/auth.html`:"https://darmmunginsa.github.io/itservices-addin/auth.html";function pe(){var e,a;const t=(a=(e=Office.context)==null?void 0:e.diagnostics)==null?void 0:a.platform;return t===Office.PlatformType.iOS||t===Office.PlatformType.Android}function fe(){return new Promise((t,e)=>{Office.context.ui.displayDialogAsync(ot,{height:60,width:30,promptBeforeOpen:!1},a=>{if(a.status!==Office.AsyncResultStatus.Succeeded){e(new Error("เปิดหน้าเข้าสู่ระบบไม่ได้"));return}const n=a.value;n.addEventHandler(Office.EventType.DialogMessageReceived,o=>{n.close();const r=o.message;if(!r){e(new Error("auth message error"));return}try{const s=JSON.parse(r);s.ok?t():e(new Error(s.error||"auth failed"))}catch{e(new Error("auth message error"))}}),n.addEventHandler(Office.EventType.DialogEventReceived,()=>e(new Error("ปิดหน้าเข้าสู่ระบบก่อนเสร็จ")))})})}const i={account:null,tab:"ticket",emailSubject:"",emailBodyPreview:"",emailSenderName:"",emailSenderEmail:"",loading:!1,projects:[],agents:[],myRole:"",emailAttachments:[],signatureContact:null,droppedFiles:[],tickets:[],contactEmails:[],emailCc:[]};async function N(){const t=L.getAllAccounts();if(t.length===0)throw new Error("Not signed in");const e={scopes:[X],account:t[0]};try{return(await L.acquireTokenSilent(e)).accessToken}catch{if(pe()){await fe();const a=L.getAllAccounts()[0];if(!a)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await L.acquireTokenSilent({scopes:[X],account:a})).accessToken}return(await L.acquireTokenPopup(e)).accessToken}}async function V(t=!1){const e=L.getAllAccounts();if(e.length===0)throw new Error("Not signed in");const a={scopes:ae,account:e[0],forceRefresh:t};try{return(await L.acquireTokenSilent(a)).accessToken}catch{if(pe()){await fe();const o=L.getAllAccounts()[0];if(!o)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await L.acquireTokenSilent({scopes:ae,account:o})).accessToken}return(await L.acquireTokenPopup({scopes:ae,account:e[0]})).accessToken}}async function it(t){const e=await V(),a={subject:t.subject,start:{dateTime:t.start,timeZone:"Asia/Bangkok"},end:{dateTime:t.end,timeZone:"Asia/Bangkok"},body:t.body?{contentType:"HTML",content:t.body.replace(/\n/g,"<br>")}:void 0,attendees:t.attendees.filter(Boolean).map(o=>({emailAddress:{address:o},type:"required"})),isOnlineMeeting:t.isOnlineMeeting,onlineMeetingProvider:t.isOnlineMeeting?"teamsForBusiness":void 0},n=await fetch("https://graph.microsoft.com/v1.0/me/events",{method:"POST",headers:{Authorization:`Bearer ${e}`,"Content-Type":"application/json"},body:JSON.stringify(a)});if(!n.ok)throw new Error(`Calendar error ${n.status}: ${await n.text()}`)}async function Be(){try{const t=await N(),e=`${H}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$orderby=Title asc&$top=500`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();i.projects=n.value.map(o=>({id:o.Id,Title:o.Title}))}}catch{}}async function je(){var t,e;try{const a=await N(),n=`${H}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText,Role&$orderby=Title asc`,o=await fetch(n,{headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata"}});if(o.ok){const r=await o.json();i.agents=r.value.map(c=>({email:c.EmailText,name:c.Title}));const s=(((t=i.account)==null?void 0:t.username)??"").toLowerCase();i.myRole=((e=r.value.find(c=>(c.EmailText??"").toLowerCase()===s))==null?void 0:e.Role)??""}}catch{}}async function Pe(){try{const t=await N(),e=`${H}/_api/web/lists/getbytitle('HD_Tickets')/items?$select=Id,Title,TicketNumber,Status&$filter=Status ne 'Closed'&$orderby=Modified desc&$top=200`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();i.tickets=n.value.map(o=>({id:o.Id,Title:o.Title,TicketNumber:o.TicketNumber,Status:o.Status}))}}catch{}}async function De(){try{const t=await N(),e=`${H}/_api/web/lists/getbytitle('HD_Contracts')/items?$select=CustomerEmail&$top=2000`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();i.contactEmails=n.value.map(o=>(o.CustomerEmail||"").trim().toLowerCase()).filter(Boolean)}}catch{}}async function ve(){const t=document.getElementById("btn-login-main"),e=document.getElementById("btn-login");t&&(t.disabled=!0,t.textContent="กำลังเข้าสู่ระบบ…"),e&&(e.disabled=!0);try{if(pe()){if(await fe(),i.account=L.getAllAccounts()[0]??null,!i.account)throw new Error("เข้าสู่ระบบไม่สำเร็จ")}else{const a=await L.loginPopup({scopes:[X]});i.account=a.account}await Promise.all([Be(),je(),Pe(),De()]),z()}catch{t&&(t.disabled=!1,t.textContent="เข้าสู่ระบบ"),e&&(e.disabled=!1)}}async function rt(){i.account&&await L.logoutPopup({account:i.account}),i.account=null,z()}async function U(t,e){const a=await N(),n=`${H}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items`,o=await fetch(n,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(e)});if(!o.ok){const s=await o.text();throw new Error(`SharePoint error ${o.status}: ${s}`)}return(await o.json()).Id}let W=null;const ne="support@itservices.co.th",Oe="engineer@itservices.co.th";async function Le(){if(W)return W;try{const t=await N(),e=`${H}/_api/web/lists/getbytitle('HD_EmailTemplates')/items?$select=EventKey,Subject,Body,IsEnabled&$top=50`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});return a.ok?(W=(await a.json()).value,W):[]}catch{return[]}}async function ct(){var t,e;try{const a=await N(),n=`${H}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title,Category&$filter=Category eq 'EmailConfig'&$top=1`,o=await fetch(n,{headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata"}});return o.ok&&((e=(t=(await o.json()).value[0])==null?void 0:t.Title)==null?void 0:e.trim())||ne}catch{return ne}}function ie(t,e){return t.replace(/\{\{(\w+)\}\}/g,(a,n)=>e[n]??`{{${n}}}`)}async function lt(t,e,a,n=[]){try{const r=(await Le()).find(d=>d.EventKey===t&&d.IsEnabled);if(!r)return;const s=ie(r.Subject||"",e),c=ie(r.Body||"",e);if(!s||!c)return;const h=d=>d.trim().toLowerCase(),p=[...new Map(a.filter(Boolean).map(d=>[h(d),d])).values()];if(p.length===0)return;const u=new Set(p.map(h)),A=t==="ticket_created"?[...n,Oe]:n,O=[...new Map(A.filter(Boolean).map(d=>[h(d),d])).values()].filter(d=>!u.has(h(d))),_=await ct(),P=await V(),E={subject:s,body:{contentType:"HTML",content:c},toRecipients:p.map(d=>({emailAddress:{address:d}}))};O.length&&(E.ccRecipients=O.map(d=>({emailAddress:{address:d}}))),_&&(E.from={emailAddress:{address:_}}),await fetch("https://graph.microsoft.com/v1.0/me/sendMail",{method:"POST",headers:{Authorization:`Bearer ${P}`,"Content-Type":"application/json"},body:JSON.stringify({message:E,saveToSentItems:!0})})}catch{}}async function dt(t,e=[]){try{const a=Office.context.mailbox.item;if(!(a!=null&&a.itemId))return!1;const n=Office.context.mailbox.convertToRestId(a.itemId,Office.MailboxEnums.RestVersion.v2_0),r={Authorization:`Bearer ${await V()}`,"Content-Type":"application/json"},s=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${n}/createReplyAll`,{method:"POST",headers:r});if(!s.ok)return!1;const c=await s.json(),h=E=>E.trim().toLowerCase(),p=c.ccRecipients??[],u=new Set(p.map(E=>h(E.emailAddress.address))),A=[...new Set(e.filter(Boolean).map(E=>E.trim()))].filter(E=>!u.has(h(E))).map(E=>({emailAddress:{address:E}})),O={body:{contentType:"HTML",content:t}};return A.length&&(O.ccRecipients=[...p,...A]),(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${c.id}`,{method:"PATCH",headers:r,body:JSON.stringify(O)})).ok?(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${c.id}/send`,{method:"POST",headers:r})).ok:!1}catch{return!1}}async function mt(t,e){const n=(await Le()).find(r=>r.EventKey===t&&r.IsEnabled);return n&&ie(n.Body||"",e)||null}async function J(t){var r;const e=s=>s.trim().toLowerCase(),a=e(((r=i.account)==null?void 0:r.username)??""),n=new Set,o=t.recipients.filter(Boolean).filter(s=>{const c=e(s);return!c||c===a||n.has(c)?!1:(n.add(c),!0)});if(o.length!==0)try{const s=await N(),c=`${H}/_api/web/lists/getbytitle('HD_Notifications')/items`;await Promise.all(o.map(h=>fetch(c,{method:"POST",headers:{Authorization:`Bearer ${s}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:t.title.slice(0,255),RecipientEmail:h,EventType:t.eventType,Message:t.message,LinkPath:t.linkPath,IsRead:!1})})))}catch{}}async function re(t,e){const a=document.querySelectorAll(".email-att-cb:checked");if(a.length===0)return;const n=await N();for(const o of Array.from(a)){const r=o.dataset.attId,s=o.dataset.attName,c=await new Promise((P,E)=>{Office.context.mailbox.item.getAttachmentContentAsync(r,{},d=>{d.status===Office.AsyncResultStatus.Succeeded?P(d):E(new Error(d.error.message))})}),{content:h,format:p}=c.value;let u;if(p===Office.MailboxEnums.AttachmentContentFormat.Base64){const P=atob(h);u=new Uint8Array(P.length);for(let E=0;E<P.length;E++)u[E]=P.charCodeAt(E)}else if(p===Office.MailboxEnums.AttachmentContentFormat.Eml||p===Office.MailboxEnums.AttachmentContentFormat.ICalendar)u=new TextEncoder().encode(h);else continue;const A=encodeURIComponent(s),O=`${H}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${A}')`;if(!(await fetch(O,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:u.buffer})).ok)throw new Error(`Upload ${s} failed`)}}async function ut(t){const e=`https://graph.microsoft.com/v1.0/me/messages/${t}/$value`;let a=await V(),n=await fetch(e,{headers:{Authorization:`Bearer ${a}`}});if((n.status===401||n.status===403)&&(a=await V(!0),n=await fetch(e,{headers:{Authorization:`Bearer ${a}`}})),!n.ok)throw new Error(`Graph ${n.status}`);return n.arrayBuffer()}async function pt(t){const e=await new Promise((n,o)=>{Office.context.mailbox.getCallbackTokenAsync({isRest:!0},r=>{r.status===Office.AsyncResultStatus.Succeeded?n(r.value):o(new Error("callback token failed"))})}),a=await fetch(`${Office.context.mailbox.restUrl}/v2.0/me/messages/${t}/$value`,{headers:{Authorization:`Bearer ${e}`}});if(!a.ok)throw new Error(`REST ${a.status}`);return a.arrayBuffer()}async function ce(t,e){const a=document.getElementById("f-attach-eml");if(!(a!=null&&a.checked))return;const n=Office.context.mailbox.item;if(!n)return;const o=Office.context.mailbox.convertToRestId(n.itemId,Office.MailboxEnums.RestVersion.v2_0);let r,s="",c="";try{r=await ut(o)}catch(O){s=O instanceof Error?O.message:String(O);try{r=await pt(o)}catch(_){c=_ instanceof Error?_.message:String(_),console.error("[eml] graph:",s,"| callback:",c),S(`ดึง .eml ไม่ได้ (Graph: ${s} / REST: ${c}) — ไฟล์อื่นบันทึกแล้ว`,"error");return}}const h=(n.subject||"email").replace(/[\\/:*?"<>|#%&{}~]/g,"_").slice(0,100).trim()||"email",p=await N(),u=`${H}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${encodeURIComponent(h+".eml")}')`;(await fetch(u,{method:"POST",headers:{Authorization:`Bearer ${p}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok||S("บันทึกไฟล์ .eml ไม่สำเร็จ","error")}async function le(t,e,a){const n=await N();for(const o of a){const r=await o.arrayBuffer(),s=encodeURIComponent(o.name),c=`${H}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${s}')`;if(!(await fetch(c,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok)throw new Error(`Upload ${o.name} failed`)}}function S(t,e="success"){const a=document.getElementById("toast-container");if(!a)return;const n=e==="success"?"bg-green-500":e==="error"?"bg-red-500":"bg-slate-700",o=e==="success"?"✅":e==="error"?"❌":"ℹ️",r=document.createElement("div");r.className=`toast pointer-events-auto ${n} text-white text-sm font-medium px-4 py-3 rounded-lg shadow-lg max-w-xs mx-2`,r.textContent=`${o} ${t}`,a.appendChild(r),setTimeout(()=>r.remove(),4e3)}function ft(t){const e=t.split(`
`).map(u=>u.trim()).filter(Boolean),a=/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/,n=/(\+?[\d\s()\-]{7,}(?:ext\.?\s*\d+)?)/i,o=/\b(co\.?,?\s*ltd\.?|co\.?,?\s*limited|corporation|corp\.?|บริษัท|จำกัด|holding|group|inc\.?|llc)\b/i;let r="",s="",c="";const h=[];for(const u of e)if(!/^[-_=*]{2,}$/.test(u)&&!/^(best regards|regards|sincerely|ขอแสดงความนับถือ|ด้วยความนับถือ|from|sent|thanks|thank you)/i.test(u)){if(!r){const A=u.match(a);if(A){r=A[0];continue}}if(!s){const A=u.match(n);if(A&&A[0].replace(/\D/g,"").length>=7){s=A[0].trim();continue}}if(!c&&o.test(u)){c=u;continue}u.length>=2&&u.length<=50&&!/\d{4,}/.test(u)&&h.push(u)}const p=h.find(u=>!a.test(u)&&!o.test(u))??"";return!r&&!p?null:{name:p,company:c,email:r,phone:s}}async function ht(){const t=i.signatureContact;if(!t)return;const e=(i.emailSenderEmail||"").toLowerCase();if(e&&i.contactEmails.includes(e)){S("ลูกค้านี้มีในระบบแล้ว","success"),i.signatureContact=null,z();return}const a=document.getElementById("btn-import-customer");a&&(a.disabled=!0,a.textContent="กำลังบันทึก…");try{await U("HD_Contracts",{Title:i.emailSenderName||t.name,CustomerEmail:i.emailSenderEmail,Phone:t.phone||void 0,Company:t.company||void 0,Status:"Active"}),e&&i.contactEmails.push(e),S("เพิ่มลูกค้าสำเร็จ!"),i.signatureContact=null,z()}catch(n){const o=n instanceof Error?n.message:String(n);S(`เกิดข้อผิดพลาด: ${o}`,"error"),a&&(a.disabled=!1,a.textContent="เพิ่มเป็นลูกค้า")}}function _e(){return new Date().toISOString().split("T")[0]}function gt(){const t=new Date;return`HD-${`${t.getFullYear()}${String(t.getMonth()+1).padStart(2,"0")}${String(t.getDate()).padStart(2,"0")}`}-${Math.floor(Math.random()*900+100)}`}function $e(){var t;return i.droppedFiles.length>0||document.querySelectorAll(".email-att-cb:checked").length>0||(((t=document.getElementById("f-attach-eml"))==null?void 0:t.checked)??!1)}async function Z(t,e){i.droppedFiles.length>0&&await le(t,e,i.droppedFiles),await re(t,e),await ce(t,e)}let se=!1;async function bt(){var e,a,n,o,r,s,c,h,p,u,A,O,_,P,E;if(!i.account){S("กรุณาเข้าสู่ระบบก่อน","error");return}if(se)return;se=!0;const t=document.getElementById("submit-btn");t&&(t.disabled=!0,t.textContent="กำลังบันทึก…");try{if(i.tab==="phish")await Ye();else if(i.tab==="ticket"){const d=document.getElementById("f-title").value.trim(),b=document.getElementById("f-description").value.trim(),$=document.getElementById("f-priority").value,w=document.getElementById("f-customer-email").value.trim(),B=((e=document.getElementById("f-cc-enable"))==null?void 0:e.checked)??!0?(((a=document.getElementById("f-cc"))==null?void 0:a.value)||"").split(/[,;\s]+/).map(R=>R.trim()).filter(Boolean):[],j=document.getElementById("f-assigned-email").value,I=i.agents.find(R=>R.email===j),y=gt(),M=await U("HD_Tickets",{Title:d,TicketNumber:y,Description:b,Priority:$,CustomerEmail:w,CustomerName:i.emailSenderName||w,Status:"Open",AssignedEmail:j||void 0,AssignedToName:(I==null?void 0:I.name)??((n=i.account)==null?void 0:n.name)??"",ProjectID:parseInt(((o=document.getElementById("f-project"))==null?void 0:o.value)||"0")||null});if($e()){const R=await U("HD_TicketComments",{Title:"📎 ไฟล์แนบจากอีเมล",TicketID:M,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await Z("HD_TicketComments",R)}i.droppedFiles=[];const m={ticket_number:y,ticket_title:d,priority:$,category:"-",description:(b||"-").replace(/\n/g,"<br>"),customer_name:i.emailSenderName||w,assigned_name:(I==null?void 0:I.name)??((r=i.account)==null?void 0:r.name)??"-",link:"https://itservices.co.th/helpdesk/"},x=[j,i.account.username,...B,Oe].filter(Boolean);let C=!1;const D=await mt("ticket_created",m);if(D){const R=`<div style="border-left:4px solid #2563eb;background:#eff6ff;padding:10px 14px;margin:0 0 14px;font-family:Segoe UI,sans-serif">
             <div style="font-size:15px;font-weight:700;color:#1e40af">Ticket No. ${y}</div>
             <div style="font-size:12px;color:#475569;margin-top:2px">กรุณาตอบกลับในอีเมลฉบับนี้เพื่อให้ข้อมูลอยู่ใน Ticket เดียวกัน</div>
           </div>`;C=await dt(R+D,x)}C||await lt("ticket_created",m,[w],x),S(C?"สร้าง Ticket และตอบกลับในเธรดเดิมแล้ว!":"สร้าง Ticket สำเร็จ!")}else if(i.tab==="task"){const d=document.getElementById("f-title").value.trim(),b=parseInt(((s=document.getElementById("f-project"))==null?void 0:s.value)||"0"),$=document.getElementById("f-due-date").value,w=document.getElementById("f-note").value.trim(),T=document.getElementById("f-assigned-email").value,B=i.agents.find(y=>y.email===T);if(!b){S("กรุณาเลือก Project","error");return}const j=await U("PM_Tasks",{Title:d,DueDate:$||null,TaskNote:w,AssignedTo:(B==null?void 0:B.name)??i.account.name??i.account.username,AssignedEmail:T,IsCompleted:!1,IsAcknowledged:!1,ProjectID:b});if(i.droppedFiles.length>0&&await le("PM_Tasks",j,i.droppedFiles),await re("PM_Tasks",j),await ce("PM_Tasks",j),i.droppedFiles=[],await J({recipients:[T],title:`📋 ได้รับมอบหมาย Task: ${d}`,message:w||($?`กำหนดส่ง ${$}`:"มี Task ใหม่"),linkPath:b?`/projects/${b}`:"/my-work",eventType:"task_assigned"}),((c=document.getElementById("f-teams"))==null?void 0:c.checked)&&$){const y=Array.from(document.querySelectorAll(".att-internal:checked")).map(C=>C.value),M=(((h=document.getElementById("f-ext-att"))==null?void 0:h.value)||"").split(/[,;\s]+/).map(C=>C.trim()).filter(Boolean),m=`${$}T09:00:00`,x=`${$}T10:00:00`;try{await it({subject:d,start:m,end:x,body:w,attendees:[...y,...M],isOnlineMeeting:!0}),S("สร้าง Task + นัดประชุม Teams สำเร็จ!")}catch(C){S("สร้าง Task แล้ว แต่สร้างนัดประชุมไม่สำเร็จ: "+(C instanceof Error?C.message:""),"error")}}else S("สร้าง Task สำเร็จ!")}else if(i.tab==="incident"){const d=document.getElementById("f-title").value.trim(),b=parseInt(((p=document.getElementById("f-project"))==null?void 0:p.value)||"0"),$=document.getElementById("f-description").value.trim(),w=document.getElementById("f-severity").value,T=document.getElementById("f-assigned-email").value,B=i.agents.find(x=>x.email===T),j=document.getElementById("f-status").value,I=document.getElementById("f-incident-date").value,y=document.getElementById("f-resolution").value.trim();if(!b){S("กรุณาเลือก Project","error");return}const M=parseInt(((u=document.getElementById("f-sla"))==null?void 0:u.value)||"0")||null,m=await U("PM_Incidents",{Title:d,Description:$||void 0,Severity:w,Status:j,AssignedTo:(B==null?void 0:B.name)??i.account.name??i.account.username,AssignedEmail:T,ProjectID:b,IncidentDate:I||_e(),Resolution:y||void 0,SLAHours:M,SLADue:Ie(M),...j==="Resolved"?{ResolvedDate:new Date().toISOString()}:{}});i.droppedFiles.length>0&&await le("PM_Incidents",m,i.droppedFiles),await re("PM_Incidents",m),await ce("PM_Incidents",m),i.droppedFiles=[],await J({recipients:[T],title:`🚨 ได้รับมอบหมาย Incident: ${d}`,message:`ความรุนแรง ${w}${$?" — "+$.slice(0,120):""}`,linkPath:b?`/projects/${b}`:"/my-work",eventType:"incident_created"}),S("สร้าง Incident สำเร็จ!")}else if(i.tab==="comment"){const d=parseInt(((A=document.getElementById("f-ticket"))==null?void 0:A.value)||"0"),b=document.getElementById("f-comment").value.trim(),$=document.getElementById("f-comment-type").value;if(!d){S("กรุณาเลือก Ticket","error");return}if(!b){S("กรุณาพิมพ์ Comment","error");return}const w=await U("HD_TicketComments",{Title:b.slice(0,100),TicketID:d,CommentText:b,CommentType:$,CommentDate:new Date().toISOString()});await Z("HD_TicketComments",w),i.droppedFiles=[];try{const T=await N(),B=`${H}/_api/web/lists/getbytitle('HD_Tickets')/items(${d})?$select=TicketNumber,Title,AssignedEmail,Author/EMail&$expand=Author`,j=await fetch(B,{headers:{Authorization:`Bearer ${T}`,Accept:"application/json;odata=nometadata"}});if(j.ok){const I=await j.json(),y=i.account.username.toLowerCase(),M=[...new Set([I.AssignedEmail,(O=I.Author)==null?void 0:O.EMail].filter(Boolean))].filter(m=>m.toLowerCase()!==y);M.length&&await J({recipients:M,title:`💬 ${((_=i.account)==null?void 0:_.name)??"มีคน"} คอมเมนต์ใน ${I.TicketNumber||"#"+d}`,message:b.slice(0,200),linkPath:`/tickets/${d}`,eventType:"comment_added"})}}catch{}S("เพิ่ม Comment สำเร็จ!")}else if(i.tab==="project"){const d=document.getElementById("f-title").value.trim(),b=document.getElementById("f-company").value.trim(),$=document.getElementById("f-group").value,w=document.getElementById("f-status").value,T=document.getElementById("f-start").value,B=document.getElementById("f-end").value,j=document.getElementById("f-description").value.trim();if(!d){S("กรุณาใส่ชื่อโครงการ","error");return}const I=await U("PM_Projects",{Title:d,Company:b||void 0,ProjectGroup:$,Progress:0,StartDate:T||void 0,EndDate:B||null,Status:w,CreatedByEmail:i.account.username,Comment:j||void 0});if($e()){const y=await U("PM_Comments",{Title:"📎 ไฟล์แนบจากอีเมล",ProjectID:I,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await Z("PM_Comments",y)}i.droppedFiles=[],S("สร้างโครงการสำเร็จ!")}else if(i.tab==="projcomment"){const d=parseInt(((P=document.getElementById("f-project"))==null?void 0:P.value)||"0"),b=document.getElementById("f-comment").value.trim(),$=document.getElementById("f-comment-type").value;if(!d){S("กรุณาเลือกโครงการ","error");return}if(!b){S("กรุณาพิมพ์ Comment","error");return}const w=await U("PM_Comments",{Title:b.slice(0,100),ProjectID:d,CommentText:b,CommentType:$,CommentDate:new Date().toISOString(),ParentID:0});await Z("PM_Comments",w),i.droppedFiles=[];try{const T=await N(),B=`${H}/_api/web/lists/getbytitle('PM_Projects')/items(${d})?$select=Title,CreatedByEmail`,j=await fetch(B,{headers:{Authorization:`Bearer ${T}`,Accept:"application/json;odata=nometadata"}});if(j.ok){const I=await j.json(),y=i.account.username.toLowerCase();I.CreatedByEmail&&I.CreatedByEmail.toLowerCase()!==y&&await J({recipients:[I.CreatedByEmail],title:`💬 ${((E=i.account)==null?void 0:E.name)??"มีคน"} คอมเมนต์ในโครงการ ${I.Title??""}`,message:b.slice(0,200),linkPath:`/projects/${d}?tab=comments`,eventType:"comment_added"})}}catch{}S("เพิ่ม Comment สำเร็จ!")}}catch(d){const b=d instanceof Error?d.message:String(d);S(`เกิดข้อผิดพลาด: ${b}`,"error")}finally{se=!1,t&&(t.disabled=!1,t.textContent="บันทึก")}}const yt={ticket:{label:"Ticket",icon:"🎫"},comment:{label:"Comment",icon:"💬"},project:{label:"Project",icon:"📁"},task:{label:"Task",icon:"✅"},incident:{label:"Incident",icon:"🚨"},projcomment:{label:"Comment",icon:"💬"},phish:{label:"PhishGuard",icon:"🛡️"}},xt=[{title:"🎫 Helpdesk",tabs:["ticket","comment"]},{title:"📁 Project",tabs:["project","task","incident","projcomment"]},{title:"🛡️ Security",tabs:["phish"]}],Re=["f-title","f-description","f-priority","f-customer-email","f-cc","f-assigned-email","f-project","f-due-date","f-note","f-severity","f-status","f-sla","f-incident-date","f-resolution","f-ticket","f-comment","f-comment-type","f-company","f-group","f-start","f-end","f-ext-att","f-attach-eml"];let q={};function vt(){for(const e of Re){const a=document.getElementById(e);a&&(q[e]=a.value)}const t=document.getElementById("f-teams");t&&(q["f-teams"]=t.checked)}function $t(){for(const e of Re){const a=document.getElementById(e);a&&q[e]!==void 0&&q[e]!==""&&(a.value=q[e])}const t=document.getElementById("f-teams");if(t&&q["f-teams"]!==void 0){t.checked=q["f-teams"];const e=document.getElementById("teams-fields");e&&(e.style.display=t.checked?"block":"none")}}function z(){var T,B,j,I,y,M;const t=document.getElementById("app");if(!t)return;vt();const{account:e,tab:a,emailSubject:n,emailSenderName:o,emailSenderEmail:r,emailBodyPreview:s}=i,c=e!==null,h=`
    <div class="bg-gradient-to-r from-blue-700 to-blue-600 text-white px-4 py-2.5 flex items-center gap-2.5 shadow flex-shrink-0">
      <div class="w-7 h-7 bg-white/95 rounded-lg flex items-center justify-center flex-shrink-0">
        <span class="text-blue-700 font-extrabold text-xs">iT</span>
      </div>
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm leading-tight truncate">iT Services Helpdesk</div>
        ${c?`<div class="text-[10px] text-blue-100 truncate">${g((e==null?void 0:e.name)??(e==null?void 0:e.username)??"")}</div>`:""}
      </div>
      ${c?`<button id="btn-logout" title="ออกจากระบบ" class="p-1.5 rounded-lg hover:bg-white/15 transition flex-shrink-0">
             <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
           </button>`:`<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded-lg transition flex-shrink-0">
             เข้าสู่ระบบ
           </button>`}
    </div>
  `;if(!c){t.innerHTML=`
      ${h}
      <div class="flex flex-col items-center justify-center px-6 py-16 gap-4 text-center">
        <div class="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center text-3xl">🔐</div>
        <h2 class="text-base font-semibold text-slate-700">เข้าสู่ระบบด้วย Microsoft</h2>
        <p class="text-xs text-slate-500">เพื่อสร้าง Ticket / Task / Incident จาก Email นี้</p>
        <button id="btn-login-main"
          class="mt-2 w-full bg-blue-700 hover:bg-blue-600 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          เข้าสู่ระบบ
        </button>
      </div>
    `,(T=document.getElementById("btn-login"))==null||T.addEventListener("click",ve),(B=document.getElementById("btn-login-main"))==null||B.addEventListener("click",ve);return}const p=n?`<div class="mx-3 mt-3 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5 text-xs text-slate-700 space-y-1">
         <div class="font-semibold text-blue-800 truncate" title="${g(n)}">📧 ${g(n)}</div>
         ${o?`<div class="text-slate-500">จาก: <span class="font-medium text-slate-700">${g(o)}</span></div>`:""}
         ${r&&r!==o?`<div class="text-slate-400 truncate">${g(r)}</div>`:""}
       </div>`:`<div class="mx-3 mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-700">
         ⚠️ ไม่พบข้อมูล Email (โหมดทดสอบ)
       </div>`,u=i.signatureContact,A=!!r&&i.contactEmails.includes(r.toLowerCase()),O=u?`<div class="mx-3 mt-3 bg-orange-50 border border-orange-200 rounded-lg px-3 py-2.5 text-xs text-slate-700">
         <div class="flex items-center justify-between mb-2">
           <span class="font-semibold text-orange-700">👤 ข้อมูลผู้ส่ง (จาก Signature)</span>
         </div>
         <div class="space-y-0.5 mb-2.5">
           ${o?`<div><span class="text-slate-400">ชื่อ:</span> <span class="font-medium">${g(o)}</span></div>`:""}
           ${u.company?`<div><span class="text-slate-400">บริษัท:</span> ${g(u.company)}</div>`:""}
           ${r?`<div><span class="text-slate-400">Email:</span> ${g(r)}</div>`:""}
           ${u.phone?`<div><span class="text-slate-400">โทร:</span> ${g(u.phone)}</div>`:""}
         </div>
         ${A?'<div class="w-full bg-green-100 text-green-700 text-xs font-semibold py-1.5 rounded-md text-center">✓ ลูกค้านี้มีในระบบแล้ว</div>':`<button id="btn-import-customer"
                class="w-full bg-orange-500 hover:bg-orange-400 text-white text-xs font-semibold py-1.5 rounded-md transition">
                + เพิ่มเป็นลูกค้า
              </button>`}
       </div>`:"",_=`
    <div class="mx-3 mt-3 space-y-2">
      ${xt.map(m=>`
        <div>
          <div class="text-[10px] font-semibold text-slate-400 uppercase tracking-wide mb-1 px-0.5">${m.title}</div>
          <div class="grid grid-cols-4 gap-1">
            ${m.tabs.map(x=>{const C=yt[x];return`<button data-tab="${x}"
                class="tab-btn flex flex-col items-center gap-1 py-2 rounded-lg transition ${a===x?"bg-blue-700 text-white shadow":"text-slate-500 hover:bg-slate-100"}">
                <span class="text-base leading-none">${C.icon}</span>
                <span class="text-[9px] font-medium leading-none">${C.label}</span>
              </button>`}).join("")}
          </div>
        </div>
      `).join("")}
    </div>
  `;let P="";a==="phish"?P=tt(!!e):a==="ticket"?P=`
      ${k("Title / หัวข้อ",`<input id="f-title" type="text"
        class="${v}"
        value="${g(n)}" />`)}
      ${k("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${v} resize-none">${g(s)}</textarea>`)}
      ${k("Priority",`<select id="f-priority" class="${v}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
      ${k("Customer Email",`<input id="f-customer-email" type="email"
        class="${v}"
        value="${g(r)}" />`)}
      ${k("CC — ให้ผู้ที่อยู่ในเมลนี้รับรู้",`
        <label class="flex items-center gap-2 text-xs text-slate-600 mb-1.5 cursor-pointer">
          <input id="f-cc-enable" type="checkbox" ${i.emailCc.length?"checked":""} /> แนบผู้รับในเมลนี้เป็น CC อัตโนมัติ
        </label>
        <input id="f-cc" type="text" class="${v}" value="${g(i.emailCc.join(", "))}" placeholder="someone@company.com, boss@company.com" />`)}
      ${k("Assign ให้ Agent",oe(e.username))}
      ${k("โครงการ (ไม่บังคับ)",Y(!0))}
      ${G()}
    `:a==="task"?P=`
      ${k("ชื่อ Task *",`<input id="f-title" type="text" required
        class="${v}" value="${g(n)}" />`)}
      ${k("Project *",Y())}
      ${k("Assign ให้",oe(e.username))}
      ${k("Due Date",`<input id="f-due-date" type="date" class="${v}" />`)}
      ${k("Task Note",`<textarea id="f-note" rows="4"
        class="${v} resize-y">${g(s)}</textarea>`)}
      <label class="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer mb-1">
        <input id="f-teams" type="checkbox" class="rounded" onchange="document.getElementById('teams-fields').style.display=this.checked?'block':'none'" />
        💻 เพิ่มการประชุมออนไลน์ (Teams) — ใช้เวลา 09:00–10:00 ของวัน Due Date
      </label>
      <div id="teams-fields" style="display:none" class="space-y-2 mb-2">
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">ผู้เข้าร่วม Internal</label>
          <div class="max-h-28 overflow-y-auto border border-slate-200 rounded-md p-1.5 space-y-0.5">
            ${i.agents.map(m=>`<label class="flex items-center gap-2 text-xs text-slate-700 px-1 py-0.5 hover:bg-slate-50 rounded cursor-pointer">
              <input type="checkbox" class="att-internal" value="${g(m.email)}" /> ${g(m.name)}
            </label>`).join("")}
          </div>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">Email ภายนอก (คั่นด้วย ,)</label>
          <input id="f-ext-att" type="text" class="${v}" placeholder="someone@company.com, ..." />
        </div>
      </div>
      ${G()}
    `:a==="incident"?P=`
      ${k("ชื่อ Incident *",`<input id="f-title" type="text" required
        class="${v}" value="${g(n)}" />`)}
      ${k("Project *",Y())}
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
          ${Fe.map(m=>`<option value="${m.hours}" ${m.hours===te.Medium?"selected":""}>${m.labelTh}</option>`).join("")}
        </select>
        <p id="f-sla-hint" class="text-[11px] text-slate-400 mt-1">นับจากตอนนี้ · ครบกำหนด ${be(te.Medium)}</p>`)}
      ${k("Assign ให้ Agent",oe(e.username))}
      ${k("วันที่เกิด Incident",`<input id="f-incident-date" type="date" class="${v}" value="${_e()}" />`)}
      ${k("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${v} resize-y">${g(s)}</textarea>`)}
      ${k("วิธีแก้ไข (ถ้ามี)",`<textarea id="f-resolution" rows="2"
        class="${v} resize-y" placeholder="อธิบายวิธีแก้ไขปัญหา..."></textarea>`)}
      ${G()}
    `:a==="comment"?P=`
      ${k("เลือก Ticket *",`<select id="f-ticket" class="${v}">
        <option value="">-- เลือก Ticket ที่จะเพิ่ม Comment --</option>
        ${i.tickets.map(m=>`<option value="${m.id}">${g(m.TicketNumber||"#"+m.id)} · ${g(m.Title)}</option>`).join("")}
      </select>`)}
      ${k("ประเภท",`<select id="f-comment-type" class="${v}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${k("Comment *",`<textarea id="f-comment" rows="5"
        class="${v} resize-y" placeholder="พิมพ์ comment...">${g(s)}</textarea>`)}
      ${G()}
    `:a==="project"?P=`
      ${k("ชื่อโครงการ *",`<input id="f-title" type="text" required
        class="${v}" value="${g(n)}" />`)}
      ${k("บริษัท / ลูกค้า",`<input id="f-company" type="text" class="${v}" value="${g(((j=i.signatureContact)==null?void 0:j.company)??"")}" />`)}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">กลุ่มโครงการ</label>
          <select id="f-group" class="${v}">
            ${["Internal","External","R&D","Maintenance","อื่นๆ"].map(m=>`<option>${m}</option>`).join("")}
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${v}">
            ${["Planning","Active","On Hold","Completed","Cancelled"].map(m=>`<option>${m}</option>`).join("")}
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
        class="${v} resize-y">${g(s)}</textarea>`)}
      ${G()}
    `:a==="projcomment"&&(P=`
      ${k("เลือกโครงการ *",Y())}
      ${k("ประเภท",`<select id="f-comment-type" class="${v}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${k("Comment *",`<textarea id="f-comment" rows="5"
        class="${v} resize-y" placeholder="พิมพ์ comment...">${g(s)}</textarea>`)}
      ${G()}
    `);const E=a==="phish"?Xe():a==="comment"||a==="projcomment"?"เพิ่ม Comment":a==="project"?"สร้างโครงการ":a==="incident"?"แจ้ง Incident":a==="task"?"สร้าง Task":"สร้าง Ticket";t.innerHTML=`
    <div class="flex flex-col h-screen bg-slate-50">
      ${h}
      <div class="flex-1 overflow-y-auto">
        ${p}
        ${O}
        ${_}
        <div class="mx-3 mt-3 space-y-3 pb-3">
          ${P}
        </div>
      </div>
      <div class="border-t border-slate-200 bg-white px-3 py-2.5 flex-shrink-0">
        <button id="submit-btn"
          class="w-full bg-blue-700 hover:bg-blue-600 active:bg-blue-800 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          ${E}
        </button>
      </div>
    </div>
  `,(I=document.getElementById("btn-logout"))==null||I.addEventListener("click",rt),(y=document.getElementById("submit-btn"))==null||y.addEventListener("click",bt),(M=document.getElementById("btn-import-customer"))==null||M.addEventListener("click",ht),a==="phish"&&at();const d=document.getElementById("f-severity"),b=document.getElementById("f-sla");if(d&&b){let m=!1;const x=document.getElementById("f-sla-hint"),C=()=>{if(!x)return;const D=parseInt(b.value||"0")||null;x.textContent=D?`นับจากตอนนี้ · ครบกำหนด ${be(D)}`:"ไม่กำหนด SLA — เคสนี้จะวัดไม่ได้ในรายงาน"};b.addEventListener("change",()=>{m=!0,C()}),d.addEventListener("change",()=>{if(m)return;const D=te[d.value];D&&(b.value=String(D),C())})}document.querySelectorAll(".tab-btn").forEach(m=>{m.addEventListener("click",()=>{const x=m.dataset.tab;x&&x!==i.tab&&(i.tab=x,z(),x==="phish"&&Q())})});const $=document.getElementById("drop-zone"),w=document.getElementById("f-files");$&&w&&(w.addEventListener("change",()=>{w.files&&de(Array.from(w.files)),w.value=""}),$.addEventListener("dragover",m=>{m.preventDefault(),$.classList.add("border-blue-500","bg-blue-50")}),$.addEventListener("dragleave",()=>{$.classList.remove("border-blue-500","bg-blue-50")}),$.addEventListener("drop",m=>{var C;m.preventDefault(),$.classList.remove("border-blue-500","bg-blue-50");const x=Array.from(((C=m.dataTransfer)==null?void 0:C.files)??[]);x.length&&de(x)})),document.querySelectorAll(".remove-dropped").forEach(m=>{m.addEventListener("click",()=>{const x=parseInt(m.dataset.remove??"-1");x>=0&&(i.droppedFiles.splice(x,1),z())})}),$t()}function de(t){i.droppedFiles.push(...t),z()}document.addEventListener("paste",t=>{var n;if(!i.account)return;const e=Array.from(((n=t.clipboardData)==null?void 0:n.items)??[]),a=[];for(const o of e)if(o.kind==="file"){const r=o.getAsFile();if(r){const s=r.name&&r.name!=="image.png"?r.name:`screenshot-${new Date().toISOString().replace(/[:.]/g,"-").slice(0,19)}.png`;a.push(new File([r],s,{type:r.type}))}}a.length&&(t.preventDefault(),de(a),S(`แนบไฟล์แล้ว: ${a.map(o=>o.name).join(", ")}`))});const v="w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white";function we(t){return t<1024?`${t} B`:t<1024*1024?`${(t/1024).toFixed(0)} KB`:`${(t/1024/1024).toFixed(1)} MB`}function G(){const t=i.emailAttachments,e=i.droppedFiles,a=t.length>0?`<div class="mb-2 space-y-1">
        <p class="text-xs text-slate-500">📎 ไฟล์แนบจาก Email:</p>
        ${t.map(o=>`
          <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
            <input type="checkbox" class="email-att-cb" data-att-id="${g(o.id)}" data-att-name="${g(o.name)}" data-att-item="${o.isItem?"1":"0"}" checked />
            <span class="flex-1 truncate">${o.isItem?"📧 ":""}${g(o.name)}</span>
            <span class="text-slate-400 flex-shrink-0">${we(o.size)}</span>
          </label>`).join("")}
      </div>`:"",n=e.length>0?`<div class="mt-2 space-y-1">
        ${e.map((o,r)=>`<div class="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 rounded px-2 py-1">
            <span class="text-base">${o.type.startsWith("image/")?"🖼️":"📄"}</span>
            <span class="flex-1 truncate">${g(o.name)}</span>
            <span class="text-slate-400">${we(o.size)}</span>
            <button type="button" data-remove="${r}"
              class="remove-dropped text-red-400 hover:text-red-600 font-bold leading-none">✕</button>
          </div>`).join("")}
      </div>`:"";return`<div class="space-y-1">
    <label class="block text-xs font-medium text-slate-600">ไฟล์แนบ</label>
    <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
      <input type="checkbox" id="f-attach-eml" />
      <span class="flex-1">📧 แนบอีเมลต้นฉบับ (.eml)</span>
    </label>
    ${a}
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
    ${n}
  </div>`}function oe(t){var e;return`<select id="f-assigned-email" class="${v}">
    <option value="${g(t)}">${g(((e=i.account)==null?void 0:e.name)??t)} (ฉัน)</option>
    ${i.agents.filter(a=>a.email!==t).map(a=>`<option value="${g(a.email)}">${g(a.name)}</option>`).join("")}
  </select>`}function Y(t=!1){return i.projects.length===0?t?'<div class="text-xs text-slate-400">ไม่พบ Project ที่ Active</div>':'<div class="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-2">⚠️ ไม่พบ Project ที่ Active</div>':`<select id="f-project" class="${v}">
    <option value="">${t?"-- ไม่ผูกกับโครงการ --":"-- เลือก Project --"}</option>
    ${i.projects.map(e=>`<option value="${e.id}">${g(e.Title)}</option>`).join("")}
  </select>`}function k(t,e){return`
    <div class="space-y-1">
      <label class="block text-xs font-medium text-slate-600">${t}</label>
      ${e}
    </div>
  `}function g(t){return t.replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}async function wt(){ze({sharepointUrl:H,internalDomains:["itservices.co.th","rpaexpert.com","rpaexpert.onmicrosoft.com"],getToken:N,getGraphToken:()=>V(),account:()=>i.account?{name:i.account.name,username:i.account.username}:null,toast:(e,a)=>S(e,a??"success"),rerender:z,canWhitelist:()=>["Agent","Supervisor","Boss","Admin"].includes(i.myRole)}),await L.initialize(),await L.handleRedirectPromise();const t=L.getAllAccounts();if(t.length>0){i.account=t[0];try{await L.acquireTokenSilent({scopes:[X],account:t[0]}),await Promise.all([Be(),je(),Pe(),De()])}catch{i.account=null}}typeof Office<"u"?Office.onReady(e=>{var a;if(e.host===Office.HostType.Outlook){const n=Office.context.mailbox.item;if(n){i.emailSubject=n.subject??"";const o=n.from;o&&(i.emailSenderName=o.displayName??"",i.emailSenderEmail=o.emailAddress??"");const r=(((a=i.account)==null?void 0:a.username)??"").toLowerCase(),s=((o==null?void 0:o.emailAddress)??"").toLowerCase(),c=[...n.to??[],...n.cc??[]].map(p=>p.emailAddress).filter(Boolean);i.emailCc=[...new Set(c.map(p=>p.toLowerCase()))].filter(p=>p!==r&&p!==s);const h=n.attachments??[];i.emailAttachments=h.filter(p=>!p.isInline&&(p.attachmentType===Office.MailboxEnums.AttachmentType.File||p.attachmentType===Office.MailboxEnums.AttachmentType.Item)).map(p=>({id:p.id,name:p.attachmentType===Office.MailboxEnums.AttachmentType.Item?`${(p.name||"email").replace(/\.eml$/i,"")}.eml`:p.name,size:p.size,isItem:p.attachmentType===Office.MailboxEnums.AttachmentType.Item})),n.body.getAsync(Office.CoercionType.Html,{asyncContext:{}},p=>{if(p.status===Office.AsyncResultStatus.Succeeded){let u=function(y,M=!1){if(y.nodeType===3){const D=y.textContent??"";return M&&D.trim()===""?"":D}const m=y,x=(m.tagName??"").toLowerCase();if(_.includes(x))return"";if(x==="br")return" ";if(x==="tr"){const D=[];for(let R=0;R<m.childNodes.length;R++){const he=m.childNodes[R],ge=(he.tagName??"").toLowerCase();(ge==="td"||ge==="th")&&D.push((he.textContent??"").replace(/\s+/g," ").trim())}return D.length?D.join("	")+`
`:""}if(E.includes(x)){let D="";for(let R=0;R<m.childNodes.length;R++)D+=u(m.childNodes[R],!0);return D}let C="";for(let D=0;D<m.childNodes.length;D++)C+=u(m.childNodes[D],!1);return P.includes(x)&&(C=`
`+C.trim()+`
`),C};const A=p.value,O=new DOMParser().parseFromString(A,"text/html"),_=["style","script","head","img","meta","link","noscript"],P=["p","div","li","h1","h2","h3","h4","h5","h6","blockquote"],E=["table","thead","tbody","tfoot"],$=u(O.body??O.documentElement).replace(/[ \t]{2,}/g," ").replace(/\n[ \t]+/g,`
`).replace(/\n{3,}/g,`

`).trim().split(`
`),w=[];let T="";for(const y of $)y.trim()===""?T&&(w.push(T.trim()),T=""):y.includes("	")?(T&&(w.push(T.trim()),T=""),w.push(y)):T=T?T+" "+y.trim():y.trim();T&&w.push(T.trim());const B=w.join(`
`),j=/\n([-_]{3,}|From:\s|Best regards|Regards,|ขอแสดงความนับถือ|Sent:\s)/i,I=B.search(j);if(I>80){i.emailBodyPreview=B.slice(0,I).trim().slice(0,2e3);const y=B.slice(I).trim();i.signatureContact=ft(y)}else i.emailBodyPreview=B.trim().slice(0,2e3),i.signatureContact=null}z()});return}}Te(),z()}):(Te(),z())}function Te(){i.emailSubject="[DEV] Test Email Subject",i.emailSenderName="Test Sender",i.emailSenderEmail="test@example.com",i.emailBodyPreview="This is a placeholder email body for development mode."}wt().catch(t=>{console.error("Init error:",t);const e=document.getElementById("app");e&&(e.innerHTML=`<div class="p-4 text-red-600 text-sm">เกิดข้อผิดพลาด: ${String(t)}</div>`)});
