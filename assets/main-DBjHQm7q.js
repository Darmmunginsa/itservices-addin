import{P as _e}from"./PublicClientApplication-DLKYUtZW.js";import{a as Q,d as Le,p as Re,L as we,S as ke}from"./analyzer-CWd3MChg.js";const Te="HD_PhishingReports";let g;const l={mail:null,analysis:null,loading:!1,reporting:!1,reported:!1,showHeaders:!1,kasmTemplate:"",analysedItemId:"",safeDomains:[],safeDomainIds:{},savingDomain:""};function Me(t){g=t}const H=t=>(t??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"),K=()=>{var t,e;return(e=(t=Office.context)==null?void 0:t.mailbox)==null?void 0:e.item},ge=t=>new Promise(e=>{const a=K();if(!(a!=null&&a.body)){e("");return}a.body.getAsync(t,n=>e(n.status===Office.AsyncResultStatus.Succeeded?n.value??"":""))});function He(t){return t?t.split(",").map(e=>{const a=e.trim(),n=a.match(/^(.*?)\s*<([^>]+)>$/);return n?{name:n[1].replace(/^"|"$/g,"").trim(),email:n[2].trim()}:{name:"",email:a.replace(/[<>]/g,"").trim()}}).filter(e=>e.email.includes("@")):[]}function Ne(){return new Promise(t=>{let e=!1;try{e=Office.context.requirements.isSetSupported("Mailbox","1.8")}catch{e=!1}const a=K();if(!e||typeof(a==null?void 0:a.getAllInternetHeadersAsync)!="function"){t({});return}try{a.getAllInternetHeadersAsync(n=>t(n.status===Office.AsyncResultStatus.Succeeded?Re(n.value??""):{}))}catch{t({})}})}async function Fe(){try{const t=K();if(!(t!=null&&t.itemId))return{};const e=Office.context.mailbox.convertToRestId(t.itemId,Office.MailboxEnums.RestVersion.v2_0),a=await g.getGraphToken(),n=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${e}?$select=internetMessageHeaders`,{headers:{Authorization:`Bearer ${a}`}});if(!n.ok)return{};const o=await n.json(),r={};for(const s of o.internetMessageHeaders??[])r[s.name]=r[s.name]?`${r[s.name]}
${s.value}`:s.value;return r}catch{return{}}}async function de(t,e){const a=await g.getToken(),n=await fetch(`${g.sharepointUrl}/_api/web/lists/getbytitle('${t}')/items?${e}`,{headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata"}});return n.ok?(await n.json()).value:[]}async function ze(){try{return(await de("HD_AgentProfiles","$select=Title,EmailText&$top=500")).filter(e=>e.EmailText).map(e=>({name:e.Title,email:e.EmailText}))}catch{return[]}}async function Ue(){var t;try{return(((t=(await de("HD_Options","$select=Title&$filter=Category eq 'KasmConfig'&$top=1"))[0])==null?void 0:t.Title)??"").trim()}catch{return""}}const Ee="SafeDomain";async function me(){try{const t=await de("HD_Options",`$select=Id,Title&$filter=Category eq '${Ee}'&$top=500`),e={},a=[];for(const n of t){const o=(n.Title??"").trim().toLowerCase();o&&(e[o]=n.Id,a.push(o))}l.safeDomains=a,l.safeDomainIds=e}catch{}}async function qe(t){const e=t.trim().toLowerCase();if(!(!e||l.savingDomain)){if(!g.canWhitelist()){g.toast("ต้องเป็น Agent ขึ้นไปจึงจะยืนยันโดเมนได้","error");return}l.savingDomain=e,g.rerender();try{const a=await g.getToken(),n=await fetch(`${g.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items`,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:e,Category:Ee})});if(!n.ok)throw new Error(String(n.status));await me(),await X(!0),g.toast(`ยืนยันแล้วว่า ${e} ปลอดภัย`)}catch{g.toast("บันทึกไม่สำเร็จ","error")}finally{l.savingDomain="",g.rerender()}}}async function Ge(t){const e=t.trim().toLowerCase(),a=l.safeDomainIds[e];if(!(!a||l.savingDomain)){if(!g.canWhitelist()){g.toast("ต้องเป็น Agent ขึ้นไปจึงจะถอนโดเมนได้","error");return}l.savingDomain=e,g.rerender();try{const n=await g.getToken(),o=await fetch(`${g.sharepointUrl}/_api/web/lists/getbytitle('HD_Options')/items(${a})`,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","IF-MATCH":"*","X-HTTP-Method":"DELETE"}});if(!o.ok)throw new Error(String(o.status));await me(),await X(!0),g.toast(`ถอน ${e} ออกจากรายการปลอดภัยแล้ว`)}catch{g.toast("ถอนไม่สำเร็จ","error")}finally{l.savingDomain="",g.rerender()}}}async function X(t=!1){var b,h;const e=K(),a=(e==null?void 0:e.itemId)??"";if(!t&&l.analysedItemId===a&&l.analysis)return;l.loading=!0,l.reported=!1,l.analysedItemId=a,g.rerender();const[n,o]=await Promise.all([ge(Office.CoercionType.Html),ge(Office.CoercionType.Text)]),r={fromName:((b=e==null?void 0:e.from)==null?void 0:b.displayName)??"",fromEmail:((h=e==null?void 0:e.from)==null?void 0:h.emailAddress)??"",replyTo:[],subject:(e==null?void 0:e.subject)??"",bodyHtml:n,bodyText:o,attachments:((e==null?void 0:e.attachments)??[]).map(m=>({name:m.name,size:m.size??0,isInline:!!m.isInline})),headers:{},internalDomains:g.internalDomains,internalPeople:[],safeDomains:l.safeDomains},s=m=>{var w;return{...r,headers:m,replyTo:He(((w=Object.entries(m).find(([C])=>C.toLowerCase()==="reply-to"))==null?void 0:w[1])??"")}},c=await Ne();if(l.mail=s(c),l.analysis=Q(l.mail),l.loading=!1,g.rerender(),g.account()){const[m,w]=await Promise.all([ze(),Object.keys(c).length?Promise.resolve(c):Fe()]);l.mail={...s(w),internalPeople:m},l.analysis=Q(l.mail),l.kasmTemplate||(l.kasmTemplate=await Ue()),l.safeDomains.length||(await me(),l.safeDomains.length&&(l.mail={...l.mail,safeDomains:l.safeDomains},l.analysis=Q(l.mail))),g.rerender()}}function Ie(){const t=l.mail,e=l.analysis;return!t||!e?"":[`ผู้ส่ง: ${t.fromName} <${t.fromEmail}>`,`หัวข้อ: ${t.subject}`,t.replyTo.length?`Reply-To: ${t.replyTo.map(a=>a.email).join(", ")}`:"",`คะแนนความเสี่ยง: ${e.score} (${we[e.level].label})`,"","สิ่งที่ตรวจพบ:",...e.findings.map(a=>`- [${ke[a.severity].label}] (${a.category}) ${a.title} — ${a.detail.replace(/\n/g," ")}`),"",e.links.length?"ลิงก์ในอีเมล:":"",...e.links.map(a=>`- ${a.href}${a.flags.length?`  ! ${a.flags.join(" / ")}`:""}`)].filter(a=>a!=="").join(`
`)}async function Ve(t){try{const e=K();if(!(e!=null&&e.itemId))return!1;const a=Office.context.mailbox.convertToRestId(e.itemId,Office.MailboxEnums.RestVersion.v2_0),n=await g.getGraphToken(),o=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${a}/$value`,{headers:{Authorization:`Bearer ${n}`}});if(!o.ok)return!1;const r=await o.arrayBuffer(),s=(e.subject||"phishing").replace(/[\\/:*?"<>|#%&{}~]/g,"_").replace(/^_+/,"").slice(0,80).trim()||"phishing",c=await g.getToken();return(await fetch(`${g.sharepointUrl}/_api/web/lists/getbytitle('${Te}')/items(${t})/AttachmentFiles/add(FileName='${encodeURIComponent(s+".eml")}')`,{method:"POST",headers:{Authorization:`Bearer ${c}`,Accept:"application/json;odata=nometadata"},body:r})).ok}catch{return!1}}async function be(t){const e=await g.getToken(),a=await fetch(`${g.sharepointUrl}/_api/web/lists/getbytitle('${Te}')/items`,{method:"POST",headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(t)});if(!a.ok)throw new Error(`SharePoint ${a.status}: ${await a.text()}`);return(await a.json()).Id}async function Ke(){if(!(!l.mail||!l.analysis||l.reporting)){if(!g.account()){g.toast("กรุณาเข้าสู่ระบบก่อนรายงาน","error");return}l.reporting=!0,g.rerender();try{const t=l.mail,e=l.analysis,a=g.account(),n={Title:(t.subject||"(ไม่มีหัวข้อ)").slice(0,255),SenderName:t.fromName.slice(0,255),SenderEmail:t.fromEmail.slice(0,255),SenderDomain:Le(t.fromEmail),RiskScore:e.score,RiskLevel:e.level,Findings:Ie(),LinkCount:e.links.length,SuspiciousLinks:e.links.filter(c=>c.flags.length).map(c=>c.href).join(`
`).slice(0,4e3),ReportedBy:(a==null?void 0:a.name)??"",ReportedEmail:(a==null?void 0:a.username)??"",Status:"New"};let o,r=!1;try{o=await be(n)}catch(c){o=await be({Title:n.Title,Findings:n.Findings}).catch(()=>{throw c}),r=!0}const s=await Ve(o);l.reported=!0,g.toast(r?"ส่งรายงานแล้ว แต่บันทึกได้บางคอลัมน์ — ตรวจชื่อคอลัมน์ใน HD_PhishingReports":s?"ส่งรายงานพร้อมอีเมลต้นฉบับแล้ว":"ส่งรายงานแล้ว (แนบ .eml ไม่ได้)",r?"info":"success")}catch(t){g.toast(`ส่งรายงานไม่สำเร็จ: ${t instanceof Error?t.message:String(t)}`,"error")}finally{l.reporting=!1,g.rerender()}}}const We=()=>l.reported?"✓ รายงานแล้ว":l.reporting?"กำลังส่ง…":"🚩 รายงานอีเมลนี้ให้ IT";function Je(t){var e;try{const a=(e=Office.context)==null?void 0:e.ui;if(typeof(a==null?void 0:a.openBrowserWindow)=="function"){a.openBrowserWindow(t);return}}catch{}window.open(t,"_blank","noopener,noreferrer")||g.toast("เปิดหน้าต่างไม่ได้ (ถูกบล็อก)","info")}async function Ze(t){var e;try{if((e=navigator.clipboard)!=null&&e.writeText)return await navigator.clipboard.writeText(t),!0}catch{}try{const a=document.createElement("textarea");a.value=t,a.setAttribute("readonly",""),a.style.position="fixed",a.style.opacity="0",document.body.appendChild(a),a.focus(),a.select();const n=document.execCommand("copy");return a.remove(),n}catch{return!1}}function Ye(t){if(l.loading&&!l.analysis)return`<div class="py-10 text-center text-slate-500 text-sm">
      <div class="w-7 h-7 border-[3px] border-slate-200 border-t-blue-600 rounded-full animate-spin mx-auto mb-2"></div>
      กำลังตรวจอีเมล…</div>`;const e=l.analysis,a=l.mail;if(!e||!a)return'<p class="text-sm text-slate-400 text-center py-8">เปิดอีเมลเพื่อเริ่มตรวจ</p>';const n=we[e.level],o=[...e.links.filter(s=>s.flags.length),...e.links.filter(s=>!s.flags.length&&!s.trusted),...e.links.filter(s=>s.trusted)],r=g.canWhitelist();return`
    <div class="rounded-xl border-2 ${n.cls} p-3">
      <div class="flex items-center gap-2">
        <span class="text-2xl leading-none">${n.icon}</span>
        <div class="min-w-0 flex-1">
          <div class="font-bold text-sm">${H(n.label)}</div>
          <div class="text-xs opacity-80">คะแนน ${e.score} · พบสัญญาณ ${e.findings.filter(s=>s.severity!=="info").length} ข้อ</div>
        </div>
        <button id="phish-recheck" class="text-[10px] px-2 py-1 rounded-md bg-white/70 hover:bg-white text-slate-700 font-medium">ตรวจใหม่</button>
      </div>
    </div>

    ${t?"":`<div class="bg-blue-50 border border-blue-200 rounded-xl p-2.5 text-[11px] text-blue-800">
      เข้าสู่ระบบเพื่อตรวจการปลอมเป็นพนักงาน และรายงานเข้า Helpdesk ได้
    </div>`}

    ${e.findings.length===0?'<p class="text-xs text-slate-400 text-center py-3">ไม่พบสัญญาณผิดปกติจากการตรวจอัตโนมัติ</p>':e.findings.map(s=>{const c=ke[s.severity];return`<div class="bg-white rounded-xl border border-slate-200 p-2.5">
          <div class="flex items-start gap-2">
            <span class="text-[9px] font-bold px-1.5 py-0.5 rounded-full ${c.cls} flex-shrink-0 mt-0.5">${c.label}</span>
            <div class="min-w-0 flex-1">
              <div class="text-xs font-semibold text-slate-800">${H(s.title)}</div>
              <div class="text-[11px] text-slate-500 whitespace-pre-line break-all">${H(s.detail)}</div>
              <div class="text-[9px] text-slate-400 mt-0.5">${H(s.category)}</div>
            </div>
          </div>
        </div>`}).join("")}

    ${o.length?`
    <div class="bg-white rounded-xl border border-slate-200 p-3">
      <div class="text-xs font-semibold text-slate-700 mb-2">ลิงก์ในอีเมล (${o.length})</div>
      <div class="space-y-2">
        ${o.map((s,c)=>{const b=l.savingDomain===(s.host?s.host.toLowerCase():"");return`
          <div class="rounded-lg border ${s.trusted?"border-emerald-200 bg-emerald-50/50":s.flags.length?"border-red-200 bg-red-50/50":"border-slate-100"} p-2">
            <div class="text-[11px] font-medium ${s.trusted?"text-emerald-800":s.flags.length?"text-red-700":"text-slate-700"} break-all">
              ${s.trusted?"✔ ":""}${H(s.host||s.href)}
            </div>
            ${s.trusted?`<div class="text-[10px] text-emerald-700">ทีมตรวจแล้วว่าปลอดภัย${(s.suppressed??[]).length?` · ระงับการเตือน ${s.suppressed.length} ข้อ`:""}</div>`:""}
            ${s.text&&s.text!==s.href?`<div class="text-[10px] text-slate-500 break-all">แสดงว่า: "${H(s.text.slice(0,70))}"</div>`:""}
            <div class="text-[10px] text-slate-400 break-all mt-0.5">${H(s.href.slice(0,150))}</div>
            ${s.flags.map(m=>`<div class="text-[10px] text-red-600 mt-0.5">! ${H(m)}</div>`).join("")}
            <div class="flex flex-wrap gap-1 mt-1.5">
              <button data-kasm="${c}" class="text-[10px] font-semibold px-2 py-1 rounded-md bg-slate-800 hover:bg-slate-900 text-white">
                เปิดใน Kasm
              </button>
              ${r&&s.host?s.trusted?`<button data-untrust="${H(s.host)}" ${b?"disabled":""}
                     class="text-[10px] font-semibold px-2 py-1 rounded-md border border-slate-300 text-slate-600 hover:bg-slate-100 disabled:opacity-50">
                     ${b?"...":"ถอนออกจากรายการปลอดภัย"}</button>`:`<button data-trust="${H(s.host)}" ${b?"disabled":""}
                     class="text-[10px] font-semibold px-2 py-1 rounded-md bg-emerald-600 hover:bg-emerald-700 text-white disabled:opacity-50">
                     ${b?"...":"✔ ตรวจแล้ว ปลอดภัย"}</button>`:""}
            </div>
          </div>`}).join("")}
      </div>
      <p class="text-[10px] text-slate-400 mt-2">อย่าคลิกลิงก์จากอีเมลที่ไม่มั่นใจโดยตรง</p>
    </div>`:""}

    <div class="flex gap-2">
      <button id="phish-copy" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">คัดลอกผลตรวจ</button>
      <button id="phish-headers" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">${l.showHeaders?"ซ่อน header":"ดู header"}</button>
    </div>
    ${l.showHeaders?`<pre class="bg-slate-900 text-slate-100 text-[9px] p-2 rounded-lg overflow-x-auto whitespace-pre-wrap break-all max-h-56 overflow-y-auto">${H(Object.entries(a.headers).map(([s,c])=>`${s}: ${c}`).join(`
`)||"อ่าน header ไม่ได้")}</pre>`:""}
    <p class="text-[10px] text-slate-400 text-center">วิเคราะห์ในเครื่องทั้งหมด — ไม่ส่งเนื้อหาอีเมลออกนอกองค์กร</p>
  `}function Xe(){var a,n,o,r;(a=document.getElementById("phish-recheck"))==null||a.addEventListener("click",()=>{X(!0)}),(n=document.getElementById("phish-headers"))==null||n.addEventListener("click",()=>{l.showHeaders=!l.showHeaders,g.rerender()}),(o=document.getElementById("phish-copy"))==null||o.addEventListener("click",async()=>{const s=await Ze(Ie());g.toast(s?"คัดลอกผลตรวจแล้ว":"คัดลอกไม่ได้",s?"success":"error")}),document.querySelectorAll("[data-trust]").forEach(s=>s.addEventListener("click",()=>qe(s.dataset.trust??""))),document.querySelectorAll("[data-untrust]").forEach(s=>s.addEventListener("click",()=>Ge(s.dataset.untrust??"")));const t=((r=l.analysis)==null?void 0:r.links)??[],e=[...t.filter(s=>s.flags.length),...t.filter(s=>!s.flags.length&&!s.trusted),...t.filter(s=>s.trusted)];document.querySelectorAll("[data-kasm]").forEach(s=>{s.addEventListener("click",()=>{const c=e[Number(s.dataset.kasm)];if(!c)return;if(!l.kasmTemplate){g.toast("ยังไม่ได้ตั้งค่า Kasm ใน HD_Options (Category=KasmConfig)","info");return}const b=l.kasmTemplate;Je(b.includes("{url}")?b.replace("{url}",encodeURIComponent(c.href)):b+encodeURIComponent(c.href))})})}const Qe="0bab07cf-65e6-487c-89af-c917fc1a5a13",et="d569b991-89fc-4a62-9df5-eb361abcef40",_="https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd",Y="https://rpaexpert.sharepoint.com/.default",ee=["https://graph.microsoft.com/Calendars.ReadWrite","https://graph.microsoft.com/Mail.Send","https://graph.microsoft.com/Mail.Read"],D=new _e({auth:{clientId:Qe,authority:`https://login.microsoftonline.com/${et}`,redirectUri:window.location.origin.includes("localhost")?"http://localhost:3000/":"https://darmmunginsa.github.io/itservices-addin/",navigateToLoginRequestUrl:!1},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:!1}}),tt=window.location.origin.includes("localhost")?`${window.location.origin}/auth.html`:"https://darmmunginsa.github.io/itservices-addin/auth.html";function ue(){var e,a;const t=(a=(e=Office.context)==null?void 0:e.diagnostics)==null?void 0:a.platform;return t===Office.PlatformType.iOS||t===Office.PlatformType.Android}function pe(){return new Promise((t,e)=>{Office.context.ui.displayDialogAsync(tt,{height:60,width:30,promptBeforeOpen:!1},a=>{if(a.status!==Office.AsyncResultStatus.Succeeded){e(new Error("เปิดหน้าเข้าสู่ระบบไม่ได้"));return}const n=a.value;n.addEventHandler(Office.EventType.DialogMessageReceived,o=>{n.close();const r=o.message;if(!r){e(new Error("auth message error"));return}try{const s=JSON.parse(r);s.ok?t():e(new Error(s.error||"auth failed"))}catch{e(new Error("auth message error"))}}),n.addEventHandler(Office.EventType.DialogEventReceived,()=>e(new Error("ปิดหน้าเข้าสู่ระบบก่อนเสร็จ")))})})}const i={account:null,tab:"ticket",emailSubject:"",emailBodyPreview:"",emailSenderName:"",emailSenderEmail:"",loading:!1,projects:[],agents:[],myRole:"",emailAttachments:[],signatureContact:null,droppedFiles:[],tickets:[],contactEmails:[],emailCc:[]};async function L(){const t=D.getAllAccounts();if(t.length===0)throw new Error("Not signed in");const e={scopes:[Y],account:t[0]};try{return(await D.acquireTokenSilent(e)).accessToken}catch{if(ue()){await pe();const a=D.getAllAccounts()[0];if(!a)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await D.acquireTokenSilent({scopes:[Y],account:a})).accessToken}return(await D.acquireTokenPopup(e)).accessToken}}async function V(t=!1){const e=D.getAllAccounts();if(e.length===0)throw new Error("Not signed in");const a={scopes:ee,account:e[0],forceRefresh:t};try{return(await D.acquireTokenSilent(a)).accessToken}catch{if(ue()){await pe();const o=D.getAllAccounts()[0];if(!o)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await D.acquireTokenSilent({scopes:ee,account:o})).accessToken}return(await D.acquireTokenPopup({scopes:ee,account:e[0]})).accessToken}}async function at(t){const e=await V(),a={subject:t.subject,start:{dateTime:t.start,timeZone:"Asia/Bangkok"},end:{dateTime:t.end,timeZone:"Asia/Bangkok"},body:t.body?{contentType:"HTML",content:t.body.replace(/\n/g,"<br>")}:void 0,attendees:t.attendees.filter(Boolean).map(o=>({emailAddress:{address:o},type:"required"})),isOnlineMeeting:t.isOnlineMeeting,onlineMeetingProvider:t.isOnlineMeeting?"teamsForBusiness":void 0},n=await fetch("https://graph.microsoft.com/v1.0/me/events",{method:"POST",headers:{Authorization:`Bearer ${e}`,"Content-Type":"application/json"},body:JSON.stringify(a)});if(!n.ok)throw new Error(`Calendar error ${n.status}: ${await n.text()}`)}async function Ae(){try{const t=await L(),e=`${_}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$orderby=Title asc&$top=500`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();i.projects=n.value.map(o=>({id:o.Id,Title:o.Title}))}}catch{}}async function Ce(){var t,e;try{const a=await L(),n=`${_}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText,Role&$orderby=Title asc`,o=await fetch(n,{headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata"}});if(o.ok){const r=await o.json();i.agents=r.value.map(c=>({email:c.EmailText,name:c.Title}));const s=(((t=i.account)==null?void 0:t.username)??"").toLowerCase();i.myRole=((e=r.value.find(c=>(c.EmailText??"").toLowerCase()===s))==null?void 0:e.Role)??""}}catch{}}async function Se(){try{const t=await L(),e=`${_}/_api/web/lists/getbytitle('HD_Tickets')/items?$select=Id,Title,TicketNumber,Status&$filter=Status ne 'Closed'&$orderby=Modified desc&$top=200`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();i.tickets=n.value.map(o=>({id:o.Id,Title:o.Title,TicketNumber:o.TicketNumber,Status:o.Status}))}}catch{}}async function Be(){try{const t=await L(),e=`${_}/_api/web/lists/getbytitle('HD_Contracts')/items?$select=CustomerEmail&$top=2000`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();i.contactEmails=n.value.map(o=>(o.CustomerEmail||"").trim().toLowerCase()).filter(Boolean)}}catch{}}async function ye(){const t=document.getElementById("btn-login-main"),e=document.getElementById("btn-login");t&&(t.disabled=!0,t.textContent="กำลังเข้าสู่ระบบ…"),e&&(e.disabled=!0);try{if(ue()){if(await pe(),i.account=D.getAllAccounts()[0]??null,!i.account)throw new Error("เข้าสู่ระบบไม่สำเร็จ")}else{const a=await D.loginPopup({scopes:[Y]});i.account=a.account}await Promise.all([Ae(),Ce(),Se(),Be()]),N()}catch{t&&(t.disabled=!1,t.textContent="เข้าสู่ระบบ"),e&&(e.disabled=!1)}}async function nt(){i.account&&await D.logoutPopup({account:i.account}),i.account=null,N()}async function F(t,e){const a=await L(),n=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items`,o=await fetch(n,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(e)});if(!o.ok){const s=await o.text();throw new Error(`SharePoint error ${o.status}: ${s}`)}return(await o.json()).Id}let W=null;const te="support@itservices.co.th",je="engineer@itservices.co.th";async function Pe(){if(W)return W;try{const t=await L(),e=`${_}/_api/web/lists/getbytitle('HD_EmailTemplates')/items?$select=EventKey,Subject,Body,IsEnabled&$top=50`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});return a.ok?(W=(await a.json()).value,W):[]}catch{return[]}}async function st(){var t,e;try{const a=await L(),n=`${_}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title,Category&$filter=Category eq 'EmailConfig'&$top=1`,o=await fetch(n,{headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata"}});return o.ok&&((e=(t=(await o.json()).value[0])==null?void 0:t.Title)==null?void 0:e.trim())||te}catch{return te}}function oe(t,e){return t.replace(/\{\{(\w+)\}\}/g,(a,n)=>e[n]??`{{${n}}}`)}async function ot(t,e,a,n=[]){try{const r=(await Pe()).find(u=>u.EventKey===t&&u.IsEnabled);if(!r)return;const s=oe(r.Subject||"",e),c=oe(r.Body||"",e);if(!s||!c)return;const b=u=>u.trim().toLowerCase(),h=[...new Map(a.filter(Boolean).map(u=>[b(u),u])).values()];if(h.length===0)return;const m=new Set(h.map(b)),w=t==="ticket_created"?[...n,je]:n,C=[...new Map(w.filter(Boolean).map(u=>[b(u),u])).values()].filter(u=>!m.has(b(u))),O=await st(),p=await V(),d={subject:s,body:{contentType:"HTML",content:c},toRecipients:h.map(u=>({emailAddress:{address:u}}))};C.length&&(d.ccRecipients=C.map(u=>({emailAddress:{address:u}}))),O&&(d.from={emailAddress:{address:O}}),await fetch("https://graph.microsoft.com/v1.0/me/sendMail",{method:"POST",headers:{Authorization:`Bearer ${p}`,"Content-Type":"application/json"},body:JSON.stringify({message:d,saveToSentItems:!0})})}catch{}}async function it(t,e=[]){try{const a=Office.context.mailbox.item;if(!(a!=null&&a.itemId))return!1;const n=Office.context.mailbox.convertToRestId(a.itemId,Office.MailboxEnums.RestVersion.v2_0),r={Authorization:`Bearer ${await V()}`,"Content-Type":"application/json"},s=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${n}/createReplyAll`,{method:"POST",headers:r});if(!s.ok)return!1;const c=await s.json(),b=d=>d.trim().toLowerCase(),h=c.ccRecipients??[],m=new Set(h.map(d=>b(d.emailAddress.address))),w=[...new Set(e.filter(Boolean).map(d=>d.trim()))].filter(d=>!m.has(b(d))).map(d=>({emailAddress:{address:d}})),C={body:{contentType:"HTML",content:t}};return w.length&&(C.ccRecipients=[...h,...w]),(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${c.id}`,{method:"PATCH",headers:r,body:JSON.stringify(C)})).ok?(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${c.id}/send`,{method:"POST",headers:r})).ok:!1}catch{return!1}}async function rt(t,e){const n=(await Pe()).find(r=>r.EventKey===t&&r.IsEnabled);return n&&oe(n.Body||"",e)||null}async function J(t){var r;const e=s=>s.trim().toLowerCase(),a=e(((r=i.account)==null?void 0:r.username)??""),n=new Set,o=t.recipients.filter(Boolean).filter(s=>{const c=e(s);return!c||c===a||n.has(c)?!1:(n.add(c),!0)});if(o.length!==0)try{const s=await L(),c=`${_}/_api/web/lists/getbytitle('HD_Notifications')/items`;await Promise.all(o.map(b=>fetch(c,{method:"POST",headers:{Authorization:`Bearer ${s}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:t.title.slice(0,255),RecipientEmail:b,EventType:t.eventType,Message:t.message,LinkPath:t.linkPath,IsRead:!1})})))}catch{}}async function ie(t,e){const a=document.querySelectorAll(".email-att-cb:checked");if(a.length===0)return;const n=await L();for(const o of Array.from(a)){const r=o.dataset.attId,s=o.dataset.attName,c=await new Promise((p,d)=>{Office.context.mailbox.item.getAttachmentContentAsync(r,{},u=>{u.status===Office.AsyncResultStatus.Succeeded?p(u):d(new Error(u.error.message))})}),{content:b,format:h}=c.value;let m;if(h===Office.MailboxEnums.AttachmentContentFormat.Base64){const p=atob(b);m=new Uint8Array(p.length);for(let d=0;d<p.length;d++)m[d]=p.charCodeAt(d)}else if(h===Office.MailboxEnums.AttachmentContentFormat.Eml||h===Office.MailboxEnums.AttachmentContentFormat.ICalendar)m=new TextEncoder().encode(b);else continue;const w=encodeURIComponent(s),C=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${w}')`;if(!(await fetch(C,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:m.buffer})).ok)throw new Error(`Upload ${s} failed`)}}async function ct(t){const e=`https://graph.microsoft.com/v1.0/me/messages/${t}/$value`;let a=await V(),n=await fetch(e,{headers:{Authorization:`Bearer ${a}`}});if((n.status===401||n.status===403)&&(a=await V(!0),n=await fetch(e,{headers:{Authorization:`Bearer ${a}`}})),!n.ok)throw new Error(`Graph ${n.status}`);return n.arrayBuffer()}async function lt(t){const e=await new Promise((n,o)=>{Office.context.mailbox.getCallbackTokenAsync({isRest:!0},r=>{r.status===Office.AsyncResultStatus.Succeeded?n(r.value):o(new Error("callback token failed"))})}),a=await fetch(`${Office.context.mailbox.restUrl}/v2.0/me/messages/${t}/$value`,{headers:{Authorization:`Bearer ${e}`}});if(!a.ok)throw new Error(`REST ${a.status}`);return a.arrayBuffer()}async function re(t,e){const a=document.getElementById("f-attach-eml");if(!(a!=null&&a.checked))return;const n=Office.context.mailbox.item;if(!n)return;const o=Office.context.mailbox.convertToRestId(n.itemId,Office.MailboxEnums.RestVersion.v2_0);let r,s="",c="";try{r=await ct(o)}catch(C){s=C instanceof Error?C.message:String(C);try{r=await lt(o)}catch(O){c=O instanceof Error?O.message:String(O),console.error("[eml] graph:",s,"| callback:",c),A(`ดึง .eml ไม่ได้ (Graph: ${s} / REST: ${c}) — ไฟล์อื่นบันทึกแล้ว`,"error");return}}const b=(n.subject||"email").replace(/[\\/:*?"<>|#%&{}~]/g,"_").slice(0,100).trim()||"email",h=await L(),m=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${encodeURIComponent(b+".eml")}')`;(await fetch(m,{method:"POST",headers:{Authorization:`Bearer ${h}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok||A("บันทึกไฟล์ .eml ไม่สำเร็จ","error")}async function ce(t,e,a){const n=await L();for(const o of a){const r=await o.arrayBuffer(),s=encodeURIComponent(o.name),c=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${s}')`;if(!(await fetch(c,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:r})).ok)throw new Error(`Upload ${o.name} failed`)}}function A(t,e="success"){const a=document.getElementById("toast-container");if(!a)return;const n=e==="success"?"bg-green-500":e==="error"?"bg-red-500":"bg-slate-700",o=e==="success"?"✅":e==="error"?"❌":"ℹ️",r=document.createElement("div");r.className=`toast pointer-events-auto ${n} text-white text-sm font-medium px-4 py-3 rounded-lg shadow-lg max-w-xs mx-2`,r.textContent=`${o} ${t}`,a.appendChild(r),setTimeout(()=>r.remove(),4e3)}function dt(t){const e=t.split(`
`).map(m=>m.trim()).filter(Boolean),a=/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/,n=/(\+?[\d\s()\-]{7,}(?:ext\.?\s*\d+)?)/i,o=/\b(co\.?,?\s*ltd\.?|co\.?,?\s*limited|corporation|corp\.?|บริษัท|จำกัด|holding|group|inc\.?|llc)\b/i;let r="",s="",c="";const b=[];for(const m of e)if(!/^[-_=*]{2,}$/.test(m)&&!/^(best regards|regards|sincerely|ขอแสดงความนับถือ|ด้วยความนับถือ|from|sent|thanks|thank you)/i.test(m)){if(!r){const w=m.match(a);if(w){r=w[0];continue}}if(!s){const w=m.match(n);if(w&&w[0].replace(/\D/g,"").length>=7){s=w[0].trim();continue}}if(!c&&o.test(m)){c=m;continue}m.length>=2&&m.length<=50&&!/\d{4,}/.test(m)&&b.push(m)}const h=b.find(m=>!a.test(m)&&!o.test(m))??"";return!r&&!h?null:{name:h,company:c,email:r,phone:s}}async function mt(){const t=i.signatureContact;if(!t)return;const e=(i.emailSenderEmail||"").toLowerCase();if(e&&i.contactEmails.includes(e)){A("ลูกค้านี้มีในระบบแล้ว","success"),i.signatureContact=null,N();return}const a=document.getElementById("btn-import-customer");a&&(a.disabled=!0,a.textContent="กำลังบันทึก…");try{await F("HD_Contracts",{Title:i.emailSenderName||t.name,CustomerEmail:i.emailSenderEmail,Phone:t.phone||void 0,Company:t.company||void 0,Status:"Active"}),e&&i.contactEmails.push(e),A("เพิ่มลูกค้าสำเร็จ!"),i.signatureContact=null,N()}catch(n){const o=n instanceof Error?n.message:String(n);A(`เกิดข้อผิดพลาด: ${o}`,"error"),a&&(a.disabled=!1,a.textContent="เพิ่มเป็นลูกค้า")}}function De(){return new Date().toISOString().split("T")[0]}function ut(){const t=new Date;return`HD-${`${t.getFullYear()}${String(t.getMonth()+1).padStart(2,"0")}${String(t.getDate()).padStart(2,"0")}`}-${Math.floor(Math.random()*900+100)}`}function xe(){var t;return i.droppedFiles.length>0||document.querySelectorAll(".email-att-cb:checked").length>0||(((t=document.getElementById("f-attach-eml"))==null?void 0:t.checked)??!1)}async function Z(t,e){i.droppedFiles.length>0&&await ce(t,e,i.droppedFiles),await ie(t,e),await re(t,e)}let ae=!1;async function pt(){var e,a,n,o,r,s,c,b,h,m,w,C,O;if(!i.account){A("กรุณาเข้าสู่ระบบก่อน","error");return}if(ae)return;ae=!0;const t=document.getElementById("submit-btn");t&&(t.disabled=!0,t.textContent="กำลังบันทึก…");try{if(i.tab==="phish")await Ke();else if(i.tab==="ticket"){const p=document.getElementById("f-title").value.trim(),d=document.getElementById("f-description").value.trim(),u=document.getElementById("f-priority").value,T=document.getElementById("f-customer-email").value.trim(),E=((e=document.getElementById("f-cc-enable"))==null?void 0:e.checked)??!0?(((a=document.getElementById("f-cc"))==null?void 0:a.value)||"").split(/[,;\s]+/).map(M=>M.trim()).filter(Boolean):[],x=document.getElementById("f-assigned-email").value,$=i.agents.find(M=>M.email===x),B=ut(),P=await F("HD_Tickets",{Title:p,TicketNumber:B,Description:d,Priority:u,CustomerEmail:T,CustomerName:i.emailSenderName||T,Status:"Open",AssignedEmail:x||void 0,AssignedToName:($==null?void 0:$.name)??((n=i.account)==null?void 0:n.name)??""});if(xe()){const M=await F("HD_TicketComments",{Title:"📎 ไฟล์แนบจากอีเมล",TicketID:P,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await Z("HD_TicketComments",M)}i.droppedFiles=[];const f={ticket_number:B,ticket_title:p,priority:u,category:"-",description:(d||"-").replace(/\n/g,"<br>"),customer_name:i.emailSenderName||T,assigned_name:($==null?void 0:$.name)??((o=i.account)==null?void 0:o.name)??"-",link:"https://itservices.co.th/helpdesk/"},S=[x,i.account.username,...E,je].filter(Boolean);let I=!1;const z=await rt("ticket_created",f);if(z){const M=`<div style="border-left:4px solid #2563eb;background:#eff6ff;padding:10px 14px;margin:0 0 14px;font-family:Segoe UI,sans-serif">
             <div style="font-size:15px;font-weight:700;color:#1e40af">Ticket No. ${B}</div>
             <div style="font-size:12px;color:#475569;margin-top:2px">กรุณาตอบกลับในอีเมลฉบับนี้เพื่อให้ข้อมูลอยู่ใน Ticket เดียวกัน</div>
           </div>`;I=await it(M+z,S)}I||await ot("ticket_created",f,[T],S),A(I?"สร้าง Ticket และตอบกลับในเธรดเดิมแล้ว!":"สร้าง Ticket สำเร็จ!")}else if(i.tab==="task"){const p=document.getElementById("f-title").value.trim(),d=parseInt(((r=document.getElementById("f-project"))==null?void 0:r.value)||"0"),u=document.getElementById("f-due-date").value,T=document.getElementById("f-note").value.trim(),j=document.getElementById("f-assigned-email").value,E=i.agents.find(B=>B.email===j);if(!d){A("กรุณาเลือก Project","error");return}const x=await F("PM_Tasks",{Title:p,DueDate:u||null,TaskNote:T,AssignedTo:(E==null?void 0:E.name)??i.account.name??i.account.username,AssignedEmail:j,IsCompleted:!1,IsAcknowledged:!1,ProjectID:d});if(i.droppedFiles.length>0&&await ce("PM_Tasks",x,i.droppedFiles),await ie("PM_Tasks",x),await re("PM_Tasks",x),i.droppedFiles=[],await J({recipients:[j],title:`📋 ได้รับมอบหมาย Task: ${p}`,message:T||(u?`กำหนดส่ง ${u}`:"มี Task ใหม่"),linkPath:d?`/projects/${d}`:"/my-work",eventType:"task_assigned"}),((s=document.getElementById("f-teams"))==null?void 0:s.checked)&&u){const B=Array.from(document.querySelectorAll(".att-internal:checked")).map(I=>I.value),P=(((c=document.getElementById("f-ext-att"))==null?void 0:c.value)||"").split(/[,;\s]+/).map(I=>I.trim()).filter(Boolean),f=`${u}T09:00:00`,S=`${u}T10:00:00`;try{await at({subject:p,start:f,end:S,body:T,attendees:[...B,...P],isOnlineMeeting:!0}),A("สร้าง Task + นัดประชุม Teams สำเร็จ!")}catch(I){A("สร้าง Task แล้ว แต่สร้างนัดประชุมไม่สำเร็จ: "+(I instanceof Error?I.message:""),"error")}}else A("สร้าง Task สำเร็จ!")}else if(i.tab==="incident"){const p=document.getElementById("f-title").value.trim(),d=parseInt(((b=document.getElementById("f-project"))==null?void 0:b.value)||"0"),u=document.getElementById("f-description").value.trim(),T=document.getElementById("f-severity").value,j=document.getElementById("f-assigned-email").value,E=i.agents.find(f=>f.email===j),x=document.getElementById("f-status").value,$=document.getElementById("f-incident-date").value,B=document.getElementById("f-resolution").value.trim();if(!d){A("กรุณาเลือก Project","error");return}const P=await F("PM_Incidents",{Title:p,Description:u||void 0,Severity:T,Status:x,AssignedTo:(E==null?void 0:E.name)??i.account.name??i.account.username,AssignedEmail:j,ProjectID:d,IncidentDate:$||De(),Resolution:B||void 0});i.droppedFiles.length>0&&await ce("PM_Incidents",P,i.droppedFiles),await ie("PM_Incidents",P),await re("PM_Incidents",P),i.droppedFiles=[],await J({recipients:[j],title:`🚨 ได้รับมอบหมาย Incident: ${p}`,message:`ความรุนแรง ${T}${u?" — "+u.slice(0,120):""}`,linkPath:d?`/projects/${d}`:"/my-work",eventType:"incident_created"}),A("สร้าง Incident สำเร็จ!")}else if(i.tab==="comment"){const p=parseInt(((h=document.getElementById("f-ticket"))==null?void 0:h.value)||"0"),d=document.getElementById("f-comment").value.trim(),u=document.getElementById("f-comment-type").value;if(!p){A("กรุณาเลือก Ticket","error");return}if(!d){A("กรุณาพิมพ์ Comment","error");return}const T=await F("HD_TicketComments",{Title:d.slice(0,100),TicketID:p,CommentText:d,CommentType:u,CommentDate:new Date().toISOString()});await Z("HD_TicketComments",T),i.droppedFiles=[];try{const j=await L(),E=`${_}/_api/web/lists/getbytitle('HD_Tickets')/items(${p})?$select=TicketNumber,Title,AssignedEmail,Author/EMail&$expand=Author`,x=await fetch(E,{headers:{Authorization:`Bearer ${j}`,Accept:"application/json;odata=nometadata"}});if(x.ok){const $=await x.json(),B=i.account.username.toLowerCase(),P=[...new Set([$.AssignedEmail,(m=$.Author)==null?void 0:m.EMail].filter(Boolean))].filter(f=>f.toLowerCase()!==B);P.length&&await J({recipients:P,title:`💬 ${((w=i.account)==null?void 0:w.name)??"มีคน"} คอมเมนต์ใน ${$.TicketNumber||"#"+p}`,message:d.slice(0,200),linkPath:`/tickets/${p}`,eventType:"comment_added"})}}catch{}A("เพิ่ม Comment สำเร็จ!")}else if(i.tab==="project"){const p=document.getElementById("f-title").value.trim(),d=document.getElementById("f-company").value.trim(),u=document.getElementById("f-group").value,T=document.getElementById("f-status").value,j=document.getElementById("f-start").value,E=document.getElementById("f-end").value,x=document.getElementById("f-description").value.trim();if(!p){A("กรุณาใส่ชื่อโครงการ","error");return}const $=await F("PM_Projects",{Title:p,Company:d||void 0,ProjectGroup:u,Progress:0,StartDate:j||void 0,EndDate:E||null,Status:T,CreatedByEmail:i.account.username,Comment:x||void 0});if(xe()){const B=await F("PM_Comments",{Title:"📎 ไฟล์แนบจากอีเมล",ProjectID:$,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await Z("PM_Comments",B)}i.droppedFiles=[],A("สร้างโครงการสำเร็จ!")}else if(i.tab==="projcomment"){const p=parseInt(((C=document.getElementById("f-project"))==null?void 0:C.value)||"0"),d=document.getElementById("f-comment").value.trim(),u=document.getElementById("f-comment-type").value;if(!p){A("กรุณาเลือกโครงการ","error");return}if(!d){A("กรุณาพิมพ์ Comment","error");return}const T=await F("PM_Comments",{Title:d.slice(0,100),ProjectID:p,CommentText:d,CommentType:u,CommentDate:new Date().toISOString(),ParentID:0});await Z("PM_Comments",T),i.droppedFiles=[];try{const j=await L(),E=`${_}/_api/web/lists/getbytitle('PM_Projects')/items(${p})?$select=Title,CreatedByEmail`,x=await fetch(E,{headers:{Authorization:`Bearer ${j}`,Accept:"application/json;odata=nometadata"}});if(x.ok){const $=await x.json(),B=i.account.username.toLowerCase();$.CreatedByEmail&&$.CreatedByEmail.toLowerCase()!==B&&await J({recipients:[$.CreatedByEmail],title:`💬 ${((O=i.account)==null?void 0:O.name)??"มีคน"} คอมเมนต์ในโครงการ ${$.Title??""}`,message:d.slice(0,200),linkPath:`/projects/${p}?tab=comments`,eventType:"comment_added"})}}catch{}A("เพิ่ม Comment สำเร็จ!")}}catch(p){const d=p instanceof Error?p.message:String(p);A(`เกิดข้อผิดพลาด: ${d}`,"error")}finally{ae=!1,t&&(t.disabled=!1,t.textContent="บันทึก")}}const ft={ticket:{label:"Ticket",icon:"🎫"},comment:{label:"Comment",icon:"💬"},project:{label:"Project",icon:"📁"},task:{label:"Task",icon:"✅"},incident:{label:"Incident",icon:"🚨"},projcomment:{label:"Comment",icon:"💬"},phish:{label:"PhishGuard",icon:"🛡️"}},ht=[{title:"🎫 Helpdesk",tabs:["ticket","comment"]},{title:"📁 Project",tabs:["project","task","incident","projcomment"]},{title:"🛡️ Security",tabs:["phish"]}],Oe=["f-title","f-description","f-priority","f-customer-email","f-cc","f-assigned-email","f-project","f-due-date","f-note","f-severity","f-status","f-incident-date","f-resolution","f-ticket","f-comment","f-comment-type","f-company","f-group","f-start","f-end","f-ext-att","f-attach-eml"];let q={};function gt(){for(const e of Oe){const a=document.getElementById(e);a&&(q[e]=a.value)}const t=document.getElementById("f-teams");t&&(q["f-teams"]=t.checked)}function bt(){for(const e of Oe){const a=document.getElementById(e);a&&q[e]!==void 0&&q[e]!==""&&(a.value=q[e])}const t=document.getElementById("f-teams");if(t&&q["f-teams"]!==void 0){t.checked=q["f-teams"];const e=document.getElementById("teams-fields");e&&(e.style.display=t.checked?"block":"none")}}function N(){var j,E,x,$,B,P;const t=document.getElementById("app");if(!t)return;gt();const{account:e,tab:a,emailSubject:n,emailSenderName:o,emailSenderEmail:r,emailBodyPreview:s}=i,c=e!==null,b=`
    <div class="bg-gradient-to-r from-blue-700 to-blue-600 text-white px-4 py-2.5 flex items-center gap-2.5 shadow flex-shrink-0">
      <div class="w-7 h-7 bg-white/95 rounded-lg flex items-center justify-center flex-shrink-0">
        <span class="text-blue-700 font-extrabold text-xs">iT</span>
      </div>
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm leading-tight truncate">iT Services Helpdesk</div>
        ${c?`<div class="text-[10px] text-blue-100 truncate">${y((e==null?void 0:e.name)??(e==null?void 0:e.username)??"")}</div>`:""}
      </div>
      ${c?`<button id="btn-logout" title="ออกจากระบบ" class="p-1.5 rounded-lg hover:bg-white/15 transition flex-shrink-0">
             <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
           </button>`:`<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded-lg transition flex-shrink-0">
             เข้าสู่ระบบ
           </button>`}
    </div>
  `;if(!c){t.innerHTML=`
      ${b}
      <div class="flex flex-col items-center justify-center px-6 py-16 gap-4 text-center">
        <div class="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center text-3xl">🔐</div>
        <h2 class="text-base font-semibold text-slate-700">เข้าสู่ระบบด้วย Microsoft</h2>
        <p class="text-xs text-slate-500">เพื่อสร้าง Ticket / Task / Incident จาก Email นี้</p>
        <button id="btn-login-main"
          class="mt-2 w-full bg-blue-700 hover:bg-blue-600 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          เข้าสู่ระบบ
        </button>
      </div>
    `,(j=document.getElementById("btn-login"))==null||j.addEventListener("click",ye),(E=document.getElementById("btn-login-main"))==null||E.addEventListener("click",ye);return}const h=n?`<div class="mx-3 mt-3 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5 text-xs text-slate-700 space-y-1">
         <div class="font-semibold text-blue-800 truncate" title="${y(n)}">📧 ${y(n)}</div>
         ${o?`<div class="text-slate-500">จาก: <span class="font-medium text-slate-700">${y(o)}</span></div>`:""}
         ${r&&r!==o?`<div class="text-slate-400 truncate">${y(r)}</div>`:""}
       </div>`:`<div class="mx-3 mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-700">
         ⚠️ ไม่พบข้อมูล Email (โหมดทดสอบ)
       </div>`,m=i.signatureContact,w=!!r&&i.contactEmails.includes(r.toLowerCase()),C=m?`<div class="mx-3 mt-3 bg-orange-50 border border-orange-200 rounded-lg px-3 py-2.5 text-xs text-slate-700">
         <div class="flex items-center justify-between mb-2">
           <span class="font-semibold text-orange-700">👤 ข้อมูลผู้ส่ง (จาก Signature)</span>
         </div>
         <div class="space-y-0.5 mb-2.5">
           ${o?`<div><span class="text-slate-400">ชื่อ:</span> <span class="font-medium">${y(o)}</span></div>`:""}
           ${m.company?`<div><span class="text-slate-400">บริษัท:</span> ${y(m.company)}</div>`:""}
           ${r?`<div><span class="text-slate-400">Email:</span> ${y(r)}</div>`:""}
           ${m.phone?`<div><span class="text-slate-400">โทร:</span> ${y(m.phone)}</div>`:""}
         </div>
         ${w?'<div class="w-full bg-green-100 text-green-700 text-xs font-semibold py-1.5 rounded-md text-center">✓ ลูกค้านี้มีในระบบแล้ว</div>':`<button id="btn-import-customer"
                class="w-full bg-orange-500 hover:bg-orange-400 text-white text-xs font-semibold py-1.5 rounded-md transition">
                + เพิ่มเป็นลูกค้า
              </button>`}
       </div>`:"",O=`
    <div class="mx-3 mt-3 space-y-2">
      ${ht.map(f=>`
        <div>
          <div class="text-[10px] font-semibold text-slate-400 uppercase tracking-wide mb-1 px-0.5">${f.title}</div>
          <div class="grid grid-cols-4 gap-1">
            ${f.tabs.map(S=>{const I=ft[S];return`<button data-tab="${S}"
                class="tab-btn flex flex-col items-center gap-1 py-2 rounded-lg transition ${a===S?"bg-blue-700 text-white shadow":"text-slate-500 hover:bg-slate-100"}">
                <span class="text-base leading-none">${I.icon}</span>
                <span class="text-[9px] font-medium leading-none">${I.label}</span>
              </button>`}).join("")}
          </div>
        </div>
      `).join("")}
    </div>
  `;let p="";a==="phish"?p=Ye(!!e):a==="ticket"?p=`
      ${k("Title / หัวข้อ",`<input id="f-title" type="text"
        class="${v}"
        value="${y(n)}" />`)}
      ${k("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${v} resize-none">${y(s)}</textarea>`)}
      ${k("Priority",`<select id="f-priority" class="${v}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
      ${k("Customer Email",`<input id="f-customer-email" type="email"
        class="${v}"
        value="${y(r)}" />`)}
      ${k("CC — ให้ผู้ที่อยู่ในเมลนี้รับรู้",`
        <label class="flex items-center gap-2 text-xs text-slate-600 mb-1.5 cursor-pointer">
          <input id="f-cc-enable" type="checkbox" ${i.emailCc.length?"checked":""} /> แนบผู้รับในเมลนี้เป็น CC อัตโนมัติ
        </label>
        <input id="f-cc" type="text" class="${v}" value="${y(i.emailCc.join(", "))}" placeholder="someone@company.com, boss@company.com" />`)}
      ${k("Assign ให้ Agent",ne(e.username))}
      ${G()}
    `:a==="task"?p=`
      ${k("ชื่อ Task *",`<input id="f-title" type="text" required
        class="${v}" value="${y(n)}" />`)}
      ${k("Project *",se())}
      ${k("Assign ให้",ne(e.username))}
      ${k("Due Date",`<input id="f-due-date" type="date" class="${v}" />`)}
      ${k("Task Note",`<textarea id="f-note" rows="4"
        class="${v} resize-y">${y(s)}</textarea>`)}
      <label class="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer mb-1">
        <input id="f-teams" type="checkbox" class="rounded" onchange="document.getElementById('teams-fields').style.display=this.checked?'block':'none'" />
        💻 เพิ่มการประชุมออนไลน์ (Teams) — ใช้เวลา 09:00–10:00 ของวัน Due Date
      </label>
      <div id="teams-fields" style="display:none" class="space-y-2 mb-2">
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">ผู้เข้าร่วม Internal</label>
          <div class="max-h-28 overflow-y-auto border border-slate-200 rounded-md p-1.5 space-y-0.5">
            ${i.agents.map(f=>`<label class="flex items-center gap-2 text-xs text-slate-700 px-1 py-0.5 hover:bg-slate-50 rounded cursor-pointer">
              <input type="checkbox" class="att-internal" value="${y(f.email)}" /> ${y(f.name)}
            </label>`).join("")}
          </div>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">Email ภายนอก (คั่นด้วย ,)</label>
          <input id="f-ext-att" type="text" class="${v}" placeholder="someone@company.com, ..." />
        </div>
      </div>
      ${G()}
    `:a==="incident"?p=`
      ${k("ชื่อ Incident *",`<input id="f-title" type="text" required
        class="${v}" value="${y(n)}" />`)}
      ${k("Project *",se())}
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
      ${k("Assign ให้ Agent",ne(e.username))}
      ${k("วันที่เกิด Incident",`<input id="f-incident-date" type="date" class="${v}" value="${De()}" />`)}
      ${k("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${v} resize-y">${y(s)}</textarea>`)}
      ${k("วิธีแก้ไข (ถ้ามี)",`<textarea id="f-resolution" rows="2"
        class="${v} resize-y" placeholder="อธิบายวิธีแก้ไขปัญหา..."></textarea>`)}
      ${G()}
    `:a==="comment"?p=`
      ${k("เลือก Ticket *",`<select id="f-ticket" class="${v}">
        <option value="">-- เลือก Ticket ที่จะเพิ่ม Comment --</option>
        ${i.tickets.map(f=>`<option value="${f.id}">${y(f.TicketNumber||"#"+f.id)} · ${y(f.Title)}</option>`).join("")}
      </select>`)}
      ${k("ประเภท",`<select id="f-comment-type" class="${v}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${k("Comment *",`<textarea id="f-comment" rows="5"
        class="${v} resize-y" placeholder="พิมพ์ comment...">${y(s)}</textarea>`)}
      ${G()}
    `:a==="project"?p=`
      ${k("ชื่อโครงการ *",`<input id="f-title" type="text" required
        class="${v}" value="${y(n)}" />`)}
      ${k("บริษัท / ลูกค้า",`<input id="f-company" type="text" class="${v}" value="${y(((x=i.signatureContact)==null?void 0:x.company)??"")}" />`)}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">กลุ่มโครงการ</label>
          <select id="f-group" class="${v}">
            ${["Internal","External","R&D","Maintenance","อื่นๆ"].map(f=>`<option>${f}</option>`).join("")}
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${v}">
            ${["Planning","Active","On Hold","Completed","Cancelled"].map(f=>`<option>${f}</option>`).join("")}
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
        class="${v} resize-y">${y(s)}</textarea>`)}
      ${G()}
    `:a==="projcomment"&&(p=`
      ${k("เลือกโครงการ *",se())}
      ${k("ประเภท",`<select id="f-comment-type" class="${v}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${k("Comment *",`<textarea id="f-comment" rows="5"
        class="${v} resize-y" placeholder="พิมพ์ comment...">${y(s)}</textarea>`)}
      ${G()}
    `);const d=a==="phish"?We():a==="comment"||a==="projcomment"?"เพิ่ม Comment":a==="project"?"สร้างโครงการ":a==="incident"?"แจ้ง Incident":a==="task"?"สร้าง Task":"สร้าง Ticket";t.innerHTML=`
    <div class="flex flex-col h-screen bg-slate-50">
      ${b}
      <div class="flex-1 overflow-y-auto">
        ${h}
        ${C}
        ${O}
        <div class="mx-3 mt-3 space-y-3 pb-3">
          ${p}
        </div>
      </div>
      <div class="border-t border-slate-200 bg-white px-3 py-2.5 flex-shrink-0">
        <button id="submit-btn"
          class="w-full bg-blue-700 hover:bg-blue-600 active:bg-blue-800 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          ${d}
        </button>
      </div>
    </div>
  `,($=document.getElementById("btn-logout"))==null||$.addEventListener("click",nt),(B=document.getElementById("submit-btn"))==null||B.addEventListener("click",pt),(P=document.getElementById("btn-import-customer"))==null||P.addEventListener("click",mt),a==="phish"&&Xe(),document.querySelectorAll(".tab-btn").forEach(f=>{f.addEventListener("click",()=>{const S=f.dataset.tab;S&&S!==i.tab&&(i.tab=S,N(),S==="phish"&&X())})});const u=document.getElementById("drop-zone"),T=document.getElementById("f-files");u&&T&&(T.addEventListener("change",()=>{T.files&&le(Array.from(T.files)),T.value=""}),u.addEventListener("dragover",f=>{f.preventDefault(),u.classList.add("border-blue-500","bg-blue-50")}),u.addEventListener("dragleave",()=>{u.classList.remove("border-blue-500","bg-blue-50")}),u.addEventListener("drop",f=>{var I;f.preventDefault(),u.classList.remove("border-blue-500","bg-blue-50");const S=Array.from(((I=f.dataTransfer)==null?void 0:I.files)??[]);S.length&&le(S)})),document.querySelectorAll(".remove-dropped").forEach(f=>{f.addEventListener("click",()=>{const S=parseInt(f.dataset.remove??"-1");S>=0&&(i.droppedFiles.splice(S,1),N())})}),bt()}function le(t){i.droppedFiles.push(...t),N()}document.addEventListener("paste",t=>{var n;if(!i.account)return;const e=Array.from(((n=t.clipboardData)==null?void 0:n.items)??[]),a=[];for(const o of e)if(o.kind==="file"){const r=o.getAsFile();if(r){const s=r.name&&r.name!=="image.png"?r.name:`screenshot-${new Date().toISOString().replace(/[:.]/g,"-").slice(0,19)}.png`;a.push(new File([r],s,{type:r.type}))}}a.length&&(t.preventDefault(),le(a),A(`แนบไฟล์แล้ว: ${a.map(o=>o.name).join(", ")}`))});const v="w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white";function ve(t){return t<1024?`${t} B`:t<1024*1024?`${(t/1024).toFixed(0)} KB`:`${(t/1024/1024).toFixed(1)} MB`}function G(){const t=i.emailAttachments,e=i.droppedFiles,a=t.length>0?`<div class="mb-2 space-y-1">
        <p class="text-xs text-slate-500">📎 ไฟล์แนบจาก Email:</p>
        ${t.map(o=>`
          <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
            <input type="checkbox" class="email-att-cb" data-att-id="${y(o.id)}" data-att-name="${y(o.name)}" data-att-item="${o.isItem?"1":"0"}" checked />
            <span class="flex-1 truncate">${o.isItem?"📧 ":""}${y(o.name)}</span>
            <span class="text-slate-400 flex-shrink-0">${ve(o.size)}</span>
          </label>`).join("")}
      </div>`:"",n=e.length>0?`<div class="mt-2 space-y-1">
        ${e.map((o,r)=>`<div class="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 rounded px-2 py-1">
            <span class="text-base">${o.type.startsWith("image/")?"🖼️":"📄"}</span>
            <span class="flex-1 truncate">${y(o.name)}</span>
            <span class="text-slate-400">${ve(o.size)}</span>
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
  </div>`}function ne(t){var e;return`<select id="f-assigned-email" class="${v}">
    <option value="${y(t)}">${y(((e=i.account)==null?void 0:e.name)??t)} (ฉัน)</option>
    ${i.agents.filter(a=>a.email!==t).map(a=>`<option value="${y(a.email)}">${y(a.name)}</option>`).join("")}
  </select>`}function se(){return i.projects.length===0?'<div class="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-2">⚠️ ไม่พบ Project ที่ Active</div>':`<select id="f-project" class="${v}">
    <option value="">-- เลือก Project --</option>
    ${i.projects.map(t=>`<option value="${t.id}">${y(t.Title)}</option>`).join("")}
  </select>`}function k(t,e){return`
    <div class="space-y-1">
      <label class="block text-xs font-medium text-slate-600">${t}</label>
      ${e}
    </div>
  `}function y(t){return t.replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}async function yt(){Me({sharepointUrl:_,internalDomains:["itservices.co.th","rpaexpert.com","rpaexpert.onmicrosoft.com"],getToken:L,getGraphToken:()=>V(),account:()=>i.account?{name:i.account.name,username:i.account.username}:null,toast:(e,a)=>A(e,a??"success"),rerender:N,canWhitelist:()=>["Agent","Supervisor","Boss","Admin"].includes(i.myRole)}),await D.initialize(),await D.handleRedirectPromise();const t=D.getAllAccounts();if(t.length>0){i.account=t[0];try{await D.acquireTokenSilent({scopes:[Y],account:t[0]}),await Promise.all([Ae(),Ce(),Se(),Be()])}catch{i.account=null}}typeof Office<"u"?Office.onReady(e=>{var a;if(e.host===Office.HostType.Outlook){const n=Office.context.mailbox.item;if(n){i.emailSubject=n.subject??"";const o=n.from;o&&(i.emailSenderName=o.displayName??"",i.emailSenderEmail=o.emailAddress??"");const r=(((a=i.account)==null?void 0:a.username)??"").toLowerCase(),s=((o==null?void 0:o.emailAddress)??"").toLowerCase(),c=[...n.to??[],...n.cc??[]].map(h=>h.emailAddress).filter(Boolean);i.emailCc=[...new Set(c.map(h=>h.toLowerCase()))].filter(h=>h!==r&&h!==s);const b=n.attachments??[];i.emailAttachments=b.filter(h=>!h.isInline&&(h.attachmentType===Office.MailboxEnums.AttachmentType.File||h.attachmentType===Office.MailboxEnums.AttachmentType.Item)).map(h=>({id:h.id,name:h.attachmentType===Office.MailboxEnums.AttachmentType.Item?`${(h.name||"email").replace(/\.eml$/i,"")}.eml`:h.name,size:h.size,isItem:h.attachmentType===Office.MailboxEnums.AttachmentType.Item})),n.body.getAsync(Office.CoercionType.Html,{asyncContext:{}},h=>{if(h.status===Office.AsyncResultStatus.Succeeded){let m=function(f,S=!1){if(f.nodeType===3){const R=f.textContent??"";return S&&R.trim()===""?"":R}const I=f,z=(I.tagName??"").toLowerCase();if(O.includes(z))return"";if(z==="br")return" ";if(z==="tr"){const R=[];for(let U=0;U<I.childNodes.length;U++){const fe=I.childNodes[U],he=(fe.tagName??"").toLowerCase();(he==="td"||he==="th")&&R.push((fe.textContent??"").replace(/\s+/g," ").trim())}return R.length?R.join("	")+`
`:""}if(d.includes(z)){let R="";for(let U=0;U<I.childNodes.length;U++)R+=m(I.childNodes[U],!0);return R}let M="";for(let R=0;R<I.childNodes.length;R++)M+=m(I.childNodes[R],!1);return p.includes(z)&&(M=`
`+M.trim()+`
`),M};const w=h.value,C=new DOMParser().parseFromString(w,"text/html"),O=["style","script","head","img","meta","link","noscript"],p=["p","div","li","h1","h2","h3","h4","h5","h6","blockquote"],d=["table","thead","tbody","tfoot"],j=m(C.body??C.documentElement).replace(/[ \t]{2,}/g," ").replace(/\n[ \t]+/g,`
`).replace(/\n{3,}/g,`

`).trim().split(`
`),E=[];let x="";for(const f of j)f.trim()===""?x&&(E.push(x.trim()),x=""):f.includes("	")?(x&&(E.push(x.trim()),x=""),E.push(f)):x=x?x+" "+f.trim():f.trim();x&&E.push(x.trim());const $=E.join(`
`),B=/\n([-_]{3,}|From:\s|Best regards|Regards,|ขอแสดงความนับถือ|Sent:\s)/i,P=$.search(B);if(P>80){i.emailBodyPreview=$.slice(0,P).trim().slice(0,2e3);const f=$.slice(P).trim();i.signatureContact=dt(f)}else i.emailBodyPreview=$.trim().slice(0,2e3),i.signatureContact=null}N()});return}}$e(),N()}):($e(),N())}function $e(){i.emailSubject="[DEV] Test Email Subject",i.emailSenderName="Test Sender",i.emailSenderEmail="test@example.com",i.emailBodyPreview="This is a placeholder email body for development mode."}yt().catch(t=>{console.error("Init error:",t);const e=document.getElementById("app");e&&(e.innerHTML=`<div class="p-4 text-red-600 text-sm">เกิดข้อผิดพลาด: ${String(t)}</div>`)});
