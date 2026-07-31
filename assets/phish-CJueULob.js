import{P as M}from"./PublicClientApplication-DLKYUtZW.js";import{a as S,L as H,S as j,p as D,d as q}from"./analyzer-B-LZi1nK.js";const z="0bab07cf-65e6-487c-89af-c917fc1a5a13",U="d569b991-89fc-4a62-9df5-eb361abcef40",b="https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd",w="https://rpaexpert.sharepoint.com/.default",R=["https://graph.microsoft.com/Mail.Read"],$="HD_PhishingReports",K=["itservices.co.th","rpaexpert.com","rpaexpert.onmicrosoft.com"],C=window.location.origin.includes("localhost")?window.location.origin+"/":"https://darmmunginsa.github.io/itservices-addin/",l=new M({auth:{clientId:z,authority:`https://login.microsoftonline.com/${U}`,redirectUri:C,navigateToLoginRequestUrl:!1},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:!1}}),F=`${C}auth.html`;function T(){var t,a;const e=(a=(t=Office.context)==null?void 0:t.diagnostics)==null?void 0:a.platform;return e===Office.PlatformType.iOS||e===Office.PlatformType.Android}function E(){return new Promise((e,t)=>{Office.context.ui.displayDialogAsync(F,{height:60,width:30,promptBeforeOpen:!1},a=>{if(a.status!==Office.AsyncResultStatus.Succeeded){t(new Error("เปิดหน้าเข้าสู่ระบบไม่ได้"));return}const s=a.value;s.addEventHandler(Office.EventType.DialogMessageReceived,o=>{s.close();const r=o.message;if(!r){t(new Error("auth message error"));return}try{const c=JSON.parse(r);c.ok?e():t(new Error(c.error||"auth failed"))}catch{t(new Error("auth message error"))}}),s.addEventHandler(Office.EventType.DialogEventReceived,()=>t(new Error("ปิดหน้าเข้าสู่ระบบก่อนเสร็จ")))})})}async function y(){const e=l.getAllAccounts();if(!e.length)throw new Error("Not signed in");const t={scopes:[w],account:e[0]};try{return(await l.acquireTokenSilent(t)).accessToken}catch{if(T()){await E();const a=l.getAllAccounts()[0];if(!a)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await l.acquireTokenSilent({scopes:[w],account:a})).accessToken}return(await l.acquireTokenPopup(t)).accessToken}}async function _(){const e=l.getAllAccounts();if(!e.length)throw new Error("Not signed in");const t={scopes:R,account:e[0]};try{return(await l.acquireTokenSilent(t)).accessToken}catch{if(T()){await E();const a=l.getAllAccounts()[0];if(!a)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await l.acquireTokenSilent({scopes:R,account:a})).accessToken}return(await l.acquireTokenPopup(t)).accessToken}}const n={account:null,mail:null,analysis:null,loading:!0,reporting:!1,reported:!1,kasmTemplate:"",showHeaders:!1,headersLoaded:!1},d=e=>(e??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;");function h(e,t="success"){const a=document.getElementById("toast-container");if(!a)return;const s=t==="error"?"bg-red-600":t==="info"?"bg-slate-700":"bg-emerald-600",o=document.createElement("div");o.className=`toast ${s} text-white text-xs px-3 py-2 rounded-lg shadow-lg pointer-events-auto max-w-[90%]`,o.textContent=e,a.appendChild(o),setTimeout(()=>o.remove(),4e3)}async function L(e,t){const a=await y(),s=`${b}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items`,o=await fetch(s,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(t)});if(!o.ok)throw new Error(`SharePoint ${o.status}: ${await o.text()}`);return(await o.json()).Id}async function G(){try{const e=await y(),t=`${b}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText&$top=500`,a=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});return a.ok?(await a.json()).value.filter(o=>o.EmailText).map(o=>({name:o.Title,email:o.EmailText})):[]}catch{return[]}}async function N(){var e;try{const t=await y(),a=`${b}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title&$filter=Category eq 'KasmConfig'&$top=1`,s=await fetch(a,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});return s.ok?(((e=(await s.json()).value[0])==null?void 0:e.Title)??"").trim():""}catch{return""}}const g=()=>{var e,t;return(t=(e=Office.context)==null?void 0:e.mailbox)==null?void 0:t.item},P=e=>new Promise(t=>{const a=g();if(!(a!=null&&a.body)){t("");return}a.body.getAsync(e,s=>t(s.status===Office.AsyncResultStatus.Succeeded?s.value??"":""))});function V(){return new Promise(e=>{let t=!1;try{t=Office.context.requirements.isSetSupported("Mailbox","1.8")}catch{t=!1}const a=g();if(!t||typeof(a==null?void 0:a.getAllInternetHeadersAsync)!="function"){e({});return}try{a.getAllInternetHeadersAsync(s=>e(s.status===Office.AsyncResultStatus.Succeeded?D(s.value??""):{}))}catch{e({})}})}async function J(){try{const e=g();if(!(e!=null&&e.itemId))return{};const t=Office.context.mailbox.convertToRestId(e.itemId,Office.MailboxEnums.RestVersion.v2_0),a=await _(),s=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${t}?$select=internetMessageHeaders`,{headers:{Authorization:`Bearer ${a}`}});if(!s.ok)return{};const o=await s.json(),r={};for(const c of o.internetMessageHeaders??[])r[c.name]=r[c.name]?`${r[c.name]}
${c.value}`:c.value;return r}catch{return{}}}function W(e){return e?e.split(",").map(t=>{const a=t.trim(),s=a.match(/^(.*?)\s*<([^>]+)>$/);return s?{name:s[1].replace(/^"|"$/g,"").trim(),email:s[2].trim()}:{name:"",email:a.replace(/[<>]/g,"").trim()}}).filter(t=>t.email.includes("@")):[]}async function Q(){var s,o;const e=g(),[t,a]=await Promise.all([P(Office.CoercionType.Html),P(Office.CoercionType.Text)]);return{fromName:((s=e==null?void 0:e.from)==null?void 0:s.displayName)??"",fromEmail:((o=e==null?void 0:e.from)==null?void 0:o.emailAddress)??"",replyTo:[],subject:(e==null?void 0:e.subject)??"",bodyHtml:t,bodyText:a,attachments:((e==null?void 0:e.attachments)??[]).map(r=>({name:r.name,size:r.size??0,isInline:!!r.isInline})),headers:{},internalDomains:K,internalPeople:[]}}async function k(){n.loading=!0,n.reported=!1,n.headersLoaded=!1,m();const e=await Q(),t=await V(),a=(s,o=e)=>{var r;return{...o,headers:s,replyTo:W(((r=Object.entries(s).find(([c])=>c.toLowerCase()==="reply-to"))==null?void 0:r[1])??"")}};if(n.mail=a(t),n.analysis=S(n.mail),n.headersLoaded=Object.keys(t).length>0,n.loading=!1,m(),n.account){const[s,o]=await Promise.all([G(),n.headersLoaded?Promise.resolve(t):J()]);n.mail={...a(o),internalPeople:s},n.analysis=S(n.mail),n.headersLoaded=Object.keys(o).length>0,m()}}function B(){const e=n.mail,t=n.analysis;return!e||!t?"":[`ผู้ส่ง: ${e.fromName} <${e.fromEmail}>`,`หัวข้อ: ${e.subject}`,e.replyTo.length?`Reply-To: ${e.replyTo.map(s=>s.email).join(", ")}`:"",`คะแนนความเสี่ยง: ${t.score} (${H[t.level].label})`,"","สิ่งที่ตรวจพบ:",...t.findings.map(s=>`- [${j[s.severity].label}] (${s.category}) ${s.title} — ${s.detail.replace(/\n/g," ")}`),"",t.links.length?"ลิงก์ในอีเมล:":"",...t.links.map(s=>`- ${s.href}${s.flags.length?`  ⚠ ${s.flags.join(" / ")}`:""}`)].filter(s=>s!=="").join(`
`)}async function X(e){try{const t=g();if(!(t!=null&&t.itemId))return!1;const a=Office.context.mailbox.convertToRestId(t.itemId,Office.MailboxEnums.RestVersion.v2_0),s=await _(),o=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${a}/$value`,{headers:{Authorization:`Bearer ${s}`}});if(!o.ok)return!1;const r=await o.arrayBuffer(),c=(t.subject||"phishing").replace(/[\\/:*?"<>|#%&{}~]/g,"_").replace(/^_+/,"").slice(0,80).trim()||"phishing",f=await y(),p=`${b}/_api/web/lists/getbytitle('${$}')/items(${e})/AttachmentFiles/add(FileName='${encodeURIComponent(c+".eml")}')`;return(await fetch(p,{method:"POST",headers:{Authorization:`Bearer ${f}`,Accept:"application/json;odata=nometadata"},body:r})).ok}catch{return!1}}async function Y(){var e,t;if(!(!n.mail||!n.analysis||n.reporting)){n.reporting=!0,m();try{const a=n.mail,s=n.analysis,o={Title:(a.subject||"(ไม่มีหัวข้อ)").slice(0,255),SenderName:a.fromName.slice(0,255),SenderEmail:a.fromEmail.slice(0,255),SenderDomain:q(a.fromEmail),RiskScore:s.score,RiskLevel:s.level,Findings:B(),LinkCount:s.links.length,SuspiciousLinks:s.links.filter(p=>p.flags.length).map(p=>p.href).join(`
`).slice(0,4e3),ReportedBy:((e=n.account)==null?void 0:e.name)??"",ReportedEmail:((t=n.account)==null?void 0:t.username)??"",Status:"New"};let r,c=!1;try{r=await L($,o)}catch(p){r=await L($,{Title:o.Title,Findings:o.Findings}).catch(()=>{throw p}),c=!0,console.warn("[PhishGuard] full payload rejected, saved core fields only:",p)}const f=await X(r);n.reported=!0,h(c?"ส่งรายงานแล้ว แต่บันทึกได้บางคอลัมน์ — ตรวจชื่อ/ชนิดคอลัมน์ในลิสต์ HD_PhishingReports":f?"ส่งรายงานพร้อมอีเมลต้นฉบับแล้ว":"ส่งรายงานแล้ว (แนบ .eml ไม่ได้)",c?"info":"success")}catch(a){h(`ส่งรายงานไม่สำเร็จ: ${a instanceof Error?a.message:String(a)}`,"error")}finally{n.reporting=!1,m()}}}function Z(e){var a;try{const s=(a=Office.context)==null?void 0:a.ui;if(typeof(s==null?void 0:s.openBrowserWindow)=="function"){s.openBrowserWindow(e);return}}catch{}window.open(e,"_blank","noopener,noreferrer")||h("เปิดหน้าต่างไม่ได้ (ถูกบล็อก) — ใช้ปุ่มคัดลอกผลตรวจแล้วเปิดเองได้","info")}function ee(e){if(!n.kasmTemplate){h("ยังไม่ได้ตั้งค่า Kasm — เพิ่มใน HD_Options (Category=KasmConfig)","info");return}const t=n.kasmTemplate.includes("{url}")?n.kasmTemplate.replace("{url}",encodeURIComponent(e)):n.kasmTemplate+encodeURIComponent(e);Z(t)}async function te(e){var t;try{if((t=navigator.clipboard)!=null&&t.writeText)return await navigator.clipboard.writeText(e),!0}catch{}try{const a=document.createElement("textarea");a.value=e,a.setAttribute("readonly",""),a.style.position="fixed",a.style.top="0",a.style.opacity="0",document.body.appendChild(a),a.focus(),a.select();const s=document.execCommand("copy");return a.remove(),s}catch{return!1}}async function ae(){const e=await te(B());h(e?"คัดลอกผลวิเคราะห์แล้ว":'คัดลอกไม่ได้ — ใช้ปุ่ม "ดู header" แล้วเลือกข้อความเองได้',e?"success":"error")}async function se(){try{if(T()){if(await E(),n.account=l.getAllAccounts()[0]??null,!n.account)throw new Error("เข้าสู่ระบบไม่สำเร็จ")}else n.account=(await l.loginPopup({scopes:[w]})).account;n.kasmTemplate=await N(),await k()}catch(e){h(`เข้าสู่ระบบไม่สำเร็จ: ${e instanceof Error?e.message:String(e)}`,"error"),m()}}function m(){var v,A,O,I;const e=document.getElementById("app");if(!e)return;const{account:t,mail:a,analysis:s,loading:o}=n;if(o&&!s){e.innerHTML=`<div class="p-6 text-center text-slate-500 text-sm">
      <div class="w-8 h-8 border-[3px] border-slate-200 border-t-blue-600 rounded-full animate-spin mx-auto mb-3"></div>
      กำลังตรวจอีเมล…</div>`;return}if(!s||!a){e.innerHTML='<div class="p-6 text-sm text-slate-500">เปิดอีเมลเพื่อเริ่มตรวจ</div>';return}const r=H[s.level],c=s.links.filter(i=>i.flags.length),f=s.links.filter(i=>!i.flags.length);e.innerHTML=`
    <div class="p-3 space-y-3">

      <!-- ผลสรุป -->
      <div class="rounded-xl border-2 ${r.cls} p-3">
        <div class="flex items-center gap-2">
          <span class="text-2xl leading-none">${r.icon}</span>
          <div class="min-w-0 flex-1">
            <div class="font-bold text-sm">${d(r.label)}</div>
            <div class="text-xs opacity-80">คะแนนความเสี่ยง ${s.score} · พบสัญญาณ ${s.findings.filter(i=>i.severity!=="info").length} ข้อ</div>
          </div>
        </div>
      </div>

      <!-- ผู้ส่ง -->
      <div class="bg-white rounded-xl border border-slate-200 p-3 text-xs space-y-1">
        <div class="flex gap-2"><span class="text-slate-400 w-14 flex-shrink-0">ผู้ส่ง</span>
          <span class="font-medium text-slate-800 break-url">${d(a.fromName||"—")}</span></div>
        <div class="flex gap-2"><span class="text-slate-400 w-14 flex-shrink-0">อีเมล</span>
          <span class="text-slate-600 break-url">${d(a.fromEmail||"—")}</span></div>
        ${a.replyTo.length?`<div class="flex gap-2"><span class="text-slate-400 w-14 flex-shrink-0">Reply-To</span>
          <span class="text-slate-600 break-url">${d(a.replyTo.map(i=>i.email).join(", "))}</span></div>`:""}
        <div class="flex gap-2"><span class="text-slate-400 w-14 flex-shrink-0">หัวข้อ</span>
          <span class="text-slate-600">${d(a.subject||"—")}</span></div>
      </div>

      ${t?"":`
        <div class="bg-blue-50 border border-blue-200 rounded-xl p-3 text-xs text-blue-800">
          <p class="font-medium mb-1">เข้าสู่ระบบเพื่อตรวจให้ครบ</p>
          <p class="mb-2 opacity-80">จะได้ตรวจ SPF/DKIM/DMARC, ตรวจการปลอมเป็นพนักงาน และรายงานเข้า Helpdesk ได้</p>
          <button id="btn-login" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-medium py-1.5 rounded-lg">เข้าสู่ระบบ</button>
        </div>`}

      <!-- สิ่งที่ตรวจพบ -->
      <div class="space-y-2">
        ${s.findings.length===0?'<p class="text-xs text-slate-400 text-center py-3">ไม่พบสัญญาณผิดปกติจากการตรวจอัตโนมัติ</p>':s.findings.map(i=>{const u=j[i.severity];return`<div class="bg-white rounded-xl border border-slate-200 p-2.5">
              <div class="flex items-start gap-2">
                <span class="text-[9px] font-bold px-1.5 py-0.5 rounded-full ${u.cls} flex-shrink-0 mt-0.5">${u.label}</span>
                <div class="min-w-0 flex-1">
                  <div class="text-xs font-semibold text-slate-800">${d(i.title)}</div>
                  <div class="text-[11px] text-slate-500 whitespace-pre-line break-url">${d(i.detail)}</div>
                  <div class="text-[9px] text-slate-400 mt-0.5">${d(i.category)}</div>
                </div>
              </div>
            </div>`}).join("")}
      </div>

      <!-- ลิงก์ -->
      ${s.links.length?`
      <div class="bg-white rounded-xl border border-slate-200 p-3">
        <div class="text-xs font-semibold text-slate-700 mb-2">ลิงก์ในอีเมล (${s.links.length})</div>
        <div class="space-y-2">
          ${[...c,...f].map((i,u)=>`
            <div class="rounded-lg border ${i.flags.length?"border-red-200 bg-red-50/50":"border-slate-100"} p-2">
              <div class="text-[11px] font-medium ${i.flags.length?"text-red-700":"text-slate-700"} break-url">${d(i.host||i.href)}</div>
              ${i.text&&i.text!==i.href?`<div class="text-[10px] text-slate-500 break-url">ข้อความที่แสดง: “${d(i.text.slice(0,80))}”</div>`:""}
              <div class="text-[10px] text-slate-400 break-url mt-0.5">${d(i.href.slice(0,160))}</div>
              ${i.flags.map(x=>`<div class="text-[10px] text-red-600 mt-0.5">⚠ ${d(x)}</div>`).join("")}
              <button data-kasm="${u}" class="mt-1.5 text-[10px] font-semibold px-2 py-1 rounded-md bg-slate-800 hover:bg-slate-900 text-white">
                🛡 เปิดใน Kasm (แซนด์บ็อกซ์)
              </button>
            </div>`).join("")}
        </div>
        <p class="text-[10px] text-slate-400 mt-2">อย่าคลิกลิงก์จากอีเมลที่ไม่มั่นใจโดยตรง — เปิดผ่าน Kasm เพื่อแยกออกจากเครื่องคุณ</p>
      </div>`:""}

      <!-- ปุ่มจัดการ -->
      <div class="space-y-2 pb-4">
        <button id="btn-report" ${!t||n.reported?"disabled":""}
          class="w-full ${n.reported?"bg-emerald-600":"bg-red-600 hover:bg-red-700"} disabled:opacity-50 text-white text-sm font-semibold py-2 rounded-lg">
          ${n.reported?"✓ รายงานแล้ว":n.reporting?"กำลังส่ง…":"🚩 รายงานอีเมลนี้ให้ IT"}
        </button>
        <div class="flex gap-2">
          <button id="btn-copy" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">คัดลอกผลตรวจ</button>
          <button id="btn-headers" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">
            ${n.showHeaders?"ซ่อน header":"ดู header"}
          </button>
        </div>
        ${n.showHeaders?`
          <pre class="bg-slate-900 text-slate-100 text-[9px] p-2 rounded-lg overflow-x-auto whitespace-pre-wrap break-url max-h-64 overflow-y-auto">${d(Object.entries(a.headers).map(([i,u])=>`${i}: ${u}`).join(`
`)||(t?"อ่าน header ไม่ได้ (ตรวจสิทธิ์ Mail.Read)":"ต้องเข้าสู่ระบบก่อน"))}</pre>`:""}
        <p class="text-[10px] text-slate-400 text-center">วิเคราะห์ในเครื่องทั้งหมด — ไม่มีการส่งเนื้อหาอีเมลออกไปที่บริการภายนอก</p>
      </div>
    </div>`,(v=document.getElementById("btn-login"))==null||v.addEventListener("click",se),(A=document.getElementById("btn-report"))==null||A.addEventListener("click",Y),(O=document.getElementById("btn-copy"))==null||O.addEventListener("click",ae),(I=document.getElementById("btn-headers"))==null||I.addEventListener("click",()=>{n.showHeaders=!n.showHeaders,m()});const p=[...c,...f];document.querySelectorAll("[data-kasm]").forEach(i=>{i.addEventListener("click",()=>{const u=Number(i.dataset.kasm),x=p[u];x&&ee(x.href)})})}Office.onReady(async()=>{var e;if(!((e=Office.context)!=null&&e.mailbox)){const t=document.getElementById("app");t&&(t.innerHTML=`
      <div class="p-6 text-center">
        <div class="text-4xl mb-3">🛡️</div>
        <h1 class="text-base font-bold text-slate-800 mb-1">PhishGuard</h1>
        <p class="text-xs text-slate-500 mb-4">ส่วนเสริมตรวจอีเมลหลอกลวง (phishing) สำหรับ Outlook</p>
        <div class="bg-amber-50 border border-amber-200 rounded-xl p-3 text-xs text-amber-800 text-left">
          <p class="font-semibold mb-1">หน้านี้ต้องเปิดจากภายใน Outlook</p>
          <p class="opacity-80">เปิดอีเมลใน Outlook แล้วกดปุ่ม <b>PhishGuard</b> บนแถบเครื่องมือ —
          การเปิด URL นี้ตรง ๆ จะไม่มีอีเมลให้ตรวจ</p>
        </div>
      </div>`);return}await l.initialize();try{await l.handleRedirectPromise()}catch{}n.account=l.getAllAccounts()[0]??null,n.account&&(n.kasmTemplate=await N()),await k();try{Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged,()=>{k()})}catch{}});
