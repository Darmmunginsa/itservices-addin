import{P as Oe}from"./PublicClientApplication-DLKYUtZW.js";import{a as ue,d as _e,p as De,L as xe,S as ve}from"./analyzer-B-LZi1nK.js";const $e="HD_PhishingReports";let k;const h={mail:null,analysis:null,loading:!1,reporting:!1,reported:!1,showHeaders:!1,kasmTemplate:"",analysedItemId:""};function Le(t){k=t}const N=t=>(t??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"),K=()=>{var t,e;return(e=(t=Office.context)==null?void 0:t.mailbox)==null?void 0:e.item},pe=t=>new Promise(e=>{const a=K();if(!(a!=null&&a.body)){e("");return}a.body.getAsync(t,n=>e(n.status===Office.AsyncResultStatus.Succeeded?n.value??"":""))});function Me(t){return t?t.split(",").map(e=>{const a=e.trim(),n=a.match(/^(.*?)\s*<([^>]+)>$/);return n?{name:n[1].replace(/^"|"$/g,"").trim(),email:n[2].trim()}:{name:"",email:a.replace(/[<>]/g,"").trim()}}).filter(e=>e.email.includes("@")):[]}function Re(){return new Promise(t=>{let e=!1;try{e=Office.context.requirements.isSetSupported("Mailbox","1.8")}catch{e=!1}const a=K();if(!e||typeof(a==null?void 0:a.getAllInternetHeadersAsync)!="function"){t({});return}try{a.getAllInternetHeadersAsync(n=>t(n.status===Office.AsyncResultStatus.Succeeded?De(n.value??""):{}))}catch{t({})}})}async function He(){try{const t=K();if(!(t!=null&&t.itemId))return{};const e=Office.context.mailbox.convertToRestId(t.itemId,Office.MailboxEnums.RestVersion.v2_0),a=await k.getGraphToken(),n=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${e}?$select=internetMessageHeaders`,{headers:{Authorization:`Bearer ${a}`}});if(!n.ok)return{};const s=await n.json(),i={};for(const c of s.internetMessageHeaders??[])i[c.name]=i[c.name]?`${i[c.name]}
${c.value}`:c.value;return i}catch{return{}}}async function we(t,e){const a=await k.getToken(),n=await fetch(`${k.sharepointUrl}/_api/web/lists/getbytitle('${t}')/items?${e}`,{headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata"}});return n.ok?(await n.json()).value:[]}async function Ne(){try{return(await we("HD_AgentProfiles","$select=Title,EmailText&$top=500")).filter(e=>e.EmailText).map(e=>({name:e.Title,email:e.EmailText}))}catch{return[]}}async function Fe(){var t;try{return(((t=(await we("HD_Options","$select=Title&$filter=Category eq 'KasmConfig'&$top=1"))[0])==null?void 0:t.Title)??"").trim()}catch{return""}}async function ke(t=!1){var b,f;const e=K(),a=(e==null?void 0:e.itemId)??"";if(!t&&h.analysedItemId===a&&h.analysis)return;h.loading=!0,h.reported=!1,h.analysedItemId=a,k.rerender();const[n,s]=await Promise.all([pe(Office.CoercionType.Html),pe(Office.CoercionType.Text)]),i={fromName:((b=e==null?void 0:e.from)==null?void 0:b.displayName)??"",fromEmail:((f=e==null?void 0:e.from)==null?void 0:f.emailAddress)??"",replyTo:[],subject:(e==null?void 0:e.subject)??"",bodyHtml:n,bodyText:s,attachments:((e==null?void 0:e.attachments)??[]).map(m=>({name:m.name,size:m.size??0,isInline:!!m.isInline})),headers:{},internalDomains:k.internalDomains,internalPeople:[]},c=m=>{var $;return{...i,headers:m,replyTo:Me((($=Object.entries(m).find(([C])=>C.toLowerCase()==="reply-to"))==null?void 0:$[1])??"")}},r=await Re();if(h.mail=c(r),h.analysis=ue(h.mail),h.loading=!1,k.rerender(),k.account()){const[m,$]=await Promise.all([Ne(),Object.keys(r).length?Promise.resolve(r):He()]);h.mail={...c($),internalPeople:m},h.analysis=ue(h.mail),h.kasmTemplate||(h.kasmTemplate=await Fe()),k.rerender()}}function Te(){const t=h.mail,e=h.analysis;return!t||!e?"":[`ผู้ส่ง: ${t.fromName} <${t.fromEmail}>`,`หัวข้อ: ${t.subject}`,t.replyTo.length?`Reply-To: ${t.replyTo.map(a=>a.email).join(", ")}`:"",`คะแนนความเสี่ยง: ${e.score} (${xe[e.level].label})`,"","สิ่งที่ตรวจพบ:",...e.findings.map(a=>`- [${ve[a.severity].label}] (${a.category}) ${a.title} — ${a.detail.replace(/\n/g," ")}`),"",e.links.length?"ลิงก์ในอีเมล:":"",...e.links.map(a=>`- ${a.href}${a.flags.length?`  ! ${a.flags.join(" / ")}`:""}`)].filter(a=>a!=="").join(`
`)}async function ze(t){try{const e=K();if(!(e!=null&&e.itemId))return!1;const a=Office.context.mailbox.convertToRestId(e.itemId,Office.MailboxEnums.RestVersion.v2_0),n=await k.getGraphToken(),s=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${a}/$value`,{headers:{Authorization:`Bearer ${n}`}});if(!s.ok)return!1;const i=await s.arrayBuffer(),c=(e.subject||"phishing").replace(/[\\/:*?"<>|#%&{}~]/g,"_").replace(/^_+/,"").slice(0,80).trim()||"phishing",r=await k.getToken();return(await fetch(`${k.sharepointUrl}/_api/web/lists/getbytitle('${$e}')/items(${t})/AttachmentFiles/add(FileName='${encodeURIComponent(c+".eml")}')`,{method:"POST",headers:{Authorization:`Bearer ${r}`,Accept:"application/json;odata=nometadata"},body:i})).ok}catch{return!1}}async function fe(t){const e=await k.getToken(),a=await fetch(`${k.sharepointUrl}/_api/web/lists/getbytitle('${$e}')/items`,{method:"POST",headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(t)});if(!a.ok)throw new Error(`SharePoint ${a.status}: ${await a.text()}`);return(await a.json()).Id}async function Ue(){if(!(!h.mail||!h.analysis||h.reporting)){if(!k.account()){k.toast("กรุณาเข้าสู่ระบบก่อนรายงาน","error");return}h.reporting=!0,k.rerender();try{const t=h.mail,e=h.analysis,a=k.account(),n={Title:(t.subject||"(ไม่มีหัวข้อ)").slice(0,255),SenderName:t.fromName.slice(0,255),SenderEmail:t.fromEmail.slice(0,255),SenderDomain:_e(t.fromEmail),RiskScore:e.score,RiskLevel:e.level,Findings:Te(),LinkCount:e.links.length,SuspiciousLinks:e.links.filter(r=>r.flags.length).map(r=>r.href).join(`
`).slice(0,4e3),ReportedBy:(a==null?void 0:a.name)??"",ReportedEmail:(a==null?void 0:a.username)??"",Status:"New"};let s,i=!1;try{s=await fe(n)}catch(r){s=await fe({Title:n.Title,Findings:n.Findings}).catch(()=>{throw r}),i=!0}const c=await ze(s);h.reported=!0,k.toast(i?"ส่งรายงานแล้ว แต่บันทึกได้บางคอลัมน์ — ตรวจชื่อคอลัมน์ใน HD_PhishingReports":c?"ส่งรายงานพร้อมอีเมลต้นฉบับแล้ว":"ส่งรายงานแล้ว (แนบ .eml ไม่ได้)",i?"info":"success")}catch(t){k.toast(`ส่งรายงานไม่สำเร็จ: ${t instanceof Error?t.message:String(t)}`,"error")}finally{h.reporting=!1,k.rerender()}}}const qe=()=>h.reported?"✓ รายงานแล้ว":h.reporting?"กำลังส่ง…":"🚩 รายงานอีเมลนี้ให้ IT";function Ge(t){var e;try{const a=(e=Office.context)==null?void 0:e.ui;if(typeof(a==null?void 0:a.openBrowserWindow)=="function"){a.openBrowserWindow(t);return}}catch{}window.open(t,"_blank","noopener,noreferrer")||k.toast("เปิดหน้าต่างไม่ได้ (ถูกบล็อก)","info")}async function Ve(t){var e;try{if((e=navigator.clipboard)!=null&&e.writeText)return await navigator.clipboard.writeText(t),!0}catch{}try{const a=document.createElement("textarea");a.value=t,a.setAttribute("readonly",""),a.style.position="fixed",a.style.opacity="0",document.body.appendChild(a),a.focus(),a.select();const n=document.execCommand("copy");return a.remove(),n}catch{return!1}}function Ke(t){if(h.loading&&!h.analysis)return`<div class="py-10 text-center text-slate-500 text-sm">
      <div class="w-7 h-7 border-[3px] border-slate-200 border-t-blue-600 rounded-full animate-spin mx-auto mb-2"></div>
      กำลังตรวจอีเมล…</div>`;const e=h.analysis,a=h.mail;if(!e||!a)return'<p class="text-sm text-slate-400 text-center py-8">เปิดอีเมลเพื่อเริ่มตรวจ</p>';const n=xe[e.level],s=[...e.links.filter(i=>i.flags.length),...e.links.filter(i=>!i.flags.length)];return`
    <div class="rounded-xl border-2 ${n.cls} p-3">
      <div class="flex items-center gap-2">
        <span class="text-2xl leading-none">${n.icon}</span>
        <div class="min-w-0 flex-1">
          <div class="font-bold text-sm">${N(n.label)}</div>
          <div class="text-xs opacity-80">คะแนน ${e.score} · พบสัญญาณ ${e.findings.filter(i=>i.severity!=="info").length} ข้อ</div>
        </div>
        <button id="phish-recheck" class="text-[10px] px-2 py-1 rounded-md bg-white/70 hover:bg-white text-slate-700 font-medium">ตรวจใหม่</button>
      </div>
    </div>

    ${t?"":`<div class="bg-blue-50 border border-blue-200 rounded-xl p-2.5 text-[11px] text-blue-800">
      เข้าสู่ระบบเพื่อตรวจการปลอมเป็นพนักงาน และรายงานเข้า Helpdesk ได้
    </div>`}

    ${e.findings.length===0?'<p class="text-xs text-slate-400 text-center py-3">ไม่พบสัญญาณผิดปกติจากการตรวจอัตโนมัติ</p>':e.findings.map(i=>{const c=ve[i.severity];return`<div class="bg-white rounded-xl border border-slate-200 p-2.5">
          <div class="flex items-start gap-2">
            <span class="text-[9px] font-bold px-1.5 py-0.5 rounded-full ${c.cls} flex-shrink-0 mt-0.5">${c.label}</span>
            <div class="min-w-0 flex-1">
              <div class="text-xs font-semibold text-slate-800">${N(i.title)}</div>
              <div class="text-[11px] text-slate-500 whitespace-pre-line break-all">${N(i.detail)}</div>
              <div class="text-[9px] text-slate-400 mt-0.5">${N(i.category)}</div>
            </div>
          </div>
        </div>`}).join("")}

    ${s.length?`
    <div class="bg-white rounded-xl border border-slate-200 p-3">
      <div class="text-xs font-semibold text-slate-700 mb-2">ลิงก์ในอีเมล (${s.length})</div>
      <div class="space-y-2">
        ${s.map((i,c)=>`
          <div class="rounded-lg border ${i.flags.length?"border-red-200 bg-red-50/50":"border-slate-100"} p-2">
            <div class="text-[11px] font-medium ${i.flags.length?"text-red-700":"text-slate-700"} break-all">${N(i.host||i.href)}</div>
            ${i.text&&i.text!==i.href?`<div class="text-[10px] text-slate-500 break-all">แสดงว่า: "${N(i.text.slice(0,70))}"</div>`:""}
            <div class="text-[10px] text-slate-400 break-all mt-0.5">${N(i.href.slice(0,150))}</div>
            ${i.flags.map(r=>`<div class="text-[10px] text-red-600 mt-0.5">! ${N(r)}</div>`).join("")}
            <button data-kasm="${c}" class="mt-1.5 text-[10px] font-semibold px-2 py-1 rounded-md bg-slate-800 hover:bg-slate-900 text-white">
              เปิดใน Kasm (แซนด์บ็อกซ์)
            </button>
          </div>`).join("")}
      </div>
      <p class="text-[10px] text-slate-400 mt-2">อย่าคลิกลิงก์จากอีเมลที่ไม่มั่นใจโดยตรง</p>
    </div>`:""}

    <div class="flex gap-2">
      <button id="phish-copy" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">คัดลอกผลตรวจ</button>
      <button id="phish-headers" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">${h.showHeaders?"ซ่อน header":"ดู header"}</button>
    </div>
    ${h.showHeaders?`<pre class="bg-slate-900 text-slate-100 text-[9px] p-2 rounded-lg overflow-x-auto whitespace-pre-wrap break-all max-h-56 overflow-y-auto">${N(Object.entries(a.headers).map(([i,c])=>`${i}: ${c}`).join(`
`)||"อ่าน header ไม่ได้")}</pre>`:""}
    <p class="text-[10px] text-slate-400 text-center">วิเคราะห์ในเครื่องทั้งหมด — ไม่ส่งเนื้อหาอีเมลออกนอกองค์กร</p>
  `}function Je(){var e,a,n,s,i;(e=document.getElementById("phish-recheck"))==null||e.addEventListener("click",()=>{ke(!0)}),(a=document.getElementById("phish-headers"))==null||a.addEventListener("click",()=>{h.showHeaders=!h.showHeaders,k.rerender()}),(n=document.getElementById("phish-copy"))==null||n.addEventListener("click",async()=>{const c=await Ve(Te());k.toast(c?"คัดลอกผลตรวจแล้ว":"คัดลอกไม่ได้",c?"success":"error")});const t=[...((s=h.analysis)==null?void 0:s.links.filter(c=>c.flags.length))??[],...((i=h.analysis)==null?void 0:i.links.filter(c=>!c.flags.length))??[]];document.querySelectorAll("[data-kasm]").forEach(c=>{c.addEventListener("click",()=>{const r=t[Number(c.dataset.kasm)];if(!r)return;if(!h.kasmTemplate){k.toast("ยังไม่ได้ตั้งค่า Kasm ใน HD_Options (Category=KasmConfig)","info");return}const b=h.kasmTemplate;Ge(b.includes("{url}")?b.replace("{url}",encodeURIComponent(r.href)):b+encodeURIComponent(r.href))})})}const We="0bab07cf-65e6-487c-89af-c917fc1a5a13",Ze="d569b991-89fc-4a62-9df5-eb361abcef40",D="https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd",Y="https://rpaexpert.sharepoint.com/.default",Q=["https://graph.microsoft.com/Calendars.ReadWrite","https://graph.microsoft.com/Mail.Send","https://graph.microsoft.com/Mail.Read"],O=new Oe({auth:{clientId:We,authority:`https://login.microsoftonline.com/${Ze}`,redirectUri:window.location.origin.includes("localhost")?"http://localhost:3000/":"https://darmmunginsa.github.io/itservices-addin/",navigateToLoginRequestUrl:!1},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:!1}}),Ye=window.location.origin.includes("localhost")?`${window.location.origin}/auth.html`:"https://darmmunginsa.github.io/itservices-addin/auth.html";function re(){var e,a;const t=(a=(e=Office.context)==null?void 0:e.diagnostics)==null?void 0:a.platform;return t===Office.PlatformType.iOS||t===Office.PlatformType.Android}function le(){return new Promise((t,e)=>{Office.context.ui.displayDialogAsync(Ye,{height:60,width:30,promptBeforeOpen:!1},a=>{if(a.status!==Office.AsyncResultStatus.Succeeded){e(new Error("เปิดหน้าเข้าสู่ระบบไม่ได้"));return}const n=a.value;n.addEventHandler(Office.EventType.DialogMessageReceived,s=>{n.close();const i=s.message;if(!i){e(new Error("auth message error"));return}try{const c=JSON.parse(i);c.ok?t():e(new Error(c.error||"auth failed"))}catch{e(new Error("auth message error"))}}),n.addEventHandler(Office.EventType.DialogEventReceived,()=>e(new Error("ปิดหน้าเข้าสู่ระบบก่อนเสร็จ")))})})}const o={account:null,tab:"ticket",emailSubject:"",emailBodyPreview:"",emailSenderName:"",emailSenderEmail:"",loading:!1,projects:[],agents:[],emailAttachments:[],signatureContact:null,droppedFiles:[],tickets:[],contactEmails:[],emailCc:[]};async function L(){const t=O.getAllAccounts();if(t.length===0)throw new Error("Not signed in");const e={scopes:[Y],account:t[0]};try{return(await O.acquireTokenSilent(e)).accessToken}catch{if(re()){await le();const a=O.getAllAccounts()[0];if(!a)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await O.acquireTokenSilent({scopes:[Y],account:a})).accessToken}return(await O.acquireTokenPopup(e)).accessToken}}async function V(t=!1){const e=O.getAllAccounts();if(e.length===0)throw new Error("Not signed in");const a={scopes:Q,account:e[0],forceRefresh:t};try{return(await O.acquireTokenSilent(a)).accessToken}catch{if(re()){await le();const s=O.getAllAccounts()[0];if(!s)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await O.acquireTokenSilent({scopes:Q,account:s})).accessToken}return(await O.acquireTokenPopup({scopes:Q,account:e[0]})).accessToken}}async function Qe(t){const e=await V(),a={subject:t.subject,start:{dateTime:t.start,timeZone:"Asia/Bangkok"},end:{dateTime:t.end,timeZone:"Asia/Bangkok"},body:t.body?{contentType:"HTML",content:t.body.replace(/\n/g,"<br>")}:void 0,attendees:t.attendees.filter(Boolean).map(s=>({emailAddress:{address:s},type:"required"})),isOnlineMeeting:t.isOnlineMeeting,onlineMeetingProvider:t.isOnlineMeeting?"teamsForBusiness":void 0},n=await fetch("https://graph.microsoft.com/v1.0/me/events",{method:"POST",headers:{Authorization:`Bearer ${e}`,"Content-Type":"application/json"},body:JSON.stringify(a)});if(!n.ok)throw new Error(`Calendar error ${n.status}: ${await n.text()}`)}async function Ee(){try{const t=await L(),e=`${D}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$orderby=Title asc&$top=500`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();o.projects=n.value.map(s=>({id:s.Id,Title:s.Title}))}}catch{}}async function Ie(){try{const t=await L(),e=`${D}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText&$orderby=Title asc`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();o.agents=n.value.map(s=>({email:s.EmailText,name:s.Title}))}}catch{}}async function Ae(){try{const t=await L(),e=`${D}/_api/web/lists/getbytitle('HD_Tickets')/items?$select=Id,Title,TicketNumber,Status&$filter=Status ne 'Closed'&$orderby=Modified desc&$top=200`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();o.tickets=n.value.map(s=>({id:s.Id,Title:s.Title,TicketNumber:s.TicketNumber,Status:s.Status}))}}catch{}}async function Ce(){try{const t=await L(),e=`${D}/_api/web/lists/getbytitle('HD_Contracts')/items?$select=CustomerEmail&$top=2000`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});if(a.ok){const n=await a.json();o.contactEmails=n.value.map(s=>(s.CustomerEmail||"").trim().toLowerCase()).filter(Boolean)}}catch{}}async function he(){const t=document.getElementById("btn-login-main"),e=document.getElementById("btn-login");t&&(t.disabled=!0,t.textContent="กำลังเข้าสู่ระบบ…"),e&&(e.disabled=!0);try{if(re()){if(await le(),o.account=O.getAllAccounts()[0]??null,!o.account)throw new Error("เข้าสู่ระบบไม่สำเร็จ")}else{const a=await O.loginPopup({scopes:[Y]});o.account=a.account}await Promise.all([Ee(),Ie(),Ae(),Ce()]),H()}catch{t&&(t.disabled=!1,t.textContent="เข้าสู่ระบบ"),e&&(e.disabled=!1)}}async function Xe(){o.account&&await O.logoutPopup({account:o.account}),o.account=null,H()}async function F(t,e){const a=await L(),n=`${D}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items`,s=await fetch(n,{method:"POST",headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(e)});if(!s.ok){const c=await s.text();throw new Error(`SharePoint error ${s.status}: ${c}`)}return(await s.json()).Id}let J=null;const X="support@itservices.co.th",Se="engineer@itservices.co.th";async function Be(){if(J)return J;try{const t=await L(),e=`${D}/_api/web/lists/getbytitle('HD_EmailTemplates')/items?$select=EventKey,Subject,Body,IsEnabled&$top=50`,a=await fetch(e,{headers:{Authorization:`Bearer ${t}`,Accept:"application/json;odata=nometadata"}});return a.ok?(J=(await a.json()).value,J):[]}catch{return[]}}async function et(){var t,e;try{const a=await L(),n=`${D}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title,Category&$filter=Category eq 'EmailConfig'&$top=1`,s=await fetch(n,{headers:{Authorization:`Bearer ${a}`,Accept:"application/json;odata=nometadata"}});return s.ok&&((e=(t=(await s.json()).value[0])==null?void 0:t.Title)==null?void 0:e.trim())||X}catch{return X}}function ne(t,e){return t.replace(/\{\{(\w+)\}\}/g,(a,n)=>e[n]??`{{${n}}}`)}async function tt(t,e,a,n=[]){try{const i=(await Be()).find(d=>d.EventKey===t&&d.IsEnabled);if(!i)return;const c=ne(i.Subject||"",e),r=ne(i.Body||"",e);if(!c||!r)return;const b=d=>d.trim().toLowerCase(),f=[...new Map(a.filter(Boolean).map(d=>[b(d),d])).values()];if(f.length===0)return;const m=new Set(f.map(b)),$=t==="ticket_created"?[...n,Se]:n,C=[...new Map($.filter(Boolean).map(d=>[b(d),d])).values()].filter(d=>!m.has(b(d))),_=await et(),u=await V(),l={subject:c,body:{contentType:"HTML",content:r},toRecipients:f.map(d=>({emailAddress:{address:d}}))};C.length&&(l.ccRecipients=C.map(d=>({emailAddress:{address:d}}))),_&&(l.from={emailAddress:{address:_}}),await fetch("https://graph.microsoft.com/v1.0/me/sendMail",{method:"POST",headers:{Authorization:`Bearer ${u}`,"Content-Type":"application/json"},body:JSON.stringify({message:l,saveToSentItems:!0})})}catch{}}async function at(t,e=[]){try{const a=Office.context.mailbox.item;if(!(a!=null&&a.itemId))return!1;const n=Office.context.mailbox.convertToRestId(a.itemId,Office.MailboxEnums.RestVersion.v2_0),i={Authorization:`Bearer ${await V()}`,"Content-Type":"application/json"},c=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${n}/createReplyAll`,{method:"POST",headers:i});if(!c.ok)return!1;const r=await c.json(),b=l=>l.trim().toLowerCase(),f=r.ccRecipients??[],m=new Set(f.map(l=>b(l.emailAddress.address))),$=[...new Set(e.filter(Boolean).map(l=>l.trim()))].filter(l=>!m.has(b(l))).map(l=>({emailAddress:{address:l}})),C={body:{contentType:"HTML",content:t}};return $.length&&(C.ccRecipients=[...f,...$]),(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${r.id}`,{method:"PATCH",headers:i,body:JSON.stringify(C)})).ok?(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${r.id}/send`,{method:"POST",headers:i})).ok:!1}catch{return!1}}async function nt(t,e){const n=(await Be()).find(i=>i.EventKey===t&&i.IsEnabled);return n&&ne(n.Body||"",e)||null}async function W(t){var i;const e=c=>c.trim().toLowerCase(),a=e(((i=o.account)==null?void 0:i.username)??""),n=new Set,s=t.recipients.filter(Boolean).filter(c=>{const r=e(c);return!r||r===a||n.has(r)?!1:(n.add(r),!0)});if(s.length!==0)try{const c=await L(),r=`${D}/_api/web/lists/getbytitle('HD_Notifications')/items`;await Promise.all(s.map(b=>fetch(r,{method:"POST",headers:{Authorization:`Bearer ${c}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:t.title.slice(0,255),RecipientEmail:b,EventType:t.eventType,Message:t.message,LinkPath:t.linkPath,IsRead:!1})})))}catch{}}async function se(t,e){const a=document.querySelectorAll(".email-att-cb:checked");if(a.length===0)return;const n=await L();for(const s of Array.from(a)){const i=s.dataset.attId,c=s.dataset.attName,r=await new Promise((u,l)=>{Office.context.mailbox.item.getAttachmentContentAsync(i,{},d=>{d.status===Office.AsyncResultStatus.Succeeded?u(d):l(new Error(d.error.message))})}),{content:b,format:f}=r.value;let m;if(f===Office.MailboxEnums.AttachmentContentFormat.Base64){const u=atob(b);m=new Uint8Array(u.length);for(let l=0;l<u.length;l++)m[l]=u.charCodeAt(l)}else if(f===Office.MailboxEnums.AttachmentContentFormat.Eml||f===Office.MailboxEnums.AttachmentContentFormat.ICalendar)m=new TextEncoder().encode(b);else continue;const $=encodeURIComponent(c),C=`${D}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${$}')`;if(!(await fetch(C,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:m.buffer})).ok)throw new Error(`Upload ${c} failed`)}}async function st(t){const e=`https://graph.microsoft.com/v1.0/me/messages/${t}/$value`;let a=await V(),n=await fetch(e,{headers:{Authorization:`Bearer ${a}`}});if((n.status===401||n.status===403)&&(a=await V(!0),n=await fetch(e,{headers:{Authorization:`Bearer ${a}`}})),!n.ok)throw new Error(`Graph ${n.status}`);return n.arrayBuffer()}async function it(t){const e=await new Promise((n,s)=>{Office.context.mailbox.getCallbackTokenAsync({isRest:!0},i=>{i.status===Office.AsyncResultStatus.Succeeded?n(i.value):s(new Error("callback token failed"))})}),a=await fetch(`${Office.context.mailbox.restUrl}/v2.0/me/messages/${t}/$value`,{headers:{Authorization:`Bearer ${e}`}});if(!a.ok)throw new Error(`REST ${a.status}`);return a.arrayBuffer()}async function ie(t,e){const a=document.getElementById("f-attach-eml");if(!(a!=null&&a.checked))return;const n=Office.context.mailbox.item;if(!n)return;const s=Office.context.mailbox.convertToRestId(n.itemId,Office.MailboxEnums.RestVersion.v2_0);let i,c="",r="";try{i=await st(s)}catch(C){c=C instanceof Error?C.message:String(C);try{i=await it(s)}catch(_){r=_ instanceof Error?_.message:String(_),console.error("[eml] graph:",c,"| callback:",r),A(`ดึง .eml ไม่ได้ (Graph: ${c} / REST: ${r}) — ไฟล์อื่นบันทึกแล้ว`,"error");return}}const b=(n.subject||"email").replace(/[\\/:*?"<>|#%&{}~]/g,"_").slice(0,100).trim()||"email",f=await L(),m=`${D}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${encodeURIComponent(b+".eml")}')`;(await fetch(m,{method:"POST",headers:{Authorization:`Bearer ${f}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:i})).ok||A("บันทึกไฟล์ .eml ไม่สำเร็จ","error")}async function oe(t,e,a){const n=await L();for(const s of a){const i=await s.arrayBuffer(),c=encodeURIComponent(s.name),r=`${D}/_api/web/lists/getbytitle('${encodeURIComponent(t)}')/items(${e})/AttachmentFiles/add(FileName='${c}')`;if(!(await fetch(r,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:i})).ok)throw new Error(`Upload ${s.name} failed`)}}function A(t,e="success"){const a=document.getElementById("toast-container");if(!a)return;const n=e==="success"?"bg-green-500":e==="error"?"bg-red-500":"bg-slate-700",s=e==="success"?"✅":e==="error"?"❌":"ℹ️",i=document.createElement("div");i.className=`toast pointer-events-auto ${n} text-white text-sm font-medium px-4 py-3 rounded-lg shadow-lg max-w-xs mx-2`,i.textContent=`${s} ${t}`,a.appendChild(i),setTimeout(()=>i.remove(),4e3)}function ot(t){const e=t.split(`
`).map(m=>m.trim()).filter(Boolean),a=/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/,n=/(\+?[\d\s()\-]{7,}(?:ext\.?\s*\d+)?)/i,s=/\b(co\.?,?\s*ltd\.?|co\.?,?\s*limited|corporation|corp\.?|บริษัท|จำกัด|holding|group|inc\.?|llc)\b/i;let i="",c="",r="";const b=[];for(const m of e)if(!/^[-_=*]{2,}$/.test(m)&&!/^(best regards|regards|sincerely|ขอแสดงความนับถือ|ด้วยความนับถือ|from|sent|thanks|thank you)/i.test(m)){if(!i){const $=m.match(a);if($){i=$[0];continue}}if(!c){const $=m.match(n);if($&&$[0].replace(/\D/g,"").length>=7){c=$[0].trim();continue}}if(!r&&s.test(m)){r=m;continue}m.length>=2&&m.length<=50&&!/\d{4,}/.test(m)&&b.push(m)}const f=b.find(m=>!a.test(m)&&!s.test(m))??"";return!i&&!f?null:{name:f,company:r,email:i,phone:c}}async function ct(){const t=o.signatureContact;if(!t)return;const e=(o.emailSenderEmail||"").toLowerCase();if(e&&o.contactEmails.includes(e)){A("ลูกค้านี้มีในระบบแล้ว","success"),o.signatureContact=null,H();return}const a=document.getElementById("btn-import-customer");a&&(a.disabled=!0,a.textContent="กำลังบันทึก…");try{await F("HD_Contracts",{Title:o.emailSenderName||t.name,CustomerEmail:o.emailSenderEmail,Phone:t.phone||void 0,Company:t.company||void 0,Status:"Active"}),e&&o.contactEmails.push(e),A("เพิ่มลูกค้าสำเร็จ!"),o.signatureContact=null,H()}catch(n){const s=n instanceof Error?n.message:String(n);A(`เกิดข้อผิดพลาด: ${s}`,"error"),a&&(a.disabled=!1,a.textContent="เพิ่มเป็นลูกค้า")}}function je(){return new Date().toISOString().split("T")[0]}function rt(){const t=new Date;return`HD-${`${t.getFullYear()}${String(t.getMonth()+1).padStart(2,"0")}${String(t.getDate()).padStart(2,"0")}`}-${Math.floor(Math.random()*900+100)}`}function ge(){var t;return o.droppedFiles.length>0||document.querySelectorAll(".email-att-cb:checked").length>0||(((t=document.getElementById("f-attach-eml"))==null?void 0:t.checked)??!1)}async function Z(t,e){o.droppedFiles.length>0&&await oe(t,e,o.droppedFiles),await se(t,e),await ie(t,e)}let ee=!1;async function lt(){var e,a,n,s,i,c,r,b,f,m,$,C,_;if(!o.account){A("กรุณาเข้าสู่ระบบก่อน","error");return}if(ee)return;ee=!0;const t=document.getElementById("submit-btn");t&&(t.disabled=!0,t.textContent="กำลังบันทึก…");try{if(o.tab==="phish")await Ue();else if(o.tab==="ticket"){const u=document.getElementById("f-title").value.trim(),l=document.getElementById("f-description").value.trim(),d=document.getElementById("f-priority").value,T=document.getElementById("f-customer-email").value.trim(),E=((e=document.getElementById("f-cc-enable"))==null?void 0:e.checked)??!0?(((a=document.getElementById("f-cc"))==null?void 0:a.value)||"").split(/[,;\s]+/).map(R=>R.trim()).filter(Boolean):[],y=document.getElementById("f-assigned-email").value,v=o.agents.find(R=>R.email===y),B=rt(),P=await F("HD_Tickets",{Title:u,TicketNumber:B,Description:l,Priority:d,CustomerEmail:T,CustomerName:o.emailSenderName||T,Status:"Open",AssignedEmail:y||void 0,AssignedToName:(v==null?void 0:v.name)??((n=o.account)==null?void 0:n.name)??""});if(ge()){const R=await F("HD_TicketComments",{Title:"📎 ไฟล์แนบจากอีเมล",TicketID:P,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await Z("HD_TicketComments",R)}o.droppedFiles=[];const p={ticket_number:B,ticket_title:u,priority:d,category:"-",description:(l||"-").replace(/\n/g,"<br>"),customer_name:o.emailSenderName||T,assigned_name:(v==null?void 0:v.name)??((s=o.account)==null?void 0:s.name)??"-",link:"https://itservices.co.th/helpdesk/"},S=[y,o.account.username,...E,Se].filter(Boolean);let I=!1;const z=await nt("ticket_created",p);if(z){const R=`<div style="border-left:4px solid #2563eb;background:#eff6ff;padding:10px 14px;margin:0 0 14px;font-family:Segoe UI,sans-serif">
             <div style="font-size:15px;font-weight:700;color:#1e40af">Ticket No. ${B}</div>
             <div style="font-size:12px;color:#475569;margin-top:2px">กรุณาตอบกลับในอีเมลฉบับนี้เพื่อให้ข้อมูลอยู่ใน Ticket เดียวกัน</div>
           </div>`;I=await at(R+z,S)}I||await tt("ticket_created",p,[T],S),A(I?"สร้าง Ticket และตอบกลับในเธรดเดิมแล้ว!":"สร้าง Ticket สำเร็จ!")}else if(o.tab==="task"){const u=document.getElementById("f-title").value.trim(),l=parseInt(((i=document.getElementById("f-project"))==null?void 0:i.value)||"0"),d=document.getElementById("f-due-date").value,T=document.getElementById("f-note").value.trim(),j=document.getElementById("f-assigned-email").value,E=o.agents.find(B=>B.email===j);if(!l){A("กรุณาเลือก Project","error");return}const y=await F("PM_Tasks",{Title:u,DueDate:d||null,TaskNote:T,AssignedTo:(E==null?void 0:E.name)??o.account.name??o.account.username,AssignedEmail:j,IsCompleted:!1,IsAcknowledged:!1,ProjectID:l});if(o.droppedFiles.length>0&&await oe("PM_Tasks",y,o.droppedFiles),await se("PM_Tasks",y),await ie("PM_Tasks",y),o.droppedFiles=[],await W({recipients:[j],title:`📋 ได้รับมอบหมาย Task: ${u}`,message:T||(d?`กำหนดส่ง ${d}`:"มี Task ใหม่"),linkPath:l?`/projects/${l}`:"/my-work",eventType:"task_assigned"}),((c=document.getElementById("f-teams"))==null?void 0:c.checked)&&d){const B=Array.from(document.querySelectorAll(".att-internal:checked")).map(I=>I.value),P=(((r=document.getElementById("f-ext-att"))==null?void 0:r.value)||"").split(/[,;\s]+/).map(I=>I.trim()).filter(Boolean),p=`${d}T09:00:00`,S=`${d}T10:00:00`;try{await Qe({subject:u,start:p,end:S,body:T,attendees:[...B,...P],isOnlineMeeting:!0}),A("สร้าง Task + นัดประชุม Teams สำเร็จ!")}catch(I){A("สร้าง Task แล้ว แต่สร้างนัดประชุมไม่สำเร็จ: "+(I instanceof Error?I.message:""),"error")}}else A("สร้าง Task สำเร็จ!")}else if(o.tab==="incident"){const u=document.getElementById("f-title").value.trim(),l=parseInt(((b=document.getElementById("f-project"))==null?void 0:b.value)||"0"),d=document.getElementById("f-description").value.trim(),T=document.getElementById("f-severity").value,j=document.getElementById("f-assigned-email").value,E=o.agents.find(p=>p.email===j),y=document.getElementById("f-status").value,v=document.getElementById("f-incident-date").value,B=document.getElementById("f-resolution").value.trim();if(!l){A("กรุณาเลือก Project","error");return}const P=await F("PM_Incidents",{Title:u,Description:d||void 0,Severity:T,Status:y,AssignedTo:(E==null?void 0:E.name)??o.account.name??o.account.username,AssignedEmail:j,ProjectID:l,IncidentDate:v||je(),Resolution:B||void 0});o.droppedFiles.length>0&&await oe("PM_Incidents",P,o.droppedFiles),await se("PM_Incidents",P),await ie("PM_Incidents",P),o.droppedFiles=[],await W({recipients:[j],title:`🚨 ได้รับมอบหมาย Incident: ${u}`,message:`ความรุนแรง ${T}${d?" — "+d.slice(0,120):""}`,linkPath:l?`/projects/${l}`:"/my-work",eventType:"incident_created"}),A("สร้าง Incident สำเร็จ!")}else if(o.tab==="comment"){const u=parseInt(((f=document.getElementById("f-ticket"))==null?void 0:f.value)||"0"),l=document.getElementById("f-comment").value.trim(),d=document.getElementById("f-comment-type").value;if(!u){A("กรุณาเลือก Ticket","error");return}if(!l){A("กรุณาพิมพ์ Comment","error");return}const T=await F("HD_TicketComments",{Title:l.slice(0,100),TicketID:u,CommentText:l,CommentType:d,CommentDate:new Date().toISOString()});await Z("HD_TicketComments",T),o.droppedFiles=[];try{const j=await L(),E=`${D}/_api/web/lists/getbytitle('HD_Tickets')/items(${u})?$select=TicketNumber,Title,AssignedEmail,Author/EMail&$expand=Author`,y=await fetch(E,{headers:{Authorization:`Bearer ${j}`,Accept:"application/json;odata=nometadata"}});if(y.ok){const v=await y.json(),B=o.account.username.toLowerCase(),P=[...new Set([v.AssignedEmail,(m=v.Author)==null?void 0:m.EMail].filter(Boolean))].filter(p=>p.toLowerCase()!==B);P.length&&await W({recipients:P,title:`💬 ${(($=o.account)==null?void 0:$.name)??"มีคน"} คอมเมนต์ใน ${v.TicketNumber||"#"+u}`,message:l.slice(0,200),linkPath:`/tickets/${u}`,eventType:"comment_added"})}}catch{}A("เพิ่ม Comment สำเร็จ!")}else if(o.tab==="project"){const u=document.getElementById("f-title").value.trim(),l=document.getElementById("f-company").value.trim(),d=document.getElementById("f-group").value,T=document.getElementById("f-status").value,j=document.getElementById("f-start").value,E=document.getElementById("f-end").value,y=document.getElementById("f-description").value.trim();if(!u){A("กรุณาใส่ชื่อโครงการ","error");return}const v=await F("PM_Projects",{Title:u,Company:l||void 0,ProjectGroup:d,Progress:0,StartDate:j||void 0,EndDate:E||null,Status:T,CreatedByEmail:o.account.username,Comment:y||void 0});if(ge()){const B=await F("PM_Comments",{Title:"📎 ไฟล์แนบจากอีเมล",ProjectID:v,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await Z("PM_Comments",B)}o.droppedFiles=[],A("สร้างโครงการสำเร็จ!")}else if(o.tab==="projcomment"){const u=parseInt(((C=document.getElementById("f-project"))==null?void 0:C.value)||"0"),l=document.getElementById("f-comment").value.trim(),d=document.getElementById("f-comment-type").value;if(!u){A("กรุณาเลือกโครงการ","error");return}if(!l){A("กรุณาพิมพ์ Comment","error");return}const T=await F("PM_Comments",{Title:l.slice(0,100),ProjectID:u,CommentText:l,CommentType:d,CommentDate:new Date().toISOString(),ParentID:0});await Z("PM_Comments",T),o.droppedFiles=[];try{const j=await L(),E=`${D}/_api/web/lists/getbytitle('PM_Projects')/items(${u})?$select=Title,CreatedByEmail`,y=await fetch(E,{headers:{Authorization:`Bearer ${j}`,Accept:"application/json;odata=nometadata"}});if(y.ok){const v=await y.json(),B=o.account.username.toLowerCase();v.CreatedByEmail&&v.CreatedByEmail.toLowerCase()!==B&&await W({recipients:[v.CreatedByEmail],title:`💬 ${((_=o.account)==null?void 0:_.name)??"มีคน"} คอมเมนต์ในโครงการ ${v.Title??""}`,message:l.slice(0,200),linkPath:`/projects/${u}?tab=comments`,eventType:"comment_added"})}}catch{}A("เพิ่ม Comment สำเร็จ!")}}catch(u){const l=u instanceof Error?u.message:String(u);A(`เกิดข้อผิดพลาด: ${l}`,"error")}finally{ee=!1,t&&(t.disabled=!1,t.textContent="บันทึก")}}const dt={ticket:{label:"Ticket",icon:"🎫"},comment:{label:"Comment",icon:"💬"},project:{label:"Project",icon:"📁"},task:{label:"Task",icon:"✅"},incident:{label:"Incident",icon:"🚨"},projcomment:{label:"Comment",icon:"💬"},phish:{label:"PhishGuard",icon:"🛡️"}},mt=[{title:"🎫 Helpdesk",tabs:["ticket","comment"]},{title:"📁 Project",tabs:["project","task","incident","projcomment"]},{title:"🛡️ Security",tabs:["phish"]}],Pe=["f-title","f-description","f-priority","f-customer-email","f-cc","f-assigned-email","f-project","f-due-date","f-note","f-severity","f-status","f-incident-date","f-resolution","f-ticket","f-comment","f-comment-type","f-company","f-group","f-start","f-end","f-ext-att","f-attach-eml"];let q={};function ut(){for(const e of Pe){const a=document.getElementById(e);a&&(q[e]=a.value)}const t=document.getElementById("f-teams");t&&(q["f-teams"]=t.checked)}function pt(){for(const e of Pe){const a=document.getElementById(e);a&&q[e]!==void 0&&q[e]!==""&&(a.value=q[e])}const t=document.getElementById("f-teams");if(t&&q["f-teams"]!==void 0){t.checked=q["f-teams"];const e=document.getElementById("teams-fields");e&&(e.style.display=t.checked?"block":"none")}}function H(){var j,E,y,v,B,P;const t=document.getElementById("app");if(!t)return;ut();const{account:e,tab:a,emailSubject:n,emailSenderName:s,emailSenderEmail:i,emailBodyPreview:c}=o,r=e!==null,b=`
    <div class="bg-gradient-to-r from-blue-700 to-blue-600 text-white px-4 py-2.5 flex items-center gap-2.5 shadow flex-shrink-0">
      <div class="w-7 h-7 bg-white/95 rounded-lg flex items-center justify-center flex-shrink-0">
        <span class="text-blue-700 font-extrabold text-xs">iT</span>
      </div>
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm leading-tight truncate">iT Services Helpdesk</div>
        ${r?`<div class="text-[10px] text-blue-100 truncate">${g((e==null?void 0:e.name)??(e==null?void 0:e.username)??"")}</div>`:""}
      </div>
      ${r?`<button id="btn-logout" title="ออกจากระบบ" class="p-1.5 rounded-lg hover:bg-white/15 transition flex-shrink-0">
             <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
           </button>`:`<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded-lg transition flex-shrink-0">
             เข้าสู่ระบบ
           </button>`}
    </div>
  `;if(!r){t.innerHTML=`
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
    `,(j=document.getElementById("btn-login"))==null||j.addEventListener("click",he),(E=document.getElementById("btn-login-main"))==null||E.addEventListener("click",he);return}const f=n?`<div class="mx-3 mt-3 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5 text-xs text-slate-700 space-y-1">
         <div class="font-semibold text-blue-800 truncate" title="${g(n)}">📧 ${g(n)}</div>
         ${s?`<div class="text-slate-500">จาก: <span class="font-medium text-slate-700">${g(s)}</span></div>`:""}
         ${i&&i!==s?`<div class="text-slate-400 truncate">${g(i)}</div>`:""}
       </div>`:`<div class="mx-3 mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-700">
         ⚠️ ไม่พบข้อมูล Email (โหมดทดสอบ)
       </div>`,m=o.signatureContact,$=!!i&&o.contactEmails.includes(i.toLowerCase()),C=m?`<div class="mx-3 mt-3 bg-orange-50 border border-orange-200 rounded-lg px-3 py-2.5 text-xs text-slate-700">
         <div class="flex items-center justify-between mb-2">
           <span class="font-semibold text-orange-700">👤 ข้อมูลผู้ส่ง (จาก Signature)</span>
         </div>
         <div class="space-y-0.5 mb-2.5">
           ${s?`<div><span class="text-slate-400">ชื่อ:</span> <span class="font-medium">${g(s)}</span></div>`:""}
           ${m.company?`<div><span class="text-slate-400">บริษัท:</span> ${g(m.company)}</div>`:""}
           ${i?`<div><span class="text-slate-400">Email:</span> ${g(i)}</div>`:""}
           ${m.phone?`<div><span class="text-slate-400">โทร:</span> ${g(m.phone)}</div>`:""}
         </div>
         ${$?'<div class="w-full bg-green-100 text-green-700 text-xs font-semibold py-1.5 rounded-md text-center">✓ ลูกค้านี้มีในระบบแล้ว</div>':`<button id="btn-import-customer"
                class="w-full bg-orange-500 hover:bg-orange-400 text-white text-xs font-semibold py-1.5 rounded-md transition">
                + เพิ่มเป็นลูกค้า
              </button>`}
       </div>`:"",_=`
    <div class="mx-3 mt-3 space-y-2">
      ${mt.map(p=>`
        <div>
          <div class="text-[10px] font-semibold text-slate-400 uppercase tracking-wide mb-1 px-0.5">${p.title}</div>
          <div class="grid grid-cols-4 gap-1">
            ${p.tabs.map(S=>{const I=dt[S];return`<button data-tab="${S}"
                class="tab-btn flex flex-col items-center gap-1 py-2 rounded-lg transition ${a===S?"bg-blue-700 text-white shadow":"text-slate-500 hover:bg-slate-100"}">
                <span class="text-base leading-none">${I.icon}</span>
                <span class="text-[9px] font-medium leading-none">${I.label}</span>
              </button>`}).join("")}
          </div>
        </div>
      `).join("")}
    </div>
  `;let u="";a==="phish"?u=Ke(!!e):a==="ticket"?u=`
      ${w("Title / หัวข้อ",`<input id="f-title" type="text"
        class="${x}"
        value="${g(n)}" />`)}
      ${w("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${x} resize-none">${g(c)}</textarea>`)}
      ${w("Priority",`<select id="f-priority" class="${x}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
      ${w("Customer Email",`<input id="f-customer-email" type="email"
        class="${x}"
        value="${g(i)}" />`)}
      ${w("CC — ให้ผู้ที่อยู่ในเมลนี้รับรู้",`
        <label class="flex items-center gap-2 text-xs text-slate-600 mb-1.5 cursor-pointer">
          <input id="f-cc-enable" type="checkbox" ${o.emailCc.length?"checked":""} /> แนบผู้รับในเมลนี้เป็น CC อัตโนมัติ
        </label>
        <input id="f-cc" type="text" class="${x}" value="${g(o.emailCc.join(", "))}" placeholder="someone@company.com, boss@company.com" />`)}
      ${w("Assign ให้ Agent",te(e.username))}
      ${G()}
    `:a==="task"?u=`
      ${w("ชื่อ Task *",`<input id="f-title" type="text" required
        class="${x}" value="${g(n)}" />`)}
      ${w("Project *",ae())}
      ${w("Assign ให้",te(e.username))}
      ${w("Due Date",`<input id="f-due-date" type="date" class="${x}" />`)}
      ${w("Task Note",`<textarea id="f-note" rows="4"
        class="${x} resize-y">${g(c)}</textarea>`)}
      <label class="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer mb-1">
        <input id="f-teams" type="checkbox" class="rounded" onchange="document.getElementById('teams-fields').style.display=this.checked?'block':'none'" />
        💻 เพิ่มการประชุมออนไลน์ (Teams) — ใช้เวลา 09:00–10:00 ของวัน Due Date
      </label>
      <div id="teams-fields" style="display:none" class="space-y-2 mb-2">
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">ผู้เข้าร่วม Internal</label>
          <div class="max-h-28 overflow-y-auto border border-slate-200 rounded-md p-1.5 space-y-0.5">
            ${o.agents.map(p=>`<label class="flex items-center gap-2 text-xs text-slate-700 px-1 py-0.5 hover:bg-slate-50 rounded cursor-pointer">
              <input type="checkbox" class="att-internal" value="${g(p.email)}" /> ${g(p.name)}
            </label>`).join("")}
          </div>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">Email ภายนอก (คั่นด้วย ,)</label>
          <input id="f-ext-att" type="text" class="${x}" placeholder="someone@company.com, ..." />
        </div>
      </div>
      ${G()}
    `:a==="incident"?u=`
      ${w("ชื่อ Incident *",`<input id="f-title" type="text" required
        class="${x}" value="${g(n)}" />`)}
      ${w("Project *",ae())}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">ความรุนแรง</label>
          <select id="f-severity" class="${x}">
            <option value="Low">Low</option>
            <option value="Medium" selected>Medium</option>
            <option value="High">High</option>
            <option value="Critical">Critical</option>
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${x}">
            <option value="Open" selected>Open</option>
            <option value="In Progress">In Progress</option>
            <option value="Resolved">Resolved</option>
          </select>
        </div>
      </div>
      ${w("Assign ให้ Agent",te(e.username))}
      ${w("วันที่เกิด Incident",`<input id="f-incident-date" type="date" class="${x}" value="${je()}" />`)}
      ${w("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${x} resize-y">${g(c)}</textarea>`)}
      ${w("วิธีแก้ไข (ถ้ามี)",`<textarea id="f-resolution" rows="2"
        class="${x} resize-y" placeholder="อธิบายวิธีแก้ไขปัญหา..."></textarea>`)}
      ${G()}
    `:a==="comment"?u=`
      ${w("เลือก Ticket *",`<select id="f-ticket" class="${x}">
        <option value="">-- เลือก Ticket ที่จะเพิ่ม Comment --</option>
        ${o.tickets.map(p=>`<option value="${p.id}">${g(p.TicketNumber||"#"+p.id)} · ${g(p.Title)}</option>`).join("")}
      </select>`)}
      ${w("ประเภท",`<select id="f-comment-type" class="${x}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${w("Comment *",`<textarea id="f-comment" rows="5"
        class="${x} resize-y" placeholder="พิมพ์ comment...">${g(c)}</textarea>`)}
      ${G()}
    `:a==="project"?u=`
      ${w("ชื่อโครงการ *",`<input id="f-title" type="text" required
        class="${x}" value="${g(n)}" />`)}
      ${w("บริษัท / ลูกค้า",`<input id="f-company" type="text" class="${x}" value="${g(((y=o.signatureContact)==null?void 0:y.company)??"")}" />`)}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">กลุ่มโครงการ</label>
          <select id="f-group" class="${x}">
            ${["Internal","External","R&D","Maintenance","อื่นๆ"].map(p=>`<option>${p}</option>`).join("")}
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${x}">
            ${["Planning","Active","On Hold","Completed","Cancelled"].map(p=>`<option>${p}</option>`).join("")}
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันที่เริ่ม</label>
          <input id="f-start" type="date" class="${x}" /></div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันสิ้นสุด</label>
          <input id="f-end" type="date" class="${x}" /></div>
      </div>
      ${w("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${x} resize-y">${g(c)}</textarea>`)}
      ${G()}
    `:a==="projcomment"&&(u=`
      ${w("เลือกโครงการ *",ae())}
      ${w("ประเภท",`<select id="f-comment-type" class="${x}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${w("Comment *",`<textarea id="f-comment" rows="5"
        class="${x} resize-y" placeholder="พิมพ์ comment...">${g(c)}</textarea>`)}
      ${G()}
    `);const l=a==="phish"?qe():a==="comment"||a==="projcomment"?"เพิ่ม Comment":a==="project"?"สร้างโครงการ":a==="incident"?"แจ้ง Incident":a==="task"?"สร้าง Task":"สร้าง Ticket";t.innerHTML=`
    <div class="flex flex-col h-screen bg-slate-50">
      ${b}
      <div class="flex-1 overflow-y-auto">
        ${f}
        ${C}
        ${_}
        <div class="mx-3 mt-3 space-y-3 pb-3">
          ${u}
        </div>
      </div>
      <div class="border-t border-slate-200 bg-white px-3 py-2.5 flex-shrink-0">
        <button id="submit-btn"
          class="w-full bg-blue-700 hover:bg-blue-600 active:bg-blue-800 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          ${l}
        </button>
      </div>
    </div>
  `,(v=document.getElementById("btn-logout"))==null||v.addEventListener("click",Xe),(B=document.getElementById("submit-btn"))==null||B.addEventListener("click",lt),(P=document.getElementById("btn-import-customer"))==null||P.addEventListener("click",ct),a==="phish"&&Je(),document.querySelectorAll(".tab-btn").forEach(p=>{p.addEventListener("click",()=>{const S=p.dataset.tab;S&&S!==o.tab&&(o.tab=S,H(),S==="phish"&&ke())})});const d=document.getElementById("drop-zone"),T=document.getElementById("f-files");d&&T&&(T.addEventListener("change",()=>{T.files&&ce(Array.from(T.files)),T.value=""}),d.addEventListener("dragover",p=>{p.preventDefault(),d.classList.add("border-blue-500","bg-blue-50")}),d.addEventListener("dragleave",()=>{d.classList.remove("border-blue-500","bg-blue-50")}),d.addEventListener("drop",p=>{var I;p.preventDefault(),d.classList.remove("border-blue-500","bg-blue-50");const S=Array.from(((I=p.dataTransfer)==null?void 0:I.files)??[]);S.length&&ce(S)})),document.querySelectorAll(".remove-dropped").forEach(p=>{p.addEventListener("click",()=>{const S=parseInt(p.dataset.remove??"-1");S>=0&&(o.droppedFiles.splice(S,1),H())})}),pt()}function ce(t){o.droppedFiles.push(...t),H()}document.addEventListener("paste",t=>{var n;if(!o.account)return;const e=Array.from(((n=t.clipboardData)==null?void 0:n.items)??[]),a=[];for(const s of e)if(s.kind==="file"){const i=s.getAsFile();if(i){const c=i.name&&i.name!=="image.png"?i.name:`screenshot-${new Date().toISOString().replace(/[:.]/g,"-").slice(0,19)}.png`;a.push(new File([i],c,{type:i.type}))}}a.length&&(t.preventDefault(),ce(a),A(`แนบไฟล์แล้ว: ${a.map(s=>s.name).join(", ")}`))});const x="w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white";function be(t){return t<1024?`${t} B`:t<1024*1024?`${(t/1024).toFixed(0)} KB`:`${(t/1024/1024).toFixed(1)} MB`}function G(){const t=o.emailAttachments,e=o.droppedFiles,a=t.length>0?`<div class="mb-2 space-y-1">
        <p class="text-xs text-slate-500">📎 ไฟล์แนบจาก Email:</p>
        ${t.map(s=>`
          <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
            <input type="checkbox" class="email-att-cb" data-att-id="${g(s.id)}" data-att-name="${g(s.name)}" data-att-item="${s.isItem?"1":"0"}" checked />
            <span class="flex-1 truncate">${s.isItem?"📧 ":""}${g(s.name)}</span>
            <span class="text-slate-400 flex-shrink-0">${be(s.size)}</span>
          </label>`).join("")}
      </div>`:"",n=e.length>0?`<div class="mt-2 space-y-1">
        ${e.map((s,i)=>`<div class="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 rounded px-2 py-1">
            <span class="text-base">${s.type.startsWith("image/")?"🖼️":"📄"}</span>
            <span class="flex-1 truncate">${g(s.name)}</span>
            <span class="text-slate-400">${be(s.size)}</span>
            <button type="button" data-remove="${i}"
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
  </div>`}function te(t){var e;return`<select id="f-assigned-email" class="${x}">
    <option value="${g(t)}">${g(((e=o.account)==null?void 0:e.name)??t)} (ฉัน)</option>
    ${o.agents.filter(a=>a.email!==t).map(a=>`<option value="${g(a.email)}">${g(a.name)}</option>`).join("")}
  </select>`}function ae(){return o.projects.length===0?'<div class="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-2">⚠️ ไม่พบ Project ที่ Active</div>':`<select id="f-project" class="${x}">
    <option value="">-- เลือก Project --</option>
    ${o.projects.map(t=>`<option value="${t.id}">${g(t.Title)}</option>`).join("")}
  </select>`}function w(t,e){return`
    <div class="space-y-1">
      <label class="block text-xs font-medium text-slate-600">${t}</label>
      ${e}
    </div>
  `}function g(t){return t.replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}async function ft(){Le({sharepointUrl:D,internalDomains:["itservices.co.th","rpaexpert.com","rpaexpert.onmicrosoft.com"],getToken:L,getGraphToken:()=>V(),account:()=>o.account?{name:o.account.name,username:o.account.username}:null,toast:(e,a)=>A(e,a??"success"),rerender:H}),await O.initialize(),await O.handleRedirectPromise();const t=O.getAllAccounts();if(t.length>0){o.account=t[0];try{await O.acquireTokenSilent({scopes:[Y],account:t[0]}),await Promise.all([Ee(),Ie(),Ae(),Ce()])}catch{o.account=null}}typeof Office<"u"?Office.onReady(e=>{var a;if(e.host===Office.HostType.Outlook){const n=Office.context.mailbox.item;if(n){o.emailSubject=n.subject??"";const s=n.from;s&&(o.emailSenderName=s.displayName??"",o.emailSenderEmail=s.emailAddress??"");const i=(((a=o.account)==null?void 0:a.username)??"").toLowerCase(),c=((s==null?void 0:s.emailAddress)??"").toLowerCase(),r=[...n.to??[],...n.cc??[]].map(f=>f.emailAddress).filter(Boolean);o.emailCc=[...new Set(r.map(f=>f.toLowerCase()))].filter(f=>f!==i&&f!==c);const b=n.attachments??[];o.emailAttachments=b.filter(f=>!f.isInline&&(f.attachmentType===Office.MailboxEnums.AttachmentType.File||f.attachmentType===Office.MailboxEnums.AttachmentType.Item)).map(f=>({id:f.id,name:f.attachmentType===Office.MailboxEnums.AttachmentType.Item?`${(f.name||"email").replace(/\.eml$/i,"")}.eml`:f.name,size:f.size,isItem:f.attachmentType===Office.MailboxEnums.AttachmentType.Item})),n.body.getAsync(Office.CoercionType.Html,{asyncContext:{}},f=>{if(f.status===Office.AsyncResultStatus.Succeeded){let m=function(p,S=!1){if(p.nodeType===3){const M=p.textContent??"";return S&&M.trim()===""?"":M}const I=p,z=(I.tagName??"").toLowerCase();if(_.includes(z))return"";if(z==="br")return" ";if(z==="tr"){const M=[];for(let U=0;U<I.childNodes.length;U++){const de=I.childNodes[U],me=(de.tagName??"").toLowerCase();(me==="td"||me==="th")&&M.push((de.textContent??"").replace(/\s+/g," ").trim())}return M.length?M.join("	")+`
`:""}if(l.includes(z)){let M="";for(let U=0;U<I.childNodes.length;U++)M+=m(I.childNodes[U],!0);return M}let R="";for(let M=0;M<I.childNodes.length;M++)R+=m(I.childNodes[M],!1);return u.includes(z)&&(R=`
`+R.trim()+`
`),R};const $=f.value,C=new DOMParser().parseFromString($,"text/html"),_=["style","script","head","img","meta","link","noscript"],u=["p","div","li","h1","h2","h3","h4","h5","h6","blockquote"],l=["table","thead","tbody","tfoot"],j=m(C.body??C.documentElement).replace(/[ \t]{2,}/g," ").replace(/\n[ \t]+/g,`
`).replace(/\n{3,}/g,`

`).trim().split(`
`),E=[];let y="";for(const p of j)p.trim()===""?y&&(E.push(y.trim()),y=""):p.includes("	")?(y&&(E.push(y.trim()),y=""),E.push(p)):y=y?y+" "+p.trim():p.trim();y&&E.push(y.trim());const v=E.join(`
`),B=/\n([-_]{3,}|From:\s|Best regards|Regards,|ขอแสดงความนับถือ|Sent:\s)/i,P=v.search(B);if(P>80){o.emailBodyPreview=v.slice(0,P).trim().slice(0,2e3);const p=v.slice(P).trim();o.signatureContact=ot(p)}else o.emailBodyPreview=v.trim().slice(0,2e3),o.signatureContact=null}H()});return}}ye(),H()}):(ye(),H())}function ye(){o.emailSubject="[DEV] Test Email Subject",o.emailSenderName="Test Sender",o.emailSenderEmail="test@example.com",o.emailBodyPreview="This is a placeholder email body for development mode."}ft().catch(t=>{console.error("Init error:",t);const e=document.getElementById("app");e&&(e.innerHTML=`<div class="p-4 text-red-600 text-sm">เกิดข้อผิดพลาด: ${String(t)}</div>`)});
