import{P as xe}from"./PublicClientApplication-DLKYUtZW.js";const $e="0bab07cf-65e6-487c-89af-c917fc1a5a13",we="d569b991-89fc-4a62-9df5-eb361abcef40",_="https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd",J="https://rpaexpert.sharepoint.com/.default",K=["https://graph.microsoft.com/Calendars.ReadWrite","https://graph.microsoft.com/Mail.Send","https://graph.microsoft.com/Mail.Read"],j=new xe({auth:{clientId:$e,authority:`https://login.microsoftonline.com/${we}`,redirectUri:window.location.origin.includes("localhost")?"http://localhost:3000/":"https://darmmunginsa.github.io/itservices-addin/",navigateToLoginRequestUrl:!1},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:!1}}),ke=window.location.origin.includes("localhost")?`${window.location.origin}/auth.html`:"https://darmmunginsa.github.io/itservices-addin/auth.html";function oe(){var t,n;const e=(n=(t=Office.context)==null?void 0:t.diagnostics)==null?void 0:n.platform;return e===Office.PlatformType.iOS||e===Office.PlatformType.Android}function ie(){return new Promise((e,t)=>{Office.context.ui.displayDialogAsync(ke,{height:60,width:30,promptBeforeOpen:!1},n=>{if(n.status!==Office.AsyncResultStatus.Succeeded){t(new Error("เปิดหน้าเข้าสู่ระบบไม่ได้"));return}const o=n.value;o.addEventHandler(Office.EventType.DialogMessageReceived,i=>{o.close();const s=i.message;if(!s){t(new Error("auth message error"));return}try{const m=JSON.parse(s);m.ok?e():t(new Error(m.error||"auth failed"))}catch{t(new Error("auth message error"))}}),o.addEventHandler(Office.EventType.DialogEventReceived,()=>t(new Error("ปิดหน้าเข้าสู่ระบบก่อนเสร็จ")))})})}const a={account:null,tab:"ticket",emailSubject:"",emailBodyPreview:"",emailSenderName:"",emailSenderEmail:"",loading:!1,projects:[],agents:[],emailAttachments:[],signatureContact:null,droppedFiles:[],tickets:[],contactEmails:[],emailCc:[]};async function D(){const e=j.getAllAccounts();if(e.length===0)throw new Error("Not signed in");const t={scopes:[J],account:e[0]};try{return(await j.acquireTokenSilent(t)).accessToken}catch{if(oe()){await ie();const n=j.getAllAccounts()[0];if(!n)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await j.acquireTokenSilent({scopes:[J],account:n})).accessToken}return(await j.acquireTokenPopup(t)).accessToken}}async function q(e=!1){const t=j.getAllAccounts();if(t.length===0)throw new Error("Not signed in");const n={scopes:K,account:t[0],forceRefresh:e};try{return(await j.acquireTokenSilent(n)).accessToken}catch{if(oe()){await ie();const i=j.getAllAccounts()[0];if(!i)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await j.acquireTokenSilent({scopes:K,account:i})).accessToken}return(await j.acquireTokenPopup({scopes:K,account:t[0]})).accessToken}}async function Te(e){const t=await q(),n={subject:e.subject,start:{dateTime:e.start,timeZone:"Asia/Bangkok"},end:{dateTime:e.end,timeZone:"Asia/Bangkok"},body:e.body?{contentType:"HTML",content:e.body.replace(/\n/g,"<br>")}:void 0,attendees:e.attendees.filter(Boolean).map(i=>({emailAddress:{address:i},type:"required"})),isOnlineMeeting:e.isOnlineMeeting,onlineMeetingProvider:e.isOnlineMeeting?"teamsForBusiness":void 0},o=await fetch("https://graph.microsoft.com/v1.0/me/events",{method:"POST",headers:{Authorization:`Bearer ${t}`,"Content-Type":"application/json"},body:JSON.stringify(n)});if(!o.ok)throw new Error(`Calendar error ${o.status}: ${await o.text()}`)}async function ue(){try{const e=await D(),t=`${_}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$orderby=Title asc&$top=500`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const o=await n.json();a.projects=o.value.map(i=>({id:i.Id,Title:i.Title}))}}catch{}}async function pe(){try{const e=await D(),t=`${_}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText&$orderby=Title asc`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const o=await n.json();a.agents=o.value.map(i=>({email:i.EmailText,name:i.Title}))}}catch{}}async function fe(){try{const e=await D(),t=`${_}/_api/web/lists/getbytitle('HD_Tickets')/items?$select=Id,Title,TicketNumber,Status&$filter=Status ne 'Closed'&$orderby=Modified desc&$top=200`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const o=await n.json();a.tickets=o.value.map(i=>({id:i.Id,Title:i.Title,TicketNumber:i.TicketNumber,Status:i.Status}))}}catch{}}async function ge(){try{const e=await D(),t=`${_}/_api/web/lists/getbytitle('HD_Contracts')/items?$select=CustomerEmail&$top=2000`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const o=await n.json();a.contactEmails=o.value.map(i=>(i.CustomerEmail||"").trim().toLowerCase()).filter(Boolean)}}catch{}}async function le(){const e=document.getElementById("btn-login-main"),t=document.getElementById("btn-login");e&&(e.disabled=!0,e.textContent="กำลังเข้าสู่ระบบ…"),t&&(t.disabled=!0);try{if(oe()){if(await ie(),a.account=j.getAllAccounts()[0]??null,!a.account)throw new Error("เข้าสู่ระบบไม่สำเร็จ")}else{const n=await j.loginPopup({scopes:[J]});a.account=n.account}await Promise.all([ue(),pe(),fe(),ge()]),L()}catch{e&&(e.disabled=!1,e.textContent="เข้าสู่ระบบ"),t&&(t.disabled=!1)}}async function Ee(){a.account&&await j.logoutPopup({account:a.account}),a.account=null,L()}async function R(e,t){const n=await D(),o=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items`,i=await fetch(o,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(t)});if(!i.ok){const m=await i.text();throw new Error(`SharePoint error ${i.status}: ${m}`)}return(await i.json()).Id}let U=null;const Z="support@itservices.co.th",he="engineer@itservices.co.th";async function be(){if(U)return U;try{const e=await D(),t=`${_}/_api/web/lists/getbytitle('HD_EmailTemplates')/items?$select=EventKey,Subject,Body,IsEnabled&$top=50`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});return n.ok?(U=(await n.json()).value,U):[]}catch{return[]}}async function Ie(){var e,t;try{const n=await D(),o=`${_}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title,Category&$filter=Category eq 'EmailConfig'&$top=1`,i=await fetch(o,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});return i.ok&&((t=(e=(await i.json()).value[0])==null?void 0:e.Title)==null?void 0:t.trim())||Z}catch{return Z}}function X(e,t){return e.replace(/\{\{(\w+)\}\}/g,(n,o)=>t[o]??`{{${o}}}`)}async function Ce(e,t,n,o=[]){try{const s=(await be()).find(l=>l.EventKey===e&&l.IsEnabled);if(!s)return;const m=X(s.Subject||"",t),g=X(s.Body||"",t);if(!m||!g)return;const v=l=>l.trim().toLowerCase(),u=[...new Map(n.filter(Boolean).map(l=>[v(l),l])).values()];if(u.length===0)return;const p=new Set(u.map(v)),E=e==="ticket_created"?[...o,he]:o,A=[...new Map(E.filter(Boolean).map(l=>[v(l),l])).values()].filter(l=>!p.has(v(l))),P=await Ie(),r=await q(),c={subject:m,body:{contentType:"HTML",content:g},toRecipients:u.map(l=>({emailAddress:{address:l}}))};A.length&&(c.ccRecipients=A.map(l=>({emailAddress:{address:l}}))),P&&(c.from={emailAddress:{address:P}}),await fetch("https://graph.microsoft.com/v1.0/me/sendMail",{method:"POST",headers:{Authorization:`Bearer ${r}`,"Content-Type":"application/json"},body:JSON.stringify({message:c,saveToSentItems:!0})})}catch{}}async function Ae(e,t=[]){try{const n=Office.context.mailbox.item;if(!(n!=null&&n.itemId))return!1;const o=Office.context.mailbox.convertToRestId(n.itemId,Office.MailboxEnums.RestVersion.v2_0),s={Authorization:`Bearer ${await q()}`,"Content-Type":"application/json"},m=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${o}/createReplyAll`,{method:"POST",headers:s});if(!m.ok)return!1;const g=await m.json(),v=c=>c.trim().toLowerCase(),u=g.ccRecipients??[],p=new Set(u.map(c=>v(c.emailAddress.address))),E=[...new Set(t.filter(Boolean).map(c=>c.trim()))].filter(c=>!p.has(v(c))).map(c=>({emailAddress:{address:c}})),A={body:{contentType:"HTML",content:e}};return E.length&&(A.ccRecipients=[...u,...E]),(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${g.id}`,{method:"PATCH",headers:s,body:JSON.stringify(A)})).ok?(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${g.id}/send`,{method:"POST",headers:s})).ok:!1}catch{return!1}}async function Se(e,t){const o=(await be()).find(s=>s.EventKey===e&&s.IsEnabled);return o&&X(o.Body||"",t)||null}async function G(e){var s;const t=m=>m.trim().toLowerCase(),n=t(((s=a.account)==null?void 0:s.username)??""),o=new Set,i=e.recipients.filter(Boolean).filter(m=>{const g=t(m);return!g||g===n||o.has(g)?!1:(o.add(g),!0)});if(i.length!==0)try{const m=await D(),g=`${_}/_api/web/lists/getbytitle('HD_Notifications')/items`;await Promise.all(i.map(v=>fetch(g,{method:"POST",headers:{Authorization:`Bearer ${m}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:e.title.slice(0,255),RecipientEmail:v,EventType:e.eventType,Message:e.message,LinkPath:e.linkPath,IsRead:!1})})))}catch{}}async function ee(e,t){const n=document.querySelectorAll(".email-att-cb:checked");if(n.length===0)return;const o=await D();for(const i of Array.from(n)){const s=i.dataset.attId,m=i.dataset.attName,g=await new Promise((r,c)=>{Office.context.mailbox.item.getAttachmentContentAsync(s,{},l=>{l.status===Office.AsyncResultStatus.Succeeded?r(l):c(new Error(l.error.message))})}),{content:v,format:u}=g.value;let p;if(u===Office.MailboxEnums.AttachmentContentFormat.Base64){const r=atob(v);p=new Uint8Array(r.length);for(let c=0;c<r.length;c++)p[c]=r.charCodeAt(c)}else if(u===Office.MailboxEnums.AttachmentContentFormat.Eml||u===Office.MailboxEnums.AttachmentContentFormat.ICalendar)p=new TextEncoder().encode(v);else continue;const E=encodeURIComponent(m),A=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${E}')`;if(!(await fetch(A,{method:"POST",headers:{Authorization:`Bearer ${o}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:p.buffer})).ok)throw new Error(`Upload ${m} failed`)}}async function Be(e){const t=`https://graph.microsoft.com/v1.0/me/messages/${e}/$value`;let n=await q(),o=await fetch(t,{headers:{Authorization:`Bearer ${n}`}});if((o.status===401||o.status===403)&&(n=await q(!0),o=await fetch(t,{headers:{Authorization:`Bearer ${n}`}})),!o.ok)throw new Error(`Graph ${o.status}`);return o.arrayBuffer()}async function je(e){const t=await new Promise((o,i)=>{Office.context.mailbox.getCallbackTokenAsync({isRest:!0},s=>{s.status===Office.AsyncResultStatus.Succeeded?o(s.value):i(new Error("callback token failed"))})}),n=await fetch(`${Office.context.mailbox.restUrl}/v2.0/me/messages/${e}/$value`,{headers:{Authorization:`Bearer ${t}`}});if(!n.ok)throw new Error(`REST ${n.status}`);return n.arrayBuffer()}async function te(e,t){const n=document.getElementById("f-attach-eml");if(!(n!=null&&n.checked))return;const o=Office.context.mailbox.item;if(!o)return;const i=Office.context.mailbox.convertToRestId(o.itemId,Office.MailboxEnums.RestVersion.v2_0);let s,m="",g="";try{s=await Be(i)}catch(A){m=A instanceof Error?A.message:String(A);try{s=await je(i)}catch(P){g=P instanceof Error?P.message:String(P),console.error("[eml] graph:",m,"| callback:",g),T(`ดึง .eml ไม่ได้ (Graph: ${m} / REST: ${g}) — ไฟล์อื่นบันทึกแล้ว`,"error");return}}const v=(o.subject||"email").replace(/[\\/:*?"<>|#%&{}~]/g,"_").slice(0,100).trim()||"email",u=await D(),p=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${encodeURIComponent(v+".eml")}')`;(await fetch(p,{method:"POST",headers:{Authorization:`Bearer ${u}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:s})).ok||T("บันทึกไฟล์ .eml ไม่สำเร็จ","error")}async function ne(e,t,n){const o=await D();for(const i of n){const s=await i.arrayBuffer(),m=encodeURIComponent(i.name),g=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${m}')`;if(!(await fetch(g,{method:"POST",headers:{Authorization:`Bearer ${o}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:s})).ok)throw new Error(`Upload ${i.name} failed`)}}function T(e,t="success"){const n=document.getElementById("toast-container");if(!n)return;const o=t==="success"?"bg-green-500":"bg-red-500",i=t==="success"?"✅":"❌",s=document.createElement("div");s.className=`toast pointer-events-auto ${o} text-white text-sm font-medium px-4 py-3 rounded-lg shadow-lg max-w-xs mx-2`,s.textContent=`${i} ${e}`,n.appendChild(s),setTimeout(()=>s.remove(),4e3)}function Pe(e){const t=e.split(`
`).map(p=>p.trim()).filter(Boolean),n=/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/,o=/(\+?[\d\s()\-]{7,}(?:ext\.?\s*\d+)?)/i,i=/\b(co\.?,?\s*ltd\.?|co\.?,?\s*limited|corporation|corp\.?|บริษัท|จำกัด|holding|group|inc\.?|llc)\b/i;let s="",m="",g="";const v=[];for(const p of t)if(!/^[-_=*]{2,}$/.test(p)&&!/^(best regards|regards|sincerely|ขอแสดงความนับถือ|ด้วยความนับถือ|from|sent|thanks|thank you)/i.test(p)){if(!s){const E=p.match(n);if(E){s=E[0];continue}}if(!m){const E=p.match(o);if(E&&E[0].replace(/\D/g,"").length>=7){m=E[0].trim();continue}}if(!g&&i.test(p)){g=p;continue}p.length>=2&&p.length<=50&&!/\d{4,}/.test(p)&&v.push(p)}const u=v.find(p=>!n.test(p)&&!i.test(p))??"";return!s&&!u?null:{name:u,company:g,email:s,phone:m}}async function Oe(){const e=a.signatureContact;if(!e)return;const t=(a.emailSenderEmail||"").toLowerCase();if(t&&a.contactEmails.includes(t)){T("ลูกค้านี้มีในระบบแล้ว","success"),a.signatureContact=null,L();return}const n=document.getElementById("btn-import-customer");n&&(n.disabled=!0,n.textContent="กำลังบันทึก…");try{await R("HD_Contracts",{Title:a.emailSenderName||e.name,CustomerEmail:a.emailSenderEmail,Phone:e.phone||void 0,Company:e.company||void 0,Status:"Active"}),t&&a.contactEmails.push(t),T("เพิ่มลูกค้าสำเร็จ!"),a.signatureContact=null,L()}catch(o){const i=o instanceof Error?o.message:String(o);T(`เกิดข้อผิดพลาด: ${i}`,"error"),n&&(n.disabled=!1,n.textContent="เพิ่มเป็นลูกค้า")}}function ye(){return new Date().toISOString().split("T")[0]}function _e(){const e=new Date;return`HD-${`${e.getFullYear()}${String(e.getMonth()+1).padStart(2,"0")}${String(e.getDate()).padStart(2,"0")}`}-${Math.floor(Math.random()*900+100)}`}function re(){var e;return a.droppedFiles.length>0||document.querySelectorAll(".email-att-cb:checked").length>0||(((e=document.getElementById("f-attach-eml"))==null?void 0:e.checked)??!1)}async function V(e,t){a.droppedFiles.length>0&&await ne(e,t,a.droppedFiles),await ee(e,t),await te(e,t)}let W=!1;async function De(){var t,n,o,i,s,m,g,v,u,p,E,A,P;if(!a.account){T("กรุณาเข้าสู่ระบบก่อน","error");return}if(W)return;W=!0;const e=document.getElementById("submit-btn");e&&(e.disabled=!0,e.textContent="กำลังบันทึก…");try{if(a.tab==="ticket"){const r=document.getElementById("f-title").value.trim(),c=document.getElementById("f-description").value.trim(),l=document.getElementById("f-priority").value,$=document.getElementById("f-customer-email").value.trim(),w=((t=document.getElementById("f-cc-enable"))==null?void 0:t.checked)??!0?(((n=document.getElementById("f-cc"))==null?void 0:n.value)||"").split(/[,;\s]+/).map(M=>M.trim()).filter(Boolean):[],h=document.getElementById("f-assigned-email").value,y=a.agents.find(M=>M.email===h),I=_e(),B=await R("HD_Tickets",{Title:r,TicketNumber:I,Description:c,Priority:l,CustomerEmail:$,CustomerName:a.emailSenderName||$,Status:"Open",AssignedEmail:h||void 0,AssignedToName:(y==null?void 0:y.name)??((o=a.account)==null?void 0:o.name)??""});if(re()){const M=await R("HD_TicketComments",{Title:"📎 ไฟล์แนบจากอีเมล",TicketID:B,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await V("HD_TicketComments",M)}a.droppedFiles=[];const d={ticket_number:I,ticket_title:r,priority:l,category:"-",description:(c||"-").replace(/\n/g,"<br>"),customer_name:a.emailSenderName||$,assigned_name:(y==null?void 0:y.name)??((i=a.account)==null?void 0:i.name)??"-",link:"https://itservices.co.th/helpdesk/"},C=[h,a.account.username,...w,he].filter(Boolean);let k=!1;const N=await Se("ticket_created",d);if(N){const M=`<div style="border-left:4px solid #2563eb;background:#eff6ff;padding:10px 14px;margin:0 0 14px;font-family:Segoe UI,sans-serif">
             <div style="font-size:15px;font-weight:700;color:#1e40af">Ticket No. ${I}</div>
             <div style="font-size:12px;color:#475569;margin-top:2px">กรุณาตอบกลับในอีเมลฉบับนี้เพื่อให้ข้อมูลอยู่ใน Ticket เดียวกัน</div>
           </div>`;k=await Ae(M+N,C)}k||await Ce("ticket_created",d,[$],C),T(k?"สร้าง Ticket และตอบกลับในเธรดเดิมแล้ว!":"สร้าง Ticket สำเร็จ!")}else if(a.tab==="task"){const r=document.getElementById("f-title").value.trim(),c=parseInt(((s=document.getElementById("f-project"))==null?void 0:s.value)||"0"),l=document.getElementById("f-due-date").value,$=document.getElementById("f-note").value.trim(),S=document.getElementById("f-assigned-email").value,w=a.agents.find(I=>I.email===S);if(!c){T("กรุณาเลือก Project","error");return}const h=await R("PM_Tasks",{Title:r,DueDate:l||null,TaskNote:$,AssignedTo:(w==null?void 0:w.name)??a.account.name??a.account.username,AssignedEmail:S,IsCompleted:!1,IsAcknowledged:!1,ProjectID:c});if(a.droppedFiles.length>0&&await ne("PM_Tasks",h,a.droppedFiles),await ee("PM_Tasks",h),await te("PM_Tasks",h),a.droppedFiles=[],await G({recipients:[S],title:`📋 ได้รับมอบหมาย Task: ${r}`,message:$||(l?`กำหนดส่ง ${l}`:"มี Task ใหม่"),linkPath:c?`/projects/${c}`:"/my-work",eventType:"task_assigned"}),((m=document.getElementById("f-teams"))==null?void 0:m.checked)&&l){const I=Array.from(document.querySelectorAll(".att-internal:checked")).map(k=>k.value),B=(((g=document.getElementById("f-ext-att"))==null?void 0:g.value)||"").split(/[,;\s]+/).map(k=>k.trim()).filter(Boolean),d=`${l}T09:00:00`,C=`${l}T10:00:00`;try{await Te({subject:r,start:d,end:C,body:$,attendees:[...I,...B],isOnlineMeeting:!0}),T("สร้าง Task + นัดประชุม Teams สำเร็จ!")}catch(k){T("สร้าง Task แล้ว แต่สร้างนัดประชุมไม่สำเร็จ: "+(k instanceof Error?k.message:""),"error")}}else T("สร้าง Task สำเร็จ!")}else if(a.tab==="incident"){const r=document.getElementById("f-title").value.trim(),c=parseInt(((v=document.getElementById("f-project"))==null?void 0:v.value)||"0"),l=document.getElementById("f-description").value.trim(),$=document.getElementById("f-severity").value,S=document.getElementById("f-assigned-email").value,w=a.agents.find(d=>d.email===S),h=document.getElementById("f-status").value,y=document.getElementById("f-incident-date").value,I=document.getElementById("f-resolution").value.trim();if(!c){T("กรุณาเลือก Project","error");return}const B=await R("PM_Incidents",{Title:r,Description:l||void 0,Severity:$,Status:h,AssignedTo:(w==null?void 0:w.name)??a.account.name??a.account.username,AssignedEmail:S,ProjectID:c,IncidentDate:y||ye(),Resolution:I||void 0});a.droppedFiles.length>0&&await ne("PM_Incidents",B,a.droppedFiles),await ee("PM_Incidents",B),await te("PM_Incidents",B),a.droppedFiles=[],await G({recipients:[S],title:`🚨 ได้รับมอบหมาย Incident: ${r}`,message:`ความรุนแรง ${$}${l?" — "+l.slice(0,120):""}`,linkPath:c?`/projects/${c}`:"/my-work",eventType:"incident_created"}),T("สร้าง Incident สำเร็จ!")}else if(a.tab==="comment"){const r=parseInt(((u=document.getElementById("f-ticket"))==null?void 0:u.value)||"0"),c=document.getElementById("f-comment").value.trim(),l=document.getElementById("f-comment-type").value;if(!r){T("กรุณาเลือก Ticket","error");return}if(!c){T("กรุณาพิมพ์ Comment","error");return}const $=await R("HD_TicketComments",{Title:c.slice(0,100),TicketID:r,CommentText:c,CommentType:l,CommentDate:new Date().toISOString()});await V("HD_TicketComments",$),a.droppedFiles=[];try{const S=await D(),w=`${_}/_api/web/lists/getbytitle('HD_Tickets')/items(${r})?$select=TicketNumber,Title,AssignedEmail,Author/EMail&$expand=Author`,h=await fetch(w,{headers:{Authorization:`Bearer ${S}`,Accept:"application/json;odata=nometadata"}});if(h.ok){const y=await h.json(),I=a.account.username.toLowerCase(),B=[...new Set([y.AssignedEmail,(p=y.Author)==null?void 0:p.EMail].filter(Boolean))].filter(d=>d.toLowerCase()!==I);B.length&&await G({recipients:B,title:`💬 ${((E=a.account)==null?void 0:E.name)??"มีคน"} คอมเมนต์ใน ${y.TicketNumber||"#"+r}`,message:c.slice(0,200),linkPath:`/tickets/${r}`,eventType:"comment_added"})}}catch{}T("เพิ่ม Comment สำเร็จ!")}else if(a.tab==="project"){const r=document.getElementById("f-title").value.trim(),c=document.getElementById("f-company").value.trim(),l=document.getElementById("f-group").value,$=document.getElementById("f-status").value,S=document.getElementById("f-start").value,w=document.getElementById("f-end").value,h=document.getElementById("f-description").value.trim();if(!r){T("กรุณาใส่ชื่อโครงการ","error");return}const y=await R("PM_Projects",{Title:r,Company:c||void 0,ProjectGroup:l,Progress:0,StartDate:S||void 0,EndDate:w||null,Status:$,CreatedByEmail:a.account.username,Comment:h||void 0});if(re()){const I=await R("PM_Comments",{Title:"📎 ไฟล์แนบจากอีเมล",ProjectID:y,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await V("PM_Comments",I)}a.droppedFiles=[],T("สร้างโครงการสำเร็จ!")}else if(a.tab==="projcomment"){const r=parseInt(((A=document.getElementById("f-project"))==null?void 0:A.value)||"0"),c=document.getElementById("f-comment").value.trim(),l=document.getElementById("f-comment-type").value;if(!r){T("กรุณาเลือกโครงการ","error");return}if(!c){T("กรุณาพิมพ์ Comment","error");return}const $=await R("PM_Comments",{Title:c.slice(0,100),ProjectID:r,CommentText:c,CommentType:l,CommentDate:new Date().toISOString(),ParentID:0});await V("PM_Comments",$),a.droppedFiles=[];try{const S=await D(),w=`${_}/_api/web/lists/getbytitle('PM_Projects')/items(${r})?$select=Title,CreatedByEmail`,h=await fetch(w,{headers:{Authorization:`Bearer ${S}`,Accept:"application/json;odata=nometadata"}});if(h.ok){const y=await h.json(),I=a.account.username.toLowerCase();y.CreatedByEmail&&y.CreatedByEmail.toLowerCase()!==I&&await G({recipients:[y.CreatedByEmail],title:`💬 ${((P=a.account)==null?void 0:P.name)??"มีคน"} คอมเมนต์ในโครงการ ${y.Title??""}`,message:c.slice(0,200),linkPath:`/projects/${r}?tab=comments`,eventType:"comment_added"})}}catch{}T("เพิ่ม Comment สำเร็จ!")}}catch(r){const c=r instanceof Error?r.message:String(r);T(`เกิดข้อผิดพลาด: ${c}`,"error")}finally{W=!1,e&&(e.disabled=!1,e.textContent="บันทึก")}}const Me={ticket:{label:"Ticket",icon:"🎫"},comment:{label:"Comment",icon:"💬"},project:{label:"Project",icon:"📁"},task:{label:"Task",icon:"✅"},incident:{label:"Incident",icon:"🚨"},projcomment:{label:"Comment",icon:"💬"}},Le=[{title:"🎫 Helpdesk",tabs:["ticket","comment"]},{title:"📁 Project",tabs:["project","task","incident","projcomment"]}],ve=["f-title","f-description","f-priority","f-customer-email","f-cc","f-assigned-email","f-project","f-due-date","f-note","f-severity","f-status","f-incident-date","f-resolution","f-ticket","f-comment","f-comment-type","f-company","f-group","f-start","f-end","f-ext-att","f-attach-eml"];let F={};function Re(){for(const t of ve){const n=document.getElementById(t);n&&(F[t]=n.value)}const e=document.getElementById("f-teams");e&&(F["f-teams"]=e.checked)}function Ne(){for(const t of ve){const n=document.getElementById(t);n&&F[t]!==void 0&&F[t]!==""&&(n.value=F[t])}const e=document.getElementById("f-teams");if(e&&F["f-teams"]!==void 0){e.checked=F["f-teams"];const t=document.getElementById("teams-fields");t&&(t.style.display=e.checked?"block":"none")}}function L(){var S,w,h,y,I,B;const e=document.getElementById("app");if(!e)return;Re();const{account:t,tab:n,emailSubject:o,emailSenderName:i,emailSenderEmail:s,emailBodyPreview:m}=a,g=t!==null,v=`
    <div class="bg-gradient-to-r from-blue-700 to-blue-600 text-white px-4 py-2.5 flex items-center gap-2.5 shadow flex-shrink-0">
      <div class="w-7 h-7 bg-white/95 rounded-lg flex items-center justify-center flex-shrink-0">
        <span class="text-blue-700 font-extrabold text-xs">iT</span>
      </div>
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm leading-tight truncate">iT Services Helpdesk</div>
        ${g?`<div class="text-[10px] text-blue-100 truncate">${f((t==null?void 0:t.name)??(t==null?void 0:t.username)??"")}</div>`:""}
      </div>
      ${g?`<button id="btn-logout" title="ออกจากระบบ" class="p-1.5 rounded-lg hover:bg-white/15 transition flex-shrink-0">
             <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
           </button>`:`<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded-lg transition flex-shrink-0">
             เข้าสู่ระบบ
           </button>`}
    </div>
  `;if(!g){e.innerHTML=`
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
    `,(S=document.getElementById("btn-login"))==null||S.addEventListener("click",le),(w=document.getElementById("btn-login-main"))==null||w.addEventListener("click",le);return}const u=o?`<div class="mx-3 mt-3 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5 text-xs text-slate-700 space-y-1">
         <div class="font-semibold text-blue-800 truncate" title="${f(o)}">📧 ${f(o)}</div>
         ${i?`<div class="text-slate-500">จาก: <span class="font-medium text-slate-700">${f(i)}</span></div>`:""}
         ${s&&s!==i?`<div class="text-slate-400 truncate">${f(s)}</div>`:""}
       </div>`:`<div class="mx-3 mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-700">
         ⚠️ ไม่พบข้อมูล Email (โหมดทดสอบ)
       </div>`,p=a.signatureContact,E=!!s&&a.contactEmails.includes(s.toLowerCase()),A=p?`<div class="mx-3 mt-3 bg-orange-50 border border-orange-200 rounded-lg px-3 py-2.5 text-xs text-slate-700">
         <div class="flex items-center justify-between mb-2">
           <span class="font-semibold text-orange-700">👤 ข้อมูลผู้ส่ง (จาก Signature)</span>
         </div>
         <div class="space-y-0.5 mb-2.5">
           ${i?`<div><span class="text-slate-400">ชื่อ:</span> <span class="font-medium">${f(i)}</span></div>`:""}
           ${p.company?`<div><span class="text-slate-400">บริษัท:</span> ${f(p.company)}</div>`:""}
           ${s?`<div><span class="text-slate-400">Email:</span> ${f(s)}</div>`:""}
           ${p.phone?`<div><span class="text-slate-400">โทร:</span> ${f(p.phone)}</div>`:""}
         </div>
         ${E?'<div class="w-full bg-green-100 text-green-700 text-xs font-semibold py-1.5 rounded-md text-center">✓ ลูกค้านี้มีในระบบแล้ว</div>':`<button id="btn-import-customer"
                class="w-full bg-orange-500 hover:bg-orange-400 text-white text-xs font-semibold py-1.5 rounded-md transition">
                + เพิ่มเป็นลูกค้า
              </button>`}
       </div>`:"",P=`
    <div class="mx-3 mt-3 space-y-2">
      ${Le.map(d=>`
        <div>
          <div class="text-[10px] font-semibold text-slate-400 uppercase tracking-wide mb-1 px-0.5">${d.title}</div>
          <div class="grid grid-cols-4 gap-1">
            ${d.tabs.map(C=>{const k=Me[C];return`<button data-tab="${C}"
                class="tab-btn flex flex-col items-center gap-1 py-2 rounded-lg transition ${n===C?"bg-blue-700 text-white shadow":"text-slate-500 hover:bg-slate-100"}">
                <span class="text-base leading-none">${k.icon}</span>
                <span class="text-[9px] font-medium leading-none">${k.label}</span>
              </button>`}).join("")}
          </div>
        </div>
      `).join("")}
    </div>
  `;let r="";n==="ticket"?r=`
      ${x("Title / หัวข้อ",`<input id="f-title" type="text"
        class="${b}"
        value="${f(o)}" />`)}
      ${x("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${b} resize-none">${f(m)}</textarea>`)}
      ${x("Priority",`<select id="f-priority" class="${b}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
      ${x("Customer Email",`<input id="f-customer-email" type="email"
        class="${b}"
        value="${f(s)}" />`)}
      ${x("CC — ให้ผู้ที่อยู่ในเมลนี้รับรู้",`
        <label class="flex items-center gap-2 text-xs text-slate-600 mb-1.5 cursor-pointer">
          <input id="f-cc-enable" type="checkbox" ${a.emailCc.length?"checked":""} /> แนบผู้รับในเมลนี้เป็น CC อัตโนมัติ
        </label>
        <input id="f-cc" type="text" class="${b}" value="${f(a.emailCc.join(", "))}" placeholder="someone@company.com, boss@company.com" />`)}
      ${x("Assign ให้ Agent",Y(t.username))}
      ${z()}
    `:n==="task"?r=`
      ${x("ชื่อ Task *",`<input id="f-title" type="text" required
        class="${b}" value="${f(o)}" />`)}
      ${x("Project *",Q())}
      ${x("Assign ให้",Y(t.username))}
      ${x("Due Date",`<input id="f-due-date" type="date" class="${b}" />`)}
      ${x("Task Note",`<textarea id="f-note" rows="4"
        class="${b} resize-y">${f(m)}</textarea>`)}
      <label class="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer mb-1">
        <input id="f-teams" type="checkbox" class="rounded" onchange="document.getElementById('teams-fields').style.display=this.checked?'block':'none'" />
        💻 เพิ่มการประชุมออนไลน์ (Teams) — ใช้เวลา 09:00–10:00 ของวัน Due Date
      </label>
      <div id="teams-fields" style="display:none" class="space-y-2 mb-2">
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">ผู้เข้าร่วม Internal</label>
          <div class="max-h-28 overflow-y-auto border border-slate-200 rounded-md p-1.5 space-y-0.5">
            ${a.agents.map(d=>`<label class="flex items-center gap-2 text-xs text-slate-700 px-1 py-0.5 hover:bg-slate-50 rounded cursor-pointer">
              <input type="checkbox" class="att-internal" value="${f(d.email)}" /> ${f(d.name)}
            </label>`).join("")}
          </div>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">Email ภายนอก (คั่นด้วย ,)</label>
          <input id="f-ext-att" type="text" class="${b}" placeholder="someone@company.com, ..." />
        </div>
      </div>
      ${z()}
    `:n==="incident"?r=`
      ${x("ชื่อ Incident *",`<input id="f-title" type="text" required
        class="${b}" value="${f(o)}" />`)}
      ${x("Project *",Q())}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">ความรุนแรง</label>
          <select id="f-severity" class="${b}">
            <option value="Low">Low</option>
            <option value="Medium" selected>Medium</option>
            <option value="High">High</option>
            <option value="Critical">Critical</option>
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${b}">
            <option value="Open" selected>Open</option>
            <option value="In Progress">In Progress</option>
            <option value="Resolved">Resolved</option>
          </select>
        </div>
      </div>
      ${x("Assign ให้ Agent",Y(t.username))}
      ${x("วันที่เกิด Incident",`<input id="f-incident-date" type="date" class="${b}" value="${ye()}" />`)}
      ${x("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${b} resize-y">${f(m)}</textarea>`)}
      ${x("วิธีแก้ไข (ถ้ามี)",`<textarea id="f-resolution" rows="2"
        class="${b} resize-y" placeholder="อธิบายวิธีแก้ไขปัญหา..."></textarea>`)}
      ${z()}
    `:n==="comment"?r=`
      ${x("เลือก Ticket *",`<select id="f-ticket" class="${b}">
        <option value="">-- เลือก Ticket ที่จะเพิ่ม Comment --</option>
        ${a.tickets.map(d=>`<option value="${d.id}">${f(d.TicketNumber||"#"+d.id)} · ${f(d.Title)}</option>`).join("")}
      </select>`)}
      ${x("ประเภท",`<select id="f-comment-type" class="${b}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${x("Comment *",`<textarea id="f-comment" rows="5"
        class="${b} resize-y" placeholder="พิมพ์ comment...">${f(m)}</textarea>`)}
      ${z()}
    `:n==="project"?r=`
      ${x("ชื่อโครงการ *",`<input id="f-title" type="text" required
        class="${b}" value="${f(o)}" />`)}
      ${x("บริษัท / ลูกค้า",`<input id="f-company" type="text" class="${b}" value="${f(((h=a.signatureContact)==null?void 0:h.company)??"")}" />`)}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">กลุ่มโครงการ</label>
          <select id="f-group" class="${b}">
            ${["Internal","External","R&D","Maintenance","อื่นๆ"].map(d=>`<option>${d}</option>`).join("")}
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${b}">
            ${["Planning","Active","On Hold","Completed","Cancelled"].map(d=>`<option>${d}</option>`).join("")}
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันที่เริ่ม</label>
          <input id="f-start" type="date" class="${b}" /></div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันสิ้นสุด</label>
          <input id="f-end" type="date" class="${b}" /></div>
      </div>
      ${x("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${b} resize-y">${f(m)}</textarea>`)}
      ${z()}
    `:n==="projcomment"&&(r=`
      ${x("เลือกโครงการ *",Q())}
      ${x("ประเภท",`<select id="f-comment-type" class="${b}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${x("Comment *",`<textarea id="f-comment" rows="5"
        class="${b} resize-y" placeholder="พิมพ์ comment...">${f(m)}</textarea>`)}
      ${z()}
    `);const c=n==="comment"||n==="projcomment"?"เพิ่ม Comment":n==="project"?"สร้างโครงการ":n==="incident"?"แจ้ง Incident":n==="task"?"สร้าง Task":"สร้าง Ticket";e.innerHTML=`
    <div class="flex flex-col h-screen bg-slate-50">
      ${v}
      <div class="flex-1 overflow-y-auto">
        ${u}
        ${A}
        ${P}
        <div class="mx-3 mt-3 space-y-3 pb-3">
          ${r}
        </div>
      </div>
      <div class="border-t border-slate-200 bg-white px-3 py-2.5 flex-shrink-0">
        <button id="submit-btn"
          class="w-full bg-blue-700 hover:bg-blue-600 active:bg-blue-800 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          ${c}
        </button>
      </div>
    </div>
  `,(y=document.getElementById("btn-logout"))==null||y.addEventListener("click",Ee),(I=document.getElementById("submit-btn"))==null||I.addEventListener("click",De),(B=document.getElementById("btn-import-customer"))==null||B.addEventListener("click",Oe),document.querySelectorAll(".tab-btn").forEach(d=>{d.addEventListener("click",()=>{const C=d.dataset.tab;C&&C!==a.tab&&(a.tab=C,L())})});const l=document.getElementById("drop-zone"),$=document.getElementById("f-files");l&&$&&($.addEventListener("change",()=>{$.files&&ae(Array.from($.files)),$.value=""}),l.addEventListener("dragover",d=>{d.preventDefault(),l.classList.add("border-blue-500","bg-blue-50")}),l.addEventListener("dragleave",()=>{l.classList.remove("border-blue-500","bg-blue-50")}),l.addEventListener("drop",d=>{var k;d.preventDefault(),l.classList.remove("border-blue-500","bg-blue-50");const C=Array.from(((k=d.dataTransfer)==null?void 0:k.files)??[]);C.length&&ae(C)})),document.querySelectorAll(".remove-dropped").forEach(d=>{d.addEventListener("click",()=>{const C=parseInt(d.dataset.remove??"-1");C>=0&&(a.droppedFiles.splice(C,1),L())})}),Ne()}function ae(e){a.droppedFiles.push(...e),L()}document.addEventListener("paste",e=>{var o;if(!a.account)return;const t=Array.from(((o=e.clipboardData)==null?void 0:o.items)??[]),n=[];for(const i of t)if(i.kind==="file"){const s=i.getAsFile();if(s){const m=s.name&&s.name!=="image.png"?s.name:`screenshot-${new Date().toISOString().replace(/[:.]/g,"-").slice(0,19)}.png`;n.push(new File([s],m,{type:s.type}))}}n.length&&(e.preventDefault(),ae(n),T(`แนบไฟล์แล้ว: ${n.map(i=>i.name).join(", ")}`))});const b="w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white";function de(e){return e<1024?`${e} B`:e<1024*1024?`${(e/1024).toFixed(0)} KB`:`${(e/1024/1024).toFixed(1)} MB`}function z(){const e=a.emailAttachments,t=a.droppedFiles,n=e.length>0?`<div class="mb-2 space-y-1">
        <p class="text-xs text-slate-500">📎 ไฟล์แนบจาก Email:</p>
        ${e.map(i=>`
          <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
            <input type="checkbox" class="email-att-cb" data-att-id="${f(i.id)}" data-att-name="${f(i.name)}" data-att-item="${i.isItem?"1":"0"}" checked />
            <span class="flex-1 truncate">${i.isItem?"📧 ":""}${f(i.name)}</span>
            <span class="text-slate-400 flex-shrink-0">${de(i.size)}</span>
          </label>`).join("")}
      </div>`:"",o=t.length>0?`<div class="mt-2 space-y-1">
        ${t.map((i,s)=>`<div class="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 rounded px-2 py-1">
            <span class="text-base">${i.type.startsWith("image/")?"🖼️":"📄"}</span>
            <span class="flex-1 truncate">${f(i.name)}</span>
            <span class="text-slate-400">${de(i.size)}</span>
            <button type="button" data-remove="${s}"
              class="remove-dropped text-red-400 hover:text-red-600 font-bold leading-none">✕</button>
          </div>`).join("")}
      </div>`:"";return`<div class="space-y-1">
    <label class="block text-xs font-medium text-slate-600">ไฟล์แนบ</label>
    <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
      <input type="checkbox" id="f-attach-eml" />
      <span class="flex-1">📧 แนบอีเมลต้นฉบับ (.eml)</span>
    </label>
    ${n}
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
    ${o}
  </div>`}function Y(e){var t;return`<select id="f-assigned-email" class="${b}">
    <option value="${f(e)}">${f(((t=a.account)==null?void 0:t.name)??e)} (ฉัน)</option>
    ${a.agents.filter(n=>n.email!==e).map(n=>`<option value="${f(n.email)}">${f(n.name)}</option>`).join("")}
  </select>`}function Q(){return a.projects.length===0?'<div class="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-2">⚠️ ไม่พบ Project ที่ Active</div>':`<select id="f-project" class="${b}">
    <option value="">-- เลือก Project --</option>
    ${a.projects.map(e=>`<option value="${e.id}">${f(e.Title)}</option>`).join("")}
  </select>`}function x(e,t){return`
    <div class="space-y-1">
      <label class="block text-xs font-medium text-slate-600">${e}</label>
      ${t}
    </div>
  `}function f(e){return e.replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}async function He(){await j.initialize(),await j.handleRedirectPromise();const e=j.getAllAccounts();if(e.length>0){a.account=e[0];try{await j.acquireTokenSilent({scopes:[J],account:e[0]}),await Promise.all([ue(),pe(),fe(),ge()])}catch{a.account=null}}typeof Office<"u"?Office.onReady(t=>{var n;if(t.host===Office.HostType.Outlook){const o=Office.context.mailbox.item;if(o){a.emailSubject=o.subject??"";const i=o.from;i&&(a.emailSenderName=i.displayName??"",a.emailSenderEmail=i.emailAddress??"");const s=(((n=a.account)==null?void 0:n.username)??"").toLowerCase(),m=((i==null?void 0:i.emailAddress)??"").toLowerCase(),g=[...o.to??[],...o.cc??[]].map(u=>u.emailAddress).filter(Boolean);a.emailCc=[...new Set(g.map(u=>u.toLowerCase()))].filter(u=>u!==s&&u!==m);const v=o.attachments??[];a.emailAttachments=v.filter(u=>!u.isInline&&(u.attachmentType===Office.MailboxEnums.AttachmentType.File||u.attachmentType===Office.MailboxEnums.AttachmentType.Item)).map(u=>({id:u.id,name:u.attachmentType===Office.MailboxEnums.AttachmentType.Item?`${(u.name||"email").replace(/\.eml$/i,"")}.eml`:u.name,size:u.size,isItem:u.attachmentType===Office.MailboxEnums.AttachmentType.Item})),o.body.getAsync(Office.CoercionType.Html,{asyncContext:{}},u=>{if(u.status===Office.AsyncResultStatus.Succeeded){let p=function(d,C=!1){if(d.nodeType===3){const O=d.textContent??"";return C&&O.trim()===""?"":O}const k=d,N=(k.tagName??"").toLowerCase();if(P.includes(N))return"";if(N==="br")return" ";if(N==="tr"){const O=[];for(let H=0;H<k.childNodes.length;H++){const se=k.childNodes[H],ce=(se.tagName??"").toLowerCase();(ce==="td"||ce==="th")&&O.push((se.textContent??"").replace(/\s+/g," ").trim())}return O.length?O.join("	")+`
`:""}if(c.includes(N)){let O="";for(let H=0;H<k.childNodes.length;H++)O+=p(k.childNodes[H],!0);return O}let M="";for(let O=0;O<k.childNodes.length;O++)M+=p(k.childNodes[O],!1);return r.includes(N)&&(M=`
`+M.trim()+`
`),M};const E=u.value,A=new DOMParser().parseFromString(E,"text/html"),P=["style","script","head","img","meta","link","noscript"],r=["p","div","li","h1","h2","h3","h4","h5","h6","blockquote"],c=["table","thead","tbody","tfoot"],S=p(A.body??A.documentElement).replace(/[ \t]{2,}/g," ").replace(/\n[ \t]+/g,`
`).replace(/\n{3,}/g,`

`).trim().split(`
`),w=[];let h="";for(const d of S)d.trim()===""?h&&(w.push(h.trim()),h=""):d.includes("	")?(h&&(w.push(h.trim()),h=""),w.push(d)):h=h?h+" "+d.trim():d.trim();h&&w.push(h.trim());const y=w.join(`
`),I=/\n([-_]{3,}|From:\s|Best regards|Regards,|ขอแสดงความนับถือ|Sent:\s)/i,B=y.search(I);if(B>80){a.emailBodyPreview=y.slice(0,B).trim().slice(0,2e3);const d=y.slice(B).trim();a.signatureContact=Pe(d)}else a.emailBodyPreview=y.trim().slice(0,2e3),a.signatureContact=null}L()});return}}me(),L()}):(me(),L())}function me(){a.emailSubject="[DEV] Test Email Subject",a.emailSenderName="Test Sender",a.emailSenderEmail="test@example.com",a.emailBodyPreview="This is a placeholder email body for development mode."}He().catch(e=>{console.error("Init error:",e);const t=document.getElementById("app");t&&(t.innerHTML=`<div class="p-4 text-red-600 text-sm">เกิดข้อผิดพลาด: ${String(e)}</div>`)});
