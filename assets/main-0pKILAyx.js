import{P as ve}from"./PublicClientApplication-DLKYUtZW.js";const xe="0bab07cf-65e6-487c-89af-c917fc1a5a13",$e="d569b991-89fc-4a62-9df5-eb361abcef40",_="https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd",J="https://rpaexpert.sharepoint.com/.default",K=["https://graph.microsoft.com/Calendars.ReadWrite","https://graph.microsoft.com/Mail.Send","https://graph.microsoft.com/Mail.Read"],j=new ve({auth:{clientId:xe,authority:`https://login.microsoftonline.com/${$e}`,redirectUri:window.location.origin.includes("localhost")?"http://localhost:3000/":"https://darmmunginsa.github.io/itservices-addin/",navigateToLoginRequestUrl:!1},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:!1}}),we=window.location.origin.includes("localhost")?`${window.location.origin}/auth.html`:"https://darmmunginsa.github.io/itservices-addin/auth.html";function oe(){var t,n;const e=(n=(t=Office.context)==null?void 0:t.diagnostics)==null?void 0:n.platform;return e===Office.PlatformType.iOS||e===Office.PlatformType.Android}function ie(){return new Promise((e,t)=>{Office.context.ui.displayDialogAsync(we,{height:60,width:30,promptBeforeOpen:!1},n=>{if(n.status!==Office.AsyncResultStatus.Succeeded){t(new Error("เปิดหน้าเข้าสู่ระบบไม่ได้"));return}const o=n.value;o.addEventHandler(Office.EventType.DialogMessageReceived,i=>{o.close();const c=i.message;if(!c){t(new Error("auth message error"));return}try{const d=JSON.parse(c);d.ok?e():t(new Error(d.error||"auth failed"))}catch{t(new Error("auth message error"))}}),o.addEventHandler(Office.EventType.DialogEventReceived,()=>t(new Error("ปิดหน้าเข้าสู่ระบบก่อนเสร็จ")))})})}const a={account:null,tab:"ticket",emailSubject:"",emailBodyPreview:"",emailSenderName:"",emailSenderEmail:"",loading:!1,projects:[],agents:[],emailAttachments:[],signatureContact:null,droppedFiles:[],tickets:[],contactEmails:[],emailCc:[]};async function D(){const e=j.getAllAccounts();if(e.length===0)throw new Error("Not signed in");const t={scopes:[J],account:e[0]};try{return(await j.acquireTokenSilent(t)).accessToken}catch{if(oe()){await ie();const n=j.getAllAccounts()[0];if(!n)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await j.acquireTokenSilent({scopes:[J],account:n})).accessToken}return(await j.acquireTokenPopup(t)).accessToken}}async function q(e=!1){const t=j.getAllAccounts();if(t.length===0)throw new Error("Not signed in");const n={scopes:K,account:t[0],forceRefresh:e};try{return(await j.acquireTokenSilent(n)).accessToken}catch{if(oe()){await ie();const i=j.getAllAccounts()[0];if(!i)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await j.acquireTokenSilent({scopes:K,account:i})).accessToken}return(await j.acquireTokenPopup({scopes:K,account:t[0]})).accessToken}}async function ke(e){const t=await q(),n={subject:e.subject,start:{dateTime:e.start,timeZone:"Asia/Bangkok"},end:{dateTime:e.end,timeZone:"Asia/Bangkok"},body:e.body?{contentType:"HTML",content:e.body.replace(/\n/g,"<br>")}:void 0,attendees:e.attendees.filter(Boolean).map(i=>({emailAddress:{address:i},type:"required"})),isOnlineMeeting:e.isOnlineMeeting,onlineMeetingProvider:e.isOnlineMeeting?"teamsForBusiness":void 0},o=await fetch("https://graph.microsoft.com/v1.0/me/events",{method:"POST",headers:{Authorization:`Bearer ${t}`,"Content-Type":"application/json"},body:JSON.stringify(n)});if(!o.ok)throw new Error(`Calendar error ${o.status}: ${await o.text()}`)}async function ue(){try{const e=await D(),t=`${_}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$orderby=Title asc&$top=500`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const o=await n.json();a.projects=o.value.map(i=>({id:i.Id,Title:i.Title}))}}catch{}}async function pe(){try{const e=await D(),t=`${_}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText&$orderby=Title asc`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const o=await n.json();a.agents=o.value.map(i=>({email:i.EmailText,name:i.Title}))}}catch{}}async function fe(){try{const e=await D(),t=`${_}/_api/web/lists/getbytitle('HD_Tickets')/items?$select=Id,Title,TicketNumber,Status&$filter=Status ne 'Closed'&$orderby=Modified desc&$top=200`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const o=await n.json();a.tickets=o.value.map(i=>({id:i.Id,Title:i.Title,TicketNumber:i.TicketNumber,Status:i.Status}))}}catch{}}async function ge(){try{const e=await D(),t=`${_}/_api/web/lists/getbytitle('HD_Contracts')/items?$select=CustomerEmail&$top=2000`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const o=await n.json();a.contactEmails=o.value.map(i=>(i.CustomerEmail||"").trim().toLowerCase()).filter(Boolean)}}catch{}}async function le(){const e=document.getElementById("btn-login-main"),t=document.getElementById("btn-login");e&&(e.disabled=!0,e.textContent="กำลังเข้าสู่ระบบ…"),t&&(t.disabled=!0);try{if(oe()){if(await ie(),a.account=j.getAllAccounts()[0]??null,!a.account)throw new Error("เข้าสู่ระบบไม่สำเร็จ")}else{const n=await j.loginPopup({scopes:[J]});a.account=n.account}await Promise.all([ue(),pe(),fe(),ge()]),L()}catch{e&&(e.disabled=!1,e.textContent="เข้าสู่ระบบ"),t&&(t.disabled=!1)}}async function Te(){a.account&&await j.logoutPopup({account:a.account}),a.account=null,L()}async function R(e,t){const n=await D(),o=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items`,i=await fetch(o,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(t)});if(!i.ok){const d=await i.text();throw new Error(`SharePoint error ${i.status}: ${d}`)}return(await i.json()).Id}let U=null;const Z="support@itservices.co.th";async function he(){if(U)return U;try{const e=await D(),t=`${_}/_api/web/lists/getbytitle('HD_EmailTemplates')/items?$select=EventKey,Subject,Body,IsEnabled&$top=50`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});return n.ok?(U=(await n.json()).value,U):[]}catch{return[]}}async function Ee(){var e,t;try{const n=await D(),o=`${_}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title,Category&$filter=Category eq 'EmailConfig'&$top=1`,i=await fetch(o,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});return i.ok&&((t=(e=(await i.json()).value[0])==null?void 0:e.Title)==null?void 0:t.trim())||Z}catch{return Z}}function X(e,t){return e.replace(/\{\{(\w+)\}\}/g,(n,o)=>t[o]??`{{${o}}}`)}async function Ie(e,t,n,o=[]){try{const c=(await he()).find(s=>s.EventKey===e&&s.IsEnabled);if(!c)return;const d=X(c.Subject||"",t),f=X(c.Body||"",t);if(!d||!f)return;const v=s=>s.trim().toLowerCase(),m=[...new Map(n.filter(Boolean).map(s=>[v(s),s])).values()];if(m.length===0)return;const u=new Set(m.map(v)),E=[...new Map(o.filter(Boolean).map(s=>[v(s),s])).values()].filter(s=>!u.has(v(s))),A=await Ee(),P=await q(),l={subject:d,body:{contentType:"HTML",content:f},toRecipients:m.map(s=>({emailAddress:{address:s}}))};E.length&&(l.ccRecipients=E.map(s=>({emailAddress:{address:s}}))),A&&(l.from={emailAddress:{address:A}}),await fetch("https://graph.microsoft.com/v1.0/me/sendMail",{method:"POST",headers:{Authorization:`Bearer ${P}`,"Content-Type":"application/json"},body:JSON.stringify({message:l,saveToSentItems:!0})})}catch{}}async function Ce(e,t=[]){try{const n=Office.context.mailbox.item;if(!(n!=null&&n.itemId))return!1;const o=Office.context.mailbox.convertToRestId(n.itemId,Office.MailboxEnums.RestVersion.v2_0),c={Authorization:`Bearer ${await q()}`,"Content-Type":"application/json"},d=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${o}/createReplyAll`,{method:"POST",headers:c});if(!d.ok)return!1;const f=await d.json(),v=s=>s.trim().toLowerCase(),m=f.ccRecipients??[],u=new Set(m.map(s=>v(s.emailAddress.address))),E=[...new Set(t.filter(Boolean).map(s=>s.trim()))].filter(s=>!u.has(v(s))).map(s=>({emailAddress:{address:s}})),A={body:{contentType:"HTML",content:e}};return E.length&&(A.ccRecipients=[...m,...E]),(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${f.id}`,{method:"PATCH",headers:c,body:JSON.stringify(A)})).ok?(await fetch(`https://graph.microsoft.com/v1.0/me/messages/${f.id}/send`,{method:"POST",headers:c})).ok:!1}catch{return!1}}async function Ae(e,t){const o=(await he()).find(c=>c.EventKey===e&&c.IsEnabled);return o&&X(o.Body||"",t)||null}async function G(e){var c;const t=d=>d.trim().toLowerCase(),n=t(((c=a.account)==null?void 0:c.username)??""),o=new Set,i=e.recipients.filter(Boolean).filter(d=>{const f=t(d);return!f||f===n||o.has(f)?!1:(o.add(f),!0)});if(i.length!==0)try{const d=await D(),f=`${_}/_api/web/lists/getbytitle('HD_Notifications')/items`;await Promise.all(i.map(v=>fetch(f,{method:"POST",headers:{Authorization:`Bearer ${d}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:e.title.slice(0,255),RecipientEmail:v,EventType:e.eventType,Message:e.message,LinkPath:e.linkPath,IsRead:!1})})))}catch{}}async function ee(e,t){const n=document.querySelectorAll(".email-att-cb:checked");if(n.length===0)return;const o=await D();for(const i of Array.from(n)){const c=i.dataset.attId,d=i.dataset.attName,f=await new Promise((l,s)=>{Office.context.mailbox.item.getAttachmentContentAsync(c,{},g=>{g.status===Office.AsyncResultStatus.Succeeded?l(g):s(new Error(g.error.message))})}),{content:v,format:m}=f.value;let u;if(m===Office.MailboxEnums.AttachmentContentFormat.Base64){const l=atob(v);u=new Uint8Array(l.length);for(let s=0;s<l.length;s++)u[s]=l.charCodeAt(s)}else if(m===Office.MailboxEnums.AttachmentContentFormat.Eml||m===Office.MailboxEnums.AttachmentContentFormat.ICalendar)u=new TextEncoder().encode(v);else continue;const E=encodeURIComponent(d),A=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${E}')`;if(!(await fetch(A,{method:"POST",headers:{Authorization:`Bearer ${o}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:u.buffer})).ok)throw new Error(`Upload ${d} failed`)}}async function Be(e){const t=`https://graph.microsoft.com/v1.0/me/messages/${e}/$value`;let n=await q(),o=await fetch(t,{headers:{Authorization:`Bearer ${n}`}});if((o.status===401||o.status===403)&&(n=await q(!0),o=await fetch(t,{headers:{Authorization:`Bearer ${n}`}})),!o.ok)throw new Error(`Graph ${o.status}`);return o.arrayBuffer()}async function Se(e){const t=await new Promise((o,i)=>{Office.context.mailbox.getCallbackTokenAsync({isRest:!0},c=>{c.status===Office.AsyncResultStatus.Succeeded?o(c.value):i(new Error("callback token failed"))})}),n=await fetch(`${Office.context.mailbox.restUrl}/v2.0/me/messages/${e}/$value`,{headers:{Authorization:`Bearer ${t}`}});if(!n.ok)throw new Error(`REST ${n.status}`);return n.arrayBuffer()}async function te(e,t){const n=document.getElementById("f-attach-eml");if(!(n!=null&&n.checked))return;const o=Office.context.mailbox.item;if(!o)return;const i=Office.context.mailbox.convertToRestId(o.itemId,Office.MailboxEnums.RestVersion.v2_0);let c,d="",f="";try{c=await Be(i)}catch(A){d=A instanceof Error?A.message:String(A);try{c=await Se(i)}catch(P){f=P instanceof Error?P.message:String(P),console.error("[eml] graph:",d,"| callback:",f),T(`ดึง .eml ไม่ได้ (Graph: ${d} / REST: ${f}) — ไฟล์อื่นบันทึกแล้ว`,"error");return}}const v=(o.subject||"email").replace(/[\\/:*?"<>|#%&{}~]/g,"_").slice(0,100).trim()||"email",m=await D(),u=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${encodeURIComponent(v+".eml")}')`;(await fetch(u,{method:"POST",headers:{Authorization:`Bearer ${m}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:c})).ok||T("บันทึกไฟล์ .eml ไม่สำเร็จ","error")}async function ne(e,t,n){const o=await D();for(const i of n){const c=await i.arrayBuffer(),d=encodeURIComponent(i.name),f=`${_}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${d}')`;if(!(await fetch(f,{method:"POST",headers:{Authorization:`Bearer ${o}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:c})).ok)throw new Error(`Upload ${i.name} failed`)}}function T(e,t="success"){const n=document.getElementById("toast-container");if(!n)return;const o=t==="success"?"bg-green-500":"bg-red-500",i=t==="success"?"✅":"❌",c=document.createElement("div");c.className=`toast pointer-events-auto ${o} text-white text-sm font-medium px-4 py-3 rounded-lg shadow-lg max-w-xs mx-2`,c.textContent=`${i} ${e}`,n.appendChild(c),setTimeout(()=>c.remove(),4e3)}function je(e){const t=e.split(`
`).map(u=>u.trim()).filter(Boolean),n=/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/,o=/(\+?[\d\s()\-]{7,}(?:ext\.?\s*\d+)?)/i,i=/\b(co\.?,?\s*ltd\.?|co\.?,?\s*limited|corporation|corp\.?|บริษัท|จำกัด|holding|group|inc\.?|llc)\b/i;let c="",d="",f="";const v=[];for(const u of t)if(!/^[-_=*]{2,}$/.test(u)&&!/^(best regards|regards|sincerely|ขอแสดงความนับถือ|ด้วยความนับถือ|from|sent|thanks|thank you)/i.test(u)){if(!c){const E=u.match(n);if(E){c=E[0];continue}}if(!d){const E=u.match(o);if(E&&E[0].replace(/\D/g,"").length>=7){d=E[0].trim();continue}}if(!f&&i.test(u)){f=u;continue}u.length>=2&&u.length<=50&&!/\d{4,}/.test(u)&&v.push(u)}const m=v.find(u=>!n.test(u)&&!i.test(u))??"";return!c&&!m?null:{name:m,company:f,email:c,phone:d}}async function Pe(){const e=a.signatureContact;if(!e)return;const t=(a.emailSenderEmail||"").toLowerCase();if(t&&a.contactEmails.includes(t)){T("ลูกค้านี้มีในระบบแล้ว","success"),a.signatureContact=null,L();return}const n=document.getElementById("btn-import-customer");n&&(n.disabled=!0,n.textContent="กำลังบันทึก…");try{await R("HD_Contracts",{Title:a.emailSenderName||e.name,CustomerEmail:a.emailSenderEmail,Phone:e.phone||void 0,Company:e.company||void 0,Status:"Active"}),t&&a.contactEmails.push(t),T("เพิ่มลูกค้าสำเร็จ!"),a.signatureContact=null,L()}catch(o){const i=o instanceof Error?o.message:String(o);T(`เกิดข้อผิดพลาด: ${i}`,"error"),n&&(n.disabled=!1,n.textContent="เพิ่มเป็นลูกค้า")}}function be(){return new Date().toISOString().split("T")[0]}function Oe(){const e=new Date;return`HD-${`${e.getFullYear()}${String(e.getMonth()+1).padStart(2,"0")}${String(e.getDate()).padStart(2,"0")}`}-${Math.floor(Math.random()*900+100)}`}function re(){var e;return a.droppedFiles.length>0||document.querySelectorAll(".email-att-cb:checked").length>0||(((e=document.getElementById("f-attach-eml"))==null?void 0:e.checked)??!1)}async function V(e,t){a.droppedFiles.length>0&&await ne(e,t,a.droppedFiles),await ee(e,t),await te(e,t)}let W=!1;async function _e(){var t,n,o,i,c,d,f,v,m,u,E,A,P;if(!a.account){T("กรุณาเข้าสู่ระบบก่อน","error");return}if(W)return;W=!0;const e=document.getElementById("submit-btn");e&&(e.disabled=!0,e.textContent="กำลังบันทึก…");try{if(a.tab==="ticket"){const l=document.getElementById("f-title").value.trim(),s=document.getElementById("f-description").value.trim(),g=document.getElementById("f-priority").value,$=document.getElementById("f-customer-email").value.trim(),w=((t=document.getElementById("f-cc-enable"))==null?void 0:t.checked)??!0?(((n=document.getElementById("f-cc"))==null?void 0:n.value)||"").split(/[,;\s]+/).map(M=>M.trim()).filter(Boolean):[],h=document.getElementById("f-assigned-email").value,y=a.agents.find(M=>M.email===h),I=Oe(),S=await R("HD_Tickets",{Title:l,TicketNumber:I,Description:s,Priority:g,CustomerEmail:$,CustomerName:a.emailSenderName||$,Status:"Open",AssignedEmail:h||void 0,AssignedToName:(y==null?void 0:y.name)??((o=a.account)==null?void 0:o.name)??""});if(re()){const M=await R("HD_TicketComments",{Title:"📎 ไฟล์แนบจากอีเมล",TicketID:S,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await V("HD_TicketComments",M)}a.droppedFiles=[];const r={ticket_number:I,ticket_title:l,priority:g,category:"-",description:(s||"-").replace(/\n/g,"<br>"),customer_name:a.emailSenderName||$,assigned_name:(y==null?void 0:y.name)??((i=a.account)==null?void 0:i.name)??"-",link:"https://itservices.co.th/helpdesk/"},C=[h,a.account.username,...w].filter(Boolean);let k=!1;const N=await Ae("ticket_created",r);if(N){const M=`<div style="border-left:4px solid #2563eb;background:#eff6ff;padding:10px 14px;margin:0 0 14px;font-family:Segoe UI,sans-serif">
             <div style="font-size:15px;font-weight:700;color:#1e40af">Ticket No. ${I}</div>
             <div style="font-size:12px;color:#475569;margin-top:2px">กรุณาตอบกลับในอีเมลฉบับนี้เพื่อให้ข้อมูลอยู่ใน Ticket เดียวกัน</div>
           </div>`;k=await Ce(M+N,C)}k||await Ie("ticket_created",r,[$],C),T(k?"สร้าง Ticket และตอบกลับในเธรดเดิมแล้ว!":"สร้าง Ticket สำเร็จ!")}else if(a.tab==="task"){const l=document.getElementById("f-title").value.trim(),s=parseInt(((c=document.getElementById("f-project"))==null?void 0:c.value)||"0"),g=document.getElementById("f-due-date").value,$=document.getElementById("f-note").value.trim(),B=document.getElementById("f-assigned-email").value,w=a.agents.find(I=>I.email===B);if(!s){T("กรุณาเลือก Project","error");return}const h=await R("PM_Tasks",{Title:l,DueDate:g||null,TaskNote:$,AssignedTo:(w==null?void 0:w.name)??a.account.name??a.account.username,AssignedEmail:B,IsCompleted:!1,IsAcknowledged:!1,ProjectID:s});if(a.droppedFiles.length>0&&await ne("PM_Tasks",h,a.droppedFiles),await ee("PM_Tasks",h),await te("PM_Tasks",h),a.droppedFiles=[],await G({recipients:[B],title:`📋 ได้รับมอบหมาย Task: ${l}`,message:$||(g?`กำหนดส่ง ${g}`:"มี Task ใหม่"),linkPath:s?`/projects/${s}`:"/my-work",eventType:"task_assigned"}),((d=document.getElementById("f-teams"))==null?void 0:d.checked)&&g){const I=Array.from(document.querySelectorAll(".att-internal:checked")).map(k=>k.value),S=(((f=document.getElementById("f-ext-att"))==null?void 0:f.value)||"").split(/[,;\s]+/).map(k=>k.trim()).filter(Boolean),r=`${g}T09:00:00`,C=`${g}T10:00:00`;try{await ke({subject:l,start:r,end:C,body:$,attendees:[...I,...S],isOnlineMeeting:!0}),T("สร้าง Task + นัดประชุม Teams สำเร็จ!")}catch(k){T("สร้าง Task แล้ว แต่สร้างนัดประชุมไม่สำเร็จ: "+(k instanceof Error?k.message:""),"error")}}else T("สร้าง Task สำเร็จ!")}else if(a.tab==="incident"){const l=document.getElementById("f-title").value.trim(),s=parseInt(((v=document.getElementById("f-project"))==null?void 0:v.value)||"0"),g=document.getElementById("f-description").value.trim(),$=document.getElementById("f-severity").value,B=document.getElementById("f-assigned-email").value,w=a.agents.find(r=>r.email===B),h=document.getElementById("f-status").value,y=document.getElementById("f-incident-date").value,I=document.getElementById("f-resolution").value.trim();if(!s){T("กรุณาเลือก Project","error");return}const S=await R("PM_Incidents",{Title:l,Description:g||void 0,Severity:$,Status:h,AssignedTo:(w==null?void 0:w.name)??a.account.name??a.account.username,AssignedEmail:B,ProjectID:s,IncidentDate:y||be(),Resolution:I||void 0});a.droppedFiles.length>0&&await ne("PM_Incidents",S,a.droppedFiles),await ee("PM_Incidents",S),await te("PM_Incidents",S),a.droppedFiles=[],await G({recipients:[B],title:`🚨 ได้รับมอบหมาย Incident: ${l}`,message:`ความรุนแรง ${$}${g?" — "+g.slice(0,120):""}`,linkPath:s?`/projects/${s}`:"/my-work",eventType:"incident_created"}),T("สร้าง Incident สำเร็จ!")}else if(a.tab==="comment"){const l=parseInt(((m=document.getElementById("f-ticket"))==null?void 0:m.value)||"0"),s=document.getElementById("f-comment").value.trim(),g=document.getElementById("f-comment-type").value;if(!l){T("กรุณาเลือก Ticket","error");return}if(!s){T("กรุณาพิมพ์ Comment","error");return}const $=await R("HD_TicketComments",{Title:s.slice(0,100),TicketID:l,CommentText:s,CommentType:g,CommentDate:new Date().toISOString()});await V("HD_TicketComments",$),a.droppedFiles=[];try{const B=await D(),w=`${_}/_api/web/lists/getbytitle('HD_Tickets')/items(${l})?$select=TicketNumber,Title,AssignedEmail,Author/EMail&$expand=Author`,h=await fetch(w,{headers:{Authorization:`Bearer ${B}`,Accept:"application/json;odata=nometadata"}});if(h.ok){const y=await h.json(),I=a.account.username.toLowerCase(),S=[...new Set([y.AssignedEmail,(u=y.Author)==null?void 0:u.EMail].filter(Boolean))].filter(r=>r.toLowerCase()!==I);S.length&&await G({recipients:S,title:`💬 ${((E=a.account)==null?void 0:E.name)??"มีคน"} คอมเมนต์ใน ${y.TicketNumber||"#"+l}`,message:s.slice(0,200),linkPath:`/tickets/${l}`,eventType:"comment_added"})}}catch{}T("เพิ่ม Comment สำเร็จ!")}else if(a.tab==="project"){const l=document.getElementById("f-title").value.trim(),s=document.getElementById("f-company").value.trim(),g=document.getElementById("f-group").value,$=document.getElementById("f-status").value,B=document.getElementById("f-start").value,w=document.getElementById("f-end").value,h=document.getElementById("f-description").value.trim();if(!l){T("กรุณาใส่ชื่อโครงการ","error");return}const y=await R("PM_Projects",{Title:l,Company:s||void 0,ProjectGroup:g,Progress:0,StartDate:B||void 0,EndDate:w||null,Status:$,CreatedByEmail:a.account.username,Comment:h||void 0});if(re()){const I=await R("PM_Comments",{Title:"📎 ไฟล์แนบจากอีเมล",ProjectID:y,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await V("PM_Comments",I)}a.droppedFiles=[],T("สร้างโครงการสำเร็จ!")}else if(a.tab==="projcomment"){const l=parseInt(((A=document.getElementById("f-project"))==null?void 0:A.value)||"0"),s=document.getElementById("f-comment").value.trim(),g=document.getElementById("f-comment-type").value;if(!l){T("กรุณาเลือกโครงการ","error");return}if(!s){T("กรุณาพิมพ์ Comment","error");return}const $=await R("PM_Comments",{Title:s.slice(0,100),ProjectID:l,CommentText:s,CommentType:g,CommentDate:new Date().toISOString(),ParentID:0});await V("PM_Comments",$),a.droppedFiles=[];try{const B=await D(),w=`${_}/_api/web/lists/getbytitle('PM_Projects')/items(${l})?$select=Title,CreatedByEmail`,h=await fetch(w,{headers:{Authorization:`Bearer ${B}`,Accept:"application/json;odata=nometadata"}});if(h.ok){const y=await h.json(),I=a.account.username.toLowerCase();y.CreatedByEmail&&y.CreatedByEmail.toLowerCase()!==I&&await G({recipients:[y.CreatedByEmail],title:`💬 ${((P=a.account)==null?void 0:P.name)??"มีคน"} คอมเมนต์ในโครงการ ${y.Title??""}`,message:s.slice(0,200),linkPath:`/projects/${l}?tab=comments`,eventType:"comment_added"})}}catch{}T("เพิ่ม Comment สำเร็จ!")}}catch(l){const s=l instanceof Error?l.message:String(l);T(`เกิดข้อผิดพลาด: ${s}`,"error")}finally{W=!1,e&&(e.disabled=!1,e.textContent="บันทึก")}}const De={ticket:{label:"Ticket",icon:"🎫"},comment:{label:"Comment",icon:"💬"},project:{label:"Project",icon:"📁"},task:{label:"Task",icon:"✅"},incident:{label:"Incident",icon:"🚨"},projcomment:{label:"Comment",icon:"💬"}},Me=[{title:"🎫 Helpdesk",tabs:["ticket","comment"]},{title:"📁 Project",tabs:["project","task","incident","projcomment"]}],ye=["f-title","f-description","f-priority","f-customer-email","f-cc","f-assigned-email","f-project","f-due-date","f-note","f-severity","f-status","f-incident-date","f-resolution","f-ticket","f-comment","f-comment-type","f-company","f-group","f-start","f-end","f-ext-att","f-attach-eml"];let F={};function Le(){for(const t of ye){const n=document.getElementById(t);n&&(F[t]=n.value)}const e=document.getElementById("f-teams");e&&(F["f-teams"]=e.checked)}function Re(){for(const t of ye){const n=document.getElementById(t);n&&F[t]!==void 0&&F[t]!==""&&(n.value=F[t])}const e=document.getElementById("f-teams");if(e&&F["f-teams"]!==void 0){e.checked=F["f-teams"];const t=document.getElementById("teams-fields");t&&(t.style.display=e.checked?"block":"none")}}function L(){var B,w,h,y,I,S;const e=document.getElementById("app");if(!e)return;Le();const{account:t,tab:n,emailSubject:o,emailSenderName:i,emailSenderEmail:c,emailBodyPreview:d}=a,f=t!==null,v=`
    <div class="bg-gradient-to-r from-blue-700 to-blue-600 text-white px-4 py-2.5 flex items-center gap-2.5 shadow flex-shrink-0">
      <div class="w-7 h-7 bg-white/95 rounded-lg flex items-center justify-center flex-shrink-0">
        <span class="text-blue-700 font-extrabold text-xs">iT</span>
      </div>
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm leading-tight truncate">iT Services Helpdesk</div>
        ${f?`<div class="text-[10px] text-blue-100 truncate">${p((t==null?void 0:t.name)??(t==null?void 0:t.username)??"")}</div>`:""}
      </div>
      ${f?`<button id="btn-logout" title="ออกจากระบบ" class="p-1.5 rounded-lg hover:bg-white/15 transition flex-shrink-0">
             <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
           </button>`:`<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded-lg transition flex-shrink-0">
             เข้าสู่ระบบ
           </button>`}
    </div>
  `;if(!f){e.innerHTML=`
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
    `,(B=document.getElementById("btn-login"))==null||B.addEventListener("click",le),(w=document.getElementById("btn-login-main"))==null||w.addEventListener("click",le);return}const m=o?`<div class="mx-3 mt-3 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5 text-xs text-slate-700 space-y-1">
         <div class="font-semibold text-blue-800 truncate" title="${p(o)}">📧 ${p(o)}</div>
         ${i?`<div class="text-slate-500">จาก: <span class="font-medium text-slate-700">${p(i)}</span></div>`:""}
         ${c&&c!==i?`<div class="text-slate-400 truncate">${p(c)}</div>`:""}
       </div>`:`<div class="mx-3 mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-700">
         ⚠️ ไม่พบข้อมูล Email (โหมดทดสอบ)
       </div>`,u=a.signatureContact,E=!!c&&a.contactEmails.includes(c.toLowerCase()),A=u?`<div class="mx-3 mt-3 bg-orange-50 border border-orange-200 rounded-lg px-3 py-2.5 text-xs text-slate-700">
         <div class="flex items-center justify-between mb-2">
           <span class="font-semibold text-orange-700">👤 ข้อมูลผู้ส่ง (จาก Signature)</span>
         </div>
         <div class="space-y-0.5 mb-2.5">
           ${i?`<div><span class="text-slate-400">ชื่อ:</span> <span class="font-medium">${p(i)}</span></div>`:""}
           ${u.company?`<div><span class="text-slate-400">บริษัท:</span> ${p(u.company)}</div>`:""}
           ${c?`<div><span class="text-slate-400">Email:</span> ${p(c)}</div>`:""}
           ${u.phone?`<div><span class="text-slate-400">โทร:</span> ${p(u.phone)}</div>`:""}
         </div>
         ${E?'<div class="w-full bg-green-100 text-green-700 text-xs font-semibold py-1.5 rounded-md text-center">✓ ลูกค้านี้มีในระบบแล้ว</div>':`<button id="btn-import-customer"
                class="w-full bg-orange-500 hover:bg-orange-400 text-white text-xs font-semibold py-1.5 rounded-md transition">
                + เพิ่มเป็นลูกค้า
              </button>`}
       </div>`:"",P=`
    <div class="mx-3 mt-3 space-y-2">
      ${Me.map(r=>`
        <div>
          <div class="text-[10px] font-semibold text-slate-400 uppercase tracking-wide mb-1 px-0.5">${r.title}</div>
          <div class="grid grid-cols-4 gap-1">
            ${r.tabs.map(C=>{const k=De[C];return`<button data-tab="${C}"
                class="tab-btn flex flex-col items-center gap-1 py-2 rounded-lg transition ${n===C?"bg-blue-700 text-white shadow":"text-slate-500 hover:bg-slate-100"}">
                <span class="text-base leading-none">${k.icon}</span>
                <span class="text-[9px] font-medium leading-none">${k.label}</span>
              </button>`}).join("")}
          </div>
        </div>
      `).join("")}
    </div>
  `;let l="";n==="ticket"?l=`
      ${x("Title / หัวข้อ",`<input id="f-title" type="text"
        class="${b}"
        value="${p(o)}" />`)}
      ${x("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${b} resize-none">${p(d)}</textarea>`)}
      ${x("Priority",`<select id="f-priority" class="${b}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
      ${x("Customer Email",`<input id="f-customer-email" type="email"
        class="${b}"
        value="${p(c)}" />`)}
      ${x("CC — ให้ผู้ที่อยู่ในเมลนี้รับรู้",`
        <label class="flex items-center gap-2 text-xs text-slate-600 mb-1.5 cursor-pointer">
          <input id="f-cc-enable" type="checkbox" ${a.emailCc.length?"checked":""} /> แนบผู้รับในเมลนี้เป็น CC อัตโนมัติ
        </label>
        <input id="f-cc" type="text" class="${b}" value="${p(a.emailCc.join(", "))}" placeholder="someone@company.com, boss@company.com" />`)}
      ${x("Assign ให้ Agent",Y(t.username))}
      ${z()}
    `:n==="task"?l=`
      ${x("ชื่อ Task *",`<input id="f-title" type="text" required
        class="${b}" value="${p(o)}" />`)}
      ${x("Project *",Q())}
      ${x("Assign ให้",Y(t.username))}
      ${x("Due Date",`<input id="f-due-date" type="date" class="${b}" />`)}
      ${x("Task Note",`<textarea id="f-note" rows="4"
        class="${b} resize-y">${p(d)}</textarea>`)}
      <label class="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer mb-1">
        <input id="f-teams" type="checkbox" class="rounded" onchange="document.getElementById('teams-fields').style.display=this.checked?'block':'none'" />
        💻 เพิ่มการประชุมออนไลน์ (Teams) — ใช้เวลา 09:00–10:00 ของวัน Due Date
      </label>
      <div id="teams-fields" style="display:none" class="space-y-2 mb-2">
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">ผู้เข้าร่วม Internal</label>
          <div class="max-h-28 overflow-y-auto border border-slate-200 rounded-md p-1.5 space-y-0.5">
            ${a.agents.map(r=>`<label class="flex items-center gap-2 text-xs text-slate-700 px-1 py-0.5 hover:bg-slate-50 rounded cursor-pointer">
              <input type="checkbox" class="att-internal" value="${p(r.email)}" /> ${p(r.name)}
            </label>`).join("")}
          </div>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">Email ภายนอก (คั่นด้วย ,)</label>
          <input id="f-ext-att" type="text" class="${b}" placeholder="someone@company.com, ..." />
        </div>
      </div>
      ${z()}
    `:n==="incident"?l=`
      ${x("ชื่อ Incident *",`<input id="f-title" type="text" required
        class="${b}" value="${p(o)}" />`)}
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
      ${x("วันที่เกิด Incident",`<input id="f-incident-date" type="date" class="${b}" value="${be()}" />`)}
      ${x("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${b} resize-y">${p(d)}</textarea>`)}
      ${x("วิธีแก้ไข (ถ้ามี)",`<textarea id="f-resolution" rows="2"
        class="${b} resize-y" placeholder="อธิบายวิธีแก้ไขปัญหา..."></textarea>`)}
      ${z()}
    `:n==="comment"?l=`
      ${x("เลือก Ticket *",`<select id="f-ticket" class="${b}">
        <option value="">-- เลือก Ticket ที่จะเพิ่ม Comment --</option>
        ${a.tickets.map(r=>`<option value="${r.id}">${p(r.TicketNumber||"#"+r.id)} · ${p(r.Title)}</option>`).join("")}
      </select>`)}
      ${x("ประเภท",`<select id="f-comment-type" class="${b}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${x("Comment *",`<textarea id="f-comment" rows="5"
        class="${b} resize-y" placeholder="พิมพ์ comment...">${p(d)}</textarea>`)}
      ${z()}
    `:n==="project"?l=`
      ${x("ชื่อโครงการ *",`<input id="f-title" type="text" required
        class="${b}" value="${p(o)}" />`)}
      ${x("บริษัท / ลูกค้า",`<input id="f-company" type="text" class="${b}" value="${p(((h=a.signatureContact)==null?void 0:h.company)??"")}" />`)}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">กลุ่มโครงการ</label>
          <select id="f-group" class="${b}">
            ${["Internal","External","R&D","Maintenance","อื่นๆ"].map(r=>`<option>${r}</option>`).join("")}
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${b}">
            ${["Planning","Active","On Hold","Completed","Cancelled"].map(r=>`<option>${r}</option>`).join("")}
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
        class="${b} resize-y">${p(d)}</textarea>`)}
      ${z()}
    `:n==="projcomment"&&(l=`
      ${x("เลือกโครงการ *",Q())}
      ${x("ประเภท",`<select id="f-comment-type" class="${b}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${x("Comment *",`<textarea id="f-comment" rows="5"
        class="${b} resize-y" placeholder="พิมพ์ comment...">${p(d)}</textarea>`)}
      ${z()}
    `);const s=n==="comment"||n==="projcomment"?"เพิ่ม Comment":n==="project"?"สร้างโครงการ":n==="incident"?"แจ้ง Incident":n==="task"?"สร้าง Task":"สร้าง Ticket";e.innerHTML=`
    <div class="flex flex-col h-screen bg-slate-50">
      ${v}
      <div class="flex-1 overflow-y-auto">
        ${m}
        ${A}
        ${P}
        <div class="mx-3 mt-3 space-y-3 pb-3">
          ${l}
        </div>
      </div>
      <div class="border-t border-slate-200 bg-white px-3 py-2.5 flex-shrink-0">
        <button id="submit-btn"
          class="w-full bg-blue-700 hover:bg-blue-600 active:bg-blue-800 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          ${s}
        </button>
      </div>
    </div>
  `,(y=document.getElementById("btn-logout"))==null||y.addEventListener("click",Te),(I=document.getElementById("submit-btn"))==null||I.addEventListener("click",_e),(S=document.getElementById("btn-import-customer"))==null||S.addEventListener("click",Pe),document.querySelectorAll(".tab-btn").forEach(r=>{r.addEventListener("click",()=>{const C=r.dataset.tab;C&&C!==a.tab&&(a.tab=C,L())})});const g=document.getElementById("drop-zone"),$=document.getElementById("f-files");g&&$&&($.addEventListener("change",()=>{$.files&&ae(Array.from($.files)),$.value=""}),g.addEventListener("dragover",r=>{r.preventDefault(),g.classList.add("border-blue-500","bg-blue-50")}),g.addEventListener("dragleave",()=>{g.classList.remove("border-blue-500","bg-blue-50")}),g.addEventListener("drop",r=>{var k;r.preventDefault(),g.classList.remove("border-blue-500","bg-blue-50");const C=Array.from(((k=r.dataTransfer)==null?void 0:k.files)??[]);C.length&&ae(C)})),document.querySelectorAll(".remove-dropped").forEach(r=>{r.addEventListener("click",()=>{const C=parseInt(r.dataset.remove??"-1");C>=0&&(a.droppedFiles.splice(C,1),L())})}),Re()}function ae(e){a.droppedFiles.push(...e),L()}document.addEventListener("paste",e=>{var o;if(!a.account)return;const t=Array.from(((o=e.clipboardData)==null?void 0:o.items)??[]),n=[];for(const i of t)if(i.kind==="file"){const c=i.getAsFile();if(c){const d=c.name&&c.name!=="image.png"?c.name:`screenshot-${new Date().toISOString().replace(/[:.]/g,"-").slice(0,19)}.png`;n.push(new File([c],d,{type:c.type}))}}n.length&&(e.preventDefault(),ae(n),T(`แนบไฟล์แล้ว: ${n.map(i=>i.name).join(", ")}`))});const b="w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white";function de(e){return e<1024?`${e} B`:e<1024*1024?`${(e/1024).toFixed(0)} KB`:`${(e/1024/1024).toFixed(1)} MB`}function z(){const e=a.emailAttachments,t=a.droppedFiles,n=e.length>0?`<div class="mb-2 space-y-1">
        <p class="text-xs text-slate-500">📎 ไฟล์แนบจาก Email:</p>
        ${e.map(i=>`
          <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
            <input type="checkbox" class="email-att-cb" data-att-id="${p(i.id)}" data-att-name="${p(i.name)}" data-att-item="${i.isItem?"1":"0"}" checked />
            <span class="flex-1 truncate">${i.isItem?"📧 ":""}${p(i.name)}</span>
            <span class="text-slate-400 flex-shrink-0">${de(i.size)}</span>
          </label>`).join("")}
      </div>`:"",o=t.length>0?`<div class="mt-2 space-y-1">
        ${t.map((i,c)=>`<div class="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 rounded px-2 py-1">
            <span class="text-base">${i.type.startsWith("image/")?"🖼️":"📄"}</span>
            <span class="flex-1 truncate">${p(i.name)}</span>
            <span class="text-slate-400">${de(i.size)}</span>
            <button type="button" data-remove="${c}"
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
    <option value="${p(e)}">${p(((t=a.account)==null?void 0:t.name)??e)} (ฉัน)</option>
    ${a.agents.filter(n=>n.email!==e).map(n=>`<option value="${p(n.email)}">${p(n.name)}</option>`).join("")}
  </select>`}function Q(){return a.projects.length===0?'<div class="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-2">⚠️ ไม่พบ Project ที่ Active</div>':`<select id="f-project" class="${b}">
    <option value="">-- เลือก Project --</option>
    ${a.projects.map(e=>`<option value="${e.id}">${p(e.Title)}</option>`).join("")}
  </select>`}function x(e,t){return`
    <div class="space-y-1">
      <label class="block text-xs font-medium text-slate-600">${e}</label>
      ${t}
    </div>
  `}function p(e){return e.replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}async function Ne(){await j.initialize(),await j.handleRedirectPromise();const e=j.getAllAccounts();if(e.length>0){a.account=e[0];try{await j.acquireTokenSilent({scopes:[J],account:e[0]}),await Promise.all([ue(),pe(),fe(),ge()])}catch{a.account=null}}typeof Office<"u"?Office.onReady(t=>{var n;if(t.host===Office.HostType.Outlook){const o=Office.context.mailbox.item;if(o){a.emailSubject=o.subject??"";const i=o.from;i&&(a.emailSenderName=i.displayName??"",a.emailSenderEmail=i.emailAddress??"");const c=(((n=a.account)==null?void 0:n.username)??"").toLowerCase(),d=((i==null?void 0:i.emailAddress)??"").toLowerCase(),f=[...o.to??[],...o.cc??[]].map(m=>m.emailAddress).filter(Boolean);a.emailCc=[...new Set(f.map(m=>m.toLowerCase()))].filter(m=>m!==c&&m!==d);const v=o.attachments??[];a.emailAttachments=v.filter(m=>!m.isInline&&(m.attachmentType===Office.MailboxEnums.AttachmentType.File||m.attachmentType===Office.MailboxEnums.AttachmentType.Item)).map(m=>({id:m.id,name:m.attachmentType===Office.MailboxEnums.AttachmentType.Item?`${(m.name||"email").replace(/\.eml$/i,"")}.eml`:m.name,size:m.size,isItem:m.attachmentType===Office.MailboxEnums.AttachmentType.Item})),o.body.getAsync(Office.CoercionType.Html,{asyncContext:{}},m=>{if(m.status===Office.AsyncResultStatus.Succeeded){let u=function(r,C=!1){if(r.nodeType===3){const O=r.textContent??"";return C&&O.trim()===""?"":O}const k=r,N=(k.tagName??"").toLowerCase();if(P.includes(N))return"";if(N==="br")return" ";if(N==="tr"){const O=[];for(let H=0;H<k.childNodes.length;H++){const se=k.childNodes[H],ce=(se.tagName??"").toLowerCase();(ce==="td"||ce==="th")&&O.push((se.textContent??"").replace(/\s+/g," ").trim())}return O.length?O.join("	")+`
`:""}if(s.includes(N)){let O="";for(let H=0;H<k.childNodes.length;H++)O+=u(k.childNodes[H],!0);return O}let M="";for(let O=0;O<k.childNodes.length;O++)M+=u(k.childNodes[O],!1);return l.includes(N)&&(M=`
`+M.trim()+`
`),M};const E=m.value,A=new DOMParser().parseFromString(E,"text/html"),P=["style","script","head","img","meta","link","noscript"],l=["p","div","li","h1","h2","h3","h4","h5","h6","blockquote"],s=["table","thead","tbody","tfoot"],B=u(A.body??A.documentElement).replace(/[ \t]{2,}/g," ").replace(/\n[ \t]+/g,`
`).replace(/\n{3,}/g,`

`).trim().split(`
`),w=[];let h="";for(const r of B)r.trim()===""?h&&(w.push(h.trim()),h=""):r.includes("	")?(h&&(w.push(h.trim()),h=""),w.push(r)):h=h?h+" "+r.trim():r.trim();h&&w.push(h.trim());const y=w.join(`
`),I=/\n([-_]{3,}|From:\s|Best regards|Regards,|ขอแสดงความนับถือ|Sent:\s)/i,S=y.search(I);if(S>80){a.emailBodyPreview=y.slice(0,S).trim().slice(0,2e3);const r=y.slice(S).trim();a.signatureContact=je(r)}else a.emailBodyPreview=y.trim().slice(0,2e3),a.signatureContact=null}L()});return}}me(),L()}):(me(),L())}function me(){a.emailSubject="[DEV] Test Email Subject",a.emailSenderName="Test Sender",a.emailSenderEmail="test@example.com",a.emailBodyPreview="This is a placeholder email body for development mode."}Ne().catch(e=>{console.error("Init error:",e);const t=document.getElementById("app");t&&(t.innerHTML=`<div class="p-4 text-red-600 text-sm">เกิดข้อผิดพลาด: ${String(e)}</div>`)});
