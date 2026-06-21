import{P as ye}from"./PublicClientApplication-DLKYUtZW.js";const ve="0bab07cf-65e6-487c-89af-c917fc1a5a13",xe="d569b991-89fc-4a62-9df5-eb361abcef40",O="https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd",V="https://rpaexpert.sharepoint.com/.default",J=["https://graph.microsoft.com/Calendars.ReadWrite","https://graph.microsoft.com/Mail.Send","https://graph.microsoft.com/Mail.Read"],S=new ye({auth:{clientId:ve,authority:`https://login.microsoftonline.com/${xe}`,redirectUri:window.location.origin.includes("localhost")?"http://localhost:3000/":"https://darmmunginsa.github.io/itservices-addin/",navigateToLoginRequestUrl:!1},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:!1}}),$e=window.location.origin.includes("localhost")?`${window.location.origin}/auth.html`:"https://darmmunginsa.github.io/itservices-addin/auth.html";function ae(){var t,n;const e=(n=(t=Office.context)==null?void 0:t.diagnostics)==null?void 0:n.platform;return e===Office.PlatformType.iOS||e===Office.PlatformType.Android}function oe(){return new Promise((e,t)=>{Office.context.ui.displayDialogAsync($e,{height:60,width:30,promptBeforeOpen:!1},n=>{if(n.status!==Office.AsyncResultStatus.Succeeded){t(new Error("เปิดหน้าเข้าสู่ระบบไม่ได้"));return}const i=n.value;i.addEventHandler(Office.EventType.DialogMessageReceived,o=>{i.close();const s=o.message;if(!s){t(new Error("auth message error"));return}try{const d=JSON.parse(s);d.ok?e():t(new Error(d.error||"auth failed"))}catch{t(new Error("auth message error"))}}),i.addEventHandler(Office.EventType.DialogEventReceived,()=>t(new Error("ปิดหน้าเข้าสู่ระบบก่อนเสร็จ")))})})}const a={account:null,tab:"ticket",emailSubject:"",emailBodyPreview:"",emailSenderName:"",emailSenderEmail:"",loading:!1,projects:[],agents:[],emailAttachments:[],signatureContact:null,droppedFiles:[],tickets:[],contactEmails:[],emailCc:[]};async function _(){const e=S.getAllAccounts();if(e.length===0)throw new Error("Not signed in");const t={scopes:[V],account:e[0]};try{return(await S.acquireTokenSilent(t)).accessToken}catch{if(ae()){await oe();const n=S.getAllAccounts()[0];if(!n)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await S.acquireTokenSilent({scopes:[V],account:n})).accessToken}return(await S.acquireTokenPopup(t)).accessToken}}async function Z(e=!1){const t=S.getAllAccounts();if(t.length===0)throw new Error("Not signed in");const n={scopes:J,account:t[0],forceRefresh:e};try{return(await S.acquireTokenSilent(n)).accessToken}catch{if(ae()){await oe();const o=S.getAllAccounts()[0];if(!o)throw new Error("เข้าสู่ระบบไม่สำเร็จ");return(await S.acquireTokenSilent({scopes:J,account:o})).accessToken}return(await S.acquireTokenPopup({scopes:J,account:t[0]})).accessToken}}async function we(e){const t=await Z(),n={subject:e.subject,start:{dateTime:e.start,timeZone:"Asia/Bangkok"},end:{dateTime:e.end,timeZone:"Asia/Bangkok"},body:e.body?{contentType:"HTML",content:e.body.replace(/\n/g,"<br>")}:void 0,attendees:e.attendees.filter(Boolean).map(o=>({emailAddress:{address:o},type:"required"})),isOnlineMeeting:e.isOnlineMeeting,onlineMeetingProvider:e.isOnlineMeeting?"teamsForBusiness":void 0},i=await fetch("https://graph.microsoft.com/v1.0/me/events",{method:"POST",headers:{Authorization:`Bearer ${t}`,"Content-Type":"application/json"},body:JSON.stringify(n)});if(!i.ok)throw new Error(`Calendar error ${i.status}: ${await i.text()}`)}async function ue(){try{const e=await _(),t=`${O}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$orderby=Title asc&$top=500`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const i=await n.json();a.projects=i.value.map(o=>({id:o.Id,Title:o.Title}))}}catch{}}async function pe(){try{const e=await _(),t=`${O}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText&$orderby=Title asc`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const i=await n.json();a.agents=i.value.map(o=>({email:o.EmailText,name:o.Title}))}}catch{}}async function fe(){try{const e=await _(),t=`${O}/_api/web/lists/getbytitle('HD_Tickets')/items?$select=Id,Title,TicketNumber,Status&$filter=Status ne 'Closed'&$orderby=Modified desc&$top=200`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const i=await n.json();a.tickets=i.value.map(o=>({id:o.Id,Title:o.Title,TicketNumber:o.TicketNumber,Status:o.Status}))}}catch{}}async function ge(){try{const e=await _(),t=`${O}/_api/web/lists/getbytitle('HD_Contracts')/items?$select=CustomerEmail&$top=2000`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});if(n.ok){const i=await n.json();a.contactEmails=i.value.map(o=>(o.CustomerEmail||"").trim().toLowerCase()).filter(Boolean)}}catch{}}async function ce(){const e=document.getElementById("btn-login-main"),t=document.getElementById("btn-login");e&&(e.disabled=!0,e.textContent="กำลังเข้าสู่ระบบ…"),t&&(t.disabled=!0);try{if(ae()){if(await oe(),a.account=S.getAllAccounts()[0]??null,!a.account)throw new Error("เข้าสู่ระบบไม่สำเร็จ")}else{const n=await S.loginPopup({scopes:[V]});a.account=n.account}await Promise.all([ue(),pe(),fe(),ge()]),M()}catch{e&&(e.disabled=!1,e.textContent="เข้าสู่ระบบ"),t&&(t.disabled=!1)}}async function ke(){a.account&&await S.logoutPopup({account:a.account}),a.account=null,M()}async function L(e,t){const n=await _(),i=`${O}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items`,o=await fetch(i,{method:"POST",headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify(t)});if(!o.ok){const d=await o.text();throw new Error(`SharePoint error ${o.status}: ${d}`)}return(await o.json()).Id}let q=null;const K="support@itservices.co.th";async function Ee(){if(q)return q;try{const e=await _(),t=`${O}/_api/web/lists/getbytitle('HD_EmailTemplates')/items?$select=EventKey,Subject,Body,IsEnabled&$top=50`,n=await fetch(t,{headers:{Authorization:`Bearer ${e}`,Accept:"application/json;odata=nometadata"}});return n.ok?(q=(await n.json()).value,q):[]}catch{return[]}}async function Te(){var e,t;try{const n=await _(),i=`${O}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title,Category&$filter=Category eq 'EmailConfig'&$top=1`,o=await fetch(i,{headers:{Authorization:`Bearer ${n}`,Accept:"application/json;odata=nometadata"}});return o.ok&&((t=(e=(await o.json()).value[0])==null?void 0:e.Title)==null?void 0:t.trim())||K}catch{return K}}function le(e,t){return e.replace(/\{\{(\w+)\}\}/g,(n,i)=>t[i]??`{{${i}}}`)}async function Ie(e,t,n,i=[]){try{const s=(await Ee()).find(c=>c.EventKey===e&&c.IsEnabled);if(!s)return;const d=le(s.Subject||"",t),b=le(s.Body||"",t);if(!d||!b)return;const w=c=>c.trim().toLowerCase(),m=[...new Map(n.filter(Boolean).map(c=>[w(c),c])).values()];if(m.length===0)return;const u=new Set(m.map(w)),T=[...new Map(i.filter(Boolean).map(c=>[w(c),c])).values()].filter(c=>!u.has(w(c))),j=await Te(),P=await Z(),r={subject:d,body:{contentType:"HTML",content:b},toRecipients:m.map(c=>({emailAddress:{address:c}}))};T.length&&(r.ccRecipients=T.map(c=>({emailAddress:{address:c}}))),j&&(r.from={emailAddress:{address:j}}),await fetch("https://graph.microsoft.com/v1.0/me/sendMail",{method:"POST",headers:{Authorization:`Bearer ${P}`,"Content-Type":"application/json"},body:JSON.stringify({message:r,saveToSentItems:!0})})}catch{}}async function U(e){var s;const t=d=>d.trim().toLowerCase(),n=t(((s=a.account)==null?void 0:s.username)??""),i=new Set,o=e.recipients.filter(Boolean).filter(d=>{const b=t(d);return!b||b===n||i.has(b)?!1:(i.add(b),!0)});if(o.length!==0)try{const d=await _(),b=`${O}/_api/web/lists/getbytitle('HD_Notifications')/items`;await Promise.all(o.map(w=>fetch(b,{method:"POST",headers:{Authorization:`Bearer ${d}`,Accept:"application/json;odata=nometadata","Content-Type":"application/json;odata=nometadata"},body:JSON.stringify({Title:e.title.slice(0,255),RecipientEmail:w,EventType:e.eventType,Message:e.message,LinkPath:e.linkPath,IsRead:!1})})))}catch{}}async function X(e,t){const n=document.querySelectorAll(".email-att-cb:checked");if(n.length===0)return;const i=await _();for(const o of Array.from(n)){const s=o.dataset.attId,d=o.dataset.attName,b=await new Promise((r,c)=>{Office.context.mailbox.item.getAttachmentContentAsync(s,{},f=>{f.status===Office.AsyncResultStatus.Succeeded?r(f):c(new Error(f.error.message))})}),{content:w,format:m}=b.value;let u;if(m===Office.MailboxEnums.AttachmentContentFormat.Base64){const r=atob(w);u=new Uint8Array(r.length);for(let c=0;c<r.length;c++)u[c]=r.charCodeAt(c)}else if(m===Office.MailboxEnums.AttachmentContentFormat.Eml||m===Office.MailboxEnums.AttachmentContentFormat.ICalendar)u=new TextEncoder().encode(w);else continue;const T=encodeURIComponent(d),j=`${O}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${T}')`;if(!(await fetch(j,{method:"POST",headers:{Authorization:`Bearer ${i}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:u.buffer})).ok)throw new Error(`Upload ${d} failed`)}}async function Ce(e){const t=`https://graph.microsoft.com/v1.0/me/messages/${e}/$value`;let n=await Z(),i=await fetch(t,{headers:{Authorization:`Bearer ${n}`}});if((i.status===401||i.status===403)&&(n=await Z(!0),i=await fetch(t,{headers:{Authorization:`Bearer ${n}`}})),!i.ok)throw new Error(`Graph ${i.status}`);return i.arrayBuffer()}async function Ae(e){const t=await new Promise((i,o)=>{Office.context.mailbox.getCallbackTokenAsync({isRest:!0},s=>{s.status===Office.AsyncResultStatus.Succeeded?i(s.value):o(new Error("callback token failed"))})}),n=await fetch(`${Office.context.mailbox.restUrl}/v2.0/me/messages/${e}/$value`,{headers:{Authorization:`Bearer ${t}`}});if(!n.ok)throw new Error(`REST ${n.status}`);return n.arrayBuffer()}async function ee(e,t){const n=document.getElementById("f-attach-eml");if(!(n!=null&&n.checked))return;const i=Office.context.mailbox.item;if(!i)return;const o=Office.context.mailbox.convertToRestId(i.itemId,Office.MailboxEnums.RestVersion.v2_0);let s,d="",b="";try{s=await Ce(o)}catch(j){d=j instanceof Error?j.message:String(j);try{s=await Ae(o)}catch(P){b=P instanceof Error?P.message:String(P),console.error("[eml] graph:",d,"| callback:",b),k(`ดึง .eml ไม่ได้ (Graph: ${d} / REST: ${b}) — ไฟล์อื่นบันทึกแล้ว`,"error");return}}const w=(i.subject||"email").replace(/[\\/:*?"<>|#%&{}~]/g,"_").slice(0,100).trim()||"email",m=await _(),u=`${O}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${encodeURIComponent(w+".eml")}')`;(await fetch(u,{method:"POST",headers:{Authorization:`Bearer ${m}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:s})).ok||k("บันทึกไฟล์ .eml ไม่สำเร็จ","error")}async function te(e,t,n){const i=await _();for(const o of n){const s=await o.arrayBuffer(),d=encodeURIComponent(o.name),b=`${O}/_api/web/lists/getbytitle('${encodeURIComponent(e)}')/items(${t})/AttachmentFiles/add(FileName='${d}')`;if(!(await fetch(b,{method:"POST",headers:{Authorization:`Bearer ${i}`,Accept:"application/json;odata=nometadata","Content-Type":"application/octet-stream"},body:s})).ok)throw new Error(`Upload ${o.name} failed`)}}function k(e,t="success"){const n=document.getElementById("toast-container");if(!n)return;const i=t==="success"?"bg-green-500":"bg-red-500",o=t==="success"?"✅":"❌",s=document.createElement("div");s.className=`toast pointer-events-auto ${i} text-white text-sm font-medium px-4 py-3 rounded-lg shadow-lg max-w-xs mx-2`,s.textContent=`${o} ${e}`,n.appendChild(s),setTimeout(()=>s.remove(),4e3)}function Be(e){const t=e.split(`
`).map(u=>u.trim()).filter(Boolean),n=/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/,i=/(\+?[\d\s()\-]{7,}(?:ext\.?\s*\d+)?)/i,o=/\b(co\.?,?\s*ltd\.?|co\.?,?\s*limited|corporation|corp\.?|บริษัท|จำกัด|holding|group|inc\.?|llc)\b/i;let s="",d="",b="";const w=[];for(const u of t)if(!/^[-_=*]{2,}$/.test(u)&&!/^(best regards|regards|sincerely|ขอแสดงความนับถือ|ด้วยความนับถือ|from|sent|thanks|thank you)/i.test(u)){if(!s){const T=u.match(n);if(T){s=T[0];continue}}if(!d){const T=u.match(i);if(T&&T[0].replace(/\D/g,"").length>=7){d=T[0].trim();continue}}if(!b&&o.test(u)){b=u;continue}u.length>=2&&u.length<=50&&!/\d{4,}/.test(u)&&w.push(u)}const m=w.find(u=>!n.test(u)&&!o.test(u))??"";return!s&&!m?null:{name:m,company:b,email:s,phone:d}}async function Se(){const e=a.signatureContact;if(!e)return;const t=(a.emailSenderEmail||"").toLowerCase();if(t&&a.contactEmails.includes(t)){k("ลูกค้านี้มีในระบบแล้ว","success"),a.signatureContact=null,M();return}const n=document.getElementById("btn-import-customer");n&&(n.disabled=!0,n.textContent="กำลังบันทึก…");try{await L("HD_Contracts",{Title:a.emailSenderName||e.name,CustomerEmail:a.emailSenderEmail,Phone:e.phone||void 0,Company:e.company||void 0,Status:"Active"}),t&&a.contactEmails.push(t),k("เพิ่มลูกค้าสำเร็จ!"),a.signatureContact=null,M()}catch(i){const o=i instanceof Error?i.message:String(i);k(`เกิดข้อผิดพลาด: ${o}`,"error"),n&&(n.disabled=!1,n.textContent="เพิ่มเป็นลูกค้า")}}function be(){return new Date().toISOString().split("T")[0]}function je(){const e=new Date;return`HD-${`${e.getFullYear()}${String(e.getMonth()+1).padStart(2,"0")}${String(e.getDate()).padStart(2,"0")}`}-${Math.floor(Math.random()*900+100)}`}function re(){var e;return a.droppedFiles.length>0||document.querySelectorAll(".email-att-cb:checked").length>0||(((e=document.getElementById("f-attach-eml"))==null?void 0:e.checked)??!1)}async function G(e,t){a.droppedFiles.length>0&&await te(e,t,a.droppedFiles),await X(e,t),await ee(e,t)}let W=!1;async function Pe(){var t,n,i,o,s,d,b,w,m,u,T,j,P;if(!a.account){k("กรุณาเข้าสู่ระบบก่อน","error");return}if(W)return;W=!0;const e=document.getElementById("submit-btn");e&&(e.disabled=!0,e.textContent="กำลังบันทึก…");try{if(a.tab==="ticket"){const r=document.getElementById("f-title").value.trim(),c=document.getElementById("f-description").value.trim(),f=document.getElementById("f-priority").value,x=document.getElementById("f-customer-email").value.trim(),$=((t=document.getElementById("f-cc-enable"))==null?void 0:t.checked)??!0?(((n=document.getElementById("f-cc"))==null?void 0:n.value)||"").split(/[,;\s]+/).map(l=>l.trim()).filter(Boolean):[],g=document.getElementById("f-assigned-email").value,y=a.agents.find(l=>l.email===g),C=je(),A=await L("HD_Tickets",{Title:r,TicketNumber:C,Description:c,Priority:f,CustomerEmail:x,CustomerName:a.emailSenderName||x,Status:"Open",AssignedEmail:g||void 0,AssignedToName:(y==null?void 0:y.name)??((i=a.account)==null?void 0:i.name)??""});if(re()){const l=await L("HD_TicketComments",{Title:"📎 ไฟล์แนบจากอีเมล",TicketID:A,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await G("HD_TicketComments",l)}a.droppedFiles=[],await Ie("ticket_created",{ticket_number:C,ticket_title:r,priority:f,category:"-",description:(c||"-").replace(/\n/g,"<br>"),customer_name:a.emailSenderName||x,assigned_name:(y==null?void 0:y.name)??((o=a.account)==null?void 0:o.name)??"-",link:"https://itservices.co.th/helpdesk/"},[x],[g,a.account.username,...$]),k("สร้าง Ticket สำเร็จ!")}else if(a.tab==="task"){const r=document.getElementById("f-title").value.trim(),c=parseInt(((s=document.getElementById("f-project"))==null?void 0:s.value)||"0"),f=document.getElementById("f-due-date").value,x=document.getElementById("f-note").value.trim(),I=document.getElementById("f-assigned-email").value,$=a.agents.find(C=>C.email===I);if(!c){k("กรุณาเลือก Project","error");return}const g=await L("PM_Tasks",{Title:r,DueDate:f||null,TaskNote:x,AssignedTo:($==null?void 0:$.name)??a.account.name??a.account.username,AssignedEmail:I,IsCompleted:!1,IsAcknowledged:!1,ProjectID:c});if(a.droppedFiles.length>0&&await te("PM_Tasks",g,a.droppedFiles),await X("PM_Tasks",g),await ee("PM_Tasks",g),a.droppedFiles=[],await U({recipients:[I],title:`📋 ได้รับมอบหมาย Task: ${r}`,message:x||(f?`กำหนดส่ง ${f}`:"มี Task ใหม่"),linkPath:c?`/projects/${c}`:"/my-work",eventType:"task_assigned"}),((d=document.getElementById("f-teams"))==null?void 0:d.checked)&&f){const C=Array.from(document.querySelectorAll(".att-internal:checked")).map(E=>E.value),A=(((b=document.getElementById("f-ext-att"))==null?void 0:b.value)||"").split(/[,;\s]+/).map(E=>E.trim()).filter(Boolean),l=`${f}T09:00:00`,B=`${f}T10:00:00`;try{await we({subject:r,start:l,end:B,body:x,attendees:[...C,...A],isOnlineMeeting:!0}),k("สร้าง Task + นัดประชุม Teams สำเร็จ!")}catch(E){k("สร้าง Task แล้ว แต่สร้างนัดประชุมไม่สำเร็จ: "+(E instanceof Error?E.message:""),"error")}}else k("สร้าง Task สำเร็จ!")}else if(a.tab==="incident"){const r=document.getElementById("f-title").value.trim(),c=parseInt(((w=document.getElementById("f-project"))==null?void 0:w.value)||"0"),f=document.getElementById("f-description").value.trim(),x=document.getElementById("f-severity").value,I=document.getElementById("f-assigned-email").value,$=a.agents.find(l=>l.email===I),g=document.getElementById("f-status").value,y=document.getElementById("f-incident-date").value,C=document.getElementById("f-resolution").value.trim();if(!c){k("กรุณาเลือก Project","error");return}const A=await L("PM_Incidents",{Title:r,Description:f||void 0,Severity:x,Status:g,AssignedTo:($==null?void 0:$.name)??a.account.name??a.account.username,AssignedEmail:I,ProjectID:c,IncidentDate:y||be(),Resolution:C||void 0});a.droppedFiles.length>0&&await te("PM_Incidents",A,a.droppedFiles),await X("PM_Incidents",A),await ee("PM_Incidents",A),a.droppedFiles=[],await U({recipients:[I],title:`🚨 ได้รับมอบหมาย Incident: ${r}`,message:`ความรุนแรง ${x}${f?" — "+f.slice(0,120):""}`,linkPath:c?`/projects/${c}`:"/my-work",eventType:"incident_created"}),k("สร้าง Incident สำเร็จ!")}else if(a.tab==="comment"){const r=parseInt(((m=document.getElementById("f-ticket"))==null?void 0:m.value)||"0"),c=document.getElementById("f-comment").value.trim(),f=document.getElementById("f-comment-type").value;if(!r){k("กรุณาเลือก Ticket","error");return}if(!c){k("กรุณาพิมพ์ Comment","error");return}const x=await L("HD_TicketComments",{Title:c.slice(0,100),TicketID:r,CommentText:c,CommentType:f,CommentDate:new Date().toISOString()});await G("HD_TicketComments",x),a.droppedFiles=[];try{const I=await _(),$=`${O}/_api/web/lists/getbytitle('HD_Tickets')/items(${r})?$select=TicketNumber,Title,AssignedEmail,Author/EMail&$expand=Author`,g=await fetch($,{headers:{Authorization:`Bearer ${I}`,Accept:"application/json;odata=nometadata"}});if(g.ok){const y=await g.json(),C=a.account.username.toLowerCase(),A=[...new Set([y.AssignedEmail,(u=y.Author)==null?void 0:u.EMail].filter(Boolean))].filter(l=>l.toLowerCase()!==C);A.length&&await U({recipients:A,title:`💬 ${((T=a.account)==null?void 0:T.name)??"มีคน"} คอมเมนต์ใน ${y.TicketNumber||"#"+r}`,message:c.slice(0,200),linkPath:`/tickets/${r}`,eventType:"comment_added"})}}catch{}k("เพิ่ม Comment สำเร็จ!")}else if(a.tab==="project"){const r=document.getElementById("f-title").value.trim(),c=document.getElementById("f-company").value.trim(),f=document.getElementById("f-group").value,x=document.getElementById("f-status").value,I=document.getElementById("f-start").value,$=document.getElementById("f-end").value,g=document.getElementById("f-description").value.trim();if(!r){k("กรุณาใส่ชื่อโครงการ","error");return}const y=await L("PM_Projects",{Title:r,Company:c||void 0,ProjectGroup:f,Progress:0,StartDate:I||void 0,EndDate:$||null,Status:x,CreatedByEmail:a.account.username,Comment:g||void 0});if(re()){const C=await L("PM_Comments",{Title:"📎 ไฟล์แนบจากอีเมล",ProjectID:y,CommentText:"ไฟล์แนบจาก Outlook Add-in",CommentType:"Internal",CommentDate:new Date().toISOString(),ParentID:0});await G("PM_Comments",C)}a.droppedFiles=[],k("สร้างโครงการสำเร็จ!")}else if(a.tab==="projcomment"){const r=parseInt(((j=document.getElementById("f-project"))==null?void 0:j.value)||"0"),c=document.getElementById("f-comment").value.trim(),f=document.getElementById("f-comment-type").value;if(!r){k("กรุณาเลือกโครงการ","error");return}if(!c){k("กรุณาพิมพ์ Comment","error");return}const x=await L("PM_Comments",{Title:c.slice(0,100),ProjectID:r,CommentText:c,CommentType:f,CommentDate:new Date().toISOString(),ParentID:0});await G("PM_Comments",x),a.droppedFiles=[];try{const I=await _(),$=`${O}/_api/web/lists/getbytitle('PM_Projects')/items(${r})?$select=Title,CreatedByEmail`,g=await fetch($,{headers:{Authorization:`Bearer ${I}`,Accept:"application/json;odata=nometadata"}});if(g.ok){const y=await g.json(),C=a.account.username.toLowerCase();y.CreatedByEmail&&y.CreatedByEmail.toLowerCase()!==C&&await U({recipients:[y.CreatedByEmail],title:`💬 ${((P=a.account)==null?void 0:P.name)??"มีคน"} คอมเมนต์ในโครงการ ${y.Title??""}`,message:c.slice(0,200),linkPath:`/projects/${r}?tab=comments`,eventType:"comment_added"})}}catch{}k("เพิ่ม Comment สำเร็จ!")}}catch(r){const c=r instanceof Error?r.message:String(r);k(`เกิดข้อผิดพลาด: ${c}`,"error")}finally{W=!1,e&&(e.disabled=!1,e.textContent="บันทึก")}}const De={ticket:{label:"Ticket",icon:"🎫"},comment:{label:"Comment",icon:"💬"},project:{label:"Project",icon:"📁"},task:{label:"Task",icon:"✅"},incident:{label:"Incident",icon:"🚨"},projcomment:{label:"Comment",icon:"💬"}},Oe=[{title:"🎫 Helpdesk",tabs:["ticket","comment"]},{title:"📁 Project",tabs:["project","task","incident","projcomment"]}],he=["f-title","f-description","f-priority","f-customer-email","f-cc","f-assigned-email","f-project","f-due-date","f-note","f-severity","f-status","f-incident-date","f-resolution","f-ticket","f-comment","f-comment-type","f-company","f-group","f-start","f-end","f-ext-att","f-attach-eml"];let R={};function _e(){for(const t of he){const n=document.getElementById(t);n&&(R[t]=n.value)}const e=document.getElementById("f-teams");e&&(R["f-teams"]=e.checked)}function Me(){for(const t of he){const n=document.getElementById(t);n&&R[t]!==void 0&&R[t]!==""&&(n.value=R[t])}const e=document.getElementById("f-teams");if(e&&R["f-teams"]!==void 0){e.checked=R["f-teams"];const t=document.getElementById("teams-fields");t&&(t.style.display=e.checked?"block":"none")}}function M(){var I,$,g,y,C,A;const e=document.getElementById("app");if(!e)return;_e();const{account:t,tab:n,emailSubject:i,emailSenderName:o,emailSenderEmail:s,emailBodyPreview:d}=a,b=t!==null,w=`
    <div class="bg-gradient-to-r from-blue-700 to-blue-600 text-white px-4 py-2.5 flex items-center gap-2.5 shadow flex-shrink-0">
      <div class="w-7 h-7 bg-white/95 rounded-lg flex items-center justify-center flex-shrink-0">
        <span class="text-blue-700 font-extrabold text-xs">iT</span>
      </div>
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm leading-tight truncate">iT Services Helpdesk</div>
        ${b?`<div class="text-[10px] text-blue-100 truncate">${p((t==null?void 0:t.name)??(t==null?void 0:t.username)??"")}</div>`:""}
      </div>
      ${b?`<button id="btn-logout" title="ออกจากระบบ" class="p-1.5 rounded-lg hover:bg-white/15 transition flex-shrink-0">
             <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
           </button>`:`<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded-lg transition flex-shrink-0">
             เข้าสู่ระบบ
           </button>`}
    </div>
  `;if(!b){e.innerHTML=`
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
    `,(I=document.getElementById("btn-login"))==null||I.addEventListener("click",ce),($=document.getElementById("btn-login-main"))==null||$.addEventListener("click",ce);return}const m=i?`<div class="mx-3 mt-3 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5 text-xs text-slate-700 space-y-1">
         <div class="font-semibold text-blue-800 truncate" title="${p(i)}">📧 ${p(i)}</div>
         ${o?`<div class="text-slate-500">จาก: <span class="font-medium text-slate-700">${p(o)}</span></div>`:""}
         ${s&&s!==o?`<div class="text-slate-400 truncate">${p(s)}</div>`:""}
       </div>`:`<div class="mx-3 mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-700">
         ⚠️ ไม่พบข้อมูล Email (โหมดทดสอบ)
       </div>`,u=a.signatureContact,T=!!s&&a.contactEmails.includes(s.toLowerCase()),j=u?`<div class="mx-3 mt-3 bg-orange-50 border border-orange-200 rounded-lg px-3 py-2.5 text-xs text-slate-700">
         <div class="flex items-center justify-between mb-2">
           <span class="font-semibold text-orange-700">👤 ข้อมูลผู้ส่ง (จาก Signature)</span>
         </div>
         <div class="space-y-0.5 mb-2.5">
           ${o?`<div><span class="text-slate-400">ชื่อ:</span> <span class="font-medium">${p(o)}</span></div>`:""}
           ${u.company?`<div><span class="text-slate-400">บริษัท:</span> ${p(u.company)}</div>`:""}
           ${s?`<div><span class="text-slate-400">Email:</span> ${p(s)}</div>`:""}
           ${u.phone?`<div><span class="text-slate-400">โทร:</span> ${p(u.phone)}</div>`:""}
         </div>
         ${T?'<div class="w-full bg-green-100 text-green-700 text-xs font-semibold py-1.5 rounded-md text-center">✓ ลูกค้านี้มีในระบบแล้ว</div>':`<button id="btn-import-customer"
                class="w-full bg-orange-500 hover:bg-orange-400 text-white text-xs font-semibold py-1.5 rounded-md transition">
                + เพิ่มเป็นลูกค้า
              </button>`}
       </div>`:"",P=`
    <div class="mx-3 mt-3 space-y-2">
      ${Oe.map(l=>`
        <div>
          <div class="text-[10px] font-semibold text-slate-400 uppercase tracking-wide mb-1 px-0.5">${l.title}</div>
          <div class="grid grid-cols-4 gap-1">
            ${l.tabs.map(B=>{const E=De[B];return`<button data-tab="${B}"
                class="tab-btn flex flex-col items-center gap-1 py-2 rounded-lg transition ${n===B?"bg-blue-700 text-white shadow":"text-slate-500 hover:bg-slate-100"}">
                <span class="text-base leading-none">${E.icon}</span>
                <span class="text-[9px] font-medium leading-none">${E.label}</span>
              </button>`}).join("")}
          </div>
        </div>
      `).join("")}
    </div>
  `;let r="";n==="ticket"?r=`
      ${v("Title / หัวข้อ",`<input id="f-title" type="text"
        class="${h}"
        value="${p(i)}" />`)}
      ${v("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${h} resize-none">${p(d)}</textarea>`)}
      ${v("Priority",`<select id="f-priority" class="${h}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
      ${v("Customer Email",`<input id="f-customer-email" type="email"
        class="${h}"
        value="${p(s)}" />`)}
      ${v("CC — ให้ผู้ที่อยู่ในเมลนี้รับรู้",`
        <label class="flex items-center gap-2 text-xs text-slate-600 mb-1.5 cursor-pointer">
          <input id="f-cc-enable" type="checkbox" ${a.emailCc.length?"checked":""} /> แนบผู้รับในเมลนี้เป็น CC อัตโนมัติ
        </label>
        <input id="f-cc" type="text" class="${h}" value="${p(a.emailCc.join(", "))}" placeholder="someone@company.com, boss@company.com" />`)}
      ${v("Assign ให้ Agent",Y(t.username))}
      ${F()}
    `:n==="task"?r=`
      ${v("ชื่อ Task *",`<input id="f-title" type="text" required
        class="${h}" value="${p(i)}" />`)}
      ${v("Project *",Q())}
      ${v("Assign ให้",Y(t.username))}
      ${v("Due Date",`<input id="f-due-date" type="date" class="${h}" />`)}
      ${v("Task Note",`<textarea id="f-note" rows="4"
        class="${h} resize-y">${p(d)}</textarea>`)}
      <label class="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer mb-1">
        <input id="f-teams" type="checkbox" class="rounded" onchange="document.getElementById('teams-fields').style.display=this.checked?'block':'none'" />
        💻 เพิ่มการประชุมออนไลน์ (Teams) — ใช้เวลา 09:00–10:00 ของวัน Due Date
      </label>
      <div id="teams-fields" style="display:none" class="space-y-2 mb-2">
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">ผู้เข้าร่วม Internal</label>
          <div class="max-h-28 overflow-y-auto border border-slate-200 rounded-md p-1.5 space-y-0.5">
            ${a.agents.map(l=>`<label class="flex items-center gap-2 text-xs text-slate-700 px-1 py-0.5 hover:bg-slate-50 rounded cursor-pointer">
              <input type="checkbox" class="att-internal" value="${p(l.email)}" /> ${p(l.name)}
            </label>`).join("")}
          </div>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">Email ภายนอก (คั่นด้วย ,)</label>
          <input id="f-ext-att" type="text" class="${h}" placeholder="someone@company.com, ..." />
        </div>
      </div>
      ${F()}
    `:n==="incident"?r=`
      ${v("ชื่อ Incident *",`<input id="f-title" type="text" required
        class="${h}" value="${p(i)}" />`)}
      ${v("Project *",Q())}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">ความรุนแรง</label>
          <select id="f-severity" class="${h}">
            <option value="Low">Low</option>
            <option value="Medium" selected>Medium</option>
            <option value="High">High</option>
            <option value="Critical">Critical</option>
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${h}">
            <option value="Open" selected>Open</option>
            <option value="In Progress">In Progress</option>
            <option value="Resolved">Resolved</option>
          </select>
        </div>
      </div>
      ${v("Assign ให้ Agent",Y(t.username))}
      ${v("วันที่เกิด Incident",`<input id="f-incident-date" type="date" class="${h}" value="${be()}" />`)}
      ${v("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${h} resize-y">${p(d)}</textarea>`)}
      ${v("วิธีแก้ไข (ถ้ามี)",`<textarea id="f-resolution" rows="2"
        class="${h} resize-y" placeholder="อธิบายวิธีแก้ไขปัญหา..."></textarea>`)}
      ${F()}
    `:n==="comment"?r=`
      ${v("เลือก Ticket *",`<select id="f-ticket" class="${h}">
        <option value="">-- เลือก Ticket ที่จะเพิ่ม Comment --</option>
        ${a.tickets.map(l=>`<option value="${l.id}">${p(l.TicketNumber||"#"+l.id)} · ${p(l.Title)}</option>`).join("")}
      </select>`)}
      ${v("ประเภท",`<select id="f-comment-type" class="${h}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${v("Comment *",`<textarea id="f-comment" rows="5"
        class="${h} resize-y" placeholder="พิมพ์ comment...">${p(d)}</textarea>`)}
      ${F()}
    `:n==="project"?r=`
      ${v("ชื่อโครงการ *",`<input id="f-title" type="text" required
        class="${h}" value="${p(i)}" />`)}
      ${v("บริษัท / ลูกค้า",`<input id="f-company" type="text" class="${h}" value="${p(((g=a.signatureContact)==null?void 0:g.company)??"")}" />`)}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">กลุ่มโครงการ</label>
          <select id="f-group" class="${h}">
            ${["Internal","External","R&D","Maintenance","อื่นๆ"].map(l=>`<option>${l}</option>`).join("")}
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${h}">
            ${["Planning","Active","On Hold","Completed","Cancelled"].map(l=>`<option>${l}</option>`).join("")}
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันที่เริ่ม</label>
          <input id="f-start" type="date" class="${h}" /></div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันสิ้นสุด</label>
          <input id="f-end" type="date" class="${h}" /></div>
      </div>
      ${v("รายละเอียด",`<textarea id="f-description" rows="4"
        class="${h} resize-y">${p(d)}</textarea>`)}
      ${F()}
    `:n==="projcomment"&&(r=`
      ${v("เลือกโครงการ *",Q())}
      ${v("ประเภท",`<select id="f-comment-type" class="${h}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${v("Comment *",`<textarea id="f-comment" rows="5"
        class="${h} resize-y" placeholder="พิมพ์ comment...">${p(d)}</textarea>`)}
      ${F()}
    `);const c=n==="comment"||n==="projcomment"?"เพิ่ม Comment":n==="project"?"สร้างโครงการ":n==="incident"?"แจ้ง Incident":n==="task"?"สร้าง Task":"สร้าง Ticket";e.innerHTML=`
    <div class="flex flex-col h-screen bg-slate-50">
      ${w}
      <div class="flex-1 overflow-y-auto">
        ${m}
        ${j}
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
  `,(y=document.getElementById("btn-logout"))==null||y.addEventListener("click",ke),(C=document.getElementById("submit-btn"))==null||C.addEventListener("click",Pe),(A=document.getElementById("btn-import-customer"))==null||A.addEventListener("click",Se),document.querySelectorAll(".tab-btn").forEach(l=>{l.addEventListener("click",()=>{const B=l.dataset.tab;B&&B!==a.tab&&(a.tab=B,M())})});const f=document.getElementById("drop-zone"),x=document.getElementById("f-files");f&&x&&(x.addEventListener("change",()=>{x.files&&ne(Array.from(x.files)),x.value=""}),f.addEventListener("dragover",l=>{l.preventDefault(),f.classList.add("border-blue-500","bg-blue-50")}),f.addEventListener("dragleave",()=>{f.classList.remove("border-blue-500","bg-blue-50")}),f.addEventListener("drop",l=>{var E;l.preventDefault(),f.classList.remove("border-blue-500","bg-blue-50");const B=Array.from(((E=l.dataTransfer)==null?void 0:E.files)??[]);B.length&&ne(B)})),document.querySelectorAll(".remove-dropped").forEach(l=>{l.addEventListener("click",()=>{const B=parseInt(l.dataset.remove??"-1");B>=0&&(a.droppedFiles.splice(B,1),M())})}),Me()}function ne(e){a.droppedFiles.push(...e),M()}document.addEventListener("paste",e=>{var i;if(!a.account)return;const t=Array.from(((i=e.clipboardData)==null?void 0:i.items)??[]),n=[];for(const o of t)if(o.kind==="file"){const s=o.getAsFile();if(s){const d=s.name&&s.name!=="image.png"?s.name:`screenshot-${new Date().toISOString().replace(/[:.]/g,"-").slice(0,19)}.png`;n.push(new File([s],d,{type:s.type}))}}n.length&&(e.preventDefault(),ne(n),k(`แนบไฟล์แล้ว: ${n.map(o=>o.name).join(", ")}`))});const h="w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white";function de(e){return e<1024?`${e} B`:e<1024*1024?`${(e/1024).toFixed(0)} KB`:`${(e/1024/1024).toFixed(1)} MB`}function F(){const e=a.emailAttachments,t=a.droppedFiles,n=e.length>0?`<div class="mb-2 space-y-1">
        <p class="text-xs text-slate-500">📎 ไฟล์แนบจาก Email:</p>
        ${e.map(o=>`
          <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
            <input type="checkbox" class="email-att-cb" data-att-id="${p(o.id)}" data-att-name="${p(o.name)}" data-att-item="${o.isItem?"1":"0"}" checked />
            <span class="flex-1 truncate">${o.isItem?"📧 ":""}${p(o.name)}</span>
            <span class="text-slate-400 flex-shrink-0">${de(o.size)}</span>
          </label>`).join("")}
      </div>`:"",i=t.length>0?`<div class="mt-2 space-y-1">
        ${t.map((o,s)=>`<div class="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 rounded px-2 py-1">
            <span class="text-base">${o.type.startsWith("image/")?"🖼️":"📄"}</span>
            <span class="flex-1 truncate">${p(o.name)}</span>
            <span class="text-slate-400">${de(o.size)}</span>
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
    ${i}
  </div>`}function Y(e){var t;return`<select id="f-assigned-email" class="${h}">
    <option value="${p(e)}">${p(((t=a.account)==null?void 0:t.name)??e)} (ฉัน)</option>
    ${a.agents.filter(n=>n.email!==e).map(n=>`<option value="${p(n.email)}">${p(n.name)}</option>`).join("")}
  </select>`}function Q(){return a.projects.length===0?'<div class="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-2">⚠️ ไม่พบ Project ที่ Active</div>':`<select id="f-project" class="${h}">
    <option value="">-- เลือก Project --</option>
    ${a.projects.map(e=>`<option value="${e.id}">${p(e.Title)}</option>`).join("")}
  </select>`}function v(e,t){return`
    <div class="space-y-1">
      <label class="block text-xs font-medium text-slate-600">${e}</label>
      ${t}
    </div>
  `}function p(e){return e.replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}async function Le(){await S.initialize(),await S.handleRedirectPromise();const e=S.getAllAccounts();if(e.length>0){a.account=e[0];try{await S.acquireTokenSilent({scopes:[V],account:e[0]}),await Promise.all([ue(),pe(),fe(),ge()])}catch{a.account=null}}typeof Office<"u"?Office.onReady(t=>{var n;if(t.host===Office.HostType.Outlook){const i=Office.context.mailbox.item;if(i){a.emailSubject=i.subject??"";const o=i.from;o&&(a.emailSenderName=o.displayName??"",a.emailSenderEmail=o.emailAddress??"");const s=(((n=a.account)==null?void 0:n.username)??"").toLowerCase(),d=((o==null?void 0:o.emailAddress)??"").toLowerCase(),b=[...i.to??[],...i.cc??[]].map(m=>m.emailAddress).filter(Boolean);a.emailCc=[...new Set(b.map(m=>m.toLowerCase()))].filter(m=>m!==s&&m!==d);const w=i.attachments??[];a.emailAttachments=w.filter(m=>m.attachmentType===Office.MailboxEnums.AttachmentType.File||m.attachmentType===Office.MailboxEnums.AttachmentType.Item).map(m=>({id:m.id,name:m.attachmentType===Office.MailboxEnums.AttachmentType.Item?`${(m.name||"email").replace(/\.eml$/i,"")}.eml`:m.name,size:m.size,isItem:m.attachmentType===Office.MailboxEnums.AttachmentType.Item})),i.body.getAsync(Office.CoercionType.Html,{asyncContext:{}},m=>{if(m.status===Office.AsyncResultStatus.Succeeded){let u=function(l,B=!1){if(l.nodeType===3){const D=l.textContent??"";return B&&D.trim()===""?"":D}const E=l,H=(E.tagName??"").toLowerCase();if(P.includes(H))return"";if(H==="br")return" ";if(H==="tr"){const D=[];for(let N=0;N<E.childNodes.length;N++){const ie=E.childNodes[N],se=(ie.tagName??"").toLowerCase();(se==="td"||se==="th")&&D.push((ie.textContent??"").replace(/\s+/g," ").trim())}return D.length?D.join("	")+`
`:""}if(c.includes(H)){let D="";for(let N=0;N<E.childNodes.length;N++)D+=u(E.childNodes[N],!0);return D}let z="";for(let D=0;D<E.childNodes.length;D++)z+=u(E.childNodes[D],!1);return r.includes(H)&&(z=`
`+z.trim()+`
`),z};const T=m.value,j=new DOMParser().parseFromString(T,"text/html"),P=["style","script","head","img","meta","link","noscript"],r=["p","div","li","h1","h2","h3","h4","h5","h6","blockquote"],c=["table","thead","tbody","tfoot"],I=u(j.body??j.documentElement).replace(/[ \t]{2,}/g," ").replace(/\n[ \t]+/g,`
`).replace(/\n{3,}/g,`

`).trim().split(`
`),$=[];let g="";for(const l of I)l.trim()===""?g&&($.push(g.trim()),g=""):l.includes("	")?(g&&($.push(g.trim()),g=""),$.push(l)):g=g?g+" "+l.trim():l.trim();g&&$.push(g.trim());const y=$.join(`
`),C=/\n([-_]{3,}|From:\s|Best regards|Regards,|ขอแสดงความนับถือ|Sent:\s)/i,A=y.search(C);if(A>80){a.emailBodyPreview=y.slice(0,A).trim().slice(0,2e3);const l=y.slice(A).trim();a.signatureContact=Be(l)}else a.emailBodyPreview=y.trim().slice(0,2e3),a.signatureContact=null}M()});return}}me(),M()}):(me(),M())}function me(){a.emailSubject="[DEV] Test Email Subject",a.emailSenderName="Test Sender",a.emailSenderEmail="test@example.com",a.emailBodyPreview="This is a placeholder email body for development mode."}Le().catch(e=>{console.error("Init error:",e);const t=document.getElementById("app");t&&(t.innerHTML=`<div class="p-4 text-red-600 text-sm">เกิดข้อผิดพลาด: ${String(e)}</div>`)});
