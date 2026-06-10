import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
} from '@azure/msal-browser'

// ─── Config ───────────────────────────────────────────────────────────────────
const CLIENT_ID = '0bab07cf-65e6-487c-89af-c917fc1a5a13'
const TENANT_ID = 'd569b991-89fc-4a62-9df5-eb361abcef40'
const SHAREPOINT_URL = 'https://rpaexpert.sharepoint.com/sites/iTServicesCo.Ltd'
const SP_SCOPE = 'https://rpaexpert.sharepoint.com/.default'
const GRAPH_SCOPES = ['https://graph.microsoft.com/Calendars.ReadWrite', 'https://graph.microsoft.com/Mail.Send']

// ─── MSAL setup ───────────────────────────────────────────────────────────────
const msalInstance = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin.includes('localhost')
      ? 'http://localhost:3000/'
      : 'https://darmmunginsa.github.io/itservices-addin/',
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
})

// ─── State ────────────────────────────────────────────────────────────────────
type Tab = 'ticket' | 'task' | 'incident' | 'comment' | 'project'

interface Project { id: number; Title: string }
interface Agent { email: string; name: string }
interface TicketRef { id: number; Title: string; TicketNumber: string; Status: string }

interface SignatureContact {
  name: string
  company: string
  email: string
  phone: string
}

interface AppState {
  account: AccountInfo | null
  tab: Tab
  emailSubject: string
  emailBodyPreview: string
  emailSenderName: string
  emailSenderEmail: string
  loading: boolean
  projects: Project[]
  agents: Agent[]
  emailAttachments: { id: string; name: string; size: number }[]
  signatureContact: SignatureContact | null
  droppedFiles: File[]
  tickets: TicketRef[]
  contactEmails: string[]    // อีเมลลูกค้าที่มีในระบบแล้ว (lowercase) — กันเพิ่มซ้ำ
  emailCc: string[]          // ผู้รับในเมลต้นทาง (To+CC) — ใช้ prefill ช่อง CC
}

const state: AppState = {
  account: null,
  tab: 'ticket',
  emailSubject: '',
  emailBodyPreview: '',
  emailSenderName: '',
  emailSenderEmail: '',
  loading: false,
  projects: [],
  agents: [],
  emailAttachments: [],
  signatureContact: null,
  droppedFiles: [],
  tickets: [],
  contactEmails: [],
  emailCc: [],
}

// ─── MSAL helpers ─────────────────────────────────────────────────────────────
async function getToken(): Promise<string> {
  const accounts = msalInstance.getAllAccounts()
  if (accounts.length === 0) throw new Error('Not signed in')

  const request = { scopes: [SP_SCOPE], account: accounts[0] }
  let result: AuthenticationResult

  try {
    result = await msalInstance.acquireTokenSilent(request)
  } catch {
    result = await msalInstance.acquireTokenPopup(request)
  }
  return result.accessToken
}

// Graph token (Calendar)
async function getGraphToken(): Promise<string> {
  const accounts = msalInstance.getAllAccounts()
  if (accounts.length === 0) throw new Error('Not signed in')
  const request = { scopes: GRAPH_SCOPES, account: accounts[0] }
  try {
    const r = await msalInstance.acquireTokenSilent(request)
    return r.accessToken
  } catch {
    const r = await msalInstance.acquireTokenPopup(request)
    return r.accessToken
  }
}

// Create Outlook Calendar event (+ optional Teams meeting + attendees)
async function createCalendarEvent(ev: { subject: string; start: string; end: string; body?: string; attendees: string[]; isOnlineMeeting: boolean }): Promise<void> {
  const token = await getGraphToken()
  const payload = {
    subject: ev.subject,
    start: { dateTime: ev.start, timeZone: 'Asia/Bangkok' },
    end: { dateTime: ev.end, timeZone: 'Asia/Bangkok' },
    body: ev.body ? { contentType: 'HTML', content: ev.body.replace(/\n/g, '<br>') } : undefined,
    attendees: ev.attendees.filter(Boolean).map(email => ({ emailAddress: { address: email }, type: 'required' })),
    isOnlineMeeting: ev.isOnlineMeeting,
    onlineMeetingProvider: ev.isOnlineMeeting ? 'teamsForBusiness' : undefined,
  }
  const res = await fetch('https://graph.microsoft.com/v1.0/me/events', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
  if (!res.ok) throw new Error(`Calendar error ${res.status}: ${await res.text()}`)
}

async function fetchProjects(): Promise<void> {
  try {
    const token = await getToken()
    const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$orderby=Title asc&$top=500`
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata' } })
    if (res.ok) {
      const data = await res.json() as { value: { Id: number; Title: string }[] }
      state.projects = data.value.map(p => ({ id: p.Id, Title: p.Title }))
    }
  } catch { /* silent */ }
}

async function fetchAgents(): Promise<void> {
  try {
    const token = await getToken()
    const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('HD_AgentProfiles')/items?$select=Title,EmailText&$orderby=Title asc`
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata' } })
    if (res.ok) {
      const data = await res.json() as { value: { Title: string; EmailText: string }[] }
      state.agents = data.value.map(a => ({ email: a.EmailText, name: a.Title }))
    }
  } catch { /* silent */ }
}

async function fetchTickets(): Promise<void> {
  try {
    const token = await getToken()
    // เปิดอยู่ก่อน (ไม่ Closed) เรียงล่าสุด
    const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('HD_Tickets')/items?$select=Id,Title,TicketNumber,Status&$filter=Status ne 'Closed'&$orderby=Modified desc&$top=200`
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata' } })
    if (res.ok) {
      const data = await res.json() as { value: { Id: number; Title: string; TicketNumber: string; Status: string }[] }
      state.tickets = data.value.map(t => ({ id: t.Id, Title: t.Title, TicketNumber: t.TicketNumber, Status: t.Status }))
    }
  } catch { /* silent */ }
}

async function fetchContacts(): Promise<void> {
  try {
    const token = await getToken()
    const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('HD_Contracts')/items?$select=CustomerEmail&$top=2000`
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata' } })
    if (res.ok) {
      const data = await res.json() as { value: { CustomerEmail?: string }[] }
      state.contactEmails = data.value.map(c => (c.CustomerEmail || '').trim().toLowerCase()).filter(Boolean)
    }
  } catch { /* silent */ }
}

async function login(): Promise<void> {
  const btn = document.getElementById('btn-login-main') as HTMLButtonElement | null
  const btn2 = document.getElementById('btn-login') as HTMLButtonElement | null
  if (btn) { btn.disabled = true; btn.textContent = 'กำลังเข้าสู่ระบบ…' }
  if (btn2) { btn2.disabled = true }
  try {
    const result = await msalInstance.loginPopup({ scopes: [SP_SCOPE] })
    state.account = result.account
    await Promise.all([fetchProjects(), fetchAgents(), fetchTickets(), fetchContacts()])
    render()
  } catch {
    if (btn) { btn.disabled = false; btn.textContent = 'เข้าสู่ระบบ' }
    if (btn2) { btn2.disabled = false }
  }
}

async function logout(): Promise<void> {
  if (state.account) {
    await msalInstance.logoutPopup({ account: state.account })
  }
  state.account = null
  render()
}

// ─── SharePoint REST ──────────────────────────────────────────────────────────
async function spCreate(listTitle: string, body: Record<string, unknown>): Promise<number> {
  const token = await getToken()
  const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`

  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
    },
    body: JSON.stringify(body),
  })

  if (!res.ok) {
    const errText = await res.text()
    throw new Error(`SharePoint error ${res.status}: ${errText}`)
  }
  const data = await res.json() as { Id: number }
  return data.Id
}

// ─── Email notifications (ทำงานเหมือน webapp: HD_EmailTemplates + Graph sendMail) ──
interface EmailTemplate { EventKey: string; Subject: string; Body: string; IsEnabled: boolean }
let _tplCache: EmailTemplate[] | null = null
const DEFAULT_SENDER = 'support@itservices.co.th'

async function getEmailTemplates(): Promise<EmailTemplate[]> {
  if (_tplCache) return _tplCache
  try {
    const token = await getToken()
    const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('HD_EmailTemplates')/items?$select=EventKey,Subject,Body,IsEnabled&$top=50`
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata' } })
    if (!res.ok) return []
    const data = await res.json() as { value: EmailTemplate[] }
    _tplCache = data.value
    return _tplCache
  } catch { return [] }
}

async function getSenderAddress(): Promise<string> {
  try {
    const token = await getToken()
    const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('HD_Options')/items?$select=Title,Category&$filter=Category eq 'EmailConfig'&$top=1`
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata' } })
    if (!res.ok) return DEFAULT_SENDER
    const data = await res.json() as { value: { Title: string }[] }
    return data.value[0]?.Title?.trim() || DEFAULT_SENDER
  } catch { return DEFAULT_SENDER }
}

function renderTpl(tpl: string, vars: Record<string, string>): string {
  return tpl.replace(/\{\{(\w+)\}\}/g, (_, k) => vars[k] ?? `{{${k}}}`)
}

async function sendTemplateEmail(eventKey: string, vars: Record<string, string>, to: string[], cc: string[] = []): Promise<void> {
  try {
    const templates = await getEmailTemplates()
    const tpl = templates.find(t => t.EventKey === eventKey && t.IsEnabled)
    if (!tpl) return
    const subject = renderTpl(tpl.Subject || '', vars)
    const body = renderTpl(tpl.Body || '', vars)
    if (!subject || !body) return
    const norm = (e: string) => e.trim().toLowerCase()
    const toArr = [...new Map(to.filter(Boolean).map(e => [norm(e), e])).values()]
    if (toArr.length === 0) return
    const toSet = new Set(toArr.map(norm))
    const ccArr = [...new Map(cc.filter(Boolean).map(e => [norm(e), e])).values()].filter(e => !toSet.has(norm(e)))
    const from = await getSenderAddress()
    const token = await getGraphToken()
    const message: Record<string, unknown> = {
      subject,
      body: { contentType: 'HTML', content: body },
      toRecipients: toArr.map(a => ({ emailAddress: { address: a } })),
    }
    if (ccArr.length) message.ccRecipients = ccArr.map(a => ({ emailAddress: { address: a } }))
    if (from) message.from = { emailAddress: { address: from } }
    await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ message, saveToSentItems: true }),
    })
  } catch { /* email fail = non-critical */ }
}

// In-app notification (เหมือน webapp) — เขียนลง HD_Notifications, ตัดคนกดเองออก
async function createNotification(p: { recipients: string[]; title: string; message: string; linkPath: string; eventType: string }): Promise<void> {
  const norm = (e: string) => e.trim().toLowerCase()
  const actor = norm(state.account?.username ?? '')
  const seen = new Set<string>()
  const to = p.recipients.filter(Boolean).filter(e => { const k = norm(e); if (!k || k === actor || seen.has(k)) return false; seen.add(k); return true })
  if (to.length === 0) return
  try {
    const token = await getToken()
    const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('HD_Notifications')/items`
    await Promise.all(to.map(email => fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
      body: JSON.stringify({ Title: p.title.slice(0, 255), RecipientEmail: email, EventType: p.eventType, Message: p.message, LinkPath: p.linkPath, IsRead: false }),
    })))
  } catch { /* non-critical */ }
}

async function uploadEmailAttachments(listTitle: string, itemId: number): Promise<void> {
  const checked = document.querySelectorAll<HTMLInputElement>('.email-att-cb:checked')
  if (checked.length === 0) return

  const token = await getToken()

  for (const cb of Array.from(checked)) {
    const attId = cb.dataset['attId']!
    const attName = cb.dataset['attName']!

    // Get attachment content from Outlook
    const content = await new Promise<string>((resolve, reject) => {
      Office.context.mailbox.item!.getAttachmentContentAsync(attId, {}, result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value.content)
        else reject(new Error(result.error.message))
      })
    })

    // Decode base64 → ArrayBuffer
    const binary = atob(content)
    const bytes = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)

    const fileName = encodeURIComponent(attName)
    const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`
    const res = await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/octet-stream' },
      body: bytes.buffer,
    })
    if (!res.ok) throw new Error(`Upload ${attName} failed`)
  }
}

async function spUploadFileList(listTitle: string, itemId: number, files: File[]): Promise<void> {
  const token = await getToken()
  for (const file of files) {
    const buffer = await file.arrayBuffer()
    const fileName = encodeURIComponent(file.name)
    const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`
    const res = await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/octet-stream' },
      body: buffer,
    })
    if (!res.ok) throw new Error(`Upload ${file.name} failed`)
  }
}

async function spUploadAttachments(listTitle: string, itemId: number, files: FileList): Promise<void> {
  const token = await getToken()
  for (const file of Array.from(files)) {
    const buffer = await file.arrayBuffer()
    const fileName = encodeURIComponent(file.name)
    const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/octet-stream',
      },
      body: buffer,
    })
    if (!res.ok) {
      const errText = await res.text()
      throw new Error(`Upload ${file.name} failed: ${errText}`)
    }
  }
}

// ─── Toast ────────────────────────────────────────────────────────────────────
function showToast(message: string, type: 'success' | 'error' = 'success'): void {
  const container = document.getElementById('toast-container')
  if (!container) return

  const bgColor = type === 'success' ? 'bg-green-500' : 'bg-red-500'
  const icon = type === 'success' ? '✅' : '❌'

  const toast = document.createElement('div')
  toast.className = `toast pointer-events-auto ${bgColor} text-white text-sm font-medium px-4 py-3 rounded-lg shadow-lg max-w-xs mx-2`
  toast.textContent = `${icon} ${message}`

  container.appendChild(toast)
  setTimeout(() => toast.remove(), 4000)
}

// ─── Signature parser ─────────────────────────────────────────────────────────
function parseSignature(sigText: string): SignatureContact | null {
  const lines = sigText.split('\n').map(l => l.trim()).filter(Boolean)

  // Email regex
  const emailRe = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/
  // Phone: +66, 0x-xxxx-xxxx, ext, etc.
  const phoneRe = /(\+?[\d\s()\-]{7,}(?:ext\.?\s*\d+)?)/i
  // Company keywords
  const companyRe = /\b(co\.?,?\s*ltd\.?|co\.?,?\s*limited|corporation|corp\.?|บริษัท|จำกัด|holding|group|inc\.?|llc)\b/i

  let email = ''
  let phone = ''
  let company = ''
  const nameCandidates: string[] = []

  for (const line of lines) {
    // Skip separator lines
    if (/^[-_=*]{2,}$/.test(line)) continue

    // Greeting/closing words to skip
    if (/^(best regards|regards|sincerely|ขอแสดงความนับถือ|ด้วยความนับถือ|from|sent|thanks|thank you)/i.test(line)) continue

    if (!email) {
      const m = line.match(emailRe)
      if (m) { email = m[0]; continue }
    }

    if (!phone) {
      const m = line.match(phoneRe)
      // Must have at least 7 digits
      if (m && m[0].replace(/\D/g, '').length >= 7) { phone = m[0].trim(); continue }
    }

    if (!company && companyRe.test(line)) {
      company = line; continue
    }

    // Remaining short lines (2-50 chars) are name candidates
    if (line.length >= 2 && line.length <= 50 && !/\d{4,}/.test(line)) {
      nameCandidates.push(line)
    }
  }

  // Pick first name candidate that isn't email/phone/company
  const name = nameCandidates.find(n => !emailRe.test(n) && !companyRe.test(n)) ?? ''

  if (!email && !name) return null
  return { name, company, email, phone }
}

async function importAsCustomer(): Promise<void> {
  const sig = state.signatureContact
  if (!sig) return
  // กันเพิ่มซ้ำ — ถ้ามีในระบบแล้วไม่ต้องส่ง
  const emailLc = (state.emailSenderEmail || '').toLowerCase()
  if (emailLc && state.contactEmails.includes(emailLc)) {
    showToast('ลูกค้านี้มีในระบบแล้ว', 'success')
    state.signatureContact = null; render(); return
  }
  const btn = document.getElementById('btn-import-customer') as HTMLButtonElement | null
  if (btn) { btn.disabled = true; btn.textContent = 'กำลังบันทึก…' }
  try {
    await spCreate('HD_Contracts', {
      Title: state.emailSenderName || sig.name,
      CustomerEmail: state.emailSenderEmail,
      Phone: sig.phone || undefined,
      Company: sig.company || undefined,
      Status: 'Active',
    })
    if (emailLc) state.contactEmails.push(emailLc)
    showToast('เพิ่มลูกค้าสำเร็จ!')
    state.signatureContact = null
    render()
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    showToast(`เกิดข้อผิดพลาด: ${msg}`, 'error')
    if (btn) { btn.disabled = false; btn.textContent = 'เพิ่มเป็นลูกค้า' }
  }
}

// ─── Today's date (ISO) ───────────────────────────────────────────────────────
function todayISO(): string {
  return new Date().toISOString().split('T')[0]
}

// ─── Ticket number (รูปแบบเดียวกับ webapp) ──────────────────────────────────────
function genTicketNumber(): string {
  const now = new Date()
  const ymd = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`
  return `HD-${ymd}-${Math.floor(Math.random() * 900 + 100)}`
}

// ─── Submit handlers ──────────────────────────────────────────────────────────
let _submitting = false
async function handleSubmit(): Promise<void> {
  if (!state.account) {
    showToast('กรุณาเข้าสู่ระบบก่อน', 'error')
    return
  }
  if (_submitting) return            // กันกดซ้ำระหว่างกำลังส่ง
  _submitting = true

  const btn = document.getElementById('submit-btn') as HTMLButtonElement | null
  if (btn) { btn.disabled = true; btn.textContent = 'กำลังบันทึก…' }

  try {
    if (state.tab === 'ticket') {
      const title = (document.getElementById('f-title') as HTMLInputElement).value.trim()
      const description = (document.getElementById('f-description') as HTMLTextAreaElement).value.trim()
      const priority = (document.getElementById('f-priority') as HTMLSelectElement).value
      const customerEmail = (document.getElementById('f-customer-email') as HTMLInputElement).value.trim()
      const ccEnabled = (document.getElementById('f-cc-enable') as HTMLInputElement)?.checked ?? true
      const ccEmails = ccEnabled
        ? ((document.getElementById('f-cc') as HTMLInputElement)?.value || '').split(/[,;\s]+/).map(s => s.trim()).filter(Boolean)
        : []

      const assignedEmail = (document.getElementById('f-assigned-email') as HTMLSelectElement).value
      const assignedAgent = state.agents.find(a => a.email === assignedEmail)
      const ticketNum = genTicketNumber()
      const ticketId = await spCreate('HD_Tickets', {
        Title: title,
        TicketNumber: ticketNum,
        Description: description,
        Priority: priority,
        CustomerEmail: customerEmail,
        CustomerName: state.emailSenderName || customerEmail,
        Status: 'Open',
        AssignedEmail: assignedEmail || undefined,
        AssignedToName: assignedAgent?.name ?? state.account?.name ?? '',
      })
      if (state.droppedFiles.length > 0) await spUploadFileList('HD_Tickets', ticketId, state.droppedFiles)
      await uploadEmailAttachments('HD_Tickets', ticketId)
      state.droppedFiles = []
      // Email: 1 ฉบับ — To = ลูกค้า, CC = agent + ผู้แจ้ง (เหมือน webapp)
      await sendTemplateEmail('ticket_created', {
        ticket_number: ticketNum,
        ticket_title: title,
        priority,
        category: '-',
        description: (description || '-').replace(/\n/g, '<br>'),
        customer_name: state.emailSenderName || customerEmail,
        assigned_name: assignedAgent?.name ?? state.account?.name ?? '-',
        link: 'https://itservices.co.th/helpdesk/',
      }, [customerEmail], [assignedEmail, state.account.username, ...ccEmails])
      showToast('สร้าง Ticket สำเร็จ!')

    } else if (state.tab === 'task') {
      const title = (document.getElementById('f-title') as HTMLInputElement).value.trim()
      const projectId = parseInt((document.getElementById('f-project') as HTMLSelectElement)?.value || '0')
      const dueDate = (document.getElementById('f-due-date') as HTMLInputElement).value
      const note = (document.getElementById('f-note') as HTMLTextAreaElement).value.trim()

      const assignedEmail = (document.getElementById('f-assigned-email') as HTMLSelectElement).value
      const assignedAgent = state.agents.find(a => a.email === assignedEmail)

      if (!projectId) { showToast('กรุณาเลือก Project', 'error'); return }

      const taskId = await spCreate('PM_Tasks', {
        Title: title,
        DueDate: dueDate || null,
        TaskNote: note,
        AssignedTo: assignedAgent?.name ?? state.account.name ?? state.account.username,
        AssignedEmail: assignedEmail,
        IsCompleted: false,
        IsAcknowledged: false,
        ProjectID: projectId,
      })
      if (state.droppedFiles.length > 0) await spUploadFileList('PM_Tasks', taskId, state.droppedFiles)
      await uploadEmailAttachments('PM_Tasks', taskId)
      state.droppedFiles = []

      // แจ้งเตือน agent ที่ถูก assign (in-app เหมือน webapp) — ยกเว้นคนสร้างเอง
      await createNotification({
        recipients: [assignedEmail],
        title: `📋 ได้รับมอบหมาย Task: ${title}`,
        message: note || (dueDate ? `กำหนดส่ง ${dueDate}` : 'มี Task ใหม่'),
        linkPath: projectId ? `/projects/${projectId}` : '/my-work',
        eventType: 'task_assigned',
      })

      // เพิ่มการประชุม Teams / Calendar (ถ้าติ๊ก)
      const isMeeting = (document.getElementById('f-teams') as HTMLInputElement)?.checked
      if (isMeeting && dueDate) {
        const internal = Array.from(document.querySelectorAll<HTMLInputElement>('.att-internal:checked')).map(c => c.value)
        const external = ((document.getElementById('f-ext-att') as HTMLInputElement)?.value || '')
          .split(/[,;\s]+/).map(s => s.trim()).filter(Boolean)
        const start = `${dueDate}T09:00:00`
        const end = `${dueDate}T10:00:00`
        try {
          await createCalendarEvent({ subject: title, start, end, body: note, attendees: [...internal, ...external], isOnlineMeeting: true })
          showToast('สร้าง Task + นัดประชุม Teams สำเร็จ!')
        } catch (e) {
          showToast('สร้าง Task แล้ว แต่สร้างนัดประชุมไม่สำเร็จ: ' + (e instanceof Error ? e.message : ''), 'error')
        }
      } else {
        showToast('สร้าง Task สำเร็จ!')
      }

    } else if (state.tab === 'incident') {
      const title = (document.getElementById('f-title') as HTMLInputElement).value.trim()
      const projectId = parseInt((document.getElementById('f-project') as HTMLSelectElement)?.value || '0')
      const description = (document.getElementById('f-description') as HTMLTextAreaElement).value.trim()
      const severity = (document.getElementById('f-severity') as HTMLSelectElement).value

      const assignedEmail = (document.getElementById('f-assigned-email') as HTMLSelectElement).value
      const assignedAgent = state.agents.find(a => a.email === assignedEmail)
      const status = (document.getElementById('f-status') as HTMLSelectElement).value
      const incidentDate = (document.getElementById('f-incident-date') as HTMLInputElement).value
      const resolution = (document.getElementById('f-resolution') as HTMLTextAreaElement).value.trim()

      if (!projectId) { showToast('กรุณาเลือก Project', 'error'); return }

      const incidentId = await spCreate('PM_Incidents', {
        Title: title,
        Description: description || undefined,
        Severity: severity,
        Status: status,
        AssignedTo: assignedAgent?.name ?? state.account.name ?? state.account.username,
        AssignedEmail: assignedEmail,
        ProjectID: projectId,
        IncidentDate: incidentDate || todayISO(),
        Resolution: resolution || undefined,
      })
      if (state.droppedFiles.length > 0) await spUploadFileList('PM_Incidents', incidentId, state.droppedFiles)
      await uploadEmailAttachments('PM_Incidents', incidentId)
      state.droppedFiles = []
      // แจ้งเตือน Assigned (in-app เหมือน webapp) — ยกเว้นคนสร้างเอง
      await createNotification({
        recipients: [assignedEmail],
        title: `🚨 ได้รับมอบหมาย Incident: ${title}`,
        message: `ความรุนแรง ${severity}${description ? ' — ' + description.slice(0, 120) : ''}`,
        linkPath: projectId ? `/projects/${projectId}` : '/my-work',
        eventType: 'incident_created',
      })
      showToast('สร้าง Incident สำเร็จ!')

    } else if (state.tab === 'comment') {
      const ticketId = parseInt((document.getElementById('f-ticket') as HTMLSelectElement)?.value || '0')
      const commentText = (document.getElementById('f-comment') as HTMLTextAreaElement).value.trim()
      const commentType = (document.getElementById('f-comment-type') as HTMLSelectElement).value
      if (!ticketId) { showToast('กรุณาเลือก Ticket', 'error'); return }
      if (!commentText) { showToast('กรุณาพิมพ์ Comment', 'error'); return }

      await spCreate('HD_TicketComments', {
        Title: commentText.slice(0, 100),
        TicketID: ticketId,
        CommentText: commentText,
        CommentType: commentType,
        CommentDate: new Date().toISOString(),
      })
      // แนบไฟล์เข้า ticket เดิม (ถ้ามี)
      if (state.droppedFiles.length > 0) await spUploadFileList('HD_Tickets', ticketId, state.droppedFiles)
      await uploadEmailAttachments('HD_Tickets', ticketId)
      state.droppedFiles = []
      // Email: แจ้งภายใน (agent + ผู้แจ้ง) ยกเว้นคนกดเอง — ดึงข้อมูล ticket ก่อน
      try {
        const token = await getToken()
        const tUrl = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('HD_Tickets')/items(${ticketId})?$select=TicketNumber,Title,AssignedEmail,Author/EMail&$expand=Author`
        const tRes = await fetch(tUrl, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata' } })
        if (tRes.ok) {
          const t = await tRes.json() as { TicketNumber?: string; Title?: string; AssignedEmail?: string; Author?: { EMail?: string } }
          const actor = state.account.username.toLowerCase()
          const internal = [...new Set([t.AssignedEmail, t.Author?.EMail].filter(Boolean) as string[])]
            .filter(e => e.toLowerCase() !== actor)
          if (internal.length) {
            await createNotification({
              recipients: internal,
              title: `💬 ${state.account?.name ?? 'มีคน'} คอมเมนต์ใน ${t.TicketNumber || ('#' + ticketId)}`,
              message: commentText.slice(0, 200),
              linkPath: `/tickets/${ticketId}`,
              eventType: 'comment_added',
            })
          }
        }
      } catch { /* non-critical */ }
      showToast('เพิ่ม Comment สำเร็จ!')

    } else if (state.tab === 'project') {
      const title = (document.getElementById('f-title') as HTMLInputElement).value.trim()
      const company = (document.getElementById('f-company') as HTMLInputElement).value.trim()
      const projectGroup = (document.getElementById('f-group') as HTMLSelectElement).value
      const status = (document.getElementById('f-status') as HTMLSelectElement).value
      const startDate = (document.getElementById('f-start') as HTMLInputElement).value
      const endDate = (document.getElementById('f-end') as HTMLInputElement).value
      const description = (document.getElementById('f-description') as HTMLTextAreaElement).value.trim()
      if (!title) { showToast('กรุณาใส่ชื่อโครงการ', 'error'); return }

      const newProjectId = await spCreate('PM_Projects', {
        Title: title,
        Company: company || undefined,
        ProjectGroup: projectGroup,
        Progress: 0,
        StartDate: startDate || undefined,
        EndDate: endDate || null,
        Status: status,
        CreatedByEmail: state.account.username,
        Comment: description || undefined,
      })
      if (state.droppedFiles.length > 0) await spUploadFileList('PM_Projects', newProjectId, state.droppedFiles)
      await uploadEmailAttachments('PM_Projects', newProjectId)
      state.droppedFiles = []
      showToast('สร้างโครงการสำเร็จ!')
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    showToast(`เกิดข้อผิดพลาด: ${msg}`, 'error')
  } finally {
    _submitting = false
    if (btn) { btn.disabled = false; btn.textContent = 'บันทึก' }
  }
}

// ─── Tab labels ───────────────────────────────────────────────────────────────
const TAB_META: Record<Tab, { label: string; icon: string }> = {
  ticket:   { label: 'Ticket',   icon: '🎫' },
  task:     { label: 'Task',     icon: '✅' },
  incident: { label: 'Incident', icon: '🚨' },
  comment:  { label: 'Comment',  icon: '💬' },
  project:  { label: 'Project',  icon: '📁' },
}

// ─── Render ───────────────────────────────────────────────────────────────────
const FORM_IDS = ['f-title','f-description','f-priority','f-customer-email','f-cc','f-assigned-email','f-project','f-due-date','f-note','f-severity','f-status','f-incident-date','f-resolution','f-ticket','f-comment','f-comment-type','f-company','f-group','f-start','f-end','f-ext-att']
let formCache: Record<string, string | boolean> = {}
function captureForm(): void {
  for (const id of FORM_IDS) {
    const el = document.getElementById(id) as HTMLInputElement | null
    if (el) formCache[id] = el.value
  }
  const teams = document.getElementById('f-teams') as HTMLInputElement | null
  if (teams) formCache['f-teams'] = teams.checked
}
function restoreForm(): void {
  for (const id of FORM_IDS) {
    const el = document.getElementById(id) as HTMLInputElement | null
    if (el && formCache[id] !== undefined && formCache[id] !== '') el.value = formCache[id] as string
  }
  const teams = document.getElementById('f-teams') as HTMLInputElement | null
  if (teams && formCache['f-teams'] !== undefined) {
    teams.checked = formCache['f-teams'] as boolean
    const tf = document.getElementById('teams-fields')
    if (tf) tf.style.display = teams.checked ? 'block' : 'none'
  }
}

function render(): void {
  const app = document.getElementById('app')
  if (!app) return

  captureForm()  // preserve typed values across re-render

  const { account, tab, emailSubject, emailSenderName, emailSenderEmail, emailBodyPreview } = state
  const isLoggedIn = account !== null

  // ── Header ──
  const headerHTML = `
    <div class="bg-gradient-to-r from-blue-700 to-blue-600 text-white px-4 py-2.5 flex items-center gap-2.5 shadow flex-shrink-0">
      <div class="w-7 h-7 bg-white/95 rounded-lg flex items-center justify-center flex-shrink-0">
        <span class="text-blue-700 font-extrabold text-xs">iT</span>
      </div>
      <div class="min-w-0 flex-1">
        <div class="font-semibold text-sm leading-tight truncate">iT Services Helpdesk</div>
        ${isLoggedIn ? `<div class="text-[10px] text-blue-100 truncate">${esc(account?.name ?? account?.username ?? '')}</div>` : ''}
      </div>
      ${isLoggedIn
        ? `<button id="btn-logout" title="ออกจากระบบ" class="p-1.5 rounded-lg hover:bg-white/15 transition flex-shrink-0">
             <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
           </button>`
        : `<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded-lg transition flex-shrink-0">
             เข้าสู่ระบบ
           </button>`
      }
    </div>
  `

  // ── Login wall ──
  if (!isLoggedIn) {
    app.innerHTML = `
      ${headerHTML}
      <div class="flex flex-col items-center justify-center px-6 py-16 gap-4 text-center">
        <div class="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center text-3xl">🔐</div>
        <h2 class="text-base font-semibold text-slate-700">เข้าสู่ระบบด้วย Microsoft</h2>
        <p class="text-xs text-slate-500">เพื่อสร้าง Ticket / Task / Incident จาก Email นี้</p>
        <button id="btn-login-main"
          class="mt-2 w-full bg-blue-700 hover:bg-blue-600 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          เข้าสู่ระบบ
        </button>
      </div>
    `
    document.getElementById('btn-login')?.addEventListener('click', login)
    document.getElementById('btn-login-main')?.addEventListener('click', login)
    return
  }

  // ── Email info box ──
  const emailInfoHTML = emailSubject
    ? `<div class="mx-3 mt-3 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2.5 text-xs text-slate-700 space-y-1">
         <div class="font-semibold text-blue-800 truncate" title="${esc(emailSubject)}">📧 ${esc(emailSubject)}</div>
         ${emailSenderName
           ? `<div class="text-slate-500">จาก: <span class="font-medium text-slate-700">${esc(emailSenderName)}</span></div>`
           : ''}
         ${emailSenderEmail && emailSenderEmail !== emailSenderName
           ? `<div class="text-slate-400 truncate">${esc(emailSenderEmail)}</div>`
           : ''}
       </div>`
    : `<div class="mx-3 mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-700">
         ⚠️ ไม่พบข้อมูล Email (โหมดทดสอบ)
       </div>`

  // ── Signature contact card ──
  const sig = state.signatureContact
  const custExists = !!emailSenderEmail && state.contactEmails.includes(emailSenderEmail.toLowerCase())
  const sigCardHTML = sig
    ? `<div class="mx-3 mt-3 bg-orange-50 border border-orange-200 rounded-lg px-3 py-2.5 text-xs text-slate-700">
         <div class="flex items-center justify-between mb-2">
           <span class="font-semibold text-orange-700">👤 ข้อมูลผู้ส่ง (จาก Signature)</span>
         </div>
         <div class="space-y-0.5 mb-2.5">
           ${emailSenderName  ? `<div><span class="text-slate-400">ชื่อ:</span> <span class="font-medium">${esc(emailSenderName)}</span></div>` : ''}
           ${sig.company      ? `<div><span class="text-slate-400">บริษัท:</span> ${esc(sig.company)}</div>` : ''}
           ${emailSenderEmail ? `<div><span class="text-slate-400">Email:</span> ${esc(emailSenderEmail)}</div>` : ''}
           ${sig.phone        ? `<div><span class="text-slate-400">โทร:</span> ${esc(sig.phone)}</div>` : ''}
         </div>
         ${custExists
           ? `<div class="w-full bg-green-100 text-green-700 text-xs font-semibold py-1.5 rounded-md text-center">✓ ลูกค้านี้มีในระบบแล้ว</div>`
           : `<button id="btn-import-customer"
                class="w-full bg-orange-500 hover:bg-orange-400 text-white text-xs font-semibold py-1.5 rounded-md transition">
                + เพิ่มเป็นลูกค้า
              </button>`}
       </div>`
    : ''

  // ── Tab switcher ──
  const tabsHTML = `
    <div class="grid grid-cols-5 gap-1 mx-3 mt-3">
      ${(Object.entries(TAB_META) as [Tab, { label: string; icon: string }][]).map(([key, meta]) => `
        <button data-tab="${key}"
          class="tab-btn flex flex-col items-center gap-1 py-2 rounded-lg transition ${
            tab === key ? 'bg-blue-700 text-white shadow' : 'text-slate-500 hover:bg-slate-100'}">
          <span class="text-base leading-none">${meta.icon}</span>
          <span class="text-[9px] font-medium leading-none">${meta.label}</span>
        </button>
      `).join('')}
    </div>
  `

  // ── Form fields by tab ──
  let formHTML = ''

  if (tab === 'ticket') {
    formHTML = `
      ${field('Title / หัวข้อ', `<input id="f-title" type="text"
        class="${inputCls}"
        value="${esc(emailSubject)}" />`)}
      ${field('รายละเอียด', `<textarea id="f-description" rows="4"
        class="${inputCls} resize-none">${esc(emailBodyPreview)}</textarea>`)}
      ${field('Priority', `<select id="f-priority" class="${inputCls}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
      ${field('Customer Email', `<input id="f-customer-email" type="email"
        class="${inputCls}"
        value="${esc(emailSenderEmail)}" />`)}
      ${field('CC — ให้ผู้ที่อยู่ในเมลนี้รับรู้', `
        <label class="flex items-center gap-2 text-xs text-slate-600 mb-1.5 cursor-pointer">
          <input id="f-cc-enable" type="checkbox" ${state.emailCc.length ? 'checked' : ''} /> แนบผู้รับในเมลนี้เป็น CC อัตโนมัติ
        </label>
        <input id="f-cc" type="text" class="${inputCls}" value="${esc(state.emailCc.join(', '))}" placeholder="someone@company.com, boss@company.com" />`)}
      ${field('Assign ให้ Agent', agentSelect(account.username))}
      ${fileField()}
    `
  } else if (tab === 'task') {
    formHTML = `
      ${field('ชื่อ Task *', `<input id="f-title" type="text" required
        class="${inputCls}" value="${esc(emailSubject)}" />`)}
      ${field('Project *', projectSelect())}
      ${field('Assign ให้', agentSelect(account.username))}
      ${field('Due Date', `<input id="f-due-date" type="date" class="${inputCls}" />`)}
      ${field('Task Note', `<textarea id="f-note" rows="4"
        class="${inputCls} resize-y">${esc(emailBodyPreview)}</textarea>`)}
      <label class="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer mb-1">
        <input id="f-teams" type="checkbox" class="rounded" onchange="document.getElementById('teams-fields').style.display=this.checked?'block':'none'" />
        💻 เพิ่มการประชุมออนไลน์ (Teams) — ใช้เวลา 09:00–10:00 ของวัน Due Date
      </label>
      <div id="teams-fields" style="display:none" class="space-y-2 mb-2">
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">ผู้เข้าร่วม Internal</label>
          <div class="max-h-28 overflow-y-auto border border-slate-200 rounded-md p-1.5 space-y-0.5">
            ${state.agents.map(a => `<label class="flex items-center gap-2 text-xs text-slate-700 px-1 py-0.5 hover:bg-slate-50 rounded cursor-pointer">
              <input type="checkbox" class="att-internal" value="${esc(a.email)}" /> ${esc(a.name)}
            </label>`).join('')}
          </div>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-600 mb-1">Email ภายนอก (คั่นด้วย ,)</label>
          <input id="f-ext-att" type="text" class="${inputCls}" placeholder="someone@company.com, ..." />
        </div>
      </div>
      ${fileField()}
    `
  } else if (tab === 'incident') {
    formHTML = `
      ${field('ชื่อ Incident *', `<input id="f-title" type="text" required
        class="${inputCls}" value="${esc(emailSubject)}" />`)}
      ${field('Project *', projectSelect())}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">ความรุนแรง</label>
          <select id="f-severity" class="${inputCls}">
            <option value="Low">Low</option>
            <option value="Medium" selected>Medium</option>
            <option value="High">High</option>
            <option value="Critical">Critical</option>
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${inputCls}">
            <option value="Open" selected>Open</option>
            <option value="In Progress">In Progress</option>
            <option value="Resolved">Resolved</option>
          </select>
        </div>
      </div>
      ${field('Assign ให้ Agent', agentSelect(account.username))}
      ${field('วันที่เกิด Incident', `<input id="f-incident-date" type="date" class="${inputCls}" value="${todayISO()}" />`)}
      ${field('รายละเอียด', `<textarea id="f-description" rows="4"
        class="${inputCls} resize-y">${esc(emailBodyPreview)}</textarea>`)}
      ${field('วิธีแก้ไข (ถ้ามี)', `<textarea id="f-resolution" rows="2"
        class="${inputCls} resize-y" placeholder="อธิบายวิธีแก้ไขปัญหา..."></textarea>`)}
      ${fileField()}
    `
  } else if (tab === 'comment') {
    formHTML = `
      ${field('เลือก Ticket *', `<select id="f-ticket" class="${inputCls}">
        <option value="">-- เลือก Ticket ที่จะเพิ่ม Comment --</option>
        ${state.tickets.map(t => `<option value="${t.id}">${esc(t.TicketNumber || ('#' + t.id))} · ${esc(t.Title)}</option>`).join('')}
      </select>`)}
      ${field('ประเภท', `<select id="f-comment-type" class="${inputCls}">
        <option value="Internal">Internal</option>
        <option value="External">External</option>
      </select>`)}
      ${field('Comment *', `<textarea id="f-comment" rows="5"
        class="${inputCls} resize-y" placeholder="พิมพ์ comment...">${esc(emailBodyPreview)}</textarea>`)}
      ${fileField()}
    `
  } else if (tab === 'project') {
    formHTML = `
      ${field('ชื่อโครงการ *', `<input id="f-title" type="text" required
        class="${inputCls}" value="${esc(emailSubject)}" />`)}
      ${field('บริษัท / ลูกค้า', `<input id="f-company" type="text" class="${inputCls}" value="${esc(state.signatureContact?.company ?? '')}" />`)}
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">กลุ่มโครงการ</label>
          <select id="f-group" class="${inputCls}">
            ${['Internal', 'External', 'R&D', 'Maintenance', 'อื่นๆ'].map(g => `<option>${g}</option>`).join('')}
          </select>
        </div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">สถานะ</label>
          <select id="f-status" class="${inputCls}">
            ${['Planning', 'Active', 'On Hold', 'Completed', 'Cancelled'].map(s => `<option>${s}</option>`).join('')}
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันที่เริ่ม</label>
          <input id="f-start" type="date" class="${inputCls}" /></div>
        <div><label class="block text-xs font-medium text-slate-600 mb-1">วันสิ้นสุด</label>
          <input id="f-end" type="date" class="${inputCls}" /></div>
      </div>
      ${field('รายละเอียด', `<textarea id="f-description" rows="4"
        class="${inputCls} resize-y">${esc(emailBodyPreview)}</textarea>`)}
      ${fileField()}
    `
  }

  const submitLabel = tab === 'comment' ? 'เพิ่ม Comment'
    : tab === 'project' ? 'สร้างโครงการ'
    : tab === 'incident' ? 'แจ้ง Incident'
    : tab === 'task' ? 'สร้าง Task' : 'สร้าง Ticket'

  app.innerHTML = `
    <div class="flex flex-col h-screen bg-slate-50">
      ${headerHTML}
      <div class="flex-1 overflow-y-auto">
        ${emailInfoHTML}
        ${sigCardHTML}
        ${tabsHTML}
        <div class="mx-3 mt-3 space-y-3 pb-3">
          ${formHTML}
        </div>
      </div>
      <div class="border-t border-slate-200 bg-white px-3 py-2.5 flex-shrink-0">
        <button id="submit-btn"
          class="w-full bg-blue-700 hover:bg-blue-600 active:bg-blue-800 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition">
          ${submitLabel}
        </button>
      </div>
    </div>
  `

  // ── Event listeners ──
  document.getElementById('btn-logout')?.addEventListener('click', logout)
  document.getElementById('submit-btn')?.addEventListener('click', handleSubmit)
  document.getElementById('btn-import-customer')?.addEventListener('click', importAsCustomer)


  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const newTab = (btn as HTMLElement).dataset['tab'] as Tab
      if (newTab && newTab !== state.tab) {
        state.tab = newTab
        render()
      }
    })
  })

  // ── Drop zone ──
  const dropZone = document.getElementById('drop-zone')
  const fileInput = document.getElementById('f-files') as HTMLInputElement | null

  if (dropZone && fileInput) {
    // File input change (browse)
    fileInput.addEventListener('change', () => {
      if (fileInput.files) addDroppedFiles(Array.from(fileInput.files))
      fileInput.value = ''
    })

    // Drag over
    dropZone.addEventListener('dragover', e => {
      e.preventDefault()
      dropZone.classList.add('border-blue-500', 'bg-blue-50')
    })
    dropZone.addEventListener('dragleave', () => {
      dropZone.classList.remove('border-blue-500', 'bg-blue-50')
    })

    // Drop
    dropZone.addEventListener('drop', e => {
      e.preventDefault()
      dropZone.classList.remove('border-blue-500', 'bg-blue-50')
      const files = Array.from(e.dataTransfer?.files ?? [])
      if (files.length) addDroppedFiles(files)
    })
  }

  // Remove dropped file buttons
  document.querySelectorAll<HTMLButtonElement>('.remove-dropped').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset['remove'] ?? '-1')
      if (idx >= 0) {
        state.droppedFiles.splice(idx, 1)
        render()
      }
    })
  })

  restoreForm()  // re-apply typed values after re-render
}

function addDroppedFiles(files: File[]): void {
  state.droppedFiles.push(...files)
  render()
}

// ── Global paste handler (Ctrl+V / screenshot) ──
document.addEventListener('paste', (e: ClipboardEvent) => {
  // Only handle paste when logged in and task pane is visible
  if (!state.account) return
  const items = Array.from(e.clipboardData?.items ?? [])
  const files: File[] = []
  for (const item of items) {
    if (item.kind === 'file') {
      const file = item.getAsFile()
      if (file) {
        // Give screenshot a meaningful name with timestamp
        const name = file.name && file.name !== 'image.png'
          ? file.name
          : `screenshot-${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.png`
        files.push(new File([file], name, { type: file.type }))
      }
    }
  }
  if (files.length) {
    e.preventDefault()
    addDroppedFiles(files)
    showToast(`แนบไฟล์แล้ว: ${files.map(f => f.name).join(', ')}`)
  }
})

// ─── Helpers ──────────────────────────────────────────────────────────────────
const inputCls = 'w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white'

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`
}

function fileField(): string {
  const emailAtts = state.emailAttachments
  const dropped = state.droppedFiles

  const emailAttHtml = emailAtts.length > 0
    ? `<div class="mb-2 space-y-1">
        <p class="text-xs text-slate-500">📎 ไฟล์แนบจาก Email:</p>
        ${emailAtts.map(a => `
          <label class="flex items-center gap-2 text-xs text-slate-700 cursor-pointer hover:bg-slate-50 rounded px-1 py-0.5">
            <input type="checkbox" class="email-att-cb" data-att-id="${esc(a.id)}" data-att-name="${esc(a.name)}" checked />
            <span class="flex-1 truncate">${esc(a.name)}</span>
            <span class="text-slate-400 flex-shrink-0">${formatBytes(a.size)}</span>
          </label>`).join('')}
      </div>`
    : ''

  const droppedHtml = dropped.length > 0
    ? `<div class="mt-2 space-y-1">
        ${dropped.map((f, i) => {
          const isImg = f.type.startsWith('image/')
          return `<div class="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 rounded px-2 py-1">
            <span class="text-base">${isImg ? '🖼️' : '📄'}</span>
            <span class="flex-1 truncate">${esc(f.name)}</span>
            <span class="text-slate-400">${formatBytes(f.size)}</span>
            <button type="button" data-remove="${i}"
              class="remove-dropped text-red-400 hover:text-red-600 font-bold leading-none">✕</button>
          </div>`
        }).join('')}
      </div>`
    : ''

  return `<div class="space-y-1">
    <label class="block text-xs font-medium text-slate-600">ไฟล์แนบ</label>
    ${emailAttHtml}
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
    ${droppedHtml}
  </div>`
}

function agentSelect(currentEmail: string): string {
  return `<select id="f-assigned-email" class="${inputCls}">
    <option value="${esc(currentEmail)}">${esc(state.account?.name ?? currentEmail)} (ฉัน)</option>
    ${state.agents.filter(a => a.email !== currentEmail).map(a =>
      `<option value="${esc(a.email)}">${esc(a.name)}</option>`
    ).join('')}
  </select>`
}

function projectSelect(): string {
  if (state.projects.length === 0) {
    return `<div class="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-2">⚠️ ไม่พบ Project ที่ Active</div>`
  }
  return `<select id="f-project" class="${inputCls}">
    <option value="">-- เลือก Project --</option>
    ${state.projects.map(p => `<option value="${p.id}">${esc(p.Title)}</option>`).join('')}
  </select>`
}

function field(label: string, control: string): string {
  return `
    <div class="space-y-1">
      <label class="block text-xs font-medium text-slate-600">${label}</label>
      ${control}
    </div>
  `
}

function esc(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

// ─── Office.onReady ───────────────────────────────────────────────────────────
async function init(): Promise<void> {
  // Initialize MSAL
  await msalInstance.initialize()

  // Handle redirect (just in case)
  await msalInstance.handleRedirectPromise()

  // Restore account from cache
  const accounts = msalInstance.getAllAccounts()
  if (accounts.length > 0) {
    state.account = accounts[0]
    // Try silent token to verify it's still valid
    try {
      await msalInstance.acquireTokenSilent({ scopes: [SP_SCOPE], account: accounts[0] })
      await Promise.all([fetchProjects(), fetchAgents(), fetchTickets(), fetchContacts()])
    } catch {
      state.account = null
    }
  }

  // Try to read Office context
  if (typeof Office !== 'undefined') {
    Office.onReady(info => {
      if (info.host === Office.HostType.Outlook) {
        const item = Office.context.mailbox.item as Office.MessageRead | null
        if (item) {
          state.emailSubject = item.subject ?? ''

          // Sender
          const sender = item.from
          if (sender) {
            state.emailSenderName = sender.displayName ?? ''
            state.emailSenderEmail = sender.emailAddress ?? ''
          }

          // ผู้รับในเมล (To + CC) → ใช้ prefill ช่อง CC (ตัดผู้ส่ง + ตัวเราออก)
          const me = (state.account?.username ?? '').toLowerCase()
          const fromLc = (sender?.emailAddress ?? '').toLowerCase()
          const recips = [...(item.to ?? []), ...(item.cc ?? [])]
            .map(r => r.emailAddress).filter(Boolean)
          state.emailCc = [...new Set(recips.map(e => e.toLowerCase()))]
            .filter(e => e !== me && e !== fromLc)

          // Email attachments
          const attachments = item.attachments ?? []
          state.emailAttachments = attachments
            .filter(a => a.attachmentType === Office.MailboxEnums.AttachmentType.File)
            .map(a => ({ id: a.id, name: a.name, size: a.size }))

          // Body preview (async) — recursive DOM walker, handles tables natively
          item.body.getAsync(Office.CoercionType.Html, { asyncContext: {} }, result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const html = result.value as string
              const doc = new DOMParser().parseFromString(html, 'text/html')

              const SKIP = ['style','script','head','img','meta','link','noscript']
              const BLOCK = ['p','div','li','h1','h2','h3','h4','h5','h6','blockquote']
              const TABLE_WRAP = ['table','thead','tbody','tfoot']

              function walk(node: Node, insideTable = false): string {
                // Text node
                if (node.nodeType === 3) {
                  const t = node.textContent ?? ''
                  // Skip whitespace-only text nodes inside table containers
                  if (insideTable && t.trim() === '') return ''
                  return t
                }

                const el = node as Element
                const tag = (el.tagName ?? '').toLowerCase()

                // Skip noise
                if (SKIP.includes(tag)) return ''

                // <br> → space (Office HTML uses <br> between words in same paragraph)
                if (tag === 'br') return ' '

                // <tr> → join cells with tab, single newline at end
                if (tag === 'tr') {
                  const cells: string[] = []
                  for (let i = 0; i < el.childNodes.length; i++) {
                    const child = el.childNodes[i] as Element
                    const ct = (child.tagName ?? '').toLowerCase()
                    if (ct === 'td' || ct === 'th') {
                      cells.push((child.textContent ?? '').replace(/\s+/g, ' ').trim())
                    }
                  }
                  return cells.length ? cells.join('\t') + '\n' : ''
                }

                // Table wrappers — recurse with insideTable=true to skip whitespace nodes
                if (TABLE_WRAP.includes(tag)) {
                  let out = ''
                  for (let i = 0; i < el.childNodes.length; i++) out += walk(el.childNodes[i], true)
                  return out
                }

                // Block elements — wrap with newlines
                let out = ''
                for (let i = 0; i < el.childNodes.length; i++) out += walk(el.childNodes[i], false)
                if (BLOCK.includes(tag)) out = '\n' + out.trim() + '\n'
                return out
              }

              const raw = walk(doc.body ?? doc.documentElement)
              const collapsed = raw
                .replace(/[ \t]{2,}/g, ' ')
                .replace(/\n[ \t]+/g, '\n')
                .replace(/\n{3,}/g, '\n\n')
                .trim()

              // Merge consecutive non-table lines into flowing paragraphs,
              // keeping table rows on their own lines and blank lines as separators.
              const lines = collapsed.split('\n')
              const paragraphs: string[] = []
              let buf = ''
              for (const line of lines) {
                if (line.trim() === '') {
                  // Blank line → flush buffer as paragraph
                  if (buf) { paragraphs.push(buf.trim()); buf = '' }
                } else if (line.includes('\t')) {
                  // Table row → flush buffer first, then add row as-is
                  if (buf) { paragraphs.push(buf.trim()); buf = '' }
                  paragraphs.push(line)
                } else {
                  // Flowing text → accumulate into buffer
                  buf = buf ? buf + ' ' + line.trim() : line.trim()
                }
              }
              if (buf) paragraphs.push(buf.trim())
              const cleaned = paragraphs.join('\n')

              // Split body vs signature
              const cutRe = /\n([-_]{3,}|From:\s|Best regards|Regards,|ขอแสดงความนับถือ|Sent:\s)/i
              const cutIdx = cleaned.search(cutRe)
              if (cutIdx > 80) {
                state.emailBodyPreview = cleaned.slice(0, cutIdx).trim().slice(0, 2000)
                const sigText = cleaned.slice(cutIdx).trim()
                state.signatureContact = parseSignature(sigText)
              } else {
                state.emailBodyPreview = cleaned.trim().slice(0, 2000)
                state.signatureContact = null
              }

            }
            render()
          })
          return
        }
      }
      renderDevMode()
      render()
    })
  } else {
    // Running outside Office (plain browser dev)
    renderDevMode()
    render()
  }
}

function renderDevMode(): void {
  state.emailSubject = '[DEV] Test Email Subject'
  state.emailSenderName = 'Test Sender'
  state.emailSenderEmail = 'test@example.com'
  state.emailBodyPreview = 'This is a placeholder email body for development mode.'
}

init().catch(err => {
  console.error('Init error:', err)
  const app = document.getElementById('app')
  if (app) {
    app.innerHTML = `<div class="p-4 text-red-600 text-sm">เกิดข้อผิดพลาด: ${String(err)}</div>`
  }
})
