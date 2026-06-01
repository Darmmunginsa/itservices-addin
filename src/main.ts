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
type Tab = 'ticket' | 'task' | 'incident'

interface Project { id: number; Title: string }

interface AppState {
  account: AccountInfo | null
  tab: Tab
  emailSubject: string
  emailBodyPreview: string
  emailSenderName: string
  emailSenderEmail: string
  loading: boolean
  projects: Project[]
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

async function fetchProjects(): Promise<void> {
  try {
    const token = await getToken()
    const url = `${SHAREPOINT_URL}/_api/web/lists/getbytitle('PM_Projects')/items?$select=Id,Title&$filter=Status eq 'Active'&$orderby=Title asc`
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json;odata=nometadata',
      },
    })
    if (res.ok) {
      const data = await res.json() as { value: { Id: number; Title: string }[] }
      state.projects = data.value.map(p => ({ id: p.Id, Title: p.Title }))
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
    await fetchProjects()
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
async function spCreate(listTitle: string, body: Record<string, unknown>): Promise<void> {
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

// ─── Today's date (ISO) ───────────────────────────────────────────────────────
function todayISO(): string {
  return new Date().toISOString().split('T')[0]
}

// ─── Submit handlers ──────────────────────────────────────────────────────────
async function handleSubmit(): Promise<void> {
  if (!state.account) {
    showToast('กรุณาเข้าสู่ระบบก่อน', 'error')
    return
  }

  const btn = document.getElementById('submit-btn') as HTMLButtonElement | null
  if (btn) { btn.disabled = true; btn.textContent = 'กำลังบันทึก…' }

  try {
    if (state.tab === 'ticket') {
      const title = (document.getElementById('f-title') as HTMLInputElement).value.trim()
      const description = (document.getElementById('f-description') as HTMLTextAreaElement).value.trim()
      const priority = (document.getElementById('f-priority') as HTMLSelectElement).value
      const customerEmail = (document.getElementById('f-customer-email') as HTMLInputElement).value.trim()

      await spCreate('HD_Tickets', {
        Title: title,
        Description: description,
        Priority: priority,
        CustomerEmail: customerEmail,
        CustomerName: state.emailSenderName || customerEmail,
        Status: 'Open',
      })
      showToast('สร้าง Ticket สำเร็จ!')

    } else if (state.tab === 'task') {
      const title = (document.getElementById('f-title') as HTMLInputElement).value.trim()
      const projectId = parseInt((document.getElementById('f-project') as HTMLSelectElement)?.value || '0')
      const dueDate = (document.getElementById('f-due-date') as HTMLInputElement).value
      const note = (document.getElementById('f-note') as HTMLTextAreaElement).value.trim()

      if (!projectId) { showToast('กรุณาเลือก Project', 'error'); return }

      await spCreate('PM_Tasks', {
        Title: title,
        DueDate: dueDate || null,
        TaskNote: note,
        AssignedTo: state.account.name ?? state.account.username,
        AssignedEmail: state.account.username,
        IsCompleted: false,
        IsAcknowledged: false,
        ProjectID: projectId,
      })
      showToast('สร้าง Task สำเร็จ!')

    } else if (state.tab === 'incident') {
      const title = (document.getElementById('f-title') as HTMLInputElement).value.trim()
      const projectId = parseInt((document.getElementById('f-project') as HTMLSelectElement)?.value || '0')
      const description = (document.getElementById('f-description') as HTMLTextAreaElement).value.trim()
      const severity = (document.getElementById('f-severity') as HTMLSelectElement).value

      if (!projectId) { showToast('กรุณาเลือก Project', 'error'); return }

      await spCreate('PM_Incidents', {
        Title: title,
        Description: description,
        Severity: severity,
        Status: 'Open',
        AssignedEmail: state.account.username,
        ProjectID: projectId,
        IncidentDate: todayISO(),
      })
      showToast('สร้าง Incident สำเร็จ!')
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    showToast(`เกิดข้อผิดพลาด: ${msg}`, 'error')
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = 'บันทึก' }
  }
}

// ─── Tab labels ───────────────────────────────────────────────────────────────
const TAB_META: Record<Tab, { label: string; icon: string }> = {
  ticket:   { label: 'Ticket',   icon: '🎫' },
  task:     { label: 'Task',     icon: '✅' },
  incident: { label: 'Incident', icon: '🚨' },
}

// ─── Render ───────────────────────────────────────────────────────────────────
function render(): void {
  const app = document.getElementById('app')
  if (!app) return

  const { account, tab, emailSubject, emailSenderName, emailSenderEmail, emailBodyPreview } = state
  const isLoggedIn = account !== null

  // ── Header ──
  const headerHTML = `
    <div class="bg-blue-700 text-white px-4 py-3 flex items-center justify-between shadow">
      <div class="flex items-center gap-2">
        <div class="w-7 h-7 bg-white rounded flex items-center justify-center">
          <span class="text-blue-700 font-bold text-xs">iT</span>
        </div>
        <span class="font-semibold text-sm tracking-wide">iT Services Helpdesk</span>
      </div>
      ${isLoggedIn
        ? `<button id="btn-logout" class="text-xs bg-blue-600 hover:bg-blue-500 px-2 py-1 rounded transition">
             ออกจากระบบ
           </button>`
        : `<button id="btn-login" class="text-xs bg-white text-blue-700 font-semibold hover:bg-blue-50 px-3 py-1 rounded transition">
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

  // ── Tab switcher ──
  const tabsHTML = `
    <div class="mx-3 mt-3 flex rounded-lg overflow-hidden border border-slate-200 bg-white shadow-sm">
      ${(Object.entries(TAB_META) as [Tab, { label: string; icon: string }][]).map(([key, meta]) => `
        <button
          data-tab="${key}"
          class="tab-btn flex-1 py-2 text-xs font-medium transition
            ${tab === key
              ? 'bg-blue-700 text-white'
              : 'text-slate-600 hover:bg-slate-100'}"
        ><span class="hidden sm:inline">${meta.icon} </span>${meta.label}</button>
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
    `
  } else if (tab === 'task') {
    formHTML = `
      ${field('Title / หัวข้อ', `<input id="f-title" type="text"
        class="${inputCls}"
        value="${esc(emailSubject)}" />`)}
      ${field('Project *', projectSelect())}
      ${field('Due Date', `<input id="f-due-date" type="date"
        class="${inputCls}"
        value="" />`)}
      ${field('หมายเหตุ', `<textarea id="f-note" rows="6"
        class="${inputCls} resize-y">${esc(emailBodyPreview)}</textarea>`)}
    `
  } else if (tab === 'incident') {
    formHTML = `
      ${field('Title / หัวข้อ', `<input id="f-title" type="text"
        class="${inputCls}"
        value="${esc(emailSubject)}" />`)}
      ${field('Project *', projectSelect())}
      ${field('รายละเอียด', `<textarea id="f-description" rows="3"
        class="${inputCls} resize-none">${esc(emailBodyPreview)}</textarea>`)}
      ${field('Severity', `<select id="f-severity" class="${inputCls}">
        <option value="Low">Low</option>
        <option value="Medium" selected>Medium</option>
        <option value="High">High</option>
        <option value="Critical">Critical</option>
      </select>`)}
    `
  }

  // ── Account pill ──
  const accountHTML = `
    <div class="mx-3 mt-3 flex items-center gap-2 bg-green-50 border border-green-200 rounded-lg px-3 py-2">
      <div class="w-6 h-6 rounded-full bg-green-600 flex items-center justify-center text-white text-xs font-bold flex-shrink-0">
        ${(account.name ?? account.username).charAt(0).toUpperCase()}
      </div>
      <div class="min-w-0">
        <div class="text-xs font-semibold text-slate-700 truncate">${esc(account.name ?? '')}</div>
        <div class="text-xs text-slate-400 truncate">${esc(account.username)}</div>
      </div>
    </div>
  `

  app.innerHTML = `
    ${headerHTML}
    ${accountHTML}
    ${emailInfoHTML}
    ${tabsHTML}
    <div class="mx-3 mt-3 space-y-3">
      ${formHTML}
      <button id="submit-btn"
        class="w-full bg-blue-700 hover:bg-blue-600 active:bg-blue-800 text-white text-sm font-semibold py-2.5 rounded-lg shadow transition mt-1">
        บันทึก
      </button>
    </div>
    <div class="h-6"></div>
  `

  // ── Event listeners ──
  document.getElementById('btn-logout')?.addEventListener('click', logout)
  document.getElementById('submit-btn')?.addEventListener('click', handleSubmit)

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const newTab = (btn as HTMLElement).dataset['tab'] as Tab
      if (newTab && newTab !== state.tab) {
        state.tab = newTab
        render()
      }
    })
  })
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
const inputCls = 'w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white'

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
      await fetchProjects()
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

          // Body preview (async)
          item.body.getAsync(Office.CoercionType.Text, { asyncContext: {} }, result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const full = result.value as string
              state.emailBodyPreview = full.slice(0, 500).trim()
            }
            render()
          })
          return // render will be called in body callback
        }
      }
      // Not in Outlook or no item — render with whatever we have (dev mode)
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
