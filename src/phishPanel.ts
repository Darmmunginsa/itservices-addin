// ─── PhishGuard panel ───────────────────────────────────────────────────────
// แท็บตรวจอีเมลหลอกลวง ที่ฝังอยู่ในแอดอิน Helpdesk เดิม
// อยู่ใน UI ของแอป ไม่ได้อยู่ใน manifest → เพิ่มได้โดยไม่ต้องติดตั้ง/แก้ deployment ใด ๆ
import { analyze, parseRawHeaders, LEVEL_META, SEV_META, domainOf, type Analysis, type MailInput } from './analyzer'

export interface PhishDeps {
  sharepointUrl: string
  internalDomains: string[]
  /** token ของ SharePoint (จาก main.ts) */
  getToken: () => Promise<string>
  /** token ของ Graph (จาก main.ts) */
  getGraphToken: () => Promise<string>
  /** ผู้ใช้ที่ล็อกอินอยู่ — ไม่มี = ยังไม่ได้ล็อกอิน */
  account: () => { name?: string; username?: string } | null
  toast: (msg: string, type?: 'success' | 'error' | 'info') => void
  /** ให้ main.ts วาดหน้าใหม่ */
  rerender: () => void
}

const REPORT_LIST = 'HD_PhishingReports'
let deps: PhishDeps

interface PhishState {
  mail: MailInput | null
  analysis: Analysis | null
  loading: boolean
  reporting: boolean
  reported: boolean
  showHeaders: boolean
  kasmTemplate: string
  /** itemId ที่วิเคราะห์ไปแล้ว — กันวิเคราะห์ซ้ำทุกครั้งที่ render */
  analysedItemId: string
}
const ps: PhishState = {
  mail: null, analysis: null, loading: false, reporting: false,
  reported: false, showHeaders: false, kasmTemplate: '', analysedItemId: '',
}

export function initPhish(d: PhishDeps): void { deps = d }

const esc = (s: string): string =>
  (s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;')

const mailboxItem = (): Office.MessageRead | undefined =>
  (Office.context?.mailbox?.item as Office.MessageRead | undefined)

const getBodyAsync = (type: Office.CoercionType): Promise<string> =>
  new Promise(resolve => {
    const it = mailboxItem()
    if (!it?.body) { resolve(''); return }
    it.body.getAsync(type, r => resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value ?? '') : ''))
  })

function parseAddressList(raw: string): { name: string; email: string }[] {
  if (!raw) return []
  return raw.split(',').map(part => {
    const s = part.trim()
    const m = s.match(/^(.*?)\s*<([^>]+)>$/)
    if (m) return { name: m[1].replace(/^"|"$/g, '').trim(), email: m[2].trim() }
    return { name: '', email: s.replace(/[<>]/g, '').trim() }
  }).filter(a => a.email.includes('@'))
}

/** header ตรงจาก Office.js (set 1.8) — ไม่ต้อง login */
function headersViaOfficeJs(): Promise<Record<string, string>> {
  return new Promise(resolve => {
    let supported = false
    try { supported = Office.context.requirements.isSetSupported('Mailbox', '1.8') } catch { supported = false }
    const it = mailboxItem() as (Office.MessageRead & {
      getAllInternetHeadersAsync?: (cb: (r: Office.AsyncResult<string>) => void) => void
    }) | undefined
    if (!supported || typeof it?.getAllInternetHeadersAsync !== 'function') { resolve({}); return }
    try {
      it.getAllInternetHeadersAsync(r =>
        resolve(r.status === Office.AsyncResultStatus.Succeeded ? parseRawHeaders(r.value ?? '') : {}))
    } catch { resolve({}) }
  })
}

/** header ผ่าน Graph — ใช้เมื่อ Office.js อ่านไม่ได้ (client เก่ากว่า 1.8) */
async function headersViaGraph(): Promise<Record<string, string>> {
  try {
    const item = mailboxItem()
    if (!item?.itemId) return {}
    const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0)
    const token = await deps.getGraphToken()
    const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${restId}?$select=internetMessageHeaders`,
      { headers: { Authorization: `Bearer ${token}` } })
    if (!res.ok) return {}
    const data = await res.json() as { internetMessageHeaders?: { name: string; value: string }[] }
    const out: Record<string, string> = {}
    for (const h of data.internetMessageHeaders ?? []) {
      out[h.name] = out[h.name] ? `${out[h.name]}\n${h.value}` : h.value
    }
    return out
  } catch { return {} }
}

async function spList<T>(list: string, query: string): Promise<T[]> {
  const token = await deps.getToken()
  const res = await fetch(`${deps.sharepointUrl}/_api/web/lists/getbytitle('${list}')/items?${query}`,
    { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata' } })
  if (!res.ok) return []
  return ((await res.json()) as { value: T[] }).value
}

/** รายชื่อพนักงาน — ใช้ตรวจการปลอมเป็นคนใน */
async function fetchInternalPeople(): Promise<{ name: string; email: string }[]> {
  try {
    const rows = await spList<{ Title: string; EmailText?: string }>('HD_AgentProfiles', '$select=Title,EmailText&$top=500')
    return rows.filter(p => p.EmailText).map(p => ({ name: p.Title, email: p.EmailText! }))
  } catch { return [] }
}

async function fetchKasmTemplate(): Promise<string> {
  try {
    const rows = await spList<{ Title?: string }>('HD_Options', "$select=Title&$filter=Category eq 'KasmConfig'&$top=1")
    return (rows[0]?.Title ?? '').trim()
  } catch { return '' }
}

/** วิเคราะห์อีเมลที่เปิดอยู่ (ข้ามถ้าวิเคราะห์ใบนี้ไปแล้ว) */
export async function analysePhish(force = false): Promise<void> {
  const item = mailboxItem()
  const id = item?.itemId ?? ''
  if (!force && ps.analysedItemId === id && ps.analysis) return

  ps.loading = true; ps.reported = false; ps.analysedItemId = id
  deps.rerender()

  const [html, text] = await Promise.all([
    getBodyAsync(Office.CoercionType.Html),
    getBodyAsync(Office.CoercionType.Text),
  ])
  const base: MailInput = {
    fromName: item?.from?.displayName ?? '',
    fromEmail: item?.from?.emailAddress ?? '',
    replyTo: [],
    subject: item?.subject ?? '',
    bodyHtml: html,
    bodyText: text,
    attachments: (item?.attachments ?? []).map(a => ({ name: a.name, size: a.size ?? 0, isInline: !!a.isInline })),
    headers: {},
    internalDomains: deps.internalDomains,
    internalPeople: [],
  }
  const withHeaders = (h: Record<string, string>): MailInput => ({
    ...base,
    headers: h,
    replyTo: parseAddressList(Object.entries(h).find(([k]) => k.toLowerCase() === 'reply-to')?.[1] ?? ''),
  })

  // รอบแรก: header จาก Office.js — เห็นผลเร็ว ไม่ต้อง login
  const ojs = await headersViaOfficeJs()
  ps.mail = withHeaders(ojs)
  ps.analysis = analyze(ps.mail)
  ps.loading = false
  deps.rerender()

  // รอบสอง (ถ้า login แล้ว): เติมรายชื่อพนักงาน + header ผ่าน Graph ถ้ารอบแรกอ่านไม่ได้
  if (deps.account()) {
    const [people, gh] = await Promise.all([
      fetchInternalPeople(),
      Object.keys(ojs).length ? Promise.resolve(ojs) : headersViaGraph(),
    ])
    ps.mail = { ...withHeaders(gh), internalPeople: people }
    ps.analysis = analyze(ps.mail)
    if (!ps.kasmTemplate) ps.kasmTemplate = await fetchKasmTemplate()
    deps.rerender()
  }
}

function reportText(): string {
  const m = ps.mail, a = ps.analysis
  if (!m || !a) return ''
  return [
    `ผู้ส่ง: ${m.fromName} <${m.fromEmail}>`,
    `หัวข้อ: ${m.subject}`,
    m.replyTo.length ? `Reply-To: ${m.replyTo.map(r => r.email).join(', ')}` : '',
    `คะแนนความเสี่ยง: ${a.score} (${LEVEL_META[a.level].label})`,
    '',
    'สิ่งที่ตรวจพบ:',
    ...a.findings.map(f => `- [${SEV_META[f.severity].label}] (${f.category}) ${f.title} — ${f.detail.replace(/\n/g, ' ')}`),
    '',
    a.links.length ? 'ลิงก์ในอีเมล:' : '',
    ...a.links.map(l => `- ${l.href}${l.flags.length ? `  ! ${l.flags.join(' / ')}` : ''}`),
  ].filter(l => l !== '').join('\n')
}

async function attachEml(itemId: number): Promise<boolean> {
  try {
    const item = mailboxItem()
    if (!item?.itemId) return false
    const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0)
    const gToken = await deps.getGraphToken()
    const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${restId}/$value`,
      { headers: { Authorization: `Bearer ${gToken}` } })
    if (!res.ok) return false
    const mime = await res.arrayBuffer()
    const safe = (item.subject || 'phishing')
      .replace(/[\\/:*?"<>|#%&{}~]/g, '_').replace(/^_+/, '').slice(0, 80).trim() || 'phishing'
    const token = await deps.getToken()
    const up = await fetch(
      `${deps.sharepointUrl}/_api/web/lists/getbytitle('${REPORT_LIST}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(safe + '.eml')}')`,
      { method: 'POST', headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata' }, body: mime })
    return up.ok
  } catch { return false }
}

async function spCreate(body: Record<string, unknown>): Promise<number> {
  const token = await deps.getToken()
  const res = await fetch(`${deps.sharepointUrl}/_api/web/lists/getbytitle('${REPORT_LIST}')/items`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
    },
    body: JSON.stringify(body),
  })
  if (!res.ok) throw new Error(`SharePoint ${res.status}: ${await res.text()}`)
  return ((await res.json()) as { Id: number }).Id
}

/** ปุ่มล่างของแอดอินเรียกตัวนี้เมื่ออยู่แท็บ PhishGuard */
export async function submitPhishReport(): Promise<void> {
  if (!ps.mail || !ps.analysis || ps.reporting) return
  if (!deps.account()) { deps.toast('กรุณาเข้าสู่ระบบก่อนรายงาน', 'error'); return }
  ps.reporting = true; deps.rerender()
  try {
    const m = ps.mail, a = ps.analysis
    const acc = deps.account()
    const full: Record<string, unknown> = {
      Title: (m.subject || '(ไม่มีหัวข้อ)').slice(0, 255),
      SenderName: m.fromName.slice(0, 255),
      SenderEmail: m.fromEmail.slice(0, 255),
      SenderDomain: domainOf(m.fromEmail),
      RiskScore: a.score,
      RiskLevel: a.level,
      Findings: reportText(),
      LinkCount: a.links.length,
      SuspiciousLinks: a.links.filter(l => l.flags.length).map(l => l.href).join('\n').slice(0, 4000),
      ReportedBy: acc?.name ?? '',
      ReportedEmail: acc?.username ?? '',
      Status: 'New',
    }
    // รายงานภัยคุกคามต้องไม่หายเพราะคอลัมน์ใดคอลัมน์หนึ่งชื่อไม่ตรง → มีชุดสำรอง
    let id: number, partial = false
    try {
      id = await spCreate(full)
    } catch (firstErr) {
      id = await spCreate({ Title: full.Title, Findings: full.Findings }).catch(() => { throw firstErr })
      partial = true
    }
    const withEml = await attachEml(id)
    ps.reported = true
    deps.toast(
      partial ? 'ส่งรายงานแล้ว แต่บันทึกได้บางคอลัมน์ — ตรวจชื่อคอลัมน์ใน HD_PhishingReports'
        : withEml ? 'ส่งรายงานพร้อมอีเมลต้นฉบับแล้ว' : 'ส่งรายงานแล้ว (แนบ .eml ไม่ได้)',
      partial ? 'info' : 'success')
  } catch (e) {
    deps.toast(`ส่งรายงานไม่สำเร็จ: ${e instanceof Error ? e.message : String(e)}`, 'error')
  } finally {
    ps.reporting = false; deps.rerender()
  }
}

export const phishSubmitLabel = (): string =>
  ps.reported ? '✓ รายงานแล้ว' : ps.reporting ? 'กำลังส่ง…' : '🚩 รายงานอีเมลนี้ให้ IT'

export const phishSubmitDisabled = (): boolean => ps.reporting || ps.reported || !ps.analysis

/** Outlook desktop (WebView2) บล็อก window.open → ใช้ API ของ Office ก่อน */
function openExternal(url: string): void {
  try {
    const ui = Office.context?.ui as { openBrowserWindow?: (u: string) => void } | undefined
    if (typeof ui?.openBrowserWindow === 'function') { ui.openBrowserWindow(url); return }
  } catch { /* ลองทางถัดไป */ }
  if (!window.open(url, '_blank', 'noopener,noreferrer')) {
    deps.toast('เปิดหน้าต่างไม่ได้ (ถูกบล็อก)', 'info')
  }
}

async function copyText(text: string): Promise<boolean> {
  try { if (navigator.clipboard?.writeText) { await navigator.clipboard.writeText(text); return true } } catch { /* ลองต่อ */ }
  try {
    const ta = document.createElement('textarea')
    ta.value = text; ta.setAttribute('readonly', ''); ta.style.position = 'fixed'; ta.style.opacity = '0'
    document.body.appendChild(ta); ta.focus(); ta.select()
    const ok = document.execCommand('copy'); ta.remove(); return ok
  } catch { return false }
}

/** HTML ของแท็บ (main.ts เอาไปใส่ใน formHTML) */
export function phishPanelHTML(loggedIn: boolean): string {
  if (ps.loading && !ps.analysis) {
    return `<div class="py-10 text-center text-slate-500 text-sm">
      <div class="w-7 h-7 border-[3px] border-slate-200 border-t-blue-600 rounded-full animate-spin mx-auto mb-2"></div>
      กำลังตรวจอีเมล…</div>`
  }
  const a = ps.analysis, m = ps.mail
  if (!a || !m) return `<p class="text-sm text-slate-400 text-center py-8">เปิดอีเมลเพื่อเริ่มตรวจ</p>`

  const lv = LEVEL_META[a.level]
  const links = [...a.links.filter(l => l.flags.length), ...a.links.filter(l => !l.flags.length)]

  return `
    <div class="rounded-xl border-2 ${lv.cls} p-3">
      <div class="flex items-center gap-2">
        <span class="text-2xl leading-none">${lv.icon}</span>
        <div class="min-w-0 flex-1">
          <div class="font-bold text-sm">${esc(lv.label)}</div>
          <div class="text-xs opacity-80">คะแนน ${a.score} · พบสัญญาณ ${a.findings.filter(f => f.severity !== 'info').length} ข้อ</div>
        </div>
        <button id="phish-recheck" class="text-[10px] px-2 py-1 rounded-md bg-white/70 hover:bg-white text-slate-700 font-medium">ตรวจใหม่</button>
      </div>
    </div>

    ${!loggedIn ? `<div class="bg-blue-50 border border-blue-200 rounded-xl p-2.5 text-[11px] text-blue-800">
      เข้าสู่ระบบเพื่อตรวจการปลอมเป็นพนักงาน และรายงานเข้า Helpdesk ได้
    </div>` : ''}

    ${a.findings.length === 0
      ? `<p class="text-xs text-slate-400 text-center py-3">ไม่พบสัญญาณผิดปกติจากการตรวจอัตโนมัติ</p>`
      : a.findings.map(f => {
        const s = SEV_META[f.severity]
        return `<div class="bg-white rounded-xl border border-slate-200 p-2.5">
          <div class="flex items-start gap-2">
            <span class="text-[9px] font-bold px-1.5 py-0.5 rounded-full ${s.cls} flex-shrink-0 mt-0.5">${s.label}</span>
            <div class="min-w-0 flex-1">
              <div class="text-xs font-semibold text-slate-800">${esc(f.title)}</div>
              <div class="text-[11px] text-slate-500 whitespace-pre-line break-all">${esc(f.detail)}</div>
              <div class="text-[9px] text-slate-400 mt-0.5">${esc(f.category)}</div>
            </div>
          </div>
        </div>`
      }).join('')}

    ${links.length ? `
    <div class="bg-white rounded-xl border border-slate-200 p-3">
      <div class="text-xs font-semibold text-slate-700 mb-2">ลิงก์ในอีเมล (${links.length})</div>
      <div class="space-y-2">
        ${links.map((l, i) => `
          <div class="rounded-lg border ${l.flags.length ? 'border-red-200 bg-red-50/50' : 'border-slate-100'} p-2">
            <div class="text-[11px] font-medium ${l.flags.length ? 'text-red-700' : 'text-slate-700'} break-all">${esc(l.host || l.href)}</div>
            ${l.text && l.text !== l.href ? `<div class="text-[10px] text-slate-500 break-all">แสดงว่า: "${esc(l.text.slice(0, 70))}"</div>` : ''}
            <div class="text-[10px] text-slate-400 break-all mt-0.5">${esc(l.href.slice(0, 150))}</div>
            ${l.flags.map(f => `<div class="text-[10px] text-red-600 mt-0.5">! ${esc(f)}</div>`).join('')}
            <button data-kasm="${i}" class="mt-1.5 text-[10px] font-semibold px-2 py-1 rounded-md bg-slate-800 hover:bg-slate-900 text-white">
              เปิดใน Kasm (แซนด์บ็อกซ์)
            </button>
          </div>`).join('')}
      </div>
      <p class="text-[10px] text-slate-400 mt-2">อย่าคลิกลิงก์จากอีเมลที่ไม่มั่นใจโดยตรง</p>
    </div>` : ''}

    <div class="flex gap-2">
      <button id="phish-copy" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">คัดลอกผลตรวจ</button>
      <button id="phish-headers" class="flex-1 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium py-1.5 rounded-lg">${ps.showHeaders ? 'ซ่อน header' : 'ดู header'}</button>
    </div>
    ${ps.showHeaders ? `<pre class="bg-slate-900 text-slate-100 text-[9px] p-2 rounded-lg overflow-x-auto whitespace-pre-wrap break-all max-h-56 overflow-y-auto">${
      esc(Object.entries(m.headers).map(([k, v]) => `${k}: ${v}`).join('\n') || 'อ่าน header ไม่ได้')
    }</pre>` : ''}
    <p class="text-[10px] text-slate-400 text-center">วิเคราะห์ในเครื่องทั้งหมด — ไม่ส่งเนื้อหาอีเมลออกนอกองค์กร</p>
  `
}

/** ผูก event หลัง main.ts เขียน innerHTML แล้ว */
export function bindPhishPanel(): void {
  document.getElementById('phish-recheck')?.addEventListener('click', () => { analysePhish(true) })
  document.getElementById('phish-headers')?.addEventListener('click', () => { ps.showHeaders = !ps.showHeaders; deps.rerender() })
  document.getElementById('phish-copy')?.addEventListener('click', async () => {
    const ok = await copyText(reportText())
    deps.toast(ok ? 'คัดลอกผลตรวจแล้ว' : 'คัดลอกไม่ได้', ok ? 'success' : 'error')
  })
  const links = [...(ps.analysis?.links.filter(l => l.flags.length) ?? []), ...(ps.analysis?.links.filter(l => !l.flags.length) ?? [])]
  document.querySelectorAll('[data-kasm]').forEach(btn => {
    btn.addEventListener('click', () => {
      const l = links[Number((btn as HTMLElement).dataset['kasm'])]
      if (!l) return
      if (!ps.kasmTemplate) { deps.toast('ยังไม่ได้ตั้งค่า Kasm ใน HD_Options (Category=KasmConfig)', 'info'); return }
      const t = ps.kasmTemplate
      openExternal(t.includes('{url}') ? t.replace('{url}', encodeURIComponent(l.href)) : t + encodeURIComponent(l.href))
    })
  })
}
