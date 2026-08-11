// SLA ของ Incident — ต้องตรงกับฝั่ง Helpdesk webapp (src/utils/sla.ts)
// วัด SLA ที่ Incident เท่านั้น ; Ticket คือ "คำขอให้ทำบางอย่าง" ไม่ใช่ปัญหา จึงไม่มี SLA

export const SLA_OPTIONS: { hours: number; labelTh: string }[] = [
  { hours: 1,   labelTh: '1 ชั่วโมง' },
  { hours: 2,   labelTh: '2 ชั่วโมง' },
  { hours: 4,   labelTh: '4 ชั่วโมง' },
  { hours: 8,   labelTh: '8 ชั่วโมง (1 วันทำการ)' },
  { hours: 24,  labelTh: '24 ชั่วโมง' },
  { hours: 48,  labelTh: '2 วัน' },
  { hours: 72,  labelTh: '3 วัน' },
  { hours: 168, labelTh: '7 วัน' },
]

/** SLA ที่แนะนำตามความรุนแรง — เป็นค่าตั้งต้น ผู้ใช้เปลี่ยนได้ */
export const SLA_BY_SEVERITY: Record<string, number> = {
  Critical: 1, High: 4, Medium: 24, Low: 72,
}

/** เส้นตายของเคสที่เพิ่งเปิด — นับจากตอนนี้ */
export function computeSlaDue(hours: number | null | undefined, now = new Date()): string | null {
  const h = typeof hours === 'number' && Number.isFinite(hours) && hours > 0 ? hours : null
  return h ? new Date(now.getTime() + h * 3600000).toISOString() : null
}

export function slaDueLabel(hours: number | null | undefined): string {
  const due = computeSlaDue(hours)
  return due ? new Date(due).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' }) : ''
}
