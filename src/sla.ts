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

/**
 * เส้นตายของเคส — ลำดับพารามิเตอร์ต้องตรงกับ webproject (utils/sla.ts) เป๊ะ
 *
 * ที่ต้องย้ำ: ตัวที่สองคือ "เวลาที่เปิดเคส" ไม่ใช่ "เวลาปัจจุบัน"
 * เดิมที่นี่ตัวที่สองเป็น now — ถ้ามีใครส่งค่าเข้ามาโดยเทียบกับ webapp
 * จะได้เส้นตายผิดแบบเงียบ ๆ (Add-in สร้างเคสใหม่เท่านั้น จึงไม่เคยส่ง)
 */
export function computeSlaDue(
  hours: number | null | undefined,
  createdIso?: string,
  now = new Date(),
): string | null {
  const h = typeof hours === 'number' && Number.isFinite(hours) && hours > 0 ? hours : null
  if (!h) return null
  const start = createdIso ? new Date(createdIso) : now
  const base = isNaN(start.getTime()) ? now : start
  return new Date(base.getTime() + h * 3600000).toISOString()
}

export function slaDueLabel(hours: number | null | undefined): string {
  const due = computeSlaDue(hours)
  return due ? new Date(due).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' }) : ''
}
