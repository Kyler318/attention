<script setup>
import { ref, computed, onMounted } from 'vue'
import * as XLSX from 'xlsx'

const ATTENDANCE_OPTIONS = [
  { label: '準時出席',   value: 100, cls: 'opt-present'     },
  { label: '遲到',       value: 90,  cls: 'opt-late'         },
  { label: '第二堂準時', value: 50,  cls: 'opt-second'       },
  { label: '第二堂遲到', value: 40,  cls: 'opt-second-late'  },
  { label: '缺席',       value: 0,   cls: 'opt-absent'       },
]

const workbook       = ref(null)
const sheetNames     = ref([])
const selectedClass  = ref('')
const students       = ref([])
const currentDate    = ref(new Date().toISOString().split('T')[0])
const records        = ref({})     // { name: { attendance, behavior, classScore } }
const sessions       = ref({})     // { className: [ ...session ] }
const activeTab      = ref('attendance')
const loading        = ref(true)
const loadError      = ref('')
const detailSession  = ref(null)
const sheetStructure = ref(null)
const editingEntry   = ref(null)   // { idx } — session being edited

// ── persist: server + localStorage fallback ───────────────────────────────────

async function loadStorage() {
  try {
    const res = await fetch('/api/sessions')
    if (res.ok) { sessions.value = await res.json(); return }
  } catch {}
  try {
    const raw = localStorage.getItem('attendance_app_v1')
    if (raw) sessions.value = JSON.parse(raw)
  } catch {}
}

async function saveStorage() {
  const data = JSON.stringify(sessions.value)
  try {
    await fetch('/api/sessions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: data,
    })
  } catch {}
  localStorage.setItem('attendance_app_v1', data)
}

// ── Excel load ────────────────────────────────────────────────────────────────

onMounted(async () => {
  await loadStorage()
  await loadWorkbook()
})

async function loadWorkbook() {
  try {
    loading.value = true
    const res = await fetch('/25-26下普中各班分紙3-試用.xls')
    if (!res.ok) throw new Error(`HTTP ${res.status}`)
    const buf = await res.arrayBuffer()
    workbook.value = XLSX.read(new Uint8Array(buf), { type: 'array', cellStyles: true })
    sheetNames.value = workbook.value.SheetNames
    if (sheetNames.value.length) {
      selectedClass.value = sheetNames.value[0]
      loadStudents()
    }
  } catch (e) {
    loadError.value = `無法讀取 Excel：${e.message}`
  } finally {
    loading.value = false
  }
}

// ── student extraction ────────────────────────────────────────────────────────

function loadStudents() {
  if (!workbook.value || !selectedClass.value) return
  students.value = extractStudents(selectedClass.value)
  resetForm()
}

function extractStudents(sheetName) {
  const sheet = workbook.value.Sheets[sheetName]
  const rows  = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })
  const out   = []
  for (const row of rows) {
    const c0 = String(row[0] ?? '').trim()
    const c1 = String(row[1] ?? '').trim()
    const c2 = String(row[2] ?? '').trim()
    if (/^\d+$/.test(c0) && /[一-龥]{2,}/.test(c1)) { out.push({ num: c0, name: c1 }); continue }
    if (/^\d+$/.test(c0) && /[一-龥]{2,}/.test(c2)) { out.push({ num: c0, name: c2 }); continue }
  }
  return out
}

function resetForm() {
  const r = {}
  students.value.forEach(s => { r[s.name] = { attendance: null, behavior: null, classScore: null } })
  records.value = r
  editingEntry.value = null
}

// ── class select ──────────────────────────────────────────────────────────────

function selectClass(name) {
  selectedClass.value = name
  loadStudents()
  activeTab.value = 'attendance'
}

// ── attendance logic ──────────────────────────────────────────────────────────

function onAttendanceChange(name) {
  const rec = records.value[name]
  if (rec.behavior !== null && rec.behavior > rec.attendance) rec.behavior = rec.attendance
}

function validateBehavior(name) {
  const rec = records.value[name]
  const max = rec.attendance ?? 0
  if (rec.behavior !== null) {
    if (rec.behavior > max) rec.behavior = max
    if (rec.behavior < 0)   rec.behavior = 0
  }
}

function getRowCls(name) {
  const att = records.value[name]?.attendance
  if (att === null || att === undefined) return ''
  const map = { 100: 'row-present', 90: 'row-late', 50: 'row-second', 40: 'row-second-late', 0: 'row-absent' }
  return map[att] ?? ''
}

// ── quick-fill ────────────────────────────────────────────────────────────────

function markAllPresent() {
  students.value.forEach(s => {
    records.value[s.name].attendance = 100
    onAttendanceChange(s.name)
  })
}

function fillAllBehavior() {
  students.value.forEach(s => {
    const rec = records.value[s.name]
    if (rec.attendance !== null) rec.behavior = rec.attendance
  })
}

function fillAllClassScore() {
  students.value.forEach(s => {
    const rec = records.value[s.name]
    if (rec.attendance !== null) rec.classScore = rec.attendance
  })
}

// ── submit / edit ─────────────────────────────────────────────────────────────

function startEdit(sess, idx) {
  editingEntry.value = { idx }
  currentDate.value  = sess.date
  const r = {}
  sess.records.forEach(rec => {
    r[rec.name] = {
      attendance: rec.attendance,
      behavior:   rec.behavior   ?? null,
      classScore: rec.classScore ?? null,
    }
  })
  records.value = r
  activeTab.value = 'attendance'
}

function cancelEdit() {
  resetForm()
}

function submitSession() {
  // Only attendance is mandatory; behavior & classScore are optional
  const unmarked = students.value.filter(s => records.value[s.name]?.attendance === null)
  if (unmarked.length) {
    alert(`尚未填寫出席：${unmarked.map(s => s.name).join('、')}`)
    return
  }
  const session = {
    id: editingEntry.value
      ? curSessions.value[editingEntry.value.idx].id
      : Date.now(),
    date: currentDate.value,
    records: students.value.map(s => ({
      num:        s.num,
      name:       s.name,
      attendance: records.value[s.name].attendance,
      behavior:   records.value[s.name].behavior   ?? 0,
      classScore: records.value[s.name].classScore ?? 0,
    }))
  }
  if (!sessions.value[selectedClass.value]) sessions.value[selectedClass.value] = []
  if (editingEntry.value) {
    sessions.value[selectedClass.value].splice(editingEntry.value.idx, 1, session)
  } else {
    sessions.value[selectedClass.value].push(session)
  }
  saveStorage()
  resetForm()
  activeTab.value = 'history'
}

// ── computed ──────────────────────────────────────────────────────────────────

const curSessions  = computed(() => sessions.value[selectedClass.value] ?? [])
const sessionCount = computed(() => curSessions.value.length)

const markedCount = computed(() =>
  students.value.filter(s => records.value[s.name]?.attendance !== null).length
)

const batchProgress = computed(() => {
  const n = sessionCount.value % 7
  return n === 0 && sessionCount.value > 0 ? 7 : n
})

const progressPct  = computed(() => `${(batchProgress.value / 7) * 100}%`)
const currentBatch = computed(() => Math.ceil(sessionCount.value / 7) || 1)

function getSessionCount(cls) {
  return sessions.value[cls]?.length ?? 0
}

// ── history helpers ───────────────────────────────────────────────────────────

function countAtt(session, val) {
  return session.records.filter(r => r.attendance === val).length
}

function attLabel(val) {
  return ATTENDANCE_OPTIONS.find(o => o.value === val)?.label ?? '—'
}

function deleteSession(idx) {
  if (!confirm('確定刪除此堂記錄？')) return
  sessions.value[selectedClass.value].splice(idx, 1)
  saveStorage()
}

// ── export helpers ────────────────────────────────────────────────────────────

function formatDate(iso) {
  if (!iso) return ''
  const p = iso.split('-')
  return p.length === 3 ? `${parseInt(p[2])}/${parseInt(p[1])}` : iso
}

function inspectStructure() {
  if (!workbook.value || !selectedClass.value) return
  const ws   = workbook.value.Sheets[selectedClass.value]
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false })
  sheetStructure.value = rows.slice(0, 12).map((row, r) => ({ r, cells: row.slice(0, 20) }))
}

// ── fillSheet: write one batch into a worksheet ───────────────────────────────

function fillSheet(ws, batch) {
  const range  = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1')
  const maxRow = range.e.r
  const maxCol = range.e.c

  const getVal = (r, c) => {
    const cell = ws[XLSX.utils.encode_cell({ r, c })]
    return cell ? String(cell.v ?? '').trim() : ''
  }
  const setVal = (r, c, v) => {
    const addr     = XLSX.utils.encode_cell({ r, c })
    const existing = ws[addr]
    if (existing?.f) return                   // preserve formula cells
    const cell = typeof v === 'number' ? { t: 'n', v } : { t: 's', v: String(v) }
    if (existing?.s) cell.s = existing.s      // carry original border/style
    ws[addr] = cell
    if (r > range.e.r) { range.e.r = r; ws['!ref'] = XLSX.utils.encode_range(range) }
    if (c > range.e.c) { range.e.c = c; ws['!ref'] = XLSX.utils.encode_range(range) }
  }

  // 1. Find score-header row (contains 出席 / 表現 / 堂課)
  let scoreRow = -1
  for (let r = 0; r <= Math.min(maxRow, 20); r++) {
    for (let c = 0; c <= maxCol; c++) {
      if (/出席|表現|堂課/.test(getVal(r, c))) { scoreRow = r; break }
    }
    if (scoreRow !== -1) break
  }
  if (scoreRow === -1) return { ok: false, scoreRow: 0 }

  // 2. Map score columns (repeated once per session)
  const attCols = [], behCols = [], clsCols = []
  for (let c = 0; c <= maxCol; c++) {
    const v = getVal(scoreRow, c)
    if (/出席/.test(v)) attCols.push(c)
    if (/表現/.test(v)) behCols.push(c)
    if (/堂課/.test(v)) clsCols.push(c)
  }
  if (!attCols.length) return { ok: false, scoreRow }

  // 3. Detect first student row
  let firstDataRow = scoreRow + 1
  while (firstDataRow <= maxRow && !/^\d+$/.test(getVal(firstDataRow, 0))) firstDataRow++

  // 4. Find date target cells
  //
  //  Detection order:
  //  A) Find cells with "日期" text in the header area → replace them with actual date
  //  B) Date cell is one column LEFT of each 出席分 column in the same row (e.g. C3 when D3=出席分)
  //  C) Empty row above scoreRow at attCols positions
  //
  const dateTargets = []   // [{ r, c }]

  // A: scan for "日期" text (partial match, any row before student data)
  const foundDateCells = []
  for (let r = 0; r < firstDataRow; r++) {
    for (let c = 0; c <= maxCol; c++) {
      if (/日期/.test(getVal(r, c))) foundDateCells.push({ r, c })
    }
  }

  if (foundDateCells.length >= attCols.length) {
    // Found enough "日期" cells — use them in order
    foundDateCells.sort((a, b) => a.r !== b.r ? a.r - b.r : a.c - b.c)
    dateTargets.push(...foundDateCells.slice(0, attCols.length))
  } else if (foundDateCells.length === 1 && foundDateCells[0].c <= 2) {
    // "日期" is a ROW LABEL → dates go in that row, at attCols positions
    const labelRow = foundDateCells[0].r
    attCols.forEach(c => dateTargets.push({ r: labelRow, c }))
  } else {
    // B: date column = one column to the LEFT of each 出席分 column, in scoreRow
    //    e.g. 出席分 at D3 (col 3) → date at C3 (col 2)
    const leftCols = attCols.map(ac => ac - 1)
    const leftOk = leftCols.every(dc => {
      if (dc < 2) return false                          // would overwrite 號 / 姓名
      const v = getVal(scoreRow, dc)
      return !v || /日期/.test(v)                       // must be empty or already "日期"
    })

    if (leftOk) {
      attCols.forEach(ac => dateTargets.push({ r: scoreRow, c: ac - 1 }))
    } else {
      // C: look for an empty row above scoreRow at attCols positions
      let dateRow = -1
      for (let r = scoreRow - 1; r >= Math.max(0, scoreRow - 5); r--) {
        if (attCols.every(c => !getVal(r, c))) { dateRow = r; break }
      }
      if (dateRow === -1) dateRow = Math.max(0, scoreRow - 1)
      attCols.forEach(c => dateTargets.push({ r: dateRow, c }))
    }
  }

  // 5. Fill dates ("18/6")
  batch.forEach((sess, i) => {
    if (i < dateTargets.length) setVal(dateTargets[i].r, dateTargets[i].c, formatDate(sess.date))
  })

  // 6. Fill scores — match student by exact name
  batch.forEach((sess, si) => {
    if (si >= attCols.length) return
    sess.records.forEach(rec => {
      let sRow = -1
      outer: for (let r = firstDataRow; r <= maxRow; r++) {
        for (let c = 0; c <= Math.min(maxCol, 5); c++) {
          if (getVal(r, c) === rec.name) { sRow = r; break outer }
        }
      }
      if (sRow === -1) return
      setVal(sRow, attCols[si], rec.attendance)
      if (si < behCols.length) setVal(sRow, behCols[si], rec.behavior ?? 0)
      if (si < clsCols.length) setVal(sRow, clsCols[si], rec.classScore ?? 0)
    })
  })

  return { ok: true, scoreRow }
}

function _triggerDownload(buf, filename) {
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const a = Object.assign(document.createElement('a'), {
    href: URL.createObjectURL(blob), download: filename,
  })
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(a.href)
}

// ── exportExcel: all classes in one file ─────────────────────────────────────

async function exportExcel() {
  const classesWithData = sheetNames.value.filter(n => (sessions.value[n]?.length ?? 0) > 0)
  if (!classesWithData.length) { alert('所有班級均無記錄可匯出'); return }

  let origBuf
  try {
    const res = await fetch('/25-26下普中各班分紙3-試用.xls')
    if (!res.ok) throw new Error(`HTTP ${res.status}`)
    origBuf = await res.arrayBuffer()
  } catch (e) {
    alert(`無法讀取原始分紙：${e.message}`); return
  }

  const outWb = XLSX.read(new Uint8Array(origBuf), { type: 'array', cellStyles: true })
  const sheetsToKeep = new Set()

  for (const className of classesWithData) {
    const all        = sessions.value[className] ?? []
    const batchCount = Math.ceil(all.length / 7)

    for (let b = 0; b < batchCount; b++) {
      const batch = all.slice(b * 7, b * 7 + 7)
      let ws
      if (b === 0) {
        ws = outWb.Sheets[className]
        sheetsToKeep.add(className)
      } else {
        const clone = XLSX.read(new Uint8Array(origBuf), { type: 'array', cellStyles: true })
        ws = clone.Sheets[className]
        if (!ws) continue
        const nm = `${className}_批${b + 1}`.slice(0, 31)
        outWb.Sheets[nm] = ws
        outWb.SheetNames.push(nm)
        sheetsToKeep.add(nm)
      }
      if (!ws) continue
      fillSheet(ws, batch)
    }
  }

  // Snapshot the current sheet order before removing unused sheets —
  // Workbook.Names entries use 0-based indices into this array
  const allSheetNames = [...outWb.SheetNames]

  // Remove sheets with no data
  outWb.SheetNames = outWb.SheetNames.filter(n => sheetsToKeep.has(n))
  Object.keys(outWb.Sheets).forEach(n => { if (!sheetsToKeep.has(n)) delete outWb.Sheets[n] })

  // Drop named ranges (print areas, print titles, etc.) whose sheet was removed,
  // and update the Sheet index for ones that remain. Without this, Excel reports
  // "已删除的功能: /xl/workbook.xml 部分的 命名区域" on every open.
  if (outWb.Workbook?.Names?.length) {
    outWb.Workbook.Names = outWb.Workbook.Names.reduce((acc, nm) => {
      const sheetName = nm.Sheet != null ? allSheetNames[nm.Sheet] : null
      if (!sheetName || !sheetsToKeep.has(sheetName)) return acc
      const newIdx = outWb.SheetNames.indexOf(sheetName)
      if (newIdx !== -1) acc.push({ ...nm, Sheet: newIdx })
      return acc
    }, [])
  }

  const xlsxArr = XLSX.write(outWb, { type: 'array', bookType: 'xlsx' })
  _triggerDownload(
    new Uint8Array(xlsxArr).buffer,
    `點名記錄_${formatDate(currentDate.value).replace('/', '-')}.xlsx`,
  )
}
</script>

<template>
  <div id="app">

    <!-- ── Header ── -->
    <header class="app-header">
      <div class="header-top">
        <h1>班級點名系統</h1>
        <span class="subtitle">25-26 下普中</span>
      </div>
      <nav class="class-tabs" v-if="sheetNames.length">
        <button
          v-for="n in sheetNames" :key="n"
          :class="['class-tab', { active: selectedClass === n }]"
          @click="selectClass(n)"
        >
          {{ n }}
          <span v-if="getSessionCount(n)" class="count-badge">{{ getSessionCount(n) }}</span>
        </button>
      </nav>
    </header>

    <!-- ── Loading / Error ── -->
    <div class="loading-state" v-if="loading">
      <div class="spinner"></div><p>載入 Excel…</p>
    </div>
    <div class="error-state" v-else-if="loadError">
      <p>{{ loadError }}</p>
      <p class="hint">請確認 public 資料夾中有 <code>25-26下普中各班分紙3-試用.xls</code></p>
    </div>

    <!-- ── Main ── -->
    <div v-else class="main-wrap">

      <!-- view toggle -->
      <div class="view-toggle">
        <button :class="['view-btn', { active: activeTab==='attendance' }]" @click="activeTab='attendance'">
          點名
        </button>
        <button :class="['view-btn', { active: activeTab==='history' }]" @click="activeTab='history'">
          記錄
          <span class="pill">{{ batchProgress }}/7</span>
        </button>
      </div>

      <!-- ════ Attendance tab ════ -->
      <div v-show="activeTab==='attendance'" class="att-view">

        <div class="form-bar">
          <label class="date-lbl">
            日期
            <input type="date" v-model="currentDate" class="date-input" />
          </label>
          <span v-if="editingEntry" class="edit-pill">編輯模式 — 第 {{ editingEntry.idx + 1 }} 堂</span>
          <span v-else class="session-pill">第 {{ sessionCount + 1 }} 堂（批次 {{ currentBatch }}）</span>
          <button v-if="editingEntry" class="btn-cancel-edit" @click="cancelEdit">取消編輯</button>
          <div v-if="students.length" class="quick-actions">
            <button class="btn-quick btn-qa-att" @click="markAllPresent">一鍵點名</button>
            <button class="btn-quick btn-qa-beh" :disabled="!markedCount" @click="fillAllBehavior">表現＝出席</button>
            <button class="btn-quick btn-qa-cls" :disabled="!markedCount" @click="fillAllClassScore">堂課＝出席</button>
          </div>
        </div>

        <div v-if="!students.length" class="empty-state">
          <p>未能讀取學生名單</p>
          <p class="hint">Excel 格式：第1欄＝號碼，第2欄＝中文姓名</p>
        </div>

        <div v-else class="table-wrap">
          <table class="att-table">
            <thead>
              <tr>
                <th class="th-num">號</th>
                <th class="th-name">姓名</th>
                <th class="th-att">出席狀況</th>
                <th class="th-score">表現分<br><small>≤出席分</small></th>
                <th class="th-score">堂課分</th>
              </tr>
            </thead>
            <tbody>
              <tr
                v-for="s in students" :key="s.name"
                :class="['s-row', getRowCls(s.name)]"
              >
                <td class="td-num">{{ s.num }}</td>
                <td class="td-name">{{ s.name }}</td>

                <td class="td-att">
                  <div class="radio-grp">
                    <label
                      v-for="opt in ATTENDANCE_OPTIONS" :key="opt.value"
                      :class="['r-opt', opt.cls,
                        { sel: records[s.name]?.attendance === opt.value }]"
                    >
                      <input
                        type="radio"
                        :name="`att_${s.name}`"
                        :value="opt.value"
                        v-model="records[s.name].attendance"
                        @change="onAttendanceChange(s.name)"
                      />
                      <span class="r-lbl">{{ opt.label }}</span>
                      <span class="r-score">{{ opt.value }}</span>
                    </label>
                  </div>
                </td>

                <td class="td-score">
                  <input
                    type="number" class="score-inp"
                    v-model.number="records[s.name].behavior"
                    :max="records[s.name].attendance ?? 0"
                    min="0" placeholder="0"
                    @input="validateBehavior(s.name)"
                    :disabled="records[s.name].attendance === null"
                  />
                </td>
                <td class="td-score">
                  <input
                    type="number" class="score-inp"
                    v-model.number="records[s.name].classScore"
                    min="0" placeholder="0"
                    :disabled="records[s.name].attendance === null"
                  />
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        <div class="submit-bar" v-if="students.length">
          <span class="marked-info">
            出席已填 {{ markedCount }} / {{ students.length }} 人
            <small v-if="markedCount === students.length" style="color:#43a047;margin-left:6px">✓ 可儲存</small>
          </span>
          <button
            class="btn-submit"
            :disabled="markedCount < students.length"
            @click="submitSession"
          >{{ editingEntry ? '更新記錄' : '儲存本堂記錄' }}</button>
        </div>
      </div>

      <!-- ════ History tab ════ -->
      <div v-show="activeTab==='history'" class="hist-view">

        <div class="hist-toolbar">
          <div class="batch-info">
            <div class="batch-bar"><div class="batch-fill" :style="{ width: progressPct }"></div></div>
            <span class="batch-lbl">批次 {{ currentBatch }} ─ {{ batchProgress }}/7 堂完成</span>
          </div>
          <button class="btn-inspect" @click="inspectStructure">查看結構</button>
          <button class="btn-export" :disabled="!sessionCount" @click="exportExcel">
            ↓ 填入並匯出
          </button>
        </div>

        <div class="empty-state" v-if="!sessionCount"><p>尚無記錄</p></div>

        <div class="sess-list" v-else>
          <div
            v-for="(sess, idx) in curSessions" :key="sess.id"
            class="sess-card"
          >
            <div class="sess-top">
              <div class="sess-meta">
                <span class="sess-num">第 {{ idx+1 }} 堂</span>
                <span class="sess-date">{{ formatDate(sess.date) }}</span>
              </div>
              <div class="sess-actions">
                <button class="btn-detail" @click="detailSession = sess">明細</button>
                <button class="btn-edit"   @click="startEdit(sess, idx)">編輯</button>
                <button class="btn-del"    @click="deleteSession(idx)">刪除</button>
              </div>
            </div>
            <div class="sess-stats">
              <span class="stat present">準時 {{ countAtt(sess, 100) }}</span>
              <span class="stat late">遲到 {{ countAtt(sess, 90) }}</span>
              <span class="stat second">第二堂準時 {{ countAtt(sess, 50) }}</span>
              <span class="stat second-late">第二堂遲到 {{ countAtt(sess, 40) }}</span>
              <span class="stat absent">缺席 {{ countAtt(sess, 0) }}</span>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- ── Structure Inspect Modal ── -->
    <div class="modal-overlay" v-if="sheetStructure" @click.self="sheetStructure=null">
      <div class="modal modal-wide">
        <div class="modal-head">
          <span>Excel 結構預覽 — {{ selectedClass }}（前12行）</span>
          <button class="modal-close" @click="sheetStructure=null">✕</button>
        </div>
        <div class="modal-body">
          <p class="struct-hint">確認「出席分」「表現分」「堂課分」欄位標題存在，否則填入會失敗。</p>
          <table class="struct-table">
            <thead><tr><th>行</th><th v-for="(_, ci) in sheetStructure[0]?.cells" :key="ci">欄{{ ci }}</th></tr></thead>
            <tbody>
              <tr v-for="row in sheetStructure" :key="row.r">
                <td class="struct-rownum">{{ row.r }}</td>
                <td
                  v-for="(cell, ci) in row.cells" :key="ci"
                  :class="{ 'struct-hit': /出席|表現|堂課/.test(String(cell)) }"
                >{{ cell }}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ── Detail Modal ── -->
    <div class="modal-overlay" v-if="detailSession" @click.self="detailSession=null">
      <div class="modal">
        <div class="modal-head">
          <span>{{ selectedClass }} ─ {{ formatDate(detailSession.date) }}</span>
          <button class="modal-close" @click="detailSession=null">✕</button>
        </div>
        <div class="modal-body">
          <table class="detail-table">
            <thead>
              <tr>
                <th>號</th><th>姓名</th><th>出席</th><th>出席分</th><th>表現分</th><th>堂課分</th><th>合計</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="r in detailSession.records" :key="r.name">
                <td>{{ r.num }}</td>
                <td>{{ r.name }}</td>
                <td>{{ attLabel(r.attendance) }}</td>
                <td>{{ r.attendance }}</td>
                <td>{{ r.behavior }}</td>
                <td>{{ r.classScore }}</td>
                <td class="total-cell">{{ r.attendance + r.behavior + r.classScore }}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

  </div>
</template>

<style scoped>
/* ── Reset ── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

#app {
  min-height: 100vh;
  font-family: 'Segoe UI', 'PingFang TC', 'Microsoft JhengHei', system-ui, sans-serif;
  font-size: 14px;
  background: #f0f2f5;
  color: #1a1a2e;
  display: flex;
  flex-direction: column;
}

/* ── Header ── */
.app-header {
  background: #1a1a2e;
  color: #fff;
  padding: 12px 16px 10px;
  position: sticky; top: 0; z-index: 100;
  box-shadow: 0 2px 10px rgba(0,0,0,.3);
}
.header-top { display: flex; align-items: baseline; gap: 10px; margin-bottom: 10px; }
h1 { font-size: 17px; font-weight: 700; }
.subtitle { font-size: 12px; opacity: .55; }

.class-tabs { display: flex; flex-wrap: wrap; gap: 6px; }
.class-tab {
  padding: 5px 14px; border-radius: 20px;
  border: 1px solid rgba(255,255,255,.25);
  background: transparent; color: rgba(255,255,255,.65);
  cursor: pointer; font-size: 13px;
  transition: all .18s;
}
.class-tab.active { background: #4f8ff7; border-color: #4f8ff7; color: #fff; font-weight: 600; }
.class-tab:hover:not(.active) { background: rgba(255,255,255,.1); color: #fff; }
.count-badge {
  background: #ff5252; color: #fff;
  border-radius: 9px; padding: 0 6px; font-size: 11px; margin-left: 4px;
}

/* ── Main wrap ── */
.main-wrap {
  flex: 1; width: 100%; max-width: 1300px;
  margin: 0 auto; padding: 14px 14px 32px;
}

/* ── View toggle ── */
.view-toggle {
  display: flex; background: #fff;
  border-radius: 10px; padding: 4px; width: fit-content;
  margin-bottom: 14px;
  box-shadow: 0 1px 4px rgba(0,0,0,.08);
}
.view-btn {
  padding: 7px 22px; border: none; border-radius: 7px;
  background: transparent; color: #666; cursor: pointer;
  font-size: 13px; display: flex; align-items: center; gap: 6px;
  transition: all .18s;
}
.view-btn.active { background: #4f8ff7; color: #fff; font-weight: 600; }
.pill {
  background: rgba(255,255,255,.25); border-radius: 10px;
  padding: 1px 7px; font-size: 11px;
}
.view-btn:not(.active) .pill { background: #eee; color: #777; }

/* ── Form bar ── */
.form-bar {
  display: flex; align-items: center; gap: 14px; flex-wrap: wrap;
  background: #fff; border-radius: 10px; padding: 11px 16px;
  margin-bottom: 12px; box-shadow: 0 1px 4px rgba(0,0,0,.07);
}
.date-lbl { display: flex; align-items: center; gap: 8px; font-weight: 500; }
.date-input { border: 1px solid #ddd; border-radius: 6px; padding: 5px 10px; font-size: 14px; }
.session-pill {
  background: #e8f0fe; color: #1967d2;
  border-radius: 20px; padding: 4px 13px; font-size: 12px; font-weight: 500;
}
.edit-pill {
  background: #fff3e0; color: #e65100;
  border-radius: 20px; padding: 4px 13px; font-size: 12px; font-weight: 600;
}
.btn-cancel-edit {
  padding: 5px 14px; background: none; border: 1px solid #ffcc80;
  color: #e65100; border-radius: 7px; font-size: 13px; cursor: pointer;
}
.btn-cancel-edit:hover { background: #fff3e0; }

.quick-actions { display: flex; gap: 6px; margin-left: auto; }
.btn-quick {
  padding: 6px 13px; border: none; border-radius: 7px;
  font-size: 12px; font-weight: 600; cursor: pointer; transition: filter .15s;
}
.btn-quick:disabled { opacity: .35; cursor: not-allowed; }
.btn-qa-att { background: #43a047; color: #fff; }
.btn-qa-att:hover:not(:disabled) { filter: brightness(1.1); }
.btn-qa-beh { background: #1e88e5; color: #fff; }
.btn-qa-beh:hover:not(:disabled) { filter: brightness(1.1); }
.btn-qa-cls { background: #8e24aa; color: #fff; }
.btn-qa-cls:hover:not(:disabled) { filter: brightness(1.1); }

/* ── Table ── */
.table-wrap {
  background: #fff; border-radius: 10px;
  box-shadow: 0 1px 4px rgba(0,0,0,.07);
  overflow-x: auto;
}
.att-table { width: 100%; border-collapse: collapse; }
.att-table thead th {
  background: #1a1a2e; color: #fff;
  padding: 10px 12px; text-align: left; font-size: 13px;
  white-space: nowrap;
}
.att-table thead th small { font-weight: 400; opacity: .7; }
.att-table tbody tr { border-bottom: 1px solid #f0f0f0; transition: background .12s; }
.att-table tbody tr:hover { filter: brightness(.97); }

.th-num  { width: 46px; text-align: center; }
.th-name { width: 84px; }
.th-att  { }
.th-score { width: 78px; text-align: center; }

.td-num  { text-align: center; color: #888; padding: 8px 12px; }
.td-name { font-weight: 600; padding: 8px 12px; white-space: nowrap; }
.td-att  { padding: 6px 12px; }
.td-score { text-align: center; padding: 6px 10px; }

/* Row tints */
.s-row.row-present     { background: #f1f8e9 !important; }
.s-row.row-late        { background: #fff8e1 !important; }
.s-row.row-second      { background: #e3f2fd !important; }
.s-row.row-second-late { background: #f3e5f5 !important; }
.s-row.row-absent      { background: #ffebee !important; }

/* ── Radio options ── */
.radio-grp { display: flex; gap: 5px; flex-wrap: wrap; }
.r-opt {
  display: flex; flex-direction: column; align-items: center;
  padding: 5px 9px; border: 2px solid #eee; border-radius: 8px;
  cursor: pointer; min-width: 68px; transition: all .14s; user-select: none;
}
.r-opt input[type="radio"] { display: none; }
.r-lbl   { font-size: 11px; font-weight: 500; line-height: 1.3; text-align: center; }
.r-score { font-size: 15px; font-weight: 700; }

.r-opt.opt-present            { border-color: #e8f5e9; }
.r-opt.opt-present.sel        { background: #43a047; border-color: #43a047; color: #fff; }
.r-opt.opt-present:hover:not(.sel) { border-color: #43a047; }

.r-opt.opt-late               { border-color: #fff8e1; }
.r-opt.opt-late.sel           { background: #fb8c00; border-color: #fb8c00; color: #fff; }
.r-opt.opt-late:hover:not(.sel) { border-color: #fb8c00; }

.r-opt.opt-second             { border-color: #e3f2fd; }
.r-opt.opt-second.sel         { background: #1e88e5; border-color: #1e88e5; color: #fff; }
.r-opt.opt-second:hover:not(.sel) { border-color: #1e88e5; }

.r-opt.opt-second-late        { border-color: #f3e5f5; }
.r-opt.opt-second-late.sel    { background: #8e24aa; border-color: #8e24aa; color: #fff; }
.r-opt.opt-second-late:hover:not(.sel) { border-color: #8e24aa; }

.r-opt.opt-absent             { border-color: #ffebee; }
.r-opt.opt-absent.sel         { background: #e53935; border-color: #e53935; color: #fff; }
.r-opt.opt-absent:hover:not(.sel) { border-color: #e53935; }

/* ── Score inputs ── */
.score-inp {
  width: 58px; padding: 6px 6px; border: 1px solid #ddd;
  border-radius: 6px; font-size: 14px; text-align: center;
  transition: border-color .14s;
}
.score-inp:focus { outline: none; border-color: #4f8ff7; }
.score-inp:disabled { background: #f5f5f5; color: #bbb; }

/* ── Submit bar ── */
.submit-bar {
  display: flex; justify-content: space-between; align-items: center;
  background: #fff; border-radius: 10px; padding: 11px 16px;
  margin-top: 12px; box-shadow: 0 1px 4px rgba(0,0,0,.07);
}
.marked-info { color: #666; font-size: 13px; }
.btn-submit {
  padding: 9px 26px; background: #4f8ff7; color: #fff;
  border: none; border-radius: 8px; font-size: 14px; font-weight: 600;
  cursor: pointer; transition: background .18s;
}
.btn-submit:hover:not(:disabled) { background: #3579e6; }
.btn-submit:disabled { background: #b0c4de; cursor: not-allowed; }

/* ── History toolbar ── */
.hist-toolbar {
  display: flex; align-items: center; gap: 14px; flex-wrap: wrap;
  background: #fff; border-radius: 10px; padding: 13px 16px;
  margin-bottom: 12px; box-shadow: 0 1px 4px rgba(0,0,0,.07);
}
.batch-info { flex: 1; display: flex; align-items: center; gap: 12px; min-width: 200px; }
.batch-bar { flex: 1; height: 10px; background: #eee; border-radius: 5px; overflow: hidden; }
.batch-fill { height: 100%; background: #4f8ff7; border-radius: 5px; transition: width .3s; }
.batch-lbl  { font-size: 13px; color: #666; white-space: nowrap; }
.btn-inspect {
  padding: 9px 14px; background: #fff; color: #666;
  border: 1px solid #ddd; border-radius: 8px; font-size: 13px;
  cursor: pointer; white-space: nowrap;
}
.btn-inspect:hover { background: #f5f5f5; }
.btn-export {
  padding: 9px 20px; background: #00c853; color: #fff;
  border: none; border-radius: 8px; font-size: 14px; font-weight: 600;
  cursor: pointer; white-space: nowrap;
}
.btn-export:hover:not(:disabled) { background: #00b548; }
.btn-export:disabled { background: #aaa; cursor: not-allowed; }

/* ── Structure inspect modal ── */
.modal-wide { max-width: 95vw; width: 1000px; }
.struct-hint { padding: 10px 18px 6px; font-size: 12px; color: #888; }
.struct-table { width: 100%; border-collapse: collapse; font-size: 12px; }
.struct-table th, .struct-table td {
  border: 1px solid #e0e0e0; padding: 5px 8px; white-space: nowrap; min-width: 40px;
}
.struct-table th { background: #f5f7fa; font-weight: 600; }
.struct-rownum { background: #f5f7fa; color: #999; font-weight: 600; }
.struct-hit { background: #fff3cd; font-weight: 700; color: #e65100; }

/* ── Session cards ── */
.sess-list { display: flex; flex-direction: column; gap: 8px; }
.sess-card { background: #fff; border-radius: 10px; padding: 12px 16px; box-shadow: 0 1px 4px rgba(0,0,0,.07); }
.sess-top { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }
.sess-meta { display: flex; align-items: center; gap: 10px; }
.sess-num  { font-weight: 700; font-size: 15px; }
.sess-date { color: #666; font-size: 13px; }
.sess-actions { display: flex; gap: 6px; }

.btn-detail {
  border: 1px solid #c5d8f8; color: #1967d2;
  background: none; border-radius: 6px; padding: 3px 10px; font-size: 12px; cursor: pointer;
}
.btn-detail:hover { background: #e8f0fe; }
.btn-edit {
  border: 1px solid #ffe082; color: #e65100;
  background: none; border-radius: 6px; padding: 3px 10px; font-size: 12px; cursor: pointer;
}
.btn-edit:hover { background: #fff8e1; }
.btn-del {
  border: 1px solid #ffcdd2; color: #e53935;
  background: none; border-radius: 6px; padding: 3px 10px; font-size: 12px; cursor: pointer;
}
.btn-del:hover { background: #ffebee; }

.sess-stats { display: flex; flex-wrap: wrap; gap: 6px; }
.stat { padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 500; }
.stat.present      { background: #e8f5e9; color: #2e7d32; }
.stat.late         { background: #fff8e1; color: #e65100; }
.stat.second       { background: #e3f2fd; color: #1565c0; }
.stat.second-late  { background: #f3e5f5; color: #6a1b9a; }
.stat.absent       { background: #ffebee; color: #c62828; }

/* ── Modal ── */
.modal-overlay {
  position: fixed; inset: 0; background: rgba(0,0,0,.45);
  display: flex; align-items: center; justify-content: center; z-index: 500; padding: 16px;
}
.modal {
  background: #fff; border-radius: 12px;
  width: 100%; max-width: 700px; max-height: 85vh;
  display: flex; flex-direction: column; box-shadow: 0 8px 32px rgba(0,0,0,.25);
}
.modal-head {
  display: flex; justify-content: space-between; align-items: center;
  padding: 14px 18px; font-weight: 600; font-size: 15px;
  border-bottom: 1px solid #eee;
}
.modal-close {
  background: none; border: none; font-size: 16px; cursor: pointer; color: #888; line-height: 1;
}
.modal-close:hover { color: #333; }
.modal-body { overflow: auto; flex: 1; }

.detail-table { width: 100%; border-collapse: collapse; }
.detail-table th { background: #f5f7fa; padding: 9px 14px; text-align: left; font-size: 13px; }
.detail-table td { padding: 8px 14px; border-bottom: 1px solid #f0f0f0; font-size: 13px; }
.detail-table tbody tr:hover { background: #fafafa; }
.total-cell { font-weight: 700; color: #1967d2; }

/* ── Loading / Error / Empty ── */
.loading-state, .error-state {
  flex: 1; display: flex; flex-direction: column;
  align-items: center; justify-content: center;
  gap: 14px; color: #666; padding: 60px 20px;
}
.spinner {
  width: 38px; height: 38px;
  border: 4px solid #eee; border-top-color: #4f8ff7;
  border-radius: 50%;
  animation: spin .75s linear infinite;
}
@keyframes spin { to { transform: rotate(360deg); } }
.empty-state { text-align: center; padding: 40px; color: #aaa; }
.hint { font-size: 12px; margin-top: 6px; opacity: .7; }

/* ── Responsive ── */
@media (max-width: 700px) {
  .r-opt { min-width: 58px; padding: 4px 6px; }
  .r-lbl { font-size: 10px; }
  .r-score { font-size: 13px; }
  .score-inp { width: 50px; }
  .th-score, .td-score { width: 62px; }
}
</style>
