<template>
  <div class="layout">

    <header class="topbar">
      <h1>🐟 魚種圖鑑 Excel 管理系統</h1>

      <div class="actions">
        <input v-model="excelUrl" placeholder="輸入 Excel 網址" />
        <button @click="loadFromUrl">用網址載入</button>

        <input type="file" accept=".xlsx,.xls" @change="uploadExcel" />

        <button class="cloud" @click="loadFromCloud">
          📡 從雲端同步
        </button>

        <button class="export" @click="exportExcel">
          匯出 Excel
        </button>
      </div>
    </header>

    <div class="body">

      <aside class="sidebar">
        <div
          v-for="(sheet, idx) in sheets"
          :key="sheet.name"
          :class="['sheet-btn', { active: idx === activeSheetIndex }]"
          @click="activeSheetIndex = idx"
        >
          {{ sheet.name }}
        </div>
      </aside>

      <main class="main" v-if="activeSheet">

        <div class="table-wrap">

          <table>
            <thead>
              <tr>
                <th width="30"></th>
                <th v-for="h in DISPLAY_HEADERS" :key="h">
                  {{ h }}
                </th>
                <th width="60">操作</th>
              </tr>
            </thead>

            <tbody>
              <tr
                v-for="(row, r) in rows"
                :key="r"
                draggable="!isEmptyRow(row)"
                @dragstart="dragStart(r)"
                @dragover.prevent
                @drop="dropRow(r)"
                :class="isInvalidRow(row) ? 'row-error' : ''"
              >

                <td class="drag" v-if="!isEmptyRow(row)">⋮⋮</td>
                <td v-else></td>

                <td
                  v-for="(cell, c) in row"
                  :key="c"
                  :class="[
                    isInvalidCell(row, c) ? 'error' : '',
                    editingCell.row === r && editingCell.col === c
                      ? 'cell-editing'
                      : ''
                  ]"
                >

                  <select
                    v-if="c === TYPE_COL_INDEX"
                    v-model.number="activeSheet.data[DATA_START_ROW + r][c]"
                    class="select"
                    @focus="setEditing(r, c)"
                    @change="sortRows"
                  >
                    <option
                      v-for="opt in TYPE_OPTIONS"
                      :key="opt.value"
                      :value="opt.value"
                    >
                      {{ opt.label }}
                    </option>
                  </select>

                  <select
                    v-else-if="c === TITLE_COL_INDEX"
                    v-model="activeSheet.data[DATA_START_ROW + r][c]"
                    class="select"
                    @focus="setEditing(r, c)"
                  >
                    <option
                      v-for="opt in TITLE_OPTIONS"
                      :key="opt.value"
                      :value="opt.value"
                    >
                      {{ opt.label }}
                    </option>
                  </select>

                  <input
                    v-else-if="c === MIN_COL_INDEX || c === MAX_COL_INDEX"
                    type="number"
                    class="number-input"
                    v-model.number="activeSheet.data[DATA_START_ROW + r][c]"
                    @focus="setEditing(r, c)"
                  />

                  <div
                    v-else
                    contenteditable
                    class="editable"
                    @focus="setEditing(r, c)"
                    @input="updateCell(r, c, $event)"
                  >
                    {{ cell }}
                  </div>

                </td>

                <td class="delete-cell">
                  <button
                    v-if="!isEmptyRow(row)"
                    class="delete-btn"
                    @click="deleteRow(r)"
                  >
                    🗑
                  </button>
                </td>

              </tr>
            </tbody>
          </table>

        </div>
      </main>
    </div>
  </div>
</template>

<script setup>
import { ref, computed } from "vue"
import axios from "axios"
import * as XLSX from "xlsx"
import { parseExcel } from "../utils/excel"

const DATA_START_ROW = 5
const COL_COUNT = 7

const DISPLAY_HEADERS = [
  "魚種類型",
  "魚種名稱",
  "最小倍率",
  "最高倍率",
  "Tag",
  "標題",
  "類型"
]

const MIN_COL_INDEX = 2
const MAX_COL_INDEX = 3
const TITLE_COL_INDEX = 5
const TYPE_COL_INDEX = 6

const TYPE_OPTIONS = [
  { value: 0, label: "一般魚" },
  { value: 2, label: "Boss" },
  { value: 1, label: "活動魚" }
]

const TITLE_OPTIONS = [
  { value: "NONE", label: "無" },
  { value: "J", label: "金蟬大獎" }
]

const excelUrl = ref("")
const sheets = ref([])
const activeSheetIndex = ref(0)
const editingCell = ref({ row: null, col: null })
const dragIndex = ref(null)

const activeSheet = computed(() =>
  sheets.value[activeSheetIndex.value]
)

const rows = computed(() => {
  if (!activeSheet.value) return []
  return activeSheet.value.data.slice(DATA_START_ROW)
})

function setEditing(r, c) {
  editingCell.value = { row: r, col: c }
}

function ensureEmptyRow() {
  const data = activeSheet.value.data
  const body = data.slice(DATA_START_ROW)

  // 移除多餘空白行
  const filtered = body.filter(row =>
    row.some(v => v)
  )

  activeSheet.value.data = [
    ...data.slice(0, DATA_START_ROW),
    ...filtered,
    new Array(COL_COUNT).fill("")
  ]
}

function deleteRow(index) {
  activeSheet.value.data.splice(DATA_START_ROW + index, 1)
  ensureEmptyRow()
}

function isEmptyRow(row) {
  return row.every(v => !v)
}

function updateCell(row, col, e) {
  activeSheet.value.data[DATA_START_ROW + row][col] =
    e.target.innerText
  ensureEmptyRow()
}

function isInvalidRow(row) {
  const min = Number(row[MIN_COL_INDEX])
  const max = Number(row[MAX_COL_INDEX])
  return !isNaN(min) && !isNaN(max) && min > max
}

function isInvalidCell(row, col) {
  if (col !== MIN_COL_INDEX && col !== MAX_COL_INDEX)
    return false
  return isInvalidRow(row)
}

/* ===== 排序 ===== */
function sortRows() {
  const order = { 0: 0, 2: 1, 1: 2 }

  const body = activeSheet.value.data
    .slice(DATA_START_ROW)
    .filter(row => row.some(v => v))

  body.sort((a, b) =>
    order[a[TYPE_COL_INDEX]] - order[b[TYPE_COL_INDEX]]
  )

  activeSheet.value.data = [
    ...activeSheet.value.data.slice(0, DATA_START_ROW),
    ...body,
    new Array(COL_COUNT).fill("")
  ]
}

/* ===== 拖曳 ===== */
function dragStart(index) {
  dragIndex.value = index
}

function dropRow(targetIndex) {
  if (dragIndex.value === null) return

  const from = DATA_START_ROW + dragIndex.value
  const to = DATA_START_ROW + targetIndex

  const data = activeSheet.value.data
  const moved = data.splice(from, 1)[0]
  data.splice(to, 0, moved)

  dragIndex.value = null
}

async function loadFromCloud() {
  const res = await axios.get(
    "https://excelproxy.kin169999.workers.dev/api/excel",
    { responseType: "arraybuffer" }
  )
  sheets.value = parseExcel(res.data)
  ensureEmptyRow()
}

async function loadFromUrl() {
  if (!excelUrl.value) return
  const res = await axios.get(excelUrl.value, {
    responseType: "arraybuffer"
  })
  sheets.value = parseExcel(res.data)
  ensureEmptyRow()
}

function uploadExcel(e) {
  const file = e.target.files[0]
  if (!file) return
  const reader = new FileReader()
  reader.onload = evt => {
    sheets.value = parseExcel(evt.target.result)
    ensureEmptyRow()
  }
  reader.readAsArrayBuffer(file)
}

function exportExcel() {
  const wb = XLSX.utils.book_new()
  sheets.value.forEach(sheet => {
    const ws = XLSX.utils.aoa_to_sheet(sheet.data)
    XLSX.utils.book_append_sheet(wb, ws, sheet.name)
  })
  XLSX.writeFile(wb, "fish_data_export.xlsx")
}
</script>

<style scoped>
.layout { height:100vh; display:flex; flex-direction:column; background:#020617; color:#e5e7eb }
.topbar { height:60px; border-bottom:1px solid #1e293b; display:flex; align-items:center; justify-content:space-between; padding:0 20px }
.actions { display:flex; gap:10px }
.actions button.export { background:#16a34a }
.actions button.cloud { background:#0ea5e9 }
.body { flex:1; display:flex; overflow:hidden }
.sidebar { width:220px; border-right:1px solid #1e293b; padding:10px }
.sheet-btn { padding:10px; margin-bottom:6px; border-radius:6px; cursor:pointer }
.sheet-btn.active { background:#2563eb }
.main { flex:1; padding:10px }
.table-wrap { height:100%; overflow:auto; border:1px solid #1e293b; border-radius:8px }
table { width:100%; border-collapse:collapse; table-layout:fixed }
thead th { position:sticky; top:0; background:#0f172a; padding:8px }
tr.row-error td { background:rgba(220,38,38,0.35)!important }
td { border-bottom:1px solid #1e293b; padding:4px }
td.cell-editing { box-shadow: inset 0 0 0 2px #facc15 }
.select,.number-input { width:100%; background:#020617; color:white; border:1px solid #334155; padding:4px; border-radius:4px }
.editable { min-height:22px; outline:none }
.delete-btn { background:none; border:none; cursor:pointer; font-size:16px; color:#94a3b8 }
.delete-btn:hover { color:#ef4444 }
.drag { cursor:grab; text-align:center; color:#64748b }
</style>
