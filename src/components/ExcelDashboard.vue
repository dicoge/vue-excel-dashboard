<template>
  <div class="layout">

    <!-- 🔥 Loading Overlay -->
    <div v-if="isLoading" class="loading-overlay">
      <div class="loading-box">
        <div class="spinner"></div>
        <p>正在載入 Excel...</p>
      </div>
    </div>

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
          @click="switchSheet(idx)"
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
                :draggable="!isEmptyRow(row)"
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

                <td>
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
const isLoading = ref(false)

const activeSheet = computed(() => sheets.value[activeSheetIndex.value])

const rows = computed(() => {
  if (!activeSheet.value) return []
  return activeSheet.value.data
    .slice(DATA_START_ROW)
    .map(row => {
      const fixed = [...row]
      while (fixed.length < COL_COUNT) fixed.push("")
      return fixed.slice(0, COL_COUNT)
    })
})

function switchSheet(index) {
  activeSheetIndex.value = index
  ensureEmptyRow()
}

function ensureEmptyRow() {
  if (!activeSheet.value) return
  const data = activeSheet.value.data
  const body = data.slice(DATA_START_ROW).filter(r => r.some(v => v))
  activeSheet.value.data = [
    ...data.slice(0, DATA_START_ROW),
    ...body,
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
  if (col !== MIN_COL_INDEX && col !== MAX_COL_INDEX) return false
  return isInvalidRow(row)
}

function sortRows() {
  const order = { 0: 0, 2: 1, 1: 2 }
  const body = activeSheet.value.data
    .slice(DATA_START_ROW)
    .filter(r => r.some(v => v))

  body.sort((a, b) => order[a[TYPE_COL_INDEX]] - order[b[TYPE_COL_INDEX]])

  activeSheet.value.data = [
    ...activeSheet.value.data.slice(0, DATA_START_ROW),
    ...body,
    new Array(COL_COUNT).fill("")
  ]
}

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
  isLoading.value = true
  try {
    const res = await axios.get(
      "https://excelproxy.kin169999.workers.dev/api/excel",
      { responseType: "arraybuffer" }
    )
    sheets.value = parseExcel(res.data)
    switchSheet(0)
  } finally {
    isLoading.value = false
  }
}

async function loadFromUrl() {
  if (!excelUrl.value) return
  isLoading.value = true
  try {
    const res = await axios.get(excelUrl.value, {
      responseType: "arraybuffer"
    })
    sheets.value = parseExcel(res.data)
    switchSheet(0)
  } finally {
    isLoading.value = false
  }
}

function uploadExcel(e) {
  const file = e.target.files[0]
  if (!file) return
  isLoading.value = true

  const reader = new FileReader()
  reader.onload = evt => {
    sheets.value = parseExcel(evt.target.result)
    switchSheet(0)
    isLoading.value = false
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

/* ===== 整體 ===== */
.layout {
  height: 100vh;
  display: flex;
  flex-direction: column;
  background: #020617;
  color: #e5e7eb;
  font-family: "Segoe UI", sans-serif;
}

/* ===== Topbar ===== */
.topbar {
  height: 60px;
  border-bottom: 1px solid #1e293b;
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 20px;
  background: linear-gradient(180deg, #0f172a, #020617);
}

.topbar h1 {
  font-size: 20px;
  font-weight: 600;
  letter-spacing: 1px;
}

.actions {
  display: flex;
  gap: 10px;
  align-items: center;
}

.actions input {
  background: #0f172a;
  border: 1px solid #334155;
  color: white;
  padding: 6px 8px;
  border-radius: 6px;
}

.actions button {
  padding: 6px 10px;
  border-radius: 6px;
  border: none;
  cursor: pointer;
  font-weight: 500;
  transition: 0.2s;
}

.actions button:hover {
  transform: translateY(-1px);
}

.actions button.export {
  background: #16a34a;
  color: white;
}

.actions button.cloud {
  background: #0ea5e9;
  color: white;
}

/* ===== Body ===== */
.body {
  flex: 1;
  display: flex;
  overflow: hidden;
}

/* ===== Sidebar ===== */
.sidebar {
  width: 220px;
  border-right: 1px solid #1e293b;
  padding: 10px;
  background: #0b1220;
}

.sheet-btn {
  padding: 10px;
  margin-bottom: 6px;
  border-radius: 6px;
  cursor: pointer;
  transition: 0.2s;
}

.sheet-btn:hover {
  background: #1e293b;
}

.sheet-btn.active {
  background: #2563eb;
  color: white;
}

/* ===== Main ===== */
.main {
  flex: 1;
  padding: 10px;
}

.table-wrap {
  height: 100%;
  overflow: auto;
  border: 1px solid #1e293b;
  border-radius: 10px;
  background: #020617;
}

/* ===== Table ===== */
table {
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
  font-size: 14px;
}

thead th {
  position: sticky;
  top: 0;
  background: linear-gradient(180deg, #0f172a, #020617);
  padding: 10px 6px;
  font-weight: 600;
  border-bottom: 1px solid #1e293b;
}

/* 交錯行 */
tbody tr:nth-child(even) td {
  background: rgba(255,255,255,0.04);
}

tbody tr:nth-child(odd) td {
  background: #020617;
}

/* 錯誤整行 */
tr.row-error td {
  background: rgba(220, 38, 38, 0.25) !important;
}

/* ===== Cell ===== */
td {
  border-bottom: 1px solid #1e293b;
  padding: 6px 6px;
  vertical-align: middle;
}

/* 編輯框提示 */
td.cell-editing {
  box-shadow: inset 0 0 0 2px #facc15;
}

/* 錯誤格 */
td.error {
  outline: 2px solid #ef4444;
}

/* ===== Inputs ===== */
.select,
.number-input {
  width: 100%;
  background: #0f172a;
  color: white;
  border: 1px solid #334155;
  padding: 4px 6px;
  border-radius: 6px;
}

.editable {
  min-height: 22px;
  outline: none;
}

/* ===== 拖曳 ===== */
.drag {
  cursor: grab;
  text-align: center;
  color: #64748b;
  width: 30px;
}

.drag:hover {
  color: #94a3b8;
}

/* ===== 刪除按鈕 ===== */
.delete-cell {
  text-align: center;
}

.delete-btn {
  background: none;
  border: none;
  cursor: pointer;
  font-size: 16px;
  color: #64748b;
  transition: 0.2s;
}

.delete-btn:hover {
  color: #ef4444;
  transform: scale(1.1);
}

/* ===== Scrollbar 美化 ===== */
.table-wrap::-webkit-scrollbar {
  width: 8px;
  height: 8px;
}

.table-wrap::-webkit-scrollbar-thumb {
  background: #334155;
  border-radius: 4px;
}

.table-wrap::-webkit-scrollbar-track {
  background: #020617;
}

</style>

