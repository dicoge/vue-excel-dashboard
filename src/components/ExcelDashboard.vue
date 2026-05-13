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
      <div class="topbar-left">
        <h1>📊 Excel 管理系統</h1>
        <select v-model="currentProjectId" class="project-select" @change="switchProject">
          <option v-for="(cfg, id) in PROJECTS" :key="id" :value="id">
            {{ cfg.label }}
          </option>
        </select>
      </div>

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
                <th width="30" v-if="currentConfig.enableDrag"></th>
                <th v-for="h in currentConfig.displayHeaders" :key="h">
                  {{ h }}
                </th>
                <th width="60">操作</th>
              </tr>
            </thead>

            <tbody>
              <tr
                v-for="(row, r) in rows"
                :key="r"
                :draggable="currentConfig.enableDrag && !isEmptyRow(row)"
                @dragstart="currentConfig.enableDrag ? dragStart(r) : null"
                @dragover.prevent
                @drop="currentConfig.enableDrag ? dropRow(r) : null"
                :class="isInvalidRow(row) ? 'row-error' : ''"
              >
                <td class="drag" v-if="currentConfig.enableDrag && !isEmptyRow(row)">⋮⋮</td>
                <td v-else-if="currentConfig.enableDrag"></td>

                <td
                  v-for="(cell, c) in row"
                  :key="c"
                  :class="[
                    hasCellError(row, c) ? 'error' : '',
                    editingCell.row === r && editingCell.col === c ? 'cell-editing' : ''
                  ]"
                >
                  <!-- 魚種-類型下拉 (fish project) -->
                  <select
                    v-if="c === currentConfig.typeColIndex && PROJECTS[currentProjectId].id === 'fish'"
                    class="select"
                    :value="getCellValue(r, c)"
                    @focus="setEditing(r, c)"
                    @change="setCellValue(r, c, $event.target.value); sortRows()"
                  >
                    <option
                      v-for="opt in currentConfig.typeOptions"
                      :key="opt.value"
                      :value="opt.value"
                    >
                      {{ opt.label }}
                    </option>
                  </select>

                  <!-- 魚種-標題下拉 (fish project) -->
                  <select
                    v-else-if="c === currentConfig.titleColIndex && PROJECTS[currentProjectId].id === 'fish'"
                    class="select"
                    :value="getCellValue(r, c)"
                    @focus="setEditing(r, c)"
                    @change="setCellValue(r, c, $event.target.value)"
                  >
                    <option
                      v-for="opt in currentConfig.titleOptions"
                      :key="opt.value"
                      :value="opt.value"
                    >
                      {{ opt.label }}
                    </option>
                  </select>

                  <!-- 數字輸入欄位 -->
                  <input
                    v-else-if="isNumberCol(c)"
                    type="number"
                    class="number-input"
                    :class="{ 'value-error': isValueCol(c) && !isValidValue(cell) }"
                    :value="cell"
                    @input="onNumberInput(r, c, $event)"
                    @focus="setEditing(r, c)"
                  />

                  <!-- 一般可編輯欄位 -->
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

// ===== 專案設定 =====
const PROJECTS = {
  fish: {
    id: 'fish',
    label: "🐟 魚種圖鑑",
    dataStartRow: 5,
    displayHeaders: ["魚種類型", "魚種名稱", "最小倍率", "最高倍率", "Tag", "標題", "類型"],
    visibleCols: [0, 1, 2, 3, 4, 5, 6],
    numberCols: [2, 3],        // visible indices 最小倍率, 最高倍率
    minColIndex: 2,            // visible index for 最小倍率
    maxColIndex: 3,            // visible index for 最高倍率
    typeColIndex: 6,           // visible index for 類型
    titleColIndex: 5,          // visible index for 標題
    typeOptions: [
      { value: 0, label: "一般魚" },
      { value: 2, label: "Boss" },
      { value: 1, label: "活動魚" }
    ],
    titleOptions: [
      { value: "NONE", label: "無" },
      { value: "J", label: "金蟬大獎" }
    ],
    enableDrag: true,
    enableSort: true
  },
  item: {
    id: 'item',
    label: "🎒 道具表",
    dataStartRow: 4,
    displayHeaders: ["物品ID", "ItemSID", "名稱", "道具價值", "類型", "說明"],
    visibleCols: [0, 1, 2, 5, 7, 19],  // col 19 = 說明(Tip), col 20 = 說明2(Tip2=全部NONE)
    numberCols: [0, 3],        // visible indices: 0=物品ID, 3=道具價值
    valueColIndex: 3,          // visible index for 道具價值 (需數字驗證)
    enableDrag: false,
    enableSort: false
  }
}

const currentProjectId = ref("fish")

const excelUrl = ref("")
const sheets = ref([])
const activeSheetIndex = ref(0)
const editingCell = ref({ row: null, col: null })
const dragIndex = ref(null)
const isLoading = ref(false)

const currentConfig = computed(() => PROJECTS[currentProjectId.value])
const activeSheet = computed(() => sheets.value[activeSheetIndex.value])

// ===== 資料行（根據 visibleCols 映射） =====
const rows = computed(() => {
  if (!activeSheet.value) return []
  const cfg = currentConfig.value
  return activeSheet.value.data
    .slice(cfg.dataStartRow)
    .map(row => {
      // 只取 visibleCols 對應的欄位
      return cfg.visibleCols.map(colIdx => {
        const val = row[colIdx]
        return val !== undefined && val !== null ? val : ""
      })
    })
})

function getCellValue(rowIdx, visibleColIdx) {
  const cfg = currentConfig.value
  const dataRow = cfg.dataStartRow + rowIdx
  const actualCol = cfg.visibleCols[visibleColIdx]
  return activeSheet.value.data[dataRow]?.[actualCol] ?? ""
}

function setCellValue(rowIdx, visibleColIdx, value) {
  const cfg = currentConfig.value
  const dataRow = cfg.dataStartRow + rowIdx
  const actualCol = cfg.visibleCols[visibleColIdx]
  activeSheet.value.data[dataRow][actualCol] = value
  ensureEmptyRow()
}

function isNumberCol(visibleIdx) {
  return currentConfig.value.numberCols.includes(visibleIdx)
}

function isValueCol(visibleIdx) {
  return currentConfig.value.valueColIndex !== undefined &&
         visibleIdx === currentConfig.value.valueColIndex
}

function isValidValue(val) {
  if (val === "" || val === null || val === undefined) return true
  const n = Number(val)
  return !isNaN(n) && isFinite(n)
}

function onNumberInput(r, c, e) {
  const cfg = currentConfig.value
  const dataRow = cfg.dataStartRow + r
  const actualCol = cfg.visibleCols[c]
  const raw = e.target.value
  // Allow empty or valid number
  if (raw === "" || raw === "-") {
    activeSheet.value.data[dataRow][actualCol] = raw
  } else {
    const n = Number(raw)
    if (!isNaN(n)) {
      activeSheet.value.data[dataRow][actualCol] = n
    }
    // If invalid, don't update (field stays as-is)
  }
  ensureEmptyRow()
}

// ===== 工作表切換 =====
function switchSheet(index) {
  activeSheetIndex.value = index
  ensureEmptyRow()
}

function ensureEmptyRow() {
  if (!activeSheet.value) return
  const cfg = currentConfig.value
  const data = activeSheet.value.data
  const body = data.slice(cfg.dataStartRow).filter(r => r.some(v => v && String(v).trim() !== ""))
  activeSheet.value.data = [
    ...data.slice(0, cfg.dataStartRow),
    ...body,
    new Array(cfg.colCount).fill("")
  ]
}

function deleteRow(index) {
  const cfg = currentConfig.value
  activeSheet.value.data.splice(cfg.dataStartRow + index, 1)
  ensureEmptyRow()
}

function isEmptyRow(visibleRow) {
  return visibleRow.every(v => !v || String(v).trim() === "")
}

function updateCell(row, col, e) {
  const cfg = currentConfig.value
  const dataRow = cfg.dataStartRow + row
  const actualCol = cfg.visibleCols[col]
  activeSheet.value.data[dataRow][actualCol] = e.target.innerText
  ensureEmptyRow()
}

// ===== 魚種專用：驗證 最小倍率 <= 最高倍率 =====
function isInvalidRow(row) {
  const cfg = currentConfig.value
  if (!cfg.enableDrag) return false  // only fish project has this validation
  const min = Number(row[cfg.minColIndex])
  const max = Number(row[cfg.maxColIndex])
  return !isNaN(min) && !isNaN(max) && min > max
}

function hasCellError(row, col) {
  const cfg = currentConfig.value
  // Fish: min/max ratio validation
  if (cfg.enableDrag && (col === cfg.minColIndex || col === cfg.maxColIndex)) {
    return isInvalidRow(row)
  }
  // Item: 道具價值 must be valid number
  if (isValueCol(col)) {
    return !isValidValue(row[col])
  }
  return false
}

// ===== 魚種專用：排序 =====
function sortRows() {
  if (!currentConfig.value.enableSort) return
  const cfg = currentConfig.value
  const order = { 0: 0, 2: 1, 1: 2 }
  const body = activeSheet.value.data
    .slice(cfg.dataStartRow)
    .filter(r => r.some(v => v))

  body.sort((a, b) => order[a[cfg.visibleCols[cfg.typeColIndex]]] - order[b[cfg.visibleCols[cfg.typeColIndex]]])

  activeSheet.value.data = [
    ...activeSheet.value.data.slice(0, cfg.dataStartRow),
    ...body,
    new Array(cfg.colCount).fill("")
  ]
}

// ===== 魚種專用：拖曳排序 =====
function dragStart(index) {
  dragIndex.value = index
}

function dropRow(targetIndex) {
  if (dragIndex.value === null) return
  const cfg = currentConfig.value
  const from = cfg.dataStartRow + dragIndex.value
  const to = cfg.dataStartRow + targetIndex
  const data = activeSheet.value.data
  const moved = data.splice(from, 1)[0]
  data.splice(to, 0, moved)
  dragIndex.value = null
}

// ===== 載入 =====
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
  XLSX.writeFile(wb, "data_export.xlsx")
}

function switchProject() {
  sheets.value = []
  activeSheetIndex.value = 0
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

.topbar-left {
  display: flex;
  align-items: center;
  gap: 12px;
}

.topbar h1 {
  font-size: 18px;
  font-weight: 600;
  letter-spacing: 1px;
  white-space: nowrap;
}

.project-select {
  background: #0f172a;
  border: 1px solid #334155;
  color: white;
  padding: 6px 10px;
  border-radius: 6px;
  font-size: 14px;
  cursor: pointer;
}

.project-select:hover {
  border-color: #2563eb;
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

.number-input.value-error {
  border-color: #ef4444;
  outline: 1px solid #ef4444;
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