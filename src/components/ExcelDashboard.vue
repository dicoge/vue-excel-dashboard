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
                <th v-for="h in DISPLAY_HEADERS" :key="h">
                  {{ h }}
                </th>
              </tr>
            </thead>

            <tbody>
              <tr
                v-for="(row, r) in editableRows"
                :key="r"
                :class="[
                  'row-' + (r % 2),
                  isInvalidRow(row) ? 'row-error' : ''
                ]"
              >
                <td v-for="(cell, c) in row" :key="c">

                  <!-- 類型 -->
                  <select
                    v-if="c === TYPE_COL_INDEX"
                    v-model.number="editableRows[r][c]"
                    @change="checkAppendRow"
                    class="select"
                  >
                    <option
                      v-for="opt in TYPE_OPTIONS"
                      :key="opt.value"
                      :value="opt.value"
                    >
                      {{ opt.label }}
                    </option>
                  </select>

                  <!-- 標題 -->
                  <select
                    v-else-if="c === TITLE_COL_INDEX"
                    v-model="editableRows[r][c]"
                    @change="checkAppendRow"
                    class="select"
                  >
                    <option
                      v-for="opt in TITLE_OPTIONS"
                      :key="opt.value"
                      :value="opt.value"
                    >
                      {{ opt.label }}
                    </option>
                  </select>

                  <!-- 數字 -->
                  <input
                    v-else-if="c === MIN_COL_INDEX || c === MAX_COL_INDEX"
                    type="number"
                    v-model.number="editableRows[r][c]"
                    @input="checkAppendRow"
                    class="number-input"
                  />

                  <!-- 文字 -->
                  <input
                    v-else
                    v-model="editableRows[r][c]"
                    @input="checkAppendRow"
                    class="text-input"
                  />

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
  { value: 1, label: "活動魚" },
  { value: 2, label: "Boss" }
]

const TITLE_OPTIONS = [
  { value: "NONE", label: "無" },
  { value: "J", label: "金蟬大獎" }
]

const excelUrl = ref("")
const sheets = ref([])
const activeSheetIndex = ref(0)
const editableRows = ref([])

/* ===========================
   載入
=========================== */

function switchSheet(idx) {
  activeSheetIndex.value = idx
  prepareEditableRows()
}

function prepareEditableRows() {
  const sheet = sheets.value[activeSheetIndex.value]
  if (!sheet) return

  editableRows.value = sheet.data
    .slice(DATA_START_ROW)
    .map(r => r.slice(0, COL_COUNT))

  appendEmptyRowIfNeeded()
}

function appendEmptyRowIfNeeded() {
  const last = editableRows.value.at(-1)
  if (!last || last.some(v => v !== "" && v != null)) {
    editableRows.value.push(new Array(COL_COUNT).fill(""))
  }
}

function checkAppendRow() {
  const last = editableRows.value.at(-1)
  if (last && last.some(v => v !== "" && v != null)) {
    editableRows.value.push(new Array(COL_COUNT).fill(""))
  }
}

async function loadFromCloud() {
  const res = await axios.get(
    "https://excelproxy.kin169999.workers.dev/api/excel",
    { responseType: "arraybuffer" }
  )
  sheets.value = parseExcel(res.data)
  switchSheet(0)
}

function uploadExcel(e) {
  const file = e.target.files[0]
  if (!file) return
  const reader = new FileReader()
  reader.onload = evt => {
    sheets.value = parseExcel(evt.target.result)
    switchSheet(0)
  }
  reader.readAsArrayBuffer(file)
}

const activeSheet = computed(() => sheets.value[activeSheetIndex.value])

/* ===========================
   驗證
=========================== */

function isInvalidRow(row) {
  const min = Number(row[MIN_COL_INDEX])
  const max = Number(row[MAX_COL_INDEX])
  return !isNaN(min) && !isNaN(max) && min > max
}

/* ===========================
   匯出
=========================== */

function exportExcel() {
  const wb = XLSX.utils.book_new()

  sheets.value.forEach((sheet, i) => {
    const original = sheet.data.slice(0, DATA_START_ROW)
    const merged = [...original, ...editableRows.value]

    const ws = XLSX.utils.aoa_to_sheet(merged)
    XLSX.utils.book_append_sheet(wb, ws, sheet.name)
  })

  XLSX.writeFile(wb, "fish_data_export.xlsx")
}
</script>

<style scoped>
.layout {
  height: 100vh;
  display: flex;
  flex-direction: column;
  background: #020617;
  color: #e5e7eb;
}

.topbar {
  height: 60px;
  border-bottom: 1px solid #1e293b;
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 20px;
}

.actions {
  display: flex;
  gap: 10px;
}

.body {
  flex: 1;
  display: flex;
  overflow: hidden;
}

.sidebar {
  width: 220px;
  border-right: 1px solid #1e293b;
  padding: 10px;
}

.sheet-btn {
  padding: 10px;
  margin-bottom: 6px;
  border-radius: 6px;
  cursor: pointer;
}

.sheet-btn.active {
  background: #2563eb;
}

.main {
  flex: 1;
  padding: 10px;
}

.table-wrap {
  height: 100%;
  overflow: auto;
  border: 1px solid #1e293b;
  border-radius: 8px;
}

table {
  width: 100%;
  border-collapse: collapse;
}

td {
  padding: 6px;
  border-bottom: 1px solid #1e293b;
}

.text-input,
.number-input,
.select {
  width: 100%;
  min-width: 120px;
  background: #020617;
  color: white;
  border: 1px solid #334155;
  padding: 4px;
  border-radius: 4px;
}
</style>
