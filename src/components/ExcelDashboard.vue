<template>
  <div class="layout">

    <!-- ===== Top Bar ===== -->
    <header class="topbar">
      <h1>🐟 魚種圖鑑 Excel 管理系統</h1>

      <div class="actions">
        <input v-model="excelUrl" placeholder="輸入 Excel 網址" />
        <button @click="loadFromUrl">用網址載入</button>

        <input type="file" accept=".xlsx,.xls" @change="uploadExcel" />

        <button class="export" @click="exportExcel">
          匯出 Excel
        </button>
      </div>
    </header>

    <div class="body">

      <!-- ===== Sidebar ===== -->
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

      <!-- ===== Main ===== -->
      <main class="main" v-if="activeSheet">

        <div class="table-wrap">

          <table>

            <!-- ===== Header ===== -->
            <thead>
              <tr>
                <th v-for="h in DISPLAY_HEADERS" :key="h">
                  {{ h }}
                </th>
              </tr>
            </thead>

            <!-- ===== Body ===== -->
            <tbody>
              <tr
                v-for="(row, r) in rows"
                :key="r"
                :class="[
                  'row-' + (r % 2),
                  isInvalidRow(row) ? 'row-error' : ''
                ]"
              >
                <td
                  v-for="(cell, c) in row"
                  :key="c"
                  :class="{
                    error: isInvalidCell(row, c)
                  }"
                >

                  <!-- 🔽 類型 -->
                  <select
                    v-if="c === TYPE_COL_INDEX"
                    v-model.number="rows[r][c]"
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

                  <!-- 🔽 標題 -->
                  <select
                    v-else-if="c === TITLE_COL_INDEX"
                    v-model="rows[r][c]"
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

                  <!-- 🔢 倍率 -->
                  <input
                    v-else-if="c === MIN_COL_INDEX || c === MAX_COL_INDEX"
                    type="number"
                    min="0"
                    step="1"
                    class="number-input"
                    v-model.number="rows[r][c]"
                  />

                  <!-- ✏ 其他 -->
                  <div
                    v-else
                    contenteditable
                    class="editable"
                    @input="updateCell(r, c, $event)"
                  >
                    {{ cell }}
                  </div>

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

/* =============================
   固定設定
============================= */

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

/* =============================
   狀態
============================= */

const excelUrl = ref("")
const sheets = ref([])
const activeSheetIndex = ref(0)

/* =============================
   載入
============================= */

async function loadFromUrl() {
  if (!excelUrl.value) return

  const res = await axios.get(excelUrl.value, {
    responseType: "arraybuffer"
  })

  sheets.value = parseExcel(res.data)
  activeSheetIndex.value = 0
}

function uploadExcel(e) {
  const file = e.target.files[0]
  if (!file) return

  const reader = new FileReader()
  reader.onload = evt => {
    sheets.value = parseExcel(evt.target.result)
    activeSheetIndex.value = 0
  }
  reader.readAsArrayBuffer(file)
}

/* =============================
   計算
============================= */

const activeSheet = computed(() => {
  return sheets.value[activeSheetIndex.value]
})

const rows = computed(() => {
  if (!activeSheet.value) return []

  return activeSheet.value.data
    .slice(DATA_START_ROW)
    .map(row => row.slice(0, COL_COUNT))
})

/* =============================
   驗證
============================= */

function isInvalidRow(row) {
  const min = Number(row[MIN_COL_INDEX])
  const max = Number(row[MAX_COL_INDEX])

  return !isNaN(min) && !isNaN(max) && min > max
}

function isInvalidCell(row, col) {
  if (col !== MIN_COL_INDEX && col !== MAX_COL_INDEX) return false
  return isInvalidRow(row)
}

/* =============================
   編輯
============================= */

function updateCell(row, col, e) {
  rows.value[row][col] = e.target.innerText
}

/* =============================
   匯出 Excel
============================= */

function exportExcel() {
  const wb = XLSX.utils.book_new()

  sheets.value.forEach(sheet => {
    const data = sheet.data.map((r, i) => {
      if (i < DATA_START_ROW) return r
      return r.slice(0, COL_COUNT)
    })

    const ws = XLSX.utils.aoa_to_sheet(data)
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

/* ===== Topbar ===== */

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

.actions button.export {
  background: #16a34a;
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

/* ===== Table ===== */

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
  table-layout: fixed;
}

thead th {
  position: sticky;
  top: 0;
  background: linear-gradient(180deg, #0f172a, #020617);
  padding: 10px;
  font-weight: 700;
}

/* ===== Row 交錯 ===== */

tr.row-0 td {
  background: #020617;
}

tr.row-1 td {
  background: rgba(255, 255, 255, 0.06);
}

tr.row-error td {
  background: rgba(220, 38, 38, 0.18) !important;
}

/* ===== Cell ===== */

td {
  border-bottom: 1px solid #1e293b;
  padding: 6px;
}

td.error {
  outline: 2px solid #dc2626;
}

/* ===== Inputs ===== */

.select,
.number-input {
  width: 100%;
  background: #020617;
  color: white;
  border: 1px solid #334155;
  padding: 4px;
  border-radius: 4px;
}

.editable {
  min-height: 22px;
  outline: none;
}
</style>
