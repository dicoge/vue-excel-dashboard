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
                <th v-for="h in DISPLAY_HEADERS" :key="h">
                  {{ h }}
                </th>
              </tr>
            </thead>

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
                  :class="{ error: isInvalidCell(row, c) }"
                >

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

                  <input
                    v-else-if="c === MIN_COL_INDEX || c === MAX_COL_INDEX"
                    type="number"
                    min="0"
                    class="number-input"
                    v-model.number="rows[r][c]"
                  />

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

/* ============================= */

const CLOUD_API =
  "https://excelproxy.kin169999.workers.dev/"

/* ============================= */

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

/* ============================= */

const excelUrl = ref("")
const sheets = ref([])
const activeSheetIndex = ref(0)

/* ============================= */

async function loadFromUrl() {
  if (!excelUrl.value) return

  const res = await axios.get(excelUrl.value, {
    responseType: "arraybuffer"
  })

  sheets.value = parseExcel(res.data)
  activeSheetIndex.value = 0
}

async function loadFromCloud() {
  try {
    const res = await axios.get(CLOUD_API, {
      responseType: "arraybuffer"
    })

    sheets.value = parseExcel(res.data)
    activeSheetIndex.value = 0
  } catch (err) {
    console.error(err)
    alert("雲端 Excel 讀取失敗")
  }
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

/* ============================= */

const activeSheet = computed(() => {
  return sheets.value[activeSheetIndex.value]
})

const rows = computed(() => {
  if (!activeSheet.value) return []

  return activeSheet.value.data
    .slice(DATA_START_ROW)
    .map(row => row.slice(0, COL_COUNT))
})

/* ============================= */

function isInvalidRow(row) {
  const min = Number(row[MIN_COL_INDEX])
  const max = Number(row[MAX_COL_INDEX])
  return !isNaN(min) && !isNaN(max) && min > max
}

function isInvalidCell(row, col) {
  if (col !== MIN_COL_INDEX && col !== MAX_COL_INDEX) return false
  return isInvalidRow(row)
}

function updateCell(row, col, e) {
  rows.value[row][col] = e.target.innerText
}

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
