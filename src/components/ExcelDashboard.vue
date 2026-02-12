<template>
  <div class="layout">

    <header class="topbar">
      <h1>🐟 魚種圖鑑 Excel 管理系統</h1>
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
                <th width="32"></th>
                <th v-for="h in DISPLAY_HEADERS" :key="h">
                  {{ h }}
                </th>
                <th width="48">操作</th>
              </tr>
            </thead>

            <tbody>
              <tr
                v-for="(row, r) in rows"
                :key="r"
                draggable="true"
                @dragstart="dragStart(r)"
                @dragover.prevent
                @drop="dropRow(r)"
                :class="[
                  isInvalidRow(row) ? 'row-error' : ''
                ]"
              >

                <!-- 拖曳 icon -->
                <td class="drag-handle">⋮⋮</td>

                <td
                  v-for="(cell, c) in row"
                  :key="c"
                  :class="editingCell.row === r && editingCell.col === c
                    ? 'cell-editing'
                    : ''"
                >

                  <select
                    v-if="c === TYPE_COL_INDEX"
                    v-model.number="rows[r][c]"
                    @change="sortRows"
                    @focus="setEditing(r, c)"
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

                  <input
                    v-else
                    v-model="rows[r][c]"
                    @focus="setEditing(r, c)"
                    class="text-input"
                  />

                </td>

                <td class="action-cell">
                  <button
                    class="delete-btn"
                    @click="deleteRow(r)"
                  >
                    <!-- SVG 垃圾桶 -->
                    <svg
                      viewBox="0 0 24 24"
                      width="16"
                      height="16"
                      fill="currentColor"
                    >
                      <path d="M9 3h6l1 2h4v2H4V5h4l1-2zm1 6h2v9h-2V9zm4 0h2v9h-2V9zM7 9h2v9H7V9z"/>
                    </svg>
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

const TYPE_COL_INDEX = 0

const TYPE_OPTIONS = [
  { value: 0, label: "一般魚" },
  { value: 2, label: "Boss" },
  { value: 1, label: "活動魚" }
]

const sheets = ref([])
const activeSheetIndex = ref(0)
const editingCell = ref({ row: null, col: null })
const dragIndex = ref(null)

const activeSheet = computed(() =>
  sheets.value[activeSheetIndex.value]
)

const rows = computed(() => {
  if (!activeSheet.value) return []
  return activeSheet.value.data
    .slice(DATA_START_ROW)
    .filter(r => r.some(v => v))
    .map(r => r.slice(0, COL_COUNT))
})

function setEditing(r, c) {
  editingCell.value = { row: r, col: c }
}

function dragStart(index) {
  dragIndex.value = index
}

function dropRow(targetIndex) {
  if (dragIndex.value === null) return
  const moved = rows.value.splice(dragIndex.value, 1)[0]
  rows.value.splice(targetIndex, 0, moved)
  dragIndex.value = null
  syncToSheet()
}

function deleteRow(index) {
  rows.value.splice(index, 1)
  syncToSheet()
}

function sortRows() {
  rows.value.sort((a, b) => {
    const order = { 0: 0, 2: 1, 1: 2 }
    return order[a[TYPE_COL_INDEX]] - order[b[TYPE_COL_INDEX]]
  })
  syncToSheet()
}

function syncToSheet() {
  const header = activeSheet.value.data.slice(0, DATA_START_ROW)
  activeSheet.value.data = [
    ...header,
    ...rows.value,
    new Array(COL_COUNT).fill("")
  ]
}

function isInvalidRow(row) {
  const min = Number(row[2])
  const max = Number(row[3])
  return !isNaN(min) && !isNaN(max) && min > max
}
</script>

<style scoped>
.layout {
  background:#020617;
  color:white;
  height:100vh;
  display:flex;
  flex-direction:column;
}

.body { flex:1; display:flex }

.sidebar {
  width:180px;
  border-right:1px solid #1e293b;
  padding:8px;
}

.main { flex:1; padding:8px }

table {
  width:100%;
  border-collapse:collapse;
  font-size:14px;
}

td, th {
  padding:4px 6px;
  border-bottom:1px solid #1e293b;
}

.row-error td {
  background:rgba(220,38,38,0.35);
}

.cell-editing {
  box-shadow: inset 0 0 0 2px #facc15;
}

.drag-handle {
  cursor:grab;
  color:#64748b;
  text-align:center;
  font-size:14px;
}

.action-cell {
  text-align:center;
}

.delete-btn {
  background:transparent;
  border:none;
  color:#94a3b8;
  cursor:pointer;
  padding:2px;
  border-radius:4px;
  transition:0.2s;
}

.delete-btn:hover {
  color:#ef4444;
  background:rgba(239,68,68,0.1);
}

.select,
.text-input {
  width:100%;
  background:#020617;
  color:white;
  border:1px solid #334155;
  padding:3px;
  border-radius:4px;
}
</style>
