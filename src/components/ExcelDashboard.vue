<script setup>
import { ref, computed, watch } from "vue"
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

/* ✅ 固定7欄 */
const rows = computed(() => {
  if (!activeSheet.value) return []

  return activeSheet.value.data
    .slice(DATA_START_ROW)
    .map(row => {
      const fixed = [...row]
      while (fixed.length < COL_COUNT) {
        fixed.push("")
      }
      return fixed.slice(0, COL_COUNT)
    })
})

/* ============================= */
/* 🔥 關鍵修正：切換 sheet 自動補空白 */
/* ============================= */

watch(activeSheetIndex, () => {
  setTimeout(() => {
    ensureEmptyRow()
  }, 0)
})

/* ============================= */

function ensureEmptyRow() {
  if (!activeSheet.value) return

  const data = activeSheet.value.data
  const body = data.slice(DATA_START_ROW)

  const filtered = body.filter(row =>
    row.some(v => v)
  )

  activeSheet.value.data = [
    ...data.slice(0, DATA_START_ROW),
    ...filtered,
    new Array(COL_COUNT).fill("")
  ]
}

/* ============================= */

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

/* ============================= */
/* 🔥 排序防呆 */
/* ============================= */

function sortRows() {
  if (!activeSheet.value) return

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

/* ============================= */
/* 🔥 拖曳修正 */
/* ============================= */

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

/* ============================= */

async function loadFromCloud() {
  const res = await axios.get(
    "https://excelproxy.kin169999.workers.dev/api/excel",
    { responseType: "arraybuffer" }
  )
  sheets.value = parseExcel(res.data)
  activeSheetIndex.value = 0
  ensureEmptyRow()
}

async function loadFromUrl() {
  if (!excelUrl.value) return
  const res = await axios.get(excelUrl.value, {
    responseType: "arraybuffer"
  })
  sheets.value = parseExcel(res.data)
  activeSheetIndex.value = 0
  ensureEmptyRow()
}

function uploadExcel(e) {
  const file = e.target.files[0]
  if (!file) return
  const reader = new FileReader()
  reader.onload = evt => {
    sheets.value = parseExcel(evt.target.result)
    activeSheetIndex.value = 0
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
