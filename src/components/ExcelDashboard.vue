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
                <th width="28"></th>
                <th v-for="h in DISPLAY_HEADERS" :key="h">
                  {{ h }}
                </th>
                <th class="action-col">操作</th>
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
                :class="isInvalidRow(row) ? 'row-error' : ''"
              >
                <!-- 拖曳 -->
                <td class="drag">⋮⋮</td>

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
                    v-model.number="rows[r][c]"
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
                    v-model="rows[r][c]"
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
                    v-model.number="rows[r][c]"
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

                <!-- 刪除 -->
                <td class="delete-cell">
                  <button
                    v-if="!isEmptyRow(row)"
                    class="delete-btn"
                    @click="deleteRow(r)"
                    title="刪除此行"
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
/* JS 保持原本邏輯不變 */
</script>

<style scoped>

/* ===== 操作欄寬度 ===== */
.action-col {
  width: 50px;
  text-align: center;
}

/* ===== 刪除按鈕優化 ===== */
.delete-cell {
  text-align: center;
}

.delete-btn {
  background: transparent;
  border: none;
  font-size: 18px;
  cursor: pointer;
  color: #94a3b8;
  transition: all 0.2s ease;
  padding: 6px;
  border-radius: 6px;
}

.delete-btn:hover {
  color: #ef4444;
  background: rgba(239, 68, 68, 0.15);
  transform: scale(1.1);
}

.drag {
  cursor: grab;
  text-align: center;
  color: #64748b;
}

</style>
