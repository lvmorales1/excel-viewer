<script setup>
import { ref, computed } from 'vue'
import * as XLSX from 'xlsx'

const PREVIEW_ROWS = 200

const fileName = ref('')
const sheets = ref([])
const active = ref(0)
const error = ref('')
const dragging = ref(false)

const current = computed(() => sheets.value[active.value])

async function load(file) {
  if (!file) return
  error.value = ''
  try {
    const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true })
    sheets.value = wb.SheetNames.map(name => ({
      name,
      rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: false }),
    }))
    fileName.value = file.name
    active.value = 0
  } catch (e) {
    error.value = e.message
  }
}

function onDrop(e) {
  dragging.value = false
  load(e.dataTransfer.files[0])
}
</script>

<template>
  <div class="flex flex-col h-screen bg-[#0f1117] text-slate-300">

    <header class="px-6 py-3.5 border-b border-[#1e2130] shrink-0">
      <h1 class="text-lg font-semibold text-slate-400 tracking-wide">Excel Viewer</h1>
    </header>

    <main
      class="flex flex-col flex-1 overflow-hidden p-4 gap-3"
      @dragover.prevent="dragging = true"
      @dragleave.prevent="dragging = false"
      @drop.prevent="onDrop"
    >
      <label
        :class="[
          'flex items-center justify-center shrink-0 rounded-md border border-dashed cursor-pointer transition-colors text-sm',
          sheets.length ? 'py-3' : 'flex-1',
          dragging ? 'border-[#3b4a6b] bg-[#1e2130] text-slate-200' : 'border-[#2e3a50] bg-[#141720] text-slate-500 hover:text-slate-300'
        ]"
      >
        <input type="file" accept=".xlsx,.xls,.xlsm,.ods,.csv" class="hidden" @change="load($event.target.files[0])" />
        <span v-if="fileName" class="font-mono">{{ fileName }} <span class="text-slate-600">· drop or click to replace</span></span>
        <span v-else>Drop a spreadsheet here or click to choose</span>
      </label>

      <p v-if="error" class="text-xs text-red-400 font-mono">{{ error }}</p>

      <template v-if="current">
        <div class="flex gap-1 overflow-x-auto shrink-0">
          <button
            v-for="(sheet, i) in sheets"
            :key="sheet.name"
            @click="active = i"
            :class="[
              'px-3 py-1 text-xs rounded border transition-colors whitespace-nowrap',
              active === i
                ? 'text-slate-200 border-[#3b4a6b] bg-[#1e2130]'
                : 'text-slate-500 border-transparent hover:text-slate-300'
            ]"
          >{{ sheet.name }} <span class="text-slate-600">{{ sheet.rows.length }}</span></button>
        </div>

        <div class="flex-1 overflow-auto rounded-md border border-[#1e2130]">
          <p v-if="!current.rows.length" class="text-sm text-slate-600 p-4">This sheet is empty.</p>
          <table v-else class="text-xs font-mono border-collapse">
            <tbody>
              <tr v-for="(row, r) in current.rows.slice(0, PREVIEW_ROWS)" :key="r" class="even:bg-[#141720]">
                <td class="px-2 py-1 text-slate-600 text-right border-r border-[#1e2130] sticky left-0 bg-[#0f1117]">{{ r + 1 }}</td>
                <td v-for="(cell, c) in row" :key="c" class="px-2 py-1 border-r border-[#1e2130] whitespace-nowrap max-w-xs truncate">{{ cell }}</td>
              </tr>
            </tbody>
          </table>
        </div>

        <p v-if="current.rows.length > PREVIEW_ROWS" class="text-xs text-slate-600 shrink-0">
          Showing {{ PREVIEW_ROWS }} of {{ current.rows.length }} rows
        </p>
      </template>
    </main>
  </div>
</template>
