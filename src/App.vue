<script setup>
import { ref, computed } from 'vue'
import * as XLSX from 'xlsx'

const PREVIEW_ROWS = 200
const FORMATS = { xml: 'XML + CSV', markdown: 'Markdown', json: 'JSON' }

const fileName = ref('')
const sheets = ref([])
const active = ref(0)
const error = ref('')
const dragging = ref(false)
const format = ref('xml')
const copied = ref(false)

const current = computed(() => sheets.value[active.value])

function toCsv(rows) {
  return rows
    .map(row => row.map(cell => /[",\r\n]/.test(cell) ? `"${cell.replaceAll('"', '""')}"` : cell).join(','))
    .join('\n')
}

function toMarkdown(rows) {
  if (!rows.length) return '_empty_'
  const line = row => `| ${row.map(cell => cell.replaceAll('|', '\\|').replace(/\r?\n/g, '<br>')).join(' | ')} |`
  return [line(rows[0]), line(rows[0].map(() => '---')), ...rows.slice(1).map(line)].join('\n')
}

function toObjects([header = [], ...rows]) {
  const keys = header.map((key, i) => key || XLSX.utils.encode_col(i))
  return rows.map(row => Object.fromEntries(keys.map((key, i) => [key, row[i]])))
}

const output = computed(() => {
  if (format.value === 'json') {
    return JSON.stringify(Object.fromEntries(sheets.value.map(s => [s.name, toObjects(s.rows)])), null, 2)
  }
  if (format.value === 'markdown') {
    return sheets.value.map(s => `## ${s.name}\n\n${toMarkdown(s.rows)}`).join('\n\n')
  }
  const body = sheets.value
    .map(s => `<sheet name="${s.name}" rows="${s.rows.length}" cols="${s.rows[0]?.length ?? 0}">\n${toCsv(s.rows)}\n</sheet>`)
    .join('\n')
  return `<workbook name="${fileName.value}">\n${body}\n</workbook>`
})

async function copyOutput() {
  await navigator.clipboard.writeText(output.value)
  copied.value = true
  setTimeout(() => { copied.value = false }, 1800)
}

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
      class="flex flex-1 overflow-hidden"
      @dragover.prevent="dragging = true"
      @dragleave.prevent="dragging = false"
      @drop.prevent="onDrop"
    >
      <section class="flex flex-col flex-1 min-w-0 p-4 gap-3 border-r border-[#1e2130] overflow-hidden">
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
      </section>

      <section class="flex flex-col flex-1 min-w-0 overflow-hidden">
        <div class="flex items-center justify-between px-4 pt-4 pb-2 border-b border-[#1e2130] shrink-0">
          <div class="flex gap-1">
            <button
              v-for="(label, key) in FORMATS"
              :key="key"
              @click="format = key"
              :class="[
                'px-3 py-1 text-xs rounded border transition-colors',
                format === key
                  ? 'text-slate-200 border-[#3b4a6b] bg-[#1e2130]'
                  : 'text-slate-500 border-transparent hover:text-slate-300'
              ]"
            >{{ label }}</button>
          </div>
          <button
            v-if="sheets.length"
            @click="copyOutput"
            :class="['px-3 py-1 text-xs rounded border transition-colors', copied ? 'text-green-300 border-green-300/25' : 'text-slate-400 border-[#2e3a50] bg-[#1e2130] hover:text-slate-200 hover:border-[#3b4a6b]']"
          >{{ copied ? 'Copied' : 'Copy' }}</button>
        </div>

        <div class="flex flex-1 p-4 overflow-hidden">
          <p v-if="!sheets.length" class="text-sm text-slate-600">The converted output will appear here.</p>
          <textarea
            v-else
            readonly
            wrap="off"
            spellcheck="false"
            :value="output"
            class="flex-1 resize-none outline-none rounded-md p-3.5
                  bg-[#141720] text-slate-300 border border-[#1e2130]
                  font-mono text-[13px] leading-relaxed"
          />
        </div>
      </section>
    </main>
  </div>
</template>
