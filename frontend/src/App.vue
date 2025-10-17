<template>
  <section class="page tools-page">
    <header class="page__header">
      <div>
        <h2>🛠️ Tilbudssammenligning</h2>
        <p class="page__subtitle">Sammenlign tilbud fra forskjellige leverandører</p>
      </div>
    </header>

    <div class="card">
      <h3>Slik bruker du verktøyet</h3>
      <ul>
        <li>Last opp tilbudsfiler i CSV, Excel (.xlsx/.xls) eller NS3459 XML format</li>
        <li>Du kan laste opp flere filer samtidig for å sammenligne tilbud</li>
        <li>Verktøyet genererer automatisk sammenligningsrapporter og Excel-eksport</li>
      </ul>
      <input
        type="file"
        multiple
        accept=".csv,.xlsx,.xls,.xml"
        @change="onFileSelected"
        class="file-input"
      />
      <button
        class="primary"
        :disabled="!files.length || isLoading"
        @click="runComparison"
      >
        {{ isLoading ? 'Behandler...' : 'Kjør sammenligning' }}
      </button>
      <div v-if="isLoading" class="progress">
        <div class="progress__track">
          <div class="progress__bar" :style="{ width: `${Math.min(progress, 100)}%` }"></div>
        </div>
        <span class="progress__value">{{ Math.round(progress) }}%</span>
      </div>
      <p v-if="errors.length" class="feedback feedback--error">
        <span v-for="err in errors" :key="err">{{ err }}</span>
      </p>
    </div>

    <div v-if="result" class="results">
      <div class="results-layout">
        <nav class="side-links">
          <h4>Hopp til</h4>
          <ul>
            <li><a href="#bid-summary">Oppsummering</a></li>
            <li><a href="#bid-chapters">Kapitteloppsummering</a></li>
            <li
              v-for="(rows, name) in result.normalized"
              :key="`link-${name}`"
            >
              <a :href="`#bid-${anchorId(name)}`">{{ name }}</a>
            </li>
            <li><a href="#bid-matrix">Sammenligning per postnr</a></li>
          </ul>
        </nav>

        <div class="results-content">
          <div id="bid-summary" class="card summary-card">
            <div class="card-header">
              <h3>Oppsummering</h3>
              <button class="secondary" @click="downloadExcel">
                Last ned Excel
              </button>
            </div>
            <div class="summary-grid">
              <div class="summary-item">
                <span class="summary-label">Antall poster</span>
                <span class="summary-value">{{ result.summary.post_count }}</span>
              </div>
              <div class="summary-item">
                <span class="summary-label">Laveste totale tilbud</span>
                <span class="summary-value">
                  {{ result.summary.winner.name }}
                  <small v-if="result.summary.winner.name">{{ formatAmount(result.summary.winner.total) }}</small>
                </span>
              </div>
              <div v-if="bestZScoreProvider.name" class="summary-item summary-item--highlight">
                <span class="summary-label">
                  Best Z-score
                  <span class="tooltip-container">
                    <span class="info-icon">ⓘ</span>
                    <span class="tooltip">
                      <strong>Z-score tolkning:</strong><br>
                      • Negativ z-score = Billigere enn gjennomsnittet (bedre) ✓<br>
                      • Z-score nær 0 = Nær gjennomsnittet<br>
                      • Positiv z-score = Dyrere enn gjennomsnittet (dårligere)<br><br>
                      Lavest total z-score = Mest konsekvent billig tilbud
                    </span>
                  </span>
                </span>
                <span class="summary-value">
                  {{ bestZScoreProvider.name }}
                  <small>Total z-score: {{ bestZScoreProvider.total.toFixed(2) }}</small>
                </span>
              </div>
              <div
                v-for="(total, name) in result.summary.totals"
                :key="name"
                class="summary-item"
              >
                <span class="summary-label">
                  {{ name }}
                  <span v-if="providerZScoreStats[name]" :class="getZScoreBadgeClass(providerZScoreStats[name].average)">
                    Z: {{ providerZScoreStats[name].average.toFixed(2) }}
                  </span>
                </span>
                <span class="summary-value">
                  {{ formatAmount(total) }}
                  <small v-if="optionTotals[name]">
                    + {{ formatAmount(optionTotals[name]) }} i opsjoner
                  </small>
                </span>
              </div>
            </div>
          </div>

          <div id="bid-chapters" class="card">
            <div class="card-header">
              <h3>Kapitteloppsummering</h3>
              <button
                class="secondary"
                type="button"
                @click="downloadChapterExcel"
                :disabled="!result.chapters.rows.length"
              >
                Last ned Kapitteloppsummering
              </button>
            </div>
            <div class="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th v-for="col in result.chapters.columns" :key="col">{{ formatChapterHeader(col) }}</th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="(row, idx) in result.chapters.rows" :key="`chap-${idx}`">
                    <td
                      v-for="col in result.chapters.columns"
                      :key="col"
                      :style="chapterCellStyle(col, row)"
                    >
                      {{ formatMatrixCell(col, row[col]) }}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <div
            v-for="(rows, name) in result.normalized"
            :key="name"
            class="card"
            :id="`bid-${anchorId(name)}`"
          >
            <h3>Normalisert tilbud: {{ name }}</h3>
            <div class="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th>Postnr</th>
                    <th>NS-kode</th>
                    <th>Spesifikasjon</th>
                    <th>Enhet</th>
                    <th>Mengde</th>
                    <th>Enhetspris</th>
                    <th>Sum</th>
                    <th>Kapittel</th>
                  </tr>
                </thead>
                <tbody>
                  <tr
                    v-for="(row, idx) in rows"
                    :key="`${name}-${idx}`"
                    :class="{ 'row-option': isOption(row) }"
                  >
                    <td>{{ row.postnr }}</td>
                    <td>{{ row.ns_code || '' }}</td>
                    <td class="spec-cell">{{ row.specification || row.beskrivelse }}</td>
                    <td>{{ row.enhet }}</td>
                    <td>{{ formatNumber(row.qty) }}</td>
                    <td class="option-amount">{{ formatAmountDisplay(row.unit_price, isOption(row)) }}</td>
                    <td class="option-amount">{{ formatAmountDisplay(row.sum_amount, isOption(row)) }}</td>
                    <td>{{ row.kapittel }}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <div id="bid-matrix" class="card">
            <div class="card-header">
              <div>
                <h3>Sammenligning per postnr</h3>
                <label v-if="hasZScoreColumns" class="advanced-toggle">
                  <input type="checkbox" v-model="advancedMode" />
                  <span>Vis Z-score (avansert)</span>
                </label>
                <p v-else-if="zScoreAvailableMessage" class="z-score-unavailable">
                  {{ zScoreAvailableMessage }}
                </p>
              </div>
              <button
                class="secondary"
                type="button"
                :disabled="!result.matrix_excel"
                @click="downloadMatrixExcel"
              >
                Last ned Matrise-Excel
              </button>
            </div>
            <div class="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th v-for="col in filteredMatrixColumns" :key="col" :style="matrixHeaderStyle(col)">
                      {{ formatMatrixHeader(col) }}
                    </th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="(row, idx) in result.matrix.rows" :key="`matrix-${idx}`">
                    <td v-for="col in filteredMatrixColumns" :key="col" :style="matrixCellStyle(col)">
                      {{ formatMatrixCell(col, row[col]) }}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </div>
  </section>
</template>

<script setup>
import { computed, ref, onBeforeUnmount } from 'vue'

const files = ref([])
const isLoading = ref(false)
const result = ref(null)
const errors = ref([])
const progress = ref(0)
const progressTimer = ref(null)
const advancedMode = ref(false)

const optionTotals = computed(
  () => result.value?.summary?.option_totals ?? {},
)

const matrixColumnColors = computed(() => {
  if (!result.value) return {}
  const totals = result.value.summary?.totals ? Object.keys(result.value.summary.totals) : []
  const palette = ['#bfdbfe', '#fde68a', '#e9d5ff', '#bbf7d0', '#fbcfe8', '#fecdd3', '#c7d2fe']
  const mapping = {}
  totals.forEach((name, index) => {
    const base = palette[index % palette.length]
    mapping[`${name} (enhetspris)`] = {
      header: withAlpha(base, 0.55),
      cell: withAlpha(base, 0.16),
    }
    mapping[`${name} (sum)`] = {
      header: withAlpha(base, 0.75),
      cell: withAlpha(base, 0.24),
    }
    mapping[`${name} (z-score)`] = {
      header: withAlpha(base, 0.45),
      cell: withAlpha(base, 0.12),
    }
  })
  return mapping
})

const filteredMatrixColumns = computed(() => {
  if (!result.value) return []
  const allColumns = result.value.matrix.columns
  if (advancedMode.value) {
    return allColumns
  }
  return allColumns.filter(col => !col.includes('(z-score)'))
})

const hasZScoreColumns = computed(() => {
  if (!result.value) return false
  return result.value.matrix.columns.some(col => col.includes('(z-score)'))
})

const zScoreAvailableMessage = computed(() => {
  if (hasZScoreColumns.value) {
    return ''
  }
  const numBids = result.value?.summary?.totals ? Object.keys(result.value.summary.totals).length : 0
  if (numBids < 3) {
    return 'Z-score krever minst 3 tilbud for å være meningsfull.'
  }
  return ''
})

const bestZScoreProvider = computed(() => {
  if (!result.value || !hasZScoreColumns.value) {
    return { name: '', total: 0 }
  }

  const zScoreColumns = result.value.matrix.columns.filter(col => col.includes('(z-score)'))
  const providerTotals = {}

  result.value.matrix.rows.forEach(row => {
    const postnr = row.postnr
    if (postnr === 'SUM') return

    zScoreColumns.forEach(col => {
      const provider = col.replace(' (z-score)', '')
      const zScore = Number(row[col])
      if (!Number.isNaN(zScore)) {
        providerTotals[provider] = (providerTotals[provider] || 0) + zScore
      }
    })
  })

  let bestProvider = ''
  let lowestTotal = Infinity

  for (const [provider, total] of Object.entries(providerTotals)) {
    if (total < lowestTotal) {
      lowestTotal = total
      bestProvider = provider
    }
  }

  return { name: bestProvider, total: lowestTotal }
})

const providerZScoreStats = computed(() => {
  if (!result.value || !hasZScoreColumns.value) {
    return {}
  }

  const zScoreColumns = result.value.matrix.columns.filter(col => col.includes('(z-score)'))
  const stats = {}

  result.value.matrix.rows.forEach(row => {
    const postnr = row.postnr
    if (postnr === 'SUM') return

    zScoreColumns.forEach(col => {
      const provider = col.replace(' (z-score)', '')
      const zScore = Number(row[col])
      if (!Number.isNaN(zScore)) {
        if (!stats[provider]) {
          stats[provider] = { total: 0, average: 0, count: 0 }
        }
        stats[provider].total += zScore
        stats[provider].count += 1
      }
    })
  })

  for (const provider in stats) {
    stats[provider].average = stats[provider].total / stats[provider].count
  }

  return stats
})

function getZScoreBadgeClass(average) {
  if (average < -0.5) return 'z-badge z-badge--good'
  if (average > 0.5) return 'z-badge z-badge--poor'
  return 'z-badge z-badge--neutral'
}

function anchorId(value) {
  const normalized = value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
  return normalized || 'tilbud'
}

function onFileSelected(event) {
  const target = event.target
  const list = target.files
  files.value = list ? Array.from(list) : []
  result.value = null
  errors.value = []
}

async function runComparison() {
  if (!files.value.length) return
  isLoading.value = true
  errors.value = []
  result.value = null
  startProgress()
  let succeeded = false

  const form = new FormData()
  for (const file of files.value) {
    form.append('files', file, file.name)
  }

  try {
    const response = await fetch('/api/bid-compare', {
      method: 'POST',
      body: form,
    })
    if (!response.ok) {
      const text = await response.text()
      throw new Error(text || 'Request failed')
    }
    const data = await response.json()
    result.value = data
    errors.value = data.errors || []
    succeeded = true
  } catch (err) {
    errors.value = [err.message ?? 'Unknown error']
  } finally {
    isLoading.value = false
    finishProgress(succeeded)
  }
}

function formatAmount(value) {
  if (value === null || value === undefined || value === '') return ''
  const num = typeof value === 'number' ? value : Number(value)
  if (Number.isNaN(num)) return ''
  return new Intl.NumberFormat('nb-NO', {
    style: 'currency',
    currency: 'NOK',
    maximumFractionDigits: 2,
  }).format(num)
}

function formatNumber(value) {
  if (value === null || value === undefined || value === '') return ''
  const num = typeof value === 'number' ? value : Number(value)
  if (Number.isNaN(num)) return ''
  return new Intl.NumberFormat('nb-NO', { maximumFractionDigits: 2 }).format(num)
}

function isOption(row) {
  const flag = row.is_option
  if (typeof flag === 'string') return flag.toLowerCase() === 'true'
  if (typeof flag === 'number') return flag !== 0
  return Boolean(flag)
}

function formatAmountDisplay(value, option) {
  const base = formatAmount(value)
  if (!base) return ''
  return option ? `(${base})` : base
}

function formatMatrixHeader(col) {
  const headerMap = {
    kapittel: 'Kapittel',
    kapittel_navn: 'Kapittelnavn',
    postnr: 'Postnr',
    ns_code: 'NS-kode',
    specification: 'Spesifikasjon',
    enhet: 'Enhet',
    qty: 'Mengde',
    laveste_tilbyder: 'Beste tilbyder',
    laveste_sum: 'Laveste sum',
    spann_pct: 'Spredning %',
    vinner: 'Vinner',
    lavest_sum: 'Laveste',
    std_avvik: 'Std.avvik',
    snitt: 'Snitt',
    std_pct: 'Std %',
  }

  if (col.includes('(z-score)')) {
    const provider = col.replace(' (z-score)', '')
    return `${provider} (Z-score)`
  }

  return headerMap[col] || col
}

function formatMatrixCell(col, value) {
  if (col === 'postnr' || col === 'kapittel' || col === 'kapittel_navn' || typeof value === 'string') {
    return value ?? ''
  }
  if (col === 'qty') return formatNumber(value)
  if (col === 'spann_pct') {
    const num = Number(value)
    if (Number.isNaN(num)) return ''
    return `${num.toFixed(2)} %`
  }
  if (col === 'std_pct') {
    const num = Number(value)
    if (Number.isNaN(num)) return ''
    return `${num.toFixed(2)} %`
  }
  if (col.includes('(z-score)')) {
    const num = Number(value)
    if (Number.isNaN(num)) return ''
    return num.toFixed(2)
  }
  return formatAmount(value)
}

function formatChapterHeader(col) {
  if (col === 'kapittel') return 'Kapittel'
  if (col === 'kapittel_navn') return 'Kapittelnavn'
  if (col === 'laveste_tilbyder') return 'Beste tilbyder'
  if (col === 'laveste_sum') return 'Laveste sum'
  if (col === 'spann_pct') return 'Spredning %'
  return col
}

function matrixHeaderStyle(col) {
  const colors = matrixColumnColors.value[col]
  if (!colors) return {}
  return {
    backgroundColor: colors.header,
    color: '#0F172A',
  }
}

function matrixCellStyle(col) {
  const colors = matrixColumnColors.value[col]
  if (!colors) return {}
  return {
    backgroundColor: colors.cell,
  }
}

function chapterCellStyle(col, row) {
  if (col !== 'laveste_sum' && col !== 'laveste_tilbyder') return {}
  const provider = String(row.laveste_tilbyder || '')
  const color = matrixColumnColors.value[`${provider} (sum)`]
  if (!provider || provider.toLowerCase() === 'n/a' || !color) return {}
  return {
    backgroundColor: color.header,
    color: '#0F172A',
    fontWeight: '600',
  }
}

function downloadExcel() {
  if (!result.value) return
  const link = document.createElement('a')
  const blob = b64toBlob(result.value.excel, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  link.href = URL.createObjectURL(blob)
  link.download = `Tilbudssammenligning_${new Date().toISOString().slice(0, 10)}.xlsx`
  link.click()
  URL.revokeObjectURL(link.href)
}

function downloadMatrixExcel() {
  if (!result.value?.matrix_excel) return
  const link = document.createElement('a')
  const blob = b64toBlob(
    result.value.matrix_excel,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  )
  link.href = URL.createObjectURL(blob)
  link.download = `Sammenligning_per_post_${new Date().toISOString().slice(0, 10)}.xlsx`
  link.click()
  URL.revokeObjectURL(link.href)
}

function downloadChapterExcel() {
  if (!result.value?.chapters_excel) return
  const link = document.createElement('a')
  const blob = b64toBlob(
    result.value.chapters_excel,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  )
  link.href = URL.createObjectURL(blob)
  link.download = `Kapitteloppsummering_${new Date().toISOString().slice(0, 10)}.xlsx`
  link.click()
  URL.revokeObjectURL(link.href)
}

function b64toBlob(base64, type = 'application/octet-stream') {
  const bytes = atob(base64)
  const len = bytes.length
  const buffer = new Uint8Array(len)
  for (let i = 0; i < len; i += 1) {
    buffer[i] = bytes.charCodeAt(i)
  }
  return new Blob([buffer], { type })
}

function hexToRgb(hex) {
  const normalized = hex.replace('#', '')
  if (normalized.length !== 6) return null
  const r = parseInt(normalized.slice(0, 2), 16)
  const g = parseInt(normalized.slice(2, 4), 16)
  const b = parseInt(normalized.slice(4, 6), 16)
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return null
  return { r, g, b }
}

function withAlpha(hex, alpha) {
  const rgb = hexToRgb(hex)
  if (!rgb) return 'transparent'
  const safeAlpha = Math.min(Math.max(alpha, 0), 1)
  return `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${safeAlpha})`
}

function startProgress() {
  progress.value = 0
  if (progressTimer.value) {
    window.clearInterval(progressTimer.value)
  }
  progressTimer.value = window.setInterval(() => {
    if (progress.value >= 90) return
    const increment = Math.max(1, Math.round((90 - progress.value) * 0.08))
    progress.value = Math.min(progress.value + increment, 90)
  }, 300)
}

function finishProgress(success) {
  if (progressTimer.value) {
    window.clearInterval(progressTimer.value)
    progressTimer.value = null
  }
  progress.value = success ? 100 : 0
  if (success) {
    window.setTimeout(() => {
      if (!isLoading.value) {
        progress.value = 0
      }
    }, 800)
  }
}

onBeforeUnmount(() => {
  if (progressTimer.value) {
    window.clearInterval(progressTimer.value)
  }
})
</script>
