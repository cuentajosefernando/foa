/**
 * Extrae los pedimentos de todos los Excel de la carpeta origen y genera
 * public/pedimentos.json, que es lo que consume la app.
 *
 * Uso:  npm run sync-pedimentos
 */
import ExcelJS from 'exceljs'
import { readdirSync, statSync, writeFileSync, mkdirSync } from 'fs'
import { join, extname, basename } from 'path'

const ORIGEN = process.env.PRES_ORIGEN || 'C:\\Users\\froda\\OneDrive\\PRES 2026'
const SALIDA = 'public/pedimentos.json'

const MAPA = {
  pedimento: 'PEDIMENTO',
  proveedor: 'CLIENTE',
  factura: 'FACTURA',
  agente: 'AGENTE DE CARGA',
  referencia: 'REFERENCIA',
  guia: 'GUIA',
  valorFacturaTexto: 'VALOR FACTURA',
  valorAduana: 'VALOR ADUANA',
  igi: 'IGI',
  dta: 'DTA',
  iva: 'IVA',
  totalImpuestos: 'TOTAL IMPUESTOS IMPO.',
  mercancia: 'TIPO DE MERCANCIA',
  fraccion: 'FRACCION ARANCELARIA',
  transporte: 'TRANSPORTE',
  fechaLlegada: 'FECHA DE LLEGADA',
  fechaDespacho: 'FECHA DE DESPACHO',
  fechaEntrega: 'FECHA DE ENTREGA',
  facturaPres: 'FACT PRES',
}

const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim().toUpperCase()

function valorCelda(v) {
  if (v === null || v === undefined) return ''
  if (v instanceof Date) return v.toISOString().slice(0, 10)
  if (typeof v === 'object') {
    if (v.result !== undefined) return valorCelda(v.result)
    if (v.text !== undefined) return String(v.text)
    if (v.richText) return v.richText.map((t) => t.text).join('')
    return ''
  }
  return v
}

// Las fechas llegan como Date (ya ISO) o como texto dd/mm/aaaa.
function aFecha(v) {
  const s = String(v ?? '').trim()
  if (!s) return ''
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10)
  const m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/)
  if (m) {
    const [, d, mes, a] = m
    const anio = a.length === 2 ? `20${a}` : a
    const iso = `${anio}-${mes.padStart(2, '0')}-${d.padStart(2, '0')}`
    // Descarta fechas absurdas que vienen mal capturadas en el origen.
    const y = Number(anio)
    if (y < 2015 || y > 2035) return ''
    return iso
  }
  return ''
}

const aNumero = (v) => {
  if (typeof v === 'number') return v
  const limpio = String(v ?? '').replace(/[^0-9.-]/g, '')
  const n = Number(limpio)
  return Number.isFinite(n) && limpio !== '' ? n : null
}

const divisaDe = (t) => {
  const m = String(t ?? '').match(/\b(EUR|USD|MXN|MXP)\b/i)
  return m ? m[1].toUpperCase().replace('MXP', 'MXN') : ''
}

function listarExcels(dir, acc = []) {
  let entradas
  try {
    entradas = readdirSync(dir)
  } catch {
    return acc
  }
  for (const nombre of entradas) {
    const ruta = join(dir, nombre)
    let st
    try {
      st = statSync(ruta)
    } catch {
      continue
    }
    if (st.isDirectory()) listarExcels(ruta, acc)
    else if (['.xlsx', '.xlsm'].includes(extname(nombre).toLowerCase()) && !nombre.startsWith('~$')) {
      acc.push({ ruta, mtime: st.mtimeMs })
    }
  }
  return acc
}

async function extraer(ruta) {
  const wb = new ExcelJS.Workbook()
  try {
    await wb.xlsx.readFile(ruta)
  } catch {
    return []
  }

  const filas = []
  for (const ws of wb.worksheets) {
    let filaEnc = 0
    const columnas = {}
    for (let r = 1; r <= Math.min(ws.rowCount, 20); r++) {
      const fila = ws.getRow(r)
      for (let c = 1; c <= ws.columnCount; c++) {
        if (norm(valorCelda(fila.getCell(c).value)) === 'PEDIMENTO') filaEnc = r
      }
      if (filaEnc) break
    }
    if (!filaEnc) continue

    const enc = ws.getRow(filaEnc)
    for (let c = 1; c <= ws.columnCount; c++) {
      const t = norm(valorCelda(enc.getCell(c).value))
      if (t && !(t in columnas)) columnas[t] = c
    }

    // Solo son reportes de importación las hojas que además traen valor en aduana
    // y tipo de mercancía; así se descartan relaciones de pagos o de gastos.
    const esReporteImportacion =
      norm(MAPA.valorAduana) in columnas && norm(MAPA.mercancia) in columnas
    if (!esReporteImportacion) continue
    const lee = (fila, clave) => {
      const c = columnas[norm(MAPA[clave])]
      return c ? valorCelda(fila.getCell(c).value) : ''
    }

    for (let r = filaEnc + 1; r <= ws.rowCount; r++) {
      const fila = ws.getRow(r)
      const pedimento = String(lee(fila, 'pedimento') ?? '').trim()
      if (!pedimento || norm(pedimento) === 'PEDIMENTO') continue
      const textoValor = lee(fila, 'valorFacturaTexto')
      filas.push({
        pedimento,
        proveedor: String(lee(fila, 'proveedor') ?? '').trim(),
        factura: String(lee(fila, 'factura') ?? '').trim(),
        agente: String(lee(fila, 'agente') ?? '').trim(),
        referencia: String(lee(fila, 'referencia') ?? '').trim(),
        guia: String(lee(fila, 'guia') ?? '').trim(),
        valorFactura: aNumero(textoValor),
        divisa: divisaDe(textoValor),
        valorAduana: aNumero(lee(fila, 'valorAduana')),
        igi: aNumero(lee(fila, 'igi')),
        dta: aNumero(lee(fila, 'dta')),
        iva: aNumero(lee(fila, 'iva')),
        totalImpuestos: aNumero(lee(fila, 'totalImpuestos')),
        mercancia: String(lee(fila, 'mercancia') ?? '').trim(),
        fraccion: String(lee(fila, 'fraccion') ?? '').trim(),
        transporte: String(lee(fila, 'transporte') ?? '').trim(),
        fechaLlegada: aFecha(lee(fila, 'fechaLlegada')),
        fechaDespacho: aFecha(lee(fila, 'fechaDespacho')),
        fechaEntrega: aFecha(lee(fila, 'fechaEntrega')),
        facturaPres: String(lee(fila, 'facturaPres') ?? '').trim(),
        origenArchivo: basename(ruta),
      })
    }
  }
  return filas
}

// Cuenta cuántos campos relevantes trae lleno un registro.
const completitud = (f) =>
  [
    f.proveedor,
    f.factura,
    f.valorFactura,
    f.valorAduana,
    f.totalImpuestos,
    f.mercancia,
    f.fraccion,
    f.fechaDespacho,
    f.fechaLlegada,
  ].filter((v) => v !== null && v !== undefined && v !== '').length

const archivos = listarExcels(ORIGEN).sort((a, b) => a.mtime - b.mtime)
console.log(`Revisando ${archivos.length} archivos de Excel en ${ORIGEN}...`)

const porClave = new Map()
let conPedimentos = 0

for (const { ruta } of archivos) {
  const filas = await extraer(ruta)
  if (!filas.length) continue
  conPedimentos++
  console.log(`  ${filas.length.toString().padStart(4)} filas  <-  ${basename(ruta)}`)
  for (const f of filas) {
    // La misma combinación pedimento+factura puede repetirse entre archivos.
    // Se procesan de viejo a nuevo y gana el registro con más datos llenos;
    // a igualdad, gana el más reciente.
    const clave = `${norm(f.pedimento)}|${norm(f.factura)}`
    const previo = porClave.get(clave)
    if (!previo || completitud(f) >= completitud(previo)) porClave.set(clave, f)
  }
}

const datos = [...porClave.values()].sort((a, b) =>
  String(b.fechaDespacho || b.fechaLlegada || '').localeCompare(
    String(a.fechaDespacho || a.fechaLlegada || '')
  )
)

mkdirSync('public', { recursive: true })
writeFileSync(
  SALIDA,
  JSON.stringify({ generado: new Date().toISOString(), origen: ORIGEN, pedimentos: datos }, null, 0)
)

const pedimentosUnicos = new Set(datos.map((d) => norm(d.pedimento))).size
console.log(
  `\nListo: ${datos.length} registros (${pedimentosUnicos} pedimentos únicos) de ${conPedimentos} archivos -> ${SALIDA}`
)
