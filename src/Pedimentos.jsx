import { useEffect, useMemo, useRef, useState } from 'react'
import ExcelJS from 'exceljs'
import { colors } from './theme'
import { ui, Kpi, ThOrdenable, comparar } from './ui'

const ALMACEN = 'pres.pedimentos.v1'

const COLUMNAS = [
  { campo: 'pedimento', titulo: 'Pedimento' },
  { campo: 'fechaDespacho', titulo: 'F. despacho' },
  { campo: 'fechaLlegada', titulo: 'F. llegada' },
  { campo: 'proveedor', titulo: 'Cliente / proveedor' },
  { campo: 'factura', titulo: 'Factura' },
  { campo: 'referencia', titulo: 'Referencia' },
  { campo: 'valorFactura', titulo: 'Valor factura', num: true },
  { campo: 'valorAduana', titulo: 'Valor aduana', num: true },
  { campo: 'totalImpuestos', titulo: 'Impuestos', num: true },
  { campo: 'mercancia', titulo: 'Mercancía' },
  { campo: 'fraccion', titulo: 'Fracción' },
  { campo: 'agente', titulo: 'Agente de carga' },
]

// Nombres tal como aparecen en el Excel origen, normalizados para comparar.
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

const norm = (s) =>
  String(s ?? '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase()

const money = (n) =>
  n === null || n === undefined || n === ''
    ? '—'
    : new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' }).format(n)

const fechaCorta = (f) =>
  f ? new Date(f + 'T00:00:00').toLocaleDateString('es-MX', { day: '2-digit', month: 'short', year: 'numeric' }) : '—'

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

const aNumero = (v) => {
  if (typeof v === 'number') return v
  const limpio = String(v ?? '').replace(/[^0-9.-]/g, '')
  const n = Number(limpio)
  return Number.isFinite(n) && limpio !== '' ? n : null
}

// Extrae la divisa del texto original, p. ej. "6,200 EUR" -> EUR
const divisaDe = (texto) => {
  const m = String(texto ?? '').match(/\b(EUR|USD|MXN|MXP)\b/i)
  return m ? m[1].toUpperCase().replace('MXP', 'MXN') : ''
}

async function leerExcel(file) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(await file.arrayBuffer())
  const ws = wb.worksheets[0]
  if (!ws) throw new Error('El archivo no tiene hojas de cálculo.')

  // La fila de encabezados no siempre es la primera: se busca la que contiene PEDIMENTO.
  let filaEncabezado = 0
  const columnas = {}
  for (let r = 1; r <= Math.min(ws.rowCount, 20); r++) {
    const fila = ws.getRow(r)
    for (let c = 1; c <= ws.columnCount; c++) {
      if (norm(valorCelda(fila.getCell(c).value)) === 'PEDIMENTO') filaEncabezado = r
    }
    if (filaEncabezado) break
  }
  if (!filaEncabezado) {
    throw new Error('No se encontró la columna PEDIMENTO en el archivo. ¿Es el reporte de importaciones?')
  }

  const encabezados = ws.getRow(filaEncabezado)
  for (let c = 1; c <= ws.columnCount; c++) {
    const titulo = norm(valorCelda(encabezados.getCell(c).value))
    if (titulo) columnas[titulo] = c
  }

  const col = (clave) => columnas[norm(MAPA[clave])]
  const lee = (fila, clave) => {
    const c = col(clave)
    return c ? valorCelda(fila.getCell(c).value) : ''
  }

  const filas = []
  for (let r = filaEncabezado + 1; r <= ws.rowCount; r++) {
    const fila = ws.getRow(r)
    const pedimento = String(lee(fila, 'pedimento') ?? '').trim()
    if (!pedimento) continue

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
      fechaLlegada: String(lee(fila, 'fechaLlegada') ?? '').slice(0, 10),
      fechaDespacho: String(lee(fila, 'fechaDespacho') ?? '').slice(0, 10),
      fechaEntrega: String(lee(fila, 'fechaEntrega') ?? '').slice(0, 10),
      facturaPres: String(lee(fila, 'facturaPres') ?? '').trim(),
    })
  }

  if (!filas.length) throw new Error('No se encontraron pedimentos en el archivo.')
  return filas
}

function Pedimentos() {
  const [pedimentos, setPedimentos] = useState([])
  const [actualizado, setActualizado] = useState('')
  const [archivo, setArchivo] = useState('')
  const [cargando, setCargando] = useState(false)
  const [error, setError] = useState('')
  const [busqueda, setBusqueda] = useState('')
  const [desde, setDesde] = useState('')
  const [hasta, setHasta] = useState('')
  const [orden, setOrden] = useState({ campo: 'fechaDespacho', dir: 'desc' })
  const [exportando, setExportando] = useState(false)
  const inputArchivo = useRef(null)

  useEffect(() => {
    try {
      const guardado = localStorage.getItem(ALMACEN)
      if (guardado) {
        const { filas, fecha, nombre } = JSON.parse(guardado)
        setPedimentos(filas || [])
        setActualizado(fecha || '')
        setArchivo(nombre || '')
      }
    } catch {
      // Si el dato guardado está corrupto se ignora y se pide el archivo de nuevo.
    }
  }, [])

  async function alSeleccionar(e) {
    const file = e.target.files?.[0]
    e.target.value = '' // permite volver a elegir el mismo archivo
    if (!file) return
    setCargando(true)
    setError('')
    try {
      const filas = await leerExcel(file)
      const fecha = new Date().toISOString()
      setPedimentos(filas)
      setActualizado(fecha)
      setArchivo(file.name)
      try {
        localStorage.setItem(ALMACEN, JSON.stringify({ filas, fecha, nombre: file.name }))
      } catch {
        setError('Los datos se cargaron pero no se pudieron guardar para la próxima sesión.')
      }
    } catch (err) {
      setError(err.message || 'No se pudo leer el archivo.')
    } finally {
      setCargando(false)
    }
  }

  const filtrados = useMemo(() => {
    const q = busqueda.trim().toLowerCase()
    const sel = pedimentos.filter((p) => {
      const f = p.fechaDespacho || p.fechaLlegada
      if (desde && (!f || f < desde)) return false
      if (hasta && (!f || f > hasta)) return false
      if (!q) return true
      return [p.pedimento, p.proveedor, p.factura, p.referencia, p.mercancia, p.fraccion]
        .join(' ')
        .toLowerCase()
        .includes(q)
    })
    return [...sel].sort((a, b) => comparar(a, b, orden.campo, orden.dir))
  }, [pedimentos, busqueda, desde, hasta, orden])

  const kpis = useMemo(() => {
    const aduana = filtrados.reduce((s, p) => s + (p.valorAduana || 0), 0)
    const impuestos = filtrados.reduce((s, p) => s + (p.totalImpuestos || 0), 0)
    const proveedores = new Set(filtrados.map((p) => p.proveedor).filter(Boolean)).size
    return { aduana, impuestos, proveedores }
  }, [filtrados])

  function ordenarPor(campo) {
    setOrden((prev) =>
      prev.campo === campo ? { campo, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { campo, dir: 'asc' }
    )
  }

  const hayFiltros = busqueda || desde || hasta

  async function exportar() {
    setExportando(true)
    try {
      const wb = new ExcelJS.Workbook()
      wb.creator = 'PRES'
      const ws = wb.addWorksheet('Pedimentos')
      ws.columns = [
        { header: 'Pedimento', key: 'pedimento', width: 22 },
        { header: 'Fecha despacho', key: 'fechaDespacho', width: 15 },
        { header: 'Fecha llegada', key: 'fechaLlegada', width: 15 },
        { header: 'Fecha entrega', key: 'fechaEntrega', width: 15 },
        { header: 'Cliente / proveedor', key: 'proveedor', width: 32 },
        { header: 'Factura', key: 'factura', width: 18 },
        { header: 'Referencia', key: 'referencia', width: 15 },
        { header: 'Guía', key: 'guia', width: 18 },
        { header: 'Agente de carga', key: 'agente', width: 24 },
        { header: 'Divisa', key: 'divisa', width: 8 },
        { header: 'Valor factura', key: 'valorFactura', width: 15 },
        { header: 'Valor aduana', key: 'valorAduana', width: 15 },
        { header: 'IGI', key: 'igi', width: 12 },
        { header: 'DTA', key: 'dta', width: 12 },
        { header: 'IVA', key: 'iva', width: 12 },
        { header: 'Total impuestos', key: 'totalImpuestos', width: 16 },
        { header: 'Mercancía', key: 'mercancia', width: 40 },
        { header: 'Fracción arancelaria', key: 'fraccion', width: 22 },
        { header: 'Transporte', key: 'transporte', width: 18 },
        { header: 'Factura PRES', key: 'facturaPres', width: 14 },
      ]
      const enc = ws.getRow(1)
      enc.font = { bold: true, color: { argb: 'FFFFFFFF' } }
      enc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3A3A3A' } }
      enc.height = 20

      filtrados.forEach((p) => {
        ws.addRow({
          ...p,
          fechaDespacho: p.fechaDespacho ? new Date(p.fechaDespacho + 'T00:00:00') : null,
          fechaLlegada: p.fechaLlegada ? new Date(p.fechaLlegada + 'T00:00:00') : null,
          fechaEntrega: p.fechaEntrega ? new Date(p.fechaEntrega + 'T00:00:00') : null,
        })
      })

      ;['B', 'C', 'D'].forEach((c) => (ws.getColumn(c).numFmt = 'dd/mm/yyyy'))
      ;['K', 'L', 'M', 'N', 'O', 'P'].forEach((c) => (ws.getColumn(c).numFmt = '#,##0.00'))
      ws.views = [{ state: 'frozen', ySplit: 1 }]
      ws.autoFilter = { from: 'A1', to: 'T1' }

      const buffer = await wb.xlsx.writeBuffer()
      const url = URL.createObjectURL(
        new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
      )
      const a = document.createElement('a')
      a.href = url
      a.download = `pedimentos-pres-${new Date().toISOString().slice(0, 10)}.xlsx`
      a.click()
      URL.revokeObjectURL(url)
    } catch {
      setError('No se pudo generar el archivo de Excel.')
    } finally {
      setExportando(false)
    }
  }

  const sinDatos = !pedimentos.length

  return (
    <section>
      <input
        ref={inputArchivo}
        type="file"
        accept=".xlsx,.xlsm"
        onChange={alSeleccionar}
        style={{ display: 'none' }}
      />

      <div style={ui.barraAcciones}>
        <div style={ui.nota}>
          {actualizado
            ? `Origen: ${archivo} · Actualizado el ${new Date(actualizado).toLocaleString('es-MX')}`
            : 'Aún no has cargado el reporte de importaciones.'}
        </div>
        <div style={ui.grupoBotones}>
          <button
            onClick={() => inputArchivo.current?.click()}
            disabled={cargando}
            style={{ ...ui.botonSecundario, ...(cargando ? ui.deshabilitado : null) }}
          >
            {cargando ? 'Leyendo...' : sinDatos ? 'Cargar archivo' : 'Actualizar datos'}
          </button>
          <button
            onClick={exportar}
            disabled={exportando || !filtrados.length}
            style={{
              ...ui.botonPrimario,
              ...(exportando || !filtrados.length ? ui.deshabilitado : null),
            }}
          >
            {exportando ? 'Generando...' : 'Exportar a Excel'}
          </button>
        </div>
      </div>

      {error && <p style={ui.error}>{error}</p>}

      {sinDatos && !cargando ? (
        <div style={ui.vacio}>
          <p style={{ margin: '0 0 8px', fontWeight: 600, color: colors.texto }}>
            Carga el reporte de importaciones
          </p>
          <p style={{ margin: '0 0 16px' }}>
            Selecciona el archivo <strong>IMPORTACIONES 2025 Y 2026 PRES.xlsx</strong> de tu carpeta
            PRES 2026. Los datos quedan guardados en este navegador; usa “Actualizar datos” cuando el
            archivo cambie.
          </p>
          <button onClick={() => inputArchivo.current?.click()} style={ui.botonPrimario}>
            Cargar archivo
          </button>
        </div>
      ) : (
        <>
          <div style={ui.kpis}>
            <Kpi etiqueta="Pedimentos" valor={filtrados.length} />
            <Kpi etiqueta="Valor aduana" valor={money(kpis.aduana)} />
            <Kpi etiqueta="Total impuestos" valor={money(kpis.impuestos)} />
            <Kpi etiqueta="Clientes / proveedores" valor={kpis.proveedores} />
          </div>

          <div style={ui.panelFiltros}>
            <input
              type="search"
              placeholder="Buscar por pedimento, proveedor, factura o mercancía"
              value={busqueda}
              onChange={(e) => setBusqueda(e.target.value)}
              style={ui.buscador}
            />
            <label style={ui.campo}>
              Desde
              <input
                type="date"
                value={desde}
                max={hasta || undefined}
                onChange={(e) => setDesde(e.target.value)}
                style={ui.input}
              />
            </label>
            <label style={ui.campo}>
              Hasta
              <input
                type="date"
                value={hasta}
                min={desde || undefined}
                onChange={(e) => setHasta(e.target.value)}
                style={ui.input}
              />
            </label>
            {hayFiltros && (
              <button
                onClick={() => {
                  setBusqueda('')
                  setDesde('')
                  setHasta('')
                }}
                style={ui.botonSecundario}
              >
                Limpiar filtros
              </button>
            )}
          </div>

          {!filtrados.length ? (
            <p style={ui.vacio}>No hay pedimentos para estos filtros.</p>
          ) : (
            <div style={ui.tarjetaTabla}>
              <table style={ui.tabla}>
                <thead>
                  <tr>
                    {COLUMNAS.map((c) => (
                      <ThOrdenable
                        key={c.campo}
                        columna={c}
                        orden={orden}
                        onClick={() => ordenarPor(c.campo)}
                      />
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtrados.map((p, i) => (
                    <tr key={`${p.pedimento}-${i}`} style={i % 2 ? ui.filaPar : null}>
                      <td style={{ ...ui.td, fontWeight: 600 }}>{p.pedimento}</td>
                      <td style={ui.td}>{fechaCorta(p.fechaDespacho)}</td>
                      <td style={ui.td}>{fechaCorta(p.fechaLlegada)}</td>
                      <td style={ui.td}>{p.proveedor || '—'}</td>
                      <td style={ui.td}>{p.factura || '—'}</td>
                      <td style={ui.td}>{p.referencia || '—'}</td>
                      <td style={ui.tdNum}>
                        {p.valorFactura === null
                          ? '—'
                          : `${new Intl.NumberFormat('es-MX', {
                              minimumFractionDigits: 2,
                            }).format(p.valorFactura)}${p.divisa ? ' ' + p.divisa : ''}`}
                      </td>
                      <td style={ui.tdNum}>{money(p.valorAduana)}</td>
                      <td style={ui.tdNum}>{money(p.totalImpuestos)}</td>
                      <td
                        style={{
                          ...ui.td,
                          maxWidth: 260,
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                        }}
                        title={p.mercancia}
                      >
                        {p.mercancia || '—'}
                      </td>
                      <td style={ui.td}>{p.fraccion || '—'}</td>
                      <td style={ui.td}>{p.agente || '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
    </section>
  )
}

export default Pedimentos
