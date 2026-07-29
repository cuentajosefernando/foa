import { useEffect, useMemo, useState } from 'react'
import ExcelJS from 'exceljs'
import { auth } from './firebase'
import { colors, radius, shadow, font } from './theme'

const ESTADOS = {
  draft: 'Borrador',
  posted: 'Publicada',
  cancel: 'Cancelada',
}

const ESTADOS_PAGO = {
  not_paid: 'Sin pagar',
  in_payment: 'En proceso',
  paid: 'Pagada',
  partial: 'Parcial',
  reversed: 'Revertida',
}

const COLUMNAS = [
  { campo: 'numero', titulo: 'Número' },
  { campo: 'fecha', titulo: 'Fecha' },
  { campo: 'vencimiento', titulo: 'Vencimiento' },
  { campo: 'cliente', titulo: 'Cliente' },
  { campo: 'subtotal', titulo: 'Subtotal', num: true },
  { campo: 'impuestos', titulo: 'Impuestos', num: true },
  { campo: 'total', titulo: 'Total', num: true },
  { campo: 'saldo', titulo: 'Saldo', num: true },
  { campo: 'estado', titulo: 'Estado' },
  { campo: 'estadoPago', titulo: 'Pago' },
]

const money = (n, moneda) =>
  new Intl.NumberFormat('es-MX', { style: 'currency', currency: moneda || 'MXN' }).format(n || 0)

const fecha = (f) => (f ? new Date(f + 'T00:00:00').toLocaleDateString('es-MX') : '—')

function Facturas() {
  const [facturas, setFacturas] = useState([])
  const [cargando, setCargando] = useState(true)
  const [error, setError] = useState('')
  const [busqueda, setBusqueda] = useState('')
  const [estado, setEstado] = useState('todos')
  const [desde, setDesde] = useState('')
  const [hasta, setHasta] = useState('')
  const [orden, setOrden] = useState({ campo: 'fecha', dir: 'desc' })
  const [exportando, setExportando] = useState(false)

  useEffect(() => {
    async function cargar() {
      setCargando(true)
      setError('')
      try {
        const token = await auth.currentUser.getIdToken()
        const res = await fetch('/api/facturas', {
          headers: { Authorization: `Bearer ${token}` },
        })
        const datos = await res.json()
        if (!res.ok) throw new Error(datos.error || 'No se pudo cargar el reporte.')
        setFacturas(datos.facturas)
      } catch (err) {
        setError(err.message)
      } finally {
        setCargando(false)
      }
    }
    cargar()
  }, [])

  const filtradas = useMemo(() => {
    const q = busqueda.trim().toLowerCase()
    const seleccionadas = facturas.filter((f) => {
      if (estado !== 'todos' && f.estado !== estado) return false
      // Las facturas en borrador no tienen fecha: quedan fuera si se filtra por rango.
      if (desde && (!f.fecha || f.fecha < desde)) return false
      if (hasta && (!f.fecha || f.fecha > hasta)) return false
      if (!q) return true
      return (
        f.cliente.toLowerCase().includes(q) ||
        f.numero.toLowerCase().includes(q) ||
        f.origen.toLowerCase().includes(q)
      )
    })

    const factor = orden.dir === 'asc' ? 1 : -1
    return [...seleccionadas].sort((a, b) => {
      const x = a[orden.campo]
      const y = b[orden.campo]
      // Los valores vacíos siempre al final, sin importar la dirección.
      const vacioX = x === null || x === undefined || x === ''
      const vacioY = y === null || y === undefined || y === ''
      if (vacioX && vacioY) return 0
      if (vacioX) return 1
      if (vacioY) return -1
      if (typeof x === 'number' && typeof y === 'number') return (x - y) * factor
      return String(x).localeCompare(String(y), 'es', { numeric: true }) * factor
    })
  }, [facturas, busqueda, estado, desde, hasta, orden])

  function ordenarPor(campo) {
    setOrden((prev) =>
      prev.campo === campo
        ? { campo, dir: prev.dir === 'asc' ? 'desc' : 'asc' }
        : { campo, dir: 'asc' }
    )
  }

  function limpiarFiltros() {
    setBusqueda('')
    setEstado('todos')
    setDesde('')
    setHasta('')
  }

  const hayFiltros = busqueda || estado !== 'todos' || desde || hasta

  const kpis = useMemo(() => {
    const total = filtradas.reduce((s, f) => s + (f.total || 0), 0)
    const saldo = filtradas.reduce((s, f) => s + (f.saldo || 0), 0)
    return { total, saldo, cuenta: filtradas.length }
  }, [filtradas])

  async function exportar() {
    setExportando(true)
    try {
      const wb = new ExcelJS.Workbook()
      wb.creator = 'PRES'
      const ws = wb.addWorksheet('Facturas')

      ws.columns = [
        { header: 'Número', key: 'numero', width: 20 },
        { header: 'Fecha', key: 'fecha', width: 12 },
        { header: 'Vencimiento', key: 'vencimiento', width: 14 },
        { header: 'Cliente', key: 'cliente', width: 40 },
        { header: 'Origen', key: 'origen', width: 12 },
        { header: 'Moneda', key: 'moneda', width: 9 },
        { header: 'Subtotal', key: 'subtotal', width: 14 },
        { header: 'Impuestos', key: 'impuestos', width: 14 },
        { header: 'Total', key: 'total', width: 14 },
        { header: 'Saldo', key: 'saldo', width: 14 },
        { header: 'Estado', key: 'estado', width: 13 },
        { header: 'Estado de pago', key: 'estadoPago', width: 15 },
      ]

      ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } }
      ws.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF3A3A3A' },
      }
      ws.getRow(1).alignment = { vertical: 'middle' }
      ws.getRow(1).height = 20

      filtradas.forEach((f) => {
        ws.addRow({
          numero: f.numero,
          fecha: f.fecha ? new Date(f.fecha + 'T00:00:00') : null,
          vencimiento: f.vencimiento ? new Date(f.vencimiento + 'T00:00:00') : null,
          cliente: f.cliente,
          origen: f.origen,
          moneda: f.moneda,
          subtotal: f.subtotal,
          impuestos: f.impuestos,
          total: f.total,
          saldo: f.saldo,
          estado: ESTADOS[f.estado] || f.estado,
          estadoPago: ESTADOS_PAGO[f.estadoPago] || f.estadoPago,
        })
      })

      ;['B', 'C'].forEach((c) => (ws.getColumn(c).numFmt = 'dd/mm/yyyy'))
      ;['G', 'H', 'I', 'J'].forEach((c) => (ws.getColumn(c).numFmt = '#,##0.00'))
      ws.views = [{ state: 'frozen', ySplit: 1 }]
      ws.autoFilter = { from: 'A1', to: 'L1' }

      const buffer = await wb.xlsx.writeBuffer()
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `facturas-pres-${new Date().toISOString().slice(0, 10)}.xlsx`
      a.click()
      URL.revokeObjectURL(url)
    } catch (err) {
      setError('No se pudo generar el archivo de Excel.')
    } finally {
      setExportando(false)
    }
  }

  return (
    <section>
      <div style={styles.encabezado}>
        <button
          onClick={exportar}
          disabled={exportando || !filtradas.length}
          style={{
            ...styles.botonExportar,
            opacity: exportando || !filtradas.length ? 0.5 : 1,
            cursor: exportando || !filtradas.length ? 'default' : 'pointer',
          }}
        >
          {exportando ? 'Generando...' : 'Exportar a Excel'}
        </button>
      </div>

      <div style={styles.kpis}>
        <Kpi etiqueta="Facturas" valor={kpis.cuenta} />
        <Kpi etiqueta="Total facturado" valor={money(kpis.total)} />
        <Kpi etiqueta="Saldo pendiente" valor={money(kpis.saldo)} />
      </div>

      <div style={styles.filtros}>
        <input
          type="search"
          placeholder="Buscar por cliente, número u origen"
          value={busqueda}
          onChange={(e) => setBusqueda(e.target.value)}
          style={styles.input}
        />
        <select value={estado} onChange={(e) => setEstado(e.target.value)} style={styles.select}>
          <option value="todos">Todos los estados</option>
          <option value="posted">Publicadas</option>
          <option value="draft">Borrador</option>
          <option value="cancel">Canceladas</option>
        </select>
        <label style={styles.campoFecha}>
          Desde
          <input
            type="date"
            value={desde}
            max={hasta || undefined}
            onChange={(e) => setDesde(e.target.value)}
            style={styles.inputFecha}
          />
        </label>
        <label style={styles.campoFecha}>
          Hasta
          <input
            type="date"
            value={hasta}
            min={desde || undefined}
            onChange={(e) => setHasta(e.target.value)}
            style={styles.inputFecha}
          />
        </label>
        {hayFiltros && (
          <button onClick={limpiarFiltros} style={styles.limpiar}>
            Limpiar filtros
          </button>
        )}
      </div>

      {error && <p style={styles.error}>{error}</p>}

      {cargando ? (
        <Skeleton />
      ) : !filtradas.length ? (
        <p style={styles.vacio}>No hay facturas para estos filtros.</p>
      ) : (
        <div style={styles.contenedorTabla}>
          <table style={styles.tabla}>
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
              {filtradas.map((f) => (
                <tr key={f.id}>
                  <td style={styles.td}>{f.numero}</td>
                  <td style={styles.td}>{fecha(f.fecha)}</td>
                  <td style={styles.td}>{fecha(f.vencimiento)}</td>
                  <td style={styles.td}>{f.cliente}</td>
                  <td style={styles.tdNum}>{money(f.subtotal, f.moneda)}</td>
                  <td style={styles.tdNum}>{money(f.impuestos, f.moneda)}</td>
                  <td style={{ ...styles.tdNum, fontWeight: 600 }}>{money(f.total, f.moneda)}</td>
                  <td style={styles.tdNum}>{money(f.saldo, f.moneda)}</td>
                  <td style={styles.td}>
                    <Badge estado={f.estado} />
                  </td>
                  <td style={styles.td}>{ESTADOS_PAGO[f.estadoPago] || f.estadoPago}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  )
}

function ThOrdenable({ columna, orden, onClick }) {
  const activo = orden.campo === columna.campo
  const flecha = !activo ? '↕' : orden.dir === 'asc' ? '↑' : '↓'
  return (
    <th
      style={columna.num ? styles.thNum : styles.th}
      aria-sort={activo ? (orden.dir === 'asc' ? 'ascending' : 'descending') : 'none'}
    >
      <button
        onClick={onClick}
        style={{
          ...styles.botonOrden,
          justifyContent: columna.num ? 'flex-end' : 'flex-start',
          color: activo ? colors.naranja : colors.textoSec,
        }}
      >
        {columna.titulo}
        <span style={styles.flecha}>{flecha}</span>
      </button>
    </th>
  )
}

function Kpi({ etiqueta, valor }) {
  return (
    <div style={styles.kpi}>
      <div style={styles.kpiValor}>{valor}</div>
      <div style={styles.kpiEtiqueta}>{etiqueta}</div>
    </div>
  )
}

function Badge({ estado }) {
  const fondo =
    estado === 'posted' ? colors.exito : estado === 'cancel' ? colors.alerta : colors.grafito
  return (
    <span style={{ ...styles.badge, background: fondo }}>{ESTADOS[estado] || estado}</span>
  )
}

function Skeleton() {
  return (
    <div style={styles.contenedorTabla}>
      {Array.from({ length: 6 }).map((_, i) => (
        <div key={i} style={styles.skeletonFila} />
      ))}
    </div>
  )
}

const styles = {
  encabezado: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'flex-end',
    gap: 16,
    marginBottom: 16,
    flexWrap: 'wrap',
  },
  botonExportar: {
    height: 40,
    padding: '0 20px',
    fontSize: 14,
    fontWeight: 700,
    fontFamily: font.body,
    color: colors.superficie,
    background: colors.naranja,
    border: 'none',
    borderRadius: radius.input,
  },
  kpis: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
    gap: 16,
    marginBottom: 16,
  },
  kpi: {
    background: colors.superficie,
    borderRadius: radius.card,
    boxShadow: shadow,
    padding: 16,
  },
  kpiValor: {
    fontSize: 24,
    fontWeight: 700,
    color: colors.texto,
    fontVariantNumeric: 'tabular-nums',
  },
  kpiEtiqueta: {
    fontSize: 12,
    color: colors.textoSec,
    marginTop: 4,
  },
  filtros: {
    display: 'flex',
    gap: 8,
    marginBottom: 16,
    flexWrap: 'wrap',
  },
  input: {
    flex: '1 1 260px',
    height: 40,
    padding: '0 12px',
    fontSize: 14,
    fontFamily: font.body,
    color: colors.texto,
    border: `1px solid ${colors.grafito}66`,
    borderRadius: radius.input,
    background: colors.superficie,
  },
  select: {
    height: 40,
    padding: '0 8px',
    fontSize: 14,
    fontFamily: font.body,
    color: colors.texto,
    border: `1px solid ${colors.grafito}66`,
    borderRadius: radius.input,
    background: colors.superficie,
  },
  campoFecha: {
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    fontSize: 13,
    color: colors.textoSec,
  },
  inputFecha: {
    height: 40,
    padding: '0 10px',
    fontSize: 14,
    fontFamily: font.body,
    color: colors.texto,
    border: `1px solid ${colors.grafito}66`,
    borderRadius: radius.input,
    background: colors.superficie,
  },
  limpiar: {
    height: 40,
    padding: '0 14px',
    fontSize: 13,
    fontFamily: font.body,
    color: colors.texto,
    background: 'transparent',
    border: `1px solid ${colors.grafito}66`,
    borderRadius: radius.input,
    cursor: 'pointer',
  },
  botonOrden: {
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    width: '100%',
    padding: 0,
    background: 'transparent',
    border: 'none',
    font: 'inherit',
    fontSize: 12,
    fontWeight: 600,
    cursor: 'pointer',
  },
  flecha: {
    fontSize: 11,
    opacity: 0.8,
  },
  contenedorTabla: {
    background: colors.superficie,
    borderRadius: radius.card,
    boxShadow: shadow,
    // Scroll propio: mantiene fijo el encabezado de la tabla al recorrer las filas.
    overflow: 'auto',
    maxHeight: 'calc(100vh - 300px)',
    padding: 8,
  },
  tabla: {
    width: '100%',
    borderCollapse: 'collapse',
    fontSize: 14,
  },
  th: {
    position: 'sticky',
    top: 0,
    zIndex: 5,
    background: colors.superficie,
    textAlign: 'left',
    padding: '10px 12px',
    fontSize: 12,
    fontWeight: 600,
    color: colors.textoSec,
    borderBottom: `1px solid ${colors.grafito}44`,
    whiteSpace: 'nowrap',
  },
  thNum: {
    position: 'sticky',
    top: 0,
    zIndex: 5,
    background: colors.superficie,
    textAlign: 'right',
    padding: '10px 12px',
    fontSize: 12,
    fontWeight: 600,
    color: colors.textoSec,
    borderBottom: `1px solid ${colors.grafito}44`,
    whiteSpace: 'nowrap',
  },
  td: {
    padding: '10px 12px',
    color: colors.texto,
    borderBottom: `1px solid ${colors.grafito}22`,
    whiteSpace: 'nowrap',
  },
  tdNum: {
    padding: '10px 12px',
    color: colors.texto,
    borderBottom: `1px solid ${colors.grafito}22`,
    textAlign: 'right',
    fontVariantNumeric: 'tabular-nums',
    whiteSpace: 'nowrap',
  },
  badge: {
    display: 'inline-block',
    padding: '2px 10px',
    borderRadius: radius.chip,
    fontSize: 12,
    fontWeight: 600,
    color: colors.superficie,
  },
  vacio: {
    background: colors.superficie,
    borderRadius: radius.card,
    boxShadow: shadow,
    padding: 32,
    textAlign: 'center',
    color: colors.textoSec,
  },
  error: {
    color: colors.alerta,
    fontSize: 14,
    fontWeight: 500,
  },
  skeletonFila: {
    height: 36,
    margin: 6,
    borderRadius: radius.input,
    background: `${colors.grafito}22`,
  },
}

export default Facturas
