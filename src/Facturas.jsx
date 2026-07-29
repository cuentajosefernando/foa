import { useEffect, useMemo, useState } from 'react'
import ExcelJS from 'exceljs'
import { auth } from './firebase'
import { colors } from './theme'
import { ui, Kpi, ThOrdenable, Skeleton, comparar } from './ui'

const ESTADOS = { draft: 'Borrador', posted: 'Publicada', cancel: 'Cancelada' }

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

const fechaCorta = (f) =>
  f
    ? new Date(f + 'T00:00:00').toLocaleDateString('es-MX', {
        day: '2-digit',
        month: 'short',
        year: 'numeric',
      })
    : '—'

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

  async function cargar() {
    setCargando(true)
    setError('')
    try {
      const token = await auth.currentUser.getIdToken()
      const res = await fetch('/api/facturas', { headers: { Authorization: `Bearer ${token}` } })
      const datos = await res.json()
      if (!res.ok) throw new Error(datos.error || 'No se pudo cargar el reporte.')
      setFacturas(datos.facturas)
    } catch (err) {
      setError(err.message)
    } finally {
      setCargando(false)
    }
  }

  useEffect(() => {
    cargar()
  }, [])

  const filtradas = useMemo(() => {
    const q = busqueda.trim().toLowerCase()
    const sel = facturas.filter((f) => {
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
    return [...sel].sort((a, b) => comparar(a, b, orden.campo, orden.dir))
  }, [facturas, busqueda, estado, desde, hasta, orden])

  const kpis = useMemo(() => {
    const total = filtradas.reduce((s, f) => s + (f.total || 0), 0)
    const saldo = filtradas.reduce((s, f) => s + (f.saldo || 0), 0)
    const divisas = new Set(filtradas.map((f) => f.moneda).filter(Boolean))
    return { total, saldo, cuenta: filtradas.length, mezcla: divisas.size > 1 }
  }, [filtradas])

  function ordenarPor(campo) {
    setOrden((prev) =>
      prev.campo === campo ? { campo, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { campo, dir: 'asc' }
    )
  }

  const hayFiltros = busqueda || estado !== 'todos' || desde || hasta

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
      const enc = ws.getRow(1)
      enc.font = { bold: true, color: { argb: 'FFFFFFFF' } }
      enc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3A3A3A' } }
      enc.height = 20

      filtradas.forEach((f) => {
        ws.addRow({
          ...f,
          fecha: f.fecha ? new Date(f.fecha + 'T00:00:00') : null,
          vencimiento: f.vencimiento ? new Date(f.vencimiento + 'T00:00:00') : null,
          estado: ESTADOS[f.estado] || f.estado,
          estadoPago: ESTADOS_PAGO[f.estadoPago] || f.estadoPago,
        })
      })

      ;['B', 'C'].forEach((c) => (ws.getColumn(c).numFmt = 'dd/mm/yyyy'))
      ;['G', 'H', 'I', 'J'].forEach((c) => (ws.getColumn(c).numFmt = '#,##0.00'))
      ws.views = [{ state: 'frozen', ySplit: 1 }]
      ws.autoFilter = { from: 'A1', to: 'L1' }

      const buffer = await wb.xlsx.writeBuffer()
      const url = URL.createObjectURL(
        new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
      )
      const a = document.createElement('a')
      a.href = url
      a.download = `facturas-pres-${new Date().toISOString().slice(0, 10)}.xlsx`
      a.click()
      URL.revokeObjectURL(url)
    } catch {
      setError('No se pudo generar el archivo de Excel.')
    } finally {
      setExportando(false)
    }
  }

  return (
    <section>
      <div style={ui.barraAcciones}>
        <div style={ui.nota}>Datos en vivo desde Odoo</div>
        <div style={ui.grupoBotones}>
          <button
            onClick={cargar}
            disabled={cargando}
            style={{ ...ui.botonSecundario, ...(cargando ? ui.deshabilitado : null) }}
          >
            {cargando ? 'Cargando...' : 'Actualizar datos'}
          </button>
          <button
            onClick={exportar}
            disabled={exportando || !filtradas.length}
            style={{
              ...ui.botonPrimario,
              ...(exportando || !filtradas.length ? ui.deshabilitado : null),
            }}
          >
            {exportando ? 'Generando...' : 'Exportar a Excel'}
          </button>
        </div>
      </div>

      {error && <p style={ui.error}>{error}</p>}

      <div style={ui.kpis}>
        <Kpi etiqueta="Facturas" valor={kpis.cuenta} />
        <Kpi
          etiqueta="Total facturado"
          valor={money(kpis.total)}
          nota={kpis.mezcla ? 'Incluye distintas divisas' : undefined}
        />
        <Kpi
          etiqueta="Saldo pendiente"
          valor={money(kpis.saldo)}
          nota={kpis.mezcla ? 'Incluye distintas divisas' : undefined}
        />
      </div>

      <div style={ui.panelFiltros}>
        <input
          type="search"
          placeholder="Buscar por cliente, número u origen"
          value={busqueda}
          onChange={(e) => setBusqueda(e.target.value)}
          style={ui.buscador}
        />
        <select value={estado} onChange={(e) => setEstado(e.target.value)} style={ui.input}>
          <option value="todos">Todos los estados</option>
          <option value="posted">Publicadas</option>
          <option value="draft">Borrador</option>
          <option value="cancel">Canceladas</option>
        </select>
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
              setEstado('todos')
              setDesde('')
              setHasta('')
            }}
            style={ui.botonSecundario}
          >
            Limpiar filtros
          </button>
        )}
      </div>

      {cargando ? (
        <Skeleton />
      ) : !filtradas.length ? (
        <p style={ui.vacio}>No hay facturas para estos filtros.</p>
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
              {filtradas.map((f, i) => (
                <tr key={f.id} style={i % 2 ? ui.filaPar : null}>
                  <td style={{ ...ui.td, fontWeight: 600 }}>{f.numero}</td>
                  <td style={ui.td}>{fechaCorta(f.fecha)}</td>
                  <td style={ui.td}>{fechaCorta(f.vencimiento)}</td>
                  <td style={ui.td}>{f.cliente}</td>
                  <td style={ui.tdNum}>{money(f.subtotal, f.moneda)}</td>
                  <td style={ui.tdNum}>{money(f.impuestos, f.moneda)}</td>
                  <td style={{ ...ui.tdNum, fontWeight: 700 }}>{money(f.total, f.moneda)}</td>
                  <td style={ui.tdNum}>{money(f.saldo, f.moneda)}</td>
                  <td style={ui.td}>
                    <Badge estado={f.estado} />
                  </td>
                  <td style={ui.td}>{ESTADOS_PAGO[f.estadoPago] || f.estadoPago}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  )
}

function Badge({ estado }) {
  const tono =
    estado === 'posted' ? colors.exito : estado === 'cancel' ? colors.alerta : colors.grafito
  return <span style={ui.badge(tono)}>{ESTADOS[estado] || estado}</span>
}

export default Facturas
