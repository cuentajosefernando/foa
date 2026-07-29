import { useEffect, useMemo, useState } from 'react'
import ExcelJS from 'exceljs'
import { ui, Kpi, ThOrdenable, Skeleton, comparar } from './ui'

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

const money = (n) =>
  n === null || n === undefined || n === ''
    ? '—'
    : new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' }).format(n)

const fechaCorta = (f) =>
  f
    ? new Date(f + 'T00:00:00').toLocaleDateString('es-MX', {
        day: '2-digit',
        month: 'short',
        year: 'numeric',
      })
    : '—'

function Pedimentos() {
  const [pedimentos, setPedimentos] = useState([])
  const [generado, setGenerado] = useState('')
  const [cargando, setCargando] = useState(true)
  const [error, setError] = useState('')
  const [busqueda, setBusqueda] = useState('')
  const [desde, setDesde] = useState('')
  const [hasta, setHasta] = useState('')
  const [orden, setOrden] = useState({ campo: 'fechaDespacho', dir: 'desc' })
  const [exportando, setExportando] = useState(false)

  async function cargar() {
    setCargando(true)
    setError('')
    try {
      const res = await fetch(`/pedimentos.json?t=${Date.now()}`, { cache: 'no-store' })
      if (!res.ok) throw new Error('No se pudo cargar el reporte de pedimentos.')
      const datos = await res.json()
      setPedimentos(datos.pedimentos || [])
      setGenerado(datos.generado || '')
    } catch (err) {
      setError(err.message)
    } finally {
      setCargando(false)
    }
  }

  useEffect(() => {
    cargar()
  }, [])

  const filtrados = useMemo(() => {
    const q = busqueda.trim().toLowerCase()
    const sel = pedimentos.filter((p) => {
      const f = p.fechaDespacho || p.fechaLlegada
      if (desde && (!f || f < desde)) return false
      if (hasta && (!f || f > hasta)) return false
      if (!q) return true
      return [p.pedimento, p.proveedor, p.factura, p.referencia, p.mercancia, p.fraccion, p.agente]
        .join(' ')
        .toLowerCase()
        .includes(q)
    })
    return [...sel].sort((a, b) => comparar(a, b, orden.campo, orden.dir))
  }, [pedimentos, busqueda, desde, hasta, orden])

  // Un pedimento puede abarcar varias facturas (varias filas). Los importes de
  // aduana e impuestos son del pedimento completo, así que se suman una sola vez.
  const kpis = useMemo(() => {
    const porPedimento = new Map()
    for (const p of filtrados) {
      const clave = p.pedimento.replace(/\s+/g, '')
      if (!porPedimento.has(clave)) porPedimento.set(clave, p)
    }
    const unicos = [...porPedimento.values()]
    return {
      registros: filtrados.length,
      pedimentos: unicos.length,
      aduana: unicos.reduce((s, p) => s + (p.valorAduana || 0), 0),
      impuestos: unicos.reduce((s, p) => s + (p.totalImpuestos || 0), 0),
      proveedores: new Set(filtrados.map((p) => p.proveedor).filter(Boolean)).size,
    }
  }, [filtrados])

  function ordenarPor(campo) {
    setOrden((prev) =>
      prev.campo === campo
        ? { campo, dir: prev.dir === 'asc' ? 'desc' : 'asc' }
        : { campo, dir: 'asc' }
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

  return (
    <section>
      <div style={ui.barraAcciones}>
        <div style={ui.nota}>
          {generado
            ? `Extraído de la carpeta PRES 2026 · Datos al ${new Date(generado).toLocaleDateString(
                'es-MX',
                { day: '2-digit', month: 'long', year: 'numeric' }
              )}`
            : 'Reporte de importaciones'}
        </div>
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

      <div style={ui.kpis}>
        <Kpi
          etiqueta="Pedimentos"
          valor={kpis.pedimentos}
          nota={`${kpis.registros} registros de factura`}
        />
        <Kpi etiqueta="Valor aduana" valor={money(kpis.aduana)} nota="Sin duplicar por pedimento" />
        <Kpi
          etiqueta="Total impuestos"
          valor={money(kpis.impuestos)}
          nota="Sin duplicar por pedimento"
        />
        <Kpi etiqueta="Clientes / proveedores" valor={kpis.proveedores} />
      </div>

      <div style={ui.panelFiltros}>
        <input
          type="search"
          placeholder="Buscar por pedimento, proveedor, factura, mercancía o agente"
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

      {cargando ? (
        <Skeleton />
      ) : !filtrados.length ? (
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
                <tr key={`${p.pedimento}-${p.factura}-${i}`} style={i % 2 ? ui.filaPar : null}>
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
    </section>
  )
}

export default Pedimentos
