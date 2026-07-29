import { colors, radius, font } from './theme'

// Estilos compartidos por los reportes: un cambio aquí se refleja en toda la app.
export const ui = {
  barraAcciones: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: 12,
    marginBottom: 20,
    flexWrap: 'wrap',
  },
  grupoBotones: {
    display: 'flex',
    gap: 8,
    flexWrap: 'wrap',
  },
  botonPrimario: {
    height: 40,
    padding: '0 18px',
    fontSize: 14,
    fontWeight: 600,
    fontFamily: font.body,
    color: colors.superficie,
    background: colors.naranja,
    border: 'none',
    borderRadius: radius.input,
    cursor: 'pointer',
    boxShadow: '0 1px 2px rgba(35,35,35,0.12)',
  },
  botonSecundario: {
    height: 40,
    padding: '0 18px',
    fontSize: 14,
    fontWeight: 600,
    fontFamily: font.body,
    color: colors.texto,
    background: colors.superficie,
    border: `1px solid ${colors.grafito}55`,
    borderRadius: radius.input,
    cursor: 'pointer',
  },
  deshabilitado: {
    opacity: 0.45,
    cursor: 'default',
    boxShadow: 'none',
  },

  kpis: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
    gap: 14,
    marginBottom: 20,
  },
  kpi: {
    position: 'relative',
    background: colors.superficie,
    border: `1px solid ${colors.grafito}2E`,
    borderRadius: radius.card,
    padding: '16px 18px 16px 20px',
    overflow: 'hidden',
  },
  kpiAcento: {
    position: 'absolute',
    left: 0,
    top: 0,
    bottom: 0,
    width: 4,
    background: colors.naranja,
  },
  kpiEtiqueta: {
    fontSize: 11,
    fontWeight: 600,
    letterSpacing: '0.06em',
    textTransform: 'uppercase',
    color: colors.textoSec,
    marginBottom: 6,
  },
  kpiValor: {
    fontSize: 26,
    fontWeight: 700,
    lineHeight: 1.1,
    color: colors.texto,
    fontVariantNumeric: 'tabular-nums',
  },
  kpiNota: {
    fontSize: 12,
    color: colors.textoSec,
    marginTop: 4,
  },

  panelFiltros: {
    display: 'flex',
    alignItems: 'center',
    gap: 10,
    flexWrap: 'wrap',
    background: colors.superficie,
    border: `1px solid ${colors.grafito}2E`,
    borderRadius: radius.card,
    padding: 12,
    marginBottom: 16,
  },
  campo: {
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    fontSize: 12,
    fontWeight: 600,
    letterSpacing: '0.03em',
    textTransform: 'uppercase',
    color: colors.textoSec,
  },
  input: {
    height: 38,
    padding: '0 12px',
    fontSize: 14,
    fontFamily: font.body,
    color: colors.texto,
    border: `1px solid ${colors.grafito}55`,
    borderRadius: radius.input,
    background: colors.superficie,
  },
  buscador: {
    flex: '1 1 240px',
    minWidth: 180,
    height: 38,
    padding: '0 12px',
    fontSize: 14,
    fontFamily: font.body,
    color: colors.texto,
    border: `1px solid ${colors.grafito}55`,
    borderRadius: radius.input,
    background: colors.superficie,
  },

  tarjetaTabla: {
    background: colors.superficie,
    border: `1px solid ${colors.grafito}2E`,
    borderRadius: radius.card,
    overflow: 'auto',
    maxHeight: 'calc(100vh - 330px)',
  },
  tabla: {
    width: '100%',
    borderCollapse: 'separate',
    borderSpacing: 0,
    fontSize: 13.5,
  },
  th: {
    position: 'sticky',
    top: 0,
    zIndex: 5,
    background: '#EFEEEC',
    textAlign: 'left',
    padding: '11px 14px',
    borderBottom: `1px solid ${colors.grafito}44`,
    whiteSpace: 'nowrap',
  },
  thNum: {
    position: 'sticky',
    top: 0,
    zIndex: 5,
    background: '#EFEEEC',
    textAlign: 'right',
    padding: '11px 14px',
    borderBottom: `1px solid ${colors.grafito}44`,
    whiteSpace: 'nowrap',
  },
  botonOrden: {
    display: 'flex',
    alignItems: 'center',
    gap: 5,
    width: '100%',
    padding: 0,
    background: 'transparent',
    border: 'none',
    fontFamily: font.body,
    fontSize: 11,
    fontWeight: 700,
    letterSpacing: '0.05em',
    textTransform: 'uppercase',
    cursor: 'pointer',
  },
  flecha: { fontSize: 10, opacity: 0.75 },
  td: {
    padding: '11px 14px',
    color: colors.texto,
    borderBottom: `1px solid ${colors.grafito}1F`,
    whiteSpace: 'nowrap',
  },
  tdNum: {
    padding: '11px 14px',
    color: colors.texto,
    borderBottom: `1px solid ${colors.grafito}1F`,
    textAlign: 'right',
    fontVariantNumeric: 'tabular-nums',
    whiteSpace: 'nowrap',
  },
  filaPar: { background: '#FBFBFA' },

  badge: (tono) => ({
    display: 'inline-block',
    padding: '3px 10px',
    borderRadius: radius.chip,
    fontSize: 11.5,
    fontWeight: 700,
    letterSpacing: '0.02em',
    color: tono,
    background: `${tono}1F`,
    border: `1px solid ${tono}44`,
  }),

  vacio: {
    background: colors.superficie,
    border: `1px dashed ${colors.grafito}66`,
    borderRadius: radius.card,
    padding: 40,
    textAlign: 'center',
    color: colors.textoSec,
    fontSize: 14,
  },
  error: {
    background: `${colors.alerta}12`,
    border: `1px solid ${colors.alerta}44`,
    borderRadius: radius.input,
    padding: '10px 14px',
    color: colors.alerta,
    fontSize: 13.5,
    fontWeight: 500,
    marginBottom: 16,
  },
  skeletonFila: {
    height: 38,
    margin: '8px 12px',
    borderRadius: radius.input,
    background: `${colors.grafito}1F`,
  },
  nota: {
    fontSize: 12.5,
    color: colors.textoSec,
  },
}

export function Kpi({ etiqueta, valor, nota }) {
  return (
    <div style={ui.kpi}>
      <span style={ui.kpiAcento} />
      <div style={ui.kpiEtiqueta}>{etiqueta}</div>
      <div style={ui.kpiValor}>{valor}</div>
      {nota && <div style={ui.kpiNota}>{nota}</div>}
    </div>
  )
}

export function ThOrdenable({ columna, orden, onClick }) {
  const activo = orden.campo === columna.campo
  const flecha = !activo ? '↕' : orden.dir === 'asc' ? '↑' : '↓'
  return (
    <th
      style={columna.num ? ui.thNum : ui.th}
      aria-sort={activo ? (orden.dir === 'asc' ? 'ascending' : 'descending') : 'none'}
    >
      <button
        onClick={onClick}
        style={{
          ...ui.botonOrden,
          justifyContent: columna.num ? 'flex-end' : 'flex-start',
          color: activo ? colors.naranja : colors.textoSec,
        }}
      >
        {columna.titulo}
        <span style={ui.flecha}>{flecha}</span>
      </button>
    </th>
  )
}

export function Skeleton({ filas = 6 }) {
  return (
    <div style={{ ...ui.tarjetaTabla, padding: 4 }}>
      {Array.from({ length: filas }).map((_, i) => (
        <div key={i} style={ui.skeletonFila} />
      ))}
    </div>
  )
}

// Ordena respetando tipos y dejando los vacíos siempre al final.
export function comparar(a, b, campo, dir) {
  const factor = dir === 'asc' ? 1 : -1
  const x = a[campo]
  const y = b[campo]
  const vacioX = x === null || x === undefined || x === ''
  const vacioY = y === null || y === undefined || y === ''
  if (vacioX && vacioY) return 0
  if (vacioX) return 1
  if (vacioY) return -1
  if (typeof x === 'number' && typeof y === 'number') return (x - y) * factor
  return String(x).localeCompare(String(y), 'es', { numeric: true }) * factor
}
