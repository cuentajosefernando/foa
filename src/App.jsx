import { useEffect, useState } from 'react'
import { onAuthStateChanged, signOut } from 'firebase/auth'
import { auth } from './firebase'
import { colors, radius, font } from './theme'
import Login from './Login'
import Facturas from './Facturas'
import Pedimentos from './Pedimentos'
import logo from './assets/logo-pres.jpg'

const REPORTES = [
  { id: 'facturas', titulo: 'Relación facturas', componente: Facturas },
  { id: 'pedimentos', titulo: 'Pedimentos', componente: Pedimentos },
]

function App() {
  const [user, setUser] = useState(undefined)
  const [activo, setActivo] = useState(REPORTES[0].id)

  useEffect(() => {
    return onAuthStateChanged(auth, setUser)
  }, [])

  if (user === undefined) {
    return <p style={styles.cargando}>Cargando...</p>
  }

  if (!user) {
    return <Login />
  }

  const reporte = REPORTES.find((r) => r.id === activo) || REPORTES[0]
  const Contenido = reporte.componente

  return (
    <div style={styles.app}>
      <nav style={styles.menu}>
        <img src={logo} alt="PRES" style={styles.logo} />
        <ul style={styles.lista}>
          {REPORTES.map((r) => (
            <li key={r.id}>
              <button
                onClick={() => setActivo(r.id)}
                style={{
                  ...styles.item,
                  ...(r.id === activo ? styles.itemActivo : null),
                }}
              >
                {r.titulo}
              </button>
            </li>
          ))}
        </ul>
      </nav>

      <div style={styles.columna}>
        <header style={styles.header}>
          <h1 style={styles.tituloHeader}>{reporte.titulo}</h1>
          <div style={styles.headerDerecha}>
            <span style={styles.usuario}>{user.email}</span>
            <button onClick={() => signOut(auth)} style={styles.salir}>
              Cerrar sesión
            </button>
          </div>
        </header>
        <main style={styles.main}>
          <Contenido />
        </main>
      </div>
    </div>
  )
}

const styles = {
  app: {
    minHeight: '100vh',
    display: 'flex',
    background: colors.fondo,
    fontFamily: font.body,
  },
  cargando: {
    fontFamily: font.body,
    padding: 24,
    color: colors.textoSec,
  },
  menu: {
    flex: '0 0 220px',
    background: colors.carbon,
    padding: '20px 12px',
    display: 'flex',
    flexDirection: 'column',
    gap: 24,
    position: 'sticky',
    top: 0,
    alignSelf: 'flex-start',
    height: '100vh',
    overflowY: 'auto',
  },
  logo: {
    width: 120,
    alignSelf: 'center',
  },
  lista: {
    listStyle: 'none',
    margin: 0,
    padding: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: 4,
  },
  item: {
    width: '100%',
    textAlign: 'left',
    padding: '10px 12px',
    fontSize: 14,
    fontFamily: font.body,
    color: colors.grafito,
    background: 'transparent',
    border: 'none',
    borderRadius: radius.input,
    cursor: 'pointer',
  },
  itemActivo: {
    background: colors.naranja,
    color: colors.superficie,
    fontWeight: 600,
  },
  columna: {
    flex: 1,
    minWidth: 0,
    display: 'flex',
    flexDirection: 'column',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: 16,
    padding: '14px 24px',
    background: colors.superficie,
    borderBottom: `1px solid ${colors.grafito}33`,
    position: 'sticky',
    top: 0,
    zIndex: 20,
  },
  tituloHeader: {
    margin: 0,
    fontFamily: font.display,
    fontSize: 20,
    color: colors.texto,
  },
  headerDerecha: {
    display: 'flex',
    alignItems: 'center',
    gap: 12,
  },
  usuario: {
    color: colors.textoSec,
    fontSize: 13,
  },
  salir: {
    height: 32,
    padding: '0 12px',
    fontSize: 13,
    fontFamily: font.body,
    color: colors.texto,
    background: 'transparent',
    border: `1px solid ${colors.grafito}66`,
    borderRadius: radius.input,
    cursor: 'pointer',
  },
  main: {
    padding: 24,
    flex: 1,
    minWidth: 0,
  },
}

export default App
