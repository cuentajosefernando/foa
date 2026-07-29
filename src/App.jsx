import { useEffect, useState } from 'react'
import { onAuthStateChanged, signOut } from 'firebase/auth'
import { auth } from './firebase'
import { colors, radius, font } from './theme'
import Login from './Login'
import Facturas from './Facturas'
import logo from './assets/logo-pres.jpg'

function App() {
  const [user, setUser] = useState(undefined)

  useEffect(() => {
    return onAuthStateChanged(auth, setUser)
  }, [])

  if (user === undefined) {
    return <p style={styles.cargando}>Cargando...</p>
  }

  if (!user) {
    return <Login />
  }

  return (
    <div style={styles.app}>
      <header style={styles.header}>
        <img src={logo} alt="PRES" style={styles.logo} />
        <div style={styles.headerDerecha}>
          <span style={styles.usuario}>{user.email}</span>
          <button onClick={() => signOut(auth)} style={styles.salir}>
            Cerrar sesión
          </button>
        </div>
      </header>
      <main style={styles.main}>
        <Facturas />
      </main>
    </div>
  )
}

const styles = {
  app: {
    minHeight: '100vh',
    background: colors.fondo,
    fontFamily: font.body,
  },
  cargando: {
    fontFamily: font.body,
    padding: 24,
    color: colors.textoSec,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: 16,
    padding: '12px 24px',
    background: colors.carbon,
  },
  logo: {
    height: 32,
  },
  headerDerecha: {
    display: 'flex',
    alignItems: 'center',
    gap: 12,
  },
  usuario: {
    color: colors.grafito,
    fontSize: 13,
  },
  salir: {
    height: 32,
    padding: '0 12px',
    fontSize: 13,
    fontFamily: font.body,
    color: colors.superficie,
    background: 'transparent',
    border: `1px solid ${colors.grafito}`,
    borderRadius: radius.input,
    cursor: 'pointer',
  },
  main: {
    padding: 24,
  },
}

export default App
