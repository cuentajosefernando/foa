import { useState } from 'react'
import { signInWithEmailAndPassword } from 'firebase/auth'
import { auth } from './firebase'
import { colors, radius, shadow, font } from './theme'
import logo from './assets/logo-pres.jpg'

const ERROR_MESSAGES = {
  'auth/invalid-email': 'El correo no es válido.',
  'auth/invalid-credential': 'Correo o contraseña incorrectos.',
  'auth/user-disabled': 'Esta cuenta fue deshabilitada.',
  'auth/too-many-requests': 'Demasiados intentos. Espera un momento e intenta de nuevo.',
}

function Login() {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)

  async function handleSubmit(e) {
    e.preventDefault()
    setError('')
    setLoading(true)
    try {
      await signInWithEmailAndPassword(auth, email, password)
    } catch (err) {
      setError(ERROR_MESSAGES[err.code] || 'No se pudo iniciar sesión. Intenta de nuevo.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div style={styles.page}>
      <aside style={styles.brandPanel}>
        <img src={logo} alt="PRES" style={styles.logo} />
        <p style={styles.brandTagline}>Reportes y documentos</p>
      </aside>

      <main style={styles.formPanel}>
        <form style={styles.card} onSubmit={handleSubmit}>
          <h1 style={styles.title}>Iniciar sesión</h1>

          <label style={styles.label}>
            Correo
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              required
              autoComplete="username"
              style={styles.input}
            />
          </label>

          <label style={styles.label}>
            Contraseña
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              required
              autoComplete="current-password"
              style={styles.input}
            />
          </label>

          {error && <p style={styles.error}>{error}</p>}

          <button type="submit" disabled={loading} style={styles.button}>
            {loading ? 'Ingresando...' : 'Ingresar'}
          </button>
        </form>
      </main>
    </div>
  )
}

const styles = {
  page: {
    minHeight: '100vh',
    display: 'flex',
    fontFamily: font.body,
    background: colors.fondo,
  },
  brandPanel: {
    flex: '0 0 38%',
    minWidth: 260,
    background: colors.carbon,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 12,
    padding: 32,
  },
  logo: {
    width: 200,
    maxWidth: '80%',
  },
  brandTagline: {
    color: colors.grafito,
    fontSize: 14,
    margin: 0,
  },
  formPanel: {
    flex: 1,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: 24,
  },
  card: {
    display: 'flex',
    flexDirection: 'column',
    gap: 12,
    width: '100%',
    maxWidth: 340,
    padding: 32,
    background: colors.superficie,
    borderRadius: radius.card,
    boxShadow: shadow,
  },
  title: {
    margin: '0 0 8px',
    fontFamily: font.display,
    fontWeight: 700,
    fontSize: 22,
    color: colors.texto,
  },
  label: {
    display: 'flex',
    flexDirection: 'column',
    gap: 4,
    fontSize: 13,
    fontWeight: 500,
    color: colors.textoSec,
  },
  input: {
    height: 44,
    padding: '0 12px',
    fontSize: 15,
    fontFamily: font.body,
    color: colors.texto,
    border: `1px solid ${colors.grafito}66`,
    borderRadius: radius.input,
    outlineColor: colors.naranja,
  },
  error: {
    color: colors.alerta,
    fontSize: 13,
    fontWeight: 500,
    margin: 0,
  },
  button: {
    marginTop: 8,
    height: 44,
    fontSize: 15,
    fontWeight: 700,
    fontFamily: font.body,
    color: colors.superficie,
    background: colors.naranja,
    border: 'none',
    borderRadius: radius.input,
    cursor: 'pointer',
  },
}

export default Login
