import { useEffect, useState } from 'react'
import { onAuthStateChanged, signOut } from 'firebase/auth'
import { collection, getDocs } from 'firebase/firestore'
import { auth, db } from './firebase'
import Login from './Login'

function App() {
  const [user, setUser] = useState(undefined)
  const [status, setStatus] = useState('Conectando a Firestore...')

  useEffect(() => {
    return onAuthStateChanged(auth, setUser)
  }, [])

  useEffect(() => {
    if (!user) return
    async function testConnection() {
      try {
        await getDocs(collection(db, 'connection_test'))
        setStatus('Conexión a Firestore OK')
      } catch (err) {
        setStatus(`Error de conexión: ${err.message}`)
      }
    }
    testConnection()
  }, [user])

  if (user === undefined) {
    return <p style={{ fontFamily: 'sans-serif', padding: '2rem' }}>Cargando...</p>
  }

  if (!user) {
    return <Login />
  }

  return (
    <div style={{ fontFamily: 'sans-serif', padding: '2rem' }}>
      <h1>FOASTUDIO</h1>
      <p>Sesión iniciada como {user.email}</p>
      <p>{status}</p>
      <button onClick={() => signOut(auth)}>Cerrar sesión</button>
    </div>
  )
}

export default App
