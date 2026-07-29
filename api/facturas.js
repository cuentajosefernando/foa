const ODOO_URL = process.env.ODOO_URL
const ODOO_DB = process.env.ODOO_DB
const ODOO_UID = Number(process.env.ODOO_UID)
const ODOO_KEY = process.env.ODOO_KEY
const FIREBASE_API_KEY = process.env.VITE_FIREBASE_API_KEY

const CAMPOS = [
  'name',
  'invoice_date',
  'invoice_date_due',
  'partner_id',
  'amount_untaxed',
  'amount_tax',
  'amount_total',
  'amount_residual',
  'state',
  'payment_state',
  'currency_id',
  'invoice_origin',
]

// Valida el token de Firebase del usuario; sin sesión no se consulta Odoo.
async function usuarioValido(idToken) {
  if (!idToken) return false
  const res = await fetch(
    `https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=${FIREBASE_API_KEY}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ idToken }),
    }
  )
  return res.ok
}

async function odoo(model, method, args, kwargs = {}) {
  const res = await fetch(`${ODOO_URL}/jsonrpc`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      jsonrpc: '2.0',
      method: 'call',
      id: 1,
      params: {
        service: 'object',
        method: 'execute_kw',
        args: [ODOO_DB, ODOO_UID, ODOO_KEY, model, method, args, kwargs],
      },
    }),
  })
  const json = await res.json()
  if (json.error) {
    throw new Error(json.error.data?.message || 'Error al consultar Odoo')
  }
  return json.result
}

export default async function handler(req, res) {
  const token = (req.headers.authorization || '').replace(/^Bearer /, '')
  if (!(await usuarioValido(token))) {
    return res.status(401).json({ error: 'Sesión no válida.' })
  }

  try {
    const dominio = [['move_type', '=', 'out_invoice']]
    const facturas = await odoo('account.move', 'search_read', [dominio], {
      fields: CAMPOS,
      order: 'invoice_date desc, id desc',
      limit: 500,
    })

    const filas = facturas.map((f) => ({
      id: f.id,
      numero: f.name || 'Borrador',
      fecha: f.invoice_date || null,
      vencimiento: f.invoice_date_due || null,
      cliente: f.partner_id ? f.partner_id[1] : '',
      moneda: f.currency_id ? f.currency_id[1] : '',
      subtotal: f.amount_untaxed,
      impuestos: f.amount_tax,
      total: f.amount_total,
      saldo: f.amount_residual,
      estado: f.state,
      estadoPago: f.payment_state,
      origen: f.invoice_origin || '',
    }))

    res.setHeader('Cache-Control', 'no-store')
    return res.status(200).json({ facturas: filas })
  } catch (err) {
    return res.status(502).json({ error: err.message })
  }
}
