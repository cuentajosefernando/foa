---
name: pres-design
description: >
  Guía de diseño web (React + Vite) para la app interna de PRES: un
  reporteador de documentos e información de oficina. Úsalo SIEMPRE que se
  cree o modifique cualquier pantalla, componente o estilo de la app. Define
  identidad de marca, tokens de color y tipografía, patrones de tablas/reportes
  y filtros, y reglas de copy en español.
  Dispara con: "pantalla", "componente", "diseño", "UI", "estilo", "reporte",
  "tabla", "filtro", "dashboard", "login", o cualquier .jsx de la app.
---

# Diseño de PRES

Actúa como el líder de diseño de producto de PRES. Es una herramienta interna
de oficina: la usa el equipo administrativo desde escritorio, con calma, para
consultar y generar documentos y reportes. La prioridad es claridad de datos,
densidad de información legible y flujos de filtrado/exportación rápidos —
no ergonomía táctil ni estados offline (eso no aplica aquí).

## Contexto

- *Superficie única:* Web (React + Vite), uso de escritorio/tablet en oficina.
- *Usuarios:* personal administrativo de PRES, ya autenticado, revisando
  documentos, reportes y datos de negocio definidos caso por caso.
- *Backend:* Firebase / Firestore.
- Antes de codear una pantalla, di en una frase: qué información/reporte
  muestra, quién la consulta y qué acción principal permite (ver, filtrar,
  exportar, generar). Diseña para esa acción.

## Identidad y tokens

Paleta tomada del logotipo de PRES: gris carbón, gris medio y acento naranja.
Sobria y corporativa — el acento naranja se reserva para lo importante.

### Color

--carbon      #3A3A3A   /* gris oscuro del logo: header, texto de marca       */
--grafito     #9B9B9B   /* gris medio del logo: elementos secundarios         */
--naranja     #E87722   /* acento del logo (la "S"): acción primaria, foco    */
--exito       #4A9B4A   /* verde: confirmaciones, estados "al día"            */
--alerta      #C0392B   /* rojo: errores, vencidos, datos críticos            */
--fondo       #F5F5F4   /* fondo cálido neutro de las pantallas de trabajo    */
--texto       #232323   /* texto principal                                    */
--texto-sec   #6B6B6B   /* texto secundario, labels, placeholders             */
--superficie  #FFFFFF   /* tarjetas, tablas, inputs sobre el fondo            */

Reglas de uso:
- Fondo de trabajo = --fondo. Tarjetas/tablas = --superficie. Texto = --texto.
- --naranja solo en la acción primaria de cada pantalla (botón principal,
  enlace activo, foco de input). No decorativo, no en bloques grandes.
- --exito y --alerta solo como indicadores de estado (badges, semáforos),
  nunca como color de fondo de página.
- Contraste mínimo AA (4.5:1) en todo texto sobre --fondo o --superficie.

### Tipografía

- *Display / marca:* una sans-serif geométrica con peso fuerte (Inter o
  Manrope en 700) para el logo/título de login y encabezados de sección.
- *Datos / cuerpo:* Inter con *cifras tabulares*
  (font-variant-numeric: tabular-nums) en toda tabla, cifra o fecha, para que
  las columnas alineen.

Escala (desktop): 24/20/16/14/12. Los encabezados de tabla y filtros van en
14, el cuerpo de tabla en 14 con tabular-nums, los totales/KPIs en 20-24.

### Espacio, radio, elevación

- Rejilla base de *8 px*. Padding de contenedor: 24 px en desktop.
- Radio: 12 px en tarjetas/tablas, 8 px en inputs y botones, 999 px en badges.
- Sombra suave: 0 1px 4px rgba(35,35,35,0.08). Solo para separar tarjetas/
  modales del fondo, no en filas de tabla.

## Patrones de componentes

- *Login:* identidad de marca a tope. Header o panel lateral con acento
  --carbon, logo centrado, formulario limpio en tarjeta --superficie. Ver
  implementación de referencia en `src/Login.jsx`.
- *Barra de filtros:* fila horizontal fija arriba de cada reporte (rango de
  fechas, sucursal/entidad, búsqueda). Un botón "Aplicar" o filtrado en vivo,
  nunca ambos ambiguos. Botón "Exportar" alineado a la derecha.
- *Tablas de reporte:* encabezado sticky, filas con hover sutil, cifras
  alineadas a la derecha con tabular-nums, paginación o scroll virtual si son
  muchas filas. Estado vacío con mensaje claro ("No hay resultados para estos
  filtros") en vez de tabla en blanco.
- *Tarjetas KPI:* para totales/resúmenes arriba de un reporte — número grande
  tabular-nums, etiqueta pequeña --texto-sec debajo, sin decoración excesiva.
- *Detalle de documento:* vista de lectura clara, jerarquía tipográfica
  simple, acción de exportar/imprimir visible sin scroll.
- *Loading:* skeletons con la forma del contenido real (filas de tabla,
  tarjetas), no un spinner centrado en pantalla en blanco.
- *Error:* mensaje concreto de qué pasó y qué hacer, nunca un stack técnico
  ni un "algo salió mal" vacío.

## Tema compartido

Centraliza paleta, tipografía y espaciado en `src/theme.js` y expórtalos como
constantes; todo componente importa de ahí — nunca hardcodees hex sueltos.
Un cambio de marca debe tocar un solo archivo.

## Copy en español

- Vocabulario del usuario administrativo: "Generar reporte", "Exportar",
  "Filtrar por sucursal" — no "submit", no jerga técnica.
- Voz activa; el botón dice exactamente qué hace ("Exportar" produce un
  archivo, no "Procesar").
- Nombres consistentes en todo el flujo (si es "sucursal" en un filtro, es
  "sucursal" en toda la app, no "tienda" o "punto de venta").
- Errores concretos y accionables: "No se pudo cargar el reporte. Intenta de
  nuevo o revisa tu conexión."
- Sentence case, sin relleno, tono directo y profesional.
