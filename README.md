# Sistema Comercial

Aplicación web de gestión comercial: compras con costo promedio ponderado, inventario valorizado por bodegas, kardex, cotizaciones, remisiones, facturas, cartera, órdenes de compra, usuarios con roles e informes en Excel.

Funciona como página estática (HTML + JavaScript) y usa **Supabase** como base de datos y sistema de usuarios. Se puede publicar gratis en **GitHub Pages**.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | Página principal |
| `styles.css` | Diseño |
| `app.js` | Toda la lógica de la aplicación |
| `config.js` | Datos de conexión a tu proyecto Supabase (lo editas tú) |
| `supabase/schema.sql` | Base de datos completa: tablas, reglas de inventario y seguridad |

## Instalación paso a paso

### 1. Crear el proyecto en Supabase
1. Entra a https://supabase.com y crea un proyecto nuevo (elige la región São Paulo para Colombia).
2. Ve a **SQL Editor › New query**, pega todo el contenido de `supabase/schema.sql` y pulsa **Run**.
3. Ve a **Authentication › Sign In / Providers › Email** y desactiva **Confirm email** (recomendado: así los usuarios que crees podrán entrar de inmediato). Si la dejas activa, cada usuario debe confirmar su correo antes de ingresar.
4. Ve a **Authentication › URL Configuration** y en **Site URL** pon la dirección donde publicarás la app (por ejemplo `https://tuusuario.github.io/sistema-comercial/`). Es necesario para los enlaces de recuperar contraseña.

### 2. Configurar la conexión
1. En Supabase ve a **Project Settings › API**.
2. Copia la **Project URL** y la clave **anon public**.
3. Pégalas en `config.js`.

> Nunca pongas la clave `service_role` en `config.js`. La clave `anon` es pública por diseño; la seguridad la hacen las reglas RLS del archivo SQL.

### 3. Publicar en GitHub Pages
1. Crea un repositorio en GitHub (por ejemplo `sistema-comercial`).
2. Sube todos los archivos manteniendo la estructura.
3. En el repositorio: **Settings › Pages › Branch: main / root › Save**.
4. En uno o dos minutos la app estará en `https://tuusuario.github.io/sistema-comercial/`.

Para probar en tu computador sin publicar: abre la carpeta con la extensión *Live Server* de VS Code, o ejecuta `python -m http.server` dentro de la carpeta y entra a `http://localhost:8000`.

### 4. Primer ingreso
La primera vez la pantalla muestra **Configuración inicial**: el primer usuario que se registra queda como **administrador**. Después de eso, nadie más puede auto-registrarse con acceso: las cuentas nuevas quedan inactivas hasta que un administrador las active.

Luego, como administrador:
1. **Empresa y consecutivos**: razón social, NIT, resolución y prefijos/números de documentos.
2. **Bodegas**: crea tus bodegas (ya existe "Bodega principal").
3. **Productos**: créalos o impórtalos desde Excel (hay una plantilla descargable).
4. **Ajustes y traslados**: carga el **inventario inicial** con un ajuste tipo *Entrada*, indicando cantidad y costo unitario de cada producto.
5. **Usuarios y roles**: crea las cuentas de tus trabajadores.

## Cómo funciona

### Roles
- **Administrador**: todo, incluyendo productos, bodegas, ajustes y traslados de inventario, anular documentos, eliminar abonos, usuarios y configuración.
- **Trabajador**: cotizar, remisionar, facturar, registrar compras y órdenes de compra, recibir mercancía, gestionar clientes y proveedores, registrar abonos, consultar existencias, kardex e informes.

Los permisos se validan en la base de datos, no solo en la pantalla.

### Flujo de ventas
Cotización → (aprobar) → Remisión o Factura. Una remisión descarga el inventario; al facturarla no se vuelve a descargar. Las facturas de contado quedan pagadas; las de crédito pasan a **Cartera** con fecha de vencimiento según los días de crédito del cliente. Si el cliente tiene cupo de crédito, el sistema no deja facturar por encima del cupo.

### Flujo de compras
Orden de compra → **Recibir mercancía** (total o parcial) → genera la compra y la entrada a la bodega. La orden queda *Parcial* hasta recibir todo, o se puede *Cerrar* si el resto no llegará.

### Costo promedio ponderado
Se calcula por producto **y por bodega**:
- Entrada: `nuevo promedio = (saldo × promedio + cantidad × costo) / (saldo + cantidad)`
- Salida: sale al promedio vigente.
- El costo de compra que entra al inventario es el valor neto (con descuento, sin IVA, porque el IVA es descontable).
- Anular una compra retira las unidades a su costo original y recalcula el promedio.

Cada movimiento queda en el **Kardex** con su saldo, costo promedio y valor después del movimiento. El sistema no permite existencias negativas.

## Importante sobre facturación electrónica
Las facturas de esta aplicación son documentos comerciales internos. En Colombia, la factura electrónica de venta válida ante la DIAN requiere un proveedor tecnológico habilitado (o la solución gratuita de la DIAN), con XML firmado, CUFE y validación previa. Si estás obligado a facturar electrónicamente, esta app sirve para controlar ventas, inventario y cartera, y el siguiente paso sería integrarla con un proveedor tecnológico.

## Próximos ajustes sugeridos
- Integración con proveedor de facturación electrónica DIAN.
- Notas crédito/débito y devoluciones parciales.
- Listas de precios por cliente.
- Ocultar costos y utilidades a los trabajadores.
- Retenciones (retefuente, reteIVA, reteICA) en compras y ventas.
- Cuentas por pagar a proveedores.
