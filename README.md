# 🍏 INTERFRUTD · Suite de Inteligencia Comercial & Pricing

🌐 **Web en Vivo (GitHub Pages):** [https://sanguti95-blip.github.io/analista-pricing-interfrutd/](https://sanguti95-blip.github.io/analista-pricing-interfrutd/)

Proyecto demostrativo de alto impacto diseñado para el puesto de **Analista Comercial y Pricing** en **INTERFRUTD — El Fruto Perfecto**.

> *"Convierte los datos en estrategias que impulsan el crecimiento del negocio."*

---

## 🎯 Módulos Incluidos (Alineados 100% a las Responsabilidades de la Vacante)

### 1. 🏷️ Motor de Pricing, Costos y Listas de Precios (`Pricing, Costos & Listas`)
- **211 SKUs reales** del catálogo 100% fresco agrícola de frutas y verduras (Frutas Nacionales, Frutas Importadas y Vegetales).
- **Gestión de 4 Listas de Precios Corporativas**:
  - Lista Mayorista / CEDI (Margen base ~18%)
  - Lista Supermercados / Cadenas (Margen base ~24%)
  - Lista HORECA / Food Service (Margen base ~28%)
  - Lista Detalle / Punto de Venta (Editable en vivo)
- **Simulador de Margen Meta en Tiempo Real** (15% a 50%) y botón de **Optimización Automática de Precios Críticos**.
- **Benchmark Competitivo en Vivo**:
  - Comparación contra **Auto Mercado, Walmart y Más x Menos** (163 productos homologados).
  - Comparativa mayorista de **39 presentaciones en caja contra PIMA CENADA** (Mínimo, Moda, Promedio y Veredicto Competitivo).

### 2. 📄 Cotizador Comercial B2B para Clientes (`Cotizador Comercial B2B`)
- Elaboración rápida de cotizaciones para clientes corporativos (Hoteles, Cadenas de Restaurantes, Supermercados, Comedores Industriales y Mayoristas).
- **Cálculo en tiempo real** de Costo Unitario, Precio de Lista, `% Descuento por Volumen`, Precio Ofertado, Utilidad Bruta (`₡`) y `% Margen Bruto`.
- **Semáforo de Aprobación Comercial**:
  - `>= 22%`: Aprobación Automática
  - `16% - 21.9%`: Requiere Validación de Jefatura
  - `< 16%`: Alerta de Margen Crítico
- **Vista Previa / Impresión PDF con membrete oficial INTERFRUTD**, exportación a Excel y **conversión en 1 clic a Orden de Compra (OC)**.

### 3. 🚚 Seguimiento de Órdenes de Compra (OC), Ventas, Inventarios y ABC (`Órdenes de Compra, Ventas & Inventarios`)
- **Pipeline de Órdenes de Compra (OC)**: Seguimiento por etapa (`OC_RECIBIDA` → `EN_ALISTAMIENTO` → `EN_RUTA` → `ENTREGADA`), con cálculo automático de **Fill Rate (Nivel de Servicio %)**.
- **Control de Inventarios Multi-Bodega**: Existencias en `CEDI Central 400`, `Bodega Comercial 401` y `Ruta 402`, velocidad de venta mensual, **Días de Cobertura** y alertas de **Punto de Reorden (<= 5 días)**.
- **Análisis de Pareto ABC (80/15/5)** y gráficos ejecutivos interactivos.

### 4. 🗄️ Query Builder Comercial Visual & Excel Avanzado (`Query Builder & Excel Avanzado`)
- Constructor interactivo de consultas (`SELECT`, `FROM`, `WHERE`, `GROUP BY`, `ORDER BY`) con **generación de sentencia SQL en tiempo real**.
- **6 Consultas Ejecutivas Preconfiguradas en 1 Clic** para responder preguntas de negocio durante la entrevista.
- **Barra de estado estilo Excel Avanzado** (`Recuento`, `Σ Ventas`, `Σ Utilidad`, `x̄ Margen Promedio`, `Σ Stock`) y exportación directa a `.xlsx` (SheetJS).

---

## 🚀 Cómo Ejecutar el Proyecto

### Opción 1: Servidor Local Independiente
```bash
cd interfrutd-pricing
npm start
```
Se abrirá automáticamente en `http://localhost:8090`.

### Opción 2: Integrado con el Servidor Principal
También se sincroniza automáticamente a `server/public/interfrutd-pricing/`, por lo que está disponible en:
- Local: `http://localhost:3000/interfrutd-pricing/`
- Nube (Render): `https://chsd-backend.onrender.com/interfrutd-pricing/`
