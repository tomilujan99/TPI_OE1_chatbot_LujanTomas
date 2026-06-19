# TPI - Organización Empresarial | UTN TUPaD

**Trabajo Práctico Integrador** — Simulación de chatbot WhatsApp Business para una tienda de repuestos automotrices.

**Materia:** Organización Empresarial  
**Carrera:** Tecnicatura Universitaria en Programación a Distancia (UTN)  
**Alumno:** Tomas Lujan

---

## Descripción

Simulación por consola de un bot de consulta y pedido de repuestos (**Tu Repuesto Córdoba**), modelado como una máquina de estados finitos. El proyecto incluye diagramas BPMN AS-IS y TO-BE, lógica de negocio en Python e integración con una base de datos simulada en Excel.

---

## Estructura del proyecto

```
├── bot.py                        # Lógica del chatbot (máquina de estados)
├── data/
│   └── datos_negocio.xlsx        # Catálogo de repuestos + registro de pedidos
├── docs/
│   ├── diccionario_de_datos.md   # Descripción de campos y variables
│   ├── manual_de_usuario.md      # Instrucciones de uso
│   └── pruebas_de_estres.md      # Casos de prueba
├── TPI_OE_TUPaD_Entregable.pdf   # Informe final (incluye diagramas BPMN)
└── TPI_OE_TUPaD_Entregable.docx  # Versión editable del informe
```

---

## Requisitos

- Python 3.8+
- openpyxl

```bash
pip install openpyxl
```

---

## Cómo ejecutar

```bash
python bot.py
```

La simulación imita una conversación de WhatsApp por consola. Escribí `salir` en cualquier momento para terminar.

### Ejemplos de búsqueda

| Lo que escribís | Resultado |
|---|---|
| `pastillas corsa` | Busca por nombre + modelo |
| `filtro de aceite` | Busca por nombre |
| `Hola, busco amortiguador Gol` | Saludo + consulta en un solo mensaje |
| `REP001` | Búsqueda exacta por código |

### Comandos especiales

| Comando | Acción |
|---|---|
| `asesor` | Derivar a un asesor humano |
| `hola` / `nueva consulta` | Reiniciar la conversación |
| `salir` | Cerrar la simulación |

---

## Flujo del bot

```
Cliente escribe → INICIO
  ├─ Saludo / consulta → busca en catálogo
  │    ├─ Múltiples resultados → ESPERANDO_SELECCION (elige por número)
  │    ├─ Un resultado con stock → ESPERANDO_CANTIDAD
  │    └─ Sin stock / no encontrado → FIN_SIN_STOCK
  ├─ Cantidad de unidades → ESPERANDO_CANTIDAD
  ├─ Método de pago → ESPERANDO_PAGO (efectivo / transferencia / tarjeta +10%)
  ├─ Envío o retiro → ESPERANDO_ENVIO
  └─ Pedido confirmado → ESPERANDO_CONTINUAR → ¿otra consulta?
```

Cada pedido confirmado se registra automáticamente en `data/datos_negocio.xlsx` (hoja *Pedidos*).

---

## Diagramas BPMN

Los diagramas AS-IS y TO-BE se encuentran en el informe `TPI_OE_TUPaD_Entregable.pdf`.