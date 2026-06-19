# Manual de Usuario — Bot Tu Repuesto Cordoba

## Como ejecutar la simulacion

Requisitos: Python 3.8 o superior, openpyxl instalado.

```bash
pip install openpyxl
python3 bot.py
```

Al iniciar, el bot muestra el prompt "Cliente:" esperando tu mensaje.
Escribi "salir" en cualquier momento para terminar la sesion.

---

## Flujo normal de una consulta

**Paso 1 — Saludo o consulta directa**

Podés empezar con un saludo (el bot te pide el repuesto) o ir directo con la consulta:

```
Cliente: hola
Bot: Hola! Bienvenido a Tu Repuesto Cordoba. Decime marca, modelo y año...

Cliente: pastillas de freno Corsa
Bot: Tenemos Pastillas de freno delanteras (Chevrolet Corsa) en stock. Precio: $12.500...
```

**Paso 2 — Elegir método de pago**

Escribi una de las tres opciones:

| Escribi       | Resultado                        |
|---------------|----------------------------------|
| efectivo      | Sin recargo                      |
| transferencia | Sin recargo                      |
| tarjeta       | Se aplica un 10% de recargo      |

**Paso 3 — Elegir modalidad de entrega**

| Escribi            | Resultado              |
|--------------------|------------------------|
| 1  o  envio        | Envio a domicilio      |
| 2  o  retiro       | Retiro en local        |

**Paso 4 — Confirmacion**

El bot muestra el resumen con numero de pedido y lo registra automaticamente en `data/datos_negocio.xlsx`.

---

## Formas de buscar un repuesto

- Por codigo exacto: `REP001`
- Por nombre: `filtro de aceite`
- Por marca y modelo: `pastillas Corsa`
- Combinando: `pastillas freno Chevrolet Corsa`

La busqueda no distingue mayusculas ni tildes.

---

## Comandos especiales

| Comando          | Efecto                                    |
|------------------|-------------------------------------------|
| asesor           | Deriva la conversacion a un asesor humano |
| hola             | Reinicia la sesion desde cualquier estado terminal |
| nueva consulta   | Idem anterior                             |
| salir            | Cierra la simulacion por consola          |

---

## Donde quedan registrados los pedidos

Cada pedido confirmado se agrega automaticamente como una nueva fila en la hoja **Pedidos** de `data/datos_negocio.xlsx`. Podes abrir ese archivo con Excel o LibreOffice para ver el historial.

---

## Limitaciones de la simulacion

- No se conecta a la API real de WhatsApp Business (simulacion por consola).
- Un solo cliente por ejecucion (telefono fijo de demo).
- Los precios y el stock son datos de prueba; en produccion se conectaria a un sistema de gestion real.
