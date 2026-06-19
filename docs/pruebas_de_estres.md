# Pruebas de Estrés y Camino Infeliz

Todas las pruebas se ejecutan con `python3 bot.py` (simulacion por consola).

---

## Caso 1 — Camino feliz completo (pago en efectivo, retiro en local)

```
Cliente: hola
Bot: Hola! Bienvenido a Tu Repuesto Cordoba. Decime marca, modelo...

Cliente: filtro aceite corsa
Bot: Tenemos Filtro de aceite (Chevrolet Corsa) en stock. Precio: $2.800
     Como preferis pagar? efectivo | transferencia | tarjeta

Cliente: efectivo
Bot: Perfecto, pago con efectivo.
     Como queres recibir el repuesto? 1 - Envio | 2 - Retiro

Cliente: 2
Bot: Tu pedido quedo confirmado! N° pedido: PED2025...
```
**Resultado esperado:** pedido registrado en Pedidos.xlsx, precio sin recargo.

---

## Caso 2 — Camino feliz con tarjeta y envio

```
Cliente: pastillas freno focus
Bot: [precio $14.000]...  efectivo | transferencia | tarjeta

Cliente: tarjeta
Bot: precio base $14.000 + 10% tarjeta = $15.400
     1 - Envio | 2 - Retiro

Cliente: 1
Bot: confirmado, total $15.400, Envio a domicilio
```
**Resultado esperado:** recargo aplicado correctamente.

---

## Caso 3 — Repuesto sin stock

```
Cliente: radiador peugeot
Bot: Lo siento, Radiador (Peugeot 207) esta temporalmente sin stock.
     Podes dejarnos tu numero y te avisamos cuando llegue, o escribi
     "asesor" para otras opciones.
```
**Resultado esperado:** flujo termina en FIN_SIN_STOCK, sin avanzar al pago.

---

## Caso 4 — Repuesto no encontrado en catálogo

```
Cliente: turbina de avion
Bot: No encontre 'turbina de avion' en nuestro catalogo.
     Podes probar con otra descripcion, o escribi "asesor"...
```
**Resultado esperado:** flujo termina en FIN_SIN_STOCK sin excepcion.

---

## Caso 5 — Mensaje vacío en paso de consulta

```
Cliente: hola
Bot: [pide el repuesto]

Cliente:        (solo espacios, Enter)
Bot: No entendi tu consulta. Escribi la marca, modelo o codigo...
```
**Resultado esperado:** el bot permanece en ESPERANDO_REPUESTO y solicita el dato de nuevo.

---

## Caso 6 — Método de pago inválido (un intento)

```
Cliente: filtro aceite
Bot: [precio y opciones de pago]

Cliente: bitcoin
Bot: Opcion no reconocida: 'bitcoin'.
     Por favor elegi: efectivo, transferencia o tarjeta.
```
**Resultado esperado:** permanece en ESPERANDO_PAGO, contador reintentos_pago = 1.

---

## Caso 7 — Método de pago inválido repetido (deriva a asesor)

```
Cliente: bitcoin
Bot: Opcion no reconocida...  [intento 1]

Cliente: cripto
Bot: Opcion no reconocida...  [intento 2]

Cliente: no se
Bot: Hubo demasiados intentos sin una opcion valida.
     Te derivo con un asesor para ayudarte.
```
**Resultado esperado:** a los 3 intentos el bot cambia a FIN_DERIVADO.

---

## Caso 8 — Respuesta ambigua en paso de envío

```
Cliente: no se si envio o retiro
Bot: No entendi 'no se si envio o retiro'.
     Escribi 1 para envio a domicilio o 2 para retiro en local.
```
**Resultado esperado:** permanece en ESPERANDO_ENVIO, solicita clarificacion.

---

## Caso 9 — Solicitud de asesor en cualquier punto

```
Cliente: hola
Bot: [bienvenida]

Cliente: asesor
Bot: Te derivo con un asesor. En breve alguien se va a comunicar con vos.
```
**Resultado esperado:** flujo termina en FIN_DERIVADO inmediatamente.

---

## Caso 10 — Reinicio desde estado terminal

```
[despues de un pedido confirmado]
Cliente: nueva consulta
Bot: Hola! Bienvenido a Tu Repuesto Cordoba. Decime marca, modelo...
```
**Resultado esperado:** sesion reiniciada, estado vuelve a ESPERANDO_REPUESTO.

---

## Resumen de estados y transiciones

| Estado actual       | Entrada del usuario     | Siguiente estado      |
|---------------------|-------------------------|-----------------------|
| INICIO              | saludo generico         | ESPERANDO_REPUESTO    |
| INICIO              | consulta directa        | ESPERANDO_PAGO / FIN_SIN_STOCK |
| ESPERANDO_REPUESTO  | consulta valida + stock | ESPERANDO_PAGO        |
| ESPERANDO_REPUESTO  | sin stock               | FIN_SIN_STOCK         |
| ESPERANDO_REPUESTO  | no encontrado           | FIN_SIN_STOCK         |
| ESPERANDO_REPUESTO  | vacio                   | ESPERANDO_REPUESTO    |
| ESPERANDO_REPUESTO  | "asesor"                | FIN_DERIVADO          |
| ESPERANDO_PAGO      | efectivo / transferencia / tarjeta | ESPERANDO_ENVIO |
| ESPERANDO_PAGO      | invalido (< 3 veces)    | ESPERANDO_PAGO        |
| ESPERANDO_PAGO      | invalido (>= 3 veces)   | FIN_DERIVADO          |
| ESPERANDO_ENVIO     | 1 / envio               | PEDIDO_REGISTRADO     |
| ESPERANDO_ENVIO     | 2 / retiro              | PEDIDO_REGISTRADO     |
| ESPERANDO_ENVIO     | ambiguo (< 3 veces)     | ESPERANDO_ENVIO       |
| ESPERANDO_ENVIO     | ambiguo (>= 3 veces)    | FIN_DERIVADO          |
| PEDIDO_REGISTRADO   | "hola" / "nueva consulta" | ESPERANDO_REPUESTO  |
| FIN_SIN_STOCK       | "hola" / "nueva consulta" | ESPERANDO_REPUESTO  |
| FIN_DERIVADO        | "hola" / "nueva consulta" | ESPERANDO_REPUESTO  |
