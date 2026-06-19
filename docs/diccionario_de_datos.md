# Diccionario de Datos

## Hoja: Repuestos (data/datos_negocio.xlsx)

| Campo  | Tipo    | Descripcion                                        | Ejemplo        |
|--------|---------|----------------------------------------------------|----------------|
| Codigo | Texto   | ID unico del repuesto (REP + 3 digitos)            | REP001         |
| Nombre | Texto   | Descripcion comercial del repuesto                 | Pastillas de freno delanteras |
| Marca  | Texto   | Marca del vehiculo al que corresponde              | Chevrolet      |
| Modelo | Texto   | Modelo del vehiculo                                | Corsa          |
| Anio   | Entero  | Año del modelo (0 = universal)                     | 2005           |
| Precio | Decimal | Precio de venta en pesos (sin recargo)             | 12500          |
| Stock  | Entero  | Unidades disponibles en deposito                   | 8              |

Regla de negocio: si Stock = 0, el bot informa sin stock y no avanza al flujo de venta.

---

## Hoja: Pedidos (data/datos_negocio.xlsx)

| Campo           | Tipo    | Descripcion                                    | Ejemplo              |
|-----------------|---------|------------------------------------------------|----------------------|
| ID_Pedido       | Texto   | Identificador unico generado automaticamente   | PED20250618143022    |
| Fecha_Hora      | Texto   | Timestamp de confirmacion del pedido           | 2025-06-18 14:30:22  |
| Telefono        | Texto   | Numero de telefono del cliente                 | +549351000000        |
| Codigo_Repuesto | Texto   | FK - Repuestos.Codigo                          | REP001               |
| Nombre_Repuesto | Texto   | Nombre del repuesto al momento del pedido      | Pastillas delanteras |
| Precio_Base     | Decimal | Precio sin recargo                             | 12500                |
| Metodo_Pago     | Texto   | efectivo / transferencia / tarjeta             | tarjeta              |
| Recargo_Pct     | Texto   | Porcentaje de recargo aplicado                 | 10%                  |
| Precio_Final    | Decimal | Precio base x (1 + recargo)                    | 13750                |
| Envio_Retiro    | Texto   | envio / retiro                                 | envio                |
| Estado          | Texto   | Estado del pedido al registrarse               | CONFIRMADO           |

---

## Variables de sesion (en memoria, modulo bot.py)

| Variable              | Tipo | Descripcion                                                |
|-----------------------|------|------------------------------------------------------------|
| estado                | str  | Estado actual de la conversacion (maquina de estados)      |
| telefono              | str  | Identificador de la sesion del cliente                     |
| repuesto_consultado   | str  | Texto original ingresado por el cliente                    |
| repuesto_encontrado   | dict | Fila del catalogo que coincidio con la busqueda (o None)   |
| metodo_pago           | str  | Metodo elegido: efectivo / transferencia / tarjeta         |
| envio_o_retiro        | str  | Modalidad elegida: envio / retiro                          |
| reintentos_pago       | int  | Contador de intentos invalidos en el paso de pago          |
| reintentos_envio      | int  | Contador de intentos invalidos en el paso de envio         |

---

## Constantes de configuracion (bot.py)

| Constante         | Valor | Descripcion                                         |
|-------------------|-------|-----------------------------------------------------|
| RECARGO_TARJETA   | 0.10  | Recargo por pago con tarjeta (10 %)                 |
| MAX_REINTENTOS    | 3     | Intentos antes de derivar automaticamente a asesor  |
