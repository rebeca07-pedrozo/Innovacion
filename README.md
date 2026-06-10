Listo Rebeca, aquí va lo concreto.
1. Hojas que debes crear: solo una.
Crea a mano únicamente RAW_CARGA, que es donde pegas el Excel. Con su encabezado en este orden exacto en la fila 1:
Vencimientos | Compañia | Correo | Correo jefe | NIT | DV | P | Declaración | 1a Cuota | 2da Cuota | 3a Cuota | Exógena
Las otras dos (OBLIGACIONES y HISTORICO_ALERTAS) no las crees: el código las genera solo la primera vez que corren. Si las creas a mano podrías descuadrar algo, así que déjalas que nazcan solas.
2. Triggers: uno lo creas tú a mano, los otros los crea el código.
Aquí hay una distinción importante. Tienes tres triggers en total, pero se instalan distinto:
El onEdit (el que normaliza solo cuando escribes en RAW_CARGA) no requiere que hagas nada: Google lo activa automáticamente porque la función se llama exactamente onEdit. Ya está andando.
Los dos diarios (actualizar estados a las 6am y enviar alertas a las 7am) los instala el código por ti: solo corre una vez, desde el menú, "🔔 Alertas DIAN → Crear triggers diarios". Eso los deja programados para siempre. No tienes que ir a configurarlos a mano.
Si prefieres verlos o crearlos manualmente, se hace en el editor en el menú de la izquierda, ícono del reloj (⏰ "Activadores") → "Añadir activador", eligiendo la función, "Basado en tiempo", "Temporizador diario" y la hora. Pero con correr la opción del menú ya queda hecho, no necesitas entrar ahí.
Una cosa que no es hoja ni trigger pero es el tercer pilar y se te puede olvidar: el Web App. Es lo que hace funcionar los botones del correo (era lo del error 400). Eso se despliega una sola vez: Implementar → Nueva implementación → Aplicación web → "Ejecutar como: yo", "Acceso: cualquier usuario de Davivienda".
Entonces tu checklist de arranque, en orden, sería: crear la hoja RAW_CARGA con su encabezado → desplegar el Web App → pegar una fila de prueba en RAW_CARGA (eso dispara la normalización sola) → correr "Crear triggers diarios" desde el menú → correr "Enviar alertas" para probar que llega el correo con los botones.
¿Quieres que te confirme que quedó bien armado, o probamos enviando la primera alerta de prueba?