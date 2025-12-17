# Ejecutor de Consultas SQL para Oracle

Script automatizado en PowerShell para ejecutar múltiples consultas SQL en Oracle utilizando **Oracle SQLcl** y exportar los resultados a CSV o JSON.

## 📋 Requisitos Previos

### 1. Sistema Operativo
- Windows 10 o superior con PowerShell 5.1 o superior

### 2. PowerShell - Verificación de Versión

Para verificar su versión de PowerShell:
1. Abra PowerShell (Inicio > escriba "PowerShell")
2. Ejecute: `$PSVersionTable.PSVersion`
3. Debe mostrar versión 5.1 o superior

### 3. Oracle SQLcl

**Oracle SQLcl** es la herramienta de línea de comandos moderna de Oracle que reemplaza a SQL*Plus. Es gratuita, ligera y no requiere instalación completa de Oracle Client.

#### ¿Qué es SQLcl?

SQLcl (SQL Command Line) es:
- ✅ Gratuito y de libre uso
- ✅ Multiplataforma (Windows, Linux, Mac)
- ✅ Moderno y con más funcionalidades que SQL*Plus
- ✅ No requiere instalación de Oracle Client completo
- ✅ Incluye soporte para JSON, CSV y otros formatos
- ✅ Solo requiere Java (incluido en la descarga)

#### Descarga e Instalación de SQLcl

##### **Paso 1: Descargar SQLcl**

1. Visite: [https://www.oracle.com/database/sqldeveloper/technologies/sqlcl/download/](https://www.oracle.com/database/sqldeveloper/technologies/sqlcl/download/)
2. Descargue la versión más reciente (archivo `.zip`)
3. **No requiere cuenta de Oracle** para la descarga básica

**Tamaño aproximado:** 40-50 MB

##### **Paso 2: Instalar SQLcl**

1. **Extraer el archivo ZIP:**
   - Click derecho en el archivo descargado > "Extraer todo..."
   - Extraiga a una ubicación permanente, por ejemplo:
     - `C:\oracle\sqlcl\`
     - `C:\Program Files\Oracle\sqlcl\`
     - `%USERPROFILE%\sqlcl\`

2. **Estructura después de extraer:**
   ```
   C:\oracle\sqlcl\
   ├── bin\
   │   ├── sql.exe          ← Ejecutable principal
   │   └── sql.bat
   ├── lib\
   └── LICENSE.txt
   ```

3. **Verificar Java (Incluido):**
   - SQLcl incluye su propio Java Runtime
   - No necesita instalar Java por separado

##### **Paso 3: Verificar la Instalación**

Abra PowerShell o CMD y ejecute:

**Opción A - Si está en PATH:**
```powershell
sql -V
```

**Opción B - Ruta completa:**
```powershell
C:\oracle\sqlcl\bin\sql.exe -V
```

Debería mostrar algo como:
```
SQLcl: Release 23.4 Production
Build: 23.4.0.341.0944
```

##### **Paso 4 (Opcional): Agregar SQLcl al PATH**

Para ejecutar `sql` desde cualquier ubicación:

1. **Abra "Variables de entorno":**
   - Presione `Win + R`
   - Escriba: `sysdm.cpl`
   - Click en "Variables de entorno..."

2. **Editar PATH:**
   - En "Variables del sistema", busque `Path`
   - Click en "Editar..."
   - Click en "Nuevo"
   - Agregue: `C:\oracle\sqlcl\bin` (ajuste según su ruta)
   - Click "Aceptar" en todas las ventanas

3. **Verificar:**
   - Abra una **nueva** ventana de PowerShell
   - Ejecute: `sql -V`

#### Rutas Buscadas por el Script

El script buscará automáticamente SQLcl en:
- `C:\oracle\sqlcl\bin\sql.exe`
- `C:\Program Files\Oracle\sqlcl\bin\sql.exe`
- `%USERPROFILE%\sqlcl\bin\sql.exe`
- `%ORACLE_HOME%\sqlcl\bin\sql.exe`
- Variable `PATH` del sistema

## 🔐 Configuración de Permisos de PowerShell

### ¿Por qué es necesario?

Por defecto, Windows **bloquea la ejecución de scripts de PowerShell** por seguridad. Debe otorgar permisos temporales para ejecutar este script.

### Opción 1: Permitir Ejecución Temporal (Recomendado)

Esta opción permite ejecutar el script **una sola vez** sin cambiar la configuración del sistema:

1. **Click derecho** en el archivo `ejecutar_consultas_oracle.ps1`
2. Seleccione **"Ejecutar con PowerShell"**
3. Si aparece un mensaje de seguridad, seleccione **"Abrir"** o **"Ejecutar de todas formas"**

### Opción 2: Cambiar Política de Ejecución (Para Uso Frecuente)

Si planea ejecutar el script múltiples veces:

#### **Dar Permisos:**

1. **Abra PowerShell como Administrador:**
   - Presione `Win + X`
   - Seleccione **"Windows PowerShell (Administrador)"** o **"Terminal (Administrador)"**

2. **Ejecute el siguiente comando:**
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Confirme con** `S` (Sí) cuando se le pregunte

4. **Cierre PowerShell**

**¿Qué hace esto?**
- Permite ejecutar scripts locales que usted mismo creó
- Mantiene protección contra scripts descargados sin firmar
- Solo afecta a su usuario, no a todo el sistema

#### **Verificar Política Actual:**
```powershell
Get-ExecutionPolicy -List
```

Debería mostrar:
```
Scope          ExecutionPolicy
-----          ---------------
CurrentUser    RemoteSigned
```

#### **Remover Permisos (Restaurar Seguridad Original):**

Una vez que termine de usar el script, puede restaurar la seguridad original:

1. **Abra PowerShell como Administrador**

2. **Ejecute:**
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope CurrentUser
   ```

3. **Confirme con** `S` (Sí)

**Esto volverá a bloquear todos los scripts de PowerShell para mayor seguridad.**

### Opción 3: Ejecución Bypass (Sin Cambiar Configuración)

Ejecutar el script **sin modificar la política** del sistema:

1. **Abra PowerShell** (no requiere administrador)
2. **Navegue a la carpeta** donde está el script:
   ```powershell
   cd "C:\Ruta\A\Tu\Carpeta"
   ```
3. **Ejecute con bypass:**
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File .\ejecutar_consultas_oracle.ps1
   ```

## 🚀 Configuración del Proyecto

### Estructura de Carpetas

Antes de ejecutar el script, debe crear la siguiente estructura de carpetas en el mismo directorio donde se encuentra el archivo `.ps1`:

```
📁 [Directorio del Script]
├── 📄 ejecutar_consultas_oracle.ps1
├── 📁 consultas/
│   ├── 📄 consulta1.sql
│   ├── 📄 consulta2.sql
│   └── 📄 ...
└── 📁 resultados/
    └── (aquí se guardarán los archivos generados)
```

### Crear las Carpetas

#### Opción 1: Manualmente
1. Cree una carpeta llamada `consultas`
2. Cree una carpeta llamada `resultados`
3. Ambas deben estar en el mismo directorio que el archivo `.ps1`

#### Opción 2: Desde PowerShell
```powershell
New-Item -ItemType Directory -Name "consultas"
New-Item -ItemType Directory -Name "resultados"
```

#### Opción 3: Desde CMD
```cmd
mkdir consultas
mkdir resultados
```

### Preparar las Consultas SQL

1. Cree archivos con extensión `.sql` dentro de la carpeta `consultas`
2. Cada archivo debe contener una consulta SQL válida para Oracle
3. Los nombres de archivo pueden contener espacios
4. **NO incluya comandos SQLcl** en sus archivos (como SET, SPOOL, EXIT) - el script los agrega automáticamente

**Ejemplo de archivo SQL** (`consultas/ventas_2024.sql`):
```sql
SELECT 
    cliente_id,
    nombre_cliente,
    SUM(monto) as total_ventas
FROM ventas
WHERE fecha >= TO_DATE('2024-01-01', 'YYYY-MM-DD')
GROUP BY cliente_id, nombre_cliente
ORDER BY total_ventas DESC;
```

**IMPORTANTE:** Solo escriba la consulta SELECT (o DML). El script automáticamente agrega los comandos necesarios para formatear y exportar los resultados.

## 💻 Uso del Script

### Ejecutar el Script

#### Método 1: Click Derecho (Más Simple)
1. **Click derecho** en `ejecutar_consultas_oracle.ps1`
2. Seleccione **"Ejecutar con PowerShell"**
3. Se abrirá una ventana de PowerShell

#### Método 2: Desde PowerShell
1. **Abra PowerShell**
2. **Navegue al directorio:**
   ```powershell
   cd "C:\Ruta\Donde\Esta\El\Script"
   ```
3. **Ejecute:**
   ```powershell
   .\ejecutar_consultas_oracle.ps1
   ```

### Datos de Entrada Requeridos

El script solicitará los siguientes datos **uno por uno**:

#### 1. Usuario de Oracle
```
Ingrese el usuario de Oracle: hr_user
```
- Ingrese el nombre de usuario de su base de datos Oracle
- Presione **ENTER**

#### 2. Contraseña (Enmascarada)
```
Ingrese la contrasena: ************
```
- Ingrese la contraseña del usuario
- **La contraseña se oculta** mientras escribe (muestra asteriscos)
- Presione **ENTER**

#### 3. Host
```
Ingrese el host (ej: localhost): 192.168.1.100
```
- Ingrese la dirección IP o nombre del servidor Oracle
- Ejemplos: `localhost`, `192.168.1.100`, `oracle.empresa.com`
- Presione **ENTER**

#### 4. Puerto
```
Ingrese el puerto (ej: 1521): 1521
```
- Ingrese el puerto de conexión (por defecto Oracle usa **1521**)
- **Debe ser un número entre 1 y 65535**
- Presione **ENTER**

#### 5. SID o Service Name
```
Ingrese el SID o Service Name: ORCL
```
- Ingrese el SID o nombre del servicio de su base de datos
- Ejemplos: `ORCL`, `XE`, `PROD`, `pdborcl`
- Presione **ENTER**

#### 6. Formato de Salida
```
Formato de salida (1=CSV, 2=JSON) [Por defecto: 1]: 1
```
- Ingrese **1** para exportar a CSV (predeterminado)
- Ingrese **2** para exportar a JSON
- Si presiona ENTER sin escribir nada, se usará CSV por defecto

### Proceso de Ejecución

Una vez ingresados todos los datos:

1. El script **verifica** la existencia de las carpetas `consultas` y `resultados`
2. Si faltan carpetas, muestra un error y espera que presione ENTER
3. Busca la instalación de Oracle SQLcl en las rutas estándar
4. **Valida la conexión** a Oracle antes de procesar consultas
5. Cuenta cuántos archivos `.sql` hay en la carpeta `consultas`
6. **Procesa cada consulta** una por una:
   - Conecta a Oracle con las credenciales proporcionadas
   - Ejecuta la consulta SQL
   - Exporta los resultados con un nombre único
7. Muestra un resumen del procesamiento con colores
8. **Siempre espera** que presione ENTER antes de cerrar

### Nombres de Archivos de Salida

Los archivos de resultados se guardan con el siguiente formato:
```
[nombre_consulta]_[fecha]_[hora].[extensión]
```

**Ejemplos:**
- `ventas_2024_20241217_143055.csv`
- `clientes_activos_20241217_144512.json`

Esto permite:
- Identificar fácilmente qué consulta generó el resultado
- Mantener un historial de ejecuciones
- Evitar sobrescribir archivos anteriores
- Timestamp con segundos para mayor precisión

## ⚠️ Manejo de Errores con Try/Catch/Finally

El script utiliza el sistema **nativo de PowerShell** para manejo de errores:

### Try Block
Contiene toda la lógica principal del script

### Catch Block
Captura **cualquier error** que ocurra y muestra:
- Mensaje de error crítico
- Detalles del error específico
- Estado final del script

### Finally Block
**SIEMPRE se ejecuta**, sin importar si hubo error o no:
- Muestra mensaje de cierre
- **Espera ENTER antes de cerrar**
- Garantiza que la ventana no se cierre abruptamente

## ❌ Errores Comunes

### El script NO se cierra automáticamente en caso de error

Todos los errores mostrarán un mensaje descriptivo y **siempre** esperarán que presione **ENTER** para cerrar la ventana.

### Errores de Validación

#### Error: Campo vacío
```
[ERROR] El usuario no puede estar vacio
```
**Solución:** Ingrese un valor válido

#### Error: Puerto inválido
```
[ERROR] El puerto debe ser un numero valido
```
**Solución:** Ingrese solo números (ejemplo: 1521)

```
[ERROR] El puerto debe estar entre 1 y 65535
```
**Solución:** Ingrese un puerto en el rango válido

### Errores de Configuración

#### Error: Carpeta "consultas" no encontrada
```
[ERROR] No se encontro la carpeta 'consultas'

Por favor, cree la carpeta 'consultas' en el mismo directorio donde esta este script
y coloque sus archivos .sql dentro de ella.
```
**Solución:** Cree la carpeta `consultas` y vuelva a ejecutar el script.

#### Error: Carpeta "resultados" no encontrada
```
[ERROR] No se encontro la carpeta 'resultados'

Por favor, cree la carpeta 'resultados' en el mismo directorio donde esta este script.
Esta carpeta se utilizara para guardar los resultados de las consultas.
```
**Solución:** Cree la carpeta `resultados` y vuelva a ejecutar el script.

#### Error: Oracle SQLcl no encontrado
```
[ERROR] No se pudo encontrar Oracle SQLcl instalado

Rutas buscadas:
  - C:\oracle\sqlcl\bin\sql.exe
  - C:\Program Files\Oracle\sqlcl\bin\sql.exe
  - %USERPROFILE%\sqlcl\bin\sql.exe
  - %ORACLE_HOME%\sqlcl\bin\sql.exe
  - Variable PATH del sistema

Por favor, descargue e instale Oracle SQLcl desde:
https://www.oracle.com/database/sqldeveloper/technologies/sqlcl/download/
```
**Solución:** Descargue e instale Oracle SQLcl siguiendo las instrucciones de este README.

### Errores de Conexión

#### Error: No se puede conectar a Oracle
```
[ERROR] No se pudo establecer conexion con Oracle

Verifique los siguientes datos:
  - Usuario: admin_ventas
  - Host: db-server.empresa.com
  - Puerto: 1521
  - SID/Service: PRODDB

Posibles causas:
  - Credenciales incorrectas
  - Servidor Oracle no accesible
  - Firewall bloqueando la conexion
  - SID o Service Name incorrecto
```

**Soluciones:**
1. Verifique que las credenciales sean correctas
2. Pruebe hacer ping al servidor: `ping db-server.empresa.com`
3. Verifique que el firewall permita conexiones al puerto Oracle
4. Confirme el SID/Service Name correcto con el administrador de BD
5. Intente conectarse manualmente con SQLcl:
   ```
   sql usuario/password@host:puerto/servicio
   ```

#### Advertencia: No hay archivos SQL
```
[ADVERTENCIA] No se encontraron archivos .sql en la carpeta 'consultas'

Por favor, agregue sus consultas SQL en la carpeta 'consultas' y vuelva a ejecutar el script.
```
**Solución:** Agregue al menos un archivo `.sql` en la carpeta `consultas`.

### Errores Durante Ejecución

#### Error al ejecutar consulta específica
```
Procesando: consulta_invalida.sql
  > Salida: consulta_invalida_20241217_1430.csv
  [ERROR] Fallo la ejecucion de la consulta
  Detalles: ORA-00942: table or view does not exist
```

**Posibles causas:**
- Sintaxis SQL incorrecta
- Tabla o columna no existe
- Permisos insuficientes sobre objetos de BD
- Consulta demasiado larga (timeout)

**Solución:**
- Pruebe la consulta manualmente en SQLcl primero
- Revise los permisos del usuario en Oracle
- Simplifique consultas muy complejas

## 📊 Ejemplo Completo de Uso

### Paso a Paso

1. **Descargar e instalar Oracle SQLcl**
   - Descargue desde: https://www.oracle.com/database/sqldeveloper/technologies/sqlcl/download/
   - Extraiga a `C:\oracle\sqlcl\`

2. **Configurar permisos de PowerShell** (Opción 2):
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Crear estructura de carpetas:**
   ```
   C:\MisConsultas\
   ├── ejecutar_consultas_oracle.ps1
   ├── consultas\
   └── resultados\
   ```

4. **Agregar consultas SQL** en `C:\MisConsultas\consultas\`:
   - `ventas_mensuales.sql`
   - `top_clientes.sql`
   - `inventario_actual.sql`

5. **Ejecutar el script:**
   - Click derecho en `ejecutar_consultas_oracle.ps1`
   - "Ejecutar con PowerShell"

6. **Ingresar datos:**
   ```
   Usuario: admin_ventas
   Contrasena: ************
   Host: db-server.empresa.com
   Puerto: 1521
   SID/Service: PRODDB
   Formato: 1
   ```

7. **Ver resultados** en `C:\MisConsultas\resultados\`:
   ```
   ventas_mensuales_20241217_150033.csv
   top_clientes_20241217_150033.csv
   inventario_actual_20241217_150033.csv
   ```

8. **Opcional - Restaurar seguridad:**
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope CurrentUser
   ```

## 🔧 Solución de Problemas

### El script no se ejecuta (Error de Política)
- Siga la sección **"Configuración de Permisos de PowerShell"**
- Use el método de **Bypass** si no puede cambiar políticas

### SQLcl no se encuentra
- Verifique que extrajo SQLcl correctamente
- Asegúrese de que `sql.exe` existe en `bin\`
- Intente agregar la ruta al PATH del sistema

### Problemas de conexión a Oracle
- Verifique que el servidor Oracle esté accesible desde su red
- Pruebe la conexión manualmente con SQLcl primero:
  ```
  sql usuario/password@host:puerto/servicio
  ```
- Verifique configuraciones de firewall y TNS

### Archivos con espacios en el nombre
- PowerShell y SQLcl manejan correctamente nombres con espacios
- No es necesario renombrar archivos

### Consultas muy grandes
- SQLcl puede tardar con consultas que devuelven muchos registros
- El script mostrará el progreso en tiempo real
- Espere a que termine el procesamiento

### Formato CSV no se ve bien
- Abra el CSV con un editor de texto primero
- Asegúrese de que su Excel esté configurado para UTF-8
- Use "Importar datos" en Excel en lugar de doble click

### La ventana se cierra inmediatamente
- **Nunca debería ocurrir** gracias al bloque `finally`
- Si ocurre, ejecute desde PowerShell directamente para ver el error

## 🎨 Características del Script PowerShell

### Ventajas de Usar SQLcl

✅ **Gratuito y ligero** - No requiere Oracle Client completo  
✅ **Multiplataforma** - Funciona en Windows, Linux, Mac  
✅ **Formatos modernos** - CSV, JSON, HTML, XML  
✅ **Rápido y eficiente** - Mejor rendimiento que SQL*Plus  
✅ **Actualizado** - Soporta las últimas versiones de Oracle  
✅ **Sin instalación compleja** - Solo extraer y usar  

### Ventajas del Script PowerShell

✅ **Try/Catch/Finally nativo** - Manejo robusto de errores  
✅ **Validación de tipos** - Puerto debe ser número  
✅ **Colores en consola** - Mejor experiencia visual  
✅ **Mejor manejo de strings** - Sin problemas con espacios  
✅ **Objetos y propiedades** - Código más limpio  
✅ **Contraseña enmascarada** - Mayor seguridad  

### Colores Utilizados

- **Cyan**: Títulos y encabezados
- **Yellow**: Advertencias y configuración
- **Green**: Éxito y confirmaciones
- **Red**: Errores críticos
- **Gray**: Información secundaria
- **White**: Información principal

## 📝 Notas Adicionales

- El script utiliza codificación **UTF-8** para soportar caracteres especiales
- Cada ejecución del script es independiente (no guarda estado entre ejecuciones)
- Los archivos de resultados **nunca se sobrescriben** gracias al timestamp único con segundos
- Se recomienda **probar las consultas manualmente** en SQLcl antes de usar el script
- El script es compatible con **Oracle 11g, 12c, 18c, 19c, 21c y 23c**
- PowerShell 5.1 viene **preinstalado** en Windows 10 y 11
- SQLcl requiere Java, pero viene incluido en el paquete

## 🔒 Seguridad y Mejores Prácticas

### Recomendaciones de Seguridad

1. **Credenciales:**
   - Nunca guarde contraseñas en el script
   - La contraseña se enmascara automáticamente durante la entrada
   - Considere usar Oracle Wallet para credenciales frecuentes

2. **Permisos de PowerShell:**
   - Use `RemoteSigned` en lugar de `Unrestricted`
   - Restaure a `Restricted` cuando termine de usar el script

3. **Consultas SQL:**
   - Revise todas las consultas antes de ejecutarlas
   - Evite consultas con `DELETE` o `UPDATE` sin `WHERE`
   - Use permisos de solo lectura cuando sea posible
   - No incluya credenciales en los archivos .sql

4. **Red:**
   - Use conexiones seguras (Oracle Advanced Security)
   - Considere VPN para conexiones remotas
   - Verifique configuraciones de firewall

## 📄 Licencia

Este script es de uso libre. Oracle SQLcl está bajo licencia Oracle Technology Network License Agreement.

## 🤝 Contribuciones

Para reportar problemas o sugerir mejoras, por favor contacte al desarrollador del proyecto.

---

**Versión del Script:** 3.0 (PowerShell + Oracle SQLcl)  
**Fecha:** Diciembre 2024  
**Compatible con:** Oracle SQLcl 23.x+, Windows 10+, PowerShell 5.1+, Oracle 11g-23c