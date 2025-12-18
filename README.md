# Ejecutor de Consultas SQL para Oracle

Script automatizado en PowerShell para ejecutar múltiples consultas SQL en Oracle utilizando **Oracle SQLcl** y exportar los resultados a CSV o XLSX (Excel). Incluye validación de solo consultas SELECT, parámetros dinámicos y opciones de contraseña seguras.

## 📋 Requisitos Previos

### 1. Sistema Operativo
- Windows 10 o superior con PowerShell 5.1 o superior

### 2. PowerShell - Verificación de Versión

Para verificar su versión de PowerShell:
1. Abra PowerShell (Inicio > escriba "PowerShell")
2. Ejecute: `$PSVersionTable.PSVersion`
3. Debe mostrar versión 5.1 o superior

### 3. Java Runtime Environment (JRE) o Java Development Kit (JDK)

**Oracle SQLcl requiere Java 17 o superior** para funcionar. Debe tener instalado Java Runtime Environment (JRE) o Java Development Kit (JDK) versión **17.0.5 o superior**.

#### Verificar si Java está Instalado

Abra PowerShell o CMD y ejecute:
```powershell
java -version
```

Si Java está instalado, verá algo como:
```
java version "17.0.9" 2023-10-17 LTS
Java(TM) SE Runtime Environment (build 17.0.9+11-LTS-201)
```

Si muestra un error o una versión menor a 17, necesita instalar o actualizar Java.

#### Descargar e Instalar Java

##### **Opción 1: Oracle JDK (Recomendado para uso comercial)**

1. **Descargar Oracle JDK:**
   - Visite: https://www.oracle.com/java/technologies/downloads/
   - Seleccione **Java 17** o superior (Java 25 es la versión LTS más reciente)
   - En la sección "Windows", descargue el instalador:
     - **x64 Installer** (archivo `.exe`) para Windows 64-bit
     - Tamaño aproximado: 150-200 MB

2. **Instalar Oracle JDK:**
   - Ejecute el instalador descargado (`.exe`)
   - Siga el asistente de instalación
   - Anote la ruta de instalación (por defecto: `C:\Program Files\Java\jdk-17` o similar)
   - Complete la instalación

##### **Opción 2: OpenJDK (Gratuito y de código abierto)**

1. **Descargar OpenJDK:**
   - Visite: https://adoptium.net/ (Eclipse Temurin)
   - Seleccione:
     - **Version:** Java 17 (LTS) o superior
     - **Operating System:** Windows
     - **Architecture:** x64
   - Click en "Download JDK"
   - Tamaño aproximado: 100-150 MB

2. **Instalar OpenJDK:**
   - Ejecute el instalador descargado (`.msi`)
   - Durante la instalación, asegúrese de marcar:
     - ✅ **"Set JAVA_HOME variable"**
     - ✅ **"Add to PATH"**
   - Complete la instalación

##### **Opción 3: Microsoft Build of OpenJDK**

1. **Descargar Microsoft OpenJDK:**
   - Visite: https://learn.microsoft.com/en-us/java/openjdk/download
   - Seleccione **Java 17 LTS** o superior
   - Descargue el instalador `.msi` para Windows x64
   - Tamaño aproximado: 100-150 MB

2. **Instalar:**
   - Ejecute el instalador `.msi`
   - Siga el asistente de instalación
   - Complete la instalación

#### Configurar Variables de Entorno de Java

Si el instalador no configuró automáticamente las variables de entorno, debe hacerlo manualmente:

##### **Paso 1: Configurar JAVA_HOME**

1. **Abrir Variables de Entorno:**
   - Presione `Win + R`
   - Escriba: `sysdm.cpl`
   - Presione ENTER
   - Click en la pestaña **"Opciones avanzadas"**
   - Click en **"Variables de entorno..."**

2. **Crear/Editar JAVA_HOME:**
   - En la sección **"Variables del sistema"**, click en **"Nueva..."** (o "Editar..." si ya existe)
   - **Nombre de la variable:** `JAVA_HOME`
   - **Valor de la variable:** Ruta donde instaló Java
     - Oracle JDK: `C:\Program Files\Java\jdk-17`
     - OpenJDK (Adoptium): `C:\Program Files\Eclipse Adoptium\jdk-17.0.9.9-hotspot`
     - Microsoft OpenJDK: `C:\Program Files\Microsoft\jdk-17.0.9.9-hotspot`
   - Click **"Aceptar"**

##### **Paso 2: Agregar Java al PATH**

1. **Editar la Variable PATH:**
   - En **"Variables del sistema"**, busque la variable **`Path`**
   - Selecciónela y click en **"Editar..."**
   - Click en **"Nuevo"**
   - Agregue: `%JAVA_HOME%\bin`
   - Click **"Aceptar"** en todas las ventanas

##### **Paso 3: Verificar la Configuración**

1. **Abra una NUEVA ventana de PowerShell o CMD** (importante para que cargue las nuevas variables)
2. **Verifique JAVA_HOME:**
   ```powershell
   echo $env:JAVA_HOME
   ```
   Debería mostrar: `C:\Program Files\Java\jdk-17` (o la ruta que configuró)

3. **Verifique Java:**
   ```powershell
   java -version
   ```
   Debería mostrar la versión de Java instalada (17 o superior)

#### Solución de Problemas con Java

##### Java no se reconoce después de instalar

**Solución:**
- Cierre TODAS las ventanas de PowerShell/CMD abiertas
- Abra una NUEVA ventana de PowerShell
- Ejecute: `java -version`

##### Error: "JAVA_HOME no está definido"

**Solución:**
- Verifique que configuró JAVA_HOME correctamente
- Asegúrese de usar la ruta completa hasta la carpeta principal de Java
- No incluya `\bin` en JAVA_HOME, solo en PATH

##### Versión incorrecta de Java

Si tiene múltiples versiones de Java instaladas:
1. Verifique cuál está en PATH: `where java`
2. Asegúrese de que JAVA_HOME apunte a Java 17 o superior
3. Edite PATH para que `%JAVA_HOME%\bin` esté al INICIO de la lista

### 4. Oracle SQLcl

**Oracle SQLcl** es la herramienta de línea de comandos moderna de Oracle que reemplaza a SQL*Plus. Es gratuita, ligera y no requiere instalación completa de Oracle Client.

#### ¿Qué es SQLcl?

SQLcl (SQL Command Line) es:
- ✅ Gratuito y de libre uso
- ✅ Multiplataforma (Windows, Linux, Mac)
- ✅ Moderno y con más funcionalidades que SQL*Plus
- ✅ No requiere instalación de Oracle Client completo
- ✅ Incluye soporte para CSV y Excel (XLSX)
- ⚠️ **Requiere Java 17 o superior** (no incluido)

#### Descarga e Instalación de SQLcl

##### **Paso 1: Descargar SQLcl**

1. Visite: https://www.oracle.com/database/sqldeveloper/technologies/sqlcl/download/
2. Descargue la versión más reciente (archivo `.zip`)
3. **No requiere cuenta de Oracle** para la descarga básica

**Tamaño aproximado:** 20-30 MB

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

Si recibe un error sobre Java, asegúrese de que:
- Java 17 o superior está instalado
- JAVA_HOME está configurado correctamente
- `%JAVA_HOME%\bin` está en PATH

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

### 5. Microsoft Excel (Solo para exportar a XLSX)

Si desea exportar resultados en formato XLSX (Excel), necesita tener **Microsoft Excel instalado** en su sistema.

- El script funciona con **Excel 2007 o superior**
- **NO es necesario** si solo usa formato CSV
- El script automáticamente convierte CSV a XLSX usando Excel COM Automation

## 🚀 NUEVAS FUNCIONALIDADES

### 🔐 Sistema de Contraseña Segura por Defecto

El script ahora incluye una **contraseña por defecto preconfigurada** que se recomienda usar para mayor seguridad y evitar errores de conexión. Características:

- **Valor por defecto:** `******` (configurable en el código)

### 📄 Sistema de Parámetros para Consultas SQL

**Nueva funcionalidad:** Ahora puede pasar parámetros dinámicos a sus consultas SQL mediante archivos `.txt`:

#### Estructura de Archivos:
```
consultas/
├── mi_consulta.sql      # Consulta SQL con variables &parametro
└── mi_consulta.txt      # Archivo de parámetros (mismo nombre base)
```

#### Formato del Archivo TXT:
- **Nombres de parámetros** separados por punto y coma (`;`)
- El script solicitará interactivamente los valores de cada parámetro

**Ejemplo:**
```txt
# mi_consulta.txt
departamento;fecha_inicio;fecha_fin
```

#### Consulta SQL con Variables:
```sql
-- mi_consulta.sql
SELECT * FROM empleados 
WHERE departamento = '&departamento'
  AND fecha_contratacion BETWEEN '&fecha_inicio' AND '&fecha_fin';
```

#### Flujo de Ejecución:
1. El script detecta `mi_consulta.sql`
2. Busca automáticamente `mi_consulta.txt` en la misma carpeta
3. Lee los nombres de parámetros del archivo `.txt`
4. Solicita al usuario los valores para cada parámetro
5. Sustituye automáticamente las variables en la consulta SQL
6. Ejecuta la consulta con los valores ingresados

### 🔒 Validación Estricta de Solo SELECT

**Seguridad mejorada:** El script ahora valida automáticamente que los archivos SQL contengan **únicamente consultas SELECT**, bloqueando cualquier operación que pueda modificar datos:

#### Operaciones Bloqueadas:
- **DDL:** `CREATE`, `ALTER`, `DROP`, `TRUNCATE`, `RENAME`
- **DML:** `INSERT`, `UPDATE`, `DELETE`, `MERGE`
- **Control de Transacciones:** `COMMIT`, `ROLLBACK`, `SAVEPOINT`
- **PL/SQL:** `BEGIN`, `END`, `DECLARE`, bloques anónimos
- **Ejecución:** `EXECUTE`, `EXEC`, `CALL`
- **Otros:** `GRANT`, `REVOKE`, `AUDIT`, `FLASHBACK`

#### Ventajas:
- **Seguridad:** Previene ejecución accidental de operaciones peligrosas
- **Validación inteligente:** Ignora comentarios para evitar falsos positivos
- **Mensajes claros:** Informa exactamente qué operación no permitida se detectó
- **Compatibilidad:** Permite `WITH` (CTE) y consultas complejas válidas

### 💾 Compilación a Ejecutable (.exe)

El script puede convertirse a un archivo ejecutable autónomo:

#### Comando de Compilación:
```powershell
ps2exe -inputFile .\ejecutar_consultas_oracle.ps1 -outputFile .\ejecutar_consultas_oracle.exe -title "Ejecutor de Consultas Oracle" -version "1.0.0.0" -requireAdmin
```

#### Características del Ejecutable:
- **Parámetros preconfigurados:** Incluye contraseña por defecto
- **Sin necesidad de PowerShell:** Ejecutable nativo de Windows
- **Compatibilidad:** Funciona en cualquier sistema sin requisitos especiales
- **Seguridad:** Mantiene todas las validaciones del script original

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

### ⚠️ Solución de Problemas con Permisos

#### Error: "No se puede cargar el archivo... porque la ejecución de scripts está deshabilitada"

**Solución:**
- Use la **Opción 1** (Click derecho > Ejecutar con PowerShell)
- O use la **Opción 2** (Cambiar política de ejecución)
- O use la **Opción 3** (Bypass)

#### Error: "Acceso denegado"

**Solución:**
- Asegúrese de ejecutar PowerShell **como Administrador** al cambiar políticas
- Verifique que tiene permisos sobre la carpeta donde está el script

## 🚀 Configuración del Proyecto

### Estructura de Carpetas

Antes de ejecutar el script, debe crear la siguiente estructura de carpetas en el mismo directorio donde se encuentra el archivo `.ps1`:

```
📁 [Directorio del Script]
├── 📄 ejecutar_consultas_oracle.ps1
├── 📁 consultas/
│   ├── 📄 consulta1.sql
│   ├── 📄 consulta1.txt           ← Opcional: parámetros
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

#### Ejemplo Básico (sin parámetros):
**Archivo:** `consultas/ventas_2024.sql`
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

#### Ejemplo con Parámetros:
**Archivo SQL:** `consultas/empleados_por_departamento.sql`
```sql
SELECT 
    empleado_id,
    nombre_completo,
    fecha_contratacion,
    salario
FROM empleados
WHERE departamento = '&departamento'
  AND fecha_contratacion > '&fecha_minima';
```

**Archivo TXT (parámetros):** `consultas/empleados_por_departamento.txt`
```
departamento;fecha_minima
```

**IMPORTANTE:**
- Solo escriba la consulta SELECT (o DML)
- El script automáticamente agrega los comandos necesarios para formatear y exportar los resultados
- Use variables con formato `&nombre_parametro` en la consulta SQL
- Los nombres de parámetros en el archivo `.txt` deben coincidir exactamente con los nombres de las variables

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

#### Método 3: Ejecutable Compilado (.exe)
1. **Doble click** en `ejecutar_consultas_oracle.exe`
 
### Datos de Entrada Requeridos

El script solicitará los siguientes datos **uno por uno**:

#### 1. Usuario de Oracle
```
Ingrese el usuario de Oracle: hr_user
```
- Ingrese el nombre de usuario de su base de datos Oracle
- Presione **ENTER**

#### 2. Contraseña (Sistema Mejorado)
```
Opciones de contrasena:
  1. Usar contrasena por defecto (RECOMENDADO)
  2. Ingresar contrasena personalizada

[ADVERTENCIA] La opcion por defecto es mas segura y evita errores de conexion.

Seleccione opcion de contrasena (1 o 2) [Por defecto: 1]: 
```

**Si selecciona Opción 1:**
```
[OK] Usando contrasena por defecto
```

**Si selecciona Opción 2:**
```
Ingrese la contrasena personalizada: ************
[OK] Contrasena personalizada configurada
```

#### 3. Host
```
Ingrese el host (ej: localhost, 192.168.1.100): 192.168.1.100
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
Ingrese el SID o Service Name (ej: ORCL, XE, PDB1): ORCL
```
- Ingrese el SID o nombre del servicio de su base de datos
- Ejemplos: `ORCL`, `XE`, `PROD`, `pdborcl`
- Presione **ENTER**

#### 6. Formato de Salida
```
Formato de salida (1=CSV, 2=XLSX/Excel) [Por defecto: 1]: 1
```
- Ingrese **1** para exportar a CSV (predeterminado)
- Ingrese **2** para exportar a XLSX/Excel (**requiere Microsoft Excel instalado**)
- Si presiona ENTER sin escribir nada, se usará CSV por defecto

### Proceso de Ejecución

Una vez ingresados todos los datos:

1. El script **verifica** la existencia de las carpetas `consultas` y `resultados`
2. Si faltan carpetas, muestra un error y espera que presione ENTER
3. Busca la instalación de Oracle SQLcl en las rutas estándar
4. **Configura variables de entorno Java** para evitar warnings
5. **Valida la conexión** a Oracle antes de procesar consultas
6. Cuenta cuántos archivos `.sql` hay en la carpeta `consultas`
7. **Procesa cada consulta** una por una:
   - **VALIDACIÓN:** Verifica que sea solo consulta SELECT
   - **PARÁMETROS:** Si existe archivo `.txt`, solicita valores de parámetros
   - **CONEXIÓN:** Conecta a Oracle con las credenciales proporcionadas
   - **EJECUCIÓN:** Ejecuta la consulta SQL con parámetros sustituidos
   - **EXPORTACIÓN:** Exporta los resultados con un nombre único
   - Si eligió XLSX, convierte automáticamente de CSV a Excel
8. Muestra un resumen del procesamiento con colores
9. **Siempre espera** que presione ENTER antes de cerrar

### Ejemplo Completo con Parámetros

#### Archivo de Consulta:
**`consultas/ventas_por_periodo.sql`:**
```sql
SELECT 
    producto_id,
    nombre_producto,
    SUM(cantidad) as unidades_vendidas,
    SUM(total) as ingresos_totales
FROM ventas_detalle
WHERE fecha_venta BETWEEN '&fecha_inicio' AND '&fecha_fin'
  AND region = '&region'
GROUP BY producto_id, nombre_producto
ORDER BY ingresos_totales DESC;
```

#### Archivo de Parámetros:
**`consultas/ventas_por_periodo.txt`:**
```
fecha_inicio;fecha_fin;region
```

#### Ejecución del Script:
```
Procesando: ventas_por_periodo.sql
  > Salida: ventas_por_periodo_20241218_143022.csv
  > Leyendo definiciones de parametros desde: ventas_por_periodo.txt
  Ingrese valor para 'fecha_inicio': 2024-01-01
  Ingrese valor para 'fecha_fin': 2024-12-31
  Ingrese valor para 'region': Norte
  > Validando que sea solo consulta SELECT...
  [OK] Validacion de SELECT exitosa
  > Ejecutando consulta (timeout: 30 minutos)...
  [OK] Archivo CSV generado: ventas_por_periodo_20241218_143022.csv
```

### Nombres de Archivos de Salida

Los archivos de resultados se guardan con el siguiente formato:
```
[nombre_consulta]_[fecha]_[hora].[extensión]
```

**Ejemplos:**
- `ventas_2024_20241217_143055.csv`
- `clientes_activos_20241217_144512.xlsx`

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

#### Error: Script SQL inválido (contiene operaciones no SELECT)
```
Procesando: consulta_peligrosa.sql
  > Salida: consulta_peligrosa_20241218_143022.csv
  > Validando que sea solo consulta SELECT...
  [ERROR] Script SQL invalido
  Razon: Contiene operacion no permitida: UPDATE
  Este script solo permite consultas SELECT.
  Operaciones prohibidas: INSERT, UPDATE, DELETE, DROP, TRUNCATE, CREATE, ALTER, PL/SQL, etc.
```
**Solución:** Revise el archivo SQL y asegúrese de que solo contenga consultas SELECT.

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

#### Error: Java no encontrado o versión incorrecta
```
Error: Java no está instalado o la versión es incorrecta
SQLcl requiere Java 17 o superior
```

**Solución:**
1. Instale Java 17 o superior (ver sección "Java Runtime Environment")
2. Configure JAVA_HOME correctamente
3. Agregue `%JAVA_HOME%\bin` al PATH
4. Abra una nueva ventana de PowerShell y verifique: `java -version`

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

#### Error al convertir a Excel
```
[ERROR] Fallo la conversion a Excel
Detalles: ...
Nota: Se requiere Microsoft Excel instalado para exportar a XLSX
```

**Posibles causas:**
- Microsoft Excel no está instalado
- Excel está abierto y bloqueando archivos
- Permisos insuficientes para crear archivos COM

**Solución:**
1. Instale Microsoft Excel 2007 o superior
2. Cierre todas las instancias de Excel antes de ejecutar el script
3. Use formato CSV si no tiene Excel instalado
4. Ejecute PowerShell como administrador si hay problemas de permisos

## 📊 Ejemplo Completo de Uso

### Paso a Paso

1. **Instalar Java 17 o superior**
   - Descargue desde: https://adoptium.net/ (recomendado) o https://www.oracle.com/java/technologies/downloads/
   - Instale y configure JAVA_HOME
   - Agregue `%JAVA_HOME%\bin` al PATH
   - Verifique: `java -version`

2. **Descargar e instalar Oracle SQLcl**
   - Descargue desde: https://www.oracle.com/database/sqldeveloper/technologies/sqlcl/download/
   - Extraiga a `C:\oracle\sqlcl\`
   - Verifique: `C:\oracle\sqlcl\bin\sql.exe -V`

3. **Configurar permisos de PowerShell** (Opción 2):
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

4. **Crear estructura de carpetas:**
   ```
   C:\MisConsultas\
   ├── ejecutar_consultas_oracle.ps1
   ├── consultas\
   └── resultados\
   ```

5. **Agregar consultas SQL** en `C:\MisConsultas\consultas\`:
   - `ventas_mensuales.sql`
   - `top_clientes.sql`
   - `inventario_actual.sql`

6. **Agregar archivos de parámetros** (opcional):
   - `ventas_mensuales.txt` con contenido: `anio;mes`
   - `top_clientes.txt` con contenido: `limite_registros`

7. **Ejecutar el script:**
   - Click derecho en `ejecutar_consultas_oracle.ps1`
   - "Ejecutar con PowerShell"

8. **Ingresar datos:**
   ```
   Usuario: admin_ventas
   Opcion de contrasena: 1
   Host: db-server.empresa.com
   Puerto: 1521
   SID/Service: PRODDB
   Formato: 1
   ```

9. **Ingresar parámetros** (si aplica):
   ```
   Para consulta 'ventas_mensuales.sql':
   Ingrese valor para 'anio': 2024
   Ingrese valor para 'mes': 12
   ```

10. **Ver resultados** en `C:\MisConsultas\resultados\`:
    ```
    ventas_mensuales_20241217_150033.csv
    top_clientes_20241217_150033.csv
    inventario_actual_20241217_150033.csv
    ```
    
    O si eligió Excel (opción 2):
    ```
    ventas_mensuales_20241217_150033.xlsx
    top_clientes_20241217_150033.xlsx
    inventario_actual_20241217_150033.xlsx
    ```

11. **Opcional - Restaurar seguridad:**
    ```powershell
    Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope CurrentUser
    ```

## 🔧 Solución de Problemas

### El script no se ejecuta (Error de Política)
- Siga la sección **"Configuración de Permisos de PowerShell"**
- Use el método de **Bypass** si no puede cambiar políticas

### Java no funciona o no se encuentra
- Verifique que instaló Java 17 o superior: `java -version`
- Verifique JAVA_HOME: `echo $env:JAVA_HOME`
- Verifique PATH: `echo $env:PATH` (debe contener `%JAVA_HOME%\bin`)
- Abra una NUEVA ventana de PowerShell después de configurar variables
- Si tiene múltiples versiones, asegúrese de que Java 17+ esté primero en PATH

### SQLcl no se encuentra
- Verifique que extrajo SQLcl correctamente
- Asegúrese de que `sql.exe` existe en `bin\`
- Verifique que Java está funcionando antes de ejecutar SQLcl
- Intente agregar la ruta al PATH del sistema

### SQLcl no inicia (Error de Java)
- Ejecute: `C:\oracle\sqlcl\bin\sql.exe -V`
- Si falla, verifique que Java 17+ está instalado
- Asegúrese de que JAVA_HOME apunta a una instalación válida de Java

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

### Formato CSV no se ve bien
- Abra el CSV con un editor de texto primero
- Asegúrese de que su Excel esté configurado para UTF-8
- Use "Importar datos" en Excel en lugar de doble click

### Problemas con exportación a Excel (XLSX)
- **Requiere Microsoft Excel instalado** en el sistema
- Cierre todas las instancias de Excel antes de ejecutar el script
- Si no tiene Excel, use formato CSV (opción 1)
- El script convierte automáticamente CSV a XLSX usando Excel COM

### La ventana se cierra inmediatamente
- **Nunca debería ocurrir** gracias al bloque `finally`
- Si ocurre, ejecute desde PowerShell directamente para ver el error

### Problemas con parámetros
- **Los nombres en el .txt deben coincidir** exactamente con los nombres de variables en el SQL
- Use solo letras, números y guiones bajos en nombres de parámetros
- El archivo .txt debe usar codificación UTF-8 sin BOM
- Asegúrese de que el archivo .txt no tenga espacios adicionales al final de las líneas

## 🎨 Características del Script PowerShell

### Ventajas de Usar SQLcl

✅ **Gratuito y ligero** - No requiere Oracle Client completo  
✅ **Multiplataforma** - Funciona en Windows, Linux, Mac  
✅ **Formato CSV nativo** - Exportación directa y eficiente  
✅ **Rápido y eficiente** - Mejor rendimiento que SQL*Plus  
✅ **Actualizado** - Soporta las últimas versiones de Oracle  
✅ **Sin instalación compleja** - Solo extraer y usar  

### Cómo Funciona la Exportación a Excel

1. **SQLcl exporta a CSV**: Oracle SQLcl genera el archivo CSV (formato nativo)
2. **PowerShell convierte a XLSX**: Si eligió Excel, el script usa COM Automation para convertir
3. **Resultado final**: Archivo Excel nativo (.xlsx) listo para usar

**Nota:** La conversión a XLSX requiere Microsoft Excel instalado. Si no lo tiene, use CSV que es universalmente compatible.

### Ventajas del Script PowerShell

✅ **Try/Catch/Finally nativo** - Manejo robusto de errores  
✅ **Validación de tipos** - Puerto debe ser número  
✅ **Colores en consola** - Mejor experiencia visual  
✅ **Mejor manejo de strings** - Sin problemas con espacios  
✅ **Objetos y propiedades** - Código más limpio  
✅ **Contraseña enmascarada** - Mayor seguridad  
✅ **Conversión automática a Excel** - CSV a XLSX con un click  
✅ **Validación de solo SELECT** - Seguridad mejorada  
✅ **Parámetros dinámicos** - Consultas parametrizadas flexibles  
✅ **Contraseña por defecto** - Configuración simplificada  

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
- La conversión a XLSX usa **Excel COM Automation** (requiere Excel instalado)
- Formato CSV funciona sin necesidad de Microsoft Excel

## 🔒 Seguridad y Mejores Prácticas

### Recomendaciones de Seguridad

1. **Credenciales:**
   - La **contraseña por defecto** es más segura para entornos controlados
   - Nunca guarde contraseñas en el script o archivos de texto plano
   - La contraseña se enmascara automáticamente durante la entrada
   - Considere usar Oracle Wallet para credenciales frecuentes

2. **Validación de Consultas:**
   - El script valida automáticamente que solo sean consultas SELECT
   - Revise todas las consultas antes de ejecutarlas
   - Evite consultas con `DELETE` o `UPDATE` sin `WHERE`
   - Use permisos de solo lectura cuando sea posible

3. **Parámetros:**
   - Los archivos .txt solo contienen nombres de parámetros, no valores
   - Los valores se solicitan interactivamente y no se almacenan
   - Use nombres descriptivos para los parámetros

4. **Permisos de PowerShell:**
   - Use `RemoteSigned` en lugar de `Unrestricted`
   - Restaure a `Restricted` cuando termine de usar el script

5. **Exportación a Excel:**
   - Si usa formato XLSX, asegúrese de cerrar Excel antes de ejecutar
   - Los archivos Excel pueden ser más grandes que CSV
   - CSV es más seguro y portable si no necesita formato específico

6. **Red:**
   - Use conexiones seguras (Oracle Advanced Security)
   - Considere VPN para conexiones remotas
   - Verifique configuraciones de firewall

## 🆚 Comparación: SQLcl vs SQL*Plus

| Característica | SQLcl | SQL*Plus |
|---------------|-------|----------|
| **Gratuito** | ✅ Sí | ✅ Sí |
| **Formato CSV** | ✅ Nativo | ⚠️ Manual |
| **Formato Excel** | ⚠️ Via conversión | ❌ No |
| **Instalación** | ✅ Extraer ZIP | ⚠️ Requiere Oracle Client |
| **Tamaño** | ~50 MB | ~200+ MB |
| **Scripting** | ✅ Excelente | ✅ Bueno |
| **Multiplataforma** | ✅ Sí | ✅ Sí |
| **Moderno** | ✅ Sí | ❌ Antiguo |
| **Automatización** | ✅ Excelente | ✅ Bueno |
| **Validación SQL** | ✅ Con este script | ❌ No |

**Conclusión:** SQLcl es la mejor opción para automatización moderna con Oracle.

## 📄 Licencia

Este script es de uso libre. Oracle SQLcl está bajo licencia Oracle Technology Network License Agreement.

## 🤝 Contribuciones

Para reportar problemas o sugerir mejoras, por favor contacte al desarrollador del proyecto.

---

**Versión del Script:** 4.0 (PowerShell + Oracle SQLcl)  
**Características Principales:** Validación SELECT, parámetros dinámicos, contraseña por defecto  
**Fecha:** Diciembre 2025  
**Compatible con:** Oracle SQLcl 23.x+, Windows 10+, PowerShell 5.1+, Oracle 11g-23c