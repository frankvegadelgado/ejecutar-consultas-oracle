# Ejecutor de Consultas SQL para Oracle - Guía de Usuario

## 📋 Requisitos Previos

### 1. Sistema Operativo
- Windows 10 o superior

### 2. Java - Instalación Requerida

**Oracle SQLcl requiere Java 17 o superior** para funcionar. **DEBE instalar Java antes de continuar.**

#### Paso 1: Verificar si Java está Instalado
Abra PowerShell o CMD y ejecute:
```powershell
java -version
```

Si muestra un error o una versión menor a 17, necesita instalar Java.

#### Paso 2: Instalar Java (Oracle JDK)
1. Visite: **https://www.oracle.com/java/technologies/downloads/**
2. Descargue **Java 17 o superior** (versión LTS)
3. Ejecute el instalador `.exe` descargado
4. Siga el asistente de instalación completo
5. **No omita ningún paso** durante la instalación

#### Paso 3: Verificar la Instalación
Después de instalar, cierre y abra una nueva ventana de PowerShell/CMD y ejecute:
```powershell
java -version
```

Debería mostrar algo como:
```
java version "17.0.9" 2023-10-17 LTS
Java(TM) SE Runtime Environment (build 17.0.9+11-LTS-201)
```

### 3. Oracle SQLcl - Instalación Requerida

**DEBE instalar Oracle SQLcl para que el programa funcione.**

#### Paso 1: Descargar Oracle SQLcl
1. Visite: **https://www.oracle.com/database/sqldeveloper/technologies/sqlcl/download/**
2. Descargue la versión más reciente (archivo `.zip`)

#### Paso 2: Instalar Oracle SQLcl
1. **Extraiga el archivo ZIP completo**
2. Extraiga a: `C:\oracle\sqlcl\`
   - Esta ruta es **IMPORTANTE** - úsela exactamente así
3. La estructura debe quedar así:
   ```
   C:\oracle\sqlcl\
   ├── bin\
   │   ├── sql.exe
   │   └── sql.bat
   ├── lib\
   └── LICENSE.txt
   ```

#### Paso 3: Verificar la Instalación
Abra PowerShell o CMD y ejecute:
```powershell
C:\oracle\sqlcl\bin\sql.exe -V
```

Debería mostrar:
```
SQLcl: Release 23.4 Production
Build: 23.4.0.341.0944
```

## 🚀 Configuración del Proyecto

### Estructura de Carpetas

**ANTES de ejecutar el programa, DEBE crear estas carpetas:**

En el mismo directorio donde está `ejecutar_consultas_oracle.exe`, cree:

```
📁 [Directorio del Proyecto]
├── 📄 ejecutar_consultas_oracle.exe
├── 📁 consultas/          ← CREAR ESTA CARPETA
│   └── (aquí se colocan los archivos .sql)
└── 📁 resultados/         ← CREAR ESTA CARPETA
    └── (aquí se guardarán los archivos generados)
```

### Crear las Carpetas

#### Método 1: Desde el Explorador de Archivos
1. Haga clic derecho en el área vacía donde está el archivo `.exe`
2. Seleccione **"Nuevo" > "Carpeta"**
3. Nombre la carpeta: `consultas`
4. Repita para crear: `resultados`

#### Método 2: Desde CMD
```cmd
mkdir consultas
mkdir resultados
```

## 💻 Uso del Programa

### Ejecutar el Programa
1. **Doble clic** en `ejecutar_consultas_oracle.exe`
2. Se abrirá una ventana de consola

### Datos de Entrada Requeridos

El programa solicitará:

#### 1. Usuario de Oracle
```
Ingrese el usuario de Oracle:
```
- Ingrese su nombre de usuario de Oracle
- Presione **ENTER**

#### 2. Contraseña
```
Opciones de contrasena:
  1. Usar contrasena por defecto (RECOMENDADO)
  2. Ingresar contrasena personalizada

Seleccione opcion de contrasena (1 o 2) [Por defecto: 1]: 
```
- Ingrese `1` para usar la contraseña por defecto (recomendado)
- O ingrese `2` para ingresar una contraseña personalizada

#### 3. Host
```
Ingrese el host (ej: localhost, 192.168.1.100):
```
- Ingrese la dirección del servidor Oracle
- Presione **ENTER**

#### 4. Puerto
```
Ingrese el puerto (ej: 1521):
```
- Ingrese el puerto (Oracle usa 1521 por defecto)
- Presione **ENTER**

#### 5. SID o Service Name
```
Ingrese el SID o Service Name (ej: ORCL, XE, PDB1):
```
- Ingrese el nombre de la base de datos
- Presione **ENTER**

#### 6. Formato de Salida
```
Formato de salida (1=CSV, 2=XLSX/Excel) [Por defecto: 1]:
```
- Ingrese `1` para CSV
- Ingrese `2` para Excel (requiere Microsoft Excel instalado)
- Presione **ENTER**

### Proceso de Ejecución

1. Verifica que las carpetas `consultas` y `resultados` existan
2. Busca archivos `.sql` en la carpeta `consultas`
3. Procesa cada archivo SQL
4. Guarda resultados en la carpeta `resultados`
5. Muestra resumen final
6. Espera que presione **ENTER** para cerrar

## ❌ Errores Comunes

### Error: Java no instalado
```
Error: Java no está instalado
```
**Solución:** Instale Java 17 o superior siguiendo los pasos en "Java - Instalación Requerida".

### Error: Oracle SQLcl no encontrado
```
[ERROR] No se pudo encontrar Oracle SQLcl instalado
```
**Solución:** Instale Oracle SQLcl siguiendo los pasos en "Oracle SQLcl - Instalación Requerida".

### Error: Carpetas no encontradas
```
[ERROR] No se encontro la carpeta 'consultas'
```
**Solución:** Cree las carpetas `consultas` y `resultados` como se explica en "Estructura de Carpetas".

### Error: No hay archivos SQL
```
[ADVERTENCIA] No se encontraron archivos .sql en la carpeta 'consultas'
```
**Solución:** Coloque archivos `.sql` dentro de la carpeta `consultas`.

---

**Fecha:** Diciembre 2025