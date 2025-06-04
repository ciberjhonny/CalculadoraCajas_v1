# 🏭 Calculadora Unificada de Cajas de Cartón Corrugado

## 📋 Descripción
Sistema completo que unifica todas las calculadoras de desarrollo de cajas de cartón con lectura automática de factores desde Excel.

## 🎯 Características Principales
- ✅ **8 tipos de cajas diferentes** unificados en un solo programa
- ✅ **Lectura automática** de factores desde archivo Excel
- ✅ **Interfaz gráfica intuitiva** con Tkinter
- ✅ **Cálculos automáticos** de dimensiones y costos
- ✅ **Generación de ejecutable** independiente (.exe)

## 📦 Tipos de Cajas Soportadas

1. **Corrugado Sencillo** - Caja básica
2. **Corrugado Doble** - Caja reforzada
3. **Medio Fondo Sencillo** - Con refuerzo inferior sencillo
4. **Medio Fondo Doble** - Con refuerzo inferior doble
5. **Doble Armado Sencillo** - Construcción doble sencilla
6. **Doble Armado Doble** - Construcción doble reforzada
7. **Doble Armado Medio Fondo Sencillo** - Combinación completa sencilla
8. **Doble Armado Medio Fondo Doble** - Combinación completa doble

## 🚀 Instalación y Uso

### Opción 1: Usar el Ejecutable (Recomendado)
```bash
# 1. Descargar los archivos:
# - CalculadoraCajas.exe
# - Factores.xlsx (si no está incluido)

# 2. Ejecutar directamente:
CalculadoraCajas.exe
```

### Opción 2: Ejecutar desde Python
```bash
# 1. Instalar Python 3.8 o superior
# 2. Instalar dependencias:
pip install pandas openpyxl

# 3. Ejecutar:
python main.py
```

### Opción 3: Generar tu Propio Ejecutable
```bash
# 1. Tener todos los archivos:
# - main.py
# - Factores.xlsx
# - build_exe.py

# 2. Ejecutar el generador:
python build_exe.py

# 3. El ejecutable estará en dist/CalculadoraCajas.exe
```

## 📁 Estructura de Archivos

```
proyecto/
├── main.py                    # Programa principal
├── Factores.xlsx             # Archivo con factores y resistencias
├── build_exe.py              # Generador de ejecutable
├── requirements.txt          # Dependencias Python
├── INSTRUCCIONES_COMPLETAS.md # Este archivo
└── dist/                     # Carpeta con ejecutable generado
    └── CalculadoraCajas.exe
```

## 🔧 Formato del Archivo Excel (Factores.xlsx)

El archivo debe tener estas columnas:
- **Columna A**: RESISTENCIA (ej: "CORRUGADO SENCILLO KRAFT")
- **Columna B**: MATERIAL (ej: "KRAFT", "BLANCO")
- **Columna C**: PAPEL (ej: "NACIONAL", "AMERICANO EXT.")
- **Columna D**: Resistencia (código, ej: "32K", "42DCK")
- **Columna E**: P.LAMINA ACTUAL o FACTORES (valor numérico)

### Ejemplo de datos:
```
RESISTENCIA                 | MATERIAL | PAPEL     | Resistencia | Factor
CORRUGADO SENCILLO KRAFT   | KRAFT    | NACIONAL  | 32K         | 1.3456751616
DOBLE CORRUGADO KRAFT      | KRAFT    | NACIONAL  | 42DCK       | 1.9146710016
```

## 💡 Uso del Programa

### 1. Carga de Factores
- Al iniciar, el programa busca automáticamente `Factores.xlsx`
- También puedes cargar manualmente otro archivo Excel
- El status se muestra en la interfaz (✅ cargado / ❌ error)

### 2. Selección de Parámetros
1. **Tipo de Caja**: Elegir del menú desplegable
2. **Dimensiones**: Ingresar largo, ancho, alto y ceja en cm
3. **Material**: Se actualiza automáticamente según el tipo de caja
4. **Resistencia**: Se filtra según el material seleccionado
5. **Parámetros Financieros**: Ganancia % y descuento %

### 3. Cálculo
- Hacer clic en "🧮 CALCULAR"
- Los resultados aparecen en tiempo real
- Incluye dimensiones finales, área y costos

### 4. Resultados
El programa muestra:
- **Dimensiones calculadas** (largo total, ancho total, área)
- **Configuración de aletones** según el tipo
- **Cálculos financieros** completos
- **Material y factor** utilizados
- **Timestamp** de cuando se hizo el cálculo

## 🔍 Fórmulas Utilizadas

### Corrugado Sencillo:
- Largo Total: `((largo + 0.4) + (ancho + 0.5)) * 2 + ceja`
- Aletón: `((ancho + 0.5 + ajuste) / 2)`
- Central: `alto + 0.8`

### Corrugado Doble:
- Largo Total: `((largo + 0.8) + (ancho + 1)) * 2 + ceja`
- Aletón: `((ancho + 1 + ajuste) / 2)`
- Central: `alto + 1.5`

### Medio Fondo:
- Mismo cálculo base pero `ancho_total = aletón + central` (sin segundo aletón)

### Doble Armado:
- Mismo cálculo base pero `área_final = área * 2`

### Cálculos Financieros:
- Costo Materia: `(área * factor) / 10000`
- Precio Venta: `costo * ganancia`
- Precio Ajustado: `precio * (1 + descuento/100)`
- Ganancia Total: `precio_ajustado - costo`

## 🛠️ Personalización

### Agregar Nuevos Tipos de Caja:
1. Editar `self.tipos_caja` en `main.py`
2. Agregar lógica en `calcular_dimensiones()`
3. Regenerar ejecutable

### Modificar Factores:
1. Editar `Factores.xlsx` con Excel
2. Mantener la estructura de columnas
3. El programa carga automáticamente los cambios

### Cambiar Interfaz:
- Los estilos están en la función `setup_ui()`
- Usar temas de ttk para diferentes apariencias
- Agregar validaciones o campos adicionales

## 🚨 Solución de Problemas

### Error: "Factores.xlsx no encontrado"
- Verificar que el archivo esté en la misma carpeta que el programa
- O usar el botón "Cargar Factores.xlsx" para seleccionar otro archivo

### Error: "Valores numéricos inválidos"
- Verificar que todos los campos tengan números válidos
- Usar punto (.) como separador decimal, no coma (,)

### Error: "No se puede cargar Excel"
- Verificar que el archivo no esté abierto en Excel
- Comprobar que el formato sea .xlsx
- Revisar que las columnas estén en el orden correcto

### El ejecutable no funciona:
- Verificar que `Factores.xlsx` esté incluido en el build
- Regenerar con `python build_exe.py`
- Comprobar antivirus (puede bloquear ejecutables nuevos)

## 📞 Soporte

Para soporte técnico o modificaciones:
1. Revisar este manual completo
2. Verificar la estructura del archivo Excel
3. Comprobar que Python y dependencias estén correctas
4. Regenerar el ejecutable si es necesario

## 📈 Versiones

- **v1.0**: Sistema unificado completo con interfaz gráfica
- Incluye todos los 8 tipos de caja originales
- Lectura automática desde Excel
- Generación de ejecutable independiente

---
*Calculadora desarrollada para optimizar el proceso de cotización y desarrollo de cajas de cartón corrugado.*