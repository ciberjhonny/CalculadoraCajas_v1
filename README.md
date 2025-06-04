# CalculadoraCajas_v1
Sistema unificado para el cálculo y cotización de cajas de cartón corrugado con interfaz gráfica intuitiva y lectura automática de factores desde Excel.
[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey.svg)](https://github.com/tu-usuario/calculadora-cajas)

> **Sistema unificado para el cálculo y cotización de cajas de cartón corrugado con interfaz gráfica intuitiva y lectura automática de factores desde Excel.**

## 🎯 ¿Qué hace este programa?

Esta calculadora está diseñada específicamente para la industria del cartón corrugado, permitiendo calcular de manera precisa y rápida las dimensiones, áreas y costos de 8 tipos diferentes de cajas. Es ideal para empresas que necesitan cotizar y diseñar envases de cartón de manera eficiente.

### ✨ Características principales

- 🚀 **Interfaz gráfica intuitiva** - No necesitas ser programador para usarla
- 📊 **8 tipos de cajas soportadas** - Desde corrugado sencillo hasta doble armado
- 📁 **Lectura automática de Excel** - Carga factores y resistencias desde tu archivo
- 💰 **Cálculos financieros completos** - Costos, precios y ganancias
- 📦 **Ejecutable independiente** - Funciona sin instalar Python
- 🔄 **Actualizaciones dinámicas** - Los materiales se filtran automáticamente

## 📦 Tipos de cajas soportadas

1. **Corrugado Sencillo** - Caja estándar básica
2. **Corrugado Doble** - Caja reforzada para mayor resistencia
3. **Medio Fondo Sencillo** - Con refuerzo inferior sencillo
4. **Medio Fondo Doble** - Con refuerzo inferior doble
5. **Doble Armado Sencillo** - Construcción doble para productos pesados
6. **Doble Armado Doble** - Máxima resistencia disponible
7. **Doble Armado Medio Fondo Sencillo** - Combinación especializada
8. **Doble Armado Medio Fondo Doble** - Solución premium completa

## 🚀 Instalación rápida

### Opción 1: Usar el ejecutable (Recomendado)
```bash
# 1. Descarga el ejecutable desde Releases
# 2. Coloca tu archivo Factores.xlsx en la misma carpeta
# 3. ¡Ejecuta y listo!
```

### Opción 2: Ejecutar desde código fuente
```bash
# Clona el repositorio
git clone https://github.com/tu-usuario/calculadora-cajas.git
cd calculadora-cajas

# Instala las dependencias
pip install -r requirements.txt

# Ejecuta el programa
python main.py
```

### Opción 3: Generar tu propio ejecutable
```bash
# Instala PyInstaller
pip install pyinstaller

# Genera el ejecutable
python build_exe.py

# El .exe estará en la carpeta dist/
```

## 📋 Uso del programa

### 1. **Preparación inicial**
- Asegúrate de tener tu archivo `Factores.xlsx` con los datos de materiales y resistencias
- El programa lo buscará automáticamente al iniciar

### 2. **Proceso de cálculo**
1. Selecciona el **tipo de caja** que necesitas
2. Ingresa las **dimensiones** (largo, ancho, alto, ceja)
3. Elige el **material y resistencia** (se actualizan automáticamente)
4. Ajusta los **parámetros financieros** (ganancia y descuento)
5. Presiona **"CALCULAR"** y obtén resultados instantáneos

### 3. **Resultados obtenidos**
- ✅ Dimensiones finales calculadas
- ✅ Área total de material necesario
- ✅ Configuración de aletones
- ✅ Costos de materia prima
- ✅ Precios de venta sugeridos
- ✅ Análisis de ganancia
- ✅ Cantidad mínima de compra

## 📊 Estructura del proyecto

```
calculadora-cajas/
├── main.py                    # Programa principal
├── build_exe.py              # Generador de ejecutable
├── requirements.txt          # Dependencias Python
├── Factores.xlsx            # Datos de materiales (ejemplo)
├── INSTRUCCIONES_COMPLETAS.md # Manual técnico completo
└── dist/                     # Ejecutables generados
    └── CalculadoraCajas.exe
```

## 🔧 Configuración del archivo Excel

Tu archivo `Factores.xlsx` debe tener esta estructura:

| Columna A | Columna B | Columna C | Columna D | Columna E |
|-----------|-----------|-----------|-----------|-----------|
| RESISTENCIA | MATERIAL | PAPEL | Código | Factor |
| CORRUGADO SENCILLO KRAFT | KRAFT | NACIONAL | 32K | 1.3456751616 |
| DOBLE CORRUGADO KRAFT | KRAFT | NACIONAL | 42DCK | 1.9146710016 |

## 💡 Fórmulas utilizadas

El programa implementa fórmulas específicas para cada tipo de caja:

**Corrugado Sencillo:**
```
Largo Total = ((largo + 0.4) + (ancho + 0.5)) × 2 + ceja
Aletón = ((ancho + 0.5 + ajuste) ÷ 2)
Central = alto + 0.8
```

**Corrugado Doble:**
```
Largo Total = ((largo + 0.8) + (ancho + 1)) × 2 + ceja
Aletón = ((ancho + 1 + ajuste) ÷ 2)
Central = alto + 1.5
```

**Cálculos Financieros:**
```
Costo Materia = (área × factor) ÷ 1000
Precio Venta = costo × ganancia
Precio Final = precio × (1 + descuento/100)
```

## 🛠️ Tecnologías utilizadas

- **Python 3.8+** - Lenguaje principal
- **Tkinter** - Interfaz gráfica nativa
- **Pandas** - Procesamiento de datos Excel
- **OpenPyXL** - Lectura de archivos .xlsx
- **PyInstaller** - Generación de ejecutables

## 🤝 Contribuir al proyecto

¡Las contribuciones son bienvenidas! Si trabajas en la industria del cartón o eres desarrollador:

1. **Fork** el repositorio
2. Crea una **branch** para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. **Commit** tus cambios (`git commit -am 'Agrega nueva funcionalidad'`)
4. **Push** a la branch (`git push origin feature/nueva-funcionalidad`)
5. Abre un **Pull Request**

### 🔍 Ideas para contribuir
- Agregar nuevos tipos de cajas
- Mejorar la interfaz visual
- Añadir exportación a PDF
- Implementar base de datos local
- Crear versión web

## 📝 Licencia

Este proyecto está bajo la Licencia MIT. Ver el archivo [LICENSE](LICENSE) para más detalles.

## 📞 Soporte y contacto

- 🐛 **Reportar bugs**: [Issues](https://github.com/tu-usuario/calculadora-cajas/issues)
- 💡 **Sugerir mejoras**: [Discussions](https://github.com/tu-usuario/calculadora-cajas/discussions)
- 📧 **Contacto directo**: tu-email@ejemplo.com

## 🏆 Casos de uso reales

Esta calculadora es perfecta para:
- **Empresas de packaging** que necesitan cotizar rapidamente
- **Diseñadores de envases** que requieren precisión en medidas
- **Distribuidores de cartón** para calcular costos al instante
- **Startups** que necesitan envases personalizados
- **Departamentos de compras** para validar presupuestos

---

⭐ **¿Te gusta el proyecto?** ¡Dale una estrella y compártelo con otros profesionales del sector!

**Desarrollado con ❤️ para la industria del cartón corrugado**
