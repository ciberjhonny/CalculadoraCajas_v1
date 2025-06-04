#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Calculadora Unificada de Cajas de Cartón Corrugado
Integra todos los tipos de desarrollo con factores desde Excel
Versión: 1.0
Autor: Sistema Unificado
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
from typing import Dict, List, Tuple, Optional
import math

class CalculadoraCajas:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Calculadora de Cajas de Cartón Corrugado - v1.0")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Variables para datos
        self.factores_df = None
        self.tipos_caja = {
            "Corrugado Sencillo": "sencillo",
            "Corrugado Doble": "doble", 
            "Medio Fondo Sencillo": "medio_fondo_sencillo",
            "Medio Fondo Doble": "medio_fondo_doble",
            "Doble Armado Sencillo": "doble_armado_sencillo",
            "Doble Armado Doble": "doble_armado_doble",
            "Doble Armado Medio Fondo Sencillo": "doble_armado_medio_fondo_sencillo",
            "Doble Armado Medio Fondo Doble": "doble_armado_medio_fondo_doble"
        }
        
        self.setup_ui()
        self.cargar_factores_default()
        
    def setup_ui(self):
        """Configurar la interfaz de usuario"""
        # Título principal
        title_frame = ttk.Frame(self.root)
        title_frame.pack(pady=10, padx=20, fill="x")
        
        title_label = ttk.Label(title_frame, text="🏭 Calculadora de Cajas de Cartón Corrugado", 
                               font=("Arial", 16, "bold"))
        title_label.pack()
        
        # Frame para cargar archivo Excel
        excel_frame = ttk.LabelFrame(self.root, text="📁 Archivo de Factores", padding=10)
        excel_frame.pack(pady=5, padx=20, fill="x")
        
        ttk.Button(excel_frame, text="Cargar Factores.xlsx", 
                  command=self.cargar_archivo_excel).pack(side="left", padx=5)
        
        self.excel_status = ttk.Label(excel_frame, text="📊 Archivo no cargado", foreground="red")
        self.excel_status.pack(side="left", padx=10)
        
        # Frame principal con dos columnas
        main_frame = ttk.Frame(self.root)
        main_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Columna izquierda - Entrada de datos
        left_frame = ttk.LabelFrame(main_frame, text="📐 Datos de Entrada", padding=10)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # Tipo de caja
        ttk.Label(left_frame, text="Tipo de Caja:").pack(anchor="w")
        self.tipo_caja = ttk.Combobox(left_frame, values=list(self.tipos_caja.keys()), 
                                     state="readonly", width=30)
        self.tipo_caja.set("Corrugado Sencillo")
        self.tipo_caja.pack(pady=2, fill="x")
        self.tipo_caja.bind('<<ComboboxSelected>>', self.on_tipo_change)
        
        # Dimensiones
        dimensions_frame = ttk.Frame(left_frame)
        dimensions_frame.pack(pady=10, fill="x")
        
        ttk.Label(dimensions_frame, text="Largo (cm):").grid(row=0, column=0, sticky="w", pady=2)
        self.largo_var = tk.StringVar(value="30.00")
        ttk.Entry(dimensions_frame, textvariable=self.largo_var, width=15).grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(dimensions_frame, text="Ancho (cm):").grid(row=1, column=0, sticky="w", pady=2)
        self.ancho_var = tk.StringVar(value="20.30")
        ttk.Entry(dimensions_frame, textvariable=self.ancho_var, width=15).grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Label(dimensions_frame, text="Alto (cm):").grid(row=2, column=0, sticky="w", pady=2)
        self.alto_var = tk.StringVar(value="40.00")
        ttk.Entry(dimensions_frame, textvariable=self.alto_var, width=15).grid(row=2, column=1, padx=5, pady=2)
        
        ttk.Label(dimensions_frame, text="Ceja (cm):").grid(row=3, column=0, sticky="w", pady=2)
        self.ceja_var = tk.StringVar(value="3.5")
        ttk.Entry(dimensions_frame, textvariable=self.ceja_var, width=15).grid(row=3, column=1, padx=5, pady=2)
        
        # Selección de material y resistencia
        material_frame = ttk.Frame(left_frame)
        material_frame.pack(pady=10, fill="x")
        
        ttk.Label(material_frame, text="Tipo de Material:").pack(anchor="w")
        self.material_combo = ttk.Combobox(material_frame, state="readonly", width=30)
        self.material_combo.pack(pady=2, fill="x")
        self.material_combo.bind('<<ComboboxSelected>>', self.actualizar_resistencias)
        
        ttk.Label(material_frame, text="Resistencia:").pack(anchor="w", pady=(10,0))
        self.resistencia_combo = ttk.Combobox(material_frame, state="readonly", width=30)
        self.resistencia_combo.pack(pady=2, fill="x")
        
        # Parámetros financieros
        financial_frame = ttk.Frame(left_frame)
        financial_frame.pack(pady=10, fill="x")
        
        ttk.Label(financial_frame, text="Ganancia %:").grid(row=0, column=0, sticky="w", pady=2)
        self.utilidad_var = tk.StringVar(value="1.50")
        ttk.Entry(financial_frame, textvariable=self.utilidad_var, width=15).grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(financial_frame, text="Descuento %:").grid(row=1, column=0, sticky="w", pady=2)
        self.descuento_var = tk.StringVar(value="10.00")
        ttk.Entry(financial_frame, textvariable=self.descuento_var, width=15).grid(row=1, column=1, padx=5, pady=2)
        
        # Botón calcular
        ttk.Button(left_frame, text="🧮 CALCULAR", command=self.calcular_todo, 
                  style="Accent.TButton").pack(pady=20, fill="x")
        
        # Columna derecha - Resultados
        right_frame = ttk.LabelFrame(main_frame, text="📊 Resultados", padding=10)
        right_frame.pack(side="right", fill="both", expand=True)
        
        # Área de texto para resultados
        self.result_text = tk.Text(right_frame, wrap="word", font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        self.result_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def cargar_factores_default(self):
        """Intenta cargar Factores.xlsx del directorio actual"""
        try:
            if os.path.exists("Factores.xlsx"):
                self.factores_df = pd.read_excel("Factores.xlsx")
                self.procesar_factores()
                self.excel_status.config(text="✅ Factores.xlsx cargado", foreground="green")
            else:
                self.excel_status.config(text="⚠️ Factores.xlsx no encontrado", foreground="orange")
        except Exception as e:
            self.excel_status.config(text=f"❌ Error: {str(e)[:30]}...", foreground="red")
    
    def cargar_archivo_excel(self):
        """Cargar archivo Excel de factores"""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo de factores",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if archivo:
            try:
                self.factores_df = pd.read_excel(archivo)
                self.procesar_factores()
                self.excel_status.config(text="✅ Archivo cargado correctamente", foreground="green")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{str(e)}")
                self.excel_status.config(text="❌ Error al cargar archivo", foreground="red")
    
    def procesar_factores(self):
        """Procesar datos del archivo Excel"""
        if self.factores_df is None:
            return
            
        # Renombrar columnas para facilitar el acceso
        self.factores_df.columns = ['tipo_resistencia', 'material', 'papel', 'codigo', 'factor']
        
        # Actualizar comboboxes
        self.actualizar_materiales()
    
    def actualizar_materiales(self):
        """Actualizar lista de materiales disponibles"""
        if self.factores_df is None:
            return
            
        tipo_actual = self.tipo_caja.get()
        
        # Mapear tipo de caja a tipo de resistencia en Excel
        tipo_mapping = {
            "Corrugado Sencillo": "CORRUGADO SENCILLO KRAFT",
            "Corrugado Doble": "DOBLE CORRUGADO KRAFT",
            "Medio Fondo Sencillo": "CORRUGADO SENCILLO KRAFT",
            "Medio Fondo Doble": "DOBLE CORRUGADO KRAFT",
            "Doble Armado Sencillo": "CORRUGADO SENCILLO KRAFT",
            "Doble Armado Doble": "DOBLE CORRUGADO KRAFT",
            "Doble Armado Medio Fondo Sencillo": "CORRUGADO SENCILLO KRAFT",
            "Doble Armado Medio Fondo Doble": "DOBLE CORRUGADO KRAFT"
        }
        
        tipo_resistencia = tipo_mapping.get(tipo_actual, "CORRUGADO SENCILLO KRAFT")
        
        # Filtrar materiales disponibles para este tipo
        materiales_disponibles = self.factores_df[
            self.factores_df['tipo_resistencia'].str.contains(tipo_resistencia.split()[0])
        ]['material'].unique()
        
        self.material_combo['values'] = list(materiales_disponibles)
        if len(materiales_disponibles) > 0:
            self.material_combo.set(materiales_disponibles[0])
            self.actualizar_resistencias()
    
    def actualizar_resistencias(self, event=None):
        """Actualizar lista de resistencias según material seleccionado"""
        if self.factores_df is None:
            return
            
        material = self.material_combo.get()
        tipo_actual = self.tipo_caja.get()
        
        # Filtrar resistencias
        filtro = (self.factores_df['material'] == material)
        if "Doble" in tipo_actual:
            filtro &= self.factores_df['tipo_resistencia'].str.contains("DOBLE")
        else:
            filtro &= ~self.factores_df['tipo_resistencia'].str.contains("DOBLE")
            
        resistencias = self.factores_df[filtro]['codigo'].unique()
        
        self.resistencia_combo['values'] = list(resistencias)
        if len(resistencias) > 0:
            self.resistencia_combo.set(resistencias[0])
    
    def on_tipo_change(self, event=None):
        """Manejar cambio de tipo de caja"""
        self.actualizar_materiales()
    
    def obtener_factor(self) -> float:
        """Obtener factor del Excel según selección actual"""
        if self.factores_df is None:
            return 15.0  # Factor por defecto
            
        material = self.material_combo.get()
        resistencia = self.resistencia_combo.get()
        
        filtro = (self.factores_df['material'] == material) & (self.factores_df['codigo'] == resistencia)
        resultado = self.factores_df[filtro]
        
        if len(resultado) > 0:
            return float(resultado.iloc[0]['factor'])
        else:
            return 15.0
    
    def calcular_dimensiones(self, largo: float, ancho: float, alto: float, ceja: float, tipo: str) -> Dict:
        """Calcular dimensiones según tipo de caja"""
        
        if tipo == "sencillo":
            # Corrugado sencillo
            largo_total = ((largo + 0.4) + (ancho + 0.5)) * 2 + ceja
            base_ancho = ancho + 0.5
            ajuste = 0.2 if int((base_ancho * 100) % 100 / 10) % 2 == 0 else 0.3
            aleton = round(((base_ancho + ajuste) / 2), 1)
            central = alto + 0.8
            ancho_total = aleton + central + aleton
            
        elif tipo == "doble":
            # Corrugado doble  
            largo_total = ((largo + 0.8) + (ancho + 1)) * 2 + ceja
            base_ancho = ancho + 1
            ajuste = 0.2 if int((base_ancho * 100) % 100 / 10) % 2 == 0 else 0.3
            aleton = round(((base_ancho + ajuste) / 2), 1)
            central = alto + 1.5
            ancho_total = aleton + central + aleton
            
        elif tipo in ["medio_fondo_sencillo", "doble_armado_medio_fondo_sencillo"]:
            # Medio fondo sencillo
            largo_total = ((largo + 0.4) + (ancho + 0.5)) * 2 + ceja if "doble_armado" not in tipo else ((largo + 0.4) + (ancho + 0.5)) + ceja
            base_ancho = ancho + 0.5
            ajuste = 0.2 if int((base_ancho * 100) % 100 / 10) % 2 == 0 else 0.3
            aleton = round(((base_ancho + ajuste) / 2), 1)
            central = alto + 0.8
            ancho_total = aleton + central  # Solo un aletón para medio fondo
            
        elif tipo in ["medio_fondo_doble", "doble_armado_medio_fondo_doble"]:
            # Medio fondo doble
            largo_total = ((largo + 0.8) + (ancho + 1)) * 2 + ceja if "doble_armado" not in tipo else ((largo + 0.8) + (ancho + 1)) + ceja
            base_ancho = ancho + 1
            ajuste = 0.2 if int((base_ancho * 100) % 100 / 10) % 2 == 0 else 0.3
            aleton = round(((base_ancho + ajuste) / 2), 1)
            central = alto + 1.5
            ancho_total = aleton + central  # Solo un aletón para medio fondo
            
        elif tipo == "doble_armado_sencillo":
            # Doble armado sencillo
            largo_total = ((largo + 0.4) + (ancho + 0.5)) + ceja
            base_ancho = ancho + 0.5
            ajuste = 0.2 if int((base_ancho * 100) % 100 / 10) % 2 == 0 else 0.3
            aleton = round(((base_ancho + ajuste) / 2), 1)
            central = alto + 0.8
            ancho_total = aleton + central + aleton
            
        elif tipo == "doble_armado_doble":
            # Doble armado doble
            largo_total = ((largo + 0.8) + (ancho + 1)) + ceja
            base_ancho = ancho + 1
            ajuste = 0.2 if int((base_ancho * 100) % 100 / 10) % 2 == 0 else 0.3
            aleton = round(((base_ancho + ajuste) / 2), 1)
            central = alto + 1.5
            ancho_total = aleton + central + aleton
            
        else:
            # Fallback a sencillo
            largo_total = ((largo + 0.4) + (ancho + 0.5)) * 2 + ceja
            base_ancho = ancho + 0.5
            ajuste = 0.2 if int((base_ancho * 100) % 100 / 10) % 2 == 0 else 0.3
            aleton = round(((base_ancho + ajuste) / 2), 1)
            central = alto + 0.8
            ancho_total = aleton + central + aleton
        
        area = largo_total * ancho_total
        
        # Para doble armado, el área se multiplica por 2
        if "doble_armado" in tipo:
            area *= 2
            
        return {
            'largo_total': largo_total,
            'ancho_total': ancho_total,
            'area': area,
            'aleton': aleton,
            'central': central
        }
    
    def calcular_todo(self):
        """Realizar todos los cálculos y mostrar resultados"""
        try:
            # Obtener valores de entrada
            largo = float(self.largo_var.get())
            ancho = float(self.ancho_var.get())
            alto = float(self.alto_var.get())
            ceja = float(self.ceja_var.get())
            utilidad = float(self.utilidad_var.get())
            descuento = float(self.descuento_var.get())
            
            tipo_nombre = self.tipo_caja.get()
            tipo_codigo = self.tipos_caja[tipo_nombre]
            
            # Obtener factor desde Excel
            factor = self.obtener_factor()
            
            # Calcular dimensiones
            dimensiones = self.calcular_dimensiones(largo, ancho, alto, ceja, tipo_codigo)
            
            # Cálculos financieros
            # El cálculo correcto se realiza sobre mil unidades de área
            # para mantener consistencia con la versión ejecutada desde
            # el código fuente (main.py) y la documentación.
            costo_materia = (dimensiones['area'] * factor) / 1000
            precio_venta = costo_materia * utilidad
            precio_venta_ajustado = precio_venta * (1 + descuento / 100)
            ganancia_total = precio_venta_ajustado - costo_materia
            compra_minima = (10000 / dimensiones['area']) * 1000
            
            # Mostrar resultados
            self.mostrar_resultados(
                tipo_nombre, dimensiones, factor, 
                costo_materia, precio_venta, precio_venta_ajustado, 
                ganancia_total, compra_minima
            )
            
        except ValueError as e:
            messagebox.showerror("Error", "Por favor ingrese valores numéricos válidos")
        except Exception as e:
            messagebox.showerror("Error", f"Error en el cálculo: {str(e)}")
    
    def mostrar_resultados(self, tipo, dim, factor, costo, precio_base, precio_ajustado, ganancia, compra_min):
        """Mostrar resultados en el área de texto"""
        self.result_text.delete(1.0, tk.END)
        
        # Calcular la descripción de aletones
        if 'medio_fondo' in tipo.lower():
            aletones_desc = f"{dim['aleton']} + {dim['central']}"
        else:
            aletones_desc = f"{dim['aleton']} + {dim['central']} + {dim['aleton']}"
        
        # Separadores
        sep_mayor = "=" * 60
        sep_menor = "-" * 40
        
        resultado = f"""
🏭 CALCULADORA DE CAJAS DE CARTÓN CORRUGADO
{sep_mayor}

📦 TIPO DE CAJA: {tipo}
📏 DIMENSIONES ENTRADA:
   • Largo: {self.largo_var.get()} cm
   • Ancho: {self.ancho_var.get()} cm  
   • Alto: {self.alto_var.get()} cm
   • Ceja: {self.ceja_var.get()} cm

📐 CÁLCULOS DIMENSIONALES:
   • Largo Total: {dim['largo_total']:.2f} cm
   • Ancho Total: {dim['ancho_total']:.2f} cm
   • Área Total: {dim['area']:.2f} cm²
   • Aletones: {aletones_desc}

🧪 MATERIAL SELECCIONADO:
   • Material: {self.material_combo.get()}
   • Resistencia: {self.resistencia_combo.get()}
   • Factor: {factor:.6f}

💰 CÁLCULOS FINANCIEROS:
   • Costo Materia Prima: ${costo:.6f}
   • Precio Venta (Base): ${precio_base:.2f}
   • Precio Venta Ajustado: ${precio_ajustado:.2f}
   • Ganancia Total: ${ganancia:.2f}
   • Compra Mínima: {compra_min:.0f} unidades

📊 PARÁMETROS USADOS:
   • Ganancia %: {self.utilidad_var.get()}%
   • Descuento/Incremento: {self.descuento_var.get()}%

{sep_mayor}
💡 Resultados generados: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}
        """
        
        self.result_text.insert(tk.END, resultado)
    
    def ejecutar(self):
        """Ejecutar la aplicación"""
        self.root.mainloop()

def main():
    """Función principal"""
    try:
        app = CalculadoraCajas()
        app.ejecutar()
    except Exception as e:
        print(f"Error al iniciar la aplicación: {e}")
        input("Presione Enter para salir...")

if __name__ == "__main__":
    main()
