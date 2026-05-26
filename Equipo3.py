"""
Analizador financiero para 35 proyectos.

Lee un archivo Excel llamado Equipo##.xlsx (busca Equipo*.xlsx en el directorio
de trabajo), lee la pestaña `dataset` con las columnas:
ID_Proyecto, Nombre_Proyecto, Inversion_Inicial, Tasa_Descuento, Flujo_1..Flujo_5

Calcula VAN y TIR por fila, etiqueta como "Rentable" si VAN>0 y TIR>Tasa_Descuento,
y escribe los resultados en una nueva pestaña llamada `Resultados` sin modificar
la pestaña original.

Uso: ejecutar directamente el script. Requiere pandas, numpy y openpyxl.
"""

from pathlib import Path
import sys
import math

try:
	import pandas as pd
except Exception as e:
	print("Falta la librería 'pandas'. Instálala con: pip install pandas openpyxl numpy")
	raise

try:
	import numpy as np
except Exception as e:
	print("Falta la librería 'numpy'. Instálala con: pip install numpy")
	raise

from openpyxl import load_workbook


def npv_cashflows(cashflows, discount_rate):
	# cashflows: iterable of flows at t=0..n (CF0 may be negative for outflow)
	r = discount_rate
	npv = 0.0
	for t, cf in enumerate(cashflows, start=0):
		npv += cf / ((1 + r) ** t)
	return npv


def irr_from_cashflows(cashflows):
	# cashflows: list of CF0..CFn (CF0 may be negative)
	cfs = list(cashflows)
	# Use numpy.roots on polynomial whose coefficients are cfs
	# Polynomial p(x) = C0*x^n + C1*x^{n-1} + ... + Cn
	try:
		coeffs = np.array(cfs, dtype=float)
		# If all zeros after C0, can't compute
		if np.allclose(coeffs[1:], 0):
			return float('nan')
		roots = np.roots(coeffs)
		real_roots = [r.real for r in roots if abs(r.imag) < 1e-6 and r.real != 0]
		# convert x = 1+irr  => irr = x-1; only keep irr > -0.9999
		irr_candidates = [x - 1 for x in real_roots if x - 1 > -0.9999]
		if not irr_candidates:
			return float('nan')
		# If multiple, choose the one that makes NPV closest to zero
		best = None
		best_err = None
		for irr in irr_candidates:
			npv = sum(cf / ((1 + irr) ** t) for t, cf in enumerate(cfs, start=0))
			err = abs(npv)
			if best is None or err < best_err:
				best = irr
				best_err = err
		return best
	except Exception:
		return float('nan')


def find_excel_file(directory: Path):
	# Busca archivos que empiecen por Equipo y terminen en .xlsx
	files = sorted(directory.glob('Equipo*.xlsx'))
	return files[0] if files else None


def main(path: str = None):
	cwd = Path(path or Path.cwd())
	excel_file = None
	if path and Path(path).is_file():
		excel_file = Path(path)
	else:
		excel_file = find_excel_file(cwd)

	if not excel_file:
		print('No se encontró ningún archivo Equipo##.xlsx en el directorio:', cwd)
		sys.exit(1)

	print('Usando archivo:', excel_file)

	# Leer la hoja dataset
	try:
		df = pd.read_excel(excel_file, sheet_name='dataset')
	except Exception as e:
		print("Error leyendo la hoja 'dataset':", e)
		sys.exit(1)

	expected_cols = ['ID_Proyecto', 'Nombre_Proyecto', 'Inversion_Inicial', 'Tasa_Descuento']
	flujo_cols = [f'Flujo_{i}' for i in range(1, 6)]
	missing = [c for c in expected_cols + flujo_cols if c not in df.columns]
	if missing:
		print('Faltan columnas esperadas en la hoja dataset:', missing)
		sys.exit(1)

	results = df.copy()
	vans = []
	tirs = []
	etiquetas = []

	for idx, row in df.iterrows():
		inv = row['Inversion_Inicial']
		tasa_desc = row['Tasa_Descuento']
		cashflows_years = [row[col] for col in flujo_cols]
		# Reemplazar NaN por 0 y convertir a float
		cashflows_years = [0.0 if (pd.isna(x)) else float(x) for x in cashflows_years]

		# Normalizar CF0: si Inversion_Inicial es negativa ya es un flujo (ej. -1000),
		# si es positiva asumimos inversión y lo convertimos a -inv
		try:
			inv_val = float(inv)
		except Exception:
			inv_val = 0.0

		if inv_val > 0:
			cf0 = -inv_val
		else:
			cf0 = inv_val

		# Normalizar tasa de descuento: si está en porcentajes (>1) convertir a decimal
		try:
			tasa_val = float(tasa_desc)
		except Exception:
			tasa_val = 0.0
		if abs(tasa_val) > 1:
			tasa_val = tasa_val / 100.0

		# Cashflows completos incluyendo t=0
		cfs = [cf0] + cashflows_years
		try:
			van = npv_cashflows(cfs, tasa_val)
		except Exception:
			van = float('nan')
		tir = irr_from_cashflows(cfs)

		# Etiquetado: Rentable si VAN>0 y TIR>tasa_desc
		rentable = False
		try:
			if not math.isnan(van) and van > 0 and not math.isnan(tir) and tir > tasa_val:
				rentable = True
		except Exception:
			rentable = False

		etiqueta = 'Rentable' if rentable else 'No rentable'

		vans.append(van)
		tirs.append(tir)
		etiquetas.append(etiqueta)

	results['VAN'] = vans
	results['TIR'] = tirs
	results['Clasificacion'] = etiquetas

	# Guardar en nueva pestaña Resultados sin modificar la original
	try:
		# Use pandas ExcelWriter with if_sheet_exists='replace' to update or create the sheet
		with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
			results.to_excel(writer, sheet_name='Resultados', index=False)
	except Exception as e:
		print('Error guardando los resultados en la hoja Resultados:', e)
		sys.exit(1)

	print('Análisis completado. Resultados guardados en la pestaña "Resultados" del archivo:', excel_file)


if __name__ == '__main__':
	# Permite pasar la ruta al archivo como argumento opcional
	arg = sys.argv[1] if len(sys.argv) > 1 else None
	main(arg)

