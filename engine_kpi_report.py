import pandas as pd

excel_file = "data/Engine Programme Management.xlsx"

df = pd.read_excel(
    excel_file,
    sheet_name="radb"
)

df["Engine_usage_max"] = pd.to_numeric(
    df["Engine_usage_max"],
    errors="coerce"
)

df = df.dropna(subset=["Engine_usage_max"])

def convert_to_km(row):
    motor = str(row["Motor"])
    engine_usage = row["Engine_usage_max"]

    if motor.startswith("5CT"):
        return engine_usage / 335000
    elif motor.startswith("OMC"):
        return engine_usage / 268000
    else:
        return None

df["km"] = df.apply(convert_to_km, axis=1)
df = df.dropna(subset=["km"])

kpi = df.groupby("Motor", as_index=False).agg(
    km_desde_ultimo_service=("km", "sum")
)

kpi["km_restantes"] = 2500 - kpi["km_desde_ultimo_service"]

kpi["km_desde_ultimo_service"] = (
    kpi["km_desde_ultimo_service"].round().astype(int)
)

kpi["km_restantes"] = (
    kpi["km_restantes"].round().astype(int)
)

def semaforo(km):
    if km >= 2500:
        return "🔴 SERVICE VENCIDO"
    elif km >= 2250:
        return "🟡 PRÓXIMO SERVICE"
    else:
        return "🟢 OK"

kpi["semaforo"] = kpi["km_desde_ultimo_service"].apply(semaforo)

services = df[df["Engine_usage_max"] < 0].groupby("Motor").size().reset_index(name="services_hechos")

kpi = kpi.merge(services, on="Motor", how="left")
kpi["services_hechos"] = kpi["services_hechos"].fillna(0).astype(int)


curr_state = pd.read_excel(
    excel_file,
    sheet_name="currState"
)

curr_state = curr_state[
    ["Motor", "Piloto", "Equipo"]
]

kpi = kpi.merge(
    curr_state,
    on="Motor",
    how="left"
)

services = df[df["Engine_usage_max"] < 0]

services_check = df[df["Engine_usage_max"] < 0]


from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

def prioridad(semaforo):
    if "VENCIDO" in semaforo:
        return 1
    elif "PRÓXIMO" in semaforo:
        return 2
    else:
        return 3

kpi["prioridad"] = kpi["semaforo"].apply(prioridad)

kpi = kpi.sort_values(
    by=["prioridad", "km_restantes"],
    ascending=[True, True]
)

kpi = kpi.drop(columns=["prioridad"])

output_file = "reports/engine_kpi_report.xlsx"

kpi.to_excel(output_file, index=False)

output_file = "reports/engine_kpi_report.xlsx"

summary = pd.DataFrame({
    "Indicador": [
        "Motores monitoreados",
        "Service vencido",
        "Próximo service",
        "OK",
        "Motor con mayor uso",
        "Próximo motor a service"
    ],
    "Valor": [
        len(kpi),
        (kpi["semaforo"] == "🔴 SERVICE VENCIDO").sum(),
        (kpi["semaforo"] == "🟡 PRÓXIMO SERVICE").sum(),
        (kpi["semaforo"] == "🟢 OK").sum(),
        f'{kpi.loc[kpi["km_desde_ultimo_service"].idxmax(), "Motor"]} - {kpi["km_desde_ultimo_service"].max()} km',
        f'{kpi[kpi["km_restantes"] > 0].sort_values("km_restantes").iloc[0]["Motor"]} - {kpi[kpi["km_restantes"] > 0]["km_restantes"].min()} km restantes'
    ]
})

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="Resumen Ejecutivo", index=False)
    kpi.to_excel(writer, sheet_name="Engine KPI Report", index=False)


wb = load_workbook(output_file)
ws = wb["Engine KPI Report"]

# Congelar la fila de encabezados
ws.freeze_panes = "A2"

# Agregar filtros automáticos
ws.auto_filter.ref = ws.dimensions


for cell in ws[1]:
    cell.font = Font(bold=True)

for column_cells in ws.columns:
    max_length = max(
        len(str(cell.value)) if cell.value is not None else 0
        for cell in column_cells
    )
    column_letter = get_column_letter(column_cells[0].column)
    ws.column_dimensions[column_letter].width = max_length + 2

green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

# Buscar la columna "semaforo"
semaforo_column = None

for cell in ws[1]:
    if cell.value == "semaforo":
        semaforo_column = cell.column
        break

# Pintar la celda de semáforo según el texto
for row in range(2, ws.max_row + 1):
    cell = ws.cell(row=row, column=semaforo_column)

# Buscar la columna "semaforo"
semaforo_column = None

for cell in ws[1]:
    if cell.value == "semaforo":
        semaforo_column = cell.column
        break

kpi = kpi.rename(
    columns={
        "km_desde_ultimo_service": "Km desde último Service",
        "km_restantes": "Km restantes",
        "services_hechos": "Services realizados",
        "semaforo": "Estado",
    }
)

kpi = kpi[
    [
        "Motor",
        "Piloto",
        "Equipo",
        "Km desde último Service",
        "Km restantes",
        "Services realizados",
        "Estado",
    ]
]

# Pintar solo la celda del semáforo según el estado
for row in range(2, ws.max_row + 1):
    semaforo_cell = ws.cell(row=row, column=semaforo_column)

    if "OK" in str(semaforo_cell.value):
        semaforo_cell.fill = green_fill
    elif "PRÓXIMO" in str(semaforo_cell.value):
        semaforo_cell.fill = yellow_fill
    elif "VENCIDO" in str(semaforo_cell.value):
        semaforo_cell.fill = red_fill


kpi = kpi.rename(
    columns={
        "km_desde_ultimo_service": "Km desde último Service",
        "km_restantes": "Km restantes",
        "services_hechos": "Services realizados",
        "semaforo": "Estado",
    }
)

wb.save(output_file)

print("✅ Reporte generado correctamente con formato.")

