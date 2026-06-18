import pandas as pd

excel_file = "data/Engine Programme Management.xlsx"
event = "TC2K-2026-04"

LIMITS = {
    "Presión aceite mín FL": {
        "warning": 3.0,
        "critical": 2.5,
        "direction": "low",
    },
    "Temp. aceite máx FL": {
        "warning": 125,
        "critical": 135,
        "direction": "high",
    },
    "Temp. agua máx FL": {
        "warning": 100,
        "critical": 110,
        "direction": "high",
    },
    "Temp. máxima detectada agua": {
        "warning": 100,
        "critical": 115,
        "direction": "high",
    },
    "RPM turbo máx FL": {
        "warning": 165000,
        "critical": 175000,
        "direction": "high",
    },
    "Presión combustible mín FL": {
        "warning": 3.0,
        "critical": 2.5,
        "direction": "low",
    },
    "Temp. escape máx": {
        "warning": 990,
        "critical": 1100,
        "direction": "high",
    },
    "Voltaje batería mín FL": {
        "warning": 12.5,
        "critical": 11.5,
        "direction": "low",
    },
}

def generar_observaciones(row):
    observaciones = []
    prioridad = "🟢 OK"

    for columna, reglas in LIMITS.items():
        valor = row[columna]

        if pd.isna(valor):
            continue

        warning = reglas["warning"]
        critical = reglas["critical"]
        direction = reglas["direction"]

        if direction == "high":
            if valor >= critical:
                observaciones.append(f"🔴 {columna} fuera de rango ({valor})")
                prioridad = "🔴 CRÍTICO"
            elif valor >= warning:
                observaciones.append(f"🟡 {columna} cerca del límite ({valor})")
                if prioridad != "🔴 CRÍTICO":
                    prioridad = "🟡 ATENCIÓN"

        elif direction == "low":
            if valor <= critical:
                observaciones.append(f"🔴 {columna} fuera de rango ({valor})")
                prioridad = "🔴 CRÍTICO"
            elif valor <= warning:
                observaciones.append(f"🟡 {columna} cerca del límite ({valor})")
                if prioridad != "🔴 CRÍTICO":
                    prioridad = "🟡 ATENCIÓN"

    if not observaciones:
        observaciones.append("✅ Sin observaciones")

    return prioridad, " | ".join(observaciones)

df = pd.read_excel(excel_file, sheet_name="radb")

df = df[df["Carrera"] == event]

columns = [
    "Carrera",
    "Run name",
    "Motor",
    "Piloto",
    "Auto",
    "pOil_FL_min",
    "tOil_FL_max",
    "tWater_FL_max",
    "tWater_max",
    "nTurbo_FL_max",
    "pFuel_FL_min",
    "tAir_FL_max",
    "tExhaust_max",
    "VBatt_FL_min",
]

report = df[columns]

print(report.head(20))

output_file = "reports/Weekend Analyzer.xlsx"

report = report.rename(
    columns={
        "Carrera": "Evento",
        "Run name": "Run",
        "pOil_FL_min": "Presión aceite mín FL",
        "tOil_FL_max": "Temp. aceite máx FL",
        "tWater_FL_max": "Temp. agua máx FL",
        "tWater_max": "Temp. máxima detectada agua",
        "nTurbo_FL_max": "RPM turbo máx FL",
        "pFuel_FL_min": "Presión combustible mín FL",
        "tAir_FL_max": "Temp. aire máx FL",
        "tExhaust_max": "Temp. escape máx",
        "VBatt_FL_min": "Voltaje batería mín FL",
    }
)
for column in LIMITS.keys():
    report[column] = pd.to_numeric(report[column], errors="coerce")

report[["Prioridad", "Observaciones"]] = report.apply(
    generar_observaciones,
    axis=1,
    result_type="expand"
)

report.to_excel(output_file, index=False)

print("✅ Weekend Analyzer generado correctamente.")

