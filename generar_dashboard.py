import numpy as np
import os
import json
from openpyxl import load_workbook

PERIODO_INFO = {
    "Horas Noviembre-Diciembre 2025.xlsx": {"label": "Nov-Dic 2025", "inicio": "2025-11-15", "fin": "2025-12-14", "orden": 1},
    "Horas Diciembre - Enero 2026.xlsx":   {"label": "Dic-Ene 2026", "inicio": "2025-12-15", "fin": "2026-01-14", "orden": 2},
    "Horas Enero -Febrero 2026.xlsx":      {"label": "Ene-Feb 2026", "inicio": "2026-01-15", "fin": "2026-02-14", "orden": 3},
    "Horas Febrero - Marzo 2026.xlsx":     {"label": "Feb-Mar 2026", "inicio": "2026-02-15", "fin": "2026-03-14", "orden": 4},
    "Horas Marzo -Abril 2026.xlsx":        {"label": "Mar-Abr 2026", "inicio": "2026-03-15", "fin": "2026-04-14", "orden": 5},
    "Horas Abril - Mayo 2026.xlsx":        {"label": "Abr-May 2026", "inicio": "2026-04-15", "fin": "2026-05-14", "orden": 6},
}

EMPLEADOS_VALIDOS = {
    "MADELINE SANCHEZ":  "SANCHEZ MACIAS MADELINE NATHALY",
    "GABRIEL BAQUERIZO": "BAQUERIZO PALACIOS GABRIEL NATHANAEL",
    "BRYAN RODRIGUEZ":   "RODRIGUEZ DIEZ BRYAN XAVIER",
    "Chistian Jimenes":  "JIMENEZ CRISTHIAN",
    "CRISTHIAN JIMENEZ": "JIMENEZ CRISTHIAN",
    "HUGO PARADA":       "PARADA SILVA HUGO ANDRES",
    "ISSAC GONZALEZ":    "GONZALEZ ROMO ISAAC ANDRES",
    "FERNANDO PHILCO":   "PHILCO BAQUE JORGE FERNANDO",
    "MELANIE REYES":     "REYES MALDONADO MELANIE GABRIELA",
    "JOSE SERRANO":      "SERRANO GUAMAN JOSE MIGUEL",
}

SUELDO_OVERRIDE = 660
FERIADOS_KEYWORDS = ["FERIADO", "LIBRE", "DIA LIBRE"]


def to_num(val):
    try:
        v = float(val)
        return v if not np.isnan(v) else None
    except (TypeError, ValueError):
        return None


def fecha_to_str(fecha):
    from datetime import datetime
    if fecha is None:
        return None
    if isinstance(fecha, datetime):
        return fecha.strftime("%Y-%m-%d")
    try:
        return str(fecha)[:10]
    except:
        return None


def es_feriado_texto(nota):
    if nota is None:
        return False
    return any(k in str(nota).upper() for k in FERIADOS_KEYWORDS)


def es_dia_libre_color(cell):
    try:
        fill = cell.fill
        if fill.patternType == "solid":
            fg = fill.fgColor
            if fg and fg.type == "theme" and fg.theme == 9:
                return True
    except:
        pass
    return False


def calcular_horas_nocturnas(ing, sal, th):
    """Calcula horas entre 19:00 (7PM) y 26:00 (2AM) - recargo nocturno 25%.
    En el Excel: 25 = 1AM, 26 = 2AM.
    Se cuenta el overlap directo del turno con la ventana [19, 26]."""
    if ing is None or sal is None or th is None or th <= 0:
        return 0
    jornada_std = 8
    regular_hours = min(th, jornada_std)

    if ing >= 14:
        # Turno nocturno: overlap directo con ventana [19, 26]
        night_hours = max(0, min(sal, 26) - max(ing, 19))
        return min(night_hours, regular_hours)
    else:
        # Turno diurno/vespertino: regular_end = ing + 9 (8h trabajo + 1h almuerzo)
        regular_end = min(ing + 9, sal)
        if regular_end > 19:
            night_hours = min(regular_end, 26) - 19
            return min(night_hours, regular_hours)
    return 0


def calcular_horas_extras(ing, sal, th, es_feriado, es_dia_libre):
    h25, h50, h100 = 0, 0, 0
    if th is None or th <= 0:
        return 0, 0, 0

    # H 100%: Feriado o Dia Libre - solo se pagan horas al 100%, sin H25
    if es_feriado or es_dia_libre:
        return 0, 0, th

    # Dia laboral regular
    if ing is None or sal is None:
        return 0, 0, 0

    # H 25%: recargo nocturno solo en dias regulares
    h25 = calcular_horas_nocturnas(ing, sal, th)

    # H 50%: Overtime en dia normal
    jornada_std = 8
    h50 = max(0, th - jornada_std)

    return h25, h50, h100


# === PASO 1: Escanear TODOS los archivos para detectar feriados nacionales ===
ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Info")
archivos = [f for f in os.listdir(ruta) if f.endswith(".xlsx") and not f.startswith("~$")]

feriados_nacionales = set()
print("Detectando feriados nacionales...")
for archivo in archivos:
    if archivo not in PERIODO_INFO:
        continue
    wb = load_workbook(os.path.join(ruta, archivo), data_only=True)
    for sn in wb.sheetnames:
        ws = wb[sn]
        for row in range(7, ws.max_row + 1):
            nota = ws.cell(row=row, column=7).value
            fecha = ws.cell(row=row, column=3).value
            if nota and "FERIADO" in str(nota).upper() and fecha:
                fs = fecha_to_str(fecha)
                if fs:
                    feriados_nacionales.add(fs)
    wb.close()
print(f"Feriados: {sorted(feriados_nacionales)}")

# === PASO 2: Procesar cada empleado con feriados globales + deteccion de color verde ===
registros = []

for archivo in sorted(archivos):
    if archivo not in PERIODO_INFO:
        continue

    info = PERIODO_INFO[archivo]
    filepath = os.path.join(ruta, archivo)
    print(f"Procesando: {archivo} -> {info['label']}")

    try:
        wb = load_workbook(filepath, data_only=True)
    except Exception as e:
        print(f"  Error abriendo: {e}")
        continue

    for sheet_name in wb.sheetnames:
        sn = sheet_name.strip()
        if sn not in EMPLEADOS_VALIDOS:
            continue

        nombre_completo = EMPLEADOS_VALIDOS[sn]
        ws = wb[sheet_name]

        sueldo = SUELDO_OVERRIDE
        base_hora = sueldo / 240

        total_horas = 0
        sum_h25, sum_h50, sum_h100 = 0, 0, 0

        for row in range(7, ws.max_row + 1):
            cell_dia = ws.cell(row=row, column=2)
            dia_val = cell_dia.value
            if dia_val is None:
                continue
            dia_str = str(dia_val).strip().upper()
            if "TOTAL" in dia_str:
                break

            fecha_val = ws.cell(row=row, column=3).value
            if fecha_val is None:
                continue

            ingreso = to_num(ws.cell(row=row, column=4).value)
            salida = to_num(ws.cell(row=row, column=5).value)
            th = to_num(ws.cell(row=row, column=6).value)
            nota = ws.cell(row=row, column=7).value
            h100_directo = to_num(ws.cell(row=row, column=9).value)

            if th is None or th <= 0:
                continue

            total_horas += th

            # Detectar FERIADO vs DIA LIBRE por separado
            fecha_str = fecha_to_str(fecha_val)
            nota_upper = str(nota).upper() if nota else ""

            # FERIADO: feriado nacional (H100 + H25 coexisten)
            es_feriado_flag = (
                "FERIADO" in nota_upper
                or (fecha_str in feriados_nacionales)
            )

            # DIA LIBRE: dia libre personal/green (solo H100, sin H25)
            es_dia_libre_flag = (
                es_dia_libre_color(cell_dia)
                or ("LIBRE" in nota_upper and "FERIADO" not in nota_upper)
                or (h100_directo is not None and h100_directo > 0 and not es_feriado_flag)
            )

            h25, h50, h100 = calcular_horas_extras(ing=ingreso, sal=salida, th=th, es_feriado=es_feriado_flag, es_dia_libre=es_dia_libre_flag)
            sum_h25 += h25
            sum_h50 += h50
            sum_h100 += h100

        v25 = round(sum_h25 * base_hora * 0.25, 2)
        v50 = round(sum_h50 * base_hora * 1.50, 2)
        v100 = round(sum_h100 * base_hora * 2.00, 2)
        total_cancelar = round(v25 + v50 + v100, 2)

        registros.append({
            "nombre": nombre_completo,
            "sueldo": sueldo,
            "periodo": info["label"],
            "orden": info["orden"],
            "fechaInicio": info["inicio"],
            "fechaFin": info["fin"],
            "totalHoras": round(total_horas, 2),
            "h25": round(sum_h25, 2),
            "h50": round(sum_h50, 2),
            "h100": round(sum_h100, 2),
            "v25": v25,
            "v50": v50,
            "v100": v100,
            "total": total_cancelar,
        })

    wb.close()

print(f"\nTotal registros: {len(registros)}")
print(f"Empleados: {sorted(set(r['nombre'] for r in registros))}")

for per in ['Dic-Ene 2026', 'Mar-Abr 2026']:
    print(f"\n=== Verificacion {per} ===")
    for r in sorted([x for x in registros if x['periodo'] == per], key=lambda x: x['nombre']):
        print(f"  {r['nombre']:45s} H25={r['h25']:5.0f} H50={r['h50']:5.0f} H100={r['h100']:5.0f} | V25={r['v25']:7.2f} V50={r['v50']:7.2f} V100={r['v100']:7.2f} | TOTAL={r['total']:8.2f}")

datos_json = json.dumps(registros, ensure_ascii=False)
orden_periodos = [v["label"] for v in sorted(PERIODO_INFO.values(), key=lambda x: x["orden"])]
periodos_json = json.dumps(orden_periodos)

html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Horas Extras</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif; background:#0f172a; color:#e2e8f0; }}
.header {{ background:linear-gradient(135deg,#1e293b,#334155); padding:20px 30px; display:flex; align-items:center; justify-content:space-between; box-shadow:0 4px 12px rgba(0,0,0,.3); }}
.header h1 {{ font-size:1.5rem; color:#38bdf8; }}
.header span {{ color:#94a3b8; font-size:.85rem; }}
.filters {{ background:#1e293b; padding:16px 30px; display:flex; gap:20px; flex-wrap:wrap; align-items:center; border-bottom:1px solid #334155; }}
.filters label {{ color:#94a3b8; font-size:.8rem; margin-right:4px; }}
.filters select, .filters input {{ background:#0f172a; color:#e2e8f0; border:1px solid #475569; border-radius:6px; padding:6px 10px; font-size:.85rem; }}
.filters select {{ min-width:180px; max-height:32px; }}
.kpis {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:16px; padding:20px 30px; }}
.kpi {{ background:linear-gradient(135deg,#1e293b,#263348); border-radius:12px; padding:20px; text-align:center; border:1px solid #334155; }}
.kpi .value {{ font-size:1.8rem; font-weight:700; color:#38bdf8; }}
.kpi .label {{ color:#94a3b8; font-size:.8rem; margin-top:4px; }}
.charts {{ padding:20px 30px; display:grid; gap:20px; }}
.chart-row {{ display:grid; grid-template-columns:1fr 1fr; gap:20px; }}
.chart-box {{ background:#1e293b; border-radius:12px; padding:20px; border:1px solid #334155; }}
.chart-box.full {{ grid-column:1/-1; }}
.chart-box h3 {{ color:#38bdf8; margin-bottom:12px; font-size:.95rem; }}
canvas {{ max-height:400px; }}
table {{ width:100%; border-collapse:collapse; font-size:.78rem; }}
th {{ background:#334155; color:#38bdf8; padding:8px; text-align:left; position:sticky; top:0; }}
td {{ padding:6px 8px; border-bottom:1px solid #1e293b; }}
tr:hover td {{ background:#334155; }}
.table-wrap {{ max-height:400px; overflow-y:auto; border-radius:8px; }}
.money {{ color:#4ade80; }}
@media(max-width:900px){{ .chart-row {{ grid-template-columns:1fr; }} }}
</style>
</head>
<body>

<div class="header">
  <h1>📊 Dashboard Consolidado Horas Extras</h1>
  <span>INT FOOD SERVICES CORP</span>
</div>

<div class="filters">
  <div>
    <label>Periodo Desde:</label>
    <select id="filtroDesde"></select>
  </div>
  <div>
    <label>Periodo Hasta:</label>
    <select id="filtroHasta"></select>
  </div>
  <div>
    <label>Empleado:</label>
    <select id="filtroEmpleado"><option value="">Todos</option></select>
  </div>
  <div>
    <button onclick="aplicarFiltros()" style="background:#38bdf8;color:#0f172a;border:none;padding:8px 20px;border-radius:6px;cursor:pointer;font-weight:600;">Filtrar</button>
    <button onclick="resetFiltros()" style="background:#475569;color:#e2e8f0;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;margin-left:6px;">Reset</button>
  </div>
</div>

<div class="kpis" id="kpis"></div>

<div class="charts">
  <div class="chart-row">
    <div class="chart-box"><h3>Horas Extras por Empleado (Consolidado)</h3><canvas id="chartHorasEmp"></canvas></div>
    <div class="chart-box"><h3>Total a Cancelar por Empleado</h3><canvas id="chartPagoEmp"></canvas></div>
  </div>
  <div class="chart-row">
    <div class="chart-box"><h3>Horas Extras por Periodo</h3><canvas id="chartHorasPeriodo"></canvas></div>
    <div class="chart-box"><h3>Gasto Total por Periodo</h3><canvas id="chartGastoPeriodo"></canvas></div>
  </div>
  <div class="chart-box full"><h3>Detalle por Empleado y Periodo (Valores Pagados)</h3><canvas id="chartValoresDetalle"></canvas></div>
  <div class="chart-box full">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
      <h3>Tabla Detallada</h3>
      <button onclick="exportarExcel()" style="background:#10b981;color:#fff;border:none;padding:8px 20px;border-radius:6px;cursor:pointer;font-weight:600;display:flex;align-items:center;gap:6px;">
        📥 Exportar Excel
      </button>
    </div>
    <div class="table-wrap" id="tablaWrap"></div>
  </div>
</div>

<script>
const DATA = {datos_json};
const PERIODOS_ORDEN = {periodos_json};
let datosFiltrados = [...DATA];
let charts = {{}};

function init() {{
  const sel1 = document.getElementById('filtroDesde');
  const sel2 = document.getElementById('filtroHasta');
  PERIODOS_ORDEN.forEach((p,i) => {{
    sel1.innerHTML += `<option value="${{i}}">${{p}}</option>`;
    sel2.innerHTML += `<option value="${{i}}" ${{i===PERIODOS_ORDEN.length-1?'selected':''}}>${{p}}</option>`;
  }});
  poblarEmpleados();
  aplicarFiltros();
}}

function poblarEmpleados() {{
  const sel = document.getElementById('filtroEmpleado');
  const nombres = [...new Set(DATA.map(d=>d.nombre))].sort();
  sel.innerHTML = '<option value="">Todos</option>';
  nombres.forEach(n => sel.innerHTML += `<option value="${{n}}">${{n}}</option>`);
}}

function resetFiltros() {{
  document.getElementById('filtroDesde').value = '0';
  document.getElementById('filtroHasta').value = PERIODOS_ORDEN.length - 1;
  document.getElementById('filtroEmpleado').value = '';
  aplicarFiltros();
}}

function aplicarFiltros() {{
  const desde = parseInt(document.getElementById('filtroDesde').value);
  const hasta = parseInt(document.getElementById('filtroHasta').value);
  const emp = document.getElementById('filtroEmpleado').value;
  const periodosActivos = PERIODOS_ORDEN.slice(desde, hasta + 1);

  datosFiltrados = DATA.filter(d => {{
    if (!periodosActivos.includes(d.periodo)) return false;
    if (emp && d.nombre !== emp) return false;
    return true;
  }});

  renderKPIs();
  renderCharts(periodosActivos);
  renderTabla();
}}

function renderKPIs() {{
  const totalEmp = new Set(datosFiltrados.map(d=>d.nombre)).size;
  const totalH = datosFiltrados.reduce((s,d)=>s+d.totalHoras,0);
  const totalPago = datosFiltrados.reduce((s,d)=>s+d.total,0);
  const totalPeriodos = new Set(datosFiltrados.map(d=>d.periodo)).size;
  document.getElementById('kpis').innerHTML = `
    <div class="kpi"><div class="value">${{totalEmp}}</div><div class="label">Empleados</div></div>
    <div class="kpi"><div class="value">${{totalH.toFixed(1)}}</div><div class="label">Total Horas Extras</div></div>
    <div class="kpi"><div class="value money">${{totalPago.toLocaleString('es-EC',{{style:'currency',currency:'USD'}})}}</div><div class="label">Total a Cancelar</div></div>
    <div class="kpi"><div class="value">${{totalPeriodos}}</div><div class="label">Periodos</div></div>
  `;
}}

function agrupar(datos, campo, valCampos) {{
  const map = {{}};
  datos.forEach(d => {{
    const k = d[campo];
    if (!map[k]) map[k] = {{}};
    valCampos.forEach(v => {{ map[k][v] = (map[k][v]||0) + d[v]; }});
  }});
  return map;
}}

function destroyChart(name) {{ if(charts[name]) {{ charts[name].destroy(); }} }}

function renderCharts(periodosActivos) {{
  // 1. Horas por empleado consolidado (stacked bar)
  const horasEmp = agrupar(datosFiltrados, 'nombre', ['h25','h50','h100','totalHoras']);
  const nombresH = Object.keys(horasEmp).sort((a,b) => horasEmp[b].totalHoras - horasEmp[a].totalHoras);
  destroyChart('chartHorasEmp');
  charts.chartHorasEmp = new Chart(document.getElementById('chartHorasEmp'), {{
    type:'bar',
    data:{{
      labels: nombresH,
      datasets:[
        {{ label:'H 25%', data:nombresH.map(n=>horasEmp[n].h25), backgroundColor:'#3b82f6' }},
        {{ label:'H 50%', data:nombresH.map(n=>horasEmp[n].h50), backgroundColor:'#f59e0b' }},
        {{ label:'H 100%', data:nombresH.map(n=>horasEmp[n].h100), backgroundColor:'#ef4444' }},
      ]
    }},
    options:{{ responsive:true, plugins:{{ legend:{{ labels:{{ color:'#94a3b8' }} }} }}, scales:{{ x:{{ stacked:true, ticks:{{ color:'#94a3b8', maxRotation:45, font:{{size:9}} }} }}, y:{{ stacked:true, ticks:{{ color:'#94a3b8' }}, grid:{{ color:'#1e293b' }} }} }} }}
  }});

  // 2. Total a cancelar por empleado
  const pagoEmp = agrupar(datosFiltrados, 'nombre', ['total']);
  const nombresP = Object.keys(pagoEmp).sort((a,b) => pagoEmp[b].total - pagoEmp[a].total);
  destroyChart('chartPagoEmp');
  charts.chartPagoEmp = new Chart(document.getElementById('chartPagoEmp'), {{
    type:'bar',
    data:{{
      labels: nombresP,
      datasets:[{{ label:'Total $', data:nombresP.map(n=>pagoEmp[n].total), backgroundColor:'#4ade80' }}]
    }},
    options:{{ indexAxis:'x', responsive:true, plugins:{{ legend:{{ labels:{{ color:'#94a3b8' }} }} }}, scales:{{ x:{{ ticks:{{ color:'#94a3b8', maxRotation:45, font:{{size:9}} }} }}, y:{{ ticks:{{ color:'#94a3b8', callback:v=>'$'+v.toFixed(0) }}, grid:{{ color:'#1e293b' }} }} }} }}
  }});

  // 3. Horas por periodo
  const horasPer = {{}};
  periodosActivos.forEach(p => {{ horasPer[p] = {{ h25:0, h50:0, h100:0 }}; }});
  datosFiltrados.forEach(d => {{
    if (horasPer[d.periodo]) {{
      horasPer[d.periodo].h25 += d.h25;
      horasPer[d.periodo].h50 += d.h50;
      horasPer[d.periodo].h100 += d.h100;
    }}
  }});
  destroyChart('chartHorasPeriodo');
  charts.chartHorasPeriodo = new Chart(document.getElementById('chartHorasPeriodo'), {{
    type:'bar',
    data:{{
      labels: periodosActivos,
      datasets:[
        {{ label:'H 25%', data:periodosActivos.map(p=>horasPer[p].h25), backgroundColor:'#3b82f6' }},
        {{ label:'H 50%', data:periodosActivos.map(p=>horasPer[p].h50), backgroundColor:'#f59e0b' }},
        {{ label:'H 100%', data:periodosActivos.map(p=>horasPer[p].h100), backgroundColor:'#ef4444' }},
      ]
    }},
    options:{{ responsive:true, plugins:{{ legend:{{ labels:{{ color:'#94a3b8' }} }} }}, scales:{{ x:{{ stacked:true, ticks:{{ color:'#94a3b8' }} }}, y:{{ stacked:true, ticks:{{ color:'#94a3b8' }}, grid:{{ color:'#1e293b' }} }} }} }}
  }});

  // 4. Gasto por periodo (line)
  const gastoPer = {{}};
  periodosActivos.forEach(p => {{ gastoPer[p] = 0; }});
  datosFiltrados.forEach(d => {{ if(gastoPer[d.periodo]!==undefined) gastoPer[d.periodo] += d.total; }});
  destroyChart('chartGastoPeriodo');
  charts.chartGastoPeriodo = new Chart(document.getElementById('chartGastoPeriodo'), {{
    type:'line',
    data:{{
      labels: periodosActivos,
      datasets:[{{ label:'Gasto Total $', data:periodosActivos.map(p=>gastoPer[p]), borderColor:'#f43f5e', backgroundColor:'rgba(244,63,94,.15)', fill:true, tension:.3, pointRadius:6, pointBackgroundColor:'#f43f5e' }}]
    }},
    options:{{ responsive:true, plugins:{{ legend:{{ labels:{{ color:'#94a3b8' }} }} }}, scales:{{ x:{{ ticks:{{ color:'#94a3b8' }} }}, y:{{ ticks:{{ color:'#94a3b8', callback:v=>'$'+v.toFixed(0) }}, grid:{{ color:'#1e293b' }} }} }} }}
  }});

  // 5. Valores pagados detalle por empleado y periodo (grouped)
  const nombres = [...new Set(datosFiltrados.map(d=>d.nombre))].sort();
  const colores = ['#3b82f6','#f59e0b','#ef4444','#4ade80','#a78bfa','#f472b6'];
  const datasets = periodosActivos.map((p,i) => {{
    const vals = nombres.map(n => {{
      const match = datosFiltrados.filter(d=>d.nombre===n && d.periodo===p);
      return match.reduce((s,d)=>s+d.total,0);
    }});
    return {{ label:p, data:vals, backgroundColor:colores[i%colores.length] }};
  }});
  destroyChart('chartValoresDetalle');
  charts.chartValoresDetalle = new Chart(document.getElementById('chartValoresDetalle'), {{
    type:'bar',
    data:{{ labels:nombres, datasets }},
    options:{{ responsive:true, plugins:{{ legend:{{ labels:{{ color:'#94a3b8' }} }} }}, scales:{{ x:{{ ticks:{{ color:'#94a3b8', maxRotation:45, font:{{size:9}} }} }}, y:{{ ticks:{{ color:'#94a3b8', callback:v=>'$'+v.toFixed(0) }}, grid:{{ color:'#1e293b' }} }} }} }}
  }});
}}

function renderTabla() {{
  let rows = datosFiltrados.sort((a,b) => a.periodo.localeCompare(b.periodo) || a.nombre.localeCompare(b.nombre));
  let html = '<table><thead><tr><th>Periodo</th><th>Cédula</th><th>Nombre</th><th>Cargo</th><th>H 25%</th><th>H 50%</th><th>H 100%</th><th>Total H</th><th>V 25%</th><th>V 50%</th><th>V 100%</th><th>Total $</th></tr></thead><tbody>';
  rows.forEach(d => {{
    html += `<tr>
      <td>${{d.periodo}}</td><td>${{d.cedula}}</td><td>${{d.nombre}}</td><td>${{d.cargo}}</td>
      <td>${{d.h25}}</td><td>${{d.h50}}</td><td>${{d.h100}}</td><td>${{d.totalHoras}}</td>
      <td class="money">${{d.v25.toFixed(2)}}</td><td class="money">${{d.v50.toFixed(2)}}</td><td class="money">${{d.v100.toFixed(2)}}</td>
      <td class="money" style="font-weight:700">${{d.total.toFixed(2)}}</td>
    </tr>`;
  }});
  html += '</tbody></table>';
  document.getElementById('tablaWrap').innerHTML = html;
}}

function exportarExcel() {{
  const data = datosFiltrados.map(d => ({{
    'Periodo': d.periodo,
    'Cédula': d.cedula || '',
    'Nombre': d.nombre,
    'Cargo': d.cargo || '',
    'H 25%': d.h25,
    'H 50%': d.h50,
    'H 100%': d.h100,
    'Total Horas': d.totalHoras,
    'V 25%': d.v25,
    'V 50%': d.v50,
    'V 100%': d.v100,
    'Total $': d.total
  }}));
  
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Horas Extras');
  
  const filename = 'Horas_Extras_' + new Date().toISOString().slice(0,10) + '.xlsx';
  XLSX.writeFile(wb, filename);
}}

init();
</script>
</body>
</html>"""

output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html")
with open(output_path, "w", encoding="utf-8") as f:
    f.write(html)

print(f"Dashboard generado exitosamente en: {output_path}")
print(f"Total registros: {len(registros)}")
print(f"Empleados unicos: {len(set(r['nombre'] for r in registros))}")
print(f"Periodos: {len(set(r['periodo'] for r in registros))}")
