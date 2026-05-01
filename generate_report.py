"""
Generador de reporte Word (.docx) para el taller de Crashing.

Lee el mismo archivo de Excel que `cpm_crashing.py`, ejecuta el análisis y
escribe un documento Word con todos los resultados, dejando marcadores donde
deben pegarse las capturas de pantalla de POMQM.

Uso:
    python generate_report.py CrashingWorkshop.xlsx [salida.docx]
"""

from __future__ import annotations
import sys
from pathlib import Path
from datetime import date

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor, Cm

from cpm_crashing import (
    compute_analysis,
    cpm_pass,
    fmt_num,
    path_duration,
    total_path_slack,
)


SCREENSHOT_COLOR = RGBColor(0x1F, 0x4E, 0x79)  # azul oscuro
HIGHLIGHT_COLOR = RGBColor(0xC0, 0x50, 0x4D)   # rojo apagado


def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    return h


def add_paragraph(doc, text, bold=False, italic=False, size=None, color=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    return p


def add_screenshot_placeholder(doc, descripcion):
    """Caja visible que indica dónde pegar una captura de POMQM."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"[ INSERTAR CAPTURA DE POMQM AQUÍ — {descripcion} ]")
    run.bold = True
    run.italic = True
    run.font.size = Pt(11)
    run.font.color.rgb = SCREENSHOT_COLOR
    # marco visible: agregar un párrafo en blanco para dejar espacio
    doc.add_paragraph("\n\n\n")


def add_table_from_rows(doc, headers, rows, header_color=None, mark_rows=None):
    """Agrega una tabla con encabezados en negrita.

    mark_rows: iterable opcional con índices de filas que deben resaltarse
    (por ejemplo, ruta crítica o actividades críticas).
    """
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.size = Pt(10)
        hdr[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    mark_set = set(mark_rows or [])
    for i, row in enumerate(rows, start=1):
        cells = table.rows[i].cells
        for j, val in enumerate(row):
            cells[j].text = ""
            p = cells[j].paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(10)
            if (i - 1) in mark_set:
                run.bold = True
                run.font.color.rgb = HIGHLIGHT_COLOR
    return table


def write_input_table(doc, activities, successors):
    add_heading(doc, "Datos de entrada del proyecto", level=2)

    preds_of = {a: [] for a in activities}
    for s, succs in successors.items():
        for t in succs:
            preds_of.setdefault(t, []).append(s)

    headers = [
        "Actividad",
        "Predecesores",
        "Tiempo Normal (días)",
        "Costo Normal",
        "Tiempo Intensivo (días)",
        "Costo Intensivo",
    ]
    rows = []
    for a in sorted(activities):
        info = activities[a]
        preds = ", ".join(sorted(preds_of.get(a, []))) or "-"
        rows.append([
            a, preds,
            fmt_num(info["tiempo_normal"]),
            fmt_num(info["costo_normal"]),
            fmt_num(info["tiempo_intensivo"]),
            fmt_num(info["costo_intensivo"]),
        ])
    add_table_from_rows(doc, headers, rows)
    doc.add_paragraph()


def write_paths_table(doc, paths, time_dict, titulo):
    add_heading(doc, titulo, level=3)

    durations = [path_duration(p, time_dict) for p in paths]
    critical = max(durations)
    indexed = sorted(
        enumerate(zip(durations, paths), start=1), key=lambda x: -x[1][0]
    )

    headers = ["#", "Duración del Camino", "Holgura del Camino", "Camino"]
    rows = []
    mark_rows = []
    for i, (_, (dur, p)) in enumerate(indexed):
        slack = critical - dur
        if abs(slack) < 1e-9:
            mark_rows.append(i)
        path_str = " → ".join(f"{a}({fmt_num(time_dict[a])})" for a in p)
        rows.append([str(i + 1), fmt_num(dur), fmt_num(slack), path_str])
    add_table_from_rows(doc, headers, rows, mark_rows=mark_rows)

    crit_paths = [p for d, p in zip(durations, paths) if abs(d - critical) < 1e-9]
    crit_str = "; ".join(" → ".join(p) for p in crit_paths)
    add_paragraph(
        doc,
        f"Duración de la ruta crítica: {fmt_num(critical)} días.   "
        f"Ruta(s) crítica(s): {crit_str}",
        bold=True,
    )
    doc.add_paragraph()
    return critical


def write_activity_slack_table(doc, activities, successors, time_dict, titulo):
    add_heading(doc, titulo, level=3)

    IC, TC, IL, TL, holgura, project_duration = cpm_pass(
        activities, successors, time_dict
    )
    headers = [
        "Actividad",
        "Duración",
        "Inicio Cercano",
        "Término Cercano",
        "Inicio Lejano",
        "Término Lejano",
        "Holgura Total",
        "¿Crítica?",
    ]
    rows = []
    mark_rows = []
    for i, a in enumerate(sorted(activities)):
        es_critica = abs(holgura[a]) < 1e-9
        if es_critica:
            mark_rows.append(i)
        rows.append([
            a,
            fmt_num(time_dict[a]),
            fmt_num(IC[a]),
            fmt_num(TC[a]),
            fmt_num(IL[a]),
            fmt_num(TL[a]),
            fmt_num(holgura[a]),
            "SÍ" if es_critica else "No",
        ])
    add_table_from_rows(doc, headers, rows, mark_rows=mark_rows)
    add_paragraph(
        doc,
        f"Duración total del proyecto: {fmt_num(project_duration)} días.",
        bold=True,
    )
    doc.add_paragraph()


def write_desintensification_table(doc, activities, sugerido, extensiones):
    add_heading(doc, "Tabla de actividades desintensificadas y ahorros", level=3)
    headers = [
        "Actividad",
        "Tiempo Normal",
        "Tiempo Intensivo",
        "Costo Normal",
        "Costo Intensivo",
        "Máx. Intensif.",
        "Costo por Día",
        "Tiempo Sugerido",
        "Días Desintensificados",
        "Ahorro",
    ]
    rows = []
    mark_rows = []
    for i, a in enumerate(sorted(activities)):
        info = activities[a]
        ahorro = info["costo_por_dia"] * extensiones[a]
        if extensiones[a] > 0:
            mark_rows.append(i)
        rows.append([
            a,
            fmt_num(info["tiempo_normal"]),
            fmt_num(info["tiempo_intensivo"]),
            fmt_num(info["costo_normal"]),
            fmt_num(info["costo_intensivo"]),
            fmt_num(info["max_intensificacion"]),
            fmt_num(info["costo_por_dia"]),
            fmt_num(sugerido[a]),
            fmt_num(extensiones[a]),
            fmt_num(ahorro),
        ])
    add_table_from_rows(doc, headers, rows, mark_rows=mark_rows)
    doc.add_paragraph()


def write_benefit_cost_table(doc, data):
    add_heading(doc, "Cuadro comparativo (a) Normal · (b) Tint costo máximo · (c) Tint costo óptimo", level=3)

    paths = data["paths"]
    activities = data["activities"]
    successors = data["successors"]
    normal_time = data["normal_time"]
    crash_time = data["crash_time"]
    sugerido = data["sugerido"]
    t_normal = data["t_normal"]
    t_optimal = data["t_optimal"]
    total_normal_cost = data["total_normal_cost"]
    total_crash_cost = data["total_crash_cost"]
    costo_optimo = data["costo_optimo"]

    _, _, _, _, holg_a, _ = cpm_pass(activities, successors, normal_time)
    _, _, _, _, holg_b, _ = cpm_pass(activities, successors, crash_time)
    _, _, _, _, holg_c, _ = cpm_pass(activities, successors, sugerido)
    holg_total_act_a = sum(holg_a.values())
    holg_total_act_b = sum(holg_b.values())
    holg_total_act_c = sum(holg_c.values())
    holg_total_cam_a, _ = total_path_slack(paths, normal_time)
    holg_total_cam_b, _ = total_path_slack(paths, crash_time)
    holg_total_cam_c, _ = total_path_slack(paths, sugerido)

    sobrecosto_b = total_crash_cost - total_normal_cost
    sobrecosto_c = costo_optimo - total_normal_cost
    dias_ahorrados = t_normal - t_optimal

    headers = ["Métrica", "(a) Normal", "(b) Tint costo máximo", "(c) Tint costo óptimo"]
    rows = [
        ["Duración del proyecto",
         fmt_num(t_normal), fmt_num(t_optimal), fmt_num(t_optimal)],
        ["Costo total del proyecto",
         fmt_num(total_normal_cost), fmt_num(total_crash_cost), fmt_num(costo_optimo)],
        ["Días ahorrados vs Tn",
         "0", fmt_num(dias_ahorrados), fmt_num(dias_ahorrados)],
        ["Sobrecosto vs Costo Normal",
         "0", fmt_num(sobrecosto_b), fmt_num(sobrecosto_c)],
        ["Costo por día ahorrado",
         "-",
         fmt_num(sobrecosto_b / dias_ahorrados) if dias_ahorrados else "-",
         fmt_num(sobrecosto_c / dias_ahorrados) if dias_ahorrados else "-"],
        ["Holgura total (actividades)",
         fmt_num(holg_total_act_a),
         fmt_num(holg_total_act_b),
         fmt_num(holg_total_act_c)],
        ["Margen de maniobra (suma de holguras de caminos)",
         fmt_num(holg_total_cam_a),
         fmt_num(holg_total_cam_b),
         fmt_num(holg_total_cam_c)],
    ]
    add_table_from_rows(doc, headers, rows)
    doc.add_paragraph()

    add_heading(doc, "Interpretación beneficio / costo", level=3)
    if sobrecosto_b > 0:
        ahorro_vs_b = sobrecosto_b - sobrecosto_c
        pct = 100 * ahorro_vs_b / sobrecosto_b
        add_paragraph(
            doc,
            f"La opción (c) ahorra {fmt_num(ahorro_vs_b)} respecto a la opción (b), "
            f"un {fmt_num(pct)}% menos de sobrecosto, "
            f"manteniendo el mismo Tint = {fmt_num(t_optimal)} días.",
        )
    add_paragraph(
        doc,
        f"Margen de maniobra opción (b): {fmt_num(holg_total_cam_b)} unidades-día.",
    )
    add_paragraph(
        doc,
        f"Margen de maniobra opción (c): {fmt_num(holg_total_cam_c)} unidades-día.",
    )
    if holg_total_cam_b > holg_total_cam_c:
        add_paragraph(
            doc,
            "La opción (b) conserva más holgura — ofrece más flexibilidad operativa, "
            "pero paga el costo intensivo máximo. La opción (c) consume parte de esa "
            "holgura para convertirla en ahorros de costo.",
        )
    doc.add_paragraph()


def build_report(xlsx_path, output_path):
    data = compute_analysis(xlsx_path)
    if data is None:
        raise RuntimeError("No se encontraron caminos en la red.")

    doc = Document()

    # ----- Carátula -----
    title = doc.add_heading("Taller Grupal — Crashing de Proyecto (CPM)", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Investigación de Operaciones — INVOPER").italic = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"Fecha: {date.today().strftime('%d/%m/%Y')}").italic = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Integrantes: ____________________________________________").italic = True

    doc.add_paragraph()
    add_paragraph(
        doc,
        f"Archivo de entrada: {Path(xlsx_path).name}",
        italic=True,
    )
    doc.add_page_break()

    # ----- Datos de entrada -----
    add_heading(doc, "Resumen del proyecto", level=1)
    add_paragraph(
        doc,
        f"El proyecto tiene {len(data['activities'])} actividades. "
        f"Actividades iniciales: {', '.join(data['starts'])}. "
        f"Actividades finales: {', '.join(data['ends'])}. "
        f"Costo normal total: {fmt_num(data['total_normal_cost'])}. "
        f"Costo intensivo total: {fmt_num(data['total_crash_cost'])}.",
    )
    write_input_table(doc, data["activities"], data["successors"])
    doc.add_paragraph()

    # ----- PUNTO 1 -----
    add_heading(
        doc,
        "Punto 1 — Resolver la red con tiempos normales y costos normales",
        level=1,
    )
    add_paragraph(
        doc,
        "Se resuelve la red usando los tiempos y costos normales. Se reportan "
        "los caminos con sus duraciones y holguras, la tabla de holguras por "
        "actividad y la ruta crítica.",
    )
    write_paths_table(
        doc, data["paths"], data["normal_time"],
        "Caminos con tiempos normales",
    )
    write_activity_slack_table(
        doc, data["activities"], data["successors"], data["normal_time"],
        "Tabla de holguras por actividad — tiempos normales",
    )
    add_paragraph(
        doc,
        f"Tn = {fmt_num(data['t_normal'])} días.   "
        f"Costo total del proyecto = {fmt_num(data['total_normal_cost'])}.",
        bold=True,
    )
    add_screenshot_placeholder(
        doc,
        "POM-QM ▸ Project Management (PERT/CPM) ▸ Solution con tiempos NORMALES "
        "(tabla de actividades y ruta crítica)",
    )
    add_screenshot_placeholder(
        doc,
        "POM-QM ▸ Project Management (PERT/CPM) ▸ Gantt Chart con tiempos NORMALES",
    )
    add_screenshot_placeholder(
        doc,
        "POM-QM ▸ Project Management (PERT/CPM) ▸ Precedence Graph (red AON) — NORMAL",
    )
    doc.add_page_break()

    # ----- PUNTO 2 -----
    add_heading(
        doc,
        "Punto 2 — Resolver la red con tiempos intensivos y costos intensivos",
        level=1,
    )
    add_paragraph(
        doc,
        "Se resuelve la red usando los tiempos intensivos. Se reporta Tint, "
        "el costo máximo intensivo (todas las actividades intensificadas) y el "
        "costo óptimo para Tint que entrega el modelo.",
    )
    write_paths_table(
        doc, data["paths"], data["crash_time"],
        "Caminos con tiempos intensivos",
    )
    write_activity_slack_table(
        doc, data["activities"], data["successors"], data["crash_time"],
        "Tabla de holguras por actividad — tiempos intensivos, costo máximo (opción b)",
    )
    add_paragraph(
        doc,
        f"Tint = {fmt_num(data['t_optimal'])} días.   "
        f"Costo intensivo MÁXIMO = {fmt_num(data['total_crash_cost'])}.   "
        f"Costo ÓPTIMO para Tint = {fmt_num(data['costo_optimo'])}.",
        bold=True,
    )
    add_screenshot_placeholder(
        doc,
        "POM-QM ▸ Project Management - CPM/Costing ▸ Solution con tiempos INTENSIVOS",
    )
    add_screenshot_placeholder(
        doc,
        "POM-QM ▸ Project Management - CPM/Costing ▸ Gantt Chart INTENSIVO",
    )
    add_screenshot_placeholder(
        doc,
        "POM-QM ▸ Project Management - CPM/Costing ▸ Crashing Schedule "
        "(reporte de costo óptimo por etapa)",
    )
    doc.add_page_break()

    # ----- PUNTO 3 -----
    add_heading(
        doc,
        "Punto 3 — Tabla de actividades desintensificadas y ahorros",
        level=1,
    )
    add_paragraph(
        doc,
        "Se presenta la tabla con días desintensificados por actividad y los "
        "ahorros respectivos. Se verifica que la suma de los ahorros (ahorro "
        "total) sea igual a la diferencia entre el costo intensivo máximo y el "
        "costo óptimo entregado por el modelo.",
    )
    write_desintensification_table(
        doc, data["activities"], data["sugerido"], data["extensiones"]
    )
    add_heading(doc, "Verificación de ahorros", level=3)
    diff = data["total_crash_cost"] - data["costo_optimo"]
    coincide = abs(data["ahorros"] - diff) < 1e-6
    add_paragraph(
        doc,
        f"Suma de ahorros individuales (ahorro total) = {fmt_num(data['ahorros'])}.",
    )
    add_paragraph(
        doc,
        f"Costo intensivo máximo − Costo óptimo "
        f"= {fmt_num(data['total_crash_cost'])} − {fmt_num(data['costo_optimo'])} "
        f"= {fmt_num(diff)}.",
    )
    add_paragraph(
        doc,
        f"¿Coinciden? {'SÍ' if coincide else 'NO'}.",
        bold=True,
    )
    add_screenshot_placeholder(
        doc,
        "POM-QM ▸ Crashing Schedule completo (mostrando costo por etapa hasta llegar a Tint)",
    )
    doc.add_page_break()

    # ----- PUNTO 4 -----
    add_heading(
        doc,
        "Punto 4 — Análisis beneficio/costo de las opciones (b) y (c)",
        level=1,
    )
    add_paragraph(
        doc,
        "Se analizan las tablas de holguras de las opciones (b) Tint con costo "
        "máximo y (c) Tint con costo óptimo, tanto para las actividades como "
        "para los caminos, junto con el sobrecosto y el margen de maniobra de "
        "cada una.",
    )
    write_benefit_cost_table(doc, data)

    add_heading(
        doc,
        "Caminos con tiempos óptimos sugeridos (opción c)",
        level=3,
    )
    write_paths_table(
        doc, data["paths"], data["sugerido"],
        "Caminos con tiempos óptimos sugeridos (opción c)",
    )
    write_activity_slack_table(
        doc, data["activities"], data["successors"], data["sugerido"],
        "Tabla de holguras por actividad — tiempos intensivos, costo óptimo (opción c)",
    )
    add_screenshot_placeholder(
        doc,
        "POM-QM ▸ Comparativo/screenshot de la solución con costo ÓPTIMO para Tint",
    )

    # ----- Diagramas de Gantt en Word (opcional) -----
    add_heading(doc, "Diagramas de Gantt — visualización opcional", level=2)
    add_paragraph(
        doc,
        "Se recomienda copiar las capturas de Gantt directamente desde POM-QM. "
        "Si se desea, se puede pegar abajo el diagrama generado por el script "
        "`cpm_crashing.py` para cada opción.",
        italic=True,
    )
    add_screenshot_placeholder(
        doc, "Gantt opción (a) Tiempos NORMALES (POM-QM)"
    )
    add_screenshot_placeholder(
        doc, "Gantt opción (b) Tiempos INTENSIVOS, costo máximo (POM-QM)"
    )
    add_screenshot_placeholder(
        doc, "Gantt opción (c) Tiempos ÓPTIMOS, costo óptimo (POM-QM)"
    )

    # ----- Conclusiones -----
    add_heading(doc, "Conclusiones", level=1)
    add_paragraph(
        doc,
        "• La duración mínima del proyecto (Tint) es "
        f"{fmt_num(data['t_optimal'])} días, "
        f"frente a Tn = {fmt_num(data['t_normal'])} días con tiempos normales.",
    )
    add_paragraph(
        doc,
        f"• Para alcanzar Tint, el costo intensivo máximo es "
        f"{fmt_num(data['total_crash_cost'])}, pero existe un costo ÓPTIMO de "
        f"{fmt_num(data['costo_optimo'])} que mantiene la misma duración con "
        f"un ahorro de {fmt_num(data['ahorros'])}.",
    )
    add_paragraph(
        doc,
        "• La elección entre la opción (b) y la opción (c) depende del valor "
        "estratégico del margen de maniobra: la opción (b) conserva más "
        "holgura operativa, mientras la opción (c) la convierte en ahorro.",
    )
    add_paragraph(doc, "• ____________________________________________________")
    add_paragraph(doc, "• ____________________________________________________")

    doc.save(output_path)
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python generate_report.py <archivo.xlsx> [salida.docx]")
        sys.exit(1)
    xlsx = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "Reporte_Crashing.docx"
    generated = build_report(xlsx, out)
    print(f"Reporte generado en: {generated}")
