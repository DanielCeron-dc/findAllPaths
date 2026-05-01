"""
Analizador de Crashing de Redes CPM
-----------------------------------
Lee un archivo de Excel con la siguiente estructura:

    Actividad | Tiempo Normal | Tiempo Intensivo | Costo Normal | Costo Intensivo |
        Predecesor 1 | Predecesor 2 | ... | Predecesor N

Para cada actividad conoce los tiempos y costos normales/intensivos. El programa:
  - Construye el grafo de actividades en los nodos (AON)
  - Enumera todos los caminos del nodo inicial al nodo final
  - Lista cada camino con tiempos NORMALES y holguras
  - Lista cada camino con tiempos INTENSIVOS y holguras
  - Calcula el TIEMPO ÓPTIMO del proyecto (= duración de la ruta crítica intensiva)
  - Calcula el COSTO ÓPTIMO del proyecto (= cronograma más barato que cumple Tint)
  - Imprime las tablas de holguras de actividades para cada escenario
  - Compara las opciones (a) Normal, (b) Tint costo máximo, (c) Tint costo óptimo
"""

from __future__ import annotations
import sys
from pathlib import Path
import pandas as pd
import numpy as np
from scipy.optimize import linprog


def find_header_row(xlsx_path: str | Path) -> int:
    """Detecta automáticamente la fila de encabezados buscando 'Activity' o 'Actividad'."""
    raw = pd.read_excel(xlsx_path, header=None)
    for i, row in raw.iterrows():
        for c in row.values:
            if isinstance(c, str) and c.strip().lower() in ("activity", "actividad"):
                return i
    return 0


def load_network(xlsx_path: str | Path):
    """
    Devuelve:
      activities:  {act: {'tiempo_normal', 'tiempo_intensivo', 'costo_normal',
                          'costo_intensivo', 'max_intensificacion', 'costo_por_dia'}}
      successors:  {act: [actividades que la siguen]}
    """
    header_row = find_header_row(xlsx_path)
    df = pd.read_excel(xlsx_path, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    cols = list(df.columns)
    activity_col = cols[0]
    normal_time_col = cols[1]
    crash_time_col = cols[2]
    normal_cost_col = cols[3]
    crash_cost_col = cols[4]
    predecessor_cols = cols[5:]

    activities: dict[str, dict] = {}
    successors: dict[str, list[str]] = {}

    for _, row in df.iterrows():
        act_raw = row[activity_col]
        if pd.isna(act_raw):
            continue
        act = str(act_raw).strip()
        if not act or act.lower() == "nan":
            continue

        tiempo_normal = float(row[normal_time_col])
        tiempo_intensivo = float(row[crash_time_col])
        costo_normal = float(row[normal_cost_col])
        costo_intensivo = float(row[crash_cost_col])
        max_intensificacion = tiempo_normal - tiempo_intensivo
        costo_por_dia = (
            (costo_intensivo - costo_normal) / max_intensificacion
            if max_intensificacion > 0
            else 0.0
        )

        activities[act] = {
            "tiempo_normal": tiempo_normal,
            "tiempo_intensivo": tiempo_intensivo,
            "costo_normal": costo_normal,
            "costo_intensivo": costo_intensivo,
            "max_intensificacion": max_intensificacion,
            "costo_por_dia": costo_por_dia,
        }
        successors.setdefault(act, [])

    for _, row in df.iterrows():
        act_raw = row[activity_col]
        if pd.isna(act_raw):
            continue
        act = str(act_raw).strip()
        if not act or act.lower() == "nan":
            continue
        for pcol in predecessor_cols:
            val = row[pcol]
            if pd.isna(val):
                continue
            pred = str(val).strip()
            if not pred or pred.lower() == "nan":
                continue
            successors.setdefault(pred, []).append(act)

    return activities, successors


def find_starts_and_ends(activities, successors):
    all_nodes = set(activities)
    has_predecessor = {child for ch in successors.values() for child in ch}
    starts = sorted(all_nodes - has_predecessor)
    ends = sorted(n for n in all_nodes if not successors.get(n))
    return starts, ends


def all_paths(successors, starts, ends):
    paths = []
    end_set = set(ends)

    def dfs(node, current):
        current.append(node)
        if node in end_set and not successors.get(node):
            paths.append(current.copy())
        else:
            for nxt in successors.get(node, []):
                dfs(nxt, current)
        current.pop()

    for s in starts:
        dfs(s, [])

    return paths


def path_duration(path, time_dict):
    return sum(time_dict[a] for a in path)


def predecessors_of(successors):
    preds: dict[str, list[str]] = {a: [] for a in successors}
    for s, succs in successors.items():
        for t in succs:
            preds.setdefault(t, []).append(s)
    return preds


def topological_order(activities, successors):
    in_deg = {a: 0 for a in activities}
    for s, succs in successors.items():
        for t in succs:
            in_deg[t] = in_deg.get(t, 0) + 1
    queue = [a for a in activities if in_deg[a] == 0]
    order = []
    while queue:
        n = queue.pop(0)
        order.append(n)
        for nxt in successors.get(n, []):
            in_deg[nxt] -= 1
            if in_deg[nxt] == 0:
                queue.append(nxt)
    return order


def cpm_pass(activities, successors, time_dict):
    """Pasada hacia adelante y hacia atrás. Devuelve IC, TC, IL, TL, holgura, duración."""
    preds = predecessors_of(successors)
    topo = topological_order(activities, successors)

    IC, TC = {}, {}
    for a in topo:
        IC[a] = max((TC[p] for p in preds.get(a, [])), default=0)
        TC[a] = IC[a] + time_dict[a]

    project_duration = max(TC.values())

    TL, IL = {}, {}
    for a in reversed(topo):
        succs = successors.get(a, [])
        TL[a] = min((IL[s] for s in succs), default=project_duration)
        IL[a] = TL[a] - time_dict[a]

    holgura = {a: IL[a] - IC[a] for a in activities}
    return IC, TC, IL, TL, holgura, project_duration


def fmt_num(n):
    if n is None:
        return "-"
    if abs(n - round(n)) < 1e-9:
        return f"{int(round(n))}"
    return f"{n:.2f}"


WIDTH = 130


def print_paths_table(title, paths, time_dict):
    durations = [path_duration(p, time_dict) for p in paths]
    critical = max(durations)
    scored = sorted(zip(durations, paths), key=lambda x: -x[0])
    print(title)
    print("-" * WIDTH)
    print(
        f"{'#':<6}"
        f"{'Duración del Camino':<22}"
        f"{'Holgura del Camino':<22}"
        f"{'Camino (actividades con su tiempo)'}"
    )
    print("-" * WIDTH)
    for i, (dur, p) in enumerate(scored, 1):
        slack = critical - dur
        marker = "  <-- RUTA CRÍTICA" if abs(dur - critical) < 1e-9 else ""
        path_str = " -> ".join(f"{a}({fmt_num(time_dict[a])})" for a in p)
        print(
            f"{i:<6}"
            f"{fmt_num(dur):<22}"
            f"{fmt_num(slack):<22}"
            f"{path_str}{marker}"
        )
    print("-" * WIDTH)
    print(f"Duración de la ruta crítica: {fmt_num(critical)}")
    print()
    return critical


def solve_optimal(activities, paths, t_optimal):
    """
    PL para el cronograma de mínimo costo cuya duración del proyecto es t_optimal.

    Variable x_a = días que la actividad se EXTIENDE desde el tiempo intensivo
    hacia el normal (0 <= x_a <= max_intensificacion_a). El costo total es
    costo_intensivo_total - sum(costo_por_dia_a * x_a). Por tanto, MAXIMIZAMOS
    los ahorros sum(costo_por_dia_a * x_a) (equivalente a minimizar el negativo).

    Restricciones: para cada camino P,
        sum_{a en P} x_a <= t_optimal - sum_{a en P} tiempo_intensivo_a
    """
    act_list = list(activities.keys())
    n = len(act_list)
    idx = {a: i for i, a in enumerate(act_list)}

    c = np.array([-activities[a]["costo_por_dia"] for a in act_list], dtype=float)

    A_rows, b_rows = [], []
    for path in paths:
        row = np.zeros(n)
        for a in path:
            row[idx[a]] += 1
        crash_sum = sum(activities[a]["tiempo_intensivo"] for a in path)
        A_rows.append(row)
        b_rows.append(t_optimal - crash_sum)

    A_ub = np.array(A_rows)
    b_ub = np.array(b_rows)
    bounds = [(0.0, activities[a]["max_intensificacion"]) for a in act_list]

    res = linprog(c, A_ub=A_ub, b_ub=b_ub, bounds=bounds, method="highs")
    if not res.success:
        raise RuntimeError(f"El solver de PL falló: {res.message}")

    extensiones = {a: int(round(res.x[idx[a]])) for a in act_list}
    sugerido = {
        a: activities[a]["tiempo_intensivo"] + extensiones[a] for a in act_list
    }
    ahorros = sum(
        activities[a]["costo_por_dia"] * extensiones[a] for a in act_list
    )
    costo_intensivo_total = sum(
        activities[a]["costo_intensivo"] for a in act_list
    )
    costo_optimo = costo_intensivo_total - ahorros
    return sugerido, extensiones, ahorros, costo_optimo


def print_activity_slack_table(title, activities, successors, time_dict):
    """Tabla de holguras a nivel de actividad (IC, TC, IL, TL, Holgura)."""
    IC, TC, IL, TL, holgura, project_duration = cpm_pass(
        activities, successors, time_dict
    )
    print(title)
    print("-" * WIDTH)
    print(
        f"{'Actividad':<14}"
        f"{'Duración':<14}"
        f"{'Inicio Cercano':<18}"
        f"{'Término Cercano':<18}"
        f"{'Inicio Lejano':<18}"
        f"{'Término Lejano':<18}"
        f"{'Holgura Total':<16}"
        f"{'¿Crítica?'}"
    )
    print("-" * WIDTH)
    for a in sorted(activities):
        es_critica = abs(holgura[a]) < 1e-9
        print(
            f"{a:<14}"
            f"{fmt_num(time_dict[a]):<14}"
            f"{fmt_num(IC[a]):<18}"
            f"{fmt_num(TC[a]):<18}"
            f"{fmt_num(IL[a]):<18}"
            f"{fmt_num(TL[a]):<18}"
            f"{fmt_num(holgura[a]):<16}"
            f"{'SÍ' if es_critica else 'No'}"
        )
    print("-" * WIDTH)
    print(f"Duración del proyecto: {fmt_num(project_duration)}")
    print()
    return holgura, project_duration


def print_activity_table(activities, sugerido, extensiones):
    print("Análisis de intensificación por actividad")
    print("-" * WIDTH)
    headers = [
        "Actividad",
        "Tiempo Normal",
        "Tiempo Intensivo",
        "Costo Normal",
        "Costo Intensivo",
        "Máx. Intensif.",
        "Costo por Día",
        "Tiempo Sugerido",
        "Días Desintens.",
        "Ahorro",
    ]
    fmt = (
        "{:<11}"  # Actividad
        "{:<16}"  # Tiempo Normal
        "{:<19}"  # Tiempo Intensivo
        "{:<15}"  # Costo Normal
        "{:<18}"  # Costo Intensivo
        "{:<17}"  # Máx. Intensif.
        "{:<16}"  # Costo por Día
        "{:<18}"  # Tiempo Sugerido
        "{:<18}"  # Días Desintens.
        "{:<10}"  # Ahorro
    )
    print(fmt.format(*headers))
    print("-" * WIDTH)
    for a in sorted(activities):
        info = activities[a]
        ahorro = info["costo_por_dia"] * extensiones[a]
        print(fmt.format(
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
        ))
    print("-" * WIDTH)
    print()


def print_gantt_chart(title, activities, successors, time_dict):
    """
    Diagrama de Gantt en ASCII basado en la programación temprana (Inicio Cercano):
        █  trabajo programado en la actividad
        ░  holgura disponible (margen para retrasarse sin afectar Tint)
        Las actividades críticas (holgura=0) llevan la marca <-- CRÍTICA.
    """
    IC, _TC, IL, _TL, holgura, project_duration = cpm_pass(
        activities, successors, time_dict
    )
    T = int(round(project_duration))

    print(title)
    print("-" * (16 + T + 16))

    label_width = 18
    scale = " " * label_width + " "
    for i in range(0, T + 1):
        scale += str(i % 10) if i % 5 != 0 else "|"
    print(scale)
    ruler = " " * label_width + " "
    for i in range(0, T + 1):
        ruler += "+" if i % 10 == 0 else ("." if i % 5 == 0 else " ")
    print(ruler)

    for a in sorted(activities):
        dur = int(round(time_dict[a]))
        es = int(round(IC[a]))
        slack = int(round(holgura[a]))
        es_critica = slack == 0

        bar = [" "] * (T + 1)
        for k in range(es, es + dur):
            if 0 <= k < len(bar):
                bar[k] = "█"
        for k in range(es + dur, es + dur + slack):
            if 0 <= k < len(bar):
                bar[k] = "░"

        marca = "  <-- CRÍTICA" if es_critica else ""
        etiqueta = f"{a} (dur={dur}, h={slack})"
        print(f"{etiqueta:<{label_width}} {''.join(bar)}{marca}")

    print("-" * (16 + T + 16))
    print(f"Duración total del proyecto: {fmt_num(project_duration)} días")
    print("Leyenda: █ = trabajo programado    ░ = holgura disponible")
    print()


def total_path_slack(paths, time_dict):
    durations = [path_duration(p, time_dict) for p in paths]
    critical = max(durations)
    return sum(critical - d for d in durations), critical


def benefit_cost_summary(
    paths, activities, successors,
    normal_time, crash_time, sugerido,
    total_normal_cost, total_crash_cost, costo_optimo,
    t_normal, t_optimal,
):
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

    rows = [
        ("Duración del proyecto",
         fmt_num(t_normal), fmt_num(t_optimal), fmt_num(t_optimal)),
        ("Costo total del proyecto",
         fmt_num(total_normal_cost),
         fmt_num(total_crash_cost),
         fmt_num(costo_optimo)),
        ("Días ahorrados vs Tn",
         "0", fmt_num(dias_ahorrados), fmt_num(dias_ahorrados)),
        ("Sobrecosto vs Costo Normal",
         "0", fmt_num(sobrecosto_b), fmt_num(sobrecosto_c)),
        ("Costo por día ahorrado",
         "-",
         fmt_num(sobrecosto_b / dias_ahorrados) if dias_ahorrados else "-",
         fmt_num(sobrecosto_c / dias_ahorrados) if dias_ahorrados else "-"),
        ("Holgura total (actividades)",
         fmt_num(holg_total_act_a),
         fmt_num(holg_total_act_b),
         fmt_num(holg_total_act_c)),
        ("Margen de maniobra (caminos)",
         fmt_num(holg_total_cam_a),
         fmt_num(holg_total_cam_b),
         fmt_num(holg_total_cam_c)),
    ]

    print(
        f"{'Métrica':<34}"
        f"{'(a) Normal':<22}"
        f"{'(b) Tint costo máximo':<26}"
        f"{'(c) Tint costo óptimo':<26}"
    )
    print("-" * WIDTH)
    for r in rows:
        print(f"{r[0]:<34}{r[1]:<22}{r[2]:<26}{r[3]:<26}")
    print("-" * WIDTH)
    print()

    print("Interpretación beneficio / costo")
    print("-" * WIDTH)
    if sobrecosto_c > 0 and sobrecosto_b > 0:
        ahorro_vs_b = sobrecosto_b - sobrecosto_c
        print(
            f"  La opción (c) ahorra {fmt_num(ahorro_vs_b)} respecto a la opción (b) "
            f"({fmt_num(100 * ahorro_vs_b / sobrecosto_b)}% menos sobrecosto),"
        )
        print(f"  manteniendo el mismo Tint = {fmt_num(t_optimal)}.")
    print(
        f"  Margen de maniobra opción (b): {fmt_num(holg_total_cam_b)} "
        f"unidades-día (suma de holguras de caminos)"
    )
    print(
        f"  Margen de maniobra opción (c): {fmt_num(holg_total_cam_c)} "
        f"unidades-día (suma de holguras de caminos)"
    )
    if holg_total_cam_b > holg_total_cam_c:
        print(
            "  La opción (b) conserva más holgura -> más flexibilidad operativa,"
        )
        print("  pero paga el costo intensivo máximo.")
        print(
            "  La opción (c) consume parte de esa holgura para convertirla en "
            "ahorros de costo."
        )
    print()


def compute_analysis(xlsx_path):
    """Carga el archivo y devuelve un diccionario con todos los resultados."""
    activities, successors = load_network(xlsx_path)
    starts, ends = find_starts_and_ends(activities, successors)
    paths = all_paths(successors, starts, ends)
    if not paths:
        return None

    total_normal_cost = sum(a["costo_normal"] for a in activities.values())
    total_crash_cost = sum(a["costo_intensivo"] for a in activities.values())
    normal_time = {a: info["tiempo_normal"] for a, info in activities.items()}
    crash_time = {a: info["tiempo_intensivo"] for a, info in activities.items()}

    t_normal = max(path_duration(p, normal_time) for p in paths)
    t_optimal = max(path_duration(p, crash_time) for p in paths)

    sugerido, extensiones, ahorros, costo_optimo = solve_optimal(
        activities, paths, t_optimal
    )

    return {
        "xlsx_path": str(xlsx_path),
        "activities": activities,
        "successors": successors,
        "starts": starts,
        "ends": ends,
        "paths": paths,
        "normal_time": normal_time,
        "crash_time": crash_time,
        "sugerido": sugerido,
        "extensiones": extensiones,
        "ahorros": ahorros,
        "costo_optimo": costo_optimo,
        "total_normal_cost": total_normal_cost,
        "total_crash_cost": total_crash_cost,
        "t_normal": t_normal,
        "t_optimal": t_optimal,
    }


def analyze(xlsx_path):
    activities, successors = load_network(xlsx_path)
    starts, ends = find_starts_and_ends(activities, successors)
    paths = all_paths(successors, starts, ends)

    if not paths:
        print(
            "No se encontraron caminos. Verifique que la red tenga "
            "actividades de inicio y de final."
        )
        return

    total_normal_cost = sum(a["costo_normal"] for a in activities.values())
    total_crash_cost = sum(a["costo_intensivo"] for a in activities.values())

    print("=" * WIDTH)
    print(f"Archivo analizado: {xlsx_path}")
    print(f"Actividades ({len(activities)}): {', '.join(sorted(activities))}")
    print(f"Actividades iniciales: {', '.join(starts)}")
    print(f"Actividades finales:   {', '.join(ends)}")
    print(f"Costo NORMAL total:    {fmt_num(total_normal_cost)}")
    print(f"Costo INTENSIVO total: {fmt_num(total_crash_cost)}")
    print("=" * WIDTH)
    print()

    normal_time = {a: info["tiempo_normal"] for a, info in activities.items()}
    crash_time = {a: info["tiempo_intensivo"] for a, info in activities.items()}

    # ---- Punto 1: red NORMAL ----
    print(
        "########## PUNTO 1: RED NORMAL "
        "(Tn, costo total normal, ruta crítica) ##########\n"
    )
    t_normal = print_paths_table(
        "CAMINOS CON TIEMPOS NORMALES", paths, normal_time
    )
    print_activity_slack_table(
        "TABLA DE HOLGURAS DE ACTIVIDADES - TIEMPOS NORMALES (opción a)",
        activities, successors, normal_time,
    )

    # ---- Punto 2: red INTENSIVA ----
    print(
        "########## PUNTO 2: RED INTENSIVA "
        "(Tint, costo intensivo máximo, costo óptimo, rutas críticas) ##########\n"
    )
    t_optimal = print_paths_table(
        "CAMINOS CON TIEMPOS INTENSIVOS", paths, crash_time
    )
    print_activity_slack_table(
        "TABLA DE HOLGURAS DE ACTIVIDADES - TIEMPOS INTENSIVOS, COSTO MÁXIMO (opción b)",
        activities, successors, crash_time,
    )

    sugerido, extensiones, ahorros, costo_optimo = solve_optimal(
        activities, paths, t_optimal
    )

    print(f"Tint (duración mínima del proyecto, todo intensificado): {fmt_num(t_optimal)}")
    print(f"Costo intensivo MÁXIMO (todas las actividades intensificadas): {fmt_num(total_crash_cost)}")
    print(f"Costo ÓPTIMO para Tint (solución del PL):                      {fmt_num(costo_optimo)}")
    print()

    # ---- Punto 3: actividades desintensificadas y ahorros ----
    print(
        "########## PUNTO 3: ACTIVIDADES DESINTENSIFICADAS Y AHORROS ##########\n"
    )
    print_activity_table(activities, sugerido, extensiones)

    print("Verificación de ahorros:")
    print(
        f"  Suma de ahorros individuales (ahorro total)                     "
        f"= {fmt_num(ahorros)}"
    )
    print(
        f"  Costo intensivo máximo - costo óptimo "
        f"(= {fmt_num(total_crash_cost)} - {fmt_num(costo_optimo)}) "
        f"= {fmt_num(total_crash_cost - costo_optimo)}"
    )
    coincide = abs(ahorros - (total_crash_cost - costo_optimo)) < 1e-6
    print(f"  ¿Coinciden?: {'SÍ' if coincide else 'NO'}")
    print()

    print_paths_table(
        "CAMINOS CON TIEMPOS ÓPTIMOS SUGERIDOS (opción c)",
        paths,
        sugerido,
    )
    print_activity_slack_table(
        "TABLA DE HOLGURAS DE ACTIVIDADES - TIEMPOS INTENSIVOS, COSTO ÓPTIMO (opción c)",
        activities, successors, sugerido,
    )

    # ---- Punto 4: análisis beneficio / costo ----
    print(
        "########## PUNTO 4: ANÁLISIS BENEFICIO / COSTO "
        "(opciones b y c) ##########\n"
    )
    benefit_cost_summary(
        paths, activities, successors,
        normal_time, crash_time, sugerido,
        total_normal_cost, total_crash_cost, costo_optimo,
        t_normal, t_optimal,
    )

    # ---- Diagramas de Gantt ----
    print("########## DIAGRAMAS DE GANTT ##########\n")
    print_gantt_chart(
        "DIAGRAMA DE GANTT - OPCIÓN (a) TIEMPOS NORMALES",
        activities, successors, normal_time,
    )
    print_gantt_chart(
        "DIAGRAMA DE GANTT - OPCIÓN (b) TIEMPOS INTENSIVOS, COSTO MÁXIMO",
        activities, successors, crash_time,
    )
    print_gantt_chart(
        "DIAGRAMA DE GANTT - OPCIÓN (c) TIEMPOS ÓPTIMOS, COSTO ÓPTIMO",
        activities, successors, sugerido,
    )


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python cpm_crashing.py <ruta_al_archivo_excel.xlsx>")
        sys.exit(1)
    analyze(sys.argv[1])
