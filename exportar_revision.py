"""
Excel de la revisión de cotización (descargable desde la app)
============================================================

Se genera con XlsxWriter para poder guardar CADA fórmula junto con su
resultado ya calculado. Con openpyxl las fórmulas quedaban sin valor
guardado y, al abrir el archivo en la vista previa de la Mac, en el correo
o en Drive, las columnas de % y semáforo aparecían vacías. Ahora se ven en
cualquier visor y en Excel siguen siendo fórmulas vivas.

Hojas:
  1. Resumen      una fila por partida con las 4 fuentes lado a lado
                  (P.U., estado y %), semáforo final, referencia para
                  negociar y ahorro.
  2. Por fuente   las 4 comparaciones con % y semáforo calculados por
                  fórmula a partir del precio de referencia.
  3. Evidencia    qué encontró cada fuente y por qué quedó en ese estado.
  4. Metodología  reglas de validación.
"""

from __future__ import annotations

import datetime as _dt
import io

try:
    import xlsxwriter
    from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell
    XLSXWRITER_DISPONIBLE = True
except ImportError:  # el servidor aún no instala requirements.txt
    XLSXWRITER_DISPONIBLE = False

import ajuste_inflacion as _infl
import validacion_referencias as v

FUENTES = ("historico", "nl", "cdmx", "ia")
PF = 6   # columnas por fuente en la hoja 'Por fuente'
NOMBRE_CORTO = {"historico": "Histórico", "nl": "Nuevo León", "cdmx": "CDMX", "ia": "IA internet"}

AZUL, AZUL2 = "#1F3864", "#2E5496"
ROJO_F, ROJO_T = "#F8C9C9", "#712121"
VERDE_F, VERDE_T = "#CCEBD2", "#14532D"
AMBAR_F, AMBAR_T = "#FFF0BB", "#705000"
NARANJA_F, NARANJA_T = "#FDE7D2", "#7A3E00"
GRIS_F, GRIS_T = "#E4E7EB", "#475467"
ESTADO_COLOR = {
    v.VALIDADA: VERDE_T, v.RECHAZADA: "#B42318",
    v.NO_CONCLUYENTE: "#8A6100", v.POR_CONFIRMAR: "#667085", v.SIN_DATO: "#98A2B3",
    v.EQUIVALENCIA_PARCIAL: "#8A6100", v.NO_COMPARABLE: "#98A2B3",
}


def pd_median(valores):
    v = sorted(float(x) for x in valores)
    n = len(v)
    return (v[n // 2] if n % 2 else (v[n // 2 - 1] + v[n // 2]) / 2) if n else 0.0


def _num(valor):
    try:
        if valor is None or valor == "":
            return None
        n = float(valor)
        return None if n != n else n  # NaN
    except (TypeError, ValueError):
        return None


def _clasif(precio, ref):
    if precio is None or not ref:
        return ""
    return "ALTO" if precio > ref else "BAJO"


def generar_excel(filas: list[dict], proveedor: str = "", proyecto: str = "",
                  revision_cantidades: dict | None = None, configuracion: str = "",
                  datos_mercado: dict | None = None, recomendacion: dict | None = None) -> bytes:
    if not XLSXWRITER_DISPONIBLE:
        import exportar_revision_respaldo
        return exportar_revision_respaldo.generar_excel(filas, proveedor=proveedor, proyecto=proyecto)
    buffer = io.BytesIO()
    wb = xlsxwriter.Workbook(buffer, {"in_memory": True, "nan_inf_to_errors": True})

    base = {"font_name": "Arial", "font_size": 10, "border": 1, "border_color": "#BFBFBF",
            "valign": "top"}

    def fmt(**extra):
        d = dict(base)
        d.update(extra)
        return wb.add_format(d)

    f_titulo = wb.add_format({"font_name": "Arial", "bold": True, "font_size": 14,
                              "font_color": "#FFFFFF", "bg_color": AZUL, "valign": "vcenter"})
    f_sub = wb.add_format({"font_name": "Arial", "font_size": 9, "font_color": "#595959"})
    f_cab = fmt(bold=True, font_color="#FFFFFF", bg_color=AZUL2, align="center",
                valign="vcenter", text_wrap=True)
    f_cab_fuente = {k: fmt(bold=True, font_color="#FFFFFF", bg_color=c, align="center",
                           valign="vcenter", text_wrap=True)
                    for k, c in zip(FUENTES, ("#4B5D7A", "#2E5496", "#3F6E8C", "#5A4E8C"))}
    f_txt = fmt(text_wrap=True)
    f_centro = fmt(align="center")
    f_input_num = fmt(num_format="#,##0.00", font_color="#0000FF")
    f_input_mon = fmt(num_format="$#,##0.00", font_color="#0000FF")
    f_mon = fmt(num_format='$#,##0.00;($#,##0.00);"—"')
    f_dif_mon = fmt(num_format="$#,##0.00;-$#,##0.00")   # mismo signo que en pantalla
    f_mon_neg = fmt(num_format='$#,##0.00;($#,##0.00);"—"', bold=True)
    f_pct = fmt(num_format='+0.0%;-0.0%;0.0%')
    f_pct_b = fmt(num_format="0.0%", bold=True)
    f_bold = fmt(bold=True)
    f_estado = {e: fmt(align="center", bold=True, font_color=c) for e, c in ESTADO_COLOR.items()}
    # P.U. de referencia: pintado como en la app según estado y semáforo.
    f_ref = {
        "ALTO": fmt(num_format="$#,##0.00", bg_color=ROJO_F, font_color=ROJO_T, bold=True),
        "BAJO": fmt(num_format="$#,##0.00", bg_color=VERDE_F, font_color=VERDE_T, bold=True),
        "EN MERCADO": fmt(num_format="$#,##0.00", bg_color=AMBAR_F, font_color=AMBAR_T, bold=True),
        "ORIENTATIVA": fmt(num_format="$#,##0.00", bg_color="#F1F3F5", font_color="#60666D", italic=True),
        # Orientativas con color claro según caro / en mercado / barato.
        "O_ALTO": fmt(num_format="$#,##0.00", bg_color="#FDE4E4", font_color="#8A2B2B", italic=True),
        "O_BAJO": fmt(num_format="$#,##0.00", bg_color="#E3F4E7", font_color="#1F6B3A", italic=True),
        "O_EN MERCADO": fmt(num_format="$#,##0.00", bg_color="#FFF6D6", font_color="#7A5D00", italic=True),
        "RECHAZADA": fmt(num_format="$#,##0.00", font_color="#98A2B3", font_strikeout=True),
        "": f_mon,
    }
    f_semaforo = {
        "ALTO": fmt(align="center", bold=True, bg_color=ROJO_F, font_color=ROJO_T),
        "BAJO": fmt(align="center", bold=True, bg_color=VERDE_F, font_color=VERDE_T),
        "EN MERCADO": fmt(align="center", bold=True, bg_color=AMBAR_F, font_color=AMBAR_T),
        "MIXTO": fmt(align="center", bold=True, bg_color=NARANJA_F, font_color=NARANJA_T),
        "OTRO": fmt(align="center", bold=True, bg_color=GRIS_F, font_color=GRIS_T),
    }

    def estilo_ref(ev):
        estado = ev.get("estado")
        if estado == v.VALIDADA:
            return f_ref.get(ev.get("clasificacion") or "", f_mon)
        if estado in v.ORIENTATIVAS:
            return f_ref.get("O_" + (ev.get("clasificacion") or ""), f_ref["ORIENTATIVA"])
        if estado in (v.RECHAZADA, v.NO_COMPARABLE):
            return f_ref["RECHAZADA"]
        return f_mon

    def semaforo_cf(ws, rango):
        for texto, clave in (("ALTO", "ALTO"), ("BAJO", "BAJO"), ("EN MERCADO", "EN MERCADO")):
            ws.conditional_format(rango, {
                "type": "cell", "criteria": "==", "value": f'"{texto}"',
                "format": wb.add_format({"bg_color": {"ALTO": ROJO_F, "BAJO": VERDE_F,
                                                      "EN MERCADO": AMBAR_F}[clave],
                                         "font_color": {"ALTO": ROJO_T, "BAJO": VERDE_T,
                                                        "EN MERCADO": AMBAR_T}[clave],
                                         "bold": True}),
            })

    # Colores que siguen al dictamen (fórmula): si cambian precios, índices
    # o estados, el color cambia junto con el texto.
    f_cf = {
        "Alto": wb.add_format({"bg_color": "#F3A5A5", "font_color": "#5C1414", "bold": True}),
        "En precio": wb.add_format({"bg_color": "#FFE08A", "font_color": "#5C4300", "bold": True}),
        "Bajo": wb.add_format({"bg_color": "#A8DBB4", "font_color": "#0F3D21", "bold": True}),
        "Posible alto": wb.add_format({"bg_color": "#FDE4E4", "font_color": "#8A2B2B", "italic": True}),
        "Posiblemente en precio": wb.add_format({"bg_color": "#FFF6D6", "font_color": "#7A5D00", "italic": True}),
        "Posible bajo": wb.add_format({"bg_color": "#E3F4E7", "font_color": "#1F6B3A", "italic": True}),
        "Rechazada": wb.add_format({"font_color": "#98A2B3", "font_strikeout": True}),
        "No comparable": wb.add_format({"font_color": "#98A2B3", "font_strikeout": True}),
    }

    def colorear_por_dictamen(hoja, rango, celda_dictamen):
        """Formato condicional del rango según el texto del dictamen."""
        for texto, formato in sorted(f_cf.items(), key=lambda kv: -len(kv[0])):
            hoja.conditional_format(rango, {
                "type": "formula", "criteria": f'={celda_dictamen}="{texto}"',
                "format": formato, "stop_if_true": True,
            })

    hoy = _dt.date.today()

    # ==================================================================
    # Hoja 1: Resumen
    # ==================================================================
    # Única captura: Cantidad y P.U. cotizado se escriben aquí; las demás
    # hojas los toman de esta. P.U. de referencia y estado de cada fuente
    # se capturan en 'Por fuente'; dictamen, semáforo final, referencia,
    # respaldo y ahorro se recalculan con fórmulas a partir de ellos.
    ws = wb.add_worksheet("Resumen")
    cab_ini = ["#", "Concepto", "Unidad", "Cantidad", "P.U. cotizado", "Importe cotizado"]
    sub_fuente = ["P.U. ref.", "Diferencia $/u", "% vs ref.", "Dictamen", "Confiab. fuente"]
    NF = len(sub_fuente)
    cab_fin = ["Resultado de los 4 filtros", "Referencia", "Respaldo", "P.U. de referencia",
               "% vs referencia", "Diferencia contra referencia", "Oportunidad validada (potencial)",
               "Nota (consulta original, no se recalcula)"]
    n_cols = len(cab_ini) + NF * len(FUENTES) + len(cab_fin)
    ws.merge_range(0, 0, 0, n_cols - 1, "Revisión de cotización CAPEX", f_titulo)
    ws.set_row(0, 26)
    ws.write(1, 0, f"Proveedor: {proveedor or '—'} · Proyecto: {proyecto or '—'} · "
                   f"Generado {hoy:%d/%m/%Y}. {configuracion} Cada fuente se compara por separado (sin promediar). "
                   "Color fuerte = validado; color claro = orientativo (pendiente de validar, no es ahorro); "
                   "tachado = rechazada. Azul = captura: Cantidad y P.U. aquí; P.U. de referencia y estado en "
                   "'Por fuente'. Todo lo demás se recalcula.", f_sub)
    for c, texto in enumerate(cab_ini):
        ws.merge_range(3, c, 4, c, texto, f_cab)
    for j, clave in enumerate(FUENTES):
        c0 = len(cab_ini) + j * NF
        ws.merge_range(3, c0, 3, c0 + NF - 1, NOMBRE_CORTO[clave], f_cab_fuente[clave])
        for k, t in enumerate(sub_fuente):
            ws.write(4, c0 + k, t, f_cab_fuente[clave])
    c_fin = len(cab_ini) + NF * len(FUENTES)
    for k, texto in enumerate(cab_fin):
        ws.merge_range(3, c_fin + k, 4, c_fin + k, texto, f_cab)
    ws.set_row(4, 28)

    fila0 = 5
    total_importe = 0.0
    total_ahorro = 0.0
    fila_resumen = {}   # partida -> número de fila de Excel (1-based) en Resumen
    for i, f in enumerate(filas):
        r = fila0 + i
        R = r + 1
        fila_resumen[str(f.get("Partida"))] = R
        evs = f.get("_evaluaciones") or {}
        fin = f.get("_final") or {}
        cantidad = _num(f.get("Cantidad"))
        precio = _num(f.get("Precio cotizado"))
        importe = (cantidad or 0) * (precio or 0)
        total_importe += importe

        ws.write(r, 0, f.get("Partida"), f_centro)
        ws.write(r, 1, f.get("Concepto"), f_txt)
        ws.write(r, 2, f.get("Unidad"), f_centro)
        ws.write_number(r, 3, cantidad or 0, f_input_num)
        ws.write_number(r, 4, precio or 0, f_input_mon)
        ws.write_formula(r, 5, f"=D{R}*E{R}", f_mon, importe)

        celdas_dict, celdas_ref, celdas_est = [], [], []
        for j, clave in enumerate(FUENTES):
            c0 = len(cab_ini) + j * NF
            ev = evs.get(clave) or {}
            ref = _num(ev.get("precio_referencia"))
            estado = ev.get("estado", v.SIN_DATO)
            pf = lambda k: f"'Por fuente'!{xl_rowcol_to_cell(4 + i, 4 + j * PF + k)}"
            col = lambda k: xl_col_to_name(c0 + k)
            ws.write_formula(r, c0, f'=IF({pf(0)}="","",{pf(0)})',
                             f_mon, ref if ref is not None else "")
            excluida = estado in (v.RECHAZADA, v.SIN_DATO, v.NO_COMPARABLE)
            ws.write_formula(
                r, c0 + 1, f'=IF(OR({col(0)}{R}="",{pf(1)}="RECHAZADA",{pf(1)}="SIN DATO",{pf(1)}="NO COMPARABLE"),"",E{R}-{col(0)}{R})',
                f_dif_mon, (precio - ref) if (ref and precio is not None and not excluida) else "")
            ws.write_formula(
                r, c0 + 2, f'=IF({col(1)}{R}="","",E{R}/{col(0)}{R}-1)',
                f_pct, (precio / ref - 1) if (ref and precio is not None and not excluida) else "")
            ws.write_formula(r, c0 + 3, f"={pf(3)}", f_centro, v.dictamen_texto(ev))
            colorear_por_dictamen(ws, f"{col(0)}{R}:{col(3)}{R}", f"${col(3)}{R}")
            ws.write_formula(r, c0 + 4, f"={pf(4)}", f_centro, v.confiabilidad_texto(ev))
            celdas_dict.append(f"{col(3)}{R}")
            celdas_ref.append(f"{col(0)}{R}")
            celdas_est.append(pf(1))

        # --- Semáforo final, referencia y respaldo: fórmulas sobre los
        # dictámenes (que a su vez salen de los estados de 'Por fuente').
        es_val = [f'LEFT({d},8)="Validado"' for d in celdas_dict]
        clase = [f"MID({d},12,20)" for d in celdas_dict]
        n_val = "+".join(f"({x})" for x in es_val)
        n_ori = "+".join(f'(LEFT({d},11)="Orientativo")' for d in celdas_dict)
        primera = '""'
        for x, c in reversed(list(zip(es_val, clase))):
            primera = f"IF({x},{c},{primera})"
        iguales = ",".join(f"OR(NOT({x}),{c}={primera})" for x, c in zip(es_val, clase))
        nc = "+".join(f'(RIGHT({d},4)="alto")' for d in celdas_dict)
        ne = "+".join(f'(RIGHT({d},9)="en precio")' for d in celdas_dict)
        nb = "+".join(f'(RIGHT({d},4)="bajo")' for d in celdas_dict)
        nt = f"(({nc})+({ne})+({nb}))"
        nv = "+".join(f'({e}="VALIDADA")' for e in celdas_est)
        sufijo = f'&" de "&{nt}&IF({nt}>1," filtros"," filtro")'
        def _pal(p):
            return f'IF(({nv})>0,"{p[0].upper() + p[1:]}","Posible {p}")'
        f_res = (f'=IF({nt}=0,"Sin datos",'
                 f'IF(AND(({nc})>({ne}),({nc})>({nb})),{_pal("alto")}&" en "&({nc}){sufijo},'
                 f'IF(AND(({ne})>({nc}),({ne})>({nb})),{_pal("en precio")}&" en "&({ne}){sufijo},'
                 f'IF(AND(({nb})>({nc}),({nb})>({ne})),{_pal("bajo")}&" en "&({nb}){sufijo},'
                 f'"No coinciden: "&MID(IF(({nc})>0,", "&({nc})&" alto","")&IF(({ne})>0,", "&({ne})&" en precio","")'
                 f'&IF(({nb})>0,", "&({nb})&" bajo",""),3,100)))))')
        rf = v.resultado_filtros(evs)
        texto_rf = rf["texto_plano"]
        ws.write_formula(r, c_fin, f_res, f_semaforo.get(rf["clave"] or "OTRO", f_semaforo["OTRO"]), texto_rf)
        for palabra, color_f, color_t in (("alto", ROJO_F, ROJO_T), ("bajo", VERDE_F, VERDE_T),
                                          ("en precio", AMBAR_F, AMBAR_T)):
            ws.conditional_format(f"{xl_col_to_name(c_fin)}{R}", {
                "type": "text", "criteria": "containing", "value": palabra + " en ",
                "format": wb.add_format({"bg_color": color_f, "font_color": color_t, "bold": True})})

        nombres = [v.NOMBRE_FUENTE[k] for k in FUENTES]
        f_ref_txt = '"—"'
        for est, nombre in reversed(list(zip(celdas_est, nombres))):
            f_ref_txt = f'IF({est}="{v.EQUIVALENCIA_PARCIAL}","{nombre}",{f_ref_txt})'
        for x, nombre in reversed(list(zip(es_val, nombres))):
            f_ref_txt = f'IF({x},"{nombre}",{f_ref_txt})'
        col_refn = xl_col_to_name(c_fin + 1)
        ws.write_formula(r, c_fin + 1, "=" + f_ref_txt, f_txt, fin.get("referencia_negociacion") or "—")
        respaldo = "VALIDADA" if rf["validados"] else ("POR VALIDAR" if rf["n"] else "—")
        ws.write_formula(
            r, c_fin + 2,
            f'=IF(({n_val})>0,"VALIDADA",IF({nt}>0,"POR VALIDAR","—"))',
            f_estado.get(v.VALIDADA if respaldo == "VALIDADA" else v.NO_CONCLUYENTE, f_centro), respaldo)
        f_pu = '""'
        for celda, nombre in reversed(list(zip(celdas_ref, nombres))):
            f_pu = f'IF({col_refn}{R}="{nombre}",{celda},{f_pu})'
        ref_neg = _num(fin.get("precio_negociacion"))
        ws.write_formula(r, c_fin + 3, "=" + f_pu, f_mon_neg, ref_neg if ref_neg is not None else "")
        col_neg = xl_col_to_name(c_fin + 3)
        col_resp = xl_col_to_name(c_fin + 2)
        pct_neg = (precio / ref_neg - 1) if (ref_neg and precio is not None) else ""
        ws.write_formula(r, c_fin + 4, f'=IF({col_neg}{R}="","",E{R}/{col_neg}{R}-1)', f_pct, pct_neg)
        dif = round((precio - ref_neg) * (cantidad or 0), 2) if (ref_neg and precio is not None) else ""
        ws.write_formula(r, c_fin + 5, f'=IF({col_neg}{R}="","",(E{R}-{col_neg}{R})*D{R})', f_dif_mon, dif)
        # Ahorro solo si la referencia está VALIDADA y el cotizado supera el ±5 %.
        ahorro = _num(fin.get("ahorro_potencial")) or 0.0
        total_ahorro += ahorro
        ws.write_formula(
            r, c_fin + 6,
            f'=IF(AND({col_resp}{R}="VALIDADA",{col_neg}{R}<>""),'
            f'IF(E{R}>{col_neg}{R},(E{R}-{col_neg}{R})*D{R},0),0)',
            f_mon, ahorro,
        )
        ws.write(r, c_fin + 7, fin.get("detalle") or "", f_txt)

    ultima = fila0 + len(filas)
    U = ultima
    ws.write(ultima, 1, "TOTAL", f_bold)
    ws.write_formula(ultima, 5, f"=SUM(F{fila0 + 1}:F{U})", f_mon_neg, total_importe)
    col_ah = xl_col_to_name(c_fin + 6)
    ws.write_formula(ultima, c_fin + 6, f"=SUM({col_ah}{fila0 + 1}:{col_ah}{U})", f_mon_neg, total_ahorro)
    ws.write(ultima + 1, 1, "Oportunidad validada sobre el importe", f_bold)
    ws.write_formula(ultima + 1, c_fin + 6, f'=IF(F{U + 1}=0,"",{col_ah}{U + 1}/F{U + 1})', f_pct_b,
                     (total_ahorro / total_importe) if total_importe else "")
    # Conteos que se recalculan con el resultado de los 4 filtros.
    col_res = xl_col_to_name(c_fin)
    rng = f"{col_res}{fila0 + 1}:{col_res}{U}"
    conteos = [
        ("Partidas altas (posible o validado)", f'=COUNTIF({rng},"*alto en*")',
         sum(1 for f in filas if v.resultado_filtros(f.get("_evaluaciones") or {})["clave"] == v.ALTO)),
        ("Partidas bajas", f'=COUNTIF({rng},"*bajo en*")',
         sum(1 for f in filas if v.resultado_filtros(f.get("_evaluaciones") or {})["clave"] == v.BAJO)),
        ("Filtros no coinciden", f'=COUNTIF({rng},"No coinciden*")',
         sum(1 for f in filas if v.resultado_filtros(f.get("_evaluaciones") or {})["clave"] == v.MIXTO)),
        ("Sin datos", f'=COUNTIF({rng},"Sin datos")',
         sum(1 for f in filas if not v.resultado_filtros(f.get("_evaluaciones") or {})["clave"])),
    ]
    for k, (etq, frm, val) in enumerate(conteos):
        ws.write(ultima + 3 + k, 1, etq, f_bold)
        ws.write_formula(ultima + 3 + k, 3, frm, f_centro, val)
    ws.write(ultima + 9, 1, "Las columnas «Nota» y la evidencia de texto corresponden a la consulta original y no se "
                            "recalculan con fórmulas; precios, diferencias, dictámenes, colores, resultado y conteos sí.",
             f_sub)

    anchos = [5, 44, 8, 10, 13, 15] + [12, 13, 9, 20, 14] * len(FUENTES) + [24, 16, 20, 14, 12, 16, 15, 44]
    for c, ancho in enumerate(anchos):
        ws.set_column(c, c, ancho)
    ws.freeze_panes(5, 2)
    ws.hide_gridlines(2)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)

    # ==================================================================
    # Hoja 2: Cantidades (aritmética, escenarios de volumen, alcances)
    # ==================================================================
    if revision_cantidades:
        wc = wb.add_worksheet("Cantidades")
        wc.merge_range(0, 0, 0, 7, "Revisión de cantidades y aritmética", f_titulo)
        wc.set_row(0, 26)
        wc.write(1, 0, "Diferencias SUJETAS A ACLARACIÓN con el generador de volúmenes del proveedor; "
                       "no son ahorros confirmados. Cantidad y P.U. vienen del Resumen (única captura); el importe "
                       "declarado (azul) es el de la cotización.", f_sub)
        cab_a = ["#", "Concepto", "Unidad", "Cantidad", "P.U.", "Importe declarado",
                 "Cantidad × P.U.", "Diferencia"]
        for c, t in enumerate(cab_a):
            wc.write(3, c, t, f_cab)
        r = 4
        for a in revision_cantidades["aritmetica"]:
            R = r + 1
            wc.write(r, 0, a["partida"], f_centro)
            wc.write(r, 1, a["concepto"], f_txt)
            wc.write(r, 2, a["unidad"], f_centro)
            Rr = fila_resumen.get(str(a["partida"]))
            if Rr:   # una sola captura: cantidad y P.U. vienen del Resumen
                wc.write_formula(r, 3, f"=Resumen!D{Rr}", fmt(num_format="#,##0.00"), a["cantidad"] or 0)
                wc.write_formula(r, 4, f"=Resumen!E{Rr}", f_mon, a["precio_unitario"] or 0)
            else:
                wc.write_number(r, 3, a["cantidad"] or 0, f_input_num)
                wc.write_number(r, 4, a["precio_unitario"] or 0, f_input_mon)
            if a["importe_declarado"] is not None:
                wc.write_number(r, 5, a["importe_declarado"], f_input_mon)
            else:
                wc.write_blank(r, 5, None, f_input_mon)
            wc.write_formula(r, 6, f"=D{R}*E{R}", f_mon, a["importe_calculado"] or 0)
            wc.write_formula(r, 7, f'=IF(F{R}="","",F{R}-G{R})', f_mon, a["diferencia"] if a["diferencia"] is not None else "")
            r += 1
        wc.write(r, 1, "TOTAL", f_bold)
        wc.write_formula(r, 5, f"=SUM(F5:F{r})", f_mon_neg,
                         sum(a["importe_declarado"] or 0 for a in revision_cantidades["aritmetica"]))
        wc.write_formula(r, 6, f"=SUM(G5:G{r})", f_mon_neg, revision_cantidades["total_calculado"])
        wc.write_formula(r, 7, f"=F{r + 1}-G{r + 1}", f_mon_neg,
                         sum(a["importe_declarado"] or 0 for a in revision_cantidades["aritmetica"])
                         - revision_cantidades["total_calculado"])
        r += 2
        for h in revision_cantidades["hallazgos"]:
            wc.merge_range(r, 0, r, 7, f"[{h['nivel']}] {h['tipo']} — partida {h['partidas']}",
                           fmt(bold=True, bg_color={"REVISAR": ROJO_F, "CONFIRMAR": AMBAR_F}.get(h["nivel"], VERDE_F)))
            r += 1
            wc.merge_range(r, 0, r, 7, h["detalle"], f_txt)
            wc.set_row(r, 32)
            r += 1
            if h["escenarios"]:
                wc.merge_range(r, 0, r, 2, "Escenario", f_cab)
                for c, t in zip(range(3, 8), ["Fórmula / base", "m² del escenario", "m² cotizados",
                                              "Diferencia m² (cotizado − escenario)",
                                              "Diferencia × P.U. cotizado (sujeto a aclaración, no es ahorro)"]):
                    wc.write(r, c, t, f_cab)
                r += 1
                idx = next((i for i, a in enumerate(revision_cantidades["aritmetica"])
                            if str(a["partida"]) == str(h["partidas"])), None)
                cant = revision_cantidades["aritmetica"][idx]["cantidad"] if idx is not None else 0
                cant = cant or 0
                # Fila de la partida en la tabla de aritmética de esta hoja
                # (cantidad en D, P.U. en E): una sola captura.
                fila_part = 5 + idx if idx is not None else None
                for e in h["escenarios"]:
                    R = r + 1
                    wc.merge_range(r, 0, r, 2, e["nombre"], f_txt)
                    wc.write(r, 3, f"{e['formula']} · {e.get('base', '')}", f_centro)
                    wc.write_number(r, 4, e["m2"], fmt(num_format="#,##0.00"))
                    if fila_part:
                        wc.write_formula(r, 5, f"=D{fila_part}", fmt(num_format="#,##0.00"), cant)
                        wc.write_formula(r, 7, f"=G{R}*E{fila_part}", f_dif_mon, e["importe"])
                    else:
                        wc.write_number(r, 5, cant, f_input_num)
                        wc.write_number(r, 7, e["importe"], f_dif_mon)
                    wc.write_formula(r, 6, f"=F{R}-E{R}", fmt(num_format="#,##0.00;-#,##0.00"), e["dif_m2"])
                    r += 1
            r += 1
        wc.merge_range(r, 0, r, 7, "Alcances e impuestos a confirmar por escrito", f_cab)
        r += 1
        for a in revision_cantidades["alcances_confirmar"]:
            wc.merge_range(r, 0, r, 7, "• " + a, f_txt)
            r += 1
        for c, ancho in enumerate([5, 46, 8, 16, 14, 17, 16, 22]):
            wc.set_column(c, c, ancho)
        wc.hide_gridlines(2)
        wc.set_landscape()
        wc.fit_to_pages(1, 0)

    # ==================================================================
    # Hoja: Mercado (referencias reales para comparar, orientativas)
    # ==================================================================
    if datos_mercado and datos_mercado.get("resumen"):
        wmk = wb.add_worksheet("Mercado")
        wmk.merge_range(0, 0, 0, 13, "Datos de mercado para comparar", f_titulo)
        wmk.set_row(0, 26)
        wmk.write(1, 0, "Precios reales publicados: NL = licitaciones de obra pública (mediana y rango p25–p75, "
                        "ajustado INPC); CDMX = tabulador oficial 2026. Orientativos: la equivalencia exacta no "
                        "está demostrada. Se excluyen materiales distintos y cifras de otra escala.", f_sub)
        # Cada fuente por separado: NL con sus percentiles p25/p75 de precios
        # contratados (no son mínimo ni máximo) y CDMX con su propio
        # concepto y especificación. No se mezclan en un solo rango.
        cab_r = ["#", "Concepto", "Unidad", "P.U. cotizado",
                 "NL · concepto más parecido", "NL · p25", "NL · mediana", "NL · p75", "NL · registros / periodo",
                 "Cotizado vs mediana NL", "CDMX · concepto más parecido", "CDMX · precio", "Cotizado vs CDMX",
                 "Dónde cae el cotizado"]
        for c, t in enumerate(cab_r):
            wmk.write(3, c, t, f_cab)
        wmk.set_row(3, 30)
        r = 4
        for m in datos_mercado["resumen"]:
            R = r + 1
            Rr = fila_resumen.get(str(m["partida"]))
            wmk.write(r, 0, m["partida"], f_centro)
            wmk.write(r, 1, m["concepto"], f_txt)
            wmk.write(r, 2, m["unidad"], f_centro)
            if Rr:
                wmk.write_formula(r, 3, f"=Resumen!E{Rr}", f_mon, m["precio_cotizado"])
            else:
                wmk.write_number(r, 3, m["precio_cotizado"], f_input_mon)
            wmk.write(r, 4, m.get("nl_concepto") or "—", f_txt)
            for c, campo in ((5, "nl_p25"), (6, "nl_mediana"), (7, "nl_p75"), (11, "cdmx_precio")):
                if m.get(campo) is not None:
                    wmk.write_number(r, c, m[campo], f_mon)
                else:
                    wmk.write_blank(r, c, None, f_mon)
            wmk.write(r, 8, (f"{m['nl_registros']} registros · {m.get('nl_periodo') or ''}"
                             if m.get("nl_registros") else "—"), f_txt)
            wmk.write_formula(r, 9, f'=IF(G{R}="","",D{R}/G{R}-1)', f_pct,
                              (m["precio_cotizado"] / m["nl_mediana"] - 1) if m.get("nl_mediana") else "")
            wmk.write(r, 10, m.get("cdmx_concepto") or "—", f_txt)
            wmk.write_formula(r, 12, f'=IF(L{R}="","",D{R}/L{R}-1)', f_pct,
                              (m["precio_cotizado"] / m["cdmx_precio"] - 1) if m.get("cdmx_precio") else "")
            wmk.write(r, 13, m["posicion"], f_txt)
            r += 1
        r += 1
        cab_d = ["#", "Fuente", "Concepto de referencia", "Relación", "Unidad", "Precio (mediana)",
                 "Rango p25", "Rango p75", "Registros", "Periodo / fecha", "Equivalencia / cotizado vs referencia"]
        for c, t in enumerate(cab_d):
            wmk.write(r, c, t, f_cab)
        r += 1
        for ref in datos_mercado["referencias"]:
            wmk.write(r, 0, ref["partida"], f_centro)
            wmk.write(r, 1, ref["fuente"], f_txt)
            wmk.write(r, 2, ref["concepto"], f_txt)
            wmk.write(r, 3, ref["relacion"], f_txt)
            wmk.write(r, 4, ref["unidad"], f_centro)
            wmk.write_number(r, 5, ref["precio"], f_mon)
            for c, campo in ((6, "rango_bajo"), (7, "rango_alto")):
                if ref.get(campo):
                    wmk.write_number(r, c, ref[campo], f_mon)
                else:
                    wmk.write_blank(r, c, None, f_mon)
            wmk.write_number(r, 8, ref["n_registros"], f_centro)
            wmk.write(r, 9, ref["fecha"], f_txt)
            wmk.write(r, 10, f"{ref['equivalencia']} · {ref['posicion']}", f_txt)
            r += 1
        for c, ancho in enumerate([5, 40, 10, 36, 40, 12, 12, 12, 22, 12, 40, 12, 12, 40]):
            wmk.set_column(c, c, ancho)
        wmk.hide_gridlines(2)
        wmk.set_landscape()
        wmk.fit_to_pages(1, 0)

    # ==================================================================
    # Hoja 3: Por fuente (semáforo por fuente calculado con fórmula)
    # ==================================================================
    # Plan de la hoja 'Inflación NL' (se escribe más abajo): renglones ordenados
    # por partida y contrato, para que el P.U. NL de 'Por fuente' sea una
    # fórmula = mediana de las medianas por contrato de esa hoja.
    plan_infl = []          # (partida_idx, f, r, contrato)
    for i_f, f in enumerate(filas):
        regs_f = (f.get("_evaluaciones") or {}).get("nl", {}).get("registros_todos") or []
        regs_f = sorted(regs_f, key=lambda r: (str(r.get("ocid") or r.get("licitacion")), r.get("fecha") or ""))
        for r in regs_f:
            plan_infl.append((i_f, f, r, str(r.get("ocid") or r.get("licitacion"))))
    celda_pu_nl = {}        # partida_idx -> celda con la mediana de medianas
    fila_ini = 5            # primera fila de datos (Excel, 1-based)
    for k, (i_f, f, r, contrato) in enumerate(plan_infl):
        if i_f not in celda_pu_nl:
            celda_pu_nl[i_f] = f"'Inflación NL'!T{fila_ini + k}"

    wf = wb.add_worksheet("Por fuente")
    cab_f = ["#", "Concepto", "Unidad", "P.U. cotizado"]
    for clave in FUENTES:
        n = NOMBRE_CORTO[clave]
        cab_f += [f"{n}\nP.U. ref.", f"{n}\nestado", f"{n}\n% vs ref.", f"{n}\ndictamen",
                  f"{n}\nconfiabilidad fuente", f"{n}\nfalta confirmar (consulta original)"]
    wf.merge_range(0, 0, 0, len(cab_f) - 1, "Comparación independiente por fuente", f_titulo)
    wf.set_row(0, 26)
    wf.write(1, 0, "Azul = captura. Cambia aquí el P.U. de referencia o el estado (VALIDADA, EQUIVALENCIA PARCIAL, "
                   "NO CONCLUYENTE, POR CONFIRMAR, RECHAZADA, SIN DATO) y el dictamen, el semáforo final, la "
                   "referencia y el respaldo del Resumen se recalculan. ±5 % = en mercado. El P.U. cotizado viene "
                   "del Resumen.", f_sub)
    for c, texto in enumerate(cab_f):
        wf.write(3, c, texto, f_cab_fuente.get(FUENTES[(c - 4) // PF], f_cab) if c >= 4 else f_cab)
    wf.set_row(3, 42)
    for i, f in enumerate(filas):
        r = 4 + i
        R = r + 1
        precio = _num(f.get("Precio cotizado"))
        wf.write(r, 0, f.get("Partida"), f_centro)
        wf.write(r, 1, f.get("Concepto"), f_txt)
        wf.write(r, 2, f.get("Unidad"), f_centro)
        wf.write_formula(r, 3, f"=Resumen!E{fila0 + 1 + i}", f_mon, precio or 0)
        evs = f.get("_evaluaciones") or {}
        for j, clave in enumerate(FUENTES):
            c0 = 4 + j * PF
            ev = evs.get(clave) or {}
            ref = _num(ev.get("precio_referencia"))
            estado = ev.get("estado", v.SIN_DATO)
            cr, ce = xl_col_to_name(c0), xl_col_to_name(c0 + 1)
            if ref is not None and clave == "nl" and i in celda_pu_nl:
                # Vinculado a la mediana de medianas por contrato ('Inflación NL').
                wf.write_formula(r, c0, f"=ROUND({celda_pu_nl[i]},2)", f_mon, ref)
            elif ref is not None:
                wf.write_number(r, c0, ref, f_input_mon)
            else:
                wf.write_blank(r, c0, None, f_input_mon)
            wf.write(r, c0 + 1, estado, f_estado.get(estado, f_centro))
            wf.data_validation(r, c0 + 1, r, c0 + 1, {
                "validate": "list",
                "source": [v.VALIDADA, v.EQUIVALENCIA_PARCIAL, v.NO_CONCLUYENTE, v.POR_CONFIRMAR,
                           v.NO_COMPARABLE, v.RECHAZADA, v.SIN_DATO],
            })
            pct = (precio / ref - 1) if (ref and precio is not None
                                        and estado not in (v.RECHAZADA, v.SIN_DATO, v.NO_COMPARABLE)) else ""
            wf.write_formula(r, c0 + 2,
                             f'=IF(OR({cr}{R}="",{ce}{R}="RECHAZADA",{ce}{R}="SIN DATO",{ce}{R}="NO COMPARABLE"),"",D{R}/{cr}{R}-1)',
                             f_pct, pct)
            clas = f'IF(D{R}>{cr}{R},"alto","bajo")'
            clas_v = f'IF(D{R}>{cr}{R},"Alto","Bajo")'
            wf.write_formula(
                r, c0 + 3,
                f'=IF(OR({cr}{R}="",{ce}{R}="SIN DATO"),"Sin dato",IF({ce}{R}="RECHAZADA","Rechazada",'
                f'IF({ce}{R}="NO COMPARABLE","No comparable",IF({ce}{R}="VALIDADA",{clas_v},"Posible "&{clas}))))',
                f_centro, v.dictamen_texto(ev))
            colorear_por_dictamen(wf, f"{cr}{R}:{xl_col_to_name(c0 + 3)}{R}", f"${xl_col_to_name(c0 + 3)}{R}")
            wf.write(r, c0 + 4, v.confiabilidad_fuente(ev), f_centro)
            wf.write(r, c0 + 5, v.falta_confirmar(ev, f.get("Concepto")), f_txt)
    anchos_f = [5, 44, 8, 13] + [12, 15, 10, 20, 13, 34] * len(FUENTES)
    for c, ancho in enumerate(anchos_f):
        wf.set_column(c, c, ancho)
    wf.freeze_panes(4, 4)
    wf.hide_gridlines(2)
    wf.set_landscape()
    wf.fit_to_pages(1, 0)

    # ==================================================================
    # Hoja 3: Evidencia
    # ==================================================================
    we = wb.add_worksheet("Evidencia")
    campos = list(v.evidencia({}, "").keys())
    cab_e = ["#", "Concepto cotizado", "Unidad", "Fuente", "P.U. referencia"] + campos
    we.merge_range(0, 0, 0, len(cab_e) - 1, "Evidencia de cada referencia", f_titulo)
    we.set_row(0, 26)
    we.write(1, 0, "TEXTO DE LA CONSULTA ORIGINAL (no se recalcula con fórmulas). "
                   "Documento, código, página, unidad, región, fechas, qué incluye el precio y ajuste por inflación "
                   "(precio original, índice, periodos, valores y factor). El ajuste por inflación no confirma "
                   "vigencia comercial ni equivalencia técnica.", f_sub)
    for c, texto in enumerate(cab_e):
        we.write(3, c, texto, f_cab)
    we.set_row(3, 30)
    f_factor, f_indice = fmt(num_format="0.0000"), fmt(num_format="0.000")
    numericos = {"Precio original", "Precio usado (actualizado)", "Valor base", "Valor final", "Factor"}
    r = 4
    for f in filas:
        for clave in FUENTES:
            ev = (f.get("_evaluaciones") or {}).get(clave) or {}
            we.write(r, 0, f.get("Partida"), f_centro)
            we.write(r, 1, f.get("Concepto"), f_txt)
            we.write(r, 2, f.get("Unidad"), f_centro)
            we.write(r, 3, v.NOMBRE_FUENTE[clave], f_txt)
            ref = _num(ev.get("precio_referencia"))
            if ref is not None:
                we.write_number(r, 4, ref, f_mon)
            else:
                we.write_blank(r, 4, None, f_mon)
            for k, (campo, valor) in enumerate(v.evidencia(ev, f.get("Concepto")).items()):
                c = 5 + k
                num = _num(valor) if campo in numericos else None
                if num is not None:
                    we.write_number(r, c, num, f_factor if campo == "Factor" else
                                    (f_indice if campo.startswith("Valor") else f_mon))
                elif campo == "Enlace" and valor:
                    we.write_url(r, c, str(valor)[:255], string=str(valor)[:255])
                else:
                    we.write(r, c, "" if valor is None else str(valor), f_txt)
            r += 1
    anchos_e = [5, 34, 8, 14, 13] + [14, 40, 18, 22, 36, 44, 10, 36, 34, 30, 10, 8, 12, 18, 13, 26, 14, 10, 14,
                                     10, 10, 13, 30, 40, 44]
    for c, ancho in enumerate(anchos_e[:len(cab_e)]):
        we.set_column(c, c, ancho)
    we.freeze_panes(4, 4)
    we.hide_gridlines(2)
    we.set_landscape()
    we.fit_to_pages(1, 0)

    # ==================================================================
    # Hoja: Inflación NL (todos los renglones usados, reproducible)
    # ==================================================================
    if plan_infl:
        wi = wb.add_worksheet("Inflación NL")
        cab_i = ["#", "Concepto cotizado", "Fila en 'Precios Contratados (real)'", "OCID", "Licitación",
                 "Dependencia", "Tipo", "Concepto del contrato", "Fecha", "Precio original", "Mes del índice",
                 "Índice base", "Cómo se obtuvo el índice base", "Mes final", "Índice final", "Factor (final ÷ base)",
                 "Precio actualizado", "Contrato (OCID)", "Mediana del contrato", "P.U. NL de la partida",
                 "Inflación acumulada del renglón"]
        wi.merge_range(0, 0, 0, len(cab_i) - 1, "Actualización por inflación de cada renglón de Nuevo León", f_titulo)
        wi.set_row(0, 26)
        wi.write(1, 0, "Método: (1) solo renglones técnicamente equivalentes a la partida (misma unidad, mismo objeto, "
                       "función, material y especificaciones compatibles); (2) a cada renglón se le aplica la inflación "
                       "ACUMULADA (compuesta) desde su mes hasta el último mes publicado: factor = índice final ÷ índice "
                       "base (columna U = factor − 1); (3) mediana de los renglones de cada contrato (OCID), porque la base no trae el número "
                       "de partida; (4) P.U. NL = mediana de las medianas por contrato: cada contrato pesa lo mismo. "
                       "Las columnas P a T son fórmulas: si cambias un precio o un índice, el P.U. de 'Por fuente' y "
                       "del Resumen se recalcula. Índice usado: " + _infl.INDICE_NOMBRE + ".", f_sub)
        wi.set_row(1, 58)
        for c, t in enumerate(cab_i):
            wi.write(3, c, t, f_cab)
        wi.set_row(3, 30)
        f_ind = fmt(num_format="0.000")
        f_fac = fmt(num_format="0.000000")
        f_pct_i = fmt(num_format="+0.00%;-0.00%", align="center")
        # bloques: partida -> lista de (fila_excel, contrato)
        bloques = {}
        for k, (i_f, f, r, contrato) in enumerate(plan_infl):
            bloques.setdefault(i_f, []).append((fila_ini + k, contrato))
        for k, (i_f, f, r, contrato) in enumerate(plan_infl):
            fila = fila_ini - 1 + k
            R = fila + 1
            valores = [f.get("Partida"), f.get("Concepto"), r.get("fila_hoja"), r.get("ocid"), r.get("licitacion"),
                       r.get("dependencia"), r.get("tipo"), r.get("concepto"), r.get("fecha")]
            for c, val in enumerate(valores):
                wi.write(fila, c, "" if val is None else val, f_txt if c in (1, 5, 7) else f_centro)
            wi.write_number(fila, 9, r["precio_original"], f_input_mon)
            wi.write(fila, 10, r.get("indice_mes"), f_centro)
            wi.write_number(fila, 11, r["indice_base"], f_ind)
            wi.write(fila, 12, r.get("metodo_indice"), f_txt)
            wi.write(fila, 13, r.get("indice_final_mes"), f_centro)
            wi.write_number(fila, 14, r["indice_final"], f_ind)
            wi.write_formula(fila, 15, f"=O{R}/L{R}", f_fac, r["factor"])
            wi.write_formula(fila, 16, f"=J{R}*P{R}", f_mon, r["precio_actualizado"])
            wi.write(fila, 17, contrato, f_centro)
            wi.write_formula(fila, 20, f"=P{R}-1", f_pct_i, r["factor"] - 1)
            # Mediana del contrato: solo en el primer renglón de cada contrato.
            filas_contrato = [fx for fx, cx in bloques[i_f] if cx == contrato]
            if R == filas_contrato[0]:
                vals = [rr["precio_actualizado"] for (ii, ff, rr, cc) in plan_infl if ii == i_f and cc == contrato]
                wi.write_formula(fila, 18, f"=MEDIAN(Q{filas_contrato[0]}:Q{filas_contrato[-1]})", f_mon,
                                 float(pd_median(vals)))
            # P.U. NL de la partida: mediana de las medianas por contrato.
            filas_partida = [fx for fx, _ in bloques[i_f]]
            if R == filas_partida[0]:
                ev_nl = (f.get("_evaluaciones") or {}).get("nl", {})
                wi.write_formula(fila, 19, f"=MEDIAN(S{filas_partida[0]}:S{filas_partida[-1]})", f_ref.get(
                    "O_" + (ev_nl.get("clasificacion") or ""), f_mon), ev_nl.get("precio_referencia") or 0)
        for c, ancho in enumerate([5, 34, 12, 30, 22, 30, 18, 50, 11, 13, 14, 10, 34, 12, 10, 12, 13, 30, 14, 14, 14]):
            wi.set_column(c, c, ancho)
        # Desglose año por año de la inflación acumulada (del renglón más
        # antiguo usado): el producto de los tramos es el factor.
        try:
            mes_viejo = min(str(r.get("fecha"))[:7] for (_i, _f, r, _c) in plan_infl)
            desg = _infl.desglose_acumulado(mes_viejo)
        except Exception:
            desg = None
        if desg and desg.get("tramos"):
            c0 = 22
            wi.merge_range(3, c0, 3, c0 + 5, f"Inflación acumulada año por año desde {desg['base']} "
                                             "(renglón más antiguo usado)", f_cab)
            for j, t in enumerate(["De", "A", "Índice inicial", "Índice final", "Inflación del tramo",
                                   "Inflación acumulada"]):
                wi.write(4, c0 + j, t, f_cab)
            from xlsxwriter.utility import xl_rowcol_to_cell as _celda
            primera = 5
            for j, t in enumerate(desg["tramos"]):
                fr = primera + j
                wi.write(fr, c0, t["de"], f_centro)
                wi.write(fr, c0 + 1, t["a"], f_centro)
                wi.write_number(fr, c0 + 2, t["valor_de"], f_ind)
                wi.write_number(fr, c0 + 3, t["valor_a"], f_ind)
                wi.write_formula(fr, c0 + 4, f"={_celda(fr, c0 + 3)}/{_celda(fr, c0 + 2)}-1", f_pct_i,
                                 t["valor_a"] / t["valor_de"] - 1)
                wi.write_formula(fr, c0 + 5, f"={_celda(fr, c0 + 3)}/{_celda(primera, c0 + 2, True, True)}-1", f_pct_i,
                                 t["valor_a"] / desg["valor_base"] - 1)
            wi.write(primera + len(desg["tramos"]), c0,
                     "La inflación acumulada se compone (se multiplica año con año); no es la suma de los porcentajes.",
                     f_sub)
            for j, ancho in enumerate([16, 16, 12, 12, 14, 14]):
                wi.set_column(c0 + j, c0 + j, ancho)
        wi.freeze_panes(4, 2)
        wi.hide_gridlines(2)
        wi.set_landscape()
        wi.fit_to_pages(1, 0)

    # ==================================================================
    # Hoja: Dónde negociar (prioridades por dinero en juego)
    # ==================================================================
    if recomendacion and recomendacion.get("items"):
        wn = wb.add_worksheet("Dónde negociar")
        cab_n = ["Prioridad", "#", "Concepto", "Importe cotizado", "% del total", "Qué hacer",
                 "En juego por precio (mín.)", "En juego por precio (máx.)", "Cantidad por aclarar (máx.)",
                 "Filtros que dicen alto", "Otros filtros", "Explicación"]
        wn.merge_range(0, 0, 0, len(cab_n) - 1, "Dónde enfocarte para negociar", f_titulo)
        wn.set_row(0, 26)
        wn.merge_range(1, 0, 1, len(cab_n) - 1, recomendacion.get("resumen", "") + " Orden: dinero en juego = "
                       "(P.U. cotizado − P.U. de referencia) × cantidad, por separado para cada filtro (rango, sin "
                       "promediar), más las cantidades por aclarar. Con referencias orientativas el monto es un "
                       "máximo posible, no un ahorro. Texto generado al exportar (no se recalcula).", f_sub)
        wn.set_row(1, 44)
        for c, t in enumerate(cab_n):
            wn.write(3, c, t, f_cab)
        wn.set_row(3, 30)
        for k, it in enumerate(recomendacion["items"]):
            fr = 4 + k
            wn.write(fr, 0, k + 1, f_centro)
            wn.write(fr, 1, it.get("partida"), f_centro)
            wn.write(fr, 2, it.get("concepto"), f_txt)
            wn.write_number(fr, 3, it.get("importe") or 0, f_mon)
            wn.write_number(fr, 4, (it.get("peso_pct") or 0) / 100, fmt(num_format="0%", align="center"))
            wn.write(fr, 5, it.get("nivel"), f_txt)
            precio_ok = it.get("tipo") in ("precio", "precio_dudoso")
            wn.write_number(fr, 6, it.get("monto_min") or 0 if precio_ok else 0, f_mon)
            wn.write_number(fr, 7, it.get("monto_max") or 0 if precio_ok else 0, f_mon)
            wn.write_number(fr, 8, (it.get("cantidad_por_aclarar") or {}).get("max", 0), f_mon)
            wn.write(fr, 9, "; ".join(f"{c['fuente']} ${c['precio_ref']:,.2f} ({c['pct']:+.1f} %"
                                      f"{'' if c['validada'] else ', orientativa'})" for c in it.get("caros", [])), f_txt)
            wn.write(fr, 10, "; ".join(f"{o['fuente']} ${o['precio_ref']:,.2f} ({o['pct']:+.1f} %)"
                                       for o in it.get("otros", [])), f_txt)
            wn.write(fr, 11, it.get("texto"), f_txt)
        if recomendacion.get("texto_ia"):
            fr = 5 + len(recomendacion["items"])
            wn.write(fr, 0, "Recomendación de la IA", f_bold)
            wn.merge_range(fr, 1, fr, len(cab_n) - 1, recomendacion["texto_ia"], f_txt)
            wn.set_row(fr, 60)
        for c, ancho in enumerate([9, 5, 44, 15, 9, 26, 15, 15, 15, 34, 34, 80]):
            wn.set_column(c, c, ancho)
        wn.freeze_panes(4, 3)
        wn.hide_gridlines(2)
        wn.set_landscape()
        wn.fit_to_pages(1, 0)

    # ==================================================================
    # Hoja 4: Metodología
    # ==================================================================
    wm = wb.add_worksheet("Metodología")
    wm.set_column(0, 0, 26)
    wm.set_column(1, 1, 100)
    wm.merge_range(0, 0, 0, 1, "Metodología de validación", f_titulo)
    wm.set_row(0, 26)
    import version as _version
    reglas = [
        ("Versión de la app", f"{_version.VERSION} ({_version.FECHA}): {_version.DESCRIPCION}."),
        ("RECHAZA (IA)", "La referencia se excluye del precio de negociación, de las diferencias y del semáforo. Queda solo como evidencia."),
        ("NO_SEGURO (IA)", "Se conserva como orientativa (NO CONCLUYENTE): se muestra su precio y su %, pero no decide el semáforo."),
        ("CONFIRMA (IA)", "No basta. VALIDADA exige además: especificaciones de la referencia (sección, espesor, f'c, "
                          f"calibre) declaradas por el proveedor y coincidentes, y precio vigente (máx. {v.VIGENCIA_MESES} meses). "
                          "Si falta algo: EQUIVALENCIA PARCIAL (orientativa)."),
        ("Filtros (no validan)", f"Escala: una referencia a más de {v.FACTOR_ESCALA:.0f}× del cotizado se descarta como otra unidad. "
                                 "El ajuste por INPC actualiza la inflación pero no demuestra vigencia."),
        ("Material / especificación distinta", "Se RECHAZA aunque no haya IA (p. ej. columnas metálicas vs. castillos de concreto)."),
        ("Sin confirmar", "Sin CONFIRMA de la IA la referencia queda POR CONFIRMAR (orientativa)."),
        ("Diferencia vs. ahorro", "Contra una referencia VALIDADA la diferencia es una oportunidad potencial (el ahorro real es lo que se negocie); contra una orientativa es "
                                  "'diferencia contra referencia, pendiente de validar'."),
        ("Precio web", "Gemini con Google Search; si no hay cuota, el respaldo busca cada partida por material, trabajo y "
                       "unidad, abre las páginas y toma el precio solo del renglón donde aparecen juntos el concepto, la "
                       "unidad y el precio. Es orientativo (confiabilidad media o baja); solo un precio de Gemini con "
                       "fuente y frase verificadas puede quedar VALIDADO."),
        ("Sin promedios", "Cada fuente se compara por separado. El P.U. de negociación sale de UNA referencia validada, con prioridad: "
                          "Histórico Ragasa > Nuevo León > CDMX > IA internet."),
        ("Resultado de los 4 filtros", "Cuenta cuántos filtros dicen alto (arriba de la referencia) o bajo (igual o abajo); gana la mayoría y, "
                                       "si empatan, 'No coinciden' con el conteo de cada dictamen. No promedia precios. 'Respaldo' dice si "
                                       "hay al menos una referencia validada."),
        ("Nuevo León", "Mediana de las medianas por contrato (OCID): cada contrato pesa lo mismo. Dentro de cada "
                       "contrato solo se agrupan renglones técnicamente equivalentes a la partida (mismo objeto, "
                       "función, material, unidad y especificaciones compatibles). Cada renglón se actualiza con el "
                       "INPC de su mes. Detalle y fórmulas en 'Inflación NL'; el P.U. de 'Por fuente' y del Resumen "
                       "está vinculado a esa hoja."),
        ("Inflación acumulada", "Los precios con fecha (contratos de Nuevo León y compras del histórico) se llevan a hoy con "
                                "la inflación ACUMULADA (compuesta): factor = índice del último mes publicado ÷ índice "
                                "del mes del precio. Índice: " + _infl.INDICE_NOMBRE + ". El desglose año por año está "
                                "en 'Inflación NL'. El tabulador CDMX es edición 2026 y no se actualiza. Un índice "
                                "general de precios no sigue por fuerza el costo de la construcción: por eso se avisa "
                                "cuando Nuevo León queda por debajo de CDMX."),
        ("Nuevo León: fecha del precio", "La columna 'fecha' de la base es el día en que el registro se publicó en datos "
                                         "abiertos (decenas de miles de renglones comparten el mismo día). El año real "
                                         "del precio es el de la licitación, que viene en su número (…-E126-2016). "
                                         "Cuando es anterior al de publicación, la inflación se acumula desde julio de "
                                         "ese año (mitad del año: el mes del concurso no está en la base)."),
        ("Coherencia NL vs CDMX", "Si la referencia de Nuevo León queda más de 10 % por debajo de la de CDMX se avisa: "
                                  "son contratos de obra pública de gran volumen y la especificación puede diferir; se "
                                  "usa como piso de negociación, no como precio objetivo."),
        ("Dónde negociar", "Las partidas se ordenan por dinero en juego = (P.U. cotizado − P.U. de referencia) × cantidad, "
                           "por separado para cada filtro (rango, sin promediar), más las cantidades por aclarar."),
        ("CDMX", "Tabulador General de Precios Unitarios del Gobierno de la CDMX."),
        ("Histórico Ragasa", "Cotizaciones guardadas previamente en el histórico interno (Google Sheets)."),
    ]
    wm.write(3, 0, "Regla", f_cab)
    wm.write(3, 1, "Aplicación", f_cab)
    for i, (regla, texto) in enumerate(reglas):
        wm.write(4 + i, 0, regla, f_bold)
        wm.write(4 + i, 1, texto, f_txt)
        wm.set_row(4 + i, 30)
    wm.hide_gridlines(2)

    wb.close()
    return buffer.getvalue()
