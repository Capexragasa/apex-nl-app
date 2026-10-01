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

import validacion_referencias as v

FUENTES = ("historico", "nl", "cdmx", "ia")
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
}


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
    if precio > ref * 1.05:
        return "ALTO"
    if precio < ref * 0.95:
        return "BAJO"
    return "EN MERCADO"


def generar_excel(filas: list[dict], proveedor: str = "", proyecto: str = "",
                  revision_cantidades: dict | None = None, configuracion: str = "",
                  datos_mercado: dict | None = None) -> bytes:
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
        if estado in (v.NO_CONCLUYENTE, v.POR_CONFIRMAR):
            return f_ref["ORIENTATIVA"]
        if estado == v.RECHAZADA:
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

    hoy = _dt.date.today()

    # ==================================================================
    # Hoja 1: Resumen
    # ==================================================================
    ws = wb.add_worksheet("Resumen")
    cab_ini = ["#", "Concepto", "Unidad", "Cantidad", "P.U. cotizado", "Importe cotizado"]
    cab_fin = ["Semáforo final", "Referencia", "Respaldo", "P.U. de referencia",
               "% vs referencia", "Diferencia contra referencia", "Ahorro respaldado", "Nota"]
    n_cols = len(cab_ini) + 3 * len(FUENTES) + len(cab_fin)
    ws.merge_range(0, 0, 0, n_cols - 1, "Revisión de cotización CAPEX", f_titulo)
    ws.set_row(0, 26)
    ws.write(1, 0, f"Proveedor: {proveedor or '—'} · Proyecto: {proyecto or '—'} · "
                   f"Generado {hoy:%d/%m/%Y}. {configuracion} Cada fuente se compara por separado (sin promediar). "
                   "P.U. de color = VALIDADA; gris cursiva = orientativa; tachado = rechazada. P.U. y estado "
                   "vienen de la hoja 'Por fuente': corrígelos ahí y todo se recalcula.", f_sub)
    # Encabezados en dos niveles: fuente arriba, P.U./estado/% abajo.
    for c, texto in enumerate(cab_ini):
        ws.merge_range(3, c, 4, c, texto, f_cab)
    for j, clave in enumerate(FUENTES):
        c0 = len(cab_ini) + j * 3
        ws.merge_range(3, c0, 3, c0 + 2, NOMBRE_CORTO[clave], f_cab_fuente[clave])
        ws.write(4, c0, "P.U. ref.", f_cab_fuente[clave])
        ws.write(4, c0 + 1, "Estado", f_cab_fuente[clave])
        ws.write(4, c0 + 2, "% vs cotizado", f_cab_fuente[clave])
    c_fin = len(cab_ini) + 3 * len(FUENTES)
    for k, texto in enumerate(cab_fin):
        ws.merge_range(3, c_fin + k, 4, c_fin + k, texto, f_cab)
    ws.set_row(4, 28)

    fila0 = 5
    total_importe = 0.0
    total_ahorro = 0.0
    for i, f in enumerate(filas):
        r = fila0 + i
        R = r + 1  # número de fila de Excel
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

        for j, clave in enumerate(FUENTES):
            c0 = len(cab_ini) + j * 3
            ev = evs.get(clave) or {}
            ref = _num(ev.get("precio_referencia"))
            estado = ev.get("estado", v.SIN_DATO)
            col_ref = xl_col_to_name(c0)
            col_est = xl_col_to_name(c0 + 1)
            # Vinculado a "Por fuente" (misma partida, misma fuente).
            origen_ref = f"'Por fuente'!{xl_rowcol_to_cell(4 + i, 4 + j * 4)}"
            origen_est = f"'Por fuente'!{xl_rowcol_to_cell(4 + i, 5 + j * 4)}"
            ws.write_formula(r, c0, f'=IF({origen_ref}="","",{origen_ref})',
                             estilo_ref(ev) if ref is not None else f_mon,
                             ref if ref is not None else "")
            ws.write_formula(r, c0 + 1, f"={origen_est}", f_estado.get(estado, f_centro), estado)
            pct = (precio / ref - 1) if (ref and precio is not None
                                        and estado not in (v.RECHAZADA, v.SIN_DATO)) else ""
            ws.write_formula(
                r, c0 + 2,
                f'=IF(OR({col_ref}{R}="",{col_est}{R}="RECHAZADA",{col_est}{R}="SIN DATO"),"",E{R}/{col_ref}{R}-1)',
                f_pct, pct,
            )

        semaforo = fin.get("semaforo") or v.SIN_VALIDADA
        ws.write(r, c_fin, semaforo, f_semaforo.get(semaforo, f_semaforo["OTRO"]))
        ws.write(r, c_fin + 1, fin.get("referencia_negociacion") or "—", f_txt)
        respaldo = fin.get("respaldo") or "—"
        ws.write(r, c_fin + 2, respaldo, f_estado.get(v.VALIDADA if respaldo == "VALIDADA" else v.NO_CONCLUYENTE, f_centro))
        ref_neg = _num(fin.get("precio_negociacion"))
        clave_neg = next((k for k, n in v.NOMBRE_FUENTE.items()
                          if n == fin.get("referencia_negociacion")), None)
        if clave_neg:
            celda_ref = xl_rowcol_to_cell(r, len(cab_ini) + FUENTES.index(clave_neg) * 3)
            ws.write_formula(r, c_fin + 3, f"={celda_ref}", f_mon_neg, ref_neg or "")
        else:
            ws.write_blank(r, c_fin + 3, None, f_mon_neg)
        col_neg = xl_col_to_name(c_fin + 3)
        col_resp = xl_col_to_name(c_fin + 2)
        pct_neg = (precio / ref_neg - 1) if (ref_neg and precio is not None) else ""
        ws.write_formula(r, c_fin + 4, f'=IF({col_neg}{R}="","",E{R}/{col_neg}{R}-1)', f_pct, pct_neg)
        dif = round((precio - ref_neg) * (cantidad or 0), 2) if (ref_neg and precio is not None) else ""
        ws.write_formula(r, c_fin + 5, f'=IF({col_neg}{R}="","",(E{R}-{col_neg}{R})*D{R})', f_mon, dif)
        # Ahorro solo si la referencia está VALIDADA y el cotizado supera el ±5 %.
        ahorro = _num(fin.get("ahorro_potencial")) or 0.0
        total_ahorro += ahorro
        ws.write_formula(
            r, c_fin + 6,
            f'=IF(AND({col_resp}{R}="VALIDADA",{col_neg}{R}<>""),'
            f'IF(E{R}>{col_neg}{R}*1.05,(E{R}-{col_neg}{R})*D{R},0),0)',
            f_mon, ahorro,
        )
        ws.write(r, c_fin + 7, fin.get("detalle") or "", f_txt)

    ultima = fila0 + len(filas)  # fila (0-based) del total
    U = ultima  # última fila de datos en Excel = ultima (1-based)
    ws.write(ultima, 1, "TOTAL", f_bold)
    ws.write_formula(ultima, 5, f"=SUM(F{fila0 + 1}:F{U})", f_mon_neg, total_importe)
    col_ah = xl_col_to_name(c_fin + 6)
    ws.write_formula(ultima, c_fin + 6, f"=SUM({col_ah}{fila0 + 1}:{col_ah}{U})", f_mon_neg, total_ahorro)
    ws.write(ultima + 1, 1, "Ahorro respaldado sobre el importe", f_bold)
    ws.write_formula(ultima + 1, c_fin + 6, f'=IF(F{U + 1}=0,"",{col_ah}{U + 1}/F{U + 1})', f_pct_b,
                     (total_ahorro / total_importe) if total_importe else "")

    anchos = [5, 44, 8, 10, 13, 15] + [12, 20, 11] * len(FUENTES) + [24, 14, 20, 14, 12, 16, 15, 44]
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
                       "no son ahorros confirmados. Azul = dato de la cotización.", f_sub)
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
                for c, t in zip(range(3, 8), ["Fórmula", "m² esperados", "m² cotizados",
                                              "Diferencia m²", "Importe sujeto a aclaración"]):
                    wc.write(r, c, t, f_cab)
                r += 1
                cant = next((a["cantidad"] for a in revision_cantidades["aritmetica"]
                             if str(a["partida"]) == str(h["partidas"])), 0) or 0
                pu = next((a["precio_unitario"] for a in revision_cantidades["aritmetica"]
                           if str(a["partida"]) == str(h["partidas"])), 0) or 0
                for e in h["escenarios"]:
                    R = r + 1
                    wc.merge_range(r, 0, r, 2, e["nombre"], f_txt)
                    wc.write(r, 3, e["formula"], f_centro)
                    wc.write_number(r, 4, e["m2"], fmt(num_format="#,##0.00"))
                    wc.write_number(r, 5, cant, f_input_num)
                    wc.write_formula(r, 6, f"=F{R}-E{R}", fmt(num_format="#,##0.00"), e["dif_m2"])
                    wc.write_formula(r, 7, f"=MAX(0,G{R})*{pu}", f_mon, e["importe"])
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
        wmk.merge_range(0, 0, 0, 10, "Datos de mercado para comparar", f_titulo)
        wmk.set_row(0, 26)
        wmk.write(1, 0, "Precios reales publicados: NL = licitaciones de obra pública (mediana y rango p25–p75, "
                        "ajustado INPC); CDMX = tabulador oficial 2026. Orientativos: la equivalencia exacta no "
                        "está demostrada. Se excluyen materiales distintos y cifras de otra escala.", f_sub)
        cab_r = ["#", "Concepto", "Unidad", "P.U. cotizado", "Referencias", "Tipo", "Mínimo",
                 "Mediana de referencias", "Máximo", "Cotizado vs mediana", "Dónde cae el cotizado"]
        for c, t in enumerate(cab_r):
            wmk.write(3, c, t, f_cab)
        r = 4
        for m in datos_mercado["resumen"]:
            R = r + 1
            wmk.write(r, 0, m["partida"], f_centro)
            wmk.write(r, 1, m["concepto"], f_txt)
            wmk.write(r, 2, m["unidad"], f_centro)
            wmk.write_number(r, 3, m["precio_cotizado"], f_input_mon)
            wmk.write_number(r, 4, m["referencias"], f_centro)
            wmk.write(r, 5, m["tipo"], f_txt)
            for c, campo in ((6, "minimo"), (7, "mediana_referencias"), (8, "maximo")):
                if m[campo] is not None:
                    wmk.write_number(r, c, m[campo], f_mon)
                else:
                    wmk.write_blank(r, c, None, f_mon)
            pct = (m["precio_cotizado"] / m["mediana_referencias"] - 1) if m["mediana_referencias"] else ""
            wmk.write_formula(r, 9, f'=IF(H{R}="","",D{R}/H{R}-1)', f_pct, pct)
            wmk.write(r, 10, m["posicion"], f_txt)
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
        for c, ancho in enumerate([5, 44, 10, 36, 14, 22, 12, 16, 12, 18, 44]):
            wmk.set_column(c, c, ancho)
        wmk.hide_gridlines(2)
        wmk.set_landscape()
        wmk.fit_to_pages(1, 0)

    # ==================================================================
    # Hoja 3: Por fuente (semáforo por fuente calculado con fórmula)
    # ==================================================================
    wf = wb.add_worksheet("Por fuente")
    cab_f = ["#", "Concepto", "Unidad", "P.U. cotizado"]
    for clave in FUENTES:
        n = NOMBRE_CORTO[clave]
        cab_f += [f"{n}\nP.U. ref.", f"{n}\nestado", f"{n}\n% vs cotizado", f"{n}\nsemáforo"]
    wf.merge_range(0, 0, 0, len(cab_f) - 1, "Comparación independiente por fuente", f_titulo)
    wf.set_row(0, 26)
    wf.write(1, 0, "Semáforo por fuente solo para referencias VALIDADAS: ±5 % = EN MERCADO. "
                   "El P.U. de referencia se puede corregir aquí y las fórmulas se recalculan.", f_sub)
    for c, texto in enumerate(cab_f):
        wf.write(3, c, texto, f_cab_fuente.get(FUENTES[(c - 4) // 4], f_cab) if c >= 4 else f_cab)
    wf.set_row(3, 42)
    for i, f in enumerate(filas):
        r = 4 + i
        R = r + 1
        precio = _num(f.get("Precio cotizado"))
        wf.write(r, 0, f.get("Partida"), f_centro)
        wf.write(r, 1, f.get("Concepto"), f_txt)
        wf.write(r, 2, f.get("Unidad"), f_centro)
        wf.write_number(r, 3, precio or 0, f_input_mon)
        evs = f.get("_evaluaciones") or {}
        for j, clave in enumerate(FUENTES):
            c0 = 4 + j * 4
            ev = evs.get(clave) or {}
            ref = _num(ev.get("precio_referencia"))
            estado = ev.get("estado", v.SIN_DATO)
            cr, ce = xl_col_to_name(c0), xl_col_to_name(c0 + 1)
            if ref is not None:
                wf.write_number(r, c0, ref, f_input_mon)
            else:
                wf.write_blank(r, c0, None, f_input_mon)
            wf.write(r, c0 + 1, estado, f_estado.get(estado, f_centro))
            pct = (precio / ref - 1) if (ref and precio is not None
                                        and estado not in (v.RECHAZADA, v.SIN_DATO)) else ""
            wf.write_formula(r, c0 + 2,
                             f'=IF(OR({cr}{R}="",{ce}{R}="RECHAZADA",{ce}{R}="SIN DATO"),"",D{R}/{cr}{R}-1)',
                             f_pct, pct)
            sem = _clasif(precio, ref) if estado == v.VALIDADA else ""
            wf.write_formula(r, c0 + 3,
                             f'=IF({ce}{R}<>"VALIDADA","",IF(D{R}>{cr}{R}*1.05,"ALTO",'
                             f'IF(D{R}<{cr}{R}*0.95,"BAJO","EN MERCADO")))',
                             f_centro, sem)
            semaforo_cf(wf, f"{xl_col_to_name(c0 + 3)}{R}")
    anchos_f = [5, 44, 8, 13] + [12, 15, 11, 12] * len(FUENTES)
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
    cab_e = ["#", "Concepto cotizado", "Unidad", "Fuente", "Estado", "Motivo / verificaciones",
             "Concepto encontrado en la fuente", "Confianza texto", "P.U. referencia",
             "Revisión IA", "Evidencia web (frase o URL)"]
    we.merge_range(0, 0, 0, len(cab_e) - 1, "Evidencia de cada referencia", f_titulo)
    we.set_row(0, 26)
    we.write(1, 0, "Por qué cada referencia quedó VALIDADA, NO CONCLUYENTE, POR CONFIRMAR, "
                   "RECHAZADA o SIN DATO.", f_sub)
    for c, texto in enumerate(cab_e):
        we.write(3, c, texto, f_cab)
    r = 4
    for f in filas:
        for clave in FUENTES:
            ev = (f.get("_evaluaciones") or {}).get(clave) or {}
            estado = ev.get("estado", v.SIN_DATO)
            we.write(r, 0, f.get("Partida"), f_centro)
            we.write(r, 1, f.get("Concepto"), f_txt)
            we.write(r, 2, f.get("Unidad"), f_centro)
            we.write(r, 3, v.NOMBRE_FUENTE[clave], f_txt)
            we.write(r, 4, estado, f_estado.get(estado, f_centro))
            we.write(r, 5, ev.get("motivo") or "", f_txt)
            we.write(r, 6, ev.get("descripcion") or "", f_txt)
            we.write(r, 7, ev.get("confianza") or "", f_centro)
            ref = _num(ev.get("precio_referencia"))
            if ref is not None:
                we.write_number(r, 8, ref, f_mon)
            else:
                we.write_blank(r, 8, None, f_mon)
            we.write(r, 9, ev.get("revision_ia") or "", f_txt)
            we.write(r, 10, ev.get("evidencia") or "", f_txt)
            r += 1
    for c, ancho in enumerate([5, 38, 8, 16, 16, 44, 50, 11, 14, 44, 50]):
        we.set_column(c, c, ancho)
    we.freeze_panes(4, 4)
    we.hide_gridlines(2)
    we.set_landscape()
    we.fit_to_pages(1, 0)

    # ==================================================================
    # Hoja 4: Metodología
    # ==================================================================
    wm = wb.add_worksheet("Metodología")
    wm.set_column(0, 0, 26)
    wm.set_column(1, 1, 100)
    wm.merge_range(0, 0, 0, 1, "Metodología de validación", f_titulo)
    wm.set_row(0, 26)
    reglas = [
        ("RECHAZA (IA)", "La referencia se excluye del precio de negociación, de las diferencias y del semáforo. Queda solo como evidencia."),
        ("NO_SEGURO (IA)", "Se conserva como orientativa (NO CONCLUYENTE): se muestra su precio y su %, pero no decide el semáforo."),
        ("CONFIRMA (IA)", "No basta. VALIDADA exige además: especificaciones de la referencia (sección, espesor, f'c, "
                          f"calibre) declaradas por el proveedor y coincidentes, y precio vigente (máx. {v.VIGENCIA_MESES} meses). "
                          "Si falta algo: EQUIVALENCIA PARCIAL (orientativa)."),
        ("Filtros (no validan)", f"Escala: una referencia a más de {v.FACTOR_ESCALA:.0f}× del cotizado se descarta como otra unidad. "
                                 "El ajuste por INPC actualiza la inflación pero no demuestra vigencia."),
        ("Material / especificación distinta", "Se RECHAZA aunque no haya IA (p. ej. columnas metálicas vs. castillos de concreto)."),
        ("Sin confirmar", "Sin CONFIRMA de la IA la referencia queda POR CONFIRMAR (orientativa)."),
        ("Diferencia vs. ahorro", "Contra una referencia VALIDADA la diferencia es ahorro respaldado; contra una orientativa es "
                                  "'diferencia contra referencia, pendiente de validar'."),
        ("Precio web", "Debe aparecer en la misma frase que el concepto y la unidad. Un fragmento web sin validar es orientativo; "
                       "solo un precio de Gemini con fuente y frase verificadas puede quedar VALIDADO."),
        ("Sin promedios", "Cada fuente se compara por separado. El P.U. de negociación sale de UNA referencia validada, con prioridad: "
                          "Histórico Ragasa > Nuevo León > CDMX > IA internet."),
        ("Semáforo final", "Todas las validadas coinciden = ese semáforo. Si discrepan = MIXTO. Sin validadas pero con orientativas = "
                           "NO CONCLUYENTE. Sin nada = SIN DATOS SUFICIENTES. EN MERCADO = ±5 % de la referencia."),
        ("Nuevo León", "Mediana del tabulador homologado de licitaciones de NL, actualizada a hoy con INPC (INEGI)."),
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
