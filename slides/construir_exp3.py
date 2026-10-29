"""Diapositivas para entender el experimento 3 del examen parcial.

Uso:
    python EP/slides/construir_exp3.py

Sale Experimento3_EP.pptx y .pdf. Las cifras se leen de EP/results/exp3/.
"""
import sys
from pathlib import Path

import pandas as pd
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

aqui = Path(__file__).resolve().parent
ep = aqui.parent
res = ep / "results" / "exp3"
sys.path.insert(0, str(ep.parent / "pc2" / "slides"))
from comun import exportar, nueva_presentacion  # noqa: E402

NEGRO = RGBColor(0, 0, 0)

resumen = pd.read_csv(res / "resumen.csv", index_col=0)
por_caso = pd.read_csv(res / "cumplimiento_por_caso.csv", index_col=0)
A, B, C, D = (resumen.loc[k] for k in "ABCD")


def num(x, d=0):
    return f"{round(float(x) + 1e-9, d):.{d}f}".replace(".", ",")


def sn(v):
    return "sí" if v else "no"


prs = nueva_presentacion()


def slide_titulo(titulo):
    s = prs.slides.add_slide(prs.slide_layouts[5])
    t = s.shapes.title
    t.left, t.top, t.width, t.height = Inches(0.6), Inches(0.3), Inches(12.1), Inches(0.9)
    t.text_frame.text = titulo
    t.text_frame.paragraphs[0].runs[0].font.size = Pt(32)
    t.text_frame.paragraphs[0].alignment = 1
    return s


def bullets(s, items, x=0.6, y=1.4, w=12.1, h=5.6, size=20):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        nivel = 1 if it.startswith("  ") else 0
        p.text = ("- " if nivel else "• ") + it.strip()
        p.level = nivel
        p.space_after = Pt(8)
        for r in p.runs:
            r.font.size = Pt(size - 3 * nivel)
    return tb


def tabla(s, datos, x, y, anchos, alto=0.4, size=15):
    gf = s.shapes.add_table(len(datos), len(datos[0]), Inches(x), Inches(y), Inches(sum(anchos)),
                            Inches(alto * len(datos)))
    t = gf.table
    for j, a in enumerate(anchos):
        t.columns[j].width = Inches(a)
    for i, fila in enumerate(datos):
        t.rows[i].height = Inches(alto)
        for j, v in enumerate(fila):
            c = t.cell(i, j)
            c.text = str(v)
            c.margin_top = c.margin_bottom = Inches(0.03)
            for p in c.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(size)
    return t


def caja(s, x, y, w, h, texto, size=16):
    b = s.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    b.fill.background()
    b.line.color.rgb = NEGRO
    b.line.width = Pt(1.25)
    b.shadow.inherit = False
    tf = b.text_frame
    tf.word_wrap = True
    for i, linea in enumerate(texto.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = linea
        p.alignment = 2
        for r in p.runs:
            r.font.size = Pt(size)
            r.font.color.rgb = NEGRO
    return b


def flecha(s, x1, y1, x2, y2):
    from lxml import etree
    c = s.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = NEGRO
    c.line.width = Pt(1.5)
    ln = c.line._get_or_add_ln()
    e = etree.SubElement(ln, "{http://schemas.openxmlformats.org/drawingml/2006/main}tailEnd")
    e.set("type", "triangle")


def notas(s, texto):
    s.notes_slide.notes_text_frame.text = texto


# 1. portada
s = prs.slides.add_slide(prs.slide_layouts[0])
s.shapes.title.text = "Experimento 3: contratos de una herramienta"
s.placeholders[1].text = "Examen Parcial CC0F4 - Tópicos de Ciencia de la Computación I"
s.shapes.title.left, s.shapes.title.top, s.shapes.title.width, s.shapes.title.height = Inches(0.9), Inches(2.0), Inches(11.5), Inches(1.8)
s.placeholders[1].left, s.placeholders[1].top, s.placeholders[1].width, s.placeholders[1].height = Inches(0.9), Inches(4.1), Inches(11.5), Inches(1.0)

# 2. de que trata
s = slide_titulo("¿De qué trata el experimento?")
bullets(s, [
    "Los experimentos 1 y 2 buscan texto y generan respuestas. Aquí no interviene ningún LLM",
    "  Se prueba una herramienta que el LLM podría llamar: ocds.indicadores_proceso",
    "La herramienta puede recibir argumentos mal escritos y puede fallar. El experimento mide cómo reacciona",
    "Pregunta del experimento",
    "  Con los mismos 16 casos, ¿qué agregan los contratos y el reintento selectivo frente a ejecutar sin contratos o reintentar todo?",
    "Se comparan cuatro políticas de ejecución, de A (sin nada) a D (contratos y reintento selectivo)",
    "La medida es objetiva: cada caso tiene un resultado esperado fijado antes de correr",
])
notas(s, "Un contrato es una promesa verificable: qué argumentos acepta la herramienta y qué forma tiene lo que devuelve. "
         "Si alguien no la cumple, el error se detecta en el borde y no se propaga.")

# 3. la herramienta
s = slide_titulo("La herramienta: ocds.indicadores_proceso")
datos = [["Señal", "Regla", "Para qué sirve"],
         ["unico_postor", "n_postores == 1", "indicador de riesgo (Fazekas y Kocsis, 2020)"],
         ["adjudicado_sobre_vr", "monto adjudicado / valor referencial > 1", "sobrecosto"],
         ["equivalente_solo_en_objeto", "el objeto dice \"o equivalente\" y el ítem no", "la inconsistencia de Q15"]]
tabla(s, datos, 0.6, 1.4, [3.6, 4.8, 3.7], alto=0.55, size=15)
bullets(s, [
    "Recibe un ocid (y un ítem si el proceso tiene varios) y calcula las tres señales desde los datos OCDS",
    "No decide si un requisito está direccionado: eso sigue siendo del LLM y de la regla de la PC1",
    "Solo entrega hechos del registro, para contrastar lo que dice el modelo",
    "  Ejemplo con Q15: el LLM se contradice sobre \"o equivalente\" y la herramienta devuelve objeto sí, ítem no",
    "No tiene efectos: solo lee data/indice_procesos.jsonl",
], y=3.8, h=3.4, size=19)

# 4. ejemplo de salida
s = slide_titulo("Ejemplo de una llamada válida (proceso de Q15)")
caja(s, 0.6, 1.4, 3.7, 0.8, "Entrada\n{\"ocid\": \"ocds-dgv273-seacev3-2025-200121-10\"}", size=12)
flecha(s, 4.3, 1.8, 4.9, 1.8)
caja(s, 4.9, 1.4, 1.9, 0.8, "ocds.indicadores\n_proceso", size=14)
flecha(s, 6.8, 1.8, 7.4, 1.8)
caja(s, 7.4, 1.4, 5.3, 0.8, "ToolResult { status: \"ok\", data: {…} }", size=14)
datos = [["Campo de data", "Valor"],
         ["nomenclatura", "LP-ABR-1-2025-APN-1"],
         ["entidad", "MTC - AUTORIDAD PORTUARIA NACIONAL"],
         ["n_postores / unico_postor", "11 / false"],
         ["ratio_adjudicado_vr / adjudicado_sobre_vr", "0,7865 / false"],
         ["objeto_dice_equivalente / item_dice_equivalente", "true / false"],
         ["alertas", "[\"equivalente_solo_en_objeto\"]"]]
tabla(s, datos, 0.6, 2.6, [6.2, 5.9], alto=0.5, size=16)
bullets(s, [
    "El proceso no tuvo un solo postor ni se adjudicó sobre el valor referencial, pero sí tiene la alerta de Q15",
], y=6.3, h=0.9, size=18)

# 5. arquitectura de la herramienta
s = slide_titulo("Por dónde pasa una llamada")
xs = [0.6, 2.8, 5.0, 7.2, 9.4, 11.0]
anchos = [1.9, 1.9, 1.9, 1.9, 1.5, 1.7]
nombres = ["Entrada\n{ocid, tender_item_id}", "Contrato de entrada\n(Pydantic)", "Lectura del índice\n(aquí se inyectan fallos)",
           "Validación de dominio", "Cálculo de las señales", "Contrato de salida"]
for x, w, n in zip(xs, anchos, nombres):
    caja(s, x, 1.4, w, 1.2, n, size=13)
for i in range(5):
    flecha(s, xs[i] + anchos[i], 2.0, xs[i + 1], 2.0)
errores = {1: "validation", 2: "transient\ntimeout\nexecution", 3: "domain", 5: "output_contract"}
for i, e in errores.items():
    cx = xs[i] + anchos[i] / 2
    flecha(s, cx, 2.6, cx, 3.3)
    caja(s, xs[i] + 0.1, 3.3, anchos[i] - 0.2, 0.95 if "\n" in e else 0.55, e, size=13)
    flecha(s, cx, 4.25 if "\n" in e else 3.85, cx, 4.9)
caja(s, 0.6, 4.9, 12.1, 0.8, "ToolResult { status, data, error }\nok con los indicadores, o error con un ToolError { code, kind, retryable, message }", size=15)
bullets(s, [
    "El contrato de entrada y el de salida son los dos controles de la herramienta; los errores de abajo salen tipados",
    "Si el contrato de salida detecta una salida rota, no llega al LLM y no se reintenta",
], y=5.9, h=1.4, size=17)

# 6. contratos
s = slide_titulo("Los contratos")
datos = [["Parte", "Qué exige"],
         ["Entrada", "ocid con el patrón ocds-dgv273-seacev3-N (hasta dos sufijos); tender_item_id entero estricto ≥ 1 u omitido"],
         ["", "extra=\"forbid\": un argumento inventado es un error, no se ignora"],
         ["Dominio", "el ocid existe en el índice; el ítem pertenece al proceso; si hay varios ítems hay que indicar cuál"],
         ["Salida", "tipos y rangos (n_postores ≥ 0, ratio ≥ 0)"],
         ["", "coherencia: unico_postor coincide con n_postores, adjudicado_sobre_vr con el ratio, y alertas con los indicadores"]]
tabla(s, datos, 0.6, 1.4, [2.1, 10.0], alto=0.7, size=16)
bullets(s, [
    "Un argumento mal escrito falla en la entrada y no dentro del cálculo",
    "Una salida incoherente se bloquea antes de que el LLM la lea",
], y=5.8, h=1.4, size=19)

# 7. tipos de error
s = slide_titulo("Los tipos de error")
datos = [["kind", "Cuándo ocurre", "¿Se reintenta?"],
         ["validation", "la entrada no cumple el contrato", "no"],
         ["domain", "ocid o ítem que no existen, o falta elegir ítem", "no"],
         ["transient", "el índice no está disponible por un momento", "sí"],
         ["timeout", "un intento pasa de 100 ms", "sí, si queda deadline"],
         ["execution", "excepción interna no esperada", "no"],
         ["output_contract", "la salida no cumple el contrato", "no"],
         ["protocol", "herramienta desconocida", "no"]]
tabla(s, datos, 0.6, 1.4, [2.8, 6.4, 2.9], alto=0.52, size=16)
bullets(s, [
    "Cada error lleva retryable: el llamador sabe si tiene sentido volver a intentar",
], y=5.8, h=1.2, size=19)

# 8. cuando reintentar
s = slide_titulo("Cuándo reintentar y cuándo no")
bullets(s, [
    "La herramienta solo lee, así que repetir una llamada no duplica nada: reintentar es seguro",
    "Pero solo es útil si el error puede desaparecer solo",
    "  Sí: transient (el índice no responde un momento) y timeout (un intento pasa de 100 ms)",
    "  No: validation, domain, output_contract, execution y protocol",
    "  Reintentar una entrada inválida o un ocid inexistente da el mismo error otra vez",
    "  Una salida rota es un bug de la herramienta; repetirla da la misma salida rota",
    "Parámetros: hasta 3 intentos, espera de 20 y 40 ms entre intentos, deadline total de 300 ms",
    "Los fallos transitorios y lentos se inyectan de forma determinista, porque el índice es un archivo local y no falla solo",
], size=19)
notas(s, "Si la herramienta escribiera algo, reintentar sin cuidado podría duplicar el efecto. Aquí no pasa, y por eso el "
         "reintento ingenuo no rompe nada; solo gasta.")

# 9. las cuatro politicas
s = slide_titulo("Las cuatro políticas")
datos = [["Política", "Contratos y ToolError", "Timeout", "Reintento", "Intentos máx."],
         ["A  sin contratos", "no", "no", "no", "1"],
         ["B  contratos", "sí", "no", "no", "1"],
         ["C  reintenta todo", "sí", "100 ms por intento, deadline 300 ms", "cualquier error", "3"],
         ["D  reintento selectivo", "sí", "100 ms por intento, deadline 300 ms", "solo transient y timeout", "3"]]
tabla(s, datos, 0.6, 1.4, [3.0, 2.6, 3.4, 2.1, 1.0], alto=0.62, size=15)
bullets(s, [
    "A es la línea base: ejecuta lo que llegue, y los errores salen sin tipo",
    "B agrega los contratos y el ToolError tipado",
    "C agrega a B un timeout y reintenta cualquier error",
    "D agrega a B un timeout por intento y un deadline total, y reintenta solo lo que puede arreglarse solo",
], y=4.7, h=2.5, size=19)

# 10. los 16 casos
s = slide_titulo("Los 16 casos")
datos = [["Grupo", "Casos", "Resultado que pide la especificación"],
         ["Válidos (4)", "alerta de Q15, único postor, sobre el VR, ítem de un proceso de 6 ítems", "ok con la alerta esperada"],
         ["Entrada inválida (3)", "ocid mal formado, ítem como texto, argumento extra", "validation"],
         ["Error de dominio (3)", "ocid inexistente, ítem de otro proceso, falta elegir ítem (104 ítems)", "domain"],
         ["Transitorio (2)", "falla solo en el primer intento / falla siempre", "ok tras reintentar / transient"],
         ["Lento (2)", "tarda 150 ms solo la primera vez / siempre", "ok tras reintentar / timeout"],
         ["Error interno (1)", "una excepción dentro de la herramienta", "execution"],
         ["Salida mal formada (1)", "devuelve n_postores = -1", "output_contract"]]
tabla(s, datos, 0.6, 1.4, [2.9, 6.3, 2.9], alto=0.62, size=14)
bullets(s, [
    "El resultado esperado de cada caso está escrito en el código antes de correr (CASOS en exp3_herramienta.py)",
], y=6.5, h=0.8, size=17)

# 11. fijo
s = slide_titulo("Qué se mantiene igual")
bullets(s, [
    "Los mismos 16 casos, con los mismos argumentos y los mismos fallos inyectados, para las cuatro políticas",
    "El resultado esperado de cada caso, fijado antes de correr",
    "El índice de procesos, con su SHA-256 guardado en results/exp3/manifiesto.json",
    "Los tiempos: 3 intentos, esperas de 20 y 40 ms, 100 ms por intento y 300 ms de deadline",
    "No hay azar ni modelo: la misma entrada da siempre el mismo resultado",
    "Lo único que cambia entre políticas es lo que rodea a la herramienta, no la herramienta",
], size=20)

# 12. como se mide
s = slide_titulo("Cómo se mide")
bullets(s, [
    "Casos que cumplen la especificación",
    "  Un caso cumple si el status y el tipo de error coinciden con lo esperado (y, si es ok, las alertas)",
    "Entradas inválidas ejecutadas: entradas que el contrato habría rechazado y que la política dejó correr",
    "Salidas inválidas aceptadas: salidas que rompen el contrato de salida y llegaron como ok",
    "Intentos totales",
    "Reintentos inútiles",
    "  Intentos hechos después de un error que reintentando no se arregla",
    "Latencia media y máxima por caso",
], size=19)

# 13. ejemplos A vs B
s = slide_titulo("Ejemplo: la misma entrada bajo A y bajo B-D")
datos = [["Caso", "Política A (sin contratos)", "Políticas B, C y D"],
         ["Argumento extra\n{ocid, entidad: \"HOSPITAL GENERAL DE JAEN\"}", "ok: ignora el argumento en silencio",
          "validation: el argumento extra no existe"],
         ["Ítem como texto\ntender_item_id: \"21314522\"", "error genérico: busca el ítem como texto y no lo halla, como si no existiera",
          "validation: el tipo del argumento es incorrecto"],
         ["Salida mal formada\nn_postores = -1", "ok con n_postores = -1",
          "output_contract: se bloquea y no se reintenta"]]
tabla(s, datos, 0.6, 1.4, [4.3, 4.0, 3.8], alto=1.0, size=15)
bullets(s, [
    "Sin contrato, A acepta un dato falso y lo pasa al LLM como si fuera correcto",
    "El error de A en el ítem es engañoso: el ítem sí existe, lo que está mal es su tipo",
], y=5.7, h=1.6, size=19)

# 14. resultados
s = slide_titulo("Resultados")
datos = [["Política", "Cumplen", "Entradas inválidas ejecutadas", "Salidas inválidas aceptadas", "Intentos", "Reintentos inútiles",
          "Latencia media (ms)", "Latencia máx. (ms)"]]
for n, r in zip(["A sin contratos", "B contratos", "C reintenta todo", "D reintento selectivo"], [A, B, C, D]):
    datos.append([n, f"{int(r.cumplen_especificacion)}/16", int(r.entradas_invalidas_ejecutadas),
                  int(r.salidas_invalidas_aceptadas), int(r.intentos_totales), int(r.reintentos_inutiles),
                  num(r.latencia_media_ms), num(r.latencia_max_ms)])
tabla(s, datos, 0.4, 1.4, [2.7, 1.2, 1.7, 1.6, 1.1, 1.4, 1.5, 1.5], alto=0.6, size=14)
bullets(s, [
    "C y D cumplen los 16 casos; B cumple 14 y A cumple 6",
    "D logra lo mismo que C con 19 intentos en vez de 29 y sin reintentos inútiles",
    "El detalle de cada caso está en cumplimiento_por_caso.csv y las trazas de cada intento en trazas.jsonl",
], y=5.4, h=1.9, size=19)

# 15. grafico
s = slide_titulo("Resultados: gráfico")
cd = CategoryChartData()
cd.categories = ["A", "B", "C", "D"]
cd.add_series("Casos que cumplen (de 16)", [int(r.cumplen_especificacion) for r in (A, B, C, D)])
cd.add_series("Intentos totales", [int(r.intentos_totales) for r in (A, B, C, D)])
cd.add_series("Reintentos inútiles", [int(r.reintentos_inutiles) for r in (A, B, C, D)])
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.4), Inches(8.2), Inches(5.4), cd)
ch = gf.chart
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
pl = ch.plots[0]
pl.gap_width = 60
pl.has_data_labels = True
pl.data_labels.show_value = True
pl.data_labels.font.size = Pt(14)
ch.category_axis.tick_labels.font.size = Pt(14)
ch.value_axis.maximum_scale = 30
ch.value_axis.minimum_scale = 0
ch.value_axis.has_major_gridlines = False
bullets(s, [
    "Azul: casos que cumplen; rojo: intentos; verde: reintentos inútiles",
    "Cumplir sube con cada capa: 6, 14, 16, 16",
    "C gasta 29 intentos, 10 de ellos inútiles",
    "D cumple lo mismo con 19",
], x=9.1, y=1.8, w=3.8, h=4.5, size=18)

# 16. por caso
s = slide_titulo("Resultados: los casos donde alguna política falla")
fallan = por_caso[~por_caso.all(axis=1)]
datos = [["Caso", "A", "B", "C", "D"]]
for caso, fila in fallan.iterrows():
    datos.append([caso, sn(fila["A"]), sn(fila["B"]), sn(fila["C"]), sn(fila["D"])])
tabla(s, datos, 0.6, 1.3, [6.9, 1.2, 1.2, 1.2, 1.2], alto=0.4, size=14)
bullets(s, [
    f"Los otros {len(por_caso) - len(fallan)} casos (los 4 válidos, lento_una_vez y error_interno) los cumplen las cuatro políticas",
    "A cumple error_interno por coincidencia: todo error le sale como execution",
], y=6.0, h=1.4, size=17)

# 17. B
s = slide_titulo("Lo que B no resuelve")
bullets(s, [
    "B cumple 14 de 16. Los dos casos que no cumple son de infraestructura, no de entrada",
    "transitorio_una_vez: el índice falla en el primer intento",
    "  B devuelve el error transient y se rinde, aunque un segundo intento habría funcionado",
    "lento_siempre: la llamada tarda más de lo permitido",
    "  B no tiene timeout, así que espera hasta el final y devuelve ok cuando la especificación pide timeout",
    "Para esos dos casos hacen falta un timeout y un reintento",
    "  Eso es lo que agregan C y D",
], size=20)

# 18. C vs D
s = slide_titulo("C y D: mismo cumplimiento, distinto gasto")
datos = [["Caso", "Intentos en C", "Intentos en D"],
         ["dominio_ocid_inexistente", 3, 1],
         ["dominio_item_de_otro_proceso", 3, 1],
         ["dominio_falta_item_en_proceso_de_104_items", 3, 1],
         ["error_interno", 3, 1],
         ["salida_mal_formada", 3, 1]]
tabla(s, datos, 0.6, 1.4, [7.3, 2.4, 2.4], alto=0.5, size=16)
bullets(s, [
    f"C repite 2 veces más estos 5 casos: {int(C.reintentos_inutiles)} intentos inútiles",
    f"La latencia media pasa de {num(D.latencia_media_ms)} ms en D a {num(C.latencia_media_ms)} ms en C",
    "Aquí no rompe nada, porque los fallos son deterministas y la herramienta no escribe",
    "  Con una herramienta que escriba, o con un deadline más corto, el reintento ingenuo sí rompería",
    "D reintenta solo los casos con transient y timeout, que son los que pueden arreglarse",
], y=4.6, h=2.7, size=18)

# 19. pruebas
s = slide_titulo("Pruebas automáticas")
bullets(s, [
    "tests/test_herramienta.py tiene 22 pruebas (pytest); todas pasan (results/exp3/pytest.txt)",
    "Cubren",
    "  la ejecución válida y el contrato de salida",
    "  5 entradas inválidas y 3 errores de dominio sin reintento",
    "  transitorio una vez (se recupera), sin reintento (falla) y persistente (se corta a los 3 intentos)",
    "  lento una vez (timeout y luego ok) y lento siempre (respeta el deadline)",
    "  error interno y salida mal formada, sin reintento",
    "  herramienta desconocida, la función debe_reintentar, y que el contrato detecta alertas incoherentes",
    "Las pruebas fijan el comportamiento esperado; el experimento lo compara entre las cuatro políticas",
], size=19)

# 20. limitaciones
s = slide_titulo("Limitaciones")
bullets(s, [
    "Los fallos son simulados y deterministas: un transitorio real puede durar distinto y repetirse",
    "  Los tiempos solo prueban que el runtime respeta sus límites, no cómo se comporta el SEACE real",
    "16 casos elegidos por nosotros: cubren cada tipo de error una o dos veces, no su frecuencia real",
    "El cumplimiento depende de la especificación que escribimos; otra especificación daría otro conteo",
    "Los límites de tiempo (100 ms, 300 ms, 3 intentos) se fijaron a mano y no se probaron otros valores",
    "La herramienta solo lee. Los resultados sobre reintentos no se pueden extender a una herramienta que escribe",
    "La latencia es de una sola corrida en una máquina",
], size=19)

# 21. conclusion
s = slide_titulo("Conclusión")
bullets(s, [
    "Lo que más cambia el comportamiento es el contrato de entrada",
    "  Sin él, 3 de 3 entradas inválidas se ejecutan y una de ellas devuelve ok",
    "  Con él, A pasa de 6 a 14 casos cumplidos (B)",
    "El timeout y el reintento agregan los 2 casos que B no cubre",
    "El reintento selectivo (D) recupera los fallos transitorios igual que el ingenuo (C)",
    "  Pero sin repetir errores que no se arreglan reintentando: 19 intentos y no 29",
    "Una salida rota no llega al LLM: B, C y D la bloquean como output_contract",
    "Los fallos son simulados; los tiempos no describen al SEACE real",
], size=20)

salida = aqui / "Experimento3_EP.pptx"
prs.save(salida)
exportar(salida, aqui / "Experimento3_EP.pdf", aqui / "_png_exp3")
print("listo", salida, "| diapositivas:", len(prs.slides))
