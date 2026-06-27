# -*- coding: utf-8 -*-
"""
Genera Presentacion_TFG.pptx a partir del contenido del TFG.
Reutiliza las capturas de las peliculas que estan en la raiz del repo.
Ejecutar desde la carpeta del repo:  python3 presentacion/generar_pptx.py
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(REPO, "presentacion", "Presentacion_TFG.pptx")

# Paleta
AZUL = RGBColor(0x1F, 0x3A, 0x5F)      # azul profundo
ACENTO = RGBColor(0xC0, 0x6B, 0x3E)    # terracota
GRIS = RGBColor(0x3A, 0x3A, 0x3A)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)
CLARO = RGBColor(0xF2, 0xEF, 0xE9)

prs = Presentation()
prs.slide_width = Inches(13.333)   # 16:9
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

IMG = {
    "hogar": "Tranquilidad y felicidad del hogar de Alicia frente a la verdad del exterior.png",
    "madres": "Alicia contemplando la marcha de las Madres de Mayo .png",
    "paisaje": "Plano del paisaje en San Luis.png",
    "tren": "Contraste entre el mundo del avance (tren) y mundo rural.png",
    "pacifico": "Cartel de Nueva York y Pacifico en el video de Hache.png",
    "aeropuerto": "Dante recibe a Martin y a Hache en el aeropuerto.png",
}


def slide():
    return prs.slides.add_slide(BLANK)


def rect(s, x, y, w, h, color):
    from pptx.enum.shapes import MSO_SHAPE
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, sp_after=6):
    """runs: lista de parrafos; cada parrafo es lista de (texto, size, bold, color, italic)."""
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(sp_after)
        for (t, size, bold, color, italic) in para:
            r = p.add_run()
            r.text = t
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.color.rgb = color
            r.font.name = "Calibri"
    return tb


def bullets(s, x, y, w, h, items, size=18, color=GRIS):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(10)
        # it puede ser (texto) o (label, texto)
        if isinstance(it, tuple):
            r = p.add_run(); r.text = "•  " + it[0]
            r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = AZUL; r.font.name = "Calibri"
            r2 = p.add_run(); r2.text = " — " + it[1]
            r2.font.size = Pt(size); r2.font.color.rgb = color; r2.font.name = "Calibri"
        else:
            r = p.add_run(); r.text = "•  " + it
            r.font.size = Pt(size); r.font.color.rgb = color; r.font.name = "Calibri"
    return tb


def header(s, kicker, title):
    rect(s, 0, 0, SW, Inches(1.25), AZUL)
    rect(s, 0, Inches(1.25), SW, Inches(0.06), ACENTO)
    txt(s, Inches(0.6), Inches(0.12), Inches(12), Inches(0.4),
        [[(kicker.upper(), 12, True, RGBColor(0xD9, 0xC7, 0xB8), False)]])
    txt(s, Inches(0.6), Inches(0.42), Inches(12.1), Inches(0.8),
        [[(title, 26, True, BLANCO, False)]], anchor=MSO_ANCHOR.MIDDLE)


def add_image(s, key, x, y, w, caption=None):
    path = os.path.join(REPO, IMG[key])
    if not os.path.exists(path):
        print("  ! Falta imagen:", IMG[key]); return
    pic = s.shapes.add_picture(path, x, y, width=w)
    if caption:
        txt(s, x, y + pic.height + Pt(2), w, Inches(0.5),
            [[(caption, 11, False, GRIS, True)]], align=PP_ALIGN.CENTER)
    return pic


# ----------------------------------------------------------------------------
# 1. PORTADA
s = slide()
rect(s, 0, 0, SW, SH, AZUL)
rect(s, 0, Inches(3.05), SW, Inches(0.07), ACENTO)
txt(s, Inches(1), Inches(0.7), Inches(11.3), Inches(0.5),
    [[("GRADO EN HISTORIA DEL ARTE  ·  UNIVERSIDAD DE MURCIA", 14, True, RGBColor(0xD9,0xC7,0xB8), False)]])
txt(s, Inches(1), Inches(1.7), Inches(11.3), Inches(1.4),
    [[("La visibilidad del cine argentino en España (1980–2000)", 36, True, BLANCO, False)]])
txt(s, Inches(1), Inches(3.25), Inches(11.3), Inches(0.7),
    [[("Memoria, exilio y coproducciones transatlánticas", 22, False, RGBColor(0xE9,0xDD,0xCF), True)]])
txt(s, Inches(1), Inches(5.4), Inches(11.3), Inches(1.5),
    [[("Trabajo de Fin de Grado · Facundo Nicolás Vallejos Ribles", 18, True, BLANCO, False)],
     [("Director: Prof. Joaquín Cánovas Belchí", 15, False, RGBColor(0xD9,0xC7,0xB8), False)],
     [("Convocatoria junio–julio · Curso 2025/2026", 13, False, RGBColor(0xB9,0xA8,0x99), False)]])

# ----------------------------------------------------------------------------
# 2. JUSTIFICACION / OBJETO
s = slide()
header(s, "1. Introducción", "Objeto de estudio y justificación")
bullets(s, Inches(0.7), Inches(1.7), Inches(12), Inches(4.8), [
    "El cine argentino en España (1980–2000), entendido no solo como fenómeno industrial sino como proceso cultural.",
    "Parte de una larga historia de intercambios cinematográficos entre las dos orillas del Atlántico.",
    "España no era un mercado extranjero cualquiera: venía de su propia dictadura y transición democrática.",
    "El cine argentino de la posdictadura podía dialogar con los silencios y heridas del público español.",
    "Enfoque desde la Historia del Arte: no solo qué cuentan las películas, sino cómo lo construyen visualmente.",
], size=19)

# ----------------------------------------------------------------------------
# 3. HIPOTESIS Y OBJETIVOS
s = slide()
header(s, "1. Introducción", "Hipótesis y objetivos")
bullets(s, Inches(0.7), Inches(1.7), Inches(12), Inches(4.8), [
    ("Memoria", "el cine argentino tuvo una recepción singular por tratar la represión política, reconocible desde la experiencia española."),
    ("Exilio", "el exiliado argentino no fue un sujeto pasivo, sino un mediador cultural que tejió redes de intercambio."),
    ("Coproducciones", "fueron una de las vías más eficaces de visibilidad, con sus posibilidades y sus tensiones."),
], size=20)
txt(s, Inches(0.7), Inches(5.0), Inches(12), Inches(1.5),
    [[("Objetivo transversal: ", 18, True, ACENTO, False),
      ("analizar cómo la memoria, el exilio y la industria se traducen en el lenguaje narrativo y estético de las películas.", 18, False, GRIS, False)]])

# ----------------------------------------------------------------------------
# 4. ESTADO DE LA CUESTION + METODOLOGIA
s = slide()
header(s, "1. Introducción", "Estado de la cuestión y metodología")
txt(s, Inches(0.7), Inches(1.55), Inches(5.8), Inches(0.5), [[("Estado de la cuestión", 18, True, AZUL, False)]])
bullets(s, Inches(0.7), Inches(2.1), Inches(5.9), Inches(4.5), [
    "Posdictadura: Amado y la «imagen justa».",
    "Exilio: Silvina Jensen.",
    "Circulación transatlántica: González; datos de Elena.",
    "Hueco detectado: la bibliografía relega el análisis formal de las películas.",
], size=16)
txt(s, Inches(6.9), Inches(1.55), Inches(5.7), Inches(0.5), [[("Metodología", 18, True, AZUL, False)]])
bullets(s, Inches(6.9), Inches(2.1), Inches(5.8), Inches(4.5), [
    "Combina análisis formal (Historia del Arte) e histórico-cultural.",
    "Películas leídas como textos visuales: puesta en escena, espacio, luz, encuadre, montaje.",
    "Pautas de Cánovas Belchí y Aliaga Cárceles, Historia del cine (2020).",
    "Citas en estilo Chicago-Deusto.",
], size=16)

# ----------------------------------------------------------------------------
# 5. CONTEXTO HISTORICO
s = slide()
header(s, "2. Contexto histórico y cultural", "Dos sociedades saliendo de la violencia de Estado")
rect(s, Inches(0.7), Inches(1.7), Inches(5.85), Inches(4.7), CLARO)
txt(s, Inches(0.95), Inches(1.85), Inches(5.4), Inches(0.5), [[("ARGENTINA", 16, True, ACENTO, False)]])
bullets(s, Inches(0.95), Inches(2.4), Inches(5.4), Inches(4.0), [
    "Antecedente de la Triple A (1973–76): primeros exilios (Briski, Guevara, Alterio).",
    "Golpe de 1976: Proceso, secuestro, tortura y desaparición.",
    "El exilio se vuelve necesidad de supervivencia.",
], size=15)
rect(s, Inches(6.8), Inches(1.7), Inches(5.85), Inches(4.7), CLARO)
txt(s, Inches(7.05), Inches(1.85), Inches(5.4), Inches(0.5), [[("ESPAÑA", 16, True, ACENTO, False)]])
bullets(s, Inches(7.05), Inches(2.4), Inches(5.4), Inches(4.0), [
    "Muerte de Franco (1975) y Constitución de 1978.",
    "Transición y apertura cultural.",
    "Madrid y Barcelona: espacios receptivos para el exilio argentino.",
], size=15)
txt(s, Inches(0.7), Inches(6.55), Inches(12), Inches(0.7),
    [[("→ El cine se convierte en un espacio común de memoria entre las dos sociedades.", 16, True, AZUL, False)]],
    align=PP_ALIGN.CENTER)

# ----------------------------------------------------------------------------
# 6. LA MEMORIA: 80s vs 90s
s = slide()
header(s, "3. Memoria y reconstrucción", "El eje de la memoria: de los 80 a los 90")
rect(s, Inches(0.7), Inches(1.7), Inches(5.85), Inches(4.7), CLARO)
txt(s, Inches(0.95), Inches(1.85), Inches(5.4), Inches(0.5), [[("AÑOS 80", 16, True, ACENTO, False)]])
bullets(s, Inches(0.95), Inches(2.4), Inches(5.4), Inches(4.0), [
    "Realismo afectivo / «imaginación melodramática».",
    "Vocación didáctica, relatos transparentes.",
    "El espacio familiar como metonimia del cuerpo social.",
], size=15)
rect(s, Inches(6.8), Inches(1.7), Inches(5.85), Inches(4.7), CLARO)
txt(s, Inches(7.05), Inches(1.85), Inches(5.4), Inches(0.5), [[("AÑOS 90 · NUEVO CINE ARGENTINO", 15, True, ACENTO, False)]])
bullets(s, Inches(7.05), Inches(2.4), Inches(5.4), Inches(4.0), [
    "Ambigüedad, fragmento y distancia crítica.",
    "Richard: la memoria es un territorio en disputa.",
    "Sarlo: el riesgo del «giro subjetivo».",
], size=15)

# ----------------------------------------------------------------------------
# 7. CASO 1: LA HISTORIA OFICIAL
s = slide()
header(s, "Caso de estudio 1", "La historia oficial (Luis Puenzo, 1985)")
bullets(s, Inches(0.7), Inches(1.6), Inches(5.3), Inches(4.8), [
    "Filme paradigmático de la primera revisión histórica; estreno durante el Juicio a las Juntas.",
    "El espacio como ceguera: interiores cálidos y burgueses frente a exteriores fríos (Madres de Plaza de Mayo).",
    "La transformación de Alicia como alegoría nacional.",
    "Óscar a mejor película extranjera (1986).",
    "Matiz crítico: el melodrama puede rozar la «teoría de los dos demonios».",
], size=15)
add_image(s, "hogar", Inches(6.35), Inches(1.7), Inches(3.05), "El hogar de Alicia")
add_image(s, "madres", Inches(9.55), Inches(1.7), Inches(3.05), "Las Madres de Plaza de Mayo")

# ----------------------------------------------------------------------------
# 8. EL EXILIO COMO PUENTE
s = slide()
header(s, "4. El exilio, un vínculo entre países", "El exiliado como arquitecto de puentes de memoria")
bullets(s, Inches(0.7), Inches(1.7), Inches(12), Inches(4.8), [
    "El exilio no es migración económica: es huida de una violencia que amenaza la vida.",
    "Rodríguez Marino: «figuras del destierro» (melancolía, pérdida, retorno imposible).",
    "Reflexiones de un salvaje (Gerardo Vallejo, 1978), rodada en España.",
    "Paralelismo franquismo / represión argentina: los niños del pueblo, Don Quijote y Martín Fierro, los republicanos asesinados.",
    "Montaje en deuda con Eisenstein: el exiliado tiende un puente de memoria entre las dos historias.",
], size=18)

# ----------------------------------------------------------------------------
# 9. COPRODUCCIONES
s = slide()
header(s, "4. El exilio, un vínculo entre países", "Las coproducciones: la vía industrial")
bullets(s, Inches(0.7), Inches(1.7), Inches(12), Inches(4.8), [
    "Más que una fórmula administrativa: estrategia de supervivencia frente al dominio de Hollywood.",
    "Convenio Hispano-Argentino de 1969 (BOE n.º 238): doble nacionalidad; participación del 30 %–70 %.",
    "Datos de Elena: Argentina, segundo origen latinoamericano en España (1933–1995) tras México.",
    "Distintos modelos: Aristarain (autoral) frente a Piñeyro (industrial: Plata quemada, Kamchatka).",
    "Tensión clave: la asimetría económica condiciona qué imagen de Argentina circula en España.",
], size=18)

# ----------------------------------------------------------------------------
# 10. CASO 2: UN LUGAR EN EL MUNDO
s = slide()
header(s, "Caso de estudio 2", "Un lugar en el mundo (Aristarain, 1992)")
bullets(s, Inches(0.7), Inches(1.6), Inches(5.3), Inches(4.8), [
    "Coproducción Argentina/Uruguay/España. Concha de Oro (San Sebastián) y Goya.",
    "Diálogo con el western: el paisaje árido de San Luis como aislamiento y resistencia.",
    "El tren: imagen del progreso que avanza pero arrasa.",
    "El espacio unitario: los desplazamientos articulan un mundo con unidad → «tener un lugar».",
    "Superó los 500.000 espectadores en España.",
], size=15)
add_image(s, "paisaje", Inches(6.35), Inches(1.7), Inches(3.05), "El paisaje de San Luis")
add_image(s, "tren", Inches(9.55), Inches(1.7), Inches(3.05), "El tren frente al mundo rural")

# ----------------------------------------------------------------------------
# 11. CASO 3: MARTIN (HACHE)
s = slide()
header(s, "Caso de estudio 3", "Martín (Hache) (Aristarain, 1997)")
bullets(s, Inches(0.7), Inches(1.6), Inches(5.3), Inches(4.8), [
    "Registro íntimo y urbano; predominio del diálogo (cerca del cine de autor europeo).",
    "Espacio fragmentado = identidad fragmentada. Madrid y Buenos Aires casi indiferenciados.",
    "Omisión de la dictadura: la memoria queda latente en el silencio.",
    "El título: la «hache» muda, el tiempo en suspenso; el brindis «por Martín» restituye el nombre.",
    "Cecilia Roth y Eusebio Poncela refuerzan la condición transatlántica.",
], size=15)
add_image(s, "pacifico", Inches(6.35), Inches(1.7), Inches(3.05), "«Pacífico» sobre Nueva York")
add_image(s, "aeropuerto", Inches(9.55), Inches(1.7), Inches(3.05), "Llegada al aeropuerto")

# ----------------------------------------------------------------------------
# 12. RECEPCION EN ESPAÑA
s = slide()
header(s, "4.4. Recepción en España", "Éxito crítico e institucional, más que de taquilla")
bullets(s, Inches(0.7), Inches(1.7), Inches(12), Inches(4.8), [
    "Los festivales fueron decisivos; la crítica leyó el cine argentino como una cinematografía con identidad propia.",
    "Elena llama a Aristarain el «buque insignia» del cine argentino en España.",
    "Premios Goya: Un lugar en el mundo, Cenizas del paraíso y el Goya de Cecilia Roth por Martín (Hache).",
    "Salvo Un lugar en el mundo, la mayoría circuló en salas especializadas y versión original.",
    "La visibilidad se apoyó más en la vía crítica e institucional que en el público masivo.",
], size=18)

# ----------------------------------------------------------------------------
# 13. CONCLUSIONES
s = slide()
header(s, "5. Conclusiones", "Resultados y resolución de las hipótesis")
bullets(s, Inches(0.7), Inches(1.65), Inches(12), Inches(4.8), [
    ("El espacio como sentido", "ceguera (La historia oficial), unidad (Un lugar en el mundo), fragmentación (Martín (Hache))."),
    ("Dos modos de memoria", "realismo afectivo y didáctico (80) frente a ambigüedad y distanciamiento (90)."),
    ("Hipótesis confirmadas", "recepción por resonancia, exilio como tejido conectivo, coproducciones como vía de difusión."),
], size=18)
txt(s, Inches(0.7), Inches(5.5), Inches(12), Inches(1.5),
    [[("Entre 1980 y 2000 las imágenes cruzaron el Atlántico cargadas de memoria, pero también de nuevas posibilidades estéticas e industriales.", 18, True, ACENTO, True)]])

# ----------------------------------------------------------------------------
# 14. CIERRE
s = slide()
rect(s, 0, 0, SW, SH, AZUL)
rect(s, Inches(0), Inches(3.55), SW, Inches(0.06), ACENTO)
txt(s, Inches(1), Inches(2.6), Inches(11.3), Inches(1.0),
    [[("Muchas gracias por su atención", 34, True, BLANCO, False)]], align=PP_ALIGN.CENTER)
txt(s, Inches(1), Inches(3.8), Inches(11.3), Inches(0.8),
    [[("Quedo a disposición del tribunal para sus preguntas", 18, False, RGBColor(0xE9,0xDD,0xCF), True)]], align=PP_ALIGN.CENTER)
txt(s, Inches(1), Inches(6.4), Inches(11.3), Inches(0.6),
    [[("Facundo Nicolás Vallejos Ribles · Grado en Historia del Arte · Universidad de Murcia", 13, False, RGBColor(0xB9,0xA8,0x99), False)]], align=PP_ALIGN.CENTER)

prs.save(OUT)
print("OK ->", OUT, "| diapositivas:", len(prs.slides._sldIdLst))
