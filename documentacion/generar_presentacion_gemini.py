#!/usr/bin/env python3
"""Presentación práctica: manejo y uso de Google Gemini (15 min)."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

HERE = Path(__file__).resolve().parent
OUT = HERE / "Manejo-y-Uso-de-Google-Gemini.pptx"

NAVY = RGBColor(0x1F, 0x4E, 0x79)
NAVY_DARK = RGBColor(0x15, 0x36, 0x54)
BLUE = RGBColor(0x2E, 0x86, 0xAB)
TEAL = RGBColor(0x14, 0x8F, 0x77)
GOLD = RGBColor(0xB7, 0x95, 0x0B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x1C, 0x28, 0x33)
MUTED = RGBColor(0x5D, 0x6D, 0x7E)
CARD = RGBColor(0xEA, 0xF2, 0xF8)
CARD2 = RGBColor(0xE8, 0xF8, 0xF5)
CARD3 = RGBColor(0xFC, 0xF3, 0xCF)
RED_SOFT = RGBColor(0xF5, 0xB7, 0xB1)
LINE = RGBColor(0xD4, 0xE6, 0xF1)


def set_run(run, size=18, bold=False, color=INK, font="Calibri"):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_text(tf, text, size=18, bold=False, color=INK, align=PP_ALIGN.LEFT, space_after=6):
    p = tf.paragraphs[0] if not tf.paragraphs[0].text else tf.add_paragraph()
    if not tf.paragraphs[0].text and len(tf.paragraphs) == 1:
        p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(space_after)
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)
    return p


def fill_shape(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def rect(slide, l, t, w, h, color):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    fill_shape(sh, color)
    return sh


def round_rect(slide, l, t, w, h, color):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    fill_shape(sh, color)
    return sh


def textbox(slide, l, t, w, h, text, size=18, bold=False, color=INK, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)
    return box


def bullets(slide, l, t, w, h, items, size=18, color=INK, bold_first=False):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = 0
        p.space_after = Pt(8)
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = "•  " + item
        set_run(run, size=size, bold=(bold_first and i == 0), color=color)
    return box


def footer(slide, page, total=11):
    rect(slide, Inches(0), Inches(7.15), Inches(13.333), Inches(0.35), NAVY)
    textbox(slide, Inches(0.4), Inches(7.16), Inches(10), Inches(0.3),
            "Manejo y uso de Google Gemini  ·  Sesión práctica de 15 minutos",
            size=11, color=WHITE)
    textbox(slide, Inches(11.6), Inches(7.16), Inches(1.4), Inches(0.3),
            f"{page}  /  {total}", size=11, color=WHITE, align=PP_ALIGN.RIGHT)


def header_bar(slide, kicker, title):
    rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(1.15), NAVY)
    rect(slide, Inches(0), Inches(1.15), Inches(13.333), Inches(0.08), TEAL)
    textbox(slide, Inches(0.5), Inches(0.12), Inches(12), Inches(0.3),
            kicker.upper(), size=12, bold=True, color=RGBColor(0xA3, 0xE4, 0xD7))
    textbox(slide, Inches(0.5), Inches(0.42), Inches(12.2), Inches(0.6),
            title, size=26, bold=True, color=WHITE)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def card(slide, l, t, w, h, title, body, fill=CARD, title_color=NAVY):
    round_rect(slide, l, t, w, h, fill)
    textbox(slide, l + Inches(0.18), t + Inches(0.12), w - Inches(0.3), Inches(0.4),
            title, size=16, bold=True, color=title_color)
    box = slide.shapes.add_textbox(l + Inches(0.18), t + Inches(0.5), w - Inches(0.36), h - Inches(0.62))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = body
    set_run(run, size=14, color=INK)


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    total = 11

    # 1 Cover
    s = prs.slides.add_slide(blank)
    rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), NAVY_DARK)
    rect(s, Inches(0), Inches(0), Inches(0.18), Inches(7.5), TEAL)
    textbox(s, Inches(0.7), Inches(1.7), Inches(11.5), Inches(0.4),
            "GUÍA PRÁCTICA  ·  15 MINUTOS", size=14, bold=True, color=RGBColor(0xA3, 0xE4, 0xD7))
    textbox(s, Inches(0.7), Inches(2.15), Inches(12), Inches(1.6),
            "Manejo y uso de\nGoogle Gemini",
            size=40, bold=True, color=WHITE)
    textbox(s, Inches(0.7), Inches(4.15), Inches(11.5), Inches(0.9),
            "Cómo entrar, cómo preguntar y cómo aprovechar la herramienta\npara el trabajo diario — con criterio y sin perder el control.",
            size=18, color=RGBColor(0xD6, 0xEA, 0xF8))
    textbox(s, Inches(0.7), Inches(6.4), Inches(11), Inches(0.4),
            "Área de Tecnología de la Información  ·  gemini.google.com",
            size=14, color=RGBColor(0xAE, 0xB6, 0xBF))
    notes(s,
          "Saludo (30 s). Dejar claro desde el inicio: hoy no es una charla teórica sobre inteligencia artificial. "
          "Es una guía de manejo de la herramienta Google Gemini: dónde se abre, qué botones usar, cómo pedir bien y qué no subir. "
          "Al final, cada persona debería poder entrar y hacer una consulta útil.")

    # 2 Agenda
    s = prs.slides.add_slide(blank)
    header_bar(s, "Mapa de la sesión", "Qué vamos a aprender a usar")
    items = [
        ("01", "2 min", "Cómo entrar y reconocer la pantalla"),
        ("02", "4 min", "Cómo escribir una instrucción que sirva"),
        ("03", "4 min", "Archivos, Drive y funciones clave"),
        ("04", "3 min", "Recorrido práctico de uso"),
        ("05", "2 min", "Cuidados al usar la herramienta y preguntas"),
    ]
    y = 1.5
    for num, tmin, title in items:
        round_rect(s, Inches(0.55), Inches(y), Inches(12.2), Inches(0.85), CARD)
        textbox(s, Inches(0.75), Inches(y + 0.18), Inches(1.1), Inches(0.5), num, size=22, bold=True, color=TEAL)
        textbox(s, Inches(2.0), Inches(y + 0.22), Inches(8.2), Inches(0.45), title, size=20, bold=True, color=NAVY)
        textbox(s, Inches(10.6), Inches(y + 0.24), Inches(1.8), Inches(0.4), tmin, size=16, color=MUTED, align=PP_ALIGN.RIGHT)
        y += 1.0
    footer(s, 2, total)
    notes(s,
          "40 s. Decir que si hay computador, pueden abrir gemini.google.com en paralelo. "
          "La meta: que sepan operar la herramienta, no recitar definiciones.")

    # 3 Acceso
    s = prs.slides.add_slide(blank)
    header_bar(s, "Primer paso", "Cómo entrar a Google Gemini")
    card(s, Inches(0.45), Inches(1.45), Inches(4.05), Inches(5.25),
         "1. Abrir",
         "En el computador: gemini.google.com\n\nEn el celular: aplicación Gemini (Android o iOS).\n\nInicie sesión con su cuenta de Google institucional, si TI así lo definió.",
         CARD)
    card(s, Inches(4.65), Inches(1.45), Inches(4.05), Inches(5.25),
         "2. Qué verá",
         "Arriba o a la izquierda: chats anteriores.\n\nAl centro: caja para escribir.\n\nJunto a la caja: botón para adjuntar archivos y herramientas (Deep Research, Canvas, etc.).\n\nCada conversación es un hilo. Puede volver a ella.",
         CARD2, TEAL)
    card(s, Inches(8.85), Inches(1.45), Inches(4.05), Inches(5.25),
         "3. Antes de escribir",
         "Revise que está en la cuenta correcta.\n\nUn chat nuevo = un tema nuevo. Así no mezcla un procedimiento con un correo personal.\n\nSi la empresa tiene Gemini en Workspace, úselo ahí para Gmail, Docs y Drive.",
         CARD3, RGBColor(0x7D, 0x66, 0x08))
    footer(s, 3, total)
    notes(s,
          "2 min. Si puede, proyecte gemini.google.com 10 segundos y señale: historial, caja de texto, clip de archivos. "
          "Insistir en la cuenta institucional. No mezclar temas en el mismo chat.")

    # 4 Prompt
    s = prs.slides.add_slide(blank)
    header_bar(s, "El corazón de la herramienta", "Cómo pedirle trabajo (la instrucción)")
    steps = [
        ("1. Diga el rol", "«Actúa como analista de sistemas» o «explica para un usuario no técnico»."),
        ("2. Diga el objetivo", "Resumir, redactar, comparar, armar pasos, corregir, traducir."),
        ("3. Entregue el material", "Pegue el texto o adjunte el archivo. Sin datos, Gemini inventa."),
        ("4. Pida el formato", "Lista, tabla, correo, 8 viñetas, procedimiento de 12 pasos."),
        ("5. Ponga un límite", "«No inventes datos que no estén en el archivo. Si falta algo, dímelo»."),
    ]
    y = 1.4
    for title, body in steps:
        round_rect(s, Inches(0.5), Inches(y), Inches(12.3), Inches(0.9), CARD)
        textbox(s, Inches(0.7), Inches(y + 0.2), Inches(3.3), Inches(0.5), title, size=18, bold=True, color=TEAL)
        textbox(s, Inches(4.1), Inches(y + 0.22), Inches(8.4), Inches(0.5), body, size=16, color=INK)
        y += 1.0
    footer(s, 4, total)
    notes(s,
          "2,5 min. Contraste rápido: mal prompt «ayúdame con TI» vs buen prompt con rol, objetivo, archivo y «no inventes». "
          "Diga que se puede continuar el chat: «acorta a media página», «pásalo a correo», «hazlo para guardias».")

    # 5 Ejemplo de instrucción
    s = prs.slides.add_slide(blank)
    header_bar(s, "Ejemplo para copiar", "Una instrucción débil vs una que sí se puede usar")
    round_rect(s, Inches(0.45), Inches(1.45), Inches(6.05), Inches(5.25), RGBColor(0xFD, 0xED, 0xEC))
    textbox(s, Inches(0.7), Inches(1.65), Inches(5.6), Inches(0.45), "Poco útil", size=20, bold=True, color=RGBColor(0x92, 0x2B, 0x21))
    box = s.shapes.add_textbox(Inches(0.7), Inches(2.25), Inches(5.55), Inches(4.1))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = (
        "«Háblame de Gemini.»\n\n"
        "«Hazme un procedimiento.»\n\n"
        "«Resume esto.»  (sin adjuntar nada)\n\n"
        "Resultado: texto genérico, largo y difícil de aplicar."
    )
    set_run(run, size=16, color=INK)
    round_rect(s, Inches(6.8), Inches(1.45), Inches(6.05), Inches(5.25), CARD2)
    textbox(s, Inches(7.05), Inches(1.65), Inches(5.6), Inches(0.45), "Útil", size=20, bold=True, color=TEAL)
    box = s.shapes.add_textbox(Inches(7.05), Inches(2.25), Inches(5.55), Inches(4.1))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = (
        "«Eres asistente de TI. Con el PDF adjunto, arma un resumen de 8 viñetas para el usuario final. "
        "Luego un correo corto de convocatoria. "
        "No agregues pasos que no estén en el documento. "
        "Marca con «Falta dato» lo que no aparezca.»\n\n"
        "Resultado: un borrador revisable en minutos."
    )
    set_run(run, size=16, color=INK)
    footer(s, 5, total)
    notes(s,
          "1,5 min. Leer en voz alta las dos columnas. Invitar a que copien la de la derecha y la adapten. "
          "Mensaje: la calidad de Gemini depende de la calidad de lo que usted escribe y adjunta.")

    # 6 Archivos
    s = prs.slides.add_slide(blank)
    header_bar(s, "Traer su material", "Cómo adjuntar archivos y usar Drive")
    bullets(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(2.2), [
        "En la caja de texto, pulse el ícono de adjuntar (clip o «Añadir archivos»).",
        "Puede subir PDF, Word, imágenes, hojas de cálculo o capturas de pantalla.",
        "También puede tomar un archivo desde Google Drive, si la cuenta está conectada.",
        "Después de adjuntar, escriba qué quiere que haga con ese archivo: resumir, extraer pasos, comparar, armar checklist.",
    ], size=17)
    card(s, Inches(0.45), Inches(4.0), Inches(4.05), Inches(2.7),
         "Imagen o captura",
         "«¿Qué dice esta pantalla?» «Lista los campos del formulario.» Útil para tickets y errores.",
         CARD)
    card(s, Inches(4.65), Inches(4.0), Inches(4.05), Inches(2.7),
         "PDF o procedimiento",
         "«Convierte esto en pasos numerados y responsables.» Ideal para TI-PRO y actas.",
         CARD2, TEAL)
    card(s, Inches(8.85), Inches(4.0), Inches(4.05), Inches(2.7),
         "Hoja o listado",
         "«Encuentra duplicados, fechas vencidas o equipos sin usuario.» Revise siempre las cifras.",
         CARD3, RGBColor(0x7D, 0x66, 0x08))
    footer(s, 6, total)
    notes(s,
          "2 min. Si hay demo: adjunte un PDF inofensivo (un procedimiento interno no confidencial o un texto de ejemplo) y pida 8 viñetas. "
          "Advertir: no subir cédulas, contraseñas ni datos de clientes.")

    # 7 Funciones
    s = prs.slides.add_slide(blank)
    header_bar(s, "Caja de herramientas", "Funciones de Gemini que sí conviene conocer")
    feats = [
        ("Chat", "Conversación normal. Siga el hilo: corrija, acorte, cambie el tono o pida otra versión."),
        ("Deep Research", "En la barra, elija Deep Research. Gemini arma un plan, busca fuentes y entrega un informe. Revise el plan antes de iniciar."),
        ("Canvas", "Espacio al lado del chat para redactar, editar o armar un entregable más largo (documento, esquema, código)."),
        ("Gems", "Menú Gems: cree un asistente con instrucciones fijas (ej. «redacta actas de TI»). Se reutiliza sin repetir el prompt."),
        ("Gemini Live", "En el celular: voz, cámara o pantalla. Útil para explicar un error en el equipo con las manos ocupadas."),
        ("Exportar", "Muchas respuestas se pueden pasar a Docs, Gmail o copiar. Usted pule y envía; Gemini no envía solo."),
    ]
    y = 1.4
    for i, (title, body) in enumerate(feats):
        col = i % 2
        if col == 0 and i:
            y += 1.7
        x = 0.45 + col * 6.4
        fill = CARD if col == 0 else CARD2
        tc = NAVY if col == 0 else TEAL
        card(s, Inches(x), Inches(y), Inches(6.15), Inches(1.55), title, body, fill, tc)
    footer(s, 7, total)
    notes(s,
          "2,5 min. No explicar todas a fondo. Priorizar: Chat + adjuntar (uso diario), Gems (si el área repite la misma tarea), "
          "Deep Research (solo para investigar un tema, no para inventar política interna). Live: mención breve para campo/soporte.")

    # 8 Recorrido práctico
    s = prs.slides.add_slide(blank)
    header_bar(s, "Hágalo así", "Recorrido de uso en 6 clics")
    flow = [
        ("1", "Entre a gemini.google.com e inicie un chat nuevo."),
        ("2", "Adjunte el archivo o pegue el texto de trabajo."),
        ("3", "Escriba rol + objetivo + formato + «no inventes»."),
        ("4", "Lea la respuesta. Pida ajuste: más corto, en tabla, en tono de correo."),
        ("5", "Copie o exporte. Corrija nombres, fechas y reglas de la empresa."),
        ("6", "Publique o envíe usted. Guarde el chat si el prompt quedó bueno."),
    ]
    y = 1.4
    for n, t in flow:
        round_rect(s, Inches(0.5), Inches(y), Inches(12.3), Inches(0.8), CARD)
        textbox(s, Inches(0.7), Inches(y + 0.16), Inches(0.7), Inches(0.5), n, size=22, bold=True, color=TEAL)
        textbox(s, Inches(1.6), Inches(y + 0.2), Inches(10.8), Inches(0.45), t, size=18, color=INK)
        y += 0.88
    footer(s, 8, total)
    notes(s,
          "2 min. Este es el corazón de la exposición. Recorrerlo despacio. Si hay tiempo, ejecutarlo en vivo con un ejemplo inofensivo. "
          "El paso 5 no se salta: revisar siempre.")

    # 9 Para qué usarla hoy
    s = prs.slides.add_slide(blank)
    header_bar(s, "Valor para el usuario", "Qué puede hacer hoy con la herramienta")
    uses = [
        ("Redactar", "Correos, convocatorias, actas y el primer borrador de un procedimiento."),
        ("Entender", "Resumir un PDF largo, una política o un hilo de correos en 10 líneas."),
        ("Preparar", "Checklist de mantenimiento, preguntas para una reunión, guía para el usuario."),
        ("Apoyar soporte", "A partir de una captura: posibles causas y qué verificar (sin cerrar el ticket)."),
        ("Traducir / adaptar", "Pasar un texto técnico a lenguaje de personal administrativo o de campo."),
        ("Ordenar ideas", "De apuntes sueltos a pasos, tabla de responsables o comparación de opciones."),
    ]
    y = 1.4
    for i, (title, body) in enumerate(uses):
        col = i % 3
        row = i // 3
        x = 0.4 + col * 4.25
        yy = y + row * 2.55
        card(s, Inches(x), Inches(yy), Inches(4.05), Inches(2.35), title, body, CARD if row == 0 else CARD2, NAVY if row == 0 else TEAL)
    footer(s, 9, total)
    notes(s,
          "2 min. Un ejemplo por bloque, del trabajo real: acta, PDF de procedimiento, checklist semestral, captura de error, texto para guardias. "
          "Cerrar: el valor está en minutos ahorrados en tareas que ya hacen, no en «probar IA».")

    # 10 Cuidados
    s = prs.slides.add_slide(blank)
    header_bar(s, "Uso responsable", "Al manejar Gemini, tenga presentes estas reglas")
    rules = [
        ("No suba", "Contraseñas, cédulas, datos de clientes, nómina ni información que no publicaría en un correo abierto."),
        ("Verifique", "Nombres, fechas, cifras y pasos. Gemini puede afirmar con seguridad algo que no está en su archivo."),
        ("Separe chats", "Un tema por conversación. Facilita reutilizar el prompt y no mezclar contextos."),
        ("Usted envía", "Gemini no reemplaza la aprobación de jefatura ni la firma de un procedimiento."),
    ]
    y = 1.45
    for title, body in rules:
        round_rect(s, Inches(0.5), Inches(y), Inches(12.3), Inches(1.2), RGBColor(0xFD, 0xED, 0xEC) if title == "No suba" else CARD)
        textbox(s, Inches(0.75), Inches(y + 0.32), Inches(2.6), Inches(0.55), title, size=18, bold=True, color=RGBColor(0x92, 0x2B, 0x21) if title == "No suba" else NAVY)
        textbox(s, Inches(3.5), Inches(y + 0.28), Inches(9.0), Inches(0.7), body, size=16, color=INK)
        y += 1.3
    footer(s, 10, total)
    notes(s,
          "1,5 min. Tono práctico, no alarmista. Analogía: es como reenviar un borrador a un colega externo: no le mande lo que no debe salir del área. "
          "Si la empresa tiene Gemini de Workspace, preferirlo para archivos de Drive.")

    # 11 Cierre
    s = prs.slides.add_slide(blank)
    rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), NAVY_DARK)
    rect(s, Inches(0), Inches(0), Inches(0.18), Inches(7.5), TEAL)
    textbox(s, Inches(0.7), Inches(1.5), Inches(12), Inches(0.4),
            "PARA LLEVARSE HOY", size=14, bold=True, color=RGBColor(0xA3, 0xE4, 0xD7))
    textbox(s, Inches(0.7), Inches(2.0), Inches(12), Inches(1.8),
            "Entre, adjunte, pida con claridad,\nrevise y envíe usted.",
            size=32, bold=True, color=WHITE)
    textbox(s, Inches(0.7), Inches(4.2), Inches(12), Inches(1.3),
            "gemini.google.com  ·  Un chat por tema  ·  «No inventes» en la instrucción.\nPreguntas.",
            size=20, color=RGBColor(0xD6, 0xEA, 0xF8))
    textbox(s, Inches(0.7), Inches(6.4), Inches(12), Inches(0.4),
            "Área de Tecnología de la Información",
            size=14, color=RGBColor(0xAE, 0xB6, 0xBF))
    notes(s,
          "Cierre 40 s. Repetir la secuencia. Abrir preguntas. Si hay silencio: «¿quién quiere que recorramos juntos un ejemplo de correo o de resumen de PDF?»")

    prs.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
