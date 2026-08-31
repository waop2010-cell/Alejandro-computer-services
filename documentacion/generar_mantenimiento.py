#!/usr/bin/env python3
"""Genera el procedimiento TI-PRO-004 de mantenimiento semestral
en la misma plantilla institucional y el diagrama Draw.io."""

from __future__ import annotations

from pathlib import Path
import sys

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from aplicar_formato_plantilla import Builder, render_from_template
OUT_DOCX = HERE / "TI-PRO-004-Mantenimiento-Equipos-Informaticos.docx"
OUT_DRAWIO = HERE / "TI-PRO-004-Diagrama.drawio"
OUT_MD = HERE / "TI-PRO-004-Pasos-a-seguir.md"
OUT_HTML = HERE / "vista-previa-mantenimiento.html"


def content(b: Builder):
    els = []
    a = els.append

    a(b.blank())
    a(b.title("PROCEDIMIENTO PARA EL MANTENIMIENTO PREVENTIVO DE EQUIPOS INFORMÁTICOS"))
    a(b.meta("Código:", " TI-PRO-004"))
    a(b.meta("Versión:", " 1.0"))
    a(b.meta("Área Responsable:", " Tecnología de la Información (TI)"))
    a(b.meta("Periodicidad:", " Cada seis (6) meses"))
    a(b.blank())

    a(b.section("1. OBJETIVO"))
    a(
        b.body(
            "Establecer el procedimiento para planificar, ejecutar, documentar y cerrar el mantenimiento preventivo "
            "de los equipos informáticos de la organización cada seis meses, con el fin de conservar su disponibilidad, "
            "rendimiento y seguridad, reducir fallas no programadas y prolongar la vida útil de los activos tecnológicos."
        )
    )
    a(
        b.body(
            "Este procedimiento complementa al TI-PRO-003 (adquisición de tecnología). Cuando el mantenimiento detecte "
            "un equipo sin arreglo técnico o al final de su vida útil, se derivará al proceso de compra correspondiente."
        )
    )
    a(b.blank())

    a(b.section("2. ALCANCE"))
    a(
        b.body(
            "Aplica a todos los equipos informáticos registrados en el inventario tecnológico, asignados a clientes internos "
            "y externos, en Quito, Guayaquil y sitios de operación. Incluye, como mínimo:"
        )
    )
    for item in [
        "Computadoras de escritorio y laptops",
        "Tablets de operación",
        "Monitores y periféricos asociados al puesto de trabajo",
        "Equipos de red de usuario (cuando estén bajo custodia de TI)",
        "Impresoras de uso interno, cuando aplique",
        "Celulares corporativos, cuando aplique",
        "Accesorios tecnológicos registrados como activo",
    ]:
        a(b.item(item))
    a(
        b.body(
            "No sustituye al soporte correctivo por ticket de falla. El correctivo de emergencia se atiende de inmediato; "
            "el preventivo se ejecuta de forma programada cada seis meses, sin perjuicio de adelantarlo si el diagnóstico lo justifica."
        )
    )
    a(b.blank())
    a(b.section("Definición de clientes"))
    a(b.section("Cliente interno"))
    a(b.body("Personal administrativo de oficinas en Quito y Guayaquil."))
    a(b.body("El mantenimiento se realiza preferentemente en sitio o en el taller de TI, previa coordinación de agenda."))
    a(b.section("Cliente externo"))
    a(b.body("Guardias, supervisores y demás clientes por fuera (tablets y laptops de operación)."))
    a(
        b.body(
            "El mantenimiento se coordina con el área responsable del personal de campo. El equipo se recolecta, se recibe "
            "en TI o se interviene en sitio, según disponibilidad operativa."
        )
    )
    a(b.blank())
    a(b.section("Definiciones"))
    a(
        b.body(
            "Mantenimiento preventivo: conjunto de actividades planificadas de limpieza, inspección, actualización, "
            "optimización y verificación que se ejecutan cada seis meses para evitar fallas."
        )
    )
    a(
        b.body(
            "Mantenimiento correctivo: intervención no programada para reparar una falla ya presentada. Si el equipo no tiene "
            "arreglo, se aplica el TI-PRO-003."
        )
    )
    a(
        b.body(
            "Ciclo semestral: período de seis meses contado desde la fecha de entrega del equipo, desde el último mantenimiento "
            "cerrado o desde la fecha de alta en inventario, la que resulte más reciente."
        )
    )
    a(b.blank())

    a(b.section("3. RESPONSABILIDADES"))
    a(b.section("Área de Tecnología de la Información"))
    for item in [
        "Elaborar el calendario semestral a partir del inventario de activos",
        "Notificar al usuario o área con al menos cinco días hábiles de anticipación",
        "Ejecutar el checklist de hardware, software y seguridad",
        "Realizar o verificar el respaldo de información antes de intervenir",
        "Registrar hallazgos, tiempos de intervención y próximo vencimiento",
        "Entregar el equipo con acta de mantenimiento y actualizar el inventario",
        "Derivar a TI-PRO-003 cuando el equipo no tenga arreglo o deba reemplazarse",
        "Emitir el informe semestral de cumplimiento a Gerencia",
    ]:
        a(b.item(item))
    a(b.section("Usuario o área solicitante"))
    for item in [
        "Facilitar el equipo en la fecha coordinada",
        "Informar claves, accesorios y novedades de uso (cuando aplique y esté autorizado)",
        "Resguardar información crítica antes de la intervención, en lo que le corresponda",
        "Firmar el acta de recepción posterior al mantenimiento",
        "Reportar por ticket cualquier anomalía posterior a la entrega",
    ]:
        a(b.item(item))
    a(b.section("Jefatura o supervisión del área"))
    for item in [
        "Autorizar ventanas de mantenimiento que no afecten la operación crítica",
        "Cliente externo: coordinar el retiro o la visita de tablets y laptops de guardias y supervisores",
    ]:
        a(b.item(item))
    a(b.section("Gerencia"))
    a(
        b.body(
            "Conocer el indicador de cumplimiento del plan semestral y autorizar reemplazos cuando TI derive el caso al proceso de adquisición."
        )
    )
    a(b.blank())

    a(b.section("4. DESCRIPCIÓN DEL PROCEDIMIENTO"))
    a(b.section("Paso 1. Programación del ciclo semestral"))
    a(
        b.body(
            "En los meses de control (o de forma continua sobre el inventario), TI identifica los equipos cuyo mantenimiento "
            "preventivo está vencido o vence en los siguientes treinta días."
        )
    )
    a(b.body("El plan debe indicar:"))
    for item in [
        "Código de inventario, serie, modelo y usuario asignado",
        "Sede (Quito, Guayaquil u operación de campo)",
        "Tipo de cliente (interno o externo)",
        "Fecha del último mantenimiento y fecha objetivo del actual",
        "Ventana tentativa de intervención",
    ]:
        a(b.item(item))
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("Paso 2. Generación del ticket de mantenimiento"))
    a(
        b.body(
            "TI abre un ticket de mantenimiento preventivo por equipo o por lote controlado (mismo usuario o misma sede), "
            "dejando trazabilidad de fecha, técnico asignado y checklist aplicable."
        )
    )
    a(b.body("El ticket no espera a que el usuario reporte una falla. Es una actividad planificada por TI."))
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("Paso 3. Notificación y coordinación"))
    a(b.body("TI notifica al usuario y a su jefatura, con al menos cinco días hábiles, indicando fecha, hora, lugar y duración estimada."))
    a(b.section("Cliente interno"))
    a(b.body("Se agenda en oficina de Quito o Guayaquil. El usuario deja el equipo disponible o acude al taller de TI, según se coordine."))
    a(b.section("Cliente externo"))
    a(
        b.body(
            "Se coordina con supervisión de campo el retiro, la visita o el reemplazo temporal de tablets y laptops, "
            "para no dejar desatendida la operación."
        )
    )
    a(b.body("Si el usuario no facilita el equipo en dos coordinaciones, TI escala a la jefatura y deja constancia en el ticket."))
    a(b.body("Responsable: Área de TI y jefatura del usuario."))
    a(b.blank())

    a(b.section("Paso 4. Recepción del equipo y respaldo"))
    a(b.body("Antes de cualquier intervención, el técnico verifica identidad del activo y realiza o confirma el respaldo de información."))
    a(b.body("Se registra:"))
    for item in [
        "Número de serie, etiqueta de inventario y estado físico de ingreso",
        "Accesorios entregados (cargador, funda, stylus, mouse, candado)",
        "Existencia de respaldo válido o ejecución del respaldo autorizado",
        "Novedades reportadas por el usuario",
    ]:
        a(b.item(item))
    a(
        b.body(
            "No se inicia el mantenimiento si no existe respaldo o autorización expresa de continuar cuando el equipo no almacene información institucional."
        )
    )
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("Paso 5. Diagnóstico inicial"))
    a(b.body("TI ejecuta una evaluación de entrada para establecer la línea base del equipo:"))
    for item in [
        "Encendido, POST y tiempos de arranque",
        "Temperatura, ventilación y ruidos anómalos",
        "Estado de batería (laptops y tablets)",
        "Pantalla, teclado, puertos, audio y conectividad",
        "Espacio en disco, memoria y carga de procesos",
        "Vigencia de antivirus, parches y licencias",
        "Cumplimiento de políticas de dominio y cifrado, cuando aplique",
    ]:
        a(b.item(item))
    a(
        b.body(
            "Si el diagnóstico evidencia daño estructural, oxidación severa, placa fuera de servicio o costo de reparación "
            "no justificado, se documenta y se continúa al Paso 9 (derivación)."
        )
    )
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("Paso 6. Mantenimiento preventivo de hardware"))
    a(b.body("Con el equipo respaldado y diagnosticado, se ejecuta el checklist de hardware, según el tipo de dispositivo:"))
    for item in [
        "Limpieza externa e interna de polvo en salidas de aire, ventiladores y disipadores",
        "Revisión de tornillería, bisagras, tapa y estado de la carcasa",
        "Inspección de puertos USB, HDMI, audio, carga y lector, cuando aplique",
        "Verificación de teclado, touchpad, cámara y micrófono",
        "Prueba de batería y adaptador de corriente",
        "Limpieza de pantalla y, en tablets, revisión de digitalizador y funda",
        "Revisión de periféricos asociados al puesto (monitor, dock, mouse)",
    ]:
        a(b.item(item))
    a(b.body("Las piezas de desgaste se reportan. Si se requiere recambio y no hay stock, se genera el requerimiento técnico."))
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("Paso 7. Mantenimiento preventivo de software y seguridad"))
    a(b.body("TI ejecuta el checklist lógico, sin instalar software no autorizado:"))
    for item in [
        "Actualización del sistema operativo y parches de seguridad pendientes",
        "Actualización de antivirus o EDR y ejecución de análisis completo",
        "Limpieza de temporales, papelera, archivos huérfanos y punto de restauración",
        "Verificación de espacio en disco y desfragmentación o optimización, según el tipo de unidad",
        "Validación de Microsoft 365, correo corporativo y software autorizado",
        "Revisión de cuenta de dominio, políticas de GPO y bloqueo de pantalla",
        "Verificación de cifrado, copias de seguridad y que no existan cuentas locales no controladas",
        "Eliminación de software no institucional detectado, con constancia en el informe",
    ]:
        a(b.item(item))
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("Paso 8. Pruebas de funcionamiento"))
    a(b.body("Al concluir hardware y software, se valida que el equipo quede operativo para el usuario:"))
    for item in [
        "Encendido y apagado correctos",
        "Inicio de sesión corporativo",
        "Navegación, correo e impresoras de red, cuando aplique",
        "Audio, cámara y conectividad Wi-Fi o cableada",
        "En tablets: aplicaciones de operación, GPS o periféricos de campo, cuando aplique",
    ]:
        a(b.item(item))
    a(b.body("Los resultados se marcan en el checklist como conforme o no conforme."))
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("Paso 9. Evaluación de resultado y derivación"))
    a(b.body("Con base en el diagnóstico y las pruebas, TI clasifica el resultado:"))
    a(b.section("Equipo conforme"))
    a(b.body("Se entrega al usuario, se firma el acta y se programa el próximo mantenimiento a seis meses."))
    a(b.section("Requiere correctivo menor"))
    a(b.body("TI repara en el mismo ciclo (cambio de pieza en stock, reimagen parcial, ajuste de configuración) y repite las pruebas."))
    a(b.section("No tiene arreglo o no es viable reparar"))
    a(
        b.body(
            "TI emite informe técnico y deriva al procedimiento TI-PRO-003. Cliente interno: TI solicita la compra. "
            "Cliente externo: el área solicita directamente a Compras por correo, con el informe de TI. "
            "En ambos casos Gerencia aprueba antes de la orden de compra."
        )
    )
    a(b.body("Mientras se gestiona el reemplazo, TI evaluará la asignación de un equipo de respaldo, si existe inventario."))
    a(b.body("Responsable: Área de TI. Derivación: conforme a TI-PRO-003."))
    a(b.blank())

    a(b.section("Paso 10. Documentación del servicio"))
    a(b.body("Toda intervención preventiva debe dejar evidencia en:"))
    for item in [
        "Checklist de mantenimiento (hardware, software y pruebas), con fecha y nombre del técnico",
        "Informe de hallazgos y acciones ejecutadas",
        "Registro de inventario: fecha de mantenimiento, resultado y próximo vencimiento",
        "Ticket actualizado con tiempos de inicio y fin",
        "Cuando exista derivación a compra, número de informe técnico asociado",
    ]:
        a(b.item(item))
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("Paso 11. Entrega al usuario"))
    a(b.body("El usuario recibe el equipo y firma el Acta de Mantenimiento, donde constarán:"))
    for item in [
        "Equipo e inventario",
        "Accesorios devueltos",
        "Trabajos realizados",
        "Resultado (conforme, correctivo o derivado a compra)",
        "Fecha y hora de entrega",
        "Responsable de TI y usuario receptor",
        "Fecha programada del próximo mantenimiento (más seis meses)",
    ]:
        a(b.item(item))
    a(b.body("Responsable: Área de TI y usuario receptor."))
    a(b.blank())

    a(b.section("Paso 12. Cierre e informe de cumplimiento"))
    a(b.body("TI cierra el ticket y actualiza:"))
    for item in [
        "Inventario de activos y bitácora de mantenimiento",
        "Calendario del siguiente ciclo semestral",
        "Indicador de equipos intervenidos versus planificados",
        "Hallazgos repetitivos que requieran cambio de estándar o de proveedor",
    ]:
        a(b.item(item))
    a(
        b.body(
            "Al cierre de cada ciclo semestral, TI emite un informe resumido a Gerencia con cobertura alcanzada, "
            "equipos derivados a compra, incidentes detectados y recomendaciones."
        )
    )
    a(b.body("Responsable: Área de TI. Informe: Gerencia."))
    a(b.blank())

    a(b.section("5. DIAGRAMA DEL PROCESO"))
    a(
        b.body(
            "El diagrama de flujo se mantiene en el archivo TI-PRO-004-Diagrama.drawio, para edición en Draw.io / diagrams.net. "
            "Representa la programación semestral, la atención a cliente interno y externo, la ejecución del preventivo, "
            "la decisión de conformidad o derivación al proceso de adquisición TI-PRO-003, y el cierre con acta e inventario."
        )
    )
    a(b.blank())
    a(b.blank())
    return els


def mx_cell(cid, value, style, x, y, w, h, parent="1"):
    value_esc = (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("\n", "&#xa;")
    )
    return (
        f'        <mxCell id="{cid}" value="{value_esc}" style="{style}" vertex="1" parent="{parent}">\n'
        f'          <mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/>\n'
        f"        </mxCell>\n"
    )


def mx_edge(eid, source, target, style="", label=""):
    if not style:
        style = (
            "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;"
            "strokeWidth=2;strokeColor=#2C3E50;endArrow=block;endFill=1;"
        )
    lab = f' value="{label}"' if label else ' value=""'
    return (
        f'        <mxCell id="{eid}"{lab} style="{style}" edge="1" parent="1" source="{source}" target="{target}">\n'
        f'          <mxGeometry relative="1" as="geometry"/>\n'
        f"        </mxCell>\n"
    )


def build_drawio() -> Path:
    title_s = "rounded=0;whiteSpace=wrap;html=1;fillColor=#1F4E79;fontColor=#FFFFFF;fontStyle=1;fontSize=16;strokeColor=#1F4E79;"
    sub_s = "rounded=0;whiteSpace=wrap;html=1;fillColor=#D4E6F1;fontColor=#1F4E79;fontSize=12;strokeColor=#5DADE2;"
    legend_s = "rounded=1;whiteSpace=wrap;html=1;fillColor=#F8F9F9;strokeColor=#7F8C8D;fontSize=11;align=left;spacingLeft=12;arcSize=6;"
    start_s = "ellipse;whiteSpace=wrap;html=1;fillColor=#148F77;fontColor=#FFFFFF;strokeColor=#0E6655;fontStyle=1;fontSize=12;strokeWidth=2;"
    end_s = "ellipse;whiteSpace=wrap;html=1;fillColor=#7B241C;fontColor=#FFFFFF;strokeColor=#641E16;fontStyle=1;fontSize=12;strokeWidth=2;"
    dec_s = "rhombus;whiteSpace=wrap;html=1;fillColor=#FCF3CF;strokeColor=#B7950B;fontColor=#7D6608;fontStyle=1;fontSize=11;strokeWidth=2;"
    ti_s = "rounded=1;whiteSpace=wrap;html=1;fillColor=#D7BDE2;strokeColor=#6C3483;fontColor=#4A235A;fontSize=11;strokeWidth=2;arcSize=10;"
    int_s = "rounded=1;whiteSpace=wrap;html=1;fillColor=#AED6F1;strokeColor=#1F4E79;fontColor=#1B4F72;fontSize=11;strokeWidth=2;arcSize=10;"
    ext_s = "rounded=1;whiteSpace=wrap;html=1;fillColor=#A3E4D7;strokeColor=#0E6655;fontColor=#0E6655;fontSize=11;strokeWidth=2;arcSize=10;"
    ok_s = "rounded=1;whiteSpace=wrap;html=1;fillColor=#ABEBC6;strokeColor=#1D8348;fontColor=#145A32;fontSize=11;strokeWidth=2;arcSize=10;"
    warn_s = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FAD7A0;strokeColor=#B9770E;fontColor=#6E2C00;fontSize=11;strokeWidth=2;arcSize=10;"
    acq_s = "rounded=1;whiteSpace=wrap;html=1;fillColor=#F5B7B1;strokeColor=#922B21;fontColor=#641E16;fontSize=11;strokeWidth=2;arcSize=10;"
    lane_i = "rounded=0;whiteSpace=wrap;html=1;fillColor=#EAF2F8;strokeColor=#1F4E79;strokeWidth=2;fontStyle=1;fontSize=13;fontColor=#1F4E79;verticalAlign=top;spacingTop=8;"
    lane_e = "rounded=0;whiteSpace=wrap;html=1;fillColor=#E8F8F5;strokeColor=#117A65;strokeWidth=2;fontStyle=1;fontSize=13;fontColor=#0E6655;verticalAlign=top;spacingTop=8;"
    lane_t = "rounded=0;whiteSpace=wrap;html=1;fillColor=#F5EEF8;strokeColor=#6C3483;strokeWidth=2;fontStyle=1;fontSize=13;fontColor=#4A235A;verticalAlign=top;spacingTop=8;"
    yes_e = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=2;strokeColor=#1D8348;endArrow=block;endFill=1;fontColor=#1D8348;fontStyle=1;fontSize=11;"
    no_e = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=2;strokeColor=#922B21;endArrow=block;endFill=1;fontColor=#922B21;fontStyle=1;fontSize=11;"
    def_e = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=2;strokeColor=#2C3E50;endArrow=block;endFill=1;"
    blue_e = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=2;strokeColor=#1F4E79;endArrow=block;endFill=1;fontColor=#1F4E79;fontStyle=1;"
    green_e = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=2;strokeColor=#0E6655;endArrow=block;endFill=1;fontColor=#0E6655;fontStyle=1;"

    c = []
    c.append('        <mxCell id="0"/>\n')
    c.append('        <mxCell id="1" parent="0"/>\n')

    c.append(mx_cell("t1", "MANTENIMIENTO PREVENTIVO DE EQUIPOS INFORMÁTICOS — CICLO SEMESTRAL", title_s, 40, 20, 1560, 48))
    c.append(mx_cell("t2", "TI-PRO-004  ·  v1.0  ·  Analista de sistemas / Área de TI  ·  Clientes internos (Quito y GYE) y externos (guardias / supervisores)", sub_s, 40, 68, 1560, 36))
    c.append(mx_cell("lg", "Leyenda:  Lila = TI    Azul = Cliente interno    Verde = Cliente externo    Verde claro = conforme    Naranja = correctivo    Rojo = deriva a compra TI-PRO-003    Rombo = decisión", legend_s, 40, 112, 1560, 36))

    c.append(mx_cell("start", "INICIO\nCiclo de 6 meses\no vencimiento en inventario", start_s, 680, 170, 240, 90))
    c.append(mx_cell("p1", "1. TI elabora el plan semestral\n(inventario, sede, usuario, fecha objetivo)", ti_s, 640, 290, 320, 70))
    c.append(mx_cell("p2", "2. TI genera ticket de mantenimiento\npreventivo (actividad planificada)", ti_s, 640, 390, 320, 70))
    c.append(mx_cell("p3", "3. Notifica al usuario / jefatura\n(mínimo 5 días hábiles)", ti_s, 640, 490, 320, 70))
    c.append(mx_cell("tipo", "¿Cliente interno\no externo?", dec_s, 670, 590, 260, 110))

    c.append(mx_cell("laneI", "CLIENTE INTERNO — Personal administrativo Quito y Guayaquil", lane_i, 40, 730, 740, 220))
    c.append(mx_cell("laneE", "CLIENTE EXTERNO — Guardias, supervisores y operación de campo", lane_e, 860, 730, 740, 220))
    c.append(mx_cell("i1", "Coordina cita en oficina o taller TI.\nUsuario deja el equipo disponible.", int_s, 80, 790, 360, 70))
    c.append(mx_cell("i2", "Recibe el activo con accesorios\ny novedades de uso.", int_s, 80, 880, 360, 50))
    c.append(mx_cell("e1", "Coordina retiro, visita o equipo de\nrespaldo para no detener la operación.", ext_s, 900, 790, 380, 70))
    c.append(mx_cell("e2", "Supervisión de campo facilita tablets\ny laptops en la ventana acordada.", ext_s, 900, 880, 380, 50))

    c.append(mx_cell("laneT", "EJECUCIÓN TÉCNICA — Área de TI (aplica a ambos clientes)", lane_t, 40, 980, 1560, 520))
    c.append(mx_cell("p4", "4. Recepción del activo e inventario de ingreso.\nRespaldo de información (obligatorio antes de intervenir).", ti_s, 80, 1040, 460, 70))
    c.append(mx_cell("p5", "5. Diagnóstico inicial\n(hardware, batería, disco, parches, políticas).", ti_s, 600, 1040, 420, 70))
    c.append(mx_cell("d0", "¿El equipo es\nviable de mantener?", dec_s, 1120, 1020, 260, 110))

    c.append(mx_cell("p6", "6. Preventivo de hardware\nLimpieza, ventilación, puertos,\nbatería, pantalla y periféricos.", ti_s, 80, 1180, 360, 90))
    c.append(mx_cell("p7", "7. Preventivo de software y seguridad\nParches, antivirus, disco, M365,\ndominio, cifrado y software no autorizado.", ti_s, 500, 1180, 420, 90))
    c.append(mx_cell("p8", "8. Pruebas funcionales\nArranque, correo, red, apps de operación.", ti_s, 980, 1180, 380, 90))

    c.append(mx_cell("d1", "¿Equipo queda\noperativo?", dec_s, 680, 1310, 240, 110))
    c.append(mx_cell("corr", "Correctivo menor\n(pieza en stock / ajuste).\nRepite pruebas.", warn_s, 1000, 1325, 280, 80))
    c.append(mx_cell("deriva", "No viable / sin arreglo.\nInforme técnico y deriva\na TI-PRO-003 (adquisición).", acq_s, 1320, 1180, 250, 90))

    c.append(mx_cell("p10", "10. Documenta checklist, hallazgos,\ninventario y próximo vencimiento +6 meses.", ti_s, 80, 1470, 420, 70))
    c.append(mx_cell("p11", "11. Entrega al usuario con\nActa de Mantenimiento firmada.", ok_s, 560, 1470, 360, 70))
    c.append(mx_cell("p12", "12. Cierra ticket e incluye el caso\nen el informe semestral a Gerencia.", ti_s, 980, 1470, 400, 70))

    c.append(mx_cell("finok", "FIN\nEquipo conforme.\nPróximo ciclo: +6 meses", start_s, 560, 1600, 280, 80))
    c.append(mx_cell("finacq", "FIN de este procedimiento.\nContinúa en TI-PRO-003", end_s, 1320, 1325, 250, 80))

    c.append(mx_edge("a1", "start", "p1", def_e))
    c.append(mx_edge("a2", "p1", "p2", def_e))
    c.append(mx_edge("a3", "p2", "p3", def_e))
    c.append(mx_edge("a4", "p3", "tipo", def_e))
    c.append(mx_edge("a5", "tipo", "i1", blue_e, "Interno"))
    c.append(mx_edge("a6", "tipo", "e1", green_e, "Externo"))
    c.append(mx_edge("a7", "i1", "i2", def_e))
    c.append(mx_edge("a8", "e1", "e2", def_e))
    c.append(mx_edge("a9", "i2", "p4", def_e))
    c.append(mx_edge("a10", "e2", "p4", def_e))
    c.append(mx_edge("a11", "p4", "p5", def_e))
    c.append(mx_edge("a12", "p5", "d0", def_e))
    c.append(mx_edge("a13", "d0", "p6", yes_e, "Sí"))
    c.append(mx_edge("a14", "d0", "deriva", no_e, "No"))
    c.append(mx_edge("a15", "p6", "p7", def_e))
    c.append(mx_edge("a16", "p7", "p8", def_e))
    c.append(mx_edge("a17", "p8", "d1", def_e))
    c.append(mx_edge("a18", "d1", "p10", yes_e, "Sí"))
    c.append(mx_edge("a19", "d1", "corr", no_e, "No"))
    c.append(mx_edge("a20", "corr", "p8", warn_s.replace("rounded=1;whiteSpace=wrap;html=1;fillColor=#FAD7A0;strokeColor=#B9770E;fontColor=#6E2C00;fontSize=11;strokeWidth=2;arcSize=10;", "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=2;strokeColor=#B9770E;endArrow=block;endFill=1;dashed=1;")))
    c.append(mx_edge("a21", "p10", "p11", def_e))
    c.append(mx_edge("a22", "p11", "p12", def_e))
    c.append(mx_edge("a23", "p12", "finok", yes_e))
    c.append(mx_edge("a24", "deriva", "finacq", no_e))

    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<mxfile host="app.diagrams.net" agent="Cursor" version="22.1.0" type="device">\n'
        '  <diagram id="ti-pro-004" name="Mantenimiento preventivo semestral">\n'
        '    <mxGraphModel dx="1400" dy="900" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" '
        'arrows="1" fold="1" page="1" pageScale="1" pageWidth="1680" pageHeight="1760" math="0" shadow="0">\n'
        "      <root>\n"
        + "".join(c)
        + "      </root>\n"
        "    </mxGraphModel>\n"
        "  </diagram>\n"
        "</mxfile>\n"
    )
    OUT_DRAWIO.write_text(xml, encoding="utf-8")
    return OUT_DRAWIO


def build_markdown() -> Path:
    md = """# Procedimiento para el mantenimiento preventivo de equipos informáticos

**Código:** TI-PRO-004  
**Versión:** 1.0  
**Periodicidad:** cada 6 meses  
**Diagrama Draw.io:** `TI-PRO-004-Diagrama.drawio`

## Objetivo

Planificar, ejecutar y documentar el mantenimiento preventivo de los equipos informáticos cada seis meses, para conservar disponibilidad, rendimiento y seguridad. Si el equipo no tiene arreglo, se deriva al **TI-PRO-003**.

## Clientes

| Tipo | Quiénes | Cómo se interviene |
| --- | --- | --- |
| Interno | Personal administrativo Quito y Guayaquil | Cita en oficina o taller TI |
| Externo | Guardias, supervisores y operación de campo | Retiro, visita o equipo de respaldo (tablets/laptops) |

## Pasos

1. TI arma el plan semestral desde el inventario.
2. Abre ticket de mantenimiento preventivo (no espera una falla).
3. Notifica con al menos 5 días hábiles.
4. Recibe el equipo y **respalda** la información.
5. Diagnóstico inicial.
6. Preventivo de hardware (limpieza, ventilación, puertos, batería).
7. Preventivo de software y seguridad (parches, antivirus, políticas).
8. Pruebas de funcionamiento.
9. Resultado: conforme, correctivo menor, o derivación a compra (TI-PRO-003).
10. Documenta checklist e inventario.
11. Entrega con **Acta de Mantenimiento**.
12. Cierra el ticket e informa cumplimiento a Gerencia. Programa el siguiente ciclo a +6 meses.

## Cómo abrir el diagrama

1. Entre a https://app.diagrams.net
2. *Open Existing Diagram* → `TI-PRO-004-Diagrama.drawio`
"""
    OUT_MD.write_text(md, encoding="utf-8")
    return OUT_MD


def build_html() -> Path:
    html = """<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>TI-PRO-004 — Mantenimiento preventivo semestral</title>
  <style>
    body { margin:0; font-family:"Segoe UI", Calibri, Arial, sans-serif; background:#f4f7fb; color:#1c2833; line-height:1.45; }
    header { background:#1f4e79; color:#fff; padding:28px 32px; }
    header p { margin:6px 0 0; opacity:.9; }
    main { max-width:1100px; margin:0 auto; padding:24px; }
    h2 { color:#1f4e79; border-bottom:2px solid #d6eaf8; padding-bottom:6px; }
    .grid { display:grid; grid-template-columns:1fr 1fr; gap:16px; }
    .card { border-radius:12px; padding:16px 18px; border:2px solid; background:#fff; }
    .interno { border-color:#1f4e79; background:#eaf2f8; }
    .externo { border-color:#0e6655; background:#e8f8f5; }
    ol.pasos { padding-left:20px; }
    ol.pasos li { margin:8px 0; }
    .note { background:#fcf3cf; border-left:4px solid #b7950b; padding:12px 14px; margin:16px 0; }
    footer { text-align:center; color:#5d6d7e; padding:24px; font-size:13px; }
    @media (max-width:800px) { .grid { grid-template-columns:1fr; } }
  </style>
</head>
<body>
  <header>
    <h1>Mantenimiento preventivo de equipos informáticos</h1>
    <p>Código TI-PRO-004 · Versión 1.0 · Periodicidad: cada 6 meses</p>
  </header>
  <main>
    <div class="note"><strong>Regla de derivación:</strong> si el preventivo determina que el equipo no tiene arreglo, se emite informe técnico y se continúa en el procedimiento de adquisición TI-PRO-003.</div>
    <h2>Los dos canales</h2>
    <div class="grid">
      <article class="card interno">
        <h3>Cliente interno</h3>
        <p>Personal administrativo de Quito y Guayaquil. Cita en oficina o taller de TI.</p>
      </article>
      <article class="card externo">
        <h3>Cliente externo</h3>
        <p>Guardias y supervisores. Tablets y laptops de campo. Se coordina retiro, visita o equipo de respaldo.</p>
      </article>
    </div>
    <h2>Pasos del ciclo semestral</h2>
    <ol class="pasos">
      <li>Plan desde inventario.</li>
      <li>Ticket preventivo (lo abre TI).</li>
      <li>Notificación (mínimo 5 días hábiles).</li>
      <li>Recepción y respaldo.</li>
      <li>Diagnóstico inicial.</li>
      <li>Preventivo de hardware.</li>
      <li>Preventivo de software y seguridad.</li>
      <li>Pruebas funcionales.</li>
      <li>Conforme, correctivo menor o derivación a compra.</li>
      <li>Documentación e inventario.</li>
      <li>Acta de mantenimiento y entrega.</li>
      <li>Cierre, informe a Gerencia y próximo ciclo a +6 meses.</li>
    </ol>
  </main>
  <footer>Documento interno · Área de Tecnología de la Información</footer>
</body>
</html>
"""
    OUT_HTML.write_text(html, encoding="utf-8")
    return OUT_HTML


def build_docx() -> Path:
    b = Builder()
    return render_from_template(content(b), OUT_DOCX)


if __name__ == "__main__":
    print(build_docx())
    print(build_drawio())
    print(build_markdown())
    print(build_html())
