#!/usr/bin/env python3
"""Genera el procedimiento Word copiando la plantilla DT360 adjunta
(encabezado, pie, marcas de agua, márgenes, Calibri justificado)."""

from __future__ import annotations

import io
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML = "http://www.w3.org/XML/1998/namespace"

# Prefijos que Word espera (sobre todo r:id en encabezados y pies).
ET.register_namespace("w", W)
ET.register_namespace("w14", W14)
ET.register_namespace("r", R)
ET.register_namespace("mc", MC)

NS = {"w": W}

HERE = Path(__file__).resolve().parent
TEMPLATE_CANDIDATES = [
    Path("/home/ubuntu/.cursor/projects/workspace/uploads/PROCESO_DE_ADQUISICI_N_DE_TECNOLOG_A_ENTRE_TI_Y_COMPRAS_26ab.docx"),
    HERE / "plantilla-DT360.docx",
]
OUT_DOCX = HERE / "TI-PRO-003-Proceso-Adquisicion-Tecnologia.docx"


def qn(tag: str) -> str:
    return f"{{{W}}}{tag}"


def w14(tag: str) -> str:
    return f"{{{W14}}}{tag}"


class Builder:
    def __init__(self) -> None:
        self.pid = 1

    def _pid(self) -> str:
        self.pid += 1
        return f"{self.pid:08X}"

    def _ppr(self, bold_mark: bool) -> ET.Element:
        ppr = ET.Element(qn("pPr"))
        sp = ET.SubElement(ppr, qn("spacing"))
        sp.set(qn("line"), "240")
        sp.set(qn("lineRule"), "auto")
        jc = ET.SubElement(ppr, qn("jc"))
        jc.set(qn("val"), "both")
        rpr = ET.SubElement(ppr, qn("rPr"))
        fonts = ET.SubElement(rpr, qn("rFonts"))
        fonts.set(qn("ascii"), "Calibri")
        fonts.set(qn("hAnsi"), "Calibri")
        fonts.set(qn("cs"), "Calibri")
        if bold_mark:
            ET.SubElement(rpr, qn("b"))
            ET.SubElement(rpr, qn("bCs"))
        color = ET.SubElement(rpr, qn("color"))
        color.set(qn("val"), "000000")
        color.set(qn("themeColor"), "text1")
        lang = ET.SubElement(rpr, qn("lang"))
        lang.set(qn("val"), "es-EC")
        return ppr

    def _rpr(self, bold: bool) -> ET.Element:
        rpr = ET.Element(qn("rPr"))
        fonts = ET.SubElement(rpr, qn("rFonts"))
        fonts.set(qn("ascii"), "Calibri")
        fonts.set(qn("hAnsi"), "Calibri")
        fonts.set(qn("cs"), "Calibri")
        if bold:
            ET.SubElement(rpr, qn("b"))
            ET.SubElement(rpr, qn("bCs"))
        color = ET.SubElement(rpr, qn("color"))
        color.set(qn("val"), "000000")
        color.set(qn("themeColor"), "text1")
        lang = ET.SubElement(rpr, qn("lang"))
        lang.set(qn("val"), "es-EC")
        return rpr

    def _run(self, text: str, bold: bool = False) -> ET.Element:
        r = ET.Element(qn("r"))
        r.append(self._rpr(bold))
        t = ET.SubElement(r, qn("t"))
        if text[:1].isspace() or text[-1:].isspace():
            t.set(f"{{{XML}}}space", "preserve")
        t.text = text
        return r

    def para(self, parts, *, heading: bool = False) -> ET.Element:
        """parts: str or list of (text, bold)."""
        p = ET.Element(qn("p"))
        p.set(w14("paraId"), self._pid())
        p.set(w14("textId"), "77777777")
        if isinstance(parts, str):
            runs = [(parts, heading)]
            bold_mark = heading
        else:
            runs = list(parts)
            bold_mark = heading or all(b for _, b in runs)
        p.append(self._ppr(bold_mark))
        if not runs or (len(runs) == 1 and runs[0][0] == ""):
            return p
        for text, bold in runs:
            p.append(self._run(text, bold))
        return p

    def blank(self) -> ET.Element:
        return self.para("")

    def title(self, text: str) -> ET.Element:
        return self.para(text, heading=True)

    def section(self, text: str) -> ET.Element:
        return self.para(text, heading=True)

    def meta(self, label: str, value: str) -> ET.Element:
        return self.para([(label, True), (value, False)])

    def body(self, text: str) -> ET.Element:
        return self.para(text)

    def item(self, text: str) -> ET.Element:
        if text and not text.endswith("."):
            text = text + "."
        return self.para(text)


def content(b: Builder) -> list[ET.Element]:
    els: list[ET.Element] = []
    a = els.append

    a(b.blank())
    a(b.title("PROCEDIMIENTO PARA LA GESTIÓN DE ADQUISICIONES TECNOLÓGICAS"))
    a(b.meta("Código:", " TI-PRO-003"))
    a(b.meta("Versión:", " 1.1"))
    a(b.meta("Área Responsable:", " Tecnología de la Información (TI)"))
    a(b.blank())

    a(b.section("1. OBJETIVO"))
    a(
        b.body(
            "Establecer el procedimiento para la adquisición de equipos tecnológicos, accesorios y demás recursos informáticos, "
            "garantizando que todas las compras se realicen de forma planificada, controlada y alineada a las necesidades de la "
            "organización, diferenciando el canal de cliente interno y el canal de cliente externo."
        )
    )
    a(b.blank())

    a(b.section("2. ALCANCE"))
    a(
        b.body(
            "Aplica a todas las adquisiciones de bienes y servicios tecnológicos requeridos por la organización, incluyendo:"
        )
    )
    for item in [
        "Computadoras y laptops",
        "Monitores",
        "Servidores",
        "Equipos de red",
        "Impresoras",
        "Celulares corporativos",
        "Tablets",
        "Equipos de videovigilancia",
        "Accesorios tecnológicos",
    ]:
        a(b.item(item))
    a(b.blank())
    a(b.section("Definición de clientes"))
    a(
        b.body(
            "El proceso tiene dos entradas según quién necesita el equipo. Esta clasificación determina quién solicita la compra al Área de Compras."
        )
    )
    a(b.section("Cliente interno"))
    a(b.body("Personal administrativo de oficinas en Quito y Guayaquil."))
    a(b.body("Equipos habituales: laptops, computadoras, monitores y accesorios de oficina."))
    a(b.section("Cliente externo"))
    a(b.body("Todos los clientes por fuera: guardias y supervisores."))
    a(b.body("Equipos habituales: tablets y laptops de operación."))
    a(
        b.body(
            "Si se daña un dispositivo de cliente externo, la compra se gestiona por el flujo de cliente externo."
        )
    )
    a(b.blank())

    a(b.section("3. RESPONSABILIDADES"))
    a(b.section("Área Solicitante"))
    for item in [
        "Identificar la necesidad o reportar la falla mediante ticket",
        "Justificar la adquisición",
        "Cliente interno: esperar la evaluación de TI; no gestionar la compra por su cuenta",
        "Cliente externo: una vez exista el informe técnico de TI, solicitar la compra directamente a Compras por correo",
        "Aprobar la recepción del bien cuando corresponda y firmar el Acta de Entrega",
    ]:
        a(b.item(item))
    a(b.blank())
    a(b.section("Área de Tecnología de la Información"))
    for item in [
        "Analizar la necesidad",
        "Evaluar si el equipo tiene arreglo o si existe inventario reutilizable",
        "Definir las especificaciones técnicas y emitir el informe técnico",
        "Validar compatibilidad",
        "Cliente interno: solicitar la compra a Compras cuando no haya arreglo ni stock",
        "Cliente externo: entregar el informe técnico para que el solicitante gestione la compra con Compras",
        "Evaluar propuestas en su componente técnico",
        "Recibir o coordinar el equipo, configurarlo y prepararlo",
        "Actualizar el inventario tecnológico",
    ]:
        a(b.item(item))
    a(b.section("Área de Compras"))
    for item in [
        "Recibir la solicitud de compra (de TI en flujo interno, o del cliente o área por correo en flujo externo)",
        "Solicitar cotizaciones comerciales",
        "Negociar con proveedores",
        "Emitir la Orden de Compra una vez exista aprobación de Gerencia",
        "Dar seguimiento al proveedor",
        "Cliente interno: entregar el equipo a TI",
        "Cliente externo: entregar el equipo directamente al usuario o enviar correo de coordinación",
        "Coordinar la entrega",
    ]:
        a(b.item(item))
    a(b.section("Gerencia"))
    a(
        b.body(
            "Revisar y aprobar las adquisiciones en ambos flujos, conforme a la política de compras de la empresa. Sin esta aprobación, Compras no emite la Orden de Compra."
        )
    )
    a(b.blank())

    a(b.section("4. DESCRIPCIÓN DEL PROCEDIMIENTO"))
    a(b.section("Paso 1. Identificación de la Necesidad"))
    a(b.body("El área solicitante identifica la necesidad de adquirir un recurso tecnológico."))
    a(b.body("Ejemplos:"))
    for item in [
        "Nuevo colaborador",
        "Equipo dañado",
        "Equipo obsoleto",
        "Incremento de personal",
        "Nuevo proyecto",
        "Dispositivo de operación que no funciona (ejemplo: el usuario indica que no le funciona el dispositivo y genera un ticket)",
    ]:
        a(b.item(item))
    a(b.body("Responsable: Área Solicitante."))
    a(b.blank())

    a(b.section("Paso 2. Generación del Requerimiento"))
    a(b.body("El área solicitante registra un ticket o solicitud formal indicando:"))
    for item in [
        "Justificación",
        "Usuario final",
        "Área",
        "Fecha requerida",
        "Tipo de recurso solicitado",
        "Si se trata de cliente interno o cliente externo",
    ]:
        a(b.item(item))
    a(b.body("Responsable: Área Solicitante."))
    a(b.blank())

    a(b.section("Paso 3. Evaluación Técnica"))
    a(b.body("TI analiza:"))
    for item in [
        "Si existe disponibilidad en inventario",
        "Si el equipo puede ser reutilizado",
        "Si el equipo tiene arreglo",
        "Compatibilidad con la infraestructura",
        "Especificaciones técnicas necesarias",
        "Licencias requeridas",
    ]:
        a(b.item(item))
    a(b.body("Si el equipo tiene arreglo o existe disponibilidad, el equipo será reparado o reasignado. No se inicia compra."))
    a(b.body("Si no existe disponibilidad y no tiene arreglo, se continúa con el proceso de compra."))
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("Paso 4. Elaboración de Especificaciones Técnicas"))
    a(b.body("TI prepara el informe o documento técnico que incluirá, según corresponda:"))
    for item in [
        "Procesador",
        "Memoria RAM",
        "Disco SSD",
        "Sistema Operativo",
        "Garantía",
        "Monitor",
        "Accesorios",
        "Marca recomendada (cuando aplique)",
    ]:
        a(b.item(item))
    a(b.body("Este informe técnico es la base obligatoria para que Compras gestione la adquisición."))
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("Paso 5. Solicitud de Compra y Cotizaciones"))
    a(b.body("La forma de solicitar la compra depende del tipo de cliente."))
    a(b.section("Cliente interno"))
    a(
        b.body(
            "TI solicita la compra al Área de Compras, con el informe técnico. El personal administrativo no gestiona el pedido de forma directa."
        )
    )
    a(b.section("Cliente externo"))
    a(
        b.body(
            "El cliente o su área debe solicitar directamente a Compras, enviando un correo basado en el informe técnico de TI. No espera a que TI haga el pedido en su nombre."
        )
    )
    a(
        b.body(
            "Compras solicita cotizaciones a proveedores autorizados utilizando las especificaciones emitidas por TI. Se recomienda obtener al menos tres cotizaciones cuando la política interna lo requiera."
        )
    )
    a(b.body("Responsable: Área de TI (flujo interno), Área Solicitante (flujo externo) y Área de Compras (cotizaciones)."))
    a(b.blank())

    a(b.section("Paso 6. Evaluación de Proveedores"))
    a(b.section("TI realiza la evaluación técnica de las propuestas considerando:"))
    for item in [
        "Cumplimiento de especificaciones",
        "Garantía",
        "Tiempo de entrega",
        "Soporte técnico",
        "Compatibilidad",
        "Calidad del fabricante",
    ]:
        a(b.item(item))
    a(b.section("Compras analiza:"))
    for item in [
        "Precio",
        "Condiciones comerciales",
        "Forma de pago",
        "Disponibilidad",
    ]:
        a(b.item(item))
    a(b.blank())

    a(b.section("Paso 7. Aprobación"))
    a(
        b.body(
            "En los dos casos, interno y externo, Gerencia revisa la adquisición para su aprobación, conforme a los niveles establecidos por la organización."
        )
    )
    a(b.body("Si se rechaza, se notifica al solicitante y se reformula o se cierra el requerimiento."))
    a(b.body("Una vez aprobada, Compras emitirá la Orden de Compra."))
    a(b.body("Responsable: Gerencia y Área de Compras."))
    a(b.blank())

    a(b.section("Paso 8. Compra"))
    a(b.body("Compras realiza la adquisición con el proveedor seleccionado y da seguimiento hasta la recepción del bien."))
    a(b.body("Responsable: Área de Compras."))
    a(b.blank())

    a(b.section("Paso 9. Recepción del Equipo"))
    a(b.body("La entrega logística depende del tipo de cliente."))
    a(b.section("Cliente interno"))
    a(b.body("Compras entrega el equipo al Área de TI."))
    a(b.section("Cliente externo"))
    a(b.body("Compras entrega el equipo directamente al usuario o envía un correo de notificación y coordinación."))
    a(b.body("Al recibir el bien, TI verificará:"))
    for item in [
        "Estado físico",
        "Número de serie",
        "Modelo",
        "Garantía",
        "Accesorios incluidos",
        "Funcionamiento",
    ]:
        a(b.item(item))
    a(b.body("En caso de inconsistencias, se notificará inmediatamente al proveedor."))
    a(b.body("Responsable: Área de Compras y Área de TI."))
    a(b.blank())

    a(b.section("Paso 10. Configuración"))
    a(b.body("En ambos casos TI prepara el equipo. TI realizará:"))
    for item in [
        "Instalación del sistema operativo",
        "Configuración de Active Directory",
        "Configuración de correo corporativo",
        "Instalación de Microsoft 365",
        "Instalación de software autorizado",
        "Configuración de antivirus",
        "Aplicación de políticas de seguridad",
        "Registro en inventario",
        "Etiquetado del activo",
    ]:
        a(b.item(item))
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("Paso 11. Entrega"))
    a(b.body("El usuario firmará el Acta de Entrega del Equipo, donde constarán:"))
    for item in [
        "Equipo entregado",
        "Accesorios",
        "Estado del equipo",
        "Fecha",
        "Responsable de TI",
        "Usuario receptor",
    ]:
        a(b.item(item))
    a(b.body("Responsable: Área de TI y usuario receptor."))
    a(b.blank())

    a(b.section("Paso 12. Cierre"))
    a(b.body("TI actualizará:"))
    for item in [
        "Inventario de activos",
        "Base de datos de equipos",
        "Licencias asignadas",
        "Ticket de solicitud",
    ]:
        a(b.item(item))
    a(b.body("Finalmente se procederá al cierre del requerimiento."))
    a(b.body("Responsable: Área de TI."))
    a(b.blank())

    a(b.section("5. DIAGRAMA DEL PROCESO"))
    a(
        b.body(
            "El diagrama de flujo se mantiene en el archivo TI-PRO-003-Diagrama.drawio, para edición en Draw.io / diagrams.net. "
            "El diagrama representa los dos flujos: cliente interno (personal administrativo Quito y Guayaquil) y cliente externo (guardias, supervisores y clientes por fuera)."
        )
    )
    a(b.blank())
    a(b.blank())
    return els


def find_template() -> Path:
    for p in TEMPLATE_CANDIDATES:
        if p.exists():
            return p
    raise FileNotFoundError("No se encontró la plantilla Word DT360.")


def build() -> Path:
    template = find_template()
    plantilla_repo = HERE / "plantilla-DT360.docx"
    if template.resolve() != plantilla_repo.resolve():
        shutil.copy2(template, plantilla_repo)

    b = Builder()
    paragraphs = content(b)

    with zipfile.ZipFile(plantilla_repo) as zin:
        doc_xml = zin.read("word/document.xml")
        others = {name: zin.read(name) for name in zin.namelist() if name != "word/document.xml"}

    root = ET.fromstring(doc_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise RuntimeError("El documento plantilla no tiene cuerpo.")
    sect = body.find("w:sectPr", NS)
    if sect is None:
        raise RuntimeError("El documento plantilla no tiene sectPr.")

    for child in list(body):
        if child.tag != qn("sectPr"):
            body.remove(child)
    for i, p in enumerate(paragraphs):
        body.insert(i, p)

    buf = io.BytesIO()
    ET.ElementTree(root).write(buf, encoding="utf-8", xml_declaration=True)
    new_doc = buf.getvalue()
    # Word expects standalone declaration similar to original
    if new_doc.startswith(b"<?xml"):
        new_doc = new_doc.replace(
            b"<?xml version='1.0' encoding='utf-8'?>",
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            1,
        )
        new_doc = new_doc.replace(
            b'<?xml version="1.0" encoding="utf-8"?>',
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            1,
        )

    tmp = OUT_DOCX.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in others.items():
            zout.writestr(name, data)
        zout.writestr("word/document.xml", new_doc)
    tmp.replace(OUT_DOCX)
    return OUT_DOCX


if __name__ == "__main__":
    path = build()
    print(path)
