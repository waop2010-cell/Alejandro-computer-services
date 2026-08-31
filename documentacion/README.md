# Procedimientos de Tecnología de la Información

Documentos internos del Área de TI, en plantilla institucional (encabezado, pie, Calibri justificado).

## TI-PRO-003 — Adquisición de tecnología entre TI y Compras

| Archivo | Uso |
| --- | --- |
| `TI-PRO-003-Proceso-Adquisicion-Tecnologia.docx` | Procedimiento completo. |
| `TI-PRO-003-Pasos-a-seguir.md` | Resumen de pasos. |
| `TI-PRO-003-Diagrama.drawio` | Diagrama Draw.io. |
| `proceso_adquisicion_ti_compras.zip` | Paquete de descarga. |

## TI-PRO-004 — Mantenimiento preventivo de equipos informáticos (cada 6 meses)

| Archivo | Uso |
| --- | --- |
| `TI-PRO-004-Mantenimiento-Equipos-Informaticos.docx` | Procedimiento completo, mismo formato que TI-PRO-003. |
| `TI-PRO-004-Pasos-a-seguir.md` | Resumen de pasos. |
| `TI-PRO-004-Diagrama.drawio` | Diagrama Draw.io del ciclo semestral. |
| `vista-previa-mantenimiento.html` | Vista rápida en el navegador. |
| `proceso_mantenimiento_equipos.zip` | Paquete de descarga. |

## Cómo abrir los diagramas

1. Entre a [https://app.diagrams.net](https://app.diagrams.net).
2. Elija **Open Existing Diagram**.
3. Seleccione el archivo `.drawio` correspondiente.
4. Para exportar: **File → Export as → PNG** o **PDF**.

## Presentación — Manejo y uso de Google Gemini (15 min)

| Archivo | Uso |
| --- | --- |
| `Manejo-y-Uso-de-Google-Gemini.pptx` | Presentación completa (Aditec + sesión Gemini), gama corporativa azul `#0A529C` y rojo `#DB261D`. |
| `Guion-Gemini-15-minutos.md` | Tiempos sugeridos y respuestas a preguntas frecuentes. |
| `presentacion_gemini_analista.zip` | Paquete de descarga. |

## Regenerar

```bash
python3 documentacion/generar_documentos.py
python3 documentacion/generar_mantenimiento.py
python3 documentacion/generar_presentacion_gemini.py
```
