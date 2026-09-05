"""Generación de PPTX editable vía html-to-pptx (Playwright Chromium) o nativo python-pptx con soporte de gráficos."""
from __future__ import annotations
import os, tempfile
from typing import Any

_SLIDE_CSS = """
* { box-sizing: border-box; margin: 0; }
body { font-family: var(--ow-font-body); color: var(--ow-text); }
.slide { width: 1280px; height: 720px; padding: 56px 64px; background: var(--ow-bg); display:flex; flex-direction:column; }
.slide h1 { font-family: var(--ow-font-title); font-size: 44px; color: var(--ow-primary); border-bottom: 3px solid var(--ow-primary); padding-bottom: 12px; margin-bottom: 24px; }
.slide h2 { font-family: var(--ow-font-title); font-size: 30px; color: var(--ow-primary); margin-bottom: 18px; }
.slide p, .slide li { font-size: 22px; line-height: 1.5; }
.slide ul { padding-left: 32px; margin-top: 10px; }
.slide table { width:100%; border-collapse:collapse; font-size:20px; margin-top:16px;}
.slide th{ background:var(--ow-primary); color:#fff; text-align:left; padding:10px 12px;}
.slide td{ padding:9px 12px; border-bottom:.5pt solid #c8d4e0;}
.slide tr:nth-child(even) td{ background:var(--ow-row-alt);}
.kicker{ color:var(--ow-accent); font-weight:bold; letter-spacing:.5px; text-transform:uppercase; font-size:18px; margin-bottom:8px;}
"""


def _hex_to_rgb(h: str):
    from pptx.dml.color import RGBColor
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _add_chart_to_slide(slide, chart_spec: dict[str, Any]):
    """Añade un gráfico nativo de PowerPoint a la diapositiva usando python-pptx."""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    c_type = str(chart_spec.get("type") or chart_spec.get("chart_type") or "bar").lower().strip()
    if c_type in ("pie", "piechart", "pie_chart"):
        xl_type = XL_CHART_TYPE.PIE
    elif c_type in ("line", "linechart", "line_chart"):
        xl_type = XL_CHART_TYPE.LINE
    elif c_type in ("horizontal_bar", "bar_horizontal"):
        xl_type = XL_CHART_TYPE.BAR_CLUSTERED
    else:
        xl_type = XL_CHART_TYPE.COLUMN_CLUSTERED

    cdata = CategoryChartData()
    categories = chart_spec.get("categories") or []
    cdata.categories = [str(c) for c in categories]

    if "series" in chart_spec and isinstance(chart_spec["series"], list):
        for s in chart_spec["series"]:
            name = str(s.get("name", "Serie"))
            vals = tuple(s.get("values", []))
            cdata.add_series(name, vals)
    else:
        vals = tuple(chart_spec.get("values") or [])
        s_name = str(chart_spec.get("series_name") or chart_spec.get("title") or "Valores")
        cdata.add_series(s_name, vals)

    left = Inches(float(chart_spec.get("left", 1.2)))
    top = Inches(float(chart_spec.get("top", 1.8)))
    width = Inches(float(chart_spec.get("width", 10.5)))
    height = Inches(float(chart_spec.get("height", 4.8)))

    chart_shape = slide.shapes.add_chart(xl_type, left, top, width, height, cdata)
    chart = chart_shape.chart
    if chart_spec.get("title"):
        chart.has_title = True
        chart.chart_title.text_frame.text = str(chart_spec["title"])
    return chart_shape


def _create_native_pptx(out_path: str, slides: list[dict], theme: dict) -> str:
    """Generador PPTX nativo usando python-pptx (sin requerir navegador)."""
    from pptx import Presentation
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    primary_color = _hex_to_rgb(theme.get("primary", "#003366"))
    accent_color = _hex_to_rgb(theme.get("accent", "#3B82F6"))

    for s in slides:
        slide = prs.slides.add_slide(blank_layout)

        # Header box (kicker + title)
        title_text = s.get("title", "")
        kicker_text = s.get("kicker", "")

        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(1.3))
        tf = tb.text_frame
        tf.word_wrap = True

        first_p = True
        if kicker_text:
            p_k = tf.paragraphs[0]
            p_k.text = kicker_text.upper()
            p_k.font.size = Pt(14)
            p_k.font.bold = True
            p_k.font.color.rgb = accent_color
            first_p = False

        if title_text:
            p_t = tf.paragraphs[0] if first_p else tf.add_paragraph()
            p_t.text = title_text
            p_t.font.size = Pt(30)
            p_t.font.bold = True
            p_t.font.color.rgb = primary_color

        # Body: bullets o subtitle
        bullets = s.get("bullets", [])
        subtitle = s.get("subtitle", "")
        has_chart = "chart" in s and s["chart"]

        if bullets:
            box_width = Inches(5.0) if has_chart else Inches(11.7)
            b_tb = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), box_width, Inches(4.5))
            b_tf = b_tb.text_frame
            b_tf.word_wrap = True
            for b_idx, bullet in enumerate(bullets):
                bp = b_tf.paragraphs[0] if b_idx == 0 else b_tf.add_paragraph()
                bp.text = f"• {bullet}"
                bp.font.size = Pt(18)
        elif subtitle and not has_chart:
            s_tb = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.7), Inches(2.0))
            s_tf = s_tb.text_frame
            sp = s_tf.paragraphs[0]
            sp.text = subtitle
            sp.font.size = Pt(18)

        if has_chart:
            chart_spec = dict(s["chart"])
            if bullets and "left" not in chart_spec:
                chart_spec["left"] = 6.2
                chart_spec["width"] = 6.3
            _add_chart_to_slide(slide, chart_spec)

    prs.save(out_path)
    return out_path


def create_pptx(out_path, slides=None, theme=None):
    """slides: lista de dicts {"title":..., "kicker":..., "bullets":[...], "chart":{type, categories, values}} o {"html":"..."} crudo.

    Si slides es None → una portada genérica. Devuelve ruta absoluta del .pptx editable.
    """
    from .themes import load_theme, css_vars
    from .security import safe_out
    from pptx import Presentation

    th = load_theme(theme)
    out_path = safe_out(out_path)

    slide_list = slides or [{"title": "The Office Worker", "bullets": ["Presentación generada"]}]

    # Intentar renderizar vía html-to-pptx si está disponible
    try:
        from html_to_pptx import convert
    except ImportError:
        convert = None

    if convert is not None:
        parts = []
        for s in slide_list:
            if "html" in s and s["html"]:
                parts.append(f'<section class="slide">{s["html"]}</section>')
                continue
            k = f'<div class="kicker">{s.get("kicker","")}</div>' if s.get("kicker") else ""
            items = "".join(f"<li>{b}</li>" for b in s.get("bullets", []))
            body = f"<ul>{items}</ul>" if items else f"<p>{s.get('subtitle','')}</p>" if s.get('subtitle') else ""
            parts.append(
                f'<section class="slide"><h1>{s.get("title","")}</h1>{k}{body}</section>'
            )

        deck_html = (f'<!DOCTYPE html><html><head><meta charset="utf-8">'
                     f'<style>{css_vars(th)}{_SLIDE_CSS}</style></head><body>{"".join(parts)}</body></html>')

        with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as tf:
            tf.write(deck_html)
            tmp_html = tf.name

        try:
            import asyncio, concurrent.futures
            with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
                result = ex.submit(asyncio.run, convert(tmp_html, out_path)).result(timeout=300)

            # Inyectar gráficos nativos en diapositivas si se declararon
            has_charts = any(isinstance(s, dict) and "chart" in s and s["chart"] for s in slide_list)
            if has_charts:
                prs = Presentation(result)
                for idx, s in enumerate(slide_list):
                    if isinstance(s, dict) and s.get("chart") and idx < len(prs.slides):
                        _add_chart_to_slide(prs.slides[idx], s["chart"])
                prs.save(result)

        except Exception:
            # Fallback a generación nativa python-pptx si falló el navegador
            result = _create_native_pptx(out_path, slide_list, th)
        finally:
            try: os.unlink(tmp_html)
            except OSError: pass
    else:
        result = _create_native_pptx(out_path, slide_list, th)

    if not os.path.exists(result) or os.path.getsize(result) < 3000:
        raise RuntimeError(f"PPTX generado inválido o vacío ({result})")
    return os.path.abspath(result)
