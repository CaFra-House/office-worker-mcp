"""Generación de PPTX editable vía html-to-pptx (Playwright Chromium) o nativo python-pptx con soporte de gráficos."""
from __future__ import annotations
import os, tempfile
from typing import Any

_SLIDE_CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: var(--ow-font-body); color: var(--ow-text); background: #f1f5f9; }
.slide {
  width: 1280px;
  height: 720px;
  padding: 48px 64px;
  background: var(--ow-bg);
  display: flex;
  flex-direction: column;
  position: relative;
  overflow: hidden;
}
.slide.cover {
  justify-content: center;
  padding: 72px 88px;
  border-top: 8px solid var(--ow-primary);
}
.slide.cover .cover-accent-bar {
  width: 6px;
  height: 100%;
  position: absolute;
  left: 0;
  top: 0;
  background: var(--ow-accent);
}
.slide.cover .kicker {
  color: var(--ow-accent);
  font-weight: 700;
  letter-spacing: 0.15em;
  text-transform: uppercase;
  font-size: 16px;
  margin-bottom: 14px;
}
.slide.cover h1 {
  font-family: var(--ow-font-title);
  font-size: 48px;
  font-weight: 800;
  color: var(--ow-primary);
  line-height: 1.15;
  border-bottom: none;
  padding-bottom: 0;
  margin-bottom: 16px;
}
.slide.cover .subtitle {
  font-size: 24px;
  color: var(--ow-muted);
  line-height: 1.4;
  margin-top: 6px;
}
.slide.cover ul {
  margin-top: 20px;
  list-style: none;
  padding-left: 0;
}
.slide.cover li {
  font-size: 18px;
  color: var(--ow-text);
  margin-bottom: 8px;
}
.slide.cover li::before {
  content: "▪ ";
  color: var(--ow-accent);
  font-weight: bold;
}
.slide .header-block {
  margin-bottom: 24px;
  border-bottom: 2.5px solid var(--ow-accent);
  padding-bottom: 12px;
}
.slide .header-block .kicker {
  color: var(--ow-accent);
  font-weight: 700;
  letter-spacing: 0.12em;
  text-transform: uppercase;
  font-size: 13px;
  margin-bottom: 4px;
}
.slide .header-block h1 {
  font-family: var(--ow-font-title);
  font-size: 34px;
  font-weight: 700;
  color: var(--ow-primary);
  border-bottom: none;
  padding-bottom: 0;
  margin: 0;
}
.slide h2 {
  font-family: var(--ow-font-title);
  font-size: 26px;
  color: var(--ow-primary);
  margin-bottom: 16px;
}
.slide p {
  font-size: 20px;
  line-height: 1.5;
  color: var(--ow-text);
}
.slide .subtitle {
  font-size: 20px;
  color: var(--ow-muted);
  font-style: italic;
  margin-bottom: 16px;
}
.slide ul {
  padding-left: 0;
  margin-top: 8px;
  list-style: none;
}
.slide li {
  font-size: 19px;
  line-height: 1.6;
  margin-bottom: 12px;
  color: var(--ow-text);
  padding-left: 24px;
  position: relative;
}
.slide li::before {
  content: "▪";
  color: var(--ow-accent);
  font-weight: bold;
  position: absolute;
  left: 0;
  top: 0;
}
.slide table {
  width: 100%;
  border-collapse: separate;
  border-spacing: 0;
  font-size: 16px;
  margin-top: 16px;
  border: 1px solid #cbd5e1;
  border-radius: 6px;
  overflow: hidden;
}
.slide th {
  background: var(--ow-primary);
  color: #ffffff;
  font-family: var(--ow-font-title);
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  font-size: 14px;
  padding: 12px 16px;
  text-align: left;
}
.slide td {
  padding: 10px 16px;
  border-bottom: 1px solid #e2e8f0;
  color: var(--ow-text);
  vertical-align: middle;
}
.slide tr:last-child td {
  border-bottom: none;
}
.slide tr:nth-child(even) td {
  background: var(--ow-row-alt);
}
.slide td:first-child {
  font-weight: 600;
}
"""


def _hex_to_rgb(h: str):
    from pptx.dml.color import RGBColor
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _clean_font_name(font_str: str | None, default: str = "Arial") -> str:
    if not font_str:
        return default
    first = str(font_str).split(",")[0].strip().strip("'").strip('"')
    return first or default


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
    """Generador PPTX nativo de calidad corporativo-agencia usando python-pptx."""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    primary_color = _hex_to_rgb(theme.get("primary", "#003366"))
    accent_color = _hex_to_rgb(theme.get("accent", "#3B82F6"))
    text_color = _hex_to_rgb(theme.get("text", "#1A202C"))
    muted_color = _hex_to_rgb(theme.get("muted", "#718096"))
    row_alt_color = _hex_to_rgb(theme.get("row_alt", "#F5F7FA"))

    font_title = _clean_font_name(theme.get("font_title"), "Helvetica")
    font_body = _clean_font_name(theme.get("font_body"), "Arial")

    for idx, s in enumerate(slides):
        slide = prs.slides.add_slide(blank_layout)

        title_text = s.get("title", "")
        kicker_text = s.get("kicker", "")
        subtitle_text = s.get("subtitle", "")
        bullets = s.get("bullets", [])
        has_chart = "chart" in s and s["chart"]
        has_table = ("table" in s and s["table"]) or ("headers" in s and "rows" in s)

        # Portada: cuando se especifica explícitamente o en slide 0 sin chart ni tabla
        is_cover = (
            s.get("is_cover") is True
            or str(s.get("type", "")).lower() == "cover"
            or (
                idx == 0
                and s.get("is_cover") is not False
                and str(s.get("type", "")).lower() not in ("slide", "content", "internal")
                and not has_chart
                and not has_table
            )
        )

        if is_cover:
            # === PORTADA CORPORATIVO-AGENCIA ===
            # Banda de color primary arriba
            top_stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.14))
            top_stripe.fill.solid()
            top_stripe.fill.fore_color.rgb = primary_color
            top_stripe.line.fill.background()

            # Barra vertical acento
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(2.2), Inches(0.08), Inches(2.6))
            bar.fill.solid()
            bar.fill.fore_color.rgb = accent_color
            bar.line.fill.background()

            # Caja de texto principal con aire generoso
            tb = slide.shapes.add_textbox(Inches(1.5), Inches(1.9), Inches(10.5), Inches(4.8))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0)
            tf.margin_top = Inches(0)
            tf.margin_right = Inches(0)
            tf.margin_bottom = Inches(0)

            first_p = True
            if kicker_text:
                pk = tf.paragraphs[0]
                pk.text = str(kicker_text).upper()
                pk.font.size = Pt(14)
                pk.font.bold = True
                pk.font.color.rgb = accent_color
                pk.font.name = font_title
                pk.space_after = Pt(12)
                first_p = False

            if title_text:
                pt = tf.paragraphs[0] if first_p else tf.add_paragraph()
                pt.text = str(title_text)
                pt.font.size = Pt(38)
                pt.font.bold = True
                pt.font.color.rgb = primary_color
                pt.font.name = font_title
                pt.space_after = Pt(14)
                first_p = False

            if subtitle_text:
                ps = tf.paragraphs[0] if first_p else tf.add_paragraph()
                ps.text = str(subtitle_text)
                ps.font.size = Pt(20)
                ps.font.color.rgb = muted_color
                ps.font.name = font_body
                ps.space_after = Pt(12)
                first_p = False

            if bullets:
                for b_item in bullets:
                    pb = tf.paragraphs[0] if first_p else tf.add_paragraph()
                    r_sym = pb.add_run()
                    r_sym.text = "▪  "
                    r_sym.font.bold = True
                    r_sym.font.size = Pt(14)
                    r_sym.font.color.rgb = accent_color
                    r_txt = pb.add_run()
                    r_txt.text = str(b_item)
                    r_txt.font.size = Pt(16)
                    r_txt.font.color.rgb = text_color
                    r_txt.font.name = font_body
                    pb.space_before = Pt(4)
                    pb.space_after = Pt(6)
                    first_p = False

        else:
            # === SLIDE INTERNO ===
            # Header box (kicker + title)
            tb = slide.shapes.add_textbox(Inches(1.0), Inches(0.55), Inches(11.333), Inches(1.2))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0)
            tf.margin_top = Inches(0)
            tf.margin_right = Inches(0)
            tf.margin_bottom = Inches(0)

            first_p = True
            if kicker_text:
                p_k = tf.paragraphs[0]
                p_k.text = str(kicker_text).upper()
                p_k.font.size = Pt(12)
                p_k.font.bold = True
                p_k.font.color.rgb = accent_color
                p_k.font.name = font_title
                p_k.space_after = Pt(4)
                first_p = False

            if title_text:
                p_t = tf.paragraphs[0] if first_p else tf.add_paragraph()
                p_t.text = str(title_text)
                p_t.font.size = Pt(30)
                p_t.font.bold = True
                p_t.font.color.rgb = primary_color
                p_t.font.name = font_title

            # Línea inferior fina accent debajo del título
            sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(1.82), Inches(11.333), Inches(0.025))
            sep.fill.solid()
            sep.fill.fore_color.rgb = accent_color
            sep.line.fill.background()

            # Contenido: Tablas / Bullets / Subtitle / Chart
            content_top = Inches(2.15)
            content_height = Inches(4.7)

            # 1. Tabla si existe
            if has_table:
                tbl_spec = s.get("table") or {}
                headers = tbl_spec.get("headers") or s.get("headers") or []
                rows = tbl_spec.get("rows") or s.get("rows") or []
                n_rows = len(rows) + (1 if headers else 0)
                n_cols = len(headers) if headers else (len(rows[0]) if rows else 0)
                if n_rows > 0 and n_cols > 0:
                    tbl_shape = slide.shapes.add_table(n_rows, n_cols, Inches(1.0), content_top, Inches(11.333), min(Inches(0.55 * n_rows), content_height))
                    tbl = tbl_shape.table
                    # Header
                    if headers:
                        for c_i, hv in enumerate(headers):
                            c = tbl.cell(0, c_i)
                            c.fill.solid()
                            c.fill.fore_color.rgb = primary_color
                            c.text = str(hv).upper()
                            for p in c.text_frame.paragraphs:
                                for r in p.runs:
                                    r.font.bold = True
                                    r.font.size = Pt(12.5)
                                    r.font.color.rgb = RGBColor(255, 255, 255)
                                    r.font.name = font_title
                    # Filas de datos
                    start_r = 1 if headers else 0
                    for r_i, row_data in enumerate(rows):
                        is_alt = (r_i % 2 == 1)
                        for c_i, cv in enumerate(row_data):
                            if c_i < n_cols:
                                c = tbl.cell(r_i + start_r, c_i)
                                if is_alt:
                                    c.fill.solid()
                                    c.fill.fore_color.rgb = row_alt_color
                                c.text = str(cv)
                                for p in c.text_frame.paragraphs:
                                    for r in p.runs:
                                        r.font.size = Pt(13)
                                        r.font.color.rgb = text_color
                                        r.font.name = font_body
                                        if c_i == 0:
                                            r.font.bold = True

            # 2. Bullets / Subtitle (cuando no hay tabla)
            elif bullets or subtitle_text:
                box_width = Inches(5.0) if has_chart else Inches(11.333)
                b_tb = slide.shapes.add_textbox(Inches(1.0), content_top, box_width, content_height)
                b_tf = b_tb.text_frame
                b_tf.word_wrap = True
                b_tf.margin_left = Inches(0)
                b_tf.margin_top = Inches(0)

                first_bp = True
                if subtitle_text:
                    sp = b_tf.paragraphs[0]
                    sp.text = subtitle_text
                    sp.font.size = Pt(18)
                    sp.font.italic = True
                    sp.font.color.rgb = muted_color
                    sp.font.name = font_body
                    sp.space_after = Pt(12)
                    first_bp = False

                for b_idx, bullet in enumerate(bullets):
                    bp = b_tf.paragraphs[0] if (first_bp and b_idx == 0) else b_tf.add_paragraph()
                    r_sym = bp.add_run()
                    r_sym.text = "▪  "
                    r_sym.font.bold = True
                    r_sym.font.size = Pt(15)
                    r_sym.font.color.rgb = accent_color
                    r_txt = bp.add_run()
                    r_txt.text = str(bullet)
                    r_txt.font.size = Pt(18)
                    r_txt.font.color.rgb = text_color
                    r_txt.font.name = font_body
                    bp.space_before = Pt(6)
                    bp.space_after = Pt(10)
                    bp.line_spacing = 1.25

            # 3. Gráficos
            if has_chart:
                chart_spec = dict(s["chart"])
                if bullets and "left" not in chart_spec:
                    chart_spec["left"] = 6.2
                    chart_spec["width"] = 6.1
                    chart_spec["top"] = 2.15
                    chart_spec["height"] = 4.7
                elif not bullets and "left" not in chart_spec:
                    chart_spec["left"] = 1.0
                    chart_spec["width"] = 11.333
                    chart_spec["top"] = 2.15
                    chart_spec["height"] = 4.7
                _add_chart_to_slide(slide, chart_spec)

    prs.save(out_path)
    return out_path


def create_pptx(out_path, slides=None, theme=None, prefer_native: bool = False):
    """slides: lista de dicts {"title":..., "kicker":..., "bullets":[...], "chart":{type, categories, values}} o {"html":"..."} crudo.

    Si slides es None → una portada genérica. Devuelve ruta absoluta del .pptx editable.
    """
    from .themes import load_theme, css_vars
    from .security import safe_out
    from pptx import Presentation

    th = load_theme(theme)
    out_path = safe_out(out_path)

    slide_list = slides or [{"title": "The Office Worker", "bullets": ["Presentación generada"]}]

    # Si se prefiere nativo directamente, saltar html_to_pptx
    if prefer_native:
        result = _create_native_pptx(out_path, slide_list, th)
        if not os.path.exists(result) or os.path.getsize(result) < 3000:
            raise RuntimeError(f"PPTX generado inválido o vacío ({result})")
        return os.path.abspath(result)

    # Intentar renderizar vía html-to-pptx si está disponible
    try:
        from html_to_pptx import convert
    except ImportError:
        convert = None

    if convert is not None:
        parts = []
        for idx, s in enumerate(slide_list):
            if "html" in s and s["html"]:
                parts.append(f'<section class="slide">{s["html"]}</section>')
                continue

            title_text = s.get("title", "")
            kicker_text = s.get("kicker", "")
            subtitle_text = s.get("subtitle", "")
            bullets = s.get("bullets", [])
            has_chart = "chart" in s and s["chart"]
            has_table = ("table" in s and s["table"]) or ("headers" in s and "rows" in s)

            is_cover = (
                s.get("is_cover") is True
                or str(s.get("type", "")).lower() == "cover"
                or (
                    idx == 0
                    and s.get("is_cover") is not False
                    and str(s.get("type", "")).lower() not in ("slide", "content", "internal")
                    and not has_chart
                    and not has_table
                )
            )

            if is_cover:
                k = f'<div class="kicker">{kicker_text}</div>' if kicker_text else ""
                t = f'<h1>{title_text}</h1>' if title_text else ""
                sub = f'<div class="subtitle">{subtitle_text}</div>' if subtitle_text else ""
                items = "".join(f"<li>{b}</li>" for b in bullets)
                b_html = f"<ul>{items}</ul>" if items else ""
                parts.append(
                    f'<section class="slide cover">'
                    f'<div class="cover-accent-bar"></div>'
                    f'{k}{t}{sub}{b_html}'
                    f'</section>'
                )
            else:
                k = f'<div class="kicker">{kicker_text}</div>' if kicker_text else ""
                t = f'<h1>{title_text}</h1>' if title_text else ""
                hdr = f'<div class="header-block">{k}{t}</div>'

                tbl_html = ""
                if has_table:
                    tbl_spec = s.get("table") or {}
                    headers = tbl_spec.get("headers") or s.get("headers") or []
                    rows = tbl_spec.get("rows") or s.get("rows") or []
                    th_str = "".join(f"<th>{h}</th>" for h in headers)
                    tr_str = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in rows)
                    tbl_html = f"<table><thead><tr>{th_str}</tr></thead><tbody>{tr_str}</tbody></table>"

                sub = f'<div class="subtitle">{subtitle_text}</div>' if subtitle_text else ""
                items = "".join(f"<li>{b}</li>" for b in bullets)
                b_html = f'<ul style="{"width:50%;" if has_chart else ""}">{items}</ul>' if items else ""

                parts.append(
                    f'<section class="slide">'
                    f'{hdr}{sub}{tbl_html}{b_html}'
                    f'</section>'
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
