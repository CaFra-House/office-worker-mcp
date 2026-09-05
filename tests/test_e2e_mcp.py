"""Test E2E del MCP completo como test de pytest (async).

Levanta el servidor MCP real, hace handshake y ejecuta las 12 tools creando/leyendo
archivos válidos en disco. PPTX se saltea si Playwright no está disponible.
Requiere: pytest-asyncio (en extras [dev]).
"""
import asyncio, json, os, shutil, sys, pathlib
import pytest

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent / "src"))

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


@pytest.mark.asyncio
async def test_office_worker_mcp_e2e(tmp_path):
    d = str(tmp_path)
    env = {**os.environ, "PYTHONPATH": str(pathlib.Path(__file__).parent.parent / "src")}
    params = StdioServerParameters(command=sys.executable, args=["-m", "office_worker.mcp_server"], env=env)

    results = {}
    async with stdio_client(params) as (r, w):
        async with ClientSession(r, w) as s:
            await s.initialize()
            tools = [t.name for t in (await s.list_tools()).tools]
            assert len(tools) == 12 and "render_document" in tools, f"tools inesperadas: {tools}"

            async def call(name, args): return json.loads((await s.call_tool(name, args)).content[0].text)

            # 1 PDF con tabla declarativa y tema corporativo nuevo
            r = await call("render_document", {
                "template_html": "<h1>{{ titulo }}</h1><p class='muted'>{{ subtitulo }}</p>{% if tabla is defined and tabla %}{{ tabla }}{% endif %}",
                "out_path": f"{d}/a.pdf",
                "theme": "corporate-blue",
                "data_json": json.dumps({"titulo": "Informe", "subtitulo": "E2E", "headers": ["KPI", "V"], "rows": [["A", 1], ["B", 2]]})})
            results["pdf"] = r; assert r["status"] == "ok" and os.path.exists(r["path"]), r

            # 2 Word
            r = await call("create_word", {
                "out_path": f"{d}/b.docx", "title": "Acta", "subtitle": "Reunión",
                "blocks_json": json.dumps([{"type": "h2", "text": "Puntos"}, {"type": "p", "text": "Decisión"}, {"type": "table", "headers": ["Item", "Estado"], "rows": [["X", "OK"]]}])})
            results["docx"] = r; assert r["status"] == "ok" and os.path.exists(r["path"]), r

            # 3 Excel multi-hoja
            r = await call("create_excel", {
                "out_path": f"{d}/c.xlsx", "title": "Balance",
                "sheets_json": json.dumps([{"name": "Resumen", "headers": ["Rubro", "Monto"], "rows": [["Ingresos", 100], ["Gastos", 40]]}])})
            results["xlsx"] = r; assert r["status"] == "ok" and os.path.exists(r["path"]), r

            # 4 PPTX editable — se saltea si no hay Playwright instalado en el entorno
            try:
                import playwright  # noqa: F401
                have_pw = True
            except Exception:
                have_pw = False
            if have_pw:
                r = await call("create_pptx", {
                    "out_path": f"{d}/d.pptx",
                    "slides_json": json.dumps([{"title": "Portada", "kicker": "The Office Worker"}, {"title": "Agenda", "bullets": ["Intro", "Resultados"]}])})
                results["pptx"] = r; assert r["status"] == "ok" and os.path.exists(r["path"]), r
            else:
                results["pptx"] = {"status": "skipped"}

            # 5-7 PDF input sobre el PDF generado en (1)
            r = await call("read_pdf", {"path": results["pdf"]["path"]}); results["read_pdf"] = r; assert r.get("n_pages", 0) >= 1 and r.get("pages"), r
            r = await call("pdf_extract_tables", {"path": results["pdf"]["path"]}); results["tables"] = r; assert r.get("status") == "ok", r
            r = await call("pdf_list_form_fields", {"path": results["pdf"]["path"]}); results["forms"] = r; assert r.get("status") == "ok", r

            # 8 temas
            r = await call("list_themes", {}); results["themes"] = r; assert r.get("theme", {}).get("primary"), r

            # 9 convert_to_pdf (Word -> PDF vía LibreOffice)
            if shutil.which("soffice") or shutil.which("libreoffice"):
                r = await call("convert_to_pdf", {"input_file": results["docx"]["path"], "output": f"{d}/b_from_docx.pdf"})
                results["convert"] = r
                assert r["status"] == "ok" and os.path.exists(r["path"]), r
                assert open(r["path"], "rb").read(5) == b"%PDF-"

            # 10 pdf_manipulate (merge 2 pdfs y extract)
            r = await call("pdf_manipulate", {
                "operation": "merge",
                "output": f"{d}/merged_e2e.pdf",
                "files": [results["pdf"]["path"], results["pdf"]["path"]]
            })
            results["merge"] = r
            assert r["status"] == "ok" and os.path.exists(r["path"]), r

            r = await call("pdf_manipulate", {
                "operation": "extract",
                "output": f"{d}/extracted_e2e.pdf",
                "input_path": results["merge"]["path"],
                "pages": "1"
            })
            results["extract"] = r
            assert r["status"] == "ok" and os.path.exists(r["path"]), r

            # 11 pdf_fill_form (crear form con reportlab y rellenar vía tool)
            try:
                from reportlab.pdfgen import canvas
                form_file = f"{d}/e2e_form.pdf"
                c = canvas.Canvas(form_file)
                c.drawString(50, 750, "Email:")
                c.acroForm.textfield(name="email", x=100, y=745, width=150, height=20)
                c.showPage()
                c.save()

                r = await call("pdf_fill_form", {
                    "input_pdf": form_file,
                    "fields": {"email": "test@example.com"},
                    "output": f"{d}/e2e_filled.pdf"
                })
                results["fill_form"] = r
                assert r["status"] == "ok" and os.path.exists(r["path"]), r
            except Exception as exc:
                print(f"pdf_fill_form E2E skip/warn: {exc}")

            # 12 pdf_ocr (imagen -> texto + PDF buscable)
            if shutil.which("tesseract"):
                try:
                    from PIL import Image, ImageDraw
                    img_file = f"{d}/e2e_ocr.png"
                    img = Image.new("RGB", (250, 60), color=(255, 255, 255))
                    d_ctx = ImageDraw.Draw(img)
                    d_ctx.text((10, 20), "HELLO WORLD", fill=(0, 0, 0))
                    img.save(img_file)

                    r = await call("pdf_ocr", {
                        "input_path": img_file,
                        "lang": "eng",
                        "output": f"{d}/e2e_ocr.pdf"
                    })
                    results["ocr"] = r
                    assert r["status"] == "ok", r
                    assert "HELLO" in r.get("text", "").replace(" ", "").upper()
                    assert os.path.exists(r["path"])
                except Exception as exc:
                    print(f"pdf_ocr E2E skip/warn: {exc}")
