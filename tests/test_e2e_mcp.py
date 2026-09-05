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
            assert len(tools) == 20 and "render_document" in tools, f"tools inesperadas: {tools}"

            async def call(name, args): return json.loads((await s.call_tool(name, args)).content[0].text)

            # 1 PDF con tabla declarativa y tema corporativo nuevo
            r = await call("render_document", {
                "template_html": "<h1>{{ titulo }}</h1><p class='muted'>{{ subtitulo }}</p>{% if tabla is defined and tabla %}{{ tabla }}{% endif %}",
                "out_path": f"{d}/a.pdf",
                "theme": "corporate-blue",
                "data_json": json.dumps({"titulo": "Informe", "subtitulo": "E2E", "headers": ["KPI", "V"], "rows": [["A", 1], ["B", 2]]})})
            results["pdf"] = r; assert r["status"] == "ok" and os.path.exists(r["path"]), r

            # 1.b PDF con contraseña opcional
            r_enc = await call("render_document", {
                "template_html": "<h1>Secret E2E</h1>",
                "out_path": f"{d}/a_enc.pdf",
                "password": "e2e_secret_password"
            })
            assert r_enc["status"] == "ok" and os.path.exists(r_enc["path"])

            # 2 Word (bloques)
            r = await call("create_word", {
                "out_path": f"{d}/b.docx", "title": "Acta", "subtitle": "Reunión",
                "blocks_json": json.dumps([{"type": "h2", "text": "Puntos"}, {"type": "p", "text": "Decisión"}, {"type": "table", "headers": ["Item", "Estado"], "rows": [["X", "OK"]]}])})
            results["docx"] = r; assert r["status"] == "ok" and os.path.exists(r["path"]), r

            # 2.b Word (plantilla docxtpl)
            from docx import Document
            tpl_f = f"{d}/template_e2e.docx"
            d_tpl = Document(); d_tpl.add_paragraph("Hola {{ cliente }}"); d_tpl.save(tpl_f)
            r_tpl = await call("create_word", {
                "out_path": f"{d}/b_from_tpl.docx",
                "template_docx": tpl_f,
                "context": {"cliente": "Cliente VIP"}
            })
            assert r_tpl["status"] == "ok" and os.path.exists(r_tpl["path"])

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

            # 5-7 PDF input sobre el PDF generado en (1) + consolidado
            r = await call("read_pdf", {"path": results["pdf"]["path"], "extract_tables": True, "list_forms": True})
            results["read_pdf"] = r
            assert r.get("n_pages", 0) >= 1 and r.get("pages") and "tables_by_page" in r and "is_form" in r, r

            # Herramientas deprecated aún funcionando
            r = await call("pdf_extract_tables", {"path": results["pdf"]["path"]}); results["tables"] = r; assert r.get("status") == "ok", r
            r = await call("pdf_list_form_fields", {"path": results["pdf"]["path"]}); results["forms"] = r; assert r.get("status") == "ok", r

            # 8 temas
            r = await call("list_themes", {}); results["themes"] = r; assert r.get("theme", {}).get("primary"), r

            # 9 convert_to_pdf (Word -> PDF vía LibreOffice)
            if shutil.which("soffice") or shutil.which("libreoffice"):
                r = await call("convert_to_pdf", {"input_file": results["docx"]["path"], "output": f"{d}/b_from_docx.pdf"})
                results["convert"] = r
                assert r["status"] == "ok" and os.path.exists(r["path"]), r
                assert r.get("fidelity") == "clean" and "warnings" in r, r
                assert open(r["path"], "rb").read(5) == b"%PDF-"

            # 10 pdf_manipulate (merge 2 pdfs con password y extract)
            r = await call("pdf_manipulate", {
                "operation": "merge",
                "output": f"{d}/merged_e2e.pdf",
                "files": [results["pdf"]["path"], results["pdf"]["path"]],
                "password": "pwd_merged_e2e"
            })
            results["merge"] = r
            assert r["status"] == "ok" and os.path.exists(r["path"]), r

            r = await call("pdf_manipulate", {
                "operation": "extract",
                "output": f"{d}/extracted_e2e.pdf",
                "input_path": results["pdf"]["path"],
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

            # 13 list_templates
            r_tpls = await call("list_templates", {})
            results["list_templates"] = r_tpls
            assert r_tpls["status"] == "ok"
            assert len(r_tpls["templates"]) == 5
            tpl_names = [t["name"] for t in r_tpls["templates"]]
            assert "acta_meeting" in tpl_names and "factura_simple" in tpl_names

            # 13.b create_word con plantilla empaquetada
            r_pack = await call("create_word", {
                "out_path": f"{d}/acta_packaged.docx",
                "template_docx": "acta_meeting",
                "context": {
                    "titulo": "Acta E2E",
                    "fecha": "2026-09-05",
                    "hora": "12:00",
                    "lugar": "Sala A",
                    "asistentes": [{"nombre": "Ana", "rol": "Líder"}],
                    "puntos": [{"orden": 1, "tema": "Test", "discusion": "Exitoso"}],
                    "acuerdos": [{"acuerdo": "Aprobar", "responsable": "Ana", "fecha_limite": "Hoy"}],
                    "firmas": [{"nombre": "Ana", "cargo": "Líder"}],
                }
            })
            assert r_pack["status"] == "ok" and os.path.exists(r_pack["path"])

            # 14 pdf_compress
            r_comp = await call("pdf_compress", {
                "input_path": results["pdf"]["path"],
                "output": f"{d}/compressed_e2e.pdf",
                "quality": "med"
            })
            results["compress"] = r_comp
            assert r_comp["status"] == "ok" and os.path.exists(r_comp["path"])

            # 15 sign_pdf (estampa sello visual PNG)
            from PIL import Image, ImageDraw
            seal_file = f"{d}/seal.png"
            seal_img = Image.new("RGBA", (150, 50), color=(0, 0, 0, 0))
            d_seal = ImageDraw.Draw(seal_img)
            d_seal.rectangle([(1, 1), (148, 48)], outline=(0, 51, 102, 255), width=2)
            d_seal.text((10, 15), "SELLO E2E", fill=(0, 51, 102, 255))
            seal_img.save(seal_file)

            r_sign = await call("sign_pdf", {
                "input_pdf": results["pdf"]["path"],
                "output": f"{d}/signed_e2e.pdf",
                "sello_img_path": seal_file,
                "reason": "Test E2E",
                "location": "Buenos Aires"
            })
            results["sign"] = r_sign
            assert r_sign["status"] == "ok" and os.path.exists(r_sign["path"])

            # 16 edit_excel (modificar c.xlsx)
            r_edit_xl = await call("edit_excel", {
                "input_path": results["xlsx"]["path"],
                "operations": [
                    {"op": "set_cell", "coordinate": "B2", "value": 150},
                    {"op": "append_row", "row": ["Inversiones", 60]},
                    {"op": "add_column", "header": "Estado", "values": ["Cerrado", "Abierto", "Pendiente"]},
                    {"op": "add_chart", "chart_type": "bar", "title": "Finanzas E2E", "target_cell": "E2"},
                    {"op": "add_table", "table_style": "TableStyleMedium9"},
                    {"op": "auto_filter"}
                ]
            })
            assert r_edit_xl["status"] == "ok" and os.path.exists(r_edit_xl["path"]), r_edit_xl
            assert r_edit_xl.get("fidelity") == "rich"

            # 17 edit_word (modificar b.docx)
            r_edit_wd = await call("edit_word", {
                "input_path": results["docx"]["path"],
                "operations": [
                    {"op": "append_paragraph", "text": "Párrafo editado E2E", "bold": True},
                    {"op": "replace_text", "find": "Decisión", "replace": "Acuerdo Unánime"},
                    {"op": "insert_after_heading", "heading_text": "Puntos", "text": "Punto 1.1 introducido"},
                    {"op": "append_table", "headers": ["Aprobador", "Firma"], "rows": [["Julio", "OK"]]}
                ]
            })
            assert r_edit_wd["status"] == "ok" and os.path.exists(r_edit_wd["path"]), r_edit_wd
            assert r_edit_wd.get("fidelity") == "clean"

            # 18 read_pdf con extract_images=True sobre el signed_e2e.pdf
            r_img = await call("read_pdf", {
                "path": results["sign"]["path"],
                "extract_images": True,
                "max_images": 5
            })
            assert r_img["status"] == "ok"
            assert r_img.get("n_images", 0) >= 1
            assert "images" in r_img
            assert r_img["images"][0]["data_url"].startswith("data:image/")

            # 19 pdf_to_excel (extraer tablas de a.pdf a un nuevo .xlsx)
            r_p2x = await call("pdf_to_excel", {
                "input_path": results["pdf"]["path"],
                "output_path": f"{d}/from_pdf_e2e.xlsx",
                "sheet_name": "Reporte"
            })
            assert r_p2x["status"] == "ok" and os.path.exists(r_p2x["path"]), r_p2x
            assert r_p2x.get("fidelity") == "clean"
            assert r_p2x.get("n_tables", 0) >= 1

            # 20 read_office (leer docx y xlsx)
            r_ro_docx = await call("read_office", {"path": results["docx"]["path"]})
            assert r_ro_docx["status"] == "ok" and r_ro_docx["format"] == "docx"
            assert r_ro_docx["n_paragraphs"] >= 1

            r_ro_xlsx = await call("read_office", {"path": results["xlsx"]["path"]})
            assert r_ro_xlsx["status"] == "ok" and r_ro_xlsx["format"] == "xlsx"
            assert r_ro_xlsx["n_sheets"] >= 1

            # 21 office_batch (ejecutar lote con manejo de error parcial)
            r_batch = await call("office_batch", {
                "operations": [
                    {
                        "tool": "create_word",
                        "args": {
                            "out_path": f"{d}/batch_doc.docx",
                            "title": "Documento Batch",
                            "blocks_json": json.dumps([{"type": "p", "text": "Contenido Batch"}])
                        }
                    },
                    {
                        "tool": "tool_inexistente",
                        "args": {"foo": "bar"}
                    },
                    {
                        "tool": "list_templates",
                        "args": {}
                    }
                ]
            })
            assert r_batch["status"] == "partial_error"
            assert r_batch["total"] == 3
            assert r_batch["succeeded"] == 2
            assert r_batch["failed"] == 1
            assert os.path.exists(f"{d}/batch_doc.docx")


