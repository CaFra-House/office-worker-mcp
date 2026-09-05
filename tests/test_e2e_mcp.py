"""Test E2E del MCP completo como test de pytest (async).

Levanta el servidor MCP real, hace handshake y ejecuta cada tool creando/leyendo
archivos válidos en disco. PPTX se saltea si Playwright no está disponible.
Requiere: pytest-asyncio (en extras [dev]).
"""
import asyncio, json, os, sys, pathlib
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
            assert len(tools) == 8 and "render_document" in tools, f"tools inesperadas: {tools}"

            async def call(name, args): return json.loads((await s.call_tool(name, args)).content[0].text)

            # 1 PDF con tabla declarativa
            r = await call("render_document", {
                "template_html": "<h1>{{ titulo }}</h1><p class='muted'>{{ subtitulo }}</p>{% if tabla is defined and tabla %}{{ tabla }}{% endif %}",
                "out_path": f"{d}/a.pdf",
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

    # El PDF ya se validó arriba vía read_pdf (n_pages>=1 + texto real extraído).
