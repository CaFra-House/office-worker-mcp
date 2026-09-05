import json
import asyncio
import sys
import pathlib

# Asegurar import de office_worker
sys.path.insert(0, str(pathlib.Path(__file__).parent / "src"))
try:
    from office_worker.mcp_server import mcp
except ImportError:
    from src.office_worker.mcp_server import mcp

V020_BASELINE = {
    "render_document": 122.5,
    "create_word": 120.0,
    "create_excel": 95.0,
    "create_pptx": 79.25,
    "read_pdf": 58.75,
    "pdf_extract_tables": 61.25,
    "pdf_list_form_fields": 35.75,
    "list_themes": 47.0,
    "pdf_fill_form": 87.75,
    "pdf_ocr": 77.0,
    "convert_to_pdf": 53.5,
    "pdf_manipulate": 154.25,
}
V020_TOTAL = 992.0

V030_BASELINE = {
    "render_document": 147.75,
    "create_word": 180.0,
    "create_excel": 95.0,
    "create_pptx": 79.25,
    "read_pdf": 98.75,
    "pdf_extract_tables": 61.25,
    "pdf_list_form_fields": 35.75,
    "list_themes": 47.0,
    "pdf_fill_form": 87.75,
    "pdf_ocr": 103.0,
    "convert_to_pdf": 53.5,
    "pdf_manipulate": 179.5,
}
V030_TOTAL = 1168.50

V040_BASELINE = {
    "render_document": 225.0,
    "create_word": 180.0,
    "create_excel": 95.0,
    "create_pptx": 79.25,
    "read_pdf": 98.75,
    "pdf_extract_tables": 61.25,
    "pdf_list_form_fields": 35.75,
    "list_themes": 47.0,
    "pdf_fill_form": 87.75,
    "pdf_ocr": 103.0,
    "convert_to_pdf": 53.5,
    "pdf_manipulate": 179.5,
    "list_templates": 18.0,
    "pdf_compress": 70.25,
    "sign_pdf": 121.25,
}
V040_TOTAL = 1455.25

async def main():
    tools = await mcp.list_tools()
    total = 0
    deprecated_tokens = 0
    current_tokens = {}

    print("=== Métrica REAL de tokens por tool (json.dumps(schema) / 4) ===")
    for t in tools:
        schema = t.inputSchema
        s = json.dumps(schema)
        toks = len(s) / 4
        current_tokens[t.name] = toks
        base_tok = V040_BASELINE.get(t.name)
        if base_tok is not None:
            diff = toks - base_tok
            diff_str = f"({diff:+.2f} vs v0.4.0)"
        else:
            diff_str = "(nueva tool v0.5.0)"
        print(f"  {t.name:20s}: {toks:6.2f} tokens {diff_str}")
        total += toks
        if t.name in ("pdf_extract_tables", "pdf_list_form_fields"):
            deprecated_tokens += toks

    print("-" * 65)
    print(f"Total ({len(tools)} tools registradas v0.5.0) : {total:.2f} tokens")
    print(f"Línea base (v0.4.0 auditada)        : {V040_TOTAL:.2f} tokens")
    delta_v040 = total - V040_TOTAL
    print(f"Delta real vs v0.4.0                : {delta_v040:+.2f} tokens")
    print(f"Línea base (v0.3.0 auditada)        : {V030_TOTAL:.2f} tokens")
    delta_v030 = total - V030_TOTAL
    print(f"Delta real vs v0.3.0                : {delta_v030:+.2f} tokens")
    consolidated_active = total - deprecated_tokens
    print(f"Consolidado (18 tools activas)      : {consolidated_active:.2f} tokens")

if __name__ == "__main__":
    asyncio.run(main())
