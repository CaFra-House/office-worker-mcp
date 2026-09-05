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

V050_BASELINE = {
    "render_document": 225.00,
    "create_word": 180.00,
    "create_excel": 141.25,
    "create_pptx": 79.25,
    "read_pdf": 138.00,
    "pdf_extract_tables": 61.25,
    "pdf_list_form_fields": 35.75,
    "list_themes": 47.00,
    "pdf_fill_form": 87.75,
    "pdf_ocr": 103.00,
    "convert_to_pdf": 53.50,
    "pdf_manipulate": 179.50,
    "list_templates": 18.00,
    "pdf_compress": 70.25,
    "sign_pdf": 121.25,
    "edit_excel": 105.75,
    "edit_word": 105.50,
    "pdf_to_excel": 102.25,
    "read_office": 51.00,
    "office_batch": 58.25,
}
V050_TOTAL = 1963.50

V051_BASELINE = {
    "render_document": 200.25,
    "create_word": 147.00,
    "create_excel": 124.75,
    "create_pptx": 71.00,
    "read_pdf": 129.50,
    "pdf_preview": 86.50,
    "list_themes": 38.75,
    "pdf_fill_form": 76.00,
    "pdf_ocr": 86.25,
    "convert_to_pdf": 53.50,
    "pdf_manipulate": 137.50,
    "list_templates": 18.00,
    "pdf_compress": 70.25,
    "sign_pdf": 121.25,
    "edit_excel": 77.25,
    "edit_word": 77.00,
    "pdf_to_excel": 93.75,
    "read_office": 51.00,
    "office_batch": 38.00,
}
V051_TOTAL = 1697.50

async def main():
    tools = await mcp.list_tools()
    total = 0
    current_tokens = {}

    print("=== Métrica REAL de tokens por tool (json.dumps(schema) / 4) ===")
    for t in tools:
        schema = t.inputSchema
        s = json.dumps(schema)
        toks = len(s) / 4
        current_tokens[t.name] = toks
        base_tok = V051_BASELINE.get(t.name)
        if base_tok is not None:
            diff = toks - base_tok
            diff_str = f"({diff:+.2f} vs v0.5.1)"
        else:
            diff_str = "(nueva tool v0.6.0)"
        print(f"  {t.name:24s}: {toks:6.2f} tokens {diff_str}")
        total += toks

    print("-" * 65)
    print(f"Total ({len(tools)} tools registradas v0.6.0) : {total:.2f} tokens")
    print(f"Línea base (v0.5.1 auditada)        : {V051_TOTAL:.2f} tokens")
    delta_v051 = total - V051_TOTAL
    print(f"Delta real vs v0.5.1                : {delta_v051:+.2f} tokens")
    print(f"Línea base (v0.5.0 auditada)        : {V050_TOTAL:.2f} tokens")
    delta_v050 = total - V050_TOTAL
    print(f"Delta real vs v0.5.0                : {delta_v050:+.2f} tokens")
    print(f"Objetivo <1900 tok v0.6.0           : {'CUMPLIDO (<1900)' if total < 1900 else 'NO CUMPLIDO'}")

if __name__ == "__main__":
    asyncio.run(main())
