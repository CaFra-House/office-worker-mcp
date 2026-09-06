"""Cara CLI mínima ('owi') para scripts/no-tecnicos. El MCP es la cara principal; esto es extra.

Uso:
  owi pdf --out informe.pdf --title "Informe" --subtitle "Q3" [--theme aden] [--rows "A|1,B|2"]
  owi word --out acta.docx --title "Acta" [--blocks-json '[{"type":"p","text":"hola"}]']
  owi excel --out bal.xlsx --title "Balance" [--sheets-json '[{"name":"R","headers":[...],"rows":[...]}]']
"""
from __future__ import annotations
import argparse, json, os, sys


def _print(r): print(json.dumps(r, ensure_ascii=False))


def main(argv=None):
    from office_worker.core.setup_notice import print_setup_notice_if_needed
    print_setup_notice_if_needed()  # stderr only; stdout stays clean for JSON output
    p = argparse.ArgumentParser(prog="owi", description="The Office Worker — CLI")
    sub = p.add_subparsers(dest="cmd", required=True)

    def common(sp):
        sp.add_argument("--out", required=True); sp.add_argument("--theme", default=None)

    pdf = sub.add_parser("pdf"); common(pdf); pdf.add_argument("--title", default="Documento")
    pdf.add_argument("--subtitle", default=""); pdf.add_argument("--body", default="")
    pdf.add_argument("--rows", default="", help="tabla: 'H1|H2' filas separadas por ',' celdas por '|'")
    pdf.add_argument("--design-mode", default="standard", choices=["standard", "premium"])

    wd = sub.add_parser("word"); common(wd); wd.add_argument("--title", default="")
    wd.add_argument("--blocks-json", default="[]")

    ex = sub.add_parser("excel"); common(ex); ex.add_argument("--title", default="")
    ex.add_argument("--sheets-json", default="[]")

    bk = sub.add_parser("book", help="Genera libro multi-capítulo profesional (PDF/EPUB)")
    common(bk); bk.add_argument("--title", required=True)
    bk.add_argument("--author", default="")
    bk.add_argument("--chapters-json", default="[]")
    bk.add_argument("--epub", action="store_true", default=False)

    doc = sub.add_parser("doctor", help="Audita capacidades del entorno y binarios del sistema")

    sk = sub.add_parser("skill", help="Gestión de skills empaquetadas para agentes")
    sk_sub = sk.add_subparsers(dest="skill_cmd", required=True)
    sk_inst = sk_sub.add_parser("install", help="Instala skill en el directorio de skills de Hermes (auto-resuelto por plataforma)")
    sk_inst.add_argument("name", nargs="?", default="office-worker", help="Nombre de la skill ('office-worker', 'google-drive-gmail', o 'all')")
    sk_inst.add_argument("--dest", default=None, help="Directorio destino opcional")
    sk_sub.add_parser("list", help="Lista skills empaquetadas disponibles")

    a = p.parse_args(argv)

    from office_worker.core import render_pdf, create_word, create_excel, create_book, check_environment

    if a.cmd == "pdf":
        rows=[]; headers=[]
        if a.rows:
            parts=a.rows.split(",")
            headers=parts[0].split("|"); rows=[r.split("|") for r in parts[1:]]
        data={"titulo":a.title,"subtitulo":a.subtitle,"resumen":a.body}
        if rows: data.update({"headers":headers,"rows":rows})
        tpl="<h1>{{ titulo }}</h1><p class='muted'>{{ subtitulo }}</p>{% if resumen %}<p>{{ resumen }}</p>{% endif %}{% if tabla is defined and tabla %}{{ tabla }}{% endif %}"
        try: _print(render_pdf(tpl,a.out,data=data,theme=a.theme,design_mode=a.design_mode)); return 0
        except Exception as e: print(json.dumps({"status":"error","error":str(e)})); return 1
    elif a.cmd == "word":
        try: blocks=json.loads(a.blocks_json or "[]"); _print(create_word(a.out,title=a.title,blocks=blocks,theme=a.theme)); return 0
        except Exception as e: print(json.dumps({"status":"error","error":str(e)})); return 1
    elif a.cmd == "excel":
        try: sheets=json.loads(a.sheets_json or "[]"); _print(create_excel(a.out,title=a.title,sheets=sheets or None,theme=a.theme)); return 0
        except Exception as e: print(json.dumps({"status":"error","error":str(e)})); return 1
    elif a.cmd == "book":
        try:
            chapters = json.loads(a.chapters_json or "[]")
            res = create_book(a.out, title=a.title, author=a.author, chapters=chapters, theme=a.theme, epub=a.epub)
            _print(res)
            return 0
        except Exception as e:
            print(json.dumps({"status": "error", "error": str(e)}))
            return 1
    elif a.cmd == "doctor":
        try:
            res = check_environment()
            _print(res)
            return 0
        except Exception as e:
            print(json.dumps({"status": "error", "error": str(e)}))
            return 1
    elif a.cmd == "skill":
        from office_worker.skills import install_skill, list_packaged_skills
        if a.skill_cmd == "install":
            try:
                res = install_skill(a.name, dest_dir=a.dest)
                _print(res)
                return 0
            except Exception as e:
                print(json.dumps({"status": "error", "error": str(e)}))
                return 1
        elif a.skill_cmd == "list":
            try:
                res = {"status": "ok", "skills": list_packaged_skills()}
                _print(res)
                return 0
            except Exception as e:
                print(json.dumps({"status": "error", "error": str(e)}))
                return 1
    return 0


if __name__ == "__main__": sys.exit(main())
