"""
Merge de tres vias para planilhas de controle (.xlsm / .xlsx).

Uso via linha de comando:
  python merge_planilhas.py BASE.xlsm MINHA.xlsm DELA.xlsm -o UNIFICADA.xlsm
                            [--prefere minha|dela] [--relatorio conflitos.txt]

  BASE  = versao original, antes das edicoes (ancestral comum)
  MINHA = primeira versao editada
  DELA  = segunda versao editada

Regras:
  * Linhas da BASE TROCA sao identificadas por (Nº PEDIDO, TIPO, ordem dentro do pedido).
  * Celula alterada so em um lado -> usa a alteracao.
  * Mesma celula alterada nos dois lados de forma diferente -> CONFLITO.
  * Formulas sao recalculadas ao abrir no Excel (fullCalcOnLoad).
  * O codigo VBA, botoes e desenhos sao preservados nativamente em arquivos .xlsm.
"""
import argparse
import html
import os
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
from copy import copy

import openpyxl
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

SHEET_MAIN = "BASE TROCA"
TABLE_NAME = "tblBaseTroca"
HDR_ROW = 5
FIRST = 6

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def headers(ws):
    out = {}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(HDR_ROW, c).value
        if h not in (None, ""):
            out[str(h).strip()] = c
    return out


def find_col(hdrs, name, contains=False):
    for h, c in hdrs.items():
        if (contains and name in h) or (not contains and h == name):
            if contains and "TOTAL" in h and name == "PEDIDO":
                continue
            return c
    raise SystemExit(f"Coluna '{name}' nao encontrada na linha {HDR_ROW}.")


def last_row(ws, c_id):
    last = FIRST - 1
    for r in range(FIRST, ws.max_row + 1):
        if ws.cell(r, c_id).value not in (None, ""):
            last = r
    return last


def read_main(ws):
    """Devolve (ordem_das_chaves, dados_por_chave, colunas_de_entrada, headers)."""
    hd = headers(ws)
    c_id = find_col(hd, "ID")
    c_tipo = find_col(hd, "TIPO")
    c_ped = find_col(hd, "PEDIDO", contains=True)
    cols = [c for c in range(1, ws.max_column + 1)
            if ws.cell(HDR_ROW, c).value not in (None, "")]
    inputs = []
    for c in cols:
        if c in (c_ped,):
            inputs.append(c)
            continue
        sample = [ws.cell(r, c).value for r in range(FIRST, last_row(ws, c_id) + 1)]
        if not any(is_formula(v) for v in sample):
            inputs.append(c)
    occ, order, data = {}, [], {}
    for r in range(FIRST, last_row(ws, c_id) + 1):
        if ws.cell(r, c_id).value in (None, ""):
            continue
        ped = ws.cell(r, c_ped).value
        tipo = str(ws.cell(r, c_tipo).value or "").strip().upper()
        base_key = (ped, tipo)
        occ[base_key] = occ.get(base_key, 0) + 1
        key = (ped, tipo, occ[base_key])
        order.append(key)
        data[key] = {L(c): ws.cell(r, c).value for c in inputs}
    return order, data, inputs, hd


def three_way(base, a, b, prefer, log, where, name_a="Arquivo 1", name_b="Arquivo 2", conflicts_list=None):
    """Merge de um valor. Devolve o valor escolhido e registra conflito se houver."""
    if a == b:
        return a
    if a == base:
        return b
    if b == base:
        return a
    msg = f"CONFLITO {where}: base={base!r} | {name_a}={a!r} | {name_b}={b!r}"
    log.append(msg)
    if conflicts_list is not None:
        conflicts_list.append({
            "where": where,
            "base": base,
            "val_a": a,
            "val_b": b,
            "name_a": name_a,
            "name_b": name_b,
            "msg": msg
        })
    return a if prefer in ("minha", "1", name_a) else b


def merge_main(wa, wb, wbase, prefer, log, name_a="Arquivo 1", name_b="Arquivo 2", conflicts_list=None):
    ob, db, cb_in, hb = read_main(wbase[SHEET_MAIN])
    oa, da, ca_in, ha = read_main(wa[SHEET_MAIN])
    ob2, dbb, cbb_in, hbb = read_main(wb[SHEET_MAIN])
    if not (set(ha) == set(hb) == set(hbb)):
        diff = sorted(set(ha) ^ set(hb) ^ set(hbb))
        raise SystemExit("As colunas da BASE TROCA sao diferentes entre as versoes: "
                         f"{diff}. Unifique a estrutura antes de fazer o merge.")

    keys = list(ob)
    for k in oa + list(ob2):
        if k in keys:
            continue
        # Linha nova: entra logo depois do ultimo item do mesmo pedido
        pos = [i for i, x in enumerate(keys) if x[0] == k[0]]
        if pos:
            keys.insert(pos[-1] + 1, k)
        else:
            keys.append(k)

    result = []
    merged = {}
    for k in keys:
        base = db.get(k)
        a = da.get(k)
        b = dbb.get(k)
        tag = f"pedido {k[0]} {k[1]}#{k[2]}"
        if base is None and a is None and b is None:
            continue
        if a is None and b is None:
            continue
        if base is None:  # Linha nova em um ou ambos os lados
            if a is not None and b is not None:
                row = {}
                for f in a:
                    row[f] = three_way(None, a[f], b[f], prefer, log, f"{tag} col {f}", name_a, name_b, conflicts_list)
                merged[k] = row
            else:
                merged[k] = dict(a or b)
            result.append(k)
            continue
        if a is None or b is None:  # Linha apagada de um lado
            other = b if a is None else a
            side = name_b if a is None else name_a
            if other == base:
                continue  # Apagada de um lado e intacta no outro: confirma exclusao
            msg = f"CONFLITO {tag}: apagada em {side} mas alterada no outro"
            log.append(msg)
            if conflicts_list is not None:
                conflicts_list.append({
                    "where": tag,
                    "base": "Linha original existente",
                    "val_a": "Apagada" if a is None else "Mantida/Alterada",
                    "val_b": "Apagada" if b is None else "Mantida/Alterada",
                    "name_a": name_a,
                    "name_b": name_b,
                    "msg": msg
                })
            if (prefer in ("minha", "1", name_a)) == (side == name_b):
                merged[k] = dict(other)
                result.append(k)
            continue
        row = {}
        for f in base:
            row[f] = three_way(base[f], a[f], b[f], prefer, log, f"{tag} col {f}", name_a, name_b, conflicts_list)
        merged[k] = row
        result.append(k)
    return result, merged, ha


def merge_cells(wa, wb, wbase, sheet, prefer, log, name_a="Arquivo 1", name_b="Arquivo 2", conflicts_list=None):
    ws0 = wbase[sheet]
    wsa = wa[sheet]
    wsb = wb[sheet]
    max_r = max(ws0.max_row, wsa.max_row, wsb.max_row)
    max_c = max(ws0.max_column, wsa.max_column, wsb.max_column)
    out = {}
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            base = ws0.cell(r, c).value
            a = wsa.cell(r, c).value
            b = wsb.cell(r, c).value
            if base == a == b:
                continue
            out[(r, c)] = three_way(base, a, b, prefer, log, f"{sheet}!{L(c)}{r}", name_a, name_b, conflicts_list)
    return out


def write_main(wout, result, merged, hd, inputs_cols=None):
    ws = wout[SHEET_MAIN]
    template = {c: copy(ws.cell(FIRST, c)._style) for c in range(1, ws.max_column + 1)}
    f_row = {c: ws.cell(FIRST, c).value for c in range(1, ws.max_column + 1)}
    
    # Limpa linhas antigas
    for r in range(FIRST, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    # Escreve novas linhas
    for i, k in enumerate(result):
        r = FIRST + i
        row = merged[k]
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell._style = copy(template[c])
            col_letter = L(c)
            if col_letter in row:
                cell.value = row[col_letter]
            elif is_formula(f_row.get(c)):
                cell.value = re.sub(r"(?<=[A-Z])(?<!\$)%d(?!\d)" % FIRST, str(r), f_row[c])
                
    new_last = FIRST + len(result) - 1
    tbl = ws.tables.get(TABLE_NAME)
    if tbl is not None:
        tbl.ref = f"A{HDR_ROW}:{L(ws.max_column)}{max(new_last, HDR_ROW + 1)}"
        if tbl.autoFilter is not None:
            tbl.autoFilter.ref = tbl.ref
    return new_last


def _sheet_part(z, name):
    """Caminho do XML da aba `name` dentro do pacote ZIP."""
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m = re.search(r'<sheet [^>]*name="%s"[^>]*/>' % re.escape(name), wbx)
    if not m:
        raise SystemExit(f"Aba '{name}' nao encontrada no pacote.")
    rid = re.search(r'r:id="([^"]+)"', m.group(0)).group(1)
    rel = re.search(r'<Relationship [^>]*Id="%s"[^>]*/>' % rid, rels).group(0)
    target = re.search(r'Target="([^"]+)"', rel).group(1).lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


def _rels_path(part):
    d, f = part.rsplit("/", 1)
    return f"{d}/_rels/{f}.rels"


def transplant_drawing(path, source, sheet_name=SHEET_MAIN):
    """Copia o desenho da aba (imagens e formas) de SOURCE para a aba de PATH."""
    with zipfile.ZipFile(source) as zs:
        src_part = _sheet_part(zs, sheet_name)
        srels_name = _rels_path(src_part)
        if srels_name not in zs.namelist():
            return
        srels = zs.read(srels_name).decode("utf-8")
        m = re.search(r'<Relationship [^>]*Type="[^"]*/drawing"[^>]*/>', srels)
        if not m:
            return
        d_target = re.search(r'Target="([^"]+)"', m.group(0)).group(1)
        d_part = posixpath.normpath(posixpath.join(posixpath.dirname(src_part), d_target))
        if d_part not in zs.namelist():
            return
        d_xml = zs.read(d_part).decode("utf-8")
        d_rels_name = _rels_path(d_part)
        d_rels = zs.read(d_rels_name).decode("utf-8") if d_rels_name in zs.namelist() else ""
        media = {}
        for t in re.findall(r'Target="([^"]+)"', d_rels):
            p = posixpath.normpath(posixpath.join(posixpath.dirname(d_part), t))
            if p in zs.namelist():
                media[t] = (p, zs.read(p))

    tmp = tempfile.mktemp(suffix=os.path.splitext(path)[1])
    with zipfile.ZipFile(path) as zin:
        out_part = _sheet_part(zin, sheet_name)
        ox = zin.read(out_part).decode("utf-8")
        names = set(zin.namelist())
        orels_name = _rels_path(out_part)
        orels = zin.read(orels_name).decode("utf-8") if orels_name in names else \
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

        old_drawing_part = None
        old = re.search(r'<drawing r:id="([^"]+)"\s*/>', ox)
        if old:
            ox = ox.replace(old.group(0), "")
            rel = re.search(r'<Relationship [^>]*Id="%s"[^>]*/>' % old.group(1), orels)
            if rel:
                t = re.search(r'Target="([^"]+)"', rel.group(0)).group(1)
                old_drawing_part = posixpath.normpath(posixpath.join(posixpath.dirname(out_part), t))
                orels = orels.replace(rel.group(0), "")

        n = 1
        while f"xl/drawings/drawingBase{n}.xml" in names:
            n += 1
        new_files = {}
        new_d_name = f"xl/drawings/drawingBase{n}.xml"
        new_d_rels_name = f"xl/drawings/_rels/drawingBase{n}.xml.rels"
        names.add(new_d_name)
        names.add(new_d_rels_name)

        rels_out = []
        k = 1
        exts = set()
        for old_t, (p, data) in media.items():
            ext = os.path.splitext(p)[1]
            while f"xl/media/botAimagem{k}{ext}" in names:
                k += 1
            newp = f"xl/media/botAimagem{k}{ext}"
            new_files[newp] = data
            names.add(newp)
            exts.add(ext[1:].lower())
            rels_out.append(f'<Relationship Id="rIdImg{k}" Type="{REL_NS}/image" Target="../media/botAimagem{k}{ext}"/>')
            old_id = re.search(r'Id="([^"]+)"[^>]*Target="%s"' % re.escape(old_t), d_rels)
            if old_id:
                d_xml = re.sub(r'r:embed="%s"' % re.escape(old_id.group(1)), f'r:embed="rIdImg{k}"', d_xml)
                d_xml = re.sub(r'r:link="%s"' % re.escape(old_id.group(1)), f'r:link="rIdImg{k}"', d_xml)
            k += 1

        new_files[new_d_name] = d_xml.encode("utf-8")
        new_files[new_d_rels_name] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(rels_out) + "</Relationships>").encode("utf-8")

        existing_rids = set(re.findall(r'Id="([^"]+)"', orels))
        idx = 1
        while f"rIdDrw{idx}" in existing_rids:
            idx += 1
        rid = f"rIdDrw{idx}"

        orels = orels.replace("</Relationships>",
            f'<Relationship Id="{rid}" Type="{REL_NS}/drawing" Target="../drawings/drawingBase{n}.xml"/></Relationships>')
        tag = f'<drawing r:id="{rid}"/>'
        if 'xmlns:r=' not in ox[:600]:
            ox = ox.replace("<worksheet ", f'<worksheet xmlns:r="{REL_NS}" ', 1)
        for marker in ("<legacyDrawing", "<legacyDrawingHF", "<picture", "<oleObjects", "<controls", "<tableParts", "<extLst", "</worksheet>"):
            if marker in ox:
                ox = ox.replace(marker, tag + marker, 1)
                break

        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if old_drawing_part and item.filename == old_drawing_part:
                    continue
                if old_drawing_part and item.filename == _rels_path(old_drawing_part):
                    continue
                if item.filename == out_part:
                    data = ox.encode("utf-8")
                elif item.filename == orels_name:
                    data = orels.encode("utf-8")
                elif item.filename == "[Content_Types].xml":
                    t = zin.read(item.filename).decode("utf-8")
                    if old_drawing_part:
                        t = re.sub(r'<Override PartName="/%s"[^>]*/>' % re.escape(old_drawing_part), "", t)
                    for e in exts:
                        mime = {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg",
                                "gif": "image/gif", "emf": "image/x-emf", "wmf": "image/x-wmf"}.get(e)
                        if mime and f'Extension="{e}"' not in t:
                            t = t.replace("</Types>", f'<Default Extension="{e}" ContentType="{mime}"/></Types>')
                    if f'PartName="/{new_d_name}"' not in t:
                        t = t.replace("</Types>",
                            f'<Override PartName="/{new_d_name}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
                    data = t.encode("utf-8")
                else:
                    data = zin.read(item.filename)
                zout.writestr(item.filename, data)
            if orels_name not in names:
                zout.writestr(orels_name, orels.encode("utf-8"))
            for p, data in new_files.items():
                zout.writestr(p, data)
    shutil.move(tmp, path)


def execute_merge(base_path, path_a, path_b, out_path, prefer="minha",
                  name_a="Arquivo 1", name_b="Arquivo 2", conflicts_list=None):
    """
    Executa a unificacao completa das planilhas.
    Devolve: (total_linhas, lista_de_conflitos_str, lista_de_conflitos_estruturados)
    """
    want_xlsm = out_path.lower().endswith(".xlsm") or base_path.lower().endswith(".xlsm")
    wbase = openpyxl.load_workbook(base_path, data_only=False)
    wa = openpyxl.load_workbook(path_a, data_only=False)
    wb = openpyxl.load_workbook(path_b, data_only=False)
    wout = openpyxl.load_workbook(base_path, keep_vba=want_xlsm)

    log = []
    if conflicts_list is None:
        conflicts_list = []

    try:
        result, merged, hd = merge_main(wa, wb, wbase, prefer, log, name_a, name_b, conflicts_list)
        write_main(wout, result, merged, hd, None)

        for name in wbase.sheetnames:
            if name == SHEET_MAIN:
                continue
            changes = merge_cells(wa, wb, wbase, name, prefer, log, name_a, name_b, conflicts_list)
            ws = wout[name]
            for (r, c), v in changes.items():
                ws.cell(r, c).value = v

        # Validacao de lista suspensa dos produtos (linha 3)
        ws_main = wout[SHEET_MAIN]
        existentes = [d for d in ws_main.data_validations.dataValidation if "K3" in str(d.sqref)]
        for d in existentes:
            ws_main.data_validations.dataValidation.remove(d)
        dv = DataValidation(type="list", formula1="PRODUTOS!$C$2:$C$200", allow_blank=True)
        dv.add("K3:BB3")
        ws_main.add_data_validation(dv)

        wout.calculation.fullCalcOnLoad = True
        wout.save(out_path)
    finally:
        try: wbase.close()
        except: pass
        try: wa.close()
        except: pass
        try: wb.close()
        except: pass
        try: wout.close()
        except: pass

    # Transplanta desenhos/imagens (logos neoenergia, etc.) de BASE para a saida
    if want_xlsm:
        with zipfile.ZipFile(base_path) as _zs:
            _wbx = _zs.read("xl/workbook.xml").decode("utf-8")
        for _nome in re.findall(r'<sheet [^>]*name="([^"]+)"', _wbx):
            nome = html.unescape(_nome)
            transplant_drawing(out_path, base_path, nome)

    return len(result), log, conflicts_list


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("base")
    ap.add_argument("minha")
    ap.add_argument("dela")
    ap.add_argument("-o", "--saida", default="unificada.xlsm")
    ap.add_argument("--prefere", choices=["minha", "dela"], default="minha")
    ap.add_argument("--relatorio", default="conflitos.txt")
    args = ap.parse_args()

    for caminho in (args.base, args.minha, args.dela):
        if not os.path.isfile(caminho):
            sys.exit(f"Arquivo nao encontrado: {caminho} (confira o nome exato e a pasta)")

    total_linhas, log, structured = execute_merge(args.base, args.minha, args.dela, args.saida, args.prefere)

    with open(args.relatorio, "w", encoding="utf-8") as f:
        f.write(f"Merge gerado em {args.saida}\n")
        f.write(f"Preferencia em conflitos: {args.prefere}\n")
        f.write(f"Linhas na unificada: {total_linhas}\n")
        f.write(f"Conflitos: {len(log)}\n\n")
        for line in log:
            f.write(line + "\n")

    print(f"OK: {args.saida} | linhas={total_linhas} | conflitos={len(log)} -> {args.relatorio}")


if __name__ == "__main__":
    sys.exit(main())