#!/usr/bin/env python3
"""
Combina dois arquivos .xlsx em um único arquivo sem reescrever os dados.
Útil quando um arquivo tem as abas de análise (pequeno, rápido de gerar)
e outro tem as bases consolidadas (grande, >200k linhas).

Uso:
  python combinar_xlsx.py --principal apuracao.xlsx --adicional bases.xlsx --output final.xlsx

O arquivo 'principal' recebe as abas do 'adicional' ao final.
Faz a mesclagem correta de styles.xml, workbook.xml, rels e Content_Types.xml.
"""

import argparse
import re
import zipfile
from copy import deepcopy

from lxml import etree

NS     = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT_NS  = 'http://schemas.openxmlformats.org/package/2006/content-types'
WS_CT  = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'


def mesclar_styles(styles_p, styles_a):
    """
    Appenda os estilos do arquivo adicional no principal,
    retornando o xf_offset (índice base para remapear as células do adicional).
    """
    def merge_list(tag, sd, sa):
        el_d = sd.find(f'{{{NS}}}{tag}')
        el_a = sa.find(f'{{{NS}}}{tag}')
        if el_d is None or el_a is None:
            return 0
        offset = len(el_d)
        for child in el_a:
            el_d.append(deepcopy(child))
        el_d.set('count', str(len(el_d)))
        return offset

    fo = merge_list('fonts',   styles_p, styles_a)
    fi = merge_list('fills',   styles_p, styles_a)
    bo = merge_list('borders', styles_p, styles_a)

    xfs_p = styles_p.find(f'{{{NS}}}cellXfs')
    xfs_a = styles_a.find(f'{{{NS}}}cellXfs')
    xf_offset = len(xfs_p)

    for xf in xfs_a:
        nx = deepcopy(xf)
        for attr, off in [('fontId', fo), ('fillId', fi), ('borderId', bo)]:
            if nx.get(attr):
                nx.set(attr, str(int(nx.get(attr)) + off))
        xfs_p.append(nx)
    xfs_p.set('count', str(len(xfs_p)))

    return xf_offset


def remap_styles(data, offset):
    """Substitui s="N" por s="N+offset" no XML da sheet."""
    return re.sub(
        r'\bs="(\d+)"',
        lambda m: f's="{int(m.group(1)) + offset}"',
        data.decode('utf-8')
    ).encode('utf-8')


def combinar(path_principal, path_adicional, path_output):
    with zipfile.ZipFile(path_principal, 'r') as zp, \
         zipfile.ZipFile(path_adicional,  'r') as za:

        # Descobrir quantas sheets o principal tem
        sheets_p = sorted(
            [n for n in zp.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', n)]
        )
        n_sheets_p = len(sheets_p)

        # Sheets do adicional
        sheets_a = sorted(
            [n for n in za.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', n)]
        )

        # Mesclar styles
        styles_p = etree.fromstring(zp.read('xl/styles.xml'))
        styles_a = etree.fromstring(za.read('xl/styles.xml'))
        xf_offset = mesclar_styles(styles_p, styles_a)
        new_styles = etree.tostring(styles_p, xml_declaration=True, encoding='UTF-8', standalone=True)

        # Remap e cache das sheets do adicional
        new_sheet_data = {}
        for i, sa_sheet in enumerate(sheets_a):
            new_idx = n_sheets_p + i + 1
            new_name = f'xl/worksheets/sheet{new_idx}.xml'
            new_sheet_data[new_name] = remap_styles(za.read(sa_sheet), xf_offset)

        # workbook.xml do adicional: pegar nomes das abas
        wb_a = etree.fromstring(za.read('xl/workbook.xml'))
        sheets_el_a = wb_a.find(f'{{{NS}}}sheets')

        # workbook.xml do principal: adicionar as novas sheets
        wb_p = etree.fromstring(zp.read('xl/workbook.xml'))
        sheets_el_p = wb_p.find(f'{{{NS}}}sheets')

        # Descobrir maior rId já usado
        rels_p = etree.fromstring(zp.read('xl/_rels/workbook.xml.rels'))
        rids_used = [int(r.get('Id', 'rId0').replace('rId', '')) for r in rels_p]
        next_rid = max(rids_used, default=0) + 1

        # Mapeamento: rId_original_adicional → novo_rId, novo_sheet_name, novo_sheetId
        sheets_a_list = list(sheets_el_a)
        rid_map = {}
        for i, s in enumerate(sheets_a_list):
            old_rid = s.get(f'{{{REL_NS}}}id')
            new_rid = f'rId{next_rid + i}'
            new_target = f'sheet{n_sheets_p + i + 1}.xml'
            new_sheetid = str(n_sheets_p + i + 1)
            rid_map[old_rid] = (new_rid, new_target, new_sheetid)

        for s in sheets_a_list:
            old_rid = s.get(f'{{{REL_NS}}}id')
            new_rid, _, new_sheetid = rid_map[old_rid]
            ns = deepcopy(s)
            ns.set(f'{{{REL_NS}}}id', new_rid)
            ns.set('sheetId', new_sheetid)
            sheets_el_p.append(ns)

        new_wb = etree.tostring(wb_p, xml_declaration=True, encoding='UTF-8', standalone=True)

        # rels: adicionar novas relações
        for r in etree.fromstring(za.read('xl/_rels/workbook.xml.rels')):
            if 'worksheet' in r.get('Type', ''):
                old_rid = r.get('Id')
                new_rid, new_target, _ = rid_map[old_rid]
                nr = deepcopy(r)
                nr.set('Id', new_rid)
                nr.set('Target', f'worksheets/{new_target}')
                rels_p.append(nr)

        new_rels = etree.tostring(rels_p, xml_declaration=True, encoding='UTF-8', standalone=True)

        # Content_Types: registrar novas sheets
        ct_p = etree.fromstring(zp.read('[Content_Types].xml'))
        for new_name in new_sheet_data:
            el = etree.SubElement(ct_p, f'{{{CT_NS}}}Override')
            el.set('PartName', f'/{new_name}')
            el.set('ContentType', WS_CT)

        new_ct = etree.tostring(ct_p, xml_declaration=True, encoding='UTF-8', standalone=True)

        # Montar arquivo final
        with zipfile.ZipFile(path_output, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zp.namelist():
                if   item == 'xl/workbook.xml':             zout.writestr(item, new_wb)
                elif item == 'xl/_rels/workbook.xml.rels':  zout.writestr(item, new_rels)
                elif item == 'xl/styles.xml':               zout.writestr(item, new_styles)
                elif item == '[Content_Types].xml':         zout.writestr(item, new_ct)
                else:                                       zout.writestr(item, zp.read(item))
            for name, data in new_sheet_data.items():
                zout.writestr(name, data)

    print(f"✅ Combinado: {path_output}")


def main():
    parser = argparse.ArgumentParser(description='Combina dois xlsx em um')
    parser.add_argument('--principal', required=True)
    parser.add_argument('--adicional', required=True)
    parser.add_argument('--output',    required=True)
    args = parser.parse_args()
    combinar(args.principal, args.adicional, args.output)


if __name__ == '__main__':
    main()
