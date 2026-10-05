#!/usr/bin/env python3
"""
=============================================================================
CONVERT_IMCA - Versão Portátil (Zero-Dependências Externas)
Conversor de Planilhas Excel (.xlsx) para o Formato Leve .IMCA (cabeceira-pwa1)
Compatível com Linux, Windows e macOS via Python 3.7+
=============================================================================
"""

import sys
import os
import argparse
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

__version__ = "1.3.0"

WEEKDAY_TRANSLATION = {
    "monday": "Segunda-feira",
    "mon": "Segunda-feira",
    "segunda": "Segunda-feira",
    "tuesday": "Terça-feira",
    "tue": "Terça-feira",
    "terça": "Terça-feira",
    "terca": "Terça-feira",
    "wednesday": "Quarta-feira",
    "wed": "Quarta-feira",
    "quarta": "Quarta-feira",
    "thursday": "Quinta-feira",
    "thu": "Quinta-feira",
    "quinta": "Quinta-feira",
    "friday": "Sexta-feira",
    "fri": "Sexta-feira",
    "sexta": "Sexta-feira",
    "saturday": "Sábado",
    "sat": "Sábado",
    "sábado": "Sábado",
    "sabado": "Sábado",
    "sunday": "Domingo",
    "sun": "Domingo",
    "domingo": "Domingo",
}

def excel_serial_to_date(serial: float) -> datetime:
    """Converte número de série de data do Excel (base 1899-12-30) para objeto datetime."""
    # O Excel simula 1900 como ano bissexto (bug histórico do Lotus 1-2-3)
    return datetime(1899, 12, 30) + timedelta(days=serial)

def extract_year_from_filename(filename: str) -> int:
    """Detecta ano de 4 dígitos presente no nome do arquivo (ex: Calendario 2026.xlsx -> 2026)."""
    import re
    matches = re.findall(r'\b(20\d\d)\b', filename)
    if matches:
        return int(matches[0])
    return 2026

def parse_xlsx(file_path: str, sheet_target: str = None, fallback_year: int = None, default_time: str = "19:30", translate_days: bool = False):
    """Extrai registros de eventos do arquivo .xlsx sem usar nenhuma dependência externa."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Arquivo '{file_path}' não foi encontrado.")

    year = fallback_year or extract_year_from_filename(os.path.basename(file_path))

    with zipfile.ZipFile(file_path, 'r') as z:
        # 1. Carregar tabela de strings compartilhadas (sharedStrings.xml)
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            sst_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in sst_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                texts = [t.text for t in si.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if t.text]
                shared_strings.append(''.join(texts))

        # 2. Descobrir abas disponíveis (workbook.xml)
        wb_root = ET.fromstring(z.read('xl/workbook.xml'))
        sheets = []
        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
              'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
        
        for idx, s in enumerate(wb_root.findall('.//main:sheet', ns)):
            s_name = s.attrib.get('name')
            r_id = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            sheets.append((idx + 1, s_name, r_id))

        if not sheets:
            raise ValueError("Nenhuma aba encontrada na planilha Excel.")

        # Selecionar a aba adequada
        chosen_sheet_file = None
        chosen_sheet_name = None

        if sheet_target:
            for s_num, s_name, _ in sheets:
                if s_name.lower() == sheet_target.lower():
                    chosen_sheet_file = f'xl/worksheets/sheet{s_num}.xml'
                    chosen_sheet_name = s_name
                    break
            if not chosen_sheet_file:
                available = [s[1] for s in sheets]
                raise ValueError(f"Aba '{sheet_target}' não encontrada. Abas disponíveis: {available}")
        else:
            # Busca prioritária por 'Calendario', 'Agenda', 'Eventos'
            for s_num, s_name, _ in sheets:
                n = s_name.lower()
                if 'calendario' in n or 'agenda' in n or 'evento' in n:
                    chosen_sheet_file = f'xl/worksheets/sheet{s_num}.xml'
                    chosen_sheet_name = s_name
                    break
            if not chosen_sheet_file:
                chosen_sheet_file = f'xl/worksheets/sheet{sheets[0][0]}.xml'
                chosen_sheet_name = sheets[0][1]

        # 3. Ler linhas da aba selecionada
        ws_root = ET.fromstring(z.read(chosen_sheet_file))
        rows = ws_root.findall('.//main:row', ns)

        events = []
        seq = 1

        # Detectar cabeçalho
        col_data = 'B'
        col_dia = 'C'
        col_evento = 'D'
        col_local = 'E'
        col_solicitante = 'F'
        data_rows = rows

        for r_idx, r in enumerate(rows[:10]):
            r_cells = {}
            for c in r.findall('main:c', ns):
                ref = c.attrib.get('r', '')
                col = ''.join([ch for ch in ref if ch.isalpha()])
                t = c.attrib.get('t')
                v_el = c.find('main:v', ns)
                val = v_el.text if v_el is not None and v_el.text is not None else ''
                if t == 's' and val.isdigit() and int(val) < len(shared_strings):
                    val = shared_strings[int(val)]
                r_cells[col] = (val or '').strip().lower()

            for col, text in r_cells.items():
                if text in ['data', 'datas']:
                    col_data = col
                elif 'evento' in text or 'nome' in text:
                    col_evento = col
                elif 'dia' in text or 'semana' in text:
                    col_dia = col
                elif 'local' in text:
                    col_local = col
                elif 'solicitante' in text or 'responsavel' in text:
                    col_solicitante = col

            if any(text in ['data', 'datas'] for text in r_cells.values()):
                data_rows = rows[r_idx + 1:]
                break

        for r in data_rows:
            cells = {}
            for c in r.findall('main:c', ns):
                ref = c.attrib.get('r', '')
                col = ''.join([ch for ch in ref if ch.isalpha()])
                t = c.attrib.get('t')
                v_el = c.find('main:v', ns)
                val = v_el.text if v_el is not None and v_el.text is not None else ''
                if t == 's' and val.isdigit() and int(val) < len(shared_strings):
                    val = shared_strings[int(val)]
                cells[col] = (val or '').strip()

            raw_evento = cells.get(col_evento, '')
            raw_data = cells.get(col_data, '')
            raw_dia = cells.get(col_dia, '')
            raw_local = cells.get(col_local, '')
            raw_solic = cells.get(col_solicitante, '')

            if not raw_evento or not raw_data:
                continue

            # Ignora notas e avisos internos de planilha
            if any(k in raw_evento.upper() for k in ['SEMPRE REORDENAR', 'DO MAIS ANTIGO']):
                continue

            # Resolução de data
            iso_date = None
            try:
                # Tenta float serial Excel
                f_val = float(raw_data)
                dt = excel_serial_to_date(f_val)
                iso_date = dt.strftime('%Y-%m-%d')
            except ValueError:
                # Tenta parse textual
                for fmt in ['%Y-%m-%d', '%d/%m/%Y']:
                    try:
                        dt = datetime.strptime(raw_data, fmt)
                        iso_date = dt.strftime('%Y-%m-%d')
                        break
                    except ValueError:
                        pass
                if not iso_date and '/' in raw_data:
                    parts = raw_data.split('/')
                    if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                        iso_date = f"{year:04d}-{int(parts[1]):02d}-{int(parts[0]):02d}"

            if not iso_date:
                continue

            # Dia da semana
            day_text = raw_dia
            if not day_text:
                dt_obj = datetime.strptime(iso_date, '%Y-%m-%d')
                day_text = dt_obj.strftime('%A')

            if translate_days:
                day_text = WEEKDAY_TRANSLATION.get(day_text.lower(), day_text)

            # Higienização de pipes e quebras de linha
            clean_name = raw_evento.replace('|', '-').replace('\n', ' ').replace('\r', '').strip()
            clean_loc = raw_local.replace('|', '-').replace('\n', ' ').replace('\r', '').strip()
            clean_solic = raw_solic.replace('|', '-').replace('\n', ' ').replace('\r', '').strip()

            date_digits = iso_date.replace('-', '')
            ev_id = f"ev_{date_digits}_{seq:03d}"
            seq += 1

            events.append({
                'id': ev_id,
                'data': iso_date,
                'dia_semana': day_text,
                'horario': default_time,
                'nome': clean_name,
                'local': clean_loc,
                'solicitante': clean_solic,
                'status': 'publicado'
            })

    metadata = {
        'versao': '1.0',
        'ano': str(year),
        'total_registros': str(len(events)),
        'gerado_em': datetime.utcnow().isoformat() + 'Z',
        'arquivo_origem': os.path.basename(file_path),
        'aplicativo_alvo': 'cabeceira-pwa1',
        'delimitador': '|'
    }

    return metadata, events

def build_imca_content(metadata: dict, events: list) -> str:
    """Gera o arquivo de texto no padrão formal .IMCA."""
    lines = [
        "# =============================================================================",
        "# FORMATO DE DADOS IMCA v1.0 - AGENDA E EVENTOS OTIMIZADOS",
        "# APLICATIVO ALVO: cabeceira-pwa1",
        "# PADRÃO: DELIMITADO POR PIPE (|) COM CABEÇALHO E METADADOS INLINE",
        "# =============================================================================",
        "",
        "[METADATA]",
        f"versao={metadata['versao']}",
        f"ano={metadata['ano']}",
        f"total_registros={len(events)}",
        f"gerado_em={metadata['gerado_em']}",
        f"arquivo_origem={metadata['arquivo_origem']}",
        f"aplicativo_alvo={metadata['aplicativo_alvo']}",
        f"delimitador={metadata['delimitador']}",
        "",
        "[SCHEMA]",
        "id|data|dia_semana|horario|nome|local|solicitante|status",
        "",
        "[DATA]"
    ]

    for ev in events:
        lines.append(f"{ev['id']}|{ev['data']}|{ev['dia_semana']}|{ev['horario']}|{ev['nome']}|{ev['local']}|{ev['solicitante']}|{ev['status']}")

    return '\n'.join(lines) + '\n'

def main():
    parser = argparse.ArgumentParser(
        prog="convert_imca.py",
        description="Conversor Portátil Excel (.xlsx) -> Formato Leve .IMCA (cabeceira-pwa1)"
    )
    parser.add_argument("input", help="Arquivo Excel (.xlsx) de entrada")
    parser.add_argument("-o", "--output", help="Arquivo .imca de saída (padrão: mesmo nome com extensão .imca)")
    parser.add_argument("-s", "--sheet", help="Nome da aba a ser processada")
    parser.add_argument("-y", "--year", type=int, help="Ano base das datas (ex: 2026)")
    parser.add_argument("-t", "--default-time", default="19:30", help="Horário padrão (padrão: 19:30)")
    parser.add_argument("--translate-days", action="store_true", help="Traduz os dias da semana para Português")
    parser.add_argument("-p", "--preview", action="store_true", help="Exibe prévia dos primeiros 5 eventos")

    args = parser.parse_args()

    start_time = datetime.now()
    output_path = args.output or os.path.splitext(args.input)[0] + ".imca"

    print("============================================================")
    print(f" 🚀 CONVERSOR IMCA v{__version__} (Motor Portátil Python)")
    print("============================================================")
    print(f" 📂 Entrada : {args.input}")
    print(f" 💾 Saída   : {output_path}")

    try:
        metadata, events = parse_xlsx(
            file_path=args.input,
            sheet_target=args.sheet,
            fallback_year=args.year,
            default_time=args.default_time,
            translate_days=args.translate_days
        )
    except Exception as e:
        print(f"❌ Erro durante a conversão: {e}", file=sys.stderr)
        sys.exit(1)

    if not events:
        print("⚠️ Aviso: Nenhum evento válido encontrado na planilha.")
        sys.exit(0)

    imca_content = build_imca_content(metadata, events)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(imca_content)

    orig_size = os.path.getsize(args.input)
    imca_size = len(imca_content.encode("utf-8"))
    reduction = ((orig_size - imca_size) / orig_size) * 100 if orig_size else 0
    elapsed = (datetime.now() - start_time).total_seconds() * 1000

    print("------------------------------------------------------------")
    print(" 📊 ESTATÍSTICAS DE CONVERSÃO:")
    print(f"   • Total de Eventos Processados : {len(events)}")
    print(f"   • Ano Base Detectado           : {metadata['ano']}")
    print(f"   • Tamanho Planilha (.xlsx)     : {orig_size / 1024:.2f} KB ({orig_size} bytes)")
    print(f"   • Tamanho Formato (.imca)      : {imca_size / 1024:.2f} KB ({imca_size} bytes)")
    print(f"   • Redução Líquida de Dados     : {reduction:.1f}% de economia")
    print(f"   • Tempo de Execução Total      : {elapsed:.2f} ms")
    print("------------------------------------------------------------")

    if args.preview:
        print(f"\n🔍 PRÉVIA DOS PRIMEIROS EVENTOS (5 de {len(events)}):")
        for i, ev in enumerate(events[:5], start=1):
            loc = ev['local'] or 'N/A'
            sol = ev['solicitante'] or 'N/A'
            print(f"   [{i:02d}] {ev['data']} | {ev['nome']} ({ev['dia_semana']}) | {ev['horario']} | Local: '{loc}' | Solicitante: '{sol}'")
        print()

    print("✨ Sucesso! Arquivo pronto para consumo no cabeceira-pwa1.\n")

if __name__ == "__main__":
    main()
