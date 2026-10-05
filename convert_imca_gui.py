#!/usr/bin/env python3
"""
=============================================================================
CONVERT_IMCA - Interface Gráfica Premium (v1.2.0)
UX/UI Baseada no Conceito Visual:
- Tela 1: Entrada com Pílula "BUSQUE O ARQUIVO XLSX" e Botão "Buscar arquivo"
- Tela 2: Carregamento com a Chama Viva Animada (Breathing Glow)
- Tela 3: Saída com Pílulas "Mostrar prévia" e "Mostrar arquivo gerado"
=============================================================================
"""

import os
import sys
import json
import subprocess
import webview
from datetime import datetime

# Importa o motor de conversão existente
try:
    from convert_imca import parse_xlsx, build_imca_content, __version__
except ImportError:
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from convert_imca import parse_xlsx, build_imca_content, __version__

UI_HTML_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "ui.html")

class ConverterApi:
    def __init__(self, window=None):
        self.window = window

    def set_window(self, window):
        self.window = window

    def select_file(self):
        """Abre o seletor nativo do SO e retorna o caminho selecionado"""
        if not self.window:
            return None
        file_types = ("Planilhas Excel (*.xlsx)", "Todos os Arquivos (*.*)")
        dialog_type = getattr(webview, "FileDialog", None)
        open_flag = dialog_type.OPEN if dialog_type else getattr(webview, "OPEN_DIALOG", 10)
        result = self.window.create_file_dialog(open_flag, allow_multiple=False, file_types=file_types)
        if result and len(result) > 0:
            return result[0]
        return None

    def convert_file(self, file_path, default_time="19:30", translate_days=True):
        """Executa a conversão ultra-rápida e retorna os dados em JSON para a interface"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Arquivo '{file_path}' não foi encontrado.")

        start_time = datetime.now()
        out_path = os.path.splitext(file_path)[0] + ".imca"

        metadata, events = parse_xlsx(
            file_path=file_path,
            default_time=default_time or "19:30",
            translate_days=bool(translate_days)
        )

        imca_content = build_imca_content(metadata, events)

        with open(out_path, "w", encoding="utf-8") as f:
            f.write(imca_content)

        orig_size = os.path.getsize(file_path)
        imca_size = len(imca_content.encode("utf-8"))
        reduction = ((orig_size - imca_size) / orig_size) * 100 if orig_size else 0
        elapsed_ms = (datetime.now() - start_time).total_seconds() * 1000

        res = {
            "success": True,
            "output_path": os.path.abspath(out_path),
            "total_events": len(events),
            "orig_size": orig_size,
            "imca_size": imca_size,
            "reduction": reduction,
            "elapsed_ms": elapsed_ms,
            "events": events[:100]
        }
        return json.dumps(res)

    def open_folder(self, output_path):
        """Abre o explorador de arquivos local com o arquivo gerado selecionado"""
        if not output_path or not os.path.exists(output_path):
            return False

        folder = os.path.dirname(os.path.abspath(output_path))
        if sys.platform.startswith("win"):
            os.system(f'explorer /select,"{os.path.abspath(output_path)}"')
        elif sys.platform.startswith("darwin"):
            subprocess.run(["open", "-R", output_path])
        else:
            subprocess.run(["xdg-open", folder])
        return True


def main():
    if not os.path.exists(UI_HTML_PATH):
        print(f"Erro: Arquivo de interface '{UI_HTML_PATH}' não encontrado.")
        sys.exit(1)

    with open(UI_HTML_PATH, "r", encoding="utf-8") as f:
        html_content = f.read()

    api = ConverterApi()
    window = webview.create_window(
        title="Convert IMCA",
        html=html_content,
        js_api=api,
        width=780,
        height=680,
        min_size=(680, 560),
        background_color="#000000",
        resizable=True
    )
    api.set_window(window)
    webview.start(debug=False)


if __name__ == "__main__":
    main()
