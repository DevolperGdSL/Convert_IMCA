# Convert_IMCA (v1.0.0)

> Conversor de alta performance de planilhas Excel (`.xlsx`) para o formato ultra-leve e otimizado **`.IMCA`** (*Interchange Model for Calendar & Agenda*), projetado para abastecer o ecossistema e o aplicativo [cabeceira-pwa1](https://github.com/DevolperGdSL/cabeceirapwa-1).

---

## ⚡ Destaques de Desempenho

- **Redução de Tamanho:** De 96.8 KB (`.xlsx`) para **25.7 KB** em texto puro (**-73.4%**) e **4.4 KB** comprimido (**-95.5%**).
- **Ultra-Rápido:** Converte a planilha anual inteira (318 eventos) em **6.36 milissegundos** no motor nativo Rust e **45 ms** no motor portátil Python.
- **Zero Dependências no PWA:** Elimina bibliotecas pesadas de mais de 1.2 MB (como SheetJS), permitindo leitura em **1.11 ms** com JavaScript/TypeScript puro.
- **Multiplataforma Nativo:** Suporte completo e builds para **Linux** e **Windows** (`.exe`).

---

## 🚀 Como Usar

### 1. No Linux (Binário Nativo Rust)
```bash
# Compilar binário release:
cargo build --release

# Converter planilha com prévia:
./target/release/convert_imca "Calendario 2026.xlsx" --preview
```

### 2. No Windows (Executável ou Scripts)
- **Prompt / CMD:** `convert_imca.bat "Calendario 2026.xlsx"`
- **PowerShell:** `.\convert_imca.ps1 "Calendario 2026.xlsx" -p`
- **Binário Nativo:** `convert_imca.exe "Calendario 2026.xlsx"`

### 3. Motor Portátil Python (Zero-Setup)
Compatível com qualquer sistema que possua Python 3 instalado:
```bash
python3 convert_imca.py "Calendario 2026.xlsx" --preview
```

---

## 📖 Documentação Completa em Formato de Livro

Consulte o arquivo [`DOCUMENTACAO.md`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/DOCUMENTACAO.md) para o guia detalhado dividido em capítulos didáticos:
- **Capítulo 1:** O Surgimento do Formato `.IMCA`
- **Capítulo 2:** A Anatomia Técnica do Formato `.IMCA`
- **Capítulo 3:** A Engenharia dos Conversores (Rust & Python)
- **Capítulo 4:** Guia de Uso, Compilação e Distribuição Multiplataforma
- **Capítulo 5:** Integração com o Aplicativo `cabeceira-pwa1`
- **Capítulo 6:** Segurança, Governança e Ciclo de Vida
