# Convert_IMCA (v1.6.0)

[![Release](https://img.shields.io/badge/Release-v1.6.0-orange.svg)](https://github.com/DevolperGdSL/Convert_IMCA/releases/tag/v1.6.0)
[![Plataformas](https://img.shields.io/badge/Plataformas-Windows%20%7C%20Linux-blue.svg)](https://github.com/DevolperGdSL/Convert_IMCA/releases)
[![Rust](https://img.shields.io/badge/Engine-Rust%20%2B%20Python-red.svg)](https://www.rust-lang.org/)
[![License](https://img.shields.io/badge/Licen%C3%A7a-MIT-green.svg)](LICENSE)

> Conversor de alta performance e interface gráfica moderna para transformar planilhas Excel (`.xlsx`) no formato ultra-leve e otimizado **`.IMCA`** (*Interchange Model for Calendar & Agenda*), desenvolvido sob medida para alimentar o aplicativo de escala e discipulado [cabeceira-pwa1](https://github.com/DevolperGdSL/cabeceirapwa-1).

---

## 📥 Baixar o Aplicativo (Downloads Oficiais)

Os instaladores e executáveis autônomos prontos para uso são gerados pela nossa esteira de integração contínua e publicados diretamente na aba de **[Releases do GitHub](https://github.com/DevolperGdSL/Convert_IMCA/releases/tag/v1.6.0)**:

| Plataforma | Pacote / Executável | Como Executar | Link Direto |
| :--- | :--- | :--- | :--- |
| **🪟 Windows (64-bit)** | `Convert_IMCA_GUI-windows.exe` (15.2 MB) | Baixar e dar duplo clique (Zero-Setup) | [⬇️ Baixar .exe](https://github.com/DevolperGdSL/Convert_IMCA/releases/download/v1.6.0/Convert_IMCA_GUI-windows.exe) |
| **🐧 Linux (64-bit)** | `Convert_IMCA-linux-x86_64.AppImage` (20.7 MB) | Duplo clique no Ubuntu/Debian/Fedora | [⬇️ Baixar .AppImage](https://github.com/DevolperGdSL/Convert_IMCA/releases/download/v1.6.0/Convert_IMCA-linux-x86_64.AppImage) |
| **🐧 Linux (Atalho)** | `Convert_IMCA.image` (20.7 MB) | Cópia com extensão simplificada `.image` | [⬇️ Baixar .image](https://github.com/DevolperGdSL/Convert_IMCA/releases/download/v1.6.0/Convert_IMCA.image) |
| **⚡ Terminal Windows** | `convert_imca-windows-x86_64.exe` (1.2 MB) | CLI nativo em Rust ultra-rápido | [⬇️ Baixar CLI Windows](https://github.com/DevolperGdSL/Convert_IMCA/releases/download/v1.6.0/convert_imca-windows-x86_64.exe) |
| **⚡ Terminal Linux** | `convert_imca-linux-x86_64` (1.4 MB) | CLI nativo em Rust ultra-rápido | [⬇️ Baixar CLI Linux](https://github.com/DevolperGdSL/Convert_IMCA/releases/download/v1.6.0/convert_imca-linux-x86_64) |

> [!NOTE]
> **Por que os arquivos executáveis não ficam misturados com o código-fonte no Git?**  
> O Git foi projetado para versionar texto e código. Arquivos executáveis binários (.exe e .AppImage) possuem de 15 a 165 MB; se fossem colocados no repositório de código, inflariam o clone para centenas de megabytes. Por isso, o `.gitignore` ignora pastas de compilação locais (`/dist/`, `/target/`) e o GitHub disponibiliza a aba **[Releases](https://github.com/DevolperGdSL/Convert_IMCA/releases)** para download oficial e seguro dos executáveis.

---

## 📊 Estrutura Esperada da Planilha Excel (`.xlsx`)

O conversor foi desenhado para ser resiliente e inteligente. Ele inspeciona automaticamente as primeiras 10 linhas da planilha à procura da linha de cabeçalho, reconhecendo as colunas mesmo que a ordem ou a caixa das letras variem.

### 📋 Colunas Reconhecidas pelo Motor

| Coluna | Nomes Aceitos no Cabeçalho | Obrigatória? | Descrição e Exemplos de Conteúdo |
| :--- | :--- | :---: | :--- |
| **Data** | `DATA`, `DATAS`, `DATE` | **SIM** | Data do evento. Aceita formato serial numérico nativo do Excel (ex: `46024`), formato brasileiro `DD/MM/AAAA` (ex: `02/01/2026`) ou padrão ISO `AAAA-MM-DD`. |
| **Nome do Evento** | `NOME DO EVENTO`, `EVENTO`, `NOME`, `TITLE` | **SIM** | Título do culto ou programação. Ex: `12 Dias de Clamor para 12 Meses de Milagres`, `Culto da Família`, `Santa Ceia`. |
| **Dia da Semana** | `DIA`, `DIA DA SEMANA`, `SEMANA`, `DAY` | Não | Nome do dia. Ex: `Sexta-feira`, `Domingo`, `Monday`. O motor traduz automaticamente dias em inglês para português. |
| **Hora** | `HORA`, `HORÁRIO`, `TIME` | Não | Horário de início do culto. Ex: `19:30`, `10:00`. **Se estiver vazio ou omitido**, o sistema atribui automaticamente as **19:30** como padrão. |
| **Local** | `LOCAL`, `ESPAÇO`, `ROOM`, `LOCALIZAÇÃO` | Não | Espaço físico da igreja. Ex: `Templo`, `Salão Social`, `Anexo 1`. Se vazio, assume `N/A`. |
| **Solicitante / Responsável** | `SOLICITANTE`, `RESPONSÁVEL`, `DEPARTAMENTO` | Não | Secretaria ou ministério encarregado. Ex: `Secretaria do discipulado`, `Mocidade`. Se vazio, assume `N/A`. |

### 🔍 Exemplo Visual de Planilha Válida

| DATA | DIA | HORA | NOME DO EVENTO | LOCAL | SOLICITANTE |
| :---: | :---: | :---: | :--- | :---: | :--- |
| 31/12/2025 | Quarta-feira | 19:30 | Reveillon cabeceira | Templo | Secretaria Pastoral |
| 02/01/2026 | Sexta-feira | 19:30 | 12 Dias de Clamor para 12 Meses de Milagres | Templo | Secretaria do discipulado |
| 03/01/2026 | Sábado | 19:30 | 12 Dias de Clamor para 12 Meses de Milagres | Templo | Secretaria do discipulado |
| 11/01/2026 | Domingo | 10:00 | Escola Bíblica Dominical | Salão | Ministério de Ensino |

### 🧹 Higienização e Descarte Automático de Lixo
- **Linhas em branco:** Ignoradas silenciosamente.
- **Linhas sem data ou sem nome:** Descartadas para preservar a integridade do banco.
- **Avisos internos de planilha:** Textos administrativos comuns em planilhas da igreja (como *"SEMPRE REORDENAR DO MAIS ANTIGO PARA O MAIS RECENTE"*) são filtrados e descartados automaticamente.

---

## 🖥️ Como Usar a Interface Gráfica

A interface foi construída na estética geométrica **Black & Orange** (`BLACK #171717` e `ORANGE #F25623`), com fluxo simples de 3 etapas:

```mermaid
graph LR
    Tela1["Tela 1: Entrada\nBotão Único de Busca"] -->|Seleciona .xlsx| Tela2["Tela 2: Processamento\nChama Viva Respirando"]
    Tela2 -->|Conversão Instantânea| Tela3["Tela 3: Conclusão\nPrévia & Mostrar Arquivo"]
```

1. **Tela 1 (Entrada):** Clique no botão central laranja **`selecionar planilha .xlsx`**. A janela nativa de busca de arquivos do seu sistema operacional (Windows Explorer ou GNOME/Zenity no Linux) se abrirá.
2. **Tela 2 (Carregamento):** A Chama do Espírito Santo se acende em uma animação suave de respiração enquanto o motor extrai, higieniza e converte os dados em milissegundos.
3. **Tela 3 (Conclusão):**
   - Clique em **`mostrar prévia`** para inspecionar a tabela completa de cultos com busca em tempo real, resumo de eventos e porcentagem de redução de peso.
   - Clique em **`mostrar arquivo gerado`** para abrir o gerenciador de arquivos com o `.imca` selecionado pronto para uso.

---

## ⚡ Como Usar via Linha de Comando (CLI)

Se preferir utilizar em scripts ou servidores, o Convert_IMCA disponibiliza motores de linha de comando:

### 1. No Linux ou Windows com o Binário Nativo Rust (Ultra-Rápido ~6 ms)
```bash
# Converter com exibição de prévia no terminal e validação:
./convert_imca "Calendario 2026.xlsx" --preview --validate

# Definir arquivo de saída personalizado e horário padrão:
./convert_imca "Calendario 2026.xlsx" -o "saida.imca" -t 20:00
```

### 2. Com o Motor Portátil em Python (Zero Instalação)
```bash
python3 convert_imca.py "Calendario 2026.xlsx" --preview
```

---

## 📄 O Que é o Formato `.IMCA`?

O formato `.IMCA` é uma representação textual estruturada, dividida em três blocos delimitados por pipe (`|`):

```ini
[METADATA]
version=1.6.0
source_file=Calendario 2026.xlsx
total_events=318
generated_at=2026-10-05T13:16:53
base_year=2026
default_time=19:30

[SCHEMA]
fields=data|dia_semana|horario|evento|local|solicitante
types=date|string|time|string|string|string

[DATA]
2025-12-31|Wednesday|19:30|Reveillon cabeceira|N/A|N/A
2026-01-02|Friday|19:30|12 Dias de Clamor para 12 Meses de Milagres|Templo|Secretaria do discipulado
2026-01-03|Saturday|19:30|12 Dias de Clamor para 12 Meses de Milagres|Templo|Secretaria do discipulado
```

### 🚀 Tabela de Economia de Dados Real

| Métrica | Planilha Original (.xlsx) | Formato Texto (.IMCA) | Formato Web (Gzip/Brotli) |
| :--- | :---: | :---: | :---: |
| **Peso do Arquivo** | **96.8 KB** | **25.8 KB** | **4.4 KB** |
| **Redução Líquida** | Base (0%) | **-73.4% de economia** | **-95.5% de economia** |
| **Tempo de Parsing no Celular** | ~280 ms (via SheetJS) | **1.1 ms (TypeScript puro)** | **Instantâneo** |
| **Dependências Externas no PWA** | 1.2 MB de bibliotecas | **0 KB (Zero dependências)** | **Zero** |

---

## 📖 Livro de Documentação (`DOCUMENTACAO.md`)

Para entender toda a engenharia, história e decisões arquiteturais do projeto, leia o nosso livro completo em [`DOCUMENTACAO.md`](DOCUMENTACAO.md), dividido rigidamente em 8 capítulos didáticos:

- **Capítulo 1:** O Surgimento do Formato `.IMCA`
- **Capítulo 2:** A Anatomia Técnica do Formato `.IMCA`
- **Capítulo 3:** A Engenharia dos Conversores (Rust & Python)
- **Capítulo 4:** Guia de Uso, Compilação e Distribuição Multiplataforma
- **Capítulo 5:** Integração com o Aplicativo `cabeceira-pwa1`
- **Capítulo 6:** Segurança, Governança e Ciclo de Vida
- **Capítulo 7:** A Interface Gráfica Desktop e o Empacotamento Multiplataforma
- **Capítulo 8:** A Identidade Visual do Sistema — A Chama, o Calendário e a Nuvem CI/CD
