# LIVRO DE DOCUMENTAÇÃO: PROJETO CONVERT_IMCA
## Guia Definitivo da Engenharia, Formato Textual Otimizado e Integração com o PWA Cabeceira

---

## Sumário Geral

- [Prólogo: A Missão e o Ponto de Partida](#prólogo-a-missão-e-o-ponto-de-partida)
- [Capítulo 1: O Surgimento do Formato .IMCA](#capítulo-1-o-surgimento-do-formato-imca)
  - [1.1 O Peso Oculto das Planilhas Modernas (.xlsx)](#11-o-peso-oculto-das-planilhas-modernas-xlsx)
  - [1.2 A Realidade dos Dispositivos Móveis e o Desafio do PWA](#12-a-realidade-dos-dispositivos-móveis-e-o-desafio-do-pwa)
  - [1.3 O Que é o Formato .IMCA e Por Que Ele Foi Criado](#13-o-que-é-o-formato-imca-e-por-que-ele-foi-criado)
  - [1.4 Benchmark Real: O Salto de 96.8 KB para 25.7 KB (e 4.4 KB Comprimido)](#14-benchmark-real-o-salto-de-968-kb-para-257-kb-e-44-kb-comprimido)
- [Capítulo 2: A Anatomia Técnica do Formato .IMCA](#capítulo-2-a-anatomia-técnica-do-formato-imca)
  - [2.1 A Estrutura Visual e a Escolha do Delimitador Pipe (|)](#21-a-estrutura-visual-e-a-escolha-do-delimitador-pipe-)
  - [2.2 O Bloco de Identificação e Metadados ([METADATA])](#22-o-bloco-de-identificação-e-metadados-metadata)
  - [2.3 O Contrato de Campos e Colunas ([SCHEMA])](#23-o-contrato-de-campos-e-colunas-schema)
  - [2.4 O Bloco de Registros Compactos ([DATA])](#24-o-bloco-de-registros-compactos-data)
  - [2.5 Tratamento de Caracteres Especiais, Quebras de Linha e Higienização](#25-tratamento-de-caracteres-especiais-quebras-de-linha-e-higienização)
- [Capítulo 3: A Engenharia dos Conversores](#capítulo-3-a-engenharia-dos-conversores)
  - [3.1 O Que Existe por Trás de um Arquivo .xlsx (ZIP e XMLs)](#31-o-que-existe-por-trás-de-um-arquivo-xlsx-zip-e-xmls)
  - [3.2 O Segredo das Tabelas de Strings Compartilhadas (sharedStrings.xml)](#32-o-segredo-das-tabelas-de-strings-compartilhadas-sharedstringsxml)
  - [3.3 A Famosa Data Serial do Excel (e o Bug Histórico de 1900)](#33-a-famosa-data-serial-do-excel-e-o-bug-histórico-de-1900)
  - [3.4 O Motor Nativo em Rust: Por Que 6 Milissegundos Mudam o Jogo](#34-o-motor-nativo-em-rust-por-que-6-milissegundos-mudam-o-jogo)
  - [3.5 O Motor Portátil em Python: Zero Dependências e Acessibilidade](#35-o-motor-portátil-em-python-zero-dependências-e-acessibilidade)
- [Capítulo 4: Guia de Uso, Compilação e Distribuição Multiplataforma](#capítulo-4-guia-de-uso-compilação-e-distribuição-multiplataforma)
  - [4.1 Rodando no Linux (Binário Nativo e Linha de Comando)](#41-rodando-no-linux-binário-nativo-e-linha-de-comando)
  - [4.2 Rodando no Windows (Binário .exe, Scripts .bat e .ps1)](#42-rodando-no-windows-binário-exe-scripts-bat-e-ps1)
  - [4.3 Dicionário Completo de Comandos e Parâmetros da CLI](#43-dicionário-completo-de-comandos-e-parâmetros-da-cli)
  - [4.4 O Pipeline Automatizado de Releases no GitHub Actions](#44-o-pipeline-automatizado-de-releases-no-github-actions)
- [Capítulo 5: Integração com o Aplicativo cabeceira-pwa1](#capítulo-5-integração-com-o-aplicativo-cabeceira-pwa1)
  - [5.1 O Parser Leve em TypeScript (imcaParser.ts)](#51-o-parser-leve-em-typescript-imcaparserts)
  - [5.2 Mapeamento Transparente para o EventoCanonico](#52-mapeamento-transparente-para-o-eventocanonico)
  - [5.3 Filtragem Mensal e Renderização Instantânea](#53-filtragem-mensal-e-renderização-instantânea)
- [Capítulo 6: Segurança, Governança e Ciclo de Vida](#capítulo-6-segurança-governança-e-ciclo-de-vida)
  - [6.1 Blindagem de Repositório com o .gitignore](#61-blindagem-de-repositório-com-o-gitignore)
  - [6.2 Política de Versionamento Semântico Incremental](#62-política-de-versionamento-semântico-incremental)
  - [6.3 Conclusão e Próximos Passos](#63-conclusão-e-próximos-passos)
- [Capítulo 7: A Interface Gráfica Desktop e o Empacotamento Multiplataforma (.exe e .AppImage)](#capítulo-7-a-interface-gráfica-desktop-e-o-empacotamento-multiplataforma-exe-e-appimage)
  - [7.1 A Filosofia da Interface de Tela Única (Simplicidade e Clareza)](#71-a-filosofia-da-interface-de-tela-única-simplicidade-e-clareza)
  - [7.2 Anatomia da Janela: O Seletor de Arquivos e os Controles de Configuração](#72-anatomia-da-janela-o-seletor-de-arquivos-e-os-controles-de-configuração)
  - [7.3 O Card do Arquivo Transformado: Métricas de Economia e Prévia](#73-o-card-do-arquivo-transformado-métricas-de-economia-e-prévia)
  - [7.4 Como o Pacote Linux AppImage (.image) Funciona por Dentro (AppDir e AppRun)](#74-como-o-pacote-linux-appimage-image-funciona-por-dentro-appdir-e-apprun)
  - [7.5 Como o Executável Windows (.exe) é Construído Sem Dependências](#75-como-o-executável-windows-exe-é-construído-sem-dependências)
- [Capítulo 8: A Identidade Visual do Sistema — A Chama e o Calendário](#capítulo-8-a-identidade-visual-do-sistema--a-chama-e-o-calendário)
  - [8.1 O Significado Semiótico: Da Planilha Bruta ao Fogo do Ministério](#81-o-significado-semiótico-da-planilha-bruta-ao-fogo-do-ministério)
  - [8.2 O Processo Criativo e a Unificação Black & Orange (#171717 e #F25623)](#82-o-processo-criativo-e-a-unificação-black--orange-171717-e-f25623)
  - [8.3 A Engenharia de Geração de Ícones Multiplataforma (PNG 512, ICO Multi-Res e SVG)](#83-a-engenharia-de-geração-de-ícones-multiplataforma-png-512-ico-multi-res-e-svg)
  - [8.4 A Integração nos Pacotes de Distribuição (.AppImage, .image e .exe)](#84-a-integração-nos-pacotes-de-distribuição-appimage-image-e-exe)
  - [8.5 A Linha de Montagem em Nuvem do GitHub Actions e Publicação Automatizada de Releases](#85-a-linha-de-montagem-em-nuvem-do-github-actions-e-publicação-automatizada-de-releases)

---

## Prólogo: A Missão e o Ponto de Partida

No contexto da gestão eclesiástica e de escalas ministeriais, as secretarias e pastorais frequentemente planejam o ano inteiro utilizando planilhas eletrônicas do Microsoft Excel (`.xlsx`). O arquivo [Calendario 2026.xlsx](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/Calendario%202026.xlsx) é um exemplo real dessa dinâmica: ele centraliza cultos, campanhas, vigílias e eventos de discipulado ao longo de todo o ano de 2026.

Entretanto, quando essa informação precisa chegar às mãos de centenas de voluntários em um aplicativo móvel progressivo (o **cabeceira-pwa1**), as planilhas do Excel tornam-se um grande obstáculo técnico:
1. São arquivos binários compactados com múltiplos esquemas XML internos.
2. Exigem bibliotecas de código gigantescas (como `SheetJS/xlsx`, com mais de 1,2 megabytes de peso) para conseguir abrir uma planilha no celular do usuário.
3. Consomem preciosos megabytes de memória RAM dos smartphones mais humildes, causando travamentos, lentidão de rolagem e drenagem excessiva de bateria.

O projeto **Convert_IMCA** nasceu para romper essa barreira: transformamos a complexidade de uma planilha pesada em um formato textual novo, compacto, seguro e transparente, denominado **`.IMCA`**, que pode ser processado no celular em apenas **1 único milissegundo** com **zero bibliotecas externas**.

---

## Capítulo 1: O Surgimento do Formato .IMCA

### 1.1 O Peso Oculto das Planilhas Modernas (.xlsx)
Para o usuário comum, uma planilha de Excel parece apenas uma grade inofensiva de linhas e colunas. No entanto, tecnicamente, qualquer arquivo com extensão `.xlsx` é um pacote comprimido em formato ZIP que oculta em seu interior dezenas de arquivos auxiliares:
- `xl/workbook.xml` (definição de pastas e relações)
- `xl/sharedStrings.xml` (dicionário com cada texto digitado)
- `xl/styles.xml` (tabelas de fontes, cores, bordas e alinhamentos)
- `xl/worksheets/sheet1.xml`, `sheet2.xml`, etc. (árvores XML para cada célula)

Para ler um único evento como *"Reveillon cabeceira"*, um programa precisa descompactar o ZIP, carregar centenas de tags XML na memória do computador, consultar o índice numérico da string e só então recuperar o texto original.

### 1.2 A Realidade dos Dispositivos Móveis e o Desafio do PWA
A aplicação consumidora dos dados é o **cabeceira-pwa1**, um aplicativo web progressivo (PWA) construído com React Native Web e TypeScript, planejado para funcionar de forma instantânea em conexões móveis (3G, 4G e redes comunitárias).

Se o PWA tivesse que baixar uma planilha do Excel de quase 100 KB e carregar uma biblioteca pesada de 1.2 MB para processá-la:
- O carregamento inicial da página ficaria cerca de 3 a 5 segundos mais lento.
- O aparelho do usuário sofreria engasgos na renderização visual do calendário.
- O consumo de franquia de dados móveis seria multiplicado desnecessariamente.

### 1.3 O Que é o Formato .IMCA e Por Que Ele Foi Criado
O formato **`.IMCA`** (*Interchange Model for Calendar & Agenda*) foi projetado sob três pilares inegociáveis:
1. **Ultra-Leveza:** Cada caractere no arquivo tem um propósito vital. Não há tags XML redundantes, formatações de cor ou estruturas vazias.
2. **Legibilidade Humana:** Qualquer líder ou pastor pode abrir o arquivo `.imca` no Bloco de Notas ou no VS Code e ler perfeitamente todas as datas e eventos.
3. **Decodificação Imediata:** O navegador ou aplicativo não precisa de nenhuma biblioteca especial. Ele utiliza o método nativo de quebra de texto (`split`) presente no JavaScript desde a sua primeira versão.

### 1.4 Benchmark Real: O Salto de 96.8 KB para 25.7 KB (e 4.4 KB Comprimido)
No teste realizado diretamente com a base real do [Calendario 2026.xlsx](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/Calendario%202026.xlsx), os resultados empíricos demonstraram a eficácia da solução:

```mermaid
pie title Tamanho Comparativo dos Dados (Bytes)
    "Planilha Excel (.xlsx)" : 99091
    "Formato IMCA Puro" : 26406
    "Formato IMCA Comprimido (GZIP)" : 4508
```

- **Planilha Excel (.xlsx):** `96.77 KB` (99.091 bytes).
- **Formato .IMCA Texto Puro:** `25.79 KB` (26.406 bytes) $\rightarrow$ **73.4% de economia direta de espaço**.
- **Formato .IMCA Comprimido (Web Gzip):** `4.40 KB` (4.508 bytes) $\rightarrow$ **95.5% de redução de tráfego de rede**.
- **Tempo de Processamento no PWA:** Reduzido de mais de 200 ms para apenas **1.11 ms**.

---

## Capítulo 2: A Anatomia Técnica do Formato .IMCA

### 2.1 A Estrutura Visual e a Escolha do Delimitador Pipe (|)
Para separar as informações de cada coluna, poderíamos utilizar vírgulas (CSV) ou tabulações (TSV). No entanto:
- Vírgulas frequentemente aparecem no nome dos eventos (ex: *"Vigília, Clamor e Louvor"*).
- Tabulações são invisíveis e fáceis de serem corrompidas por editores de texto.

Por isso, o formato `.IMCA` adota o caractere de barra vertical ou **Pipe (`|`)** como delimitador oficial. O pipe é visualmente nítido, raramente utilizado em nomes de cultos e possui custo computacional mínimo para identificação.

### 2.2 O Bloco de Identificação e Metadados ([METADATA])
Todo arquivo `.imca` se inicia com um cabeçalho informativo protegido por comentários (`#`), seguido pela seção declarativa `[METADATA]`:

```text
# =============================================================================
# FORMATO DE DADOS IMCA v1.0 - AGENDA E EVENTOS OTIMIZADOS
# APLICATIVO ALVO: cabeceira-pwa1
# PADRÃO: DELIMITADO POR PIPE (|) COM CABEÇALHO E METADADOS INLINE
# =============================================================================

[METADATA]
versao=1.0
ano=2026
total_registros=318
gerado_em=2026-10-05T14:40:43.910133288+00:00
arquivo_origem=Calendario 2026.xlsx
aplicativo_alvo=cabeceira-pwa1
delimitador=|
```

Essas chaves permitem ao sistema receptor conferir a integridade do arquivo antes de renderizar qualquer informação na tela (por exemplo, certificando-se de que o total de registros lidos confere com o valor declarado em `total_registros`).

### 2.3 O Contrato de Campos e Colunas ([SCHEMA])
A seção `[SCHEMA]` define rigidamente a ordem e o significado de cada coluna de dados:

```text
[SCHEMA]
id|data|dia_semana|horario|nome|local|solicitante|status
```

1. **`id`**: Identificador canônico único no formato `ev_YYYYMMDD_seq` (ex: `ev_20260102_002`).
2. **`data`**: Data do evento no formato padrão ISO 8601 (`AAAA-MM-DD`).
3. **`dia_semana`**: Dia da semana correspondente (ex: `Friday` ou `Sexta-feira`).
4. **`horario`**: Horário planejado para o início da atividade (ex: `19:30`).
5. **`nome`**: Nome do evento limpo e sem caracteres de quebra.
6. **`local`**: Local de realização (ex: `Templo`, `Anexo`, `Salão Social`).
7. **`solicitante`**: Ministério ou secretaria solicitante (ex: `Secretaria do discipulado`).
8. **`status`**: Ciclo de vida da escala (`publicado`, `planejamento`, `concluido`).

### 2.4 O Bloco de Registros Compactos ([DATA])
Sob a etiqueta `[DATA]`, cada linha representa um evento completo. Se um campo não tiver valor (por exemplo, quando o local ou solicitante não foram preenchidos), ele permanece vazio entre os pipes, ocupando apenas um único byte:

```text
[DATA]
ev_20251231_001|2025-12-31|Wednesday|19:30|Reveillon cabeceira|||publicado
ev_20260102_002|2026-01-02|Friday|19:30|12 Dias de Clamor para 12 Meses de Milagres|Templo|Secretaria do discipulado|publicado
ev_20260103_003|2026-01-03|Saturday|19:30|12 Dias de Clamor para 12 Meses de Milagres|Templo|Secretaria do discipulado|publicado
```

### 2.5 Tratamento de Caracteres Especiais, Quebras de Linha e Higienização
O conversor aplica regras automáticas de saneamento:
- Caracteres pipe acidentais no texto original são convertidos para traços (`-`).
- Quebras de linha (`\n`, `\r`) dentro de células do Excel são substituídas por espaços simples, impedindo que uma linha seja quebrada indevidamente no meio de um registro.
- Espaços em branco redundantes nas extremidades são eliminados (`trim`).

---

## Capítulo 3: A Engenharia dos Conversores

Para proporcionar versatilidade absoluta e performance de ponta, o ecossistema Convert_IMCA foi construído com dois motores sincronizados:

```mermaid
flowchart TD
    subgraph Entrada
        XLSX["Calendario 2026.xlsx\n(Planilha Bruta)"]
    end

    subgraph Motores
        RUST["Motor Nativo Rust\n(6.36 ms | Binário 1.5MB)"]
        PY["Motor Portátil Python\n(Zero Dependências Externas)"]
    end

    subgraph Saída
        IMCA["Calendario 2026.imca\n(Formato Otimizado)"]
    end

    XLSX --> RUST
    XLSX --> PY
    RUST --> IMCA
    PY --> IMCA
```

### 3.1 O Que Existe por Trás de um Arquivo .xlsx (ZIP e XMLs)
Quando abrimos uma planilha pelo conversor, o primeiro passo é tratar o arquivo binário como um contêiner ZIP. O leitor descompacta em memória apenas dois arquivos vitais:
1. `xl/sharedStrings.xml`: Onde residem todos os textos.
2. `xl/worksheets/sheetX.xml`: Onde residem as posições das células (`c r="B2"`, `c r="D2"`, etc.).

### 3.2 O Segredo das Tabelas de Strings Compartilhadas (sharedStrings.xml)
No formato OpenXML do Excel, palavras repetidas não são salvas várias vezes. Se a palavra *"Templo"* aparecer 200 vezes na planilha, ela é guardada apenas uma vez na tabela `sharedStrings.xml` com o índice `42`. As células da planilha guardam apenas o número `42`.
Nossos conversores realizam a resolução reversa instantânea, mapeando os índices numéricos de volta aos seus nomes reais.

### 3.3 A Famosa Data Serial do Excel (e o Bug Histórico de 1900)
Um dos maiores desafios no tratamento de planilhas é a representação de datas. O Excel não salva `"2026-01-02"` como texto; ele salva o número decimal `46024.0`.

Esse número representa a quantidade de dias transcorridos desde o dia **30 de dezembro de 1899**.
> **Curiosidade Histórica:** O Excel herdou propositalmente um erro de cálculo do clássico software Lotus 1-2-3, que considerava erroneamente o ano de 1900 como bissexto. Por isso, a data base do cálculo deve ser fixada em `1899-12-30` para que qualquer dia a partir de 1900 fique perfeitamente alinhado.

O conversor implementa a conversão precisa dessa data serial para o padrão internacional `AAAA-MM-DD`.

### 3.4 O Motor Nativo em Rust: Por Que 6 Milissegundos Mudam o Jogo
O motor principal foi desenvolvido na linguagem **Rust** (arquivos em `src/`), reconhecida mundialmente pela velocidade incomparável e segurança de memória:
- **Tempo de Execução:** Converte todas as 545 linhas e 318 eventos em míseros **6.36 milissegundos**.
- **Consumo de Memória:** Menos de 5 MB de memória durante todo o processo.
- **Binário Estático:** O executável compilado (`target/release/convert_imca`) tem apenas **1.5 MB** e não precisa de bibliotecas dinâmicas ou dependências externas instaladas no sistema operacional.

### 3.5 O Motor Portátil em Python: Zero Dependências e Acessibilidade
Para os cenários onde o operador está em uma máquina sem o compilador Rust, fornecemos o motor complementar [convert_imca.py](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/convert_imca.py).
Ele utiliza exclusivamente as bibliotecas embutidas do próprio interpretador Python (`zipfile`, `xml.etree.ElementTree`, `datetime`, `argparse`), sem exigir nenhum comando `pip install`. Ele roda em Windows, Linux e macOS com total paridade de saída.

---

## Capítulo 4: Guia de Uso, Compilação e Distribuição Multiplataforma

### 4.1 Rodando no Linux (Binário Nativo e Linha de Comando)

Para compilar o binário em modo release otimizado:
```bash
cargo build --release
```
O executável final estará disponível em:
```text
target/release/convert_imca
```

Para converter uma planilha:
```bash
# Conversão básica com detecção automática:
./target/release/convert_imca "Calendario 2026.xlsx"

# Conversão com prévia detalhada no terminal:
./target/release/convert_imca "Calendario 2026.xlsx" --preview

# Conversão traduzindo os dias para Português:
./target/release/convert_imca "Calendario 2026.xlsx" --translate-days --preview
```

### 4.2 Rodando no Windows (Binário .exe, Scripts .bat e .ps1)
No Windows, o usuário pode escolher entre três formas simples de execução:

1. **Via Executável Direto:**
   ```cmd
   convert_imca.exe "Calendario 2026.xlsx" --preview
   ```

2. **Via Inicializador Batch ([convert_imca.bat](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/convert_imca.bat)):**
   Basta arrastar a planilha por cima do arquivo `convert_imca.bat` ou executar no Prompt de Comando:
   ```cmd
   convert_imca.bat "Calendario 2026.xlsx"
   ```

3. **Via PowerShell ([convert_imca.ps1](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/convert_imca.ps1)):**
   ```powershell
   .\convert_imca.ps1 "Calendario 2026.xlsx" -p
   ```

### 4.3 Dicionário Completo de Comandos e Parâmetros da CLI

| Parâmetro | Tipo | Padrão | Descrição |
| :--- | :--- | :--- | :--- |
| `<ARQUIVO_XLSX>` | Posicional | *Obrigatório* | Caminho da planilha Excel de entrada. |
| `-o, --output` | String | `<mesmo_nome>.imca` | Caminho do arquivo de saída desejado. |
| `-s, --sheet` | String | *Auto* | Força o processamento de uma aba específica (ex: `--sheet "Calendario"`). |
| `-y, --year` | Inteiro | *Auto (2026)* | Define o ano base para datas incompletas. |
| `-t, --default-time` | String | `"19:30"` | Horário padrão atribuído aos cultos/eventos. |
| `--translate-days` | Flag | `false` | Traduz os dias da semana de inglês para português. |
| `-p, --preview` | Flag | `false` | Mostra no terminal um resumo dos primeiros 5 eventos. |
| `--validate` | Booleano | `true` | Realiza a prova real e validação do arquivo recém-gerado. |

### 4.4 O Pipeline Automatizado de Releases no GitHub Actions
Criamos o fluxo automatizado [.github/workflows/release.yml](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/.github/workflows/release.yml).
Sempre que uma tag de versão for publicada (ex: `v1.0.0`), os servidores em nuvem do GitHub disparam uma matriz de compilação:
- Compilam o executável nativo Linux (`convert_imca-linux-x86_64`) em uma máquina Ubuntu.
- Compilam o executável nativo Windows (`convert_imca-windows-x86_64.exe`) em uma máquina Windows Server.
- Anexam os binários já prontos para download na área de *Releases* do repositório, permitindo que qualquer membro da equipe baixe e use sem compilar nada.

---

## Capítulo 5: Integração com o Aplicativo cabeceira-pwa1

### 5.1 O Parser Leve em TypeScript (imcaParser.ts)
No repositório do PWA, a importação do arquivo `.imca` é intermediada pelo módulo [imcaParser.ts](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/imcaParser.ts).

O leitor possui menos de 100 linhas de código e não adiciona nenhum peso ao bundle do aplicativo:
```typescript
import { parseIMCA, EventoCanonico } from './services/imcaParser';

// Exemplo: lendo o arquivo transferido pela rede ou pelo Firebase Storage
const resposta = await fetch('/calendarios/Calendario_2026.imca');
const textoImca = await resposta.text();

// Parse instantâneo em 1.11 milissegundos
const { metadata, eventos } = parseIMCA(textoImca);
console.log(`Carregados ${eventos.length} eventos do ano ${metadata.ano}`);
```

### 5.2 Mapeamento Transparente para o EventoCanonico
Cada linha do bloco `[DATA]` é convertida diretamente na interface `EventoCanonico` utilizada pelo Firebase Realtime Database e pelo Contexto global do aplicativo:

```typescript
export interface EventoCanonico {
  id: string;              // "ev_20260102_002"
  nome: string;            // "12 Dias de Clamor para 12 Meses de Milagres"
  data: string;            // "2026-01-02"
  horario: string;         // "19:30"
  local: string;           // "Templo"
  solicitante?: string;    // "Secretaria do discipulado"
  status: StatusEvento;    // "publicado"
  minisRequisitados?: {};  // Pronto para receber escalas
  escalas?: {};
}
```

### 5.3 Filtragem Mensal e Renderização Instantânea
Para abastecer a tela de calendário (`calendar.tsx`), a função utilitária `filtrarEventosPorAnoMes` permite extrair apenas os cultos do mês selecionado pelo usuário sem atraso:
```typescript
import { filtrarEventosPorAnoMes } from './services/imcaParser';

// Filtra instantaneamente apenas os eventos de Janeiro de 2026:
const eventosJaneiro = filtrarEventosPorAnoMes(eventos, '2026-01');
```

---

## Capítulo 6: Segurança, Governança e Ciclo de Vida

### 6.1 Blindagem de Repositório com o .gitignore
Para evitar que binários de sistema operacional, arquivos temporários de compilação ou dados sensíveis sejam enviados ao repositório público, o arquivo [.gitignore](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/.gitignore) foi configurado para bloquear:
- Pastas de compilação pesadas (`/target/`, `/dist/`, `/build/`).
- Binários executáveis (`*.exe`, `*.bin`, `*.so`, `*.dll`).
- Caches de linguagem (`__pycache__/`, `*.pyc`, `.venv/`).
- Arquivos temporários de bloqueio de planilhas (`~$*.xlsx`).
- Chaves privadas e arquivos de variáveis de ambiente (`.env*`, `*.pem`, `*.key`).

### 6.2 Política de Versionamento Semântico Incremental
O projeto adota a convenção de Versionamento Semântico (`MAJOR.MINOR.PATCH`):
- **Versão `v1.0.0`:**
  - Criação da especificação formal do padrão `.IMCA v1.0`.
  - Implementação do motor nativo em Rust (compilação estática de 1.5MB).
  - Implementação do motor portátil em Python (zero dependências).
  - Módulo TypeScript de alta velocidade para o `cabeceira-pwa1`.
  - Scripts de automação para Windows e Linux.
  - Documentação completa em formato de livro.
- **Versão `v1.1.0`:**
  - Criação da Interface Gráfica Desktop de tela única.
  - Suporte a empacotamento autônomo Linux AppImage e Windows `.exe`.
- **Versão `v1.2.0`:**
  - Reformulação visual com fluxo de 3 telas lineares e introdução da animação da chama viva com *breathing glow*.
- **Versão `v1.4.0`:**
  - Introdução do design geométrico Black & Orange com elementos menos arredondados (cantos de 8px a 12px).
- **Versão `v1.5.0` (Versão Atual):**
  - **Correção da Rolagem da Tabela de Prévia:** Cabeçalho fixo (*sticky*) 100% opaco (`background: #1C1C1C; z-index: 20`) com `border-collapse: separate; border-spacing: 0;` e ajuste do container de rolagem, eliminando qualquer vazamento ou sobreposição de linhas sob o cabeçalho.
  - **Unificação de Cores dos Botões (#F25623):** O botão de busca da Tela 1 e o botão de arquivo gerado da Tela 3 adotam rigorosamente o mesmo tom de laranja sólido `#F25623`, sem lavagem de cor.
  - **Eliminação de Efeitos Luminescentes Excessivos:** Remoção da luz ambiente difusa (*ambient glow*), eliminação de halos neon em repouso e atenuação do efeito de respiração da chama para um brilho orgânico e sutil, reservando destaques luminosos exclusivamente para a passagem do mouse (*hover*).
  - Atualização dos pacotes autônomos e motor Rust para a versão 1.5.0.

### 6.3 Conclusão e Próximos Passos
O conversor e o formato `.IMCA` consolidam uma ponte de altíssima eficiência entre as secretarias da igreja (que trabalham com planilhas Excel) e os voluntários na ponta final (que utilizam o aplicativo móvel `cabeceira-pwa1`). A economia de mais de 73% de armazenamento e a velocidade de leitura em milissegundos garantem uma experiência de uso fluida, estável e moderna.

---

## Capítulo 7: A Interface Gráfica Desktop e o Empacotamento Multiplataforma (.exe e .AppImage)

### 7.1 A Filosofia da Interface de Três Telas (Entrada, Carregamento e Saída)
A experiência do usuário no Convert_IMCA foi refinada na versão **`v1.5.0`** com rigor visual e usabilidade ergonômica:
- **Zero Poluição:** Telas limpas, sem elementos conflitantes, borrões ou halos luminosos desnecessários.
- **Narrativa Clara:** Cada tela cumpre uma missão única:
  1. *Entrada*: O usuário localiza e entrega a planilha em um botão hero retangular acolhedor.
  2. *Carregamento*: A chama viva pulsa suavemente enquanto os algoritmos processam os dados em milissegundos.
  3. *Conclusão*: O usuário escolhe entre inspecionar os dados gerados (*mostrar prévia*) com rolagem perfeita ou revelar o arquivo no computador (*mostrar arquivo gerado*).

```mermaid
flowchart TD
    subgraph Tela1 ["Tela 1: Entrada (Botão Hero Geométrico)"]
        Tag["| MOTOR DE CALENDÁRIO • IMCA"]
        T1["Convert IMCA\n(Tipografia Branca Pura #FFFFFF)"]
        Hero["🟧 BUSCAR ARQUIVO XLSX\n(Laranja Sólido #F25623 + Drag&Drop)"]
        Tag --> T1 --> Hero
    end

    subgraph Tela2 ["Tela 2: Carregamento"]
        F2["🔥 Chama Viva Animada\n(Breathing Glow Suave e Natural)"]
        S2["Status Dinâmico:\n'Lendo planilha...' -> 'Extraindo 318 eventos...'"]
        F2 --> S2
    end

    subgraph Tela3 ["Tela 3: Conclusão (Cartões Geométricos)"]
        R1["mostrar prévia\n(Tabela com Cabeçalho Sticky Opaco)"]
        R2["mostrar arquivo gerado\n(Laranja Sólido #F25623)"]
        R3["↺ Converter outro arquivo"]
        R1 --- R2 --- R3
    end

    Hero --> Tela2
    Tela2 --> Tela3
    R3 --> Tela1
```

### 7.2 Tela 1: O Começo Geométrico e o Botão Único Unificado
A versão `v1.5.0` aprimora os contrastes e remove efeitos difusos:
- **Fundo Sóbrio e Limpo:** O fundo preto fosco profundo (`#171717`) agora se apresenta livre de manchas luminosas (*ambient glow* desativado), garantindo sobriedade e foco nos dados.
- **Botão Único com Cor Unificada (#F25623):** O botão de busca adota o mesmo laranja sólido oficial do botão final, com sombra discreta em repouso (`box-shadow: 0 2px 6px rgba(0, 0, 0, 0.4);`) e leve realce apenas durante o *hover*.
- **Geometria de Cantos:** Mantém cantos sutis de 12px, reforçando a identidade arquitetônica.

### 7.3 Tela 2: A Chama Viva com Iluminação Equilibrada
O símbolo oficial da chama mantém seu movimento orgânico de respiração (*breathing glow*), porém com filtros de sombra calibrados para níveis naturais (`drop-shadow: 0 4px 12px rgba(242, 86, 35, 0.45)`), sem o excesso de saturação visual das versões iniciais.

### 7.4 Tela 3: Conclusão e Rolagem Perfeita na Tabela de Prévia
O modal de prévia foi corrigido para resolver a sobreposição de linhas durante a rolagem:
- **Cabeçalho Fixo 100% Opaco:** As células de cabeçalho (`<th>`) possuem fundo sólido (`#1C1C1C`), camada `z-index: 20` e borda inferior nítida de 2px.
- **Isolamento de Células:** O uso de `border-collapse: separate; border-spacing: 0;` e a remoção de espaçamentos verticais no container de rolagem asseguram que as linhas da tabela deslizem suavemente por baixo do cabeçalho sem qualquer vazamento de texto.
- **A Chama:** O símbolo oficial da chama extraído em alta definição ([`assets/flame.png`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/assets/flame.png)) e embutido diretamente no HTML como base64, garantindo autonomia total sem necessidade de arquivos externos.
- **Animação Breathing Glow:** Um efeito de respiração orgânica em CSS3 que alterna a escala da chama entre 96% e 105% acompanhada por sombras dinâmicas de luz vermelha e âmbar.
- **Micro-Etapas:** O texto abaixo da chama informa os passos do algoritmo:
  - *"Lendo planilha Excel..."*
  - *"Extraindo cultos e escalas..."*
  - *"Normalizando datas para padrão ISO..."*
  - *"Otimizando para formato .IMCA..."*

### 7.4 Tela 3: As Duas Ações Principais ("mostrar prévia" e "mostrar arquivo gerado")
Finalizada a conversão, a interface exibe as duas opções de forma direta e limpa:

1. **Pílula `mostrar prévia`:**
   - Abre uma elegante gaveta modal com o resumo completo da conversão:
     - **Eventos:** 318 cultos e escalas validados.
     - **Original:** 96.8 KB.
     - **Formato IMCA:** 25.8 KB (**-73.4%** de redução direta e **-95.5%** comprimido na web).
     - **Tabela:** Grade com as colunas *Data (ISO)*, *Dia*, *Hora*, *Nome do Evento* e *Local*.
2. **Pílula `mostrar arquivo gerado`:**
   - Dispara o gerenciador de arquivos nativo (Nautilus no Linux ou Explorer no Windows) com a pasta aberta e o arquivo `.imca` selecionado.
   - Exibe um toast flutuante verde confirmando que o caminho foi copiado.
3. **Link `↺ Converter outro arquivo`:**
   - Reseta a interface suavemente e retorna para a Tela 1 para uma nova conversão.

### 7.5 O Motor Gráfico PyWebView, a Ponte Bidirecional e o Seletor Resiliente
A interface é alimentada por um motor leve e autônomo de **PyWebView**:
- **Renderização Nativa:** No Linux utiliza o WebKitGTK 2.50 pré-instalado; no Windows utiliza o Microsoft Edge WebView2 nativo do sistema operacional.
- **Ponte Python/JS:** Quando o usuário clica em qualquer ação no HTML/CSS, uma chamada assíncrona (`window.pywebview.api.select_file` ou `convert_file`) é enviada diretamente ao interpretador Python, que interage com o sistema operacional e devolve o resultado em JSON.
- **Arquitetura de Seletor de Arquivos com Triplo Fallback:**
  Para garantir que a janela nativa de busca do sistema operacional sempre abra com 100% de confiabilidade, mesmo diante de peculiaridades entre Wayland e X11 no Linux:
  1. *Camada 1 (Nativa do PyWebView):* Dispara `window.create_file_dialog` com o enum `FileDialog.OPEN`.
  2. *Camada 2 (Fallback Nativo GNOME/Zenity):* Se a chamada da camada 1 encontrar qualquer restrição de thread no GTK, o motor invoca o utilitário nativo `/usr/bin/zenity --file-selection`, padrão oficial do ecossistema Linux/Ubuntu.
  3. *Camada 3 (Fallback Tkinter):* Caso nenhum diálogo anterior responda, o módulo padrão `tkinter.filedialog` é disparado como rede de segurança final.
  4. *Camada Web:* No JavaScript, a higienização de caminhos multiplataforma utiliza o padrão seguro `.replace(/\\/g, '/').split('/').pop()`, imune a ambiguidades de escape em expressões regulares.

### 7.6 Empacotamento Linux AppImage (.image) e Executável Windows (.exe)
Os artefatos de distribuição foram atualizados para incorporar toda a nova interface:
- **Linux:** O script [`build_appimage.sh`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/build_appimage.sh) empacota o executável e a interface em um único arquivo `Convert_IMCA-x86_64.AppImage` (e atalho `Convert_IMCA.image`), que roda com duplo clique.
- **Windows:** Os scripts [`build_windows_gui.bat`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/build_windows_gui.bat) e [`build_windows_gui.ps1`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/build_windows_gui.ps1) compilam o executável autônomo `Convert_IMCA_GUI.exe` com o ícone embutido e sem janela de console.

---

## Capítulo 8: A Identidade Visual do Sistema — A Chama e o Calendário

### 8.1 O Significado Semiótico: Da Planilha Bruta ao Fogo do Ministério
Todo software que aspira à excelência de uso necessita de uma identidade visual que dialogue diretamente com a sua razão de existir. O **Convert_IMCA** não é apenas um utilitário técnico de conversão binária; ele é a ponte que transporta o planejamento anual da igreja — cultos, santas ceias, vigílias e escalas de voluntários — para a palma da mão da congregação no aplicativo **cabeceira-pwa1**.

A semiótica do novo ícone (`v1.6.0`) funde dois pilares fundamentais:
1. **O Calendário (A Ordem e o Planejamento):** Representado por uma folha de calendário arquitetônica e geométrica, com suas encadernações superiores e divisão de datas. Simboliza a dedicação das secretarias e pastorais em organizar o ano de forma diligente e estruturada.
2. **A Chama Viva (O Propósito e a Unção):** Posicionada com destaque sobre o corpo do calendário, a chama extraída do logotipo oficial do ministério transmite a mensagem central da fé: os eventos e escalas não são compromissos burocráticos frios em uma folha de Excel, mas momentos dedicados à presença viva do Espírito Santo na igreja.

```mermaid
graph TD
    subgraph "Identidade Visual Convert_IMCA (v1.6.0)"
        Planilha["Planilha Excel Bruta\n(Dados Frios)"] --> Fusao["Fusão de Conceitos"]
        Calendario["Folha de Calendário\n(Ordem e Planejamento)"] --> Fusao
        Chama["Chama Oficial do Ministério\n(Fogo e Vida do Espírito)"] --> Fusao
        Fusao --> IconeFinal["Ícone Oficial do Sistema\n(Black #171717 & Orange #F25623)"]
    end
```

### 8.2 O Processo Criativo e a Unificação Black & Orange (#171717 e #F25623)
Durante o desenvolvimento da versão `1.6.0`, quatro propostas conceituais foram geradas e submetidas à análise direta do usuário:
- *Proposta 1 (Digital Glow):* Um calendário em perspectiva 3D estilizada com grades luminosas e fogo superior.
- *Proposta 2 (Swiss Grid):* Uma abordagem inspirada no design suíço internacionalista, com foco em tipografia técnica e fogo sutil.
- *Proposta 3 (Minimalist Flat):* Uma silhueta vetorial monocromática plana.
- *Proposta 4 (Oficial Integrada - Escolhida e Aprovada):* A composição geométrica perfeita que alinha o fundo escuro grafite institucional (`#171717`), o cabeçalho superior na cor sólida oficial (`#F25623`) e a própria silhueta da chama ministerial em alta definição.

Essa escolha unifica todo o ecossistema visual do projeto: a tela de abertura, a animação de carregamento, os botões de ação e o ícone do sistema operacional operam sob a exata mesma linguagem visual.

### 8.3 A Engenharia de Geração de Ícones Multiplataforma (PNG 512, ICO Multi-Res e SVG)
Em sistemas operacionais modernos, um ícone de aplicativo não pode ser uma imagem comum redimensionada arbitrariamente pelo sistema. Diferentes ambientes exigem tratamentos técnicos específicos para garantir nitidez cristalina:

```mermaid
graph LR
    Master["Icon Master\n(512x512 RGBA)"] --> PNG["assets/icon.png\n(Linux / WebKit)"]
    Master --> ICO["assets/icon.ico\n(Multi-Res Windows)"]
    ICO --> W1["16x16 (Taskbar mini)"]
    ICO --> W2["32x32 (Taskbar padrão)"]
    ICO --> W3["48x48 (Desktop médio)"]
    ICO --> W4["64x64 (Menu Iniciar)"]
    ICO --> W5["128x128 (HiDPI)"]
    ICO --> W6["256x256 (4K Ultra-HD)"]
```

1. **Formato Mestre PNG (512x512 com Canal Alpha):**
   - Salvo em [`assets/icon.png`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/assets/icon.png), renderizado com reamostragem bilinear de alta precisão (*Lanczos*), preservando transparência nos cantos chanfrados e fidelidade cromática no espectro sRGB.
2. **Formato Windows Multi-Resolution (.ICO):**
   - O arquivo [`assets/icon.ico`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/assets/icon.ico) é uma cápsula binária que encapsula 6 versões distintas da mesma imagem nos tamanhos: `16x16`, `32x32`, `48x48`, `64x64`, `128x128` e `256x256` pixels.
   - Isso evita o erro comum de interpolação onde o Windows estica uma imagem pequena ou esmaga uma imagem grande, garantindo que o ícone na barra de tarefas ou na visualização em lista fique perfeitamente nítido sem serrilhamento.

### 8.4 A Integração nos Pacotes de Distribuição (.AppImage, .image e .exe)
O novo ícone foi integrado em todas as etapas da cadeia de montagem:
- **No Linux (Desktop Entry e AppDir):**
  - O script [`build_appimage.sh`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/build_appimage.sh) injeta o ícone em três locais do pacote montado: na raiz como `.DirIcon` (para os gerenciadores de arquivos do GNOME e KDE), em `convert_imca.png` na raiz do AppImage, e em `usr/share/icons/hicolor/256x256/apps/convert_imca.png`.
  - O arquivo `convert_imca.desktop` referencia `Icon=convert_imca`, assegurando que o sistema operacional exiba o ícone correto no inicializador de aplicativos e na doca do sistema.
- **No Windows (Compilação Nativa):**
  - O utilitário PyInstaller e os scripts de compilação [`build_windows_gui.ps1`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/build_windows_gui.ps1) injetam o `assets/icon.ico` diretamente na tabela de recursos do executável PE32+ (`.exe`), fazendo com que o ícone apareça no Windows Explorer mesmo antes do aplicativo ser executado.
- **Na Janela Web (Favicon e WebKit):**
  - O cabeçalho de [`assets/ui.html`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/assets/ui.html) inclui `<link rel="icon" type="image/png" href="icon.png">`, garantindo consistência até mesmo se a interface for aberta em navegadores web tradicionais.

### 8.5 A Linha de Montagem em Nuvem do GitHub Actions e Publicação Automatizada de Releases
Para que líderes de ministério, secretarias e voluntários não precisem configurar compiladores, Rust ou ambientes Python em seus computadores, o projeto conta com um esteira industrial automatizada de integração contínua (CI/CD) descrita em [`.github/workflows/release.yml`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/.github/workflows/release.yml).

```mermaid
sequenceDiagram
    autonumber
    actor Dev as Desenvolvedor (Git Push)
    participant GH as GitHub Actions Runner
    participant LNX as Matriz Linux (Ubuntu)
    participant WIN as Matriz Windows (Server)
    participant REL as GitHub Releases (Público)

    Dev->>GH: Push de Tag (ex: v1.6.0)
    par Compilação Paralela
        GH->>LNX: 1. Compila Rust CLI (x86_64) + Gera AppImage (.image)
        GH->>WIN: 2. Compila Rust CLI (.exe) + Gera Windows GUI (.exe)
    end
    LNX-->>GH: Upload dos artefatos Linux
    WIN-->>GH: Upload dos artefatos Windows
    GH->>REL: Cria Release oficial anexando todos os executáveis
    REL-->>Dev: Notificação de Release Pública Disponível
```

#### 8.5.1 Como Funciona a Matriz Multiplataforma
O GitHub Actions aloca duas máquinas virtuais isoladas simultaneamente:
1. **Ambiente Ubuntu (`ubuntu-latest`):**
   - Instala o compilador Rust estável e o alvo `x86_64-unknown-linux-gnu`.
   - Injeta via APT as bibliotecas de sistema `libfuse2` (para execução do `appimagetool` em contêineres sem FUSE) e o ecossistema `WebKitGTK` (`libwebkit2gtk-4.0-37`, `gir1.2-webkit2-4.0`).
   - Executa o script [`build_appimage.sh`](file:///home/gabriel/Documentos/GitHub/Convert_IMCA/build_appimage.sh), gerando o binário portátil `Convert_IMCA-linux-x86_64.AppImage`.
2. **Ambiente Windows (`windows-latest`):**
   - Instala o compilador Rust estável para `x86_64-pc-windows-msvc`.
   - Utiliza **PowerShell Core nativo (`pwsh`)** para disparar o PyInstaller com injeção do ícone multi-resolução (`assets\icon.ico`), empacotando o arquivo único sem dependências `Convert_IMCA_GUI-windows.exe`.

#### 8.5.2 A Publicação na Aba de Releases
Ao final das compilações paralelas, a ação `softprops/action-gh-release@v2` reúne todos os artefatos gerados:
- O pacote Linux auto-executável (`.AppImage`).
- O executável de desktop para Windows (`.exe`).
- O utilitário de alta performance CLI em Rust para Linux e Windows.
- E gera automaticamente as notas de lançamento (*release notes*), disponibilizando o download direto para qualquer usuário final da igreja.




