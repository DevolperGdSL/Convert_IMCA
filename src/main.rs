mod excel;
mod imca;
mod types;

use clap::Parser;
use excel::{extract_events_from_excel, ExcelExtractOptions};
use imca::{serialize_imca, validate_and_parse_imca};
use std::fs;
use std::path::PathBuf;
use std::time::Instant;

#[derive(Parser, Debug)]
#[command(
    name = "convert_imca",
    author = "Gabriel <DevolperGdSL>",
    version = "1.3.0",
    about = "Conversor ultra-rápido de planilhas Excel (.xlsx) para o formato leve .IMCA (cabeceira-pwa1)"
)]
struct Args {
    /// Arquivo de entrada Excel (.xlsx)
    #[arg(value_name = "ARQUIVO_XLSX")]
    input: PathBuf,

    /// Arquivo de saída (.imca). Se omitido, usa o mesmo nome da entrada com extensão .imca
    #[arg(short, long, value_name = "ARQUIVO_IMCA")]
    output: Option<PathBuf>,

    /// Nome específico da aba da planilha a ser processada
    #[arg(short, long, value_name = "NOME_ABA")]
    sheet: Option<String>,

    /// Ano padrão para contexto das datas (ex: 2026). Detectado automaticamente se omitido.
    #[arg(short = 'y', long, value_name = "ANO")]
    year: Option<i32>,

    /// Horário padrão atribuído aos eventos que não especificarem hora
    #[arg(short = 't', long, default_value = "19:30", value_name = "HH:MM")]
    default_time: String,

    /// Traduz os nomes dos dias da semana de inglês para português (ex: Wednesday -> Quarta-feira)
    #[arg(long)]
    translate_days: bool,

    /// Exibe uma prévia dos primeiros 5 eventos convertidos no terminal
    #[arg(short, long)]
    preview: bool,

    /// Valida a integridade do arquivo .imca gerado após a conversão
    #[arg(long, default_value_t = true)]
    validate: bool,
}

fn main() {
    let args = Args::parse();
    let start_time = Instant::now();

    if !args.input.exists() {
        eprintln!("❌ Erro: Arquivo de entrada '{:?}' não foi encontrado!", args.input);
        std::process::exit(1);
    }

    let output_path = args.output.unwrap_or_else(|| {
        let mut out = args.input.clone();
        out.set_extension("imca");
        out
    });

    println!("============================================================");
    println!(" 🚀 CONVERSOR IMCA v{} (High-Performance Engine)", env!("CARGO_PKG_VERSION"));
    println!("============================================================");
    println!(" 📂 Arquivo de Entrada : {}", args.input.display());
    println!(" 💾 Arquivo de Saída   : {}", output_path.display());

    let extract_opts = ExcelExtractOptions {
        sheet_name: args.sheet,
        fallback_year: args.year,
        default_time: args.default_time,
        translate_weekdays: args.translate_days,
    };

    let (metadata, events) = match extract_events_from_excel(&args.input, &extract_opts) {
        Ok(res) => res,
        Err(e) => {
            eprintln!("❌ Falha na extração da planilha: {}", e);
            std::process::exit(1);
        }
    };

    let total_events = events.len();
    if total_events == 0 {
        eprintln!("⚠️ Aviso: Nenhum evento válido foi encontrado na planilha.");
        std::process::exit(0);
    }

    let imca_content = serialize_imca(&metadata, &events);

    if let Err(e) = fs::write(&output_path, &imca_content) {
        eprintln!("❌ Erro ao gravar arquivo .imca de saída: {}", e);
        std::process::exit(1);
    }

    // Validação de integridade
    if args.validate {
        match validate_and_parse_imca(&imca_content) {
            Ok((val_meta, val_events)) => {
                println!(" ✅ Validação de Integridade: OK ({} registros validados com sucesso)", val_events.len());
                let _ = val_meta;
            }
            Err(err) => {
                eprintln!("❌ Erro de validação do arquivo gerado: {}", err);
                std::process::exit(1);
            }
        }
    }

    // Estatísticas de compressão
    let orig_size = fs::metadata(&args.input).map(|m| m.len()).unwrap_or(0);
    let imca_size = imca_content.as_bytes().len() as u64;
    let reduction_pct = if orig_size > 0 {
        (1.0 - (imca_size as f64 / orig_size as f64)) * 100.0
    } else {
        0.0
    };

    let elapsed = start_time.elapsed();

    println!("------------------------------------------------------------");
    println!(" 📊 ESTATÍSTICAS DE CONVERSÃO:");
    println!("   • Total de Eventos Processados : {}", total_events);
    println!("   • Ano Base Detectado           : {}", metadata.year);
    println!("   • Tamanho Planilha (.xlsx)     : {:.2} KB ({} bytes)", orig_size as f64 / 1024.0, orig_size);
    println!("   • Tamanho Formato (.imca)      : {:.2} KB ({} bytes)", imca_size as f64 / 1024.0, imca_size);
    println!("   • Redução Líquida de Dados     : {:.1}% de economia", reduction_pct);
    println!("   • Tempo de Execução Total      : {:.2} ms", elapsed.as_secs_f64() * 1000.0);
    println!("------------------------------------------------------------");

    // Prévia dos primeiros registros se solicitado
    if args.preview {
        println!("\n🔍 PRÉVIA DOS PRIMEIROS EVENTOS (5 de {}):", total_events);
        for (idx, ev) in events.iter().take(5).enumerate() {
            println!(
                "   [{:02}] {} | {} ({}) | {} | Local: '{}' | Solicitante: '{}'",
                idx + 1,
                ev.date,
                ev.name,
                ev.day_of_week,
                ev.time,
                if ev.location.is_empty() { "N/A" } else { &ev.location },
                if ev.requester.is_empty() { "N/A" } else { &ev.requester }
            );
        }
        println!();
    }

    println!("✨ Sucesso! Arquivo pronto para consumo no cabeceira-pwa1.\n");
}
