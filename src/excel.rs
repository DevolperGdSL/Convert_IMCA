use crate::types::{EventRecord, ImcaMetadata};
use calamine::{open_workbook_auto, Data, DataType, Reader};
use chrono::{Datelike, NaiveDate, Utc};
use std::path::Path;

/// Traduz o nome do dia da semana em inglês para português
pub fn translate_weekday(weekday: &str) -> &'static str {
    match weekday.to_lowercase().as_str() {
        "monday" | "mon" | "segunda" => "Segunda-feira",
        "tuesday" | "tue" | "terça" | "terca" => "Terça-feira",
        "wednesday" | "wed" | "quarta" => "Quarta-feira",
        "thursday" | "thu" | "quinta" => "Quinta-feira",
        "friday" | "fri" | "sexta" => "Sexta-feira",
        "saturday" | "sat" | "sábado" | "sabado" => "Sábado",
        "sunday" | "sun" | "domingo" => "Domingo",
        _ => "Data Especial",
    }
}

/// Extrai o ano a partir do nome do arquivo (ex: "Calendario 2026.xlsx" -> 2026)
pub fn extract_year_from_path(path: &Path) -> Option<i32> {
    let filename = path.file_name()?.to_str()?;
    for word in filename.split(|c: char| !c.is_numeric()) {
        if word.len() == 4 {
            if let Ok(y) = word.parse::<i32>() {
                if (2000..=2100).contains(&y) {
                    return Some(y);
                }
            }
        }
    }
    None
}

/// Configurações do processo de extração da planilha
#[derive(Debug, Clone)]
pub struct ExcelExtractOptions {
    pub sheet_name: Option<String>,
    pub fallback_year: Option<i32>,
    pub default_time: String,
    pub translate_weekdays: bool,
}

impl Default for ExcelExtractOptions {
    fn default() -> Self {
        Self {
            sheet_name: None,
            fallback_year: None,
            default_time: "19:30".to_string(),
            translate_weekdays: false,
        }
    }
}

/// Processa o arquivo Excel e extrai todos os registros de eventos válidos
pub fn extract_events_from_excel(
    file_path: &Path,
    options: &ExcelExtractOptions,
) -> Result<(ImcaMetadata, Vec<EventRecord>), Box<dyn std::error::Error>> {
    let mut workbook = open_workbook_auto(file_path)?;
    let sheet_names = workbook.sheet_names();
    
    if sheet_names.is_empty() {
        return Err("A planilha Excel não contém nenhuma aba visível.".into());
    }

    // Seleciona a aba alvo (preferência explícita, ou busca "Calendario", ou a primeira)
    let target_sheet = if let Some(ref explicit) = options.sheet_name {
        if !sheet_names.contains(explicit) {
            return Err(format!(
                "Aba '{}' não encontrada. Abas disponíveis: {:?}",
                explicit, sheet_names
            )
            .into());
        }
        explicit.clone()
    } else {
        // Tenta encontrar "Calendario", "Agenda" ou aba similar
        sheet_names
            .iter()
            .find(|name| {
                let n = name.to_lowercase();
                n.contains("calendario") || n.contains("agenda") || n.contains("eventos")
            })
            .cloned()
            .unwrap_or_else(|| sheet_names[0].clone())
    };

    let range = workbook.worksheet_range(&target_sheet)?;
    let detected_year = options
        .fallback_year
        .or_else(|| extract_year_from_path(file_path))
        .unwrap_or(2026);

    let mut events = Vec::new();
    let mut seq = 1;

    // Descobrir posições de colunas a partir do cabeçalho
    let mut col_data_idx = 0;
    let mut col_dia_idx = 1;
    let mut col_evento_idx = 2;
    let mut col_local_idx = 3;
    let mut col_solic_idx = 4;
    let mut header_found = false;
    let mut start_row_idx = 0;

    for (row_idx, row) in range.rows().enumerate().take(15) {
        let texts: Vec<String> = row
            .iter()
            .map(|c| match c {
                Data::String(s) => s.trim().to_lowercase(),
                _ => String::new(),
            })
            .collect();

        if let Some(pos_data) = texts.iter().position(|s| s == "data" || s == "datas") {
            col_data_idx = pos_data;
            if let Some(pos_ev) = texts.iter().position(|s| s.contains("evento") || s.contains("nome")) {
                col_evento_idx = pos_ev;
            }
            if let Some(pos_dia) = texts.iter().position(|s| s.contains("dia") || s.contains("semana")) {
                col_dia_idx = pos_dia;
            }
            if let Some(pos_loc) = texts.iter().position(|s| s.contains("local")) {
                col_local_idx = pos_loc;
            }
            if let Some(pos_sol) = texts.iter().position(|s| s.contains("solicitante") || s.contains("responsavel")) {
                col_solic_idx = pos_sol;
            }
            header_found = true;
            start_row_idx = row_idx + 1;
            break;
        }
    }

    if !header_found {
        start_row_idx = 1; // Assume linha 1 em diante se não achar texto exato
    }

    // Leitura das linhas de dados
    for row in range.rows().skip(start_row_idx) {
        if row.is_empty() {
            continue;
        }

        // Extrai a data
        let date_cell = row.get(col_data_idx);
        let date_opt: Option<NaiveDate> = match date_cell {
            Some(Data::DateTime(_)) | Some(Data::Float(_)) => date_cell.and_then(|c| c.as_date()),
            Some(Data::String(s)) => {
                let trimmed = s.trim();
                if trimmed.is_empty() {
                    None
                } else if let Ok(d) = NaiveDate::parse_from_str(trimmed, "%Y-%m-%d") {
                    Some(d)
                } else if let Ok(d) = NaiveDate::parse_from_str(trimmed, "%d/%m/%Y") {
                    Some(d)
                } else {
                    // Tenta formato DD/MM adicionando o ano detectado
                    let parts: Vec<&str> = trimmed.split('/').collect();
                    if parts.len() == 2 {
                        if let (Ok(dia), Ok(mes)) = (parts[0].parse::<u32>(), parts[1].parse::<u32>()) {
                            NaiveDate::from_ymd_opt(detected_year, mes, dia)
                        } else {
                            None
                        }
                    } else {
                        None
                    }
                }
            }
            _ => None,
        };

        let date = match date_opt {
            Some(d) => d,
            None => continue, // Sem data válida, ignora
        };

        // Extrai o nome do evento
        let name_cell = row.get(col_evento_idx);
        let raw_name = match name_cell {
            Some(Data::String(s)) => s.trim().to_string(),
            _ => continue, // Sem nome de evento, ignora
        };

        // Filtra notas de instrução de planilha (ex: avisos de reordenamento)
        let name_upper = raw_name.to_uppercase();
        if name_upper.contains("SEMPRE REORDENAR") || name_upper.contains("DO MAIS ANTIGO") || raw_name.is_empty() {
            continue;
        }

        // Dia da semana
        let raw_day = row
            .get(col_dia_idx)
            .and_then(|c| match c {
                Data::String(s) => Some(s.trim().to_string()),
                _ => None,
            })
            .filter(|s| !s.is_empty())
            .unwrap_or_else(|| date.weekday().to_string());

        let day_of_week = if options.translate_weekdays {
            translate_weekday(&raw_day).to_string()
        } else {
            raw_day
        };

        // Local
        let location = row
            .get(col_local_idx)
            .and_then(|c| match c {
                Data::String(s) => Some(s.trim().to_string()),
                _ => None,
            })
            .unwrap_or_default();

        // Solicitante
        let requester = row
            .get(col_solic_idx)
            .and_then(|c| match c {
                Data::String(s) => Some(s.trim().to_string()),
                _ => None,
            })
            .unwrap_or_default();

        // Limpeza de caracteres delimitadores para não quebrar o formato pipe
        let clean_name = raw_name.replace('|', "-").replace('\n', " ").replace('\r', "");
        let clean_loc = location.replace('|', "-").replace('\n', " ").replace('\r', "");
        let clean_req = requester.replace('|', "-").replace('\n', " ").replace('\r', "");

        // Gera ID canônico único: ev_YYYYMMDD_XXX
        let id = format!("ev_{:04}{:02}{:02}_{:03}", date.year(), date.month(), date.day(), seq);
        seq += 1;

        events.push(EventRecord {
            id,
            date,
            day_of_week,
            time: options.default_time.clone(),
            name: clean_name,
            location: clean_loc,
            requester: clean_req,
            status: "publicado".to_string(),
        });
    }

    let metadata = ImcaMetadata {
        version: "1.0".to_string(),
        year: detected_year,
        total_records: events.len(),
        generated_at: Utc::now().to_rfc3339(),
        source_file: file_path.file_name().and_then(|s| s.to_str()).unwrap_or("origem.xlsx").to_string(),
        target_app: "cabeceira-pwa1".to_string(),
        delimiter: '|',
    };

    Ok((metadata, events))
}
