use crate::types::{EventRecord, ImcaMetadata};
use chrono::NaiveDate;

/// Serializa os metadados e a lista de eventos no formato textual padronizado .IMCA
pub fn serialize_imca(metadata: &ImcaMetadata, events: &[EventRecord]) -> String {
    let mut out = String::with_capacity(events.len() * 120 + 512);

    // Cabeçalho de Identificação
    out.push_str("# =============================================================================\n");
    out.push_str("# FORMATO DE DADOS IMCA v1.0 - AGENDA E EVENTOS OTIMIZADOS\n");
    out.push_str("# APLICATIVO ALVO: cabeceira-pwa1\n");
    out.push_str("# PADRÃO: DELIMITADO POR PIPE (|) COM CABEÇALHO E METADADOS INLINE\n");
    out.push_str("# =============================================================================\n\n");

    // Bloco [METADATA]
    out.push_str("[METADATA]\n");
    out.push_str(&format!("versao={}\n", metadata.version));
    out.push_str(&format!("ano={}\n", metadata.year));
    out.push_str(&format!("total_registros={}\n", events.len()));
    out.push_str(&format!("gerado_em={}\n", metadata.generated_at));
    out.push_str(&format!("arquivo_origem={}\n", metadata.source_file));
    out.push_str(&format!("aplicativo_alvo={}\n", metadata.target_app));
    out.push_str(&format!("delimitador={}\n\n", metadata.delimiter));

    // Bloco [SCHEMA]
    out.push_str("[SCHEMA]\n");
    out.push_str("id|data|dia_semana|horario|nome|local|solicitante|status\n\n");

    // Bloco [DATA]
    out.push_str("[DATA]\n");
    for ev in events {
        out.push_str(&format!(
            "{}|{}|{}|{}|{}|{}|{}|{}\n",
            ev.id,
            ev.date.format("%Y-%m-%d"),
            ev.day_of_week,
            ev.time,
            ev.name,
            ev.location,
            ev.requester,
            ev.status
        ));
    }

    out
}

/// Valida e realiza o parse de um conteúdo .IMCA para conferência de integridade
pub fn validate_and_parse_imca(content: &str) -> Result<(ImcaMetadata, Vec<EventRecord>), String> {
    let mut metadata = ImcaMetadata::default();
    let mut events = Vec::new();
    let mut current_section = "";

    for line in content.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() || trimmed.starts_with('#') {
            continue;
        }

        if trimmed.starts_with('[') && trimmed.ends_with(']') {
            current_section = match trimmed {
                "[METADATA]" => "METADATA",
                "[SCHEMA]" => "SCHEMA",
                "[DATA]" => "DATA",
                _ => "UNKNOWN",
            };
            continue;
        }

        match current_section {
            "METADATA" => {
                if let Some((k, v)) = trimmed.split_once('=') {
                    match k.trim() {
                        "versao" => metadata.version = v.trim().to_string(),
                        "ano" => metadata.year = v.trim().parse().unwrap_or(2026),
                        "total_registros" => metadata.total_records = v.trim().parse().unwrap_or(0),
                        "gerado_em" => metadata.generated_at = v.trim().to_string(),
                        "arquivo_origem" => metadata.source_file = v.trim().to_string(),
                        "aplicativo_alvo" => metadata.target_app = v.trim().to_string(),
                        "delimitador" => metadata.delimiter = v.trim().chars().next().unwrap_or('|'),
                        _ => {}
                    }
                }
            }
            "SCHEMA" => {
                // Apenas validação de schema
            }
            "DATA" => {
                let parts: Vec<&str> = trimmed.split('|').collect();
                if parts.len() >= 8 {
                    let date = NaiveDate::parse_from_str(parts[1].trim(), "%Y-%m-%d")
                        .map_err(|e| format!("Data inválida '{}': {}", parts[1], e))?;

                    events.push(EventRecord {
                        id: parts[0].trim().to_string(),
                        date,
                        day_of_week: parts[2].trim().to_string(),
                        time: parts[3].trim().to_string(),
                        name: parts[4].trim().to_string(),
                        location: parts[5].trim().to_string(),
                        requester: parts[6].trim().to_string(),
                        status: parts[7].trim().to_string(),
                    });
                }
            }
            _ => {}
        }
    }

    if metadata.total_records > 0 && metadata.total_records != events.len() {
        return Err(format!(
            "Divergência de contagem: cabeçalho indica {} registros, mas foram lidos {}.",
            metadata.total_records,
            events.len()
        ));
    }

    Ok((metadata, events))
}
