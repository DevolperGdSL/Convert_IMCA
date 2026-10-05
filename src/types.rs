use chrono::NaiveDate;

/// Representação normalizada de um evento extraído da planilha
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EventRecord {
    pub id: String,
    pub date: NaiveDate,
    pub day_of_week: String,
    pub time: String,
    pub name: String,
    pub location: String,
    pub requester: String,
    pub status: String,
}

/// Metadados contidos no cabeçalho do arquivo .IMCA
#[derive(Debug, Clone)]
pub struct ImcaMetadata {
    pub version: String,
    pub year: i32,
    pub total_records: usize,
    pub generated_at: String,
    pub source_file: String,
    pub target_app: String,
    pub delimiter: char,
}

impl Default for ImcaMetadata {
    fn default() -> Self {
        Self {
            version: "1.0".to_string(),
            year: 2026,
            total_records: 0,
            generated_at: String::new(),
            source_file: String::new(),
            target_app: "cabeceira-pwa1".to_string(),
            delimiter: '|',
        }
    }
}
