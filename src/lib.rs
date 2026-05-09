pub mod models;
pub mod utils;
pub mod parser;
pub mod docx;
pub mod pdf;
pub mod api;
pub mod statistics;

pub use models::{
    ScriptToken,
    StructToken,
    Location,
    ScreenplayProperties,
    Conf,
    Position,
    Range,
    Synopsis,
    Note
};

pub use parser::{
    FountainParser,
    ParseOutput,
    TitleKeywordFormat
};

pub use docx::{
    DocxOptions,
    DocxResult,
    generate_docx
};

pub use api::{
    SimpleConf,
    ExportResult,
    parse_fountain_text,
    export_to_docx,
    export_to_docx_base64,
    test_connection
};

/// 解析Fountain格式文本
///
/// # Arguments
///
/// * `script` - Fountain格式的剧本文本
/// * `config` - 配置对象
/// * `generate_html` - 是否生成HTML输出
///
/// # Returns
///
/// 解析结果对象
/// 
/// # Arguments
///
/// * `script` - Fountain 格式的剧本文本
/// * `config` - 配置对象
/// * `generate_html` - 是否生成 HTML 输出
/// * `calc_statistics` - 是否计算统计数据（可选，默认 false）
pub fn parse(script: &str, config: &Conf, generate_html: bool, calc_statistics: Option<bool>) -> ParseOutput {
    let mut parser = FountainParser::new();
    parser.parse(script, config, generate_html, calc_statistics)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let config = Conf::default();
        let result = parse("INT. ROOM - DAY\n\nHello, world!", &config, false, None);
        assert!(!result.tokens.is_empty());
    }
}
