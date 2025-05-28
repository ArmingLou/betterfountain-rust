pub mod models;
pub mod utils;
pub mod parser;
pub mod docx;
pub mod pdf;
pub mod api;

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
pub fn parse(script: &str, config: &Conf, generate_html: bool) -> ParseOutput {
    let mut parser = FountainParser::new();
    parser.parse(script, config, generate_html)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let config = Conf::default();
        let result = parse("INT. ROOM - DAY\n\nHello, world!", &config, false);
        assert!(!result.tokens.is_empty());
    }
}
