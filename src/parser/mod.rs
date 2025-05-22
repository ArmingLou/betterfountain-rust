pub mod fountain_parser;
pub mod text_processor;

pub use fountain_parser::FountainParser;
pub use fountain_parser::ParseOutput;
pub use fountain_parser::TitleKeywordFormat;
pub use text_processor::{
    process_token_text_style_char,
    generate_html,
    generate_title_html
};
pub use crate::utils::is_blank_line_after_style;
