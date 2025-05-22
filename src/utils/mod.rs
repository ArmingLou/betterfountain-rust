pub mod fountain_constants;

use regex;
pub use fountain_constants::FountainConstants;

/// 检查一行文本是否为样式后的空行
///
/// 如果一行文本只包含空白字符或样式标记（如 *粗体*、_斜体_），则返回 true
pub fn is_blank_line_after_style(text: &str) -> bool {
    // 使用 FountainConstants 中的样式字符
    let style_chars = FountainConstants::style_chars()["all"];
    let pattern = format!(r"[{}]", regex::escape(style_chars));
    let re = regex::Regex::new(&pattern).unwrap();
    let t = re.replace_all(text, "");
    t.trim().is_empty()
}
