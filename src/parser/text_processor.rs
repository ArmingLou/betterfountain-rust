use regex::Regex;
use crate::models::ScriptToken;
use crate::utils::FountainConstants;

// 处理token文本样式
pub fn process_token_text_style_char(token: &mut ScriptToken) -> String {
    if !token.text.is_empty() {
        // 三 *** 换成当个特殊符号，以防下面split_token分行截断
        token.text = Regex::new(r"\*{3}")
            .unwrap()
            .replace_all(&token.text, FountainConstants::style_chars()["bold_italic"])
            .to_string();

        // 双 ** 换成当个特殊符号 ↭，以防下面split_token分行截断
        token.text = Regex::new(r"\*{2}")
            .unwrap()
            .replace_all(&token.text, FountainConstants::style_chars()["bold"])
            .to_string();

        // 单 * 换成特殊符号
        token.text = Regex::new(r"\*")
            .unwrap()
            .replace_all(&token.text, FountainConstants::style_chars()["italic"])
            .to_string();

        // 下划线 _ 换成特殊符号
        token.text = Regex::new(r"_")
            .unwrap()
            .replace_all(&token.text, FountainConstants::style_chars()["underline"])
            .to_string();

        // 处理转义字符
        token.text = token.text.replace(r"\*", "*").replace(r"\_", "_");
    }

    token.text.clone()
}

// 使用 utils/mod.rs 中的 is_blank_line_after_style 函数

// 生成HTML输出
pub fn generate_html(tokens: &[ScriptToken]) -> String {
    let mut buffer = String::new();
    for token in tokens {
        buffer.push_str(&token.to_html());
        buffer.push('\n');
    }
    buffer
}

// 生成标题页HTML输出
pub fn generate_title_html(title_keys: &[String], tokens: &[ScriptToken]) -> String {
    let mut buffer = String::new();
    for key in title_keys {
        if let Some(token) = tokens.iter().find(|t| {
            t.metadata.as_ref()
                .and_then(|m| m.get("key"))
                .map_or(false, |k| k == key)
        }) {
            buffer.push_str(&token.to_html());
            buffer.push('\n');
        }
    }
    buffer
}
