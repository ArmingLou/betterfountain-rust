use std::collections::HashMap;
use crate::parser::ParseOutput;

/// 元数据提取结果
#[derive(Debug, Clone)]
pub struct ExtractedMetadata {
    pub metadata: HashMap<String, String>,
    pub watermark: Option<String>,
    pub header: Option<String>,
    pub footer: Option<String>,
    pub font: String,
    pub font_bold: String,
    pub font_italic: String,
    pub font_bold_italic: String,
}

impl Default for ExtractedMetadata {
    fn default() -> Self {
        Self {
            metadata: HashMap::new(),
            watermark: None,
            header: None,
            footer: None,
            font: "Courier Prime".to_string(),
            font_bold: String::new(),
            font_italic: String::new(),
            font_bold_italic: String::new(),
        }
    }
}

/// 将嵌套的 JSON 对象转换为扁平的键值对
///
/// 例如，将以下 JSON：
/// ```json
/// {
///     "print": {
///         "chinaFormat": 3,
///         "rmBlankLine": 0
///     }
/// }
/// ```
///
/// 转换为：
/// ```
/// "print.chinaFormat" => "3"
/// "print.rmBlankLine" => "0"
/// "print" => "true"
/// ```
fn flatten_json_to_hashmap(value: &serde_json::Value, prefix: &str, result: &mut HashMap<String, String>) {
    match value {
        serde_json::Value::Object(map) => {
            // 对于对象，添加一个标记键，表示该对象存在
            if !prefix.is_empty() {
                result.insert(prefix.to_string(), "true".to_string());
            }

            // 递归处理对象的每个字段
            for (key, val) in map {
                let new_prefix = if prefix.is_empty() {
                    key.clone()
                } else {
                    format!("{}.{}", prefix, key)
                };
                flatten_json_to_hashmap(val, &new_prefix, result);
            }
        },
        serde_json::Value::Array(arr) => {
            // 对于数组，将每个元素添加为带索引的键
            for (i, val) in arr.iter().enumerate() {
                let new_prefix = format!("{}[{}]", prefix, i);
                flatten_json_to_hashmap(val, &new_prefix, result);
            }
        },
        serde_json::Value::String(s) => {
            // 对于字符串，直接添加
            result.insert(prefix.to_string(), s.clone());
        },
        serde_json::Value::Number(n) => {
            // 对于数字，转换为字符串后添加
            if let Some(i) = n.as_i64() {
                result.insert(prefix.to_string(), i.to_string());
            } else if let Some(f) = n.as_f64() {
                result.insert(prefix.to_string(), f.to_string());
            }
        },
        serde_json::Value::Bool(b) => {
            // 对于布尔值，转换为字符串后添加
            result.insert(prefix.to_string(), b.to_string());
        },
        serde_json::Value::Null => {
            // 对于 null，添加为空字符串
            result.insert(prefix.to_string(), "".to_string());
        },
    }
}

/// 从解析结果中提取元数据
///
/// # 参数
///
/// * `parsed_document` - 解析后的文档
/// * `default_font` - 默认字体
///
/// # 返回值
///
/// 提取的元数据
pub fn extract_metadata_from_parsed_document(
    parsed_document: &ParseOutput,
    default_font: &str,
) -> ExtractedMetadata {
    let mut result = ExtractedMetadata::default();
    result.font = default_font.to_string();

    // 从标题页中提取元数据
    if !parsed_document.title_page.is_empty() {
        if let Some(hidden_tokens) = parsed_document.title_page.get("hidden") {
            for token in hidden_tokens {
                match token.token_type.as_str() {
                    "watermark" => {
                        result.watermark = Some(token.text.clone());
                    },
                    "header" => {
                        result.header = Some(token.text.clone());
                    },
                    "footer" => {
                        result.footer = Some(token.text.clone());
                    },
                    "font" => {
                        result.font = token.text.clone();
                    },
                    "font_italic" => {
                        result.font_italic = token.text.clone();
                    },
                    "font_bold" => {
                        result.font_bold = token.text.clone();
                    },
                    "font_bold_italic" => {
                        result.font_bold_italic = token.text.clone();
                    },
                    "metadata" => {
                        let metadata_string = &token.text;
                        if !metadata_string.is_empty() {
                            // 首先尝试解析为 serde_json::Value，以处理嵌套的 JSON 结构
                            match serde_json::from_str::<serde_json::Value>(metadata_string) {
                                Ok(json_value) => {
                                    // 将 JSON 值转换为扁平的键值对
                                    flatten_json_to_hashmap(&json_value, "", &mut result.metadata);
                                },
                                Err(_) => {
                                    // 如果解析失败，尝试解析为简单的键值对
                                    match serde_json::from_str::<HashMap<String, String>>(metadata_string) {
                                        Ok(meta) => {
                                            result.metadata = meta;
                                        },
                                        Err(_) => {} // 如果解析失败，保持 metadata 不变
                                    }
                                }
                            }
                        }
                    },
                    _ => {} // 忽略其他类型的 token
                }
            }
        }
    }

    result
}
