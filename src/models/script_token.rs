use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use crate::models::location::Location;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScriptToken {
    pub token_type: String,  // token类型: scene_heading, character, dialogue等
    pub text: String,        // 打印文本内容
    pub line: usize,         // 所在行号
    pub start: usize,        // 起始位置
    pub end: usize,          // 结束位置
    pub is_dual_dialogue: bool, // 是否为双对话
    pub dual: Option<String>, // 双对话位置: 'left' 或 'right'
    pub duration_sec: Option<f64>, // 对话时长(秒)
    pub time: Option<f64>,   // 对话时间(秒)
    pub location_info: Option<Location>, // 场景位置信息(仅场景标题有效)
    pub metadata: Option<HashMap<String, String>>, // 额外元数据
    pub index: i32,
    pub number: Option<String>, // 场景编号
    pub text_no_notes: Option<String>, // 无注释文本
    pub character: Option<String>, // 角色名
    pub take_number: Option<i32>, // 拍摄次数
    pub level: Option<i32>,  // 层级
    pub ignore: bool,        // 是否忽略
    pub characters_action: Option<Vec<String>>, // action 类型专用，该 action 行包含哪些角色
    pub play_time_sec: f64,  // 对应行结束后在影片中的时间进度
    pub invisible_sections: Option<Vec<ScriptToken>>, // 不可见的章节（用于创建书签和生成docx侧边栏）
}

impl ScriptToken {
    pub fn new(
        token_type: String,
        text: String,
        line: usize,
        start: usize,
        end: usize,
    ) -> Self {
        ScriptToken {
            token_type,
            text,
            line,
            start,
            end,
            is_dual_dialogue: false,
            dual: None,
            duration_sec: None,
            time: None,
            location_info: None,
            metadata: None,
            index: -1,
            number: None,
            text_no_notes: None,
            character: None,
            take_number: None,
            level: None,
            ignore: false,
            characters_action: None,
            play_time_sec: 0.0,
            invisible_sections: None,
        }
    }

    // 创建一个新的空token
    pub fn empty() -> Self {
        ScriptToken {
            token_type: String::new(),
            text: String::new(),
            line: 0,
            start: 0,
            end: 0,
            is_dual_dialogue: false,
            dual: None,
            duration_sec: None,
            time: None,
            location_info: None,
            metadata: None,
            index: -1,
            number: None,
            text_no_notes: None,
            character: None,
            take_number: None,
            level: None,
            ignore: false,
            characters_action: None,
            play_time_sec: 0.0,
            invisible_sections: None,
        }
    }

    // 检查token类型是否匹配
    pub fn is_type(&self, types: &[&str]) -> bool {
        types.contains(&self.token_type.as_str())
    }

    // 获取清理后的文本(去除格式标记等)
    pub fn clean_text(&self) -> String {
        // 移除星号和下划线样式的标记
        let mut t = self.text.replace(&['*', '_'][..], "");

        // 移除内联注释 [[...]] 或 /* ... */
        t = t.replace(r"[[", "").replace(r"]]", "")
            .replace(r"/*", "").replace(r"*/", "");

        t.trim().to_string()
    }

    // 转换为HTML格式(用于预览)
    pub fn to_html(&self) -> String {
        let cleaned = self.clean_text();
        match self.token_type.as_str() {
            "scene_heading" => format!("<div class=\"scene-heading\">{}</div>", cleaned),
            "character" => format!("<div class=\"character\">{}</div>", cleaned),
            "dialogue" => format!("<div class=\"dialogue\">{}</div>", cleaned),
            "parenthetical" => format!("<div class=\"parenthetical\">{}</div>", cleaned),
            "action" => format!("<div class=\"action\">{}</div>", cleaned),
            _ => format!("<div class=\"fountain-{}\">{}</div>", self.token_type, cleaned),
        }
    }
}
