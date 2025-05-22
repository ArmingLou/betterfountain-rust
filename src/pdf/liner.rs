use crate::parser::fountain_parser::Line;
use crate::models::{ScriptToken, Conf};
use crate::utils::is_blank_line_after_style;

/// 行处理器
pub struct Liner {
    /// 是否打印拍摄次数
    pub print_take_numbers: bool,
    /// 当前状态
    pub state: String,
}

impl Liner {
    /// 创建新的行处理器
    pub fn new(print_take_numbers: bool) -> Self {
        Self {
            print_take_numbers,
            state: "normal".to_string(),
        }
    }

    /// 分割token为行
    pub fn split_token3(&self, token: &ScriptToken) -> Vec<Line> {
        let tmp_text = if token.token_type == "character" && self.print_take_numbers {
            if let Some(take_number) = token.take_number {
                format!("{} - {}", take_number, token.text)
            } else {
                token.text.clone()
            }
        } else {
            token.text.clone()
        };

        let lines = tmp_text.split('\n');
        let mut result = Vec::new();
        let mut st = token.start;

        for (i, line_text) in lines.enumerate() {
            let l = line_text.len();
            result.push(Line {
                token_type: token.token_type.clone(),
                token: Some(token.line),
                text: line_text.to_string(),
                start: st,
                end: if l > 0 { st + l - 1 } else { st },
                local_index: i,
                global_index: 0, // 将在line2中设置
                number: token.number.clone(),
                dual: token.dual.clone(),
                level: token.level.clone(),
            });
            st += l;
        }

        result
    }

    /// 处理tokens为行
    pub fn line2(&self, tokens: &[ScriptToken], config: &Conf) -> Vec<Line> {
        let mut lines: Vec<Line> = Vec::new();
        let mut global_index = 0;
        let mut last_line_blank = false;

        for token in tokens {
            if token.ignore {
                continue;
            }

            // 替换制表符为4个空格
            let token_text = token.text.replace('\t', "    ");
            let token_with_replaced_tabs = ScriptToken {
                text: token_text,
                ..token.clone()
            };

            let mut token_lines = self.split_token3(&token_with_replaced_tabs);

            // 处理场景标题的编号
            if token.token_type == "scene_heading" && !lines.is_empty() {
                if let Some(number) = &token.number {
                    token_lines[0].number = Some(number.clone());
                }
            }

            for (index, line) in token_lines.iter_mut().enumerate() {
                let mut pushed = true;

                if config.merge_empty_lines {
                    if token.token_type == "page_break" {
                        last_line_blank = false; // 需要保留行
                    } else {
                        let curr_blank = is_blank_line_after_style(&line.text);
                        if curr_blank && last_line_blank {
                            if let Some(last_line) = lines.last_mut() {
                                // 剩下样式符号
                                let t = line.text.replace(|c: char| c.is_whitespace(), "");
                                // 加到上一行
                                last_line.text.push_str(&t);
                                pushed = false;
                            }
                        }
                        last_line_blank = curr_blank;
                    }
                }

                if pushed {
                    line.local_index = index;
                    line.global_index = global_index;
                    global_index += 1;
                    lines.push(line.clone());
                }
            }
        }

        lines
    }
}
