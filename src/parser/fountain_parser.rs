use std::collections::{HashMap, HashSet};
use regex::Regex;
use lazy_static::lazy_static;
use crate::models::{
    Conf,
    Location,
    ScriptToken,
    ScreenplayProperties,
    StructToken,
    Position,
    Range,
    Synopsis,
    Note
};

/// 行结构体，用于存储处理后的行信息
#[derive(Debug, Clone, serde::Serialize)]
pub struct Line {
    /// 行类型
    pub token_type: String,
    /// 原始token引用
    pub token: Option<usize>,
    /// 行文本
    pub text: String,
    /// 起始位置
    pub start: usize,
    /// 结束位置
    pub end: usize,
    /// 本地索引（在token中的索引）
    pub local_index: usize,
    /// 全局索引
    pub global_index: usize,
    /// 场景编号
    pub number: Option<String>,
    /// 双对话位置: 'left' 或 'right'
    pub dual: Option<String>,
    pub level: Option<i32>,  // 层级
}
use crate::utils::fountain_constants::BLOCK_REGEX;
use crate::utils::{FountainConstants, is_blank_line_after_style};
use crate::parser::text_processor::{process_token_text_style_char, generate_html, generate_title_html};

// 扩展Vec类型添加pushSorted方法
pub trait SortedList<T> {
    fn push_sorted(&mut self, el: T, compare_fn: impl Fn(&T, &T) -> i32) -> usize;
}

impl<T> SortedList<T> for Vec<T> {
    fn push_sorted(&mut self, el: T, compare_fn: impl Fn(&T, &T) -> i32) -> usize {
        let mut m = 0;
        let mut n = self.len().saturating_sub(1);

        while m <= n {
            let k = (n + m) >> 1;
            let cmp = compare_fn(&el, &self[k]);

            if cmp > 0 {
                m = k + 1;
            } else if cmp < 0 {
                n = k.saturating_sub(1);
            } else {
                self.insert(k, el);
                return self.len();
            }
        }

        let insert_pos = if m > 0 { m } else { 0 };
        self.insert(insert_pos, el);
        self.len()
    }
}

#[derive(Debug, Clone)]
pub struct TitleKeywordFormat {
    pub position: String,
    pub index: i32,
}

#[derive(Debug, Clone, serde::Serialize)]
pub struct ParseOutput {
    pub tokens: Vec<ScriptToken>,
    pub properties: ScreenplayProperties,
    pub length_action: f64,
    pub length_dialogue: f64,
    pub parse_time: u64,
    pub script_html: Option<String>,
    pub title_html: Option<String>,
    pub state: String,
    pub dual_str: Option<String>,
    pub title_page: HashMap<String, Vec<ScriptToken>>,
    /// 对白中每字符耗时预估(不含标点)
    pub dial_sec_per_char: f64,
    /// 对白中每个短标点耗时预估(逗号顿号等)
    pub dial_sec_per_punc_short: f64,
    /// 对白中每个长标点耗时预估(句号问号等)
    pub dial_sec_per_punc_long: f64,
    /// action文本中每字符转化成影片时长预估(不含标点)
    pub action_sec_per_char: f64,
    /// 处理后的行列表
    pub lines: Vec<Line>,
}

impl ParseOutput {
    pub fn new() -> Self {
        ParseOutput {
            tokens: Vec::new(),
            properties: ScreenplayProperties::new(),
            length_action: 0.0,
            length_dialogue: 0.0,
            parse_time: 0,
            script_html: None,
            title_html: None,
            state: "normal".to_string(),
            dual_str: None,
            title_page: HashMap::new(),
            dial_sec_per_char: 0.3,
            dial_sec_per_punc_short: 0.3,
            dial_sec_per_punc_long: 0.75,
            action_sec_per_char: 0.4,
            lines: Vec::new(),
        }
    }
}

impl Default for ParseOutput {
    fn default() -> Self {
        Self::new()
    }
}

pub struct FountainParser {
    result: ParseOutput,
    length_action_so_far: f64,
    length_dialogue_so_far: f64,
    current_depth: usize,
    scene_number: usize,
    play_time_sec: f64,
    last_scen_structure_token: Option<StructToken>,
    last_scen_structure_token_pre: Option<StructToken>,
    last_chartor_structure_token: Option<StructToken>,
    force_not_dual: bool,
    take_count: usize,
    lines_length: usize,
    current_cursor: usize,
    new_line_length: usize,
    text_display: String,
    text_valid: String,
    shot_cut: usize,
    shot_cut_strct_tokens: Vec<HashMap<String, serde_json::Value>>,
    current_outline_note_text: Vec<String>,
    current_outline_note_linenum: Vec<usize>,
    nested_comments: i32,
    nested_notes: i32,
    regex: HashMap<String, Regex>,
    title_page_display: HashMap<String, TitleKeywordFormat>,
}

impl FountainParser {
    pub fn new() -> Self {
        let mut parser = FountainParser {
            result: ParseOutput::new(),
            length_action_so_far: 0.0,
            length_dialogue_so_far: 0.0,
            current_depth: 0,
            scene_number: 1,
            play_time_sec: 0.0,
            last_scen_structure_token: None,
            last_scen_structure_token_pre: None,
            last_chartor_structure_token: None,
            force_not_dual: true,
            take_count: 1,
            lines_length: 0,
            current_cursor: 0,
            new_line_length: 1,
            text_display: String::new(),
            text_valid: String::new(),
            shot_cut: 0,
            shot_cut_strct_tokens: Vec::new(),
            current_outline_note_text: Vec::new(),
            current_outline_note_linenum: Vec::new(),
            nested_comments: 0,
            nested_notes: 0,
            regex: HashMap::new(),
            title_page_display: HashMap::new(),
        };

        // 初始化正则表达式
        parser.init_regex();
        // 初始化标题页显示配置
        parser.init_title_page_display();

        parser
    }

    // 去除空格、标点和特殊字符
    pub fn calculate_chars(&self, text: &str) -> String {
        let re = Regex::new(r"\s|\p{P}|\p{S}").unwrap();
        re.replace_all(text, "").to_string()
    }

    // 计算动作持续时间
    pub fn calculate_action_duration(&self, text: &str, config_x: Option<f64>) -> f64 {
        let x = config_x.unwrap_or(0.4); // 默认值: 0.4秒/字符

        let sanitized = self.calculate_chars(text);
        sanitized.chars().count() as f64 * x
    }

    // 计算对话持续时间
    pub fn calculate_dialogue_duration(
        &self,
        text: &str,
        config_x: Option<f64>,
        config_long: Option<f64>,
        config_short: Option<f64>
    ) -> f64 {
        let mut duration = 0.0;
        let x = config_x.unwrap_or(0.3); // 默认值: 0.3秒/字符
        let long = config_long.unwrap_or(0.75); // 长标点默认值: 0.75秒
        let short = config_short.unwrap_or(0.3); // 短标点默认值: 0.3秒

        let sanitized = self.calculate_chars(text);
        duration += sanitized.chars().count() as f64 * x;

        // 处理标点符号
        let rec = Regex::new(r"(\.|\?|\!|\:|。|？|！|：)|(\,|，|;|；|、)").unwrap();
        for cap in rec.captures_iter(text) {
            if cap.get(1).is_some() {
                duration += long; // 长标点
            } else if cap.get(2).is_some() {
                duration += short; // 短标点
            }
        }

        duration
    }

    // 创建token
    fn create_token(
        &self,
        text: Option<&str>,
        cursor: Option<usize>,
        line: Option<usize>,
        length: Option<usize>,
        token_type: &str,
    ) -> ScriptToken {
        ScriptToken {
            token_type: token_type.to_string(),
            text: text.unwrap_or("").to_string(),
            line: line.unwrap_or(0),
            start: cursor.unwrap_or(0),
            end: cursor.unwrap_or(0) + length.unwrap_or(0),
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

    // 添加token到结果
    fn push_token(&mut self, token: ScriptToken) {
        self.result.tokens.push(token);
    }

    // 处理对话块
    fn process_dialogue_block(&mut self, mut token: ScriptToken) -> ScriptToken {
        // 第一个场景之前的时间不累计（与Flutter版本保持一致）
        if !self.result.properties.scenes.is_empty() {
            token.text_no_notes = Some(self.text_valid.clone());
            let text_without_notes = &self.text_valid;

            // 计算对话持续时间
            let time = self.calculate_dialogue_duration(
                text_without_notes,
                Some(self.result.dial_sec_per_char),
                Some(self.result.dial_sec_per_punc_long),
                Some(self.result.dial_sec_per_punc_short)
            );
            token.time = Some(time);
            self.result.length_dialogue += time;
            self.play_time_sec += time;
            token.play_time_sec = self.play_time_sec;

            if let Some(last_scen_structure_token) = &mut self.last_scen_structure_token {
                let mut need = false;
                if self.shot_cut > 0 {
                    // 检查lastScenStructureToken是否已在当前shotCutStrctTokens中
                    if let Some(last_map) = self.shot_cut_strct_tokens.last() {
                        if let Some(structs) = last_map.get("structs") {
                            if let Some(structs_array) = structs.as_array() {
                                for struct_token in structs_array {
                                    if struct_token == &serde_json::to_value(&last_scen_structure_token).unwrap() {
                                        need = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                if need {
                    if let Some(last_map) = self.shot_cut_strct_tokens.last_mut() {
                        let duration = last_map.get("duration")
                            .and_then(|v| v.as_f64())
                            .unwrap_or(0.0) + time;
                        last_map.insert("duration".to_string(), serde_json::to_value(duration).unwrap());
                    }
                } else {
                    last_scen_structure_token.duration_sec += time;
                }
            }

            if let Some(last_chartor_structure_token) = &mut self.last_chartor_structure_token {
                last_chartor_structure_token.duration_sec += time;
            }
        }
        token
    }

    // 去除角色名前的@符号
    fn trim_character_force_symbol(&self, text: &str) -> String {
        let re = Regex::new(r"^[ \t]*@").unwrap();
        re.replace(text, "").to_string()
    }

    // 去除角色名后的扩展部分
    fn trim_character_extension(&self, text: &str) -> String {
        let re = Regex::new(r"[ \t]*(\(.*\)|（.*）)[ \t]*([ \t]*\^)?$").unwrap();
        re.replace(text, "").to_string()
    }

    // 处理动作文本块
    fn process_action_block(&mut self, mut token: ScriptToken) -> ScriptToken {
        // 第一个场景之前的时间不累计（与您提供的代码保持一致）
        if !self.result.properties.scenes.is_empty() {
            token.text_no_notes = Some(self.text_valid.clone());
            let text_without_notes = &self.text_valid;

            let time = self.calculate_action_duration(text_without_notes, Some(self.result.action_sec_per_char));
            token.time = Some(time);
            self.result.length_action += time;
            self.play_time_sec += time;
            token.play_time_sec = self.play_time_sec;

            // 更新场景持续时间
            if let Some(last_scen_structure_token) = &mut self.last_scen_structure_token {
                let mut need = false;
                if self.shot_cut > 0 {
                    if let Some(last_map) = self.shot_cut_strct_tokens.last() {
                        if let Some(structs) = last_map.get("structs") {
                            if let Some(structs_array) = structs.as_array() {
                                for struct_token in structs_array {
                                    if struct_token == &serde_json::to_value(&last_scen_structure_token).unwrap() {
                                        need = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                if need {
                    if let Some(last_map) = self.shot_cut_strct_tokens.last_mut() {
                        let duration = last_map.get("duration")
                            .and_then(|v| v.as_f64())
                            .unwrap_or(0.0) + time;
                        last_map.insert("duration".to_string(), serde_json::to_value(duration).unwrap());
                    }
                } else {
                    last_scen_structure_token.duration_sec += time;
                }
            }
        }
        token
    }

    // 处理标题页结束
    fn process_title_page_end(&mut self, line: usize) {
        if self.result.properties.first_token_line.is_none() || self.result.properties.first_token_line == Some(usize::MAX) {
            self.result.properties.first_token_line = Some(line);

            // 在标题页结束时添加多个separator token（与Flutter版本保持一致）
            if self.result.state == "title_page" {
                // 添加3个separator，与Flutter版本保持一致
                for _ in 0..3 {
                    let separator_token = ScriptToken {
                        token_type: "separator".to_string(),
                        text: FountainConstants::style_chars()["style_global_clean"].to_string(),
                        line,
                        start: 0,
                        end: 0,
                        ..ScriptToken::empty()
                    };
                    self.result.tokens.push(separator_token);
                }
                self.result.state = "normal".to_string();
            }
        }
    }

    // 解析场景位置信息
    fn parse_location_information(&self, scene_heading: &str) -> Option<Location> {
        let scene_heading_regex = self.regex.get("scene_heading")?;
        let match_result = scene_heading_regex.captures(scene_heading)?;

        // input group 1 is int/ext, group 2 is location and time, group 3 is scene number
        if match_result.len() < 3 {
            return None;
        }

        let split_location_from_time = Regex::new(r"(.*?)[\-–—−](.*)").ok()?.captures(match_result.get(2)?.as_str());

        let group1 = match_result.get(1)?.as_str();
        let mut i = group1.contains('I');
        let mut e = group1.contains("EX") || group1.contains("E.");

        let mut n = if let Some(time_match) = &split_location_from_time {
            time_match.get(1)?.as_str().trim().to_string()
        } else {
            match_result.get(2)?.as_str().trim().to_string()
        };

        // 处理中文场景标记 - 严格按照 TypeScript 逻辑
        if n.starts_with("(内景)") || n.starts_with("（内景）") {
            n = n.chars().skip(4).collect::<String>().trim().to_string();
            if !i {
                i = true;
            }
        } else if n.starts_with("(外景)") || n.starts_with("（外景）") {
            n = n.chars().skip(4).collect::<String>().trim().to_string();
            if !e {
                e = true;
            }
        } else if n.starts_with("(内外景)") || n.starts_with("（内外景）") {
            n = n.chars().skip(5).collect::<String>().trim().to_string();
            if !e {
                e = true;
            }
            if !i {
                i = true;
            }
        }

        // 标准化地点名称
        n = n.to_uppercase().replace(|c: char| c.is_whitespace(), " ");

        // 处理时间部分
        let day_t = if let Some(time_match) = &split_location_from_time {
            time_match.get(2).map(|m| m.as_str().trim()).unwrap_or("")
        } else {
            ""
        };
        let day_t = day_t.to_uppercase().replace(|c: char| c.is_whitespace(), " ");

        // 注释掉的逻辑（保持与原代码一致）
        // if i && e {
        //     if n.contains('/') {
        //         // 内外景， 但是 混合多地点。 归类为 不确定： i 和 e 都false
        //         i = false;
        //         e = false;
        //     }
        // }

        Some(Location {
            name: n,
            interior: i,
            exterior: e,
            time_of_day: day_t,
            scene_number: String::new(),
            line: 0,
            start_play_sec: 0.0,
        })
    }

    // 标准化文本
    fn slugify(&self, text: &str) -> String {
        text.to_uppercase()
            .replace(|c: char| c.is_whitespace(), " ")
            .replace(" -", "-")
            .replace("- ", "-")
    }

    // 获取标题页位置
    fn get_title_page_position(&self, key: &str) -> Option<&TitleKeywordFormat> {
        self.title_page_display.get(key)
    }

    // 查找指定深度下最新的section
    fn latest_section(&self, depth: usize) -> Option<StructToken> {
        // 查找第一层中最后一个符合条件的token
        if depth <= 0 {
            return None;
        }

        // 找到第一层的section
        let mut current_section = None;
        for item in self.result.properties.structure.iter().rev() {
            if item.section {
                current_section = Some(item.clone());
                break;
            }
        }

        if depth == 1 || current_section.is_none() {
            return current_section;
        }

        // 迭代查找更深层次的section
        let mut current_depth = 1;
        while current_depth < depth {
            let section = current_section.clone().unwrap();
            if section.isscene {
                break;
            }

            // 在子节点中查找section
            let mut found_child = false;
            for child in section.children.iter().rev() {
                if child.section {
                    current_section = Some(child.clone());
                    found_child = true;
                    break;
                }
            }

            if !found_child {
                break;
            }

            current_depth += 1;
        }

        current_section
    }

    // 查找指定深度下最新的scene
    fn latest_scene(&self, depth: usize) -> Option<StructToken> {
        // 查找第一层中最后一个符合条件的token
        if depth <= 0 {
            return None;
        }

        // 找到第一层的scene
        let mut current_scene = None;
        for item in self.result.properties.structure.iter().rev() {
            if item.isscene {
                current_scene = Some(item.clone());
                break;
            }
        }

        if depth == 1 || current_scene.is_none() {
            return current_scene;
        }

        // 迭代查找更深层次的scene
        let mut current_depth = 1;
        while current_depth < depth {
            let scene = current_scene.clone().unwrap();

            // 在子节点中查找scene
            let mut found_child = false;
            for child in scene.children.iter().rev() {
                if child.isscene {
                    current_scene = Some(child.clone());
                    found_child = true;
                    break;
                }
            }

            if !found_child {
                break;
            }

            current_depth += 1;
        }

        current_scene
    }

    // 处理内联注释并将它们添加到结构树中
    fn process_inline_notes(&mut self) {
        if !self.current_outline_note_text.is_empty() {
            for i in 0..self.current_outline_note_text.len() {
                if !self.current_outline_note_text[i].trim().is_empty() {
                    let line_number = self.current_outline_note_linenum[i];
                    let note_text = self.current_outline_note_text[i].trim().to_string();

                    let struct_token = StructToken {
                        text: note_text.clone(),
                        isnote: true,
                        id: Some(format!("/{}", line_number)),
                        isscene: false,
                        ischartor: false,
                        dialogue_end_line: 0,
                        duration_sec: 0.0,
                        children: Vec::new(),
                        level: 0,
                        notes: Vec::new(),
                        range: Some(Range {
                            start: Position { line: line_number, character: 0 },
                            end: Position { line: line_number, character: note_text.len() + 4 },
                        }),
                        section: false,
                        synopses: Vec::new(),
                        play_sec: self.play_time_sec,
                        structs: Vec::new(),
                        duration: 0.0,
                    };

                    self.result.properties.structure.push(struct_token);
                }
            }
        }
    }

    // 添加对话编号装饰
    fn add_dialogue_number_decoration(&self, _token: &mut ScriptToken) {
        // 在Rust实现中可以根据需要添加具体实现
    }

    // 处理注释和注解
    fn process_comments_and_notes(&mut self, parts: Vec<&str>, line_num: usize, cfg: &Conf) {
        for part in parts {
            if !part.is_empty() {
                if part == "/*" {
                    if self.nested_notes == 0 {
                        self.nested_comments += 1;
                    } else {
                        // 是 note 的注解内容
                        self.add_outline_note("/*", line_num);
                        if cfg.print_notes {
                            self.text_display.push_str(part);
                        }
                    }
                } else if part == "*/" {
                    if self.nested_comments > 0 {
                        self.nested_comments -= 1;
                    } else {
                        if self.nested_notes == 0 {
                            // 既不是 note 也不是 comment
                            self.text_display.push_str(part);
                            self.text_valid.push_str(part);
                        } else {
                            // 是 note 的注解内容
                            self.add_outline_note("*/", line_num);
                            if cfg.print_notes {
                                self.text_display.push_str(part);
                            }
                        }
                    }
                } else if part == "[[" || part == "[[|" {
                    if self.nested_comments == 0 {
                        self.nested_notes += 1;
                        if self.nested_notes == 1 {
                            // 需要处理大纲注解
                            self.current_outline_note_text.push(String::new());
                            self.current_outline_note_linenum.push(line_num);
                            if cfg.print_notes {
                                if part == "[[|" {
                                    // 扩展 note 语法。可强制个别注解不在底部而在原位置打印。
                                    self.text_display.push_str(&format!("{}[", FountainConstants::style_chars()["note_begin_ext"]));
                                } else {
                                    self.text_display.push_str(&format!("{}[", FountainConstants::style_chars()["note_begin"]));
                                }
                            }
                        } else {
                            self.add_outline_note("[", line_num);
                            if cfg.print_notes {
                                self.text_display.push_str("["); // 嵌套的里层开口
                            }
                        }
                    } else {
                        // 是 comment 的注解内容
                    }
                } else if part == "]]" {
                    if self.nested_notes > 0 {
                        self.nested_notes -= 1;
                        if self.nested_notes == 0 {
                            if cfg.print_notes {
                                self.text_display.push_str(&format!("]{}",
                                    FountainConstants::style_chars()["note_end"])); // 闭口，转换成特殊样式字符。
                            }
                        } else {
                            self.add_outline_note("]", line_num);
                            if cfg.print_notes {
                                self.text_display.push_str("]"); // 嵌套的里层闭口
                            }
                        }
                    } else {
                        if self.nested_comments == 0 {
                            // 既不是 comment 也不是 note
                            self.text_display.push_str(part);
                            self.text_valid.push_str(part);
                        } else {
                            // 是 comment 的注解内容
                        }
                    }
                } else {
                    // 非符号文字内容
                    if self.nested_comments > 0 {
                        // 注释内容，忽略
                    } else if self.nested_notes > 0 {
                        self.add_outline_note(if !part.is_empty() { part } else { "" }, line_num);
                        if cfg.print_notes {
                            self.text_display.push_str(part);
                        }
                    } else {
                        self.text_display.push_str(part);
                        self.text_valid.push_str(part);
                    }
                }
            }
        }
    }

    // 添加大纲注解
    fn add_outline_note(&mut self, note: &str, line: usize) {
        if !self.current_outline_note_text.is_empty() {
            let last_index = self.current_outline_note_text.len() - 1;
            self.current_outline_note_text[last_index].push_str(note);
            if self.current_outline_note_text[last_index].trim().is_empty() {
                self.current_outline_note_linenum[last_index] = line;
            }
        }
    }

    // 更新前一个场景的长度
    fn update_previous_scene_length(&mut self) {
        let action = self.result.length_action - self.length_action_so_far;
        let dialogue = self.result.length_dialogue - self.length_dialogue_so_far;
        // println!("DEBUG: 更新场景长度 - 动作: {} -> {}, 对话: {} -> {}",
        //          self.length_action_so_far, self.result.length_action,
        //          self.length_dialogue_so_far, self.result.length_dialogue);
        self.length_action_so_far = self.result.length_action;
        self.length_dialogue_so_far = self.result.length_dialogue;

        if !self.result.properties.scenes.is_empty() {
            let last_index = self.result.properties.scenes.len() - 1;
            if let Some(scene) = self.result.properties.scenes.get_mut(last_index) {
                scene.insert("actionLength".to_string(), serde_json::to_value(action).unwrap());
                scene.insert("dialogueLength".to_string(), serde_json::to_value(dialogue).unwrap());
                scene.insert("endPlaySec".to_string(), serde_json::to_value(self.play_time_sec).unwrap());
            }
        }
    }

    /// 解析Fountain格式文本
    ///
    /// # Arguments
    ///
    /// * `script` - Fountain格式的剧本文本
    /// * `cfg` - 配置对象
    /// * `generate_html` - 是否生成HTML输出
    ///
    /// # Returns
    ///
    /// 解析结果对象
    pub fn parse(&mut self, script: &str, cfg: &Conf, generate_html: bool) -> ParseOutput {
        // 初始化解析结果
        self.result = ParseOutput::new();
        if script.is_empty() {
            return self.result.clone();
        }

        // 从配置中初始化时长计算参数
        self.result.dial_sec_per_char = cfg.dial_sec_per_char;
        self.result.dial_sec_per_punc_short = cfg.dial_sec_per_punc_short;
        self.result.dial_sec_per_punc_long = cfg.dial_sec_per_punc_long;
        self.result.action_sec_per_char = cfg.action_sec_per_char;

        // 记录开始时间
        self.result.parse_time = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis() as u64;

        // 处理换行符差异
        self.new_line_length = if script.contains("\r\n") { 2 } else { 1 };
        let lines: Vec<&str> = script.split(&['\r', '\n'][..]).collect();

        // 解析状态跟踪
        self.result.state = "normal".to_string(); // normal, title, dialogue
        self.scene_number = 1;
        self.nested_comments = 0;
        self.nested_notes = 0;
        let mut _need_process_outline_note = 0;
        self.current_outline_note_text.clear();
        self.current_outline_note_linenum.clear();
        self.text_display = String::new();
        self.text_valid = String::new();
        let mut ignored_last_token = false;
        self.last_chartor_structure_token = None;
        let mut last_title_page_token: Option<ScriptToken> = None;
        self.last_scen_structure_token = None;
        self.last_scen_structure_token_pre = None;
        let mut last_character_index = 0;
        let mut _previous_character: Option<String> = None;

        let mut parenthetical_open = false;
        self.result.dual_str = None;
        let mut font_title = false;
        let mut last_is_blank_title = false; // 上一个行是否是空行
        self.play_time_sec = 0.0;
        let mut last_was_separator = false;

        // 镜头交切处理
        self.shot_cut = 0;
        self.shot_cut_strct_tokens.clear();
        let mut dup_scence_nuber: HashMap<String, String> = HashMap::new();
        let mut scence_numbers: HashSet<String> = HashSet::new();

        self.length_action_so_far = 0.0;
        self.length_dialogue_so_far = 0.0;
        self.force_not_dual = true;

        let mut _empty_title_page = true;

        // 主解析循环
        let mut is_block_inner = false;
        let mut note_token: Option<ScriptToken> = None; // page_break 的打印注解

        for i in 0..lines.len() {
            let text = lines[i];

            self.text_display = String::new(); // 视乎打印设置是否打印note，可以包含 note 内容
            self.text_valid = String::new(); // 去除注解和note后的有效内容，用来判定文本内容性质

            let mut empty_break_line = false;
            let mut is_block_end_empty_line = false; // 连续块后紧接着的断块空行
            let mut is_block_begin_line = false; // 连续块后紧接着的断块空行

            let match_line_break = text.trim().is_empty() && text.len() > 1;

            // 处理空行
            if text.trim().is_empty() {
                // 注解去除前已经是空行了
                self.text_display = text.to_string();
                if self.nested_comments > 0 || self.nested_notes > 0 {
                    // 如果是在注释中，那么直接忽略
                    if self.nested_notes > 0 && cfg.print_notes && match_line_break {
                        // note 空行，双空格表示保留一个空行，否则直接去掉空行
                    } else {
                        continue;
                    }
                } else {
                    if !is_block_inner {
                        // 至少一个空行后的，再空行。且不在注解中的空行
                        // 插入一个 action 空行
                        empty_break_line = true;
                    } else {
                        // 非空行后的紧接着的空行
                        if match_line_break && self.result.state != "normal" {
                            // 区分情况，title page 和 dialogue 里面的双空格空行，转成换行，还归为块内内容
                        } else {
                            is_block_inner = false;
                            is_block_end_empty_line = true;
                        }
                    }
                }
            } else {
                // 至少不是空行了

                // 分割注释和注解
                let re = Regex::new(r"(\/\*|\*\/|\[\[\||\[\[|\]\])").unwrap();
                let mut parts = Vec::new();
                let mut last_end = 0;

                for cap in re.captures_iter(text) {
                    let m = cap.get(0).unwrap();
                    if m.start() > last_end {
                        parts.push(&text[last_end..m.start()]);
                    }
                    parts.push(m.as_str());
                    last_end = m.end();
                }

                if last_end < text.len() {
                    parts.push(&text[last_end..]);
                }

                parts.retain(|part| !part.is_empty());
                self.process_comments_and_notes(parts, i, cfg);

                if self.text_valid.trim().is_empty() {
                    // 纯注解行内容，有内容，且全部被注解了
                    if self.text_display.trim().is_empty() && self.text_display.len() <= 1 {
                        // 关闭了打印note的注解开头，或跨行的note；且空格小于2
                        continue;
                    }
                } else {
                    if !is_block_inner {
                        is_block_inner = true;
                        is_block_begin_line = true;
                    }
                }
            }

            let mut this_token = ScriptToken {
                token_type: String::new(),
                text: self.text_display.clone(),
                line: i,
                start: 0,
                end: text.len(),
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
            };
            this_token.play_time_sec = self.play_time_sec;

            // 处理各种类型的行

            // 首先处理空行
            if self.text_display.trim().is_empty() {
                // 处理空行
                if empty_break_line {
                    // 空行后的空行
                    let skip_separator = (cfg.merge_empty_lines && last_was_separator) ||
                        (ignored_last_token &&
                         self.result.tokens.len() > 1 &&
                         self.result.tokens[self.result.tokens.len() - 1].token_type == "separator");

                    if skip_separator {
                        continue;
                    }

                    this_token.token_type = "separator".to_string();
                    self.push_token(this_token);
                    last_was_separator = true;
                } else if is_block_end_empty_line {
                    // 块的结束，处理状态
                    if self.result.state == "dialogue" {
                        if let Some(last_chartor) = &mut self.last_chartor_structure_token {
                            last_chartor.dialogue_end_line = i - 1;
                        }
                        parenthetical_open = false;

                        let dialogue_end = ScriptToken {
                            token_type: "dialogue_end".to_string(),
                            text: self.text_display.clone(),
                            line: i,
                            start: 0,
                            end: text.len(),
                            ..ScriptToken::empty()
                        };
                        self.result.tokens.push(dialogue_end);
                    }

                    if self.result.state == "dual_dialogue" {
                        if let Some(last_chartor) = &mut self.last_chartor_structure_token {
                            last_chartor.dialogue_end_line = i - 1;
                        }
                        parenthetical_open = false;

                        let dual_dialogue_end = ScriptToken {
                            token_type: "dual_dialogue_end".to_string(),
                            text: self.text_display.clone(),
                            line: i,
                            start: 0,
                            end: text.len(),
                            ..ScriptToken::empty()
                        };
                        self.result.tokens.push(dual_dialogue_end);
                    }

                    self.result.dual_str = None;
                    self.result.state = "normal".to_string();

                    this_token.token_type = "separator".to_string();
                    this_token.text = FountainConstants::style_chars()["style_global_clean"].to_string();
                    self.result.tokens.push(this_token);
                    last_was_separator = true;
                } else {
                    // note里的延续空行或者延续块内容的块内空行
                    if !self.result.tokens.is_empty() {
                        if self.result.state == "title_page" {
                            if font_title {
                                continue;
                            }

                            let mut merge = false;
                            if cfg.merge_empty_lines {
                                if let Some(last_title) = &last_title_page_token {
                                    // 检查是否以换行符加空格结尾（替代前瞻断言）
                                    if last_title.text.ends_with('\n') && last_title.text.chars().rev().skip(1).take_while(|&c| c == ' ').count() > 0 {
                                        merge = true;
                                    }
                                }
                            }

                            if !merge && last_title_page_token.is_some() {
                                last_title_page_token.as_mut().unwrap().text.push('\n');
                            }
                        } else {
                            let last_index = self.result.tokens.len() - 1;
                            let last_token = &self.result.tokens[last_index];

                            let mut merge = false;
                            if cfg.merge_empty_lines && last_token.text.trim().is_empty() {
                                merge = true;
                            }

                            if !merge {
                                if last_token.token_type == "character" {
                                    this_token.token_type = "dialogue".to_string();
                                    this_token.dual = last_token.dual.clone();
                                } else if last_token.token_type == "parenthetical" {
                                    this_token.token_type = "parenthetical".to_string();
                                    this_token.dual = last_token.dual.clone();
                                } else if last_token.token_type == "dialogue" {
                                    this_token.token_type = "dialogue".to_string();
                                    this_token.dual = last_token.dual.clone();
                                } else {
                                    this_token.token_type = "action".to_string();
                                }
                                self.result.tokens.push(this_token);
                            }
                        }
                    }
                }
                continue;
            }

            // 检查是否是块开始行并且是标题页
            if is_block_begin_line && self.regex.get("title_page").unwrap().is_match(&self.text_valid) {
                self.result.state = "title_page".to_string();
            }

            // 处理标题页
            if self.result.state == "title_page" {
                // 检查是否遇到了非标题页内容，如果是则结束标题页状态
                if is_block_begin_line && !self.regex.get("title_page").unwrap().is_match(&self.text_valid) {
                    // 检查是否是标题页字段的延续行
                    if last_title_page_token.is_some() && !self.text_valid.trim().is_empty() {
                        // 这可能是标题页字段的延续行，不结束标题页状态
                        // 继续处理为标题页延续内容，跳过结束标题页的逻辑
                    } else {
                        // 遇到非标题页内容，结束标题页状态
                        self.process_title_page_end(i);
                        // 继续处理当前行，不要跳过
                    }
                }

                if self.regex.get("title_page").unwrap().is_match(&self.text_valid) {
                    self.text_valid = self.text_valid.trim().to_string();
                    let index = self.text_valid.find(':').unwrap_or(0);
                    this_token.token_type = self.text_valid[..index].to_lowercase().replace(' ', "_");

                    let font_mt = Regex::new(r"(?i)^\s*(font|font italic|font bold|font bold italic|metadata)\:(.*)").unwrap()
                        .captures(&self.text_valid);

                    if let Some(captures) = font_mt {
                        font_title = true;
                        this_token.text = captures.get(2).unwrap().as_str().trim().to_string();
                    } else {
                        font_title = false;
                        let mt = Regex::new(r"(?i)^(.*?↻)??\s*(title|credit|author[s]?|source|notes|draft date|date|watermark|contact(?: info)?|revision|copyright|tl|tc|tr|cc|br|bl|header|footer)\:(.*)").unwrap()
                            .captures(&self.text_display);

                        if let Some(captures) = mt {
                            this_token.text = captures.get(3).unwrap().as_str().trim().to_string();
                            process_token_text_style_char(&mut this_token);
                            this_token.text = format!("{}{}", FountainConstants::style_chars()["style_global_clean"], this_token.text);
                            last_is_blank_title = false;
                        }
                    }

                    last_title_page_token = Some(this_token.clone());

                    // 根据key将内容添加到正确的title page区域
                    if let Some(position) = self.get_title_page_position(&this_token.token_type) {
                        this_token.index = position.index;

                        // 克隆position.position以避免借用冲突
                        let position_key = position.position.clone();

                        if !self.result.title_page.contains_key(&position_key) {
                            self.result.title_page.insert(position_key.clone(), Vec::new());
                        }

                        if let Some(tokens) = self.result.title_page.get_mut(&position_key) {
                            tokens.push(this_token.clone());
                        }

                        _empty_title_page = false;
                    }

                    if !self.result.properties.title_keys.contains(&this_token.token_type) {
                        self.result.properties.title_keys.push(this_token.token_type.clone());
                    }

                    continue;
                }

                // 处理标题页字段的延续内容或其他标题页状态下的内容
                if self.result.state == "title_page" && !self.regex.get("title_page").unwrap().is_match(&self.text_valid) {
                    // 标题页字段内容的换行内容，或者标题页状态下的其他内容
                    if font_title {
                        this_token.text = self.text_valid.trim().to_string();
                        if !this_token.text.is_empty() && last_title_page_token.is_some() {
                            let last_text = &last_title_page_token.as_ref().unwrap().text;
                            let separator = if last_text.is_empty() { "" } else { " " };
                            let new_text = format!("{}{}{}", last_text, separator, this_token.text.trim());

                            // 更新 last_title_page_token
                            last_title_page_token.as_mut().unwrap().text = new_text.clone();

                            // 同时更新 title_page 中对应的 token
                            let token_type = &last_title_page_token.as_ref().unwrap().token_type;
                            if let Some(position) = self.get_title_page_position(token_type) {
                                let position_key = position.position.clone();
                                if let Some(tokens) = self.result.title_page.get_mut(&position_key) {
                                    // 找到最后一个相同类型的 token 并更新其文本
                                    for token in tokens.iter_mut().rev() {
                                        if token.token_type == *token_type {
                                            token.text = new_text;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    } else {
                        this_token.text = self.text_display.trim().to_string();
                        process_token_text_style_char(&mut this_token);

                        let mut handled = false;
                        if cfg.merge_empty_lines {
                            let curr_blank = is_blank_line_after_style(&this_token.text);
                            if curr_blank && last_is_blank_title && last_title_page_token.is_some() {
                                handled = true;
                                let t = Regex::new(&format!(r"[{}]", regex::escape(FountainConstants::style_chars()["all"])))
                                    .unwrap()
                                    .replace_all(&this_token.text, "")
                                    .to_string();
                                last_title_page_token.as_mut().unwrap().text.push_str(&t);
                            }
                            last_is_blank_title = curr_blank;
                        }

                        if !handled && last_title_page_token.is_some() {
                            let separator = if !is_blank_line_after_style(&last_title_page_token.as_ref().unwrap().text) {
                                "\n"
                            } else {
                                ""
                            };
                            last_title_page_token.as_mut().unwrap().text.push_str(&format!("{}{}", separator, this_token.text.trim()));

                            // 同时更新 title_page 中对应的 token
                            let token_type = &last_title_page_token.as_ref().unwrap().token_type;
                            if let Some(position) = self.get_title_page_position(token_type) {
                                let position_key = position.position.clone();
                                if let Some(tokens) = self.result.title_page.get_mut(&position_key) {
                                    // 找到最后一个相同类型的 token 并更新其文本
                                    for token in tokens.iter_mut().rev() {
                                        if token.token_type == *token_type {
                                            token.text = last_title_page_token.as_ref().unwrap().text.clone();
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    continue;
                }
            }

            // 处理对话状态下的行
            if self.result.state == "dialogue" || self.result.state == "dual_dialogue" {
                if parenthetical_open {
                    // 如果括号已经打开，强制当前行为括号内容（与Flutter版本保持一致）
                    this_token.token_type = "parenthetical".to_string();
                    this_token.dual = self.result.dual_str.clone();
                    this_token.text = self.text_display.clone();
                    process_token_text_style_char(&mut this_token);

                    // 为括号内容添加包装符号（与Flutter版本保持一致）
                    this_token.text = format!("{}{}{}",
                        FountainConstants::style_chars()["italic_global_begin"],
                        this_token.text,
                        FountainConstants::style_chars()["italic_global_end"]
                    );

                    // 检查是否是括号结束
                    if self.regex.get("parenthetical_end").unwrap().is_match(&self.text_valid) {
                        parenthetical_open = false;
                    }

                    self.push_token(this_token);
                    continue;
                } else if self.regex.get("parenthetical").unwrap().is_match(&self.text_valid) {
                    // 处理单行括号内容
                    this_token.token_type = "parenthetical".to_string();
                    this_token.dual = self.result.dual_str.clone();
                    this_token.text = self.text_display.clone();
                    process_token_text_style_char(&mut this_token);

                    // 为括号内容添加包装符号（与Flutter版本保持一致）
                    this_token.text = format!("{}{}{}",
                        FountainConstants::style_chars()["italic_global_begin"],
                        this_token.text,
                        FountainConstants::style_chars()["italic_global_end"]
                    );

                    self.push_token(this_token);
                    continue;
                } else if self.regex.get("parenthetical_start").unwrap().is_match(&self.text_valid) {
                    // 处理多行括号开始
                    this_token.token_type = "parenthetical".to_string();
                    this_token.dual = self.result.dual_str.clone();
                    this_token.text = self.text_display.clone();
                    process_token_text_style_char(&mut this_token);

                    // 为括号内容添加包装符号（与Flutter版本保持一致）
                    this_token.text = format!("{}{}{}",
                        FountainConstants::style_chars()["italic_global_begin"],
                        this_token.text,
                        FountainConstants::style_chars()["italic_global_end"]
                    );
                    parenthetical_open = true;

                    self.push_token(this_token);
                    continue;
                } else {
                    // 处理对话内容
                    this_token.token_type = "dialogue".to_string();
                    this_token.dual = self.result.dual_str.clone();
                    this_token.text = self.text_display.clone();
                    process_token_text_style_char(&mut this_token);

                    // 为对话添加包装符号（与Flutter版本保持一致）
                    this_token.text = format!("{}{}{}",
                        FountainConstants::style_chars()["italic_global_begin"],
                        this_token.text,
                        FountainConstants::style_chars()["italic_global_end"]
                    );

                    // 处理对话持续时间
                    this_token = self.process_dialogue_block(this_token);

                    self.push_token(this_token);
                    continue;
                }
            }

            // 处理正常状态下的行
            if self.result.state == "normal" {
                let mut action = false;
                if is_block_begin_line {

                    // 检查是否是场景标题
                    if self.regex.get("scene_heading").unwrap().is_match(&self.text_valid) {
                        self.process_title_page_end(i);

                        if self.result.properties.first_scene_line.is_none() || self.result.properties.first_scene_line == Some(usize::MAX) {
                            self.result.properties.first_scene_line = Some(i);

                            // 从metadata里更新用户配置时间预估参数（与TypeScript版本保持一致）
                            if let Some(hidden_tokens) = self.result.title_page.get("hidden") {
                                for token in hidden_tokens {
                                    if token.token_type == "metadata" {
                                        // 去掉 "Metadata: " 前缀，只保留JSON部分
                                        let json_text = if token.text.starts_with("Metadata: ") {
                                            &token.text[10..] // 跳过 "Metadata: " 前缀
                                        } else {
                                            &token.text
                                        };
                                        if let Ok(metadata) = serde_json::from_str::<serde_json::Value>(json_text) {
                                            if let Some(dial_sec_per_char) = metadata.get("dial_sec_per_char").and_then(|v| v.as_f64()) {
                                                self.result.dial_sec_per_char = dial_sec_per_char;
                                            }
                                            if let Some(dial_sec_per_punc_short) = metadata.get("dial_sec_per_punc_short").and_then(|v| v.as_f64()) {
                                                self.result.dial_sec_per_punc_short = dial_sec_per_punc_short;
                                            }
                                            if let Some(dial_sec_per_punc_long) = metadata.get("dial_sec_per_punc_long").and_then(|v| v.as_f64()) {
                                                self.result.dial_sec_per_punc_long = dial_sec_per_punc_long;
                                            }
                                            if let Some(action_sec_per_char) = metadata.get("action_sec_per_char").and_then(|v| v.as_f64()) {
                                                self.result.action_sec_per_char = action_sec_per_char;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                        }

                        self.force_not_dual = true;
                        // 去掉前面的点号
                        self.text_display = Regex::new(r"^[ \t]*\.").unwrap()
                            .replace(&self.text_display, "").to_string();

                        // 如果配置要求每个场景新页面，且不是第一个场景，添加分页符
                        if cfg.each_scene_on_new_page && self.scene_number != 1 {
                            let page_break = self.create_token(Some(""), Some(0), Some(i), Some(text.len()), "page_break");
                            self.push_token(page_break);
                        }

                        this_token.token_type = "scene_heading".to_string();
                        let mut nb = self.scene_number.to_string();
                        let mut scene_number_dup = false;

                        // 处理场景编号
                        let mut text_for_token = Regex::new(r"^[ \t]*\.").unwrap()
                            .replace(&self.text_valid, "").to_string();

                        if let Some(scene_num_match) = self.regex.get("scene_number").unwrap().captures(&self.text_valid) {
                            text_for_token = self.regex.get("scene_number").unwrap()
                                .replace(&text_for_token, "").to_string();
                            self.text_display = self.regex.get("scene_number").unwrap()
                                .replace(&self.text_display, "").to_string();

                            // 处理场景编号
                            if let Some(group2) = scene_num_match.get(2) {
                                if !group2.as_str().trim().is_empty() {
                                    nb = group2.as_str().trim().to_string();
                                }
                            }

                            if let Some(group1) = scene_num_match.get(1) {
                                if !group1.as_str().trim().is_empty() {
                                    let tab = group1.as_str().trim().to_string();
                                    if dup_scence_nuber.contains_key(&tab) {
                                        nb = dup_scence_nuber.get(&tab).unwrap().clone();
                                    } else {
                                        dup_scence_nuber.insert(tab, nb.clone());
                                    }
                                }
                            }
                        }

                        // 检查场景编号是否重复
                        if scence_numbers.contains(&nb) {
                            scene_number_dup = true;
                        } else {
                            scence_numbers.insert(nb.clone());
                        }

                        // 设置场景编号
                        if scene_number_dup {
                            this_token.number = Some(format!("↑{}", nb)); // 用 ↑ 打印标记提示重复
                        } else {
                            this_token.number = Some(nb.clone());
                        }

                        // 标准化场景标题格式
                        let mut idx = text_for_token.find('-');
                        if let Some(pos) = idx {
                            text_for_token = format!("{} - {}",
                                &text_for_token[..pos],
                                &text_for_token[pos+1..]);
                        }

                        text_for_token = text_for_token.to_uppercase()
                            .replace(|c: char| c.is_whitespace(), " ");

                        idx = self.text_display.find('-');
                        if let Some(pos) = idx {
                            self.text_display = format!("{} - {}",
                                &self.text_display[..pos],
                                &self.text_display[pos+1..]);
                        }

                        self.text_display = self.text_display.to_uppercase()
                            .replace(|c: char| c.is_whitespace(), " ");

                        // 规范化空格：将多个连续空格替换为单个空格，并trim两端
                        self.text_display = Regex::new(r"\s+").unwrap()
                            .replace_all(&self.text_display, " ")
                            .trim()
                            .to_string();

                        // 去掉场景标题前面的点号（与Flutter版本保持一致）
                        this_token.text = self.text_display.clone();
                        if this_token.text.trim_start().starts_with('.') {
                            this_token.text = this_token.text.trim_start().chars().skip(1).collect::<String>().trim_start().to_string();
                        }
                        this_token.text_no_notes = Some(text_for_token.clone());

                        // 创建结构树节点
                        let mut cobj = StructToken {
                            text: format!("{} {}", this_token.number.clone().unwrap_or_default(), text_for_token.clone()),
                            children: Vec::new(),
                            level: 0,
                            section: false,
                            synopses: Vec::new(),
                            notes: Vec::new(),
                            isscene: true,
                            ischartor: false,
                            dialogue_end_line: 0,
                            duration_sec: 0.0,
                            range: Some(Range {
                                start: Position { line: this_token.line, character: 0 },
                                end: Position { line: this_token.line, character: self.text_valid.len() },
                            }),
                            id: None,
                            isnote: false,
                            play_sec: self.play_time_sec,
                            structs: Vec::new(),
                            duration: 0.0,
                        };

                        // 设置结构树节点ID
                        if self.current_depth == 0 {
                            cobj.id = Some(format!("/{}", this_token.line));
                            self.result.properties.structure.push(cobj.clone());
                        } else {
                            if let Some(level) = self.latest_section(self.current_depth) {
                                cobj.id = Some(format!("{}/{}", level.id.as_ref().unwrap_or(&String::new()), this_token.line));

                                // 找到父节点并添加子节点
                                for parent in &mut self.result.properties.structure {
                                    if parent.id == level.id {
                                        parent.children.push(cobj.clone());
                                        break;
                                    }
                                }
                            } else {
                                cobj.id = Some(format!("/{}", this_token.line));
                                self.result.properties.structure.push(cobj.clone());
                            }
                        }

                        self.last_scen_structure_token_pre = self.last_scen_structure_token.clone();
                        self.last_scen_structure_token = Some(cobj.clone());

                        if self.shot_cut > 0 {
                            if let Some(last_map) = self.shot_cut_strct_tokens.last_mut() {
                                if let Some(structs) = last_map.get_mut("structs") {
                                    if let Some(structs_array) = structs.as_array_mut() {
                                        structs_array.push(serde_json::to_value(cobj.clone()).unwrap());
                                    }
                                }
                            }
                        }

                        // 更新前一个场景的长度
                        self.update_previous_scene_length();

                        if !self.result.properties.scenes.is_empty() {
                            let last_index = self.result.properties.scenes.len() - 1;
                            if let Some(scene) = self.result.properties.scenes.get_mut(last_index) {
                                scene.insert("endPlaySec".to_string(), serde_json::to_value(self.play_time_sec).unwrap());
                            }
                        }

                        // 添加新场景
                        let mut scene_map = HashMap::new();
                        scene_map.insert("scene".to_string(), serde_json::to_value(nb.clone()).unwrap());
                        scene_map.insert("text".to_string(), serde_json::to_value(text_for_token.clone()).unwrap());
                        scene_map.insert("line".to_string(), serde_json::to_value(this_token.line).unwrap());
                        scene_map.insert("actionLength".to_string(), serde_json::to_value(0).unwrap());
                        scene_map.insert("dialogueLength".to_string(), serde_json::to_value(0).unwrap());
                        scene_map.insert("number".to_string(), serde_json::to_value(nb.clone()).unwrap());
                        scene_map.insert("startPlaySec".to_string(), serde_json::to_value(self.play_time_sec).unwrap());
                        scene_map.insert("endPlaySec".to_string(), serde_json::to_value(self.play_time_sec).unwrap());

                        // println!("DEBUG: 添加场景 {} (行号: {})", nb, this_token.line);
                    self.result.properties.scenes.push(scene_map);
                        self.result.properties.scene_lines.push(this_token.line);
                        self.result.properties.scene_names.push(self.text_valid.clone());

                        // 处理场景位置信息
                        if let Some(location) = self.parse_location_information(&self.text_valid) {
                            let location_slug = self.slugify(&location.name);
                            let lslugs: Vec<String> = vec![location_slug.clone()]
                                .into_iter()
                                .map(|it| it.trim().to_string())
                                .filter(|it| !it.is_empty())
                                .collect();

                            for sl in lslugs {
                                let mut loc = location.clone();
                                loc.scene_number = nb.clone();
                                loc.line = this_token.line;
                                loc.start_play_sec = self.play_time_sec;

                                if let Some(locations) = self.result.properties.locations.get_mut(&sl) {
                                    if !locations.iter().any(|it| it.line == this_token.line) {
                                        locations.push(loc);
                                    }
                                } else {
                                    self.result.properties.locations.insert(sl, vec![loc]);
                                }
                            }
                        }

                        // 更新场景编号
                        if !scene_number_dup {
                            self.scene_number += 1;
                        }

                        self.push_token(this_token);
                        continue;
                    } else if self.regex.get("centered").unwrap().is_match(&self.text_valid) {
                        // 处理居中文本
                        action = true;
                    } else if self.regex.get("transition").unwrap().is_match(&self.text_valid) {
                        // 处理转场
                        // 处理镜头交切标志
                        let match_display = self.regex.get("transition").unwrap().captures(&self.text_valid);

                        if let Some(captures) = match_display {
                            if captures.len() > 2 && captures.get(2).is_some() {
                                let tx = captures.get(2).unwrap().as_str().trim();

                                if tx.starts_with("{+") && tx.ends_with("+} ↓") {
                                    self.shot_cut = 1;
                                    let mut shot_cut_map = HashMap::new();
                                    shot_cut_map.insert("duration".to_string(), serde_json::to_value(0.0).unwrap());

                                    let mut structs = Vec::new();
                                    if let Some(last_scen) = &self.last_scen_structure_token {
                                        structs.push(serde_json::to_value(last_scen).unwrap());
                                    }

                                    shot_cut_map.insert("structs".to_string(), serde_json::to_value(structs).unwrap());
                                    self.shot_cut_strct_tokens.push(shot_cut_map);
                                } else if tx.starts_with("{#") && tx.ends_with("#} ↓") {
                                    self.shot_cut = 2;
                                    let mut shot_cut_map = HashMap::new();
                                    shot_cut_map.insert("duration".to_string(), serde_json::to_value(0.0).unwrap());

                                    let mut structs = Vec::new();
                                    if let Some(last_scen_pre) = &self.last_scen_structure_token_pre {
                                        structs.push(serde_json::to_value(last_scen_pre).unwrap());
                                    }

                                    if let Some(last_scen) = &self.last_scen_structure_token {
                                        structs.push(serde_json::to_value(last_scen).unwrap());
                                    }

                                    shot_cut_map.insert("structs".to_string(), serde_json::to_value(structs).unwrap());
                                    self.shot_cut_strct_tokens.push(shot_cut_map);
                                } else if tx.starts_with("{=") && tx.ends_with("=} ↓") {
                                    self.shot_cut = 3;
                                    let mut shot_cut_map = HashMap::new();
                                    shot_cut_map.insert("duration".to_string(), serde_json::to_value(0.0).unwrap());
                                    shot_cut_map.insert("structs".to_string(), serde_json::to_value(Vec::<StructToken>::new()).unwrap());
                                    self.shot_cut_strct_tokens.push(shot_cut_map);
                                } else if tx.starts_with("{-") && tx.ends_with("-} ↑") {
                                    self.shot_cut = 0;
                                }
                            }
                        }

                        self.process_title_page_end(i);
                        this_token.text = Regex::new(r"^\s*>\s*").unwrap()
                            .replace(&self.text_display, "").to_string();

                        process_token_text_style_char(&mut this_token);
                        this_token.token_type = "transition".to_string();

                        self.push_token(this_token);
                        continue;
                    } else if self.regex.get("character").unwrap().is_match(&self.text_valid) {
                        // 处理角色
                        self.process_title_page_end(i);
                        self.result.state = "dialogue".to_string();
                        this_token.token_type = "character".to_string();
                        this_token.take_number = Some(self.take_count as i32);
                        self.take_count += 1;

                        // 处理角色名
                        let mut text_valid = self.trim_character_force_symbol(&self.text_valid);

                        // 检查是否是双对话
                        if text_valid.ends_with("^") {
                            if cfg.use_dual_dialogue && !self.force_not_dual {
                                self.result.state = "dual_dialogue".to_string();

                                // 更新上一个对话为dual:left
                                let dialogue_tokens = ["dialogue", "character", "parenthetical"];
                                let mut mod_last = false;
                                let mut last_character_index = last_character_index;

                                while last_character_index < self.result.tokens.len() &&
                                      dialogue_tokens.contains(&self.result.tokens[last_character_index].token_type.as_str()) {
                                    let old_last_dual = self.result.tokens[last_character_index].dual.clone();

                                    if old_last_dual.is_none() || old_last_dual.as_ref().unwrap().is_empty() {
                                        mod_last = true;
                                        self.result.tokens[last_character_index].dual = Some("left".to_string());
                                        self.result.dual_str = Some("right".to_string());
                                    } else if old_last_dual.as_ref().unwrap() == "left" {
                                        self.result.dual_str = Some("right".to_string());
                                    } else {
                                        self.result.dual_str = Some("left".to_string());
                                    }

                                    last_character_index += 1;
                                }

                                if mod_last {
                                    // 更新上一个dialogue_begin为dual_dialogue_begin，并移除最后的dialogue_end
                                    let mut found_match = false;
                                    let mut temp_index = self.result.tokens.len();

                                    while !found_match && temp_index > 0 {
                                        temp_index -= 1;
                                        match self.result.tokens[temp_index].token_type.as_str() {
                                            "dialogue_end" => {
                                                self.result.tokens.truncate(temp_index);
                                                temp_index -= 1;
                                            },
                                            "separator" | "character" | "dialogue" | "parenthetical" => {},
                                            "dialogue_begin" => {
                                                self.result.tokens[temp_index].token_type = "dual_dialogue_begin".to_string();
                                                found_match = true;
                                            },
                                            _ => found_match = true
                                        }
                                    }
                                }

                                if self.result.dual_str.as_ref().unwrap() == "left" {
                                    self.push_token(self.create_token(None, None, None, None, "dual_dialogue_begin"));
                                } else {
                                    // 删除之前left的dual_dialogue_end
                                    let mut found_match = false;
                                    let mut temp_index = self.result.tokens.len();

                                    while !found_match && temp_index > 0 {
                                        temp_index -= 1;
                                        match self.result.tokens[temp_index].token_type.as_str() {
                                            "dual_dialogue_end" => {
                                                self.result.tokens.truncate(temp_index);
                                                temp_index -= 1;
                                            },
                                            "separator" | "character" | "dialogue" | "parenthetical" => {},
                                            _ => found_match = true
                                        }
                                    }
                                }

                                this_token.dual = self.result.dual_str.clone();
                            } else {
                                self.push_token(self.create_token(None, None, None, None, "dialogue_begin"));
                            }

                            // 移除角色名后的^符号
                            text_valid = Regex::new(r"\^\s*$").unwrap().replace(&text_valid, "").to_string();

                            // 替代前瞻性判断的实现：分别处理三种情况
                            // 1. ^空白字符后跟注释开始符号 இ
                            if let Some(captures) = Regex::new(r"(\^\s*)(இ.*$)").unwrap().captures(&self.text_display) {
                                let note_part = captures.get(2).map_or("", |m| m.as_str());
                                self.text_display = note_part.to_string();
                            }
                            // 2. ^空白字符后跟注释开始符号 ↺
                            else if let Some(captures) = Regex::new(r"(\^\s*)(↺.*$)").unwrap().captures(&self.text_display) {
                                let note_part = captures.get(2).map_or("", |m| m.as_str());
                                self.text_display = note_part.to_string();
                            }
                            // 3. ^空白字符在行尾
                            else {
                                self.text_display = Regex::new(r"\^\s*$").unwrap()
                                    .replace(&self.text_display, "").to_string();
                            }
                        } else {
                            self.push_token(self.create_token(None, None, None, None, "dialogue_begin"));
                        }

                        self.force_not_dual = false;
                        let character = self.trim_character_extension(&text_valid).trim().to_string();
                        _previous_character = Some(character.clone());
                        this_token.character = Some(character.clone());

                        // 更新角色信息
                        let scene_idx = self.result.properties.scenes.len().saturating_sub(1);

                        if let Some(values) = self.result.properties.characters.get_mut(&character) {
                            if !values.contains(&scene_idx) {
                                values.push(scene_idx);
                            }
                        } else {
                            self.result.properties.characters.insert(character.clone(), vec![scene_idx]);
                        }

                        if self.result.properties.character_lines.is_none() {
                            self.result.properties.character_lines = Some(HashMap::new());
                        }

                        if let Some(char_lines) = &mut self.result.properties.character_lines {
                            char_lines.insert(this_token.line, character.clone());
                        }

                        if self.result.properties.scenes.is_empty() {
                            if self.result.properties.character_first_line.is_none() ||
                               !self.result.properties.character_first_line.as_ref().unwrap().contains_key(&character) {

                                if self.result.properties.character_first_line.is_none() {
                                    self.result.properties.character_first_line = Some(HashMap::new());
                                }

                                if let Some(char_first_line) = &mut self.result.properties.character_first_line {
                                    char_first_line.insert(character.clone(), this_token.line);
                                }

                                if self.result.properties.character_describe.is_none() {
                                    self.result.properties.character_describe = Some(HashMap::new());
                                }

                                if let Some(char_describe) = &mut self.result.properties.character_describe {
                                    char_describe.insert(character.clone(), text.to_string());
                                }
                            }
                        }

                        last_character_index = self.result.tokens.len();

                        // 对话角色加入结构树
                        if let Some(last_scen_structure_token) = &mut self.last_scen_structure_token {
                            if cfg.dialogue_foldable {
                                let cobj = StructToken {
                                    text: text_valid.clone(),
                                    children: Vec::new(),
                                    level: 0,
                                    section: false,
                                    synopses: Vec::new(),
                                    notes: Vec::new(),
                                    isscene: false,
                                    ischartor: true,
                                    dialogue_end_line: if self.lines_length > 0 { self.lines_length - 1 } else { 0 },
                                    duration_sec: 0.0,
                                    range: Some(Range {
                                        start: Position { line: this_token.line, character: 0 },
                                        end: Position { line: this_token.line, character: text_valid.len() },
                                    }),
                                    id: Some(format!("{}/{}", last_scen_structure_token.id.clone().unwrap_or_default(), this_token.line)),
                                    isnote: false,
                                    play_sec: 0.0,
                                    structs: Vec::new(),
                                    duration: 0.0,
                                };

                                last_scen_structure_token.children.push(cobj.clone());
                                self.last_chartor_structure_token = Some(cobj);
                            }
                        }

                        // 处理角色名格式
                        self.text_display = Regex::new(r"^(.*?↻)?[ \t]*@").unwrap()
                            .replace(&self.text_display, "").trim().to_string();
                        this_token.text = self.text_display.clone();

                        if cfg.print_dialogue_numbers {
                            self.add_dialogue_number_decoration(&mut this_token);
                        }

                        self.push_token(this_token);
                        continue;
                    } else if BLOCK_REGEX.get("action_force").unwrap().is_match(&self.text_valid) {
                        // 处理强制动作
                        self.process_title_page_end(i);
                        this_token.token_type = "action".to_string();

                        let mt = Regex::new(r"^((?:.*?↻)?\s*)(\!)(.*)")
                            .unwrap()
                            .captures(&self.text_display);

                        if let Some(captures) = mt {
                            let group1 = captures.get(1).map_or("", |m| m.as_str());
                            let group3 = captures.get(3).map_or("", |m| m.as_str());

                            this_token.text = format!("{}{}", group1, group3);
                            process_token_text_style_char(&mut this_token);
                            this_token = self.process_action_block(this_token);
                        }

                        self.push_token(this_token);
                        continue;
                    }  else {
                        action = true;
                    }
                } else {
                    action = true;
                }

                if action {
                    if self.regex.get("centered").unwrap().is_match(&self.text_valid) {
                        // 处理居中文本
                        self.process_title_page_end(i);
                        this_token.token_type = "centered".to_string();

                        let mt = Regex::new(r"((?:^.*?↻)|^)[ \t]*>\s*(.+)\s*?<\s*((?:இ.*$)|(?:↺.*$)|$)").unwrap()
                            .captures(&self.text_display);

                        if let Some(captures) = mt {
                            let group1 = captures.get(1).map_or("", |m| m.as_str()).trim();
                            let group2 = captures.get(2).map_or("", |m| m.as_str()).trim();
                            let group3 = captures.get(3).map_or("", |m| m.as_str()).trim();

                            this_token.text = format!("{}{}{}", group1, group2, group3);
                            process_token_text_style_char(&mut this_token);
                        }

                        self.push_token(this_token);
                        continue;
                    }else if self.regex.get("section").unwrap().is_match(&self.text_valid) {
                        // 处理章节标题
                        self.process_title_page_end(i);
                        this_token.token_type = "section".to_string();

                        let mt = Regex::new(r"^((?:.*?↻)?\s*)(#+)(?:\s*)(.*)").unwrap()
                            .captures(&self.text_display);

                        if let Some(captures) = mt {
                            let group1 = captures.get(1).map_or("", |m| m.as_str());
                            let group2 = captures.get(2).map_or("", |m| m.as_str());
                            let group3 = captures.get(3).map_or("", |m| m.as_str());

                            this_token.level = Some(group2.len() as i32);
                            this_token.text = format!("{}{}", group1, group3);
                            process_token_text_style_char(&mut this_token);

                            println!("【parse】section: {}, level: {:?}", this_token.text, this_token.level);

                            // 创建结构树节点
                            let mut cobj = StructToken {
                                text: group3.to_string(),
                                children: Vec::new(),
                                level: group2.len(),
                                section: true,
                                synopses: Vec::new(),
                                notes: Vec::new(),
                                isscene: false,
                                ischartor: false,
                                dialogue_end_line: 0,
                                duration_sec: 0.0,
                                range: Some(Range {
                                    start: Position { line: this_token.line, character: 0 },
                                    end: Position { line: this_token.line, character: self.text_valid.len() },
                                }),
                                id: None,
                                isnote: false,
                                play_sec: self.play_time_sec,
                                structs: Vec::new(),
                                duration: 0.0,
                            };

                            self.current_depth = group2.len();

                            if self.current_depth == 1 {
                                cobj.id = Some(format!("/{}", this_token.line));
                                self.result.properties.structure.push(cobj);
                            } else {
                                if let Some(level) = self.latest_section(self.current_depth - 1) {
                                    cobj.id = Some(format!("{}/{}", level.id.as_ref().unwrap_or(&String::new()), this_token.line));

                                    // 找到父节点并添加子节点
                                    for parent in &mut self.result.properties.structure {
                                        if parent.id == level.id {
                                            parent.children.push(cobj);
                                            break;
                                        }
                                    }
                                } else {
                                    cobj.id = Some(format!("/{}", this_token.line));
                                    self.result.properties.structure.push(cobj);
                                }
                            }
                        }

                        self.push_token(this_token);
                        continue;
                    } else if self.regex.get("page_break").unwrap().is_match(&self.text_valid) {
                        // 处理分页符（必须在synopsis之前检查，因为===也会匹配synopsis的正则）
                        self.process_title_page_end(i);
                        this_token.token_type = "page_break".to_string();
                        this_token.text = "".to_string();

                        self.push_token(this_token);
                        continue;
                    } else if self.regex.get("synopsis").unwrap().is_match(&self.text_valid) {
                        // 处理概要
                        self.process_title_page_end(i);
                        this_token.token_type = "synopsis".to_string();

                        let mt = Regex::new(r"^((?:.*?↻)?\s*)(?:\=)(.*)").unwrap()
                            .captures(&self.text_display);

                        if let Some(captures) = mt {
                            let group1 = captures.get(1).map_or("", |m| m.as_str());
                            let group2 = captures.get(2).map_or("", |m| m.as_str());

                            this_token.text = format!("{}{}", group1, group2);
                            process_token_text_style_char(&mut this_token);

                            // 添加概要到结构树
                            let synopsis = Synopsis {
                                synopsis: group2.trim().to_string(),
                                line: this_token.line,
                            };

                            if self.current_depth == 0 {
                                if let Some(last_scen) = &mut self.last_scen_structure_token {
                                    last_scen.synopses.push(synopsis);
                                }
                            } else {
                                if let Some(level) = self.latest_section(self.current_depth) {
                                    for parent in &mut self.result.properties.structure {
                                        if parent.id == level.id {
                                            parent.synopses.push(synopsis);
                                            break;
                                        }
                                    }
                                }
                            }
                        }

                        self.push_token(this_token);
                        continue;
                    } else if BLOCK_REGEX.get("lyric").unwrap().is_match(&self.text_valid) {
                        // 处理歌词
                        self.process_title_page_end(i);
                        this_token.token_type = "lyric".to_string();

                        let mt = Regex::new(r"^((?:.*?↻)?\s*)(\~)(\s*)(.*)").unwrap()
                            .captures(&self.text_display);

                        if let Some(captures) = mt {
                            let group1 = captures.get(1).map_or("", |m| m.as_str());
                            let group3 = captures.get(3).map_or("", |m| m.as_str());
                            let group4 = captures.get(4).map_or("", |m| m.as_str());

                            this_token.text = format!("{}{}{}", group1, group3, group4);
                            process_token_text_style_char(&mut this_token);
                        }

                        self.push_token(this_token);
                        continue;
                    } else {
                        self.process_title_page_end(i);
                        this_token.token_type = "action".to_string();
                        this_token.text = self.text_display.clone();
                        process_token_text_style_char(&mut this_token);
                        this_token = self.process_action_block(this_token);

                        self.push_token(this_token);
                        continue;
                    }
                }
            }







            // 如果没有匹配任何特定类型，默认为动作
            if this_token.token_type.is_empty() {
                // 如果仍在标题页状态，跳过默认处理
                if self.result.state == "title_page" {
                    continue;
                }

                self.process_title_page_end(i);
                this_token.token_type = "action".to_string();
                process_token_text_style_char(&mut this_token);
                this_token = self.process_action_block(this_token);
            }

            last_was_separator = false;

            if self.result.state != "ignore" {
                if this_token.token_type == "scene_heading" || this_token.token_type == "transition" {
                    this_token.text = this_token.text.to_uppercase();
                }

                if this_token.token_type != "action" && this_token.token_type != "dialogue" {
                    this_token.text = this_token.text.trim().to_string();
                }

                if let Some(note_token) = note_token.take() {
                    self.push_token(note_token);
                }

                if this_token.ignore {
                    ignored_last_token = true;
                } else {
                    ignored_last_token = false;
                    self.push_token(this_token);
                }
            }
        }

        // 所有文档行解析完后，如果是直接截断的对话block，额外添加token
        if self.result.state == "dialogue" {
            self.push_token(self.create_token(None, None, None, None, "dialogue_end"));
        }

        if self.result.state == "dual_dialogue" {
            self.push_token(self.create_token(None, None, None, None, "dual_dialogue_end"));
        }

        if _need_process_outline_note > 0 {
            self.process_inline_notes();
            _need_process_outline_note = 0;
        }

        self.update_previous_scene_length(); // 统计最后一个场景的时长

        // 处理镜头交切的场景，将场景时间平均分配调整一下
        for shot_cut_token in &mut self.shot_cut_strct_tokens {
            let duration = shot_cut_token.get("duration")
                .and_then(|v| v.as_f64())
                .unwrap_or(0.0);

            if let Some(structs) = shot_cut_token.get_mut("structs") {
                if let Some(structs_array) = structs.as_array_mut() {
                    if duration == 0.0 || structs_array.is_empty() {
                        continue;
                    }

                    let average_duration = duration / structs_array.len() as f64;
                    for struct_token in structs_array {
                        if let Some(token) = struct_token.as_object_mut() {
                            let current_duration = token.get("durationSec")
                                .and_then(|v| v.as_f64())
                                .unwrap_or(0.0);

                            token.insert("durationSec".to_string(), serde_json::to_value(current_duration + average_duration).unwrap());
                        }
                    }
                }
            }
        }

        // 保存场景变量
        self.result.properties.scene_number_vars = Some(dup_scence_nuber.keys().cloned().collect());

        // 处理完所有的token，进行最后的处理
        // 所有action里面出现过的角色，也应该计入在场景中出现过
        let mut last_scene_idx: i32 = -1;
        for token in &mut self.result.tokens {
            if token.token_type == "scene_heading" {
                last_scene_idx += 1;
            } else if last_scene_idx >= 0 && token.token_type == "action" {
                if !token.text.is_empty() {
                    let mut char_map: HashMap<String, Vec<usize>> = HashMap::new(); // 角色，在行中字符index的start和end

                    // 先将result.properties.characters按照角色名的长度排序，长的在前面
                    let mut sorted_keys: Vec<String> = self.result.properties.characters.keys().cloned().collect();
                    sorted_keys.sort_by(|a, b| b.len().cmp(&a.len()));

                    for k in sorted_keys {
                        // 需要判断，角色名匹配，不能在同一行中索引重叠
                        let mut added = false;
                        let mut i = 0;

                        while i < token.text.len() {
                            if let Some(idx) = token.text[i..].find(&k) {
                                let start = i + idx;
                                let end = start + k.len();
                                let mut overlap = false;

                                for (_existing_k, v) in &char_map {
                                    if v[0] < end && v[1] > start {
                                        overlap = true;
                                        break;
                                    }
                                }

                                if !overlap {
                                    char_map.insert(k.clone(), vec![start, end]);
                                    added = true;
                                    break;
                                } else {
                                    i = end; // 继续查找下一个
                                }
                            } else {
                                break;
                            }
                        }

                        if added {
                            if let Some(v) = self.result.properties.characters.get_mut(&k) {
                                // 角色在action中出现过，也算在场景中出现过
                                if !v.contains(&(last_scene_idx as usize)) {
                                    v.push(last_scene_idx as usize);
                                }
                            }

                            if token.characters_action.is_none() {
                                token.characters_action = Some(Vec::new());
                            }

                            if let Some(chars) = &mut token.characters_action {
                                chars.push(k);
                            }
                        }
                    }
                }
            }
        }

        // 转换result.properties.characters成result.properties.characterSceneNumber
        let mut character_scene_number: HashMap<String, HashSet<String>> = HashMap::new();

        for (k, v) in &self.result.properties.characters {
            let mut scene_numbers = HashSet::new();

            // 按照场景头索引排序，去掉负数的场景号
            let mut sorted_indices: Vec<&usize> = v.iter().collect();
            sorted_indices.sort();

            for &idx in sorted_indices {
                if idx < self.result.properties.scenes.len() {
                    if let Some(scene) = self.result.properties.scenes.get(idx) {
                        if let Some(number) = scene.get("number") {
                            if let Some(number_str) = number.as_str() {
                                scene_numbers.insert(number_str.to_string());
                            }
                        }
                    }
                }
            }

            character_scene_number.insert(k.clone(), scene_numbers);
        }

        self.result.properties.character_scene_number = Some(character_scene_number);

        // 生成HTML输出
        if generate_html {
            self.result.script_html = Some(crate::parser::text_processor::generate_html(&self.result.tokens));
            self.result.title_html = Some(crate::parser::text_processor::generate_title_html(&self.result.properties.title_keys, &self.result.tokens));
        }

        // 计算解析时间
        let end_time = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis() as u64;

        self.result.parse_time = end_time - self.result.parse_time;

        self.result.clone()
    }

    // 初始化正则表达式
    fn init_regex(&mut self) {
        self.regex.insert(
            "title_page".to_string(),
            Regex::new(r"(?i)^[ \t]*(title|credit|author[s]?|source|notes|draft date|date|watermark|contact( info)?|revision|copyright|font|font italic|font bold|font bold italic|metadata|tl|tc|tr|cc|br|bl|header|footer)\:.*").unwrap()
        );
        self.regex.insert(
            "section".to_string(),
            Regex::new(r"^[ \t]*(#+)(?:\s*)(.*)").unwrap()
        );
        self.regex.insert(
            "synopsis".to_string(),
            Regex::new(r"^[ \t]*(?:\=)(.*)").unwrap()
        );
        self.regex.insert(
            "scene_heading".to_string(),
            Regex::new(r"^[ \t]*([.]|(?i:int|ext|est|int[.]?\/ext|i[.]?\/e)[. ])\s*([^#]*)(#\s*[^\s].*#)?\s*$").unwrap()
        );
        self.regex.insert(
            "scene_number".to_string(),
            Regex::new(r"#\s*(?:\$\{\s*([^\}\s]*)\s*\})?\s*([^#]*)\s*#").unwrap()
        );
        self.regex.insert(
            "transition".to_string(),
            Regex::new(r"^\s*(?:(>)[^\n\r<]*|[A-Z ]+TO:)$").unwrap()
        );
        self.regex.insert(
            "character".to_string(),
            Regex::new(r"^[ \t]*((\p{Lu}[^\p{Ll}\r\n@]*)|(@[^\r\n\(（\^]*))(\(.*\)|（.*）)?(\s*\^)?\s*$").unwrap()
        );
        self.regex.insert(
            "parenthetical".to_string(),
            Regex::new(r"^[ \t]*(\(.+\)|（.+）)\s*$").unwrap()
        );
        self.regex.insert(
            "parenthetical_start".to_string(),
            Regex::new(r"^[ \t]*(?:\(|（)[^\)）]*$").unwrap()
        );
        self.regex.insert(
            "parenthetical_end".to_string(),
            Regex::new(r"^.*(?:\)|）)\s*$").unwrap()
        );
        self.regex.insert(
            "action".to_string(),
            Regex::new(r"^(.+)").unwrap()
        );
        self.regex.insert(
            "centered".to_string(),
            Regex::new(r"^[ \t]*>\s*(.+)\s*<\s*$").unwrap()
        );
        self.regex.insert(
            "page_break".to_string(),
            Regex::new(r"^\s*\={3,}\s*$").unwrap()
        );
        self.regex.insert(
            "line_break".to_string(),
            Regex::new(r"^ {2,}$").unwrap()
        );
        self.regex.insert(
            "note_inline".to_string(),
            Regex::new(r"\[{2}([^\[].+?)\]{2}").unwrap()
        );
        self.regex.insert(
            "emphasis".to_string(),
            Regex::new(r"(_|\*{1,3}|_\*{1,3}|\*{1,3}_)(.+)(_|\*{1,3}|_\*{1,3}|\*{1,3}_)").unwrap()
        );
        self.regex.insert(
            "bold_italic_underline".to_string(),
            Regex::new(r"(_{1}\*{3}|\*{3}_{1})(.+?)(\*{3}_{1}|_{1}\*{3})").unwrap()
        );
        self.regex.insert(
            "bold_underline".to_string(),
            Regex::new(r"(_{1}\*{2}|\*{2}_{1})(.+?)(\*{2}_{1}|_{1}\*{2})").unwrap()
        );
        self.regex.insert(
            "italic_underline".to_string(),
            Regex::new(r"(_{1}\*{1}|\*{1}_{1})(.+?)(\*{1}_{1}|_{1}\*{1})").unwrap()
        );
        self.regex.insert(
            "bold_italic".to_string(),
            Regex::new(r"(\*{3})(.+?)(\*{3})").unwrap()
        );
        self.regex.insert(
            "bold".to_string(),
            Regex::new(r"(\*{2})(.+?)(\*{2})").unwrap()
        );
        self.regex.insert(
            "italic".to_string(),
            Regex::new(r"(\*{1})(.+?)(\*{1})").unwrap()
        );
        self.regex.insert(
            "underline".to_string(),
            Regex::new(r"(_{1})(.+?)(_{1})").unwrap()
        );
    }

    // 初始化标题页显示配置
    fn init_title_page_display(&mut self) {
        self.title_page_display.insert("title".to_string(), TitleKeywordFormat { position: "cc".to_string(), index: 0 });
        self.title_page_display.insert("credit".to_string(), TitleKeywordFormat { position: "cc".to_string(), index: 1 });
        self.title_page_display.insert("author".to_string(), TitleKeywordFormat { position: "cc".to_string(), index: 2 });
        self.title_page_display.insert("authors".to_string(), TitleKeywordFormat { position: "cc".to_string(), index: 3 });
        self.title_page_display.insert("source".to_string(), TitleKeywordFormat { position: "cc".to_string(), index: 4 });

        self.title_page_display.insert("watermark".to_string(), TitleKeywordFormat { position: "hidden".to_string(), index: -1 });
        self.title_page_display.insert("font".to_string(), TitleKeywordFormat { position: "hidden".to_string(), index: -1 });
        self.title_page_display.insert("font_italic".to_string(), TitleKeywordFormat { position: "hidden".to_string(), index: -1 });
        self.title_page_display.insert("font_bold".to_string(), TitleKeywordFormat { position: "hidden".to_string(), index: -1 });
        self.title_page_display.insert("font_bold_italic".to_string(), TitleKeywordFormat { position: "hidden".to_string(), index: -1 });
        self.title_page_display.insert("header".to_string(), TitleKeywordFormat { position: "hidden".to_string(), index: -1 });
        self.title_page_display.insert("footer".to_string(), TitleKeywordFormat { position: "hidden".to_string(), index: -1 });
        self.title_page_display.insert("metadata".to_string(), TitleKeywordFormat { position: "hidden".to_string(), index: -1 });

        self.title_page_display.insert("notes".to_string(), TitleKeywordFormat { position: "bl".to_string(), index: 0 });
        self.title_page_display.insert("copyright".to_string(), TitleKeywordFormat { position: "bl".to_string(), index: 1 });

        self.title_page_display.insert("revision".to_string(), TitleKeywordFormat { position: "br".to_string(), index: 0 });
        self.title_page_display.insert("date".to_string(), TitleKeywordFormat { position: "br".to_string(), index: 1 });
        self.title_page_display.insert("draft_date".to_string(), TitleKeywordFormat { position: "br".to_string(), index: 2 });
        self.title_page_display.insert("contact".to_string(), TitleKeywordFormat { position: "br".to_string(), index: 3 });
        self.title_page_display.insert("contact_info".to_string(), TitleKeywordFormat { position: "br".to_string(), index: 4 });

        self.title_page_display.insert("br".to_string(), TitleKeywordFormat { position: "br".to_string(), index: -1 });
        self.title_page_display.insert("bl".to_string(), TitleKeywordFormat { position: "bl".to_string(), index: -1 });
        self.title_page_display.insert("tr".to_string(), TitleKeywordFormat { position: "tr".to_string(), index: -1 });
        self.title_page_display.insert("tc".to_string(), TitleKeywordFormat { position: "tc".to_string(), index: -1 });
        self.title_page_display.insert("tl".to_string(), TitleKeywordFormat { position: "tl".to_string(), index: -1 });
        self.title_page_display.insert("cc".to_string(), TitleKeywordFormat { position: "cc".to_string(), index: -1 });
    }
}
