//! DOCX 生成模块
//!
//! 该模块提供了与原始 TypeScript 版本 docxmaker.ts 兼容的 API

use crate::models::Conf;
use crate::parser::fountain_parser::Line;
use crate::parser::ParseOutput;
use crate::utils::is_blank_line_after_style;
use std::collections::HashMap;
use thiserror::Error;

// 使用适配器中的类型
use crate::docx::adapter::docx::{Document, TextRun};
use crate::docx::adapter::{
    convert_inches_to_twip, convert_point_to_inches, convert_point_to_twip, DocxAdapterError,
    DocxAsBase64, DocxStats, LineStruct, RunProps, StyleStash, UnderlineTypeConst,
};

use super::adapter::docx::ParagraphSpacing;

/// DOCX导出错误类型
#[derive(Error, Debug)]
pub enum DocxError {
    #[error("IO错误: {0}")]
    IoError(#[from] std::io::Error),

    #[error("DOCX生成错误: {0}")]
    DocxError(#[from] docx_rs::DocxError),

    #[error("适配器错误: {0}")]
    AdapterError(#[from] DocxAdapterError),

    #[error("无效的配置: {0}")]
    InvalidConfig(String),
}

/// DOCX导出结果
pub type DocxResult<T> = Result<T, DocxError>;

/// 打印配置
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct PrintProfile {
    /// 字体大小 //磅
    pub font_size: f32,
    /// 注释字体大小  //磅
    pub note_font_size: f32,
    /// 每页行数
    pub lines_per_page: usize,
    /// 页面宽度 //英寸
    pub page_width: f32,
    /// 页面高度 //英寸
    pub page_height: f32,
    /// 字体宽度 //英寸
    pub font_width: f32,
    /// 纸张大小
    pub paper_size: String,
    /// 上边距
    pub top_margin: f32,
    /// 下边距
    pub bottom_margin: f32,
    /// 左边距
    pub left_margin: f32,
    /// 右边距
    pub right_margin: f32,
    /// 注释配置
    pub note: NoteConfig,
    /// 页码上边距
    pub page_number_top_margin: f32,
    /// 场景标题配置
    pub scene_heading: ElementConfig,
    /// 动作配置
    pub action: ElementConfig,
    /// 角色配置
    pub character: ElementConfig,
    /// 对话配置
    pub dialogue: ElementConfig,
    /// 括号配置
    pub parenthetical: ElementConfig,
    /// 章节配置
    pub section: SectionConfig,
    /// 概要配置
    pub synopsis: SynopsisConfig,
    /// 注释行高 //英寸
    pub note_line_height: f32,
    /// 字距 //磅
    pub character_spacing: f32,
}

impl Default for PrintProfile {
    fn default() -> Self {
        Self {
            // 基于"中文a4"配置
            font_size: 12.0, //磅
            note_font_size: 9.0,
            lines_per_page: 30,
            page_width: 8.27,
            page_height: 11.69,
            font_width: 0.1, //英寸
            paper_size: "a4".to_string(),
            top_margin: 1.19, //英寸
            bottom_margin: 1.0,
            left_margin: 1.5,
            right_margin: 1.5,
            note: NoteConfig::default(),
            page_number_top_margin: 0.4,
            scene_heading: ElementConfig {
                feed: 1.2,
                color: None,
                italic: false,
            },
            action: ElementConfig {
                feed: 1.2,
                color: None,
                italic: false,
            },
            character: ElementConfig {
                feed: 3.0,
                color: None,
                italic: false,
            },
            dialogue: ElementConfig {
                feed: 2.2,
                color: None,
                italic: false,
            },
            parenthetical: ElementConfig {
                feed: 2.5,
                color: None,
                italic: false,
            },
            section: SectionConfig {
                feed: 0.2,
                color: Some("#555555".to_string()),
                italic: false,
                level_indent: 0.2,
            },
            synopsis: SynopsisConfig::default(),
            note_line_height: 0.17,
            character_spacing: 1.0,
        }
    }
}

/// 注释配置
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct NoteConfig {
    /// 颜色
    pub color: String,
    /// 是否斜体
    pub italic: bool,
}

impl Default for NoteConfig {
    fn default() -> Self {
        Self {
            color: "#888888".to_string(),
            italic: true,
        }
    }
}

/// 元素配置
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct ElementConfig {
    /// 缩进
    pub feed: f32,
    /// 颜色
    pub color: Option<String>,
    /// 是否斜体
    pub italic: bool,
}

impl Default for ElementConfig {
    fn default() -> Self {
        Self {
            feed: 1.5,
            color: None,
            italic: false,
        }
    }
}

/// 章节配置
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct SectionConfig {
    /// 缩进
    pub feed: f32,
    /// 颜色
    pub color: Option<String>,
    /// 是否斜体
    pub italic: bool,
    /// 层级缩进
    pub level_indent: f32,
}

/// 概要配置
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct SynopsisConfig {
    /// 缩进
    pub feed: Option<f32>,
    /// 颜色
    pub color: Option<String>,
    /// 是否斜体
    pub italic: bool,
    /// 内边距
    pub padding: Option<f32>,
    /// 是否根据最后一个章节调整缩进
    pub feed_with_last_section: bool,
}

impl Default for SectionConfig {
    fn default() -> Self {
        Self {
            feed: 1.5,
            color: Some("#666666".to_string()),
            italic: false,
            level_indent: 0.2,
        }
    }
}

impl Default for SynopsisConfig {
    fn default() -> Self {
        Self {
            feed: Some(0.2),
            color: Some("#888888".to_string()),
            italic: true,
            padding: Some(0.0),
            feed_with_last_section: true,
        }
    }
}

/// 注释
#[derive(Debug, Clone, Default)]
pub struct Note {
    /// 注释编号
    pub no: usize,
    /// 注释文本
    pub text: Vec<String>,
}

/// 当前注释
#[derive(Debug, Clone)]
pub struct CurrentNote {
    /// 页面索引
    pub page_idx: i32,
    /// 注释内容
    pub note: Note,
}

impl Default for CurrentNote {
    fn default() -> Self {
        Self {
            page_idx: -1, // 初始化为 -1，表示不在脚注中
            note: Note::default(),
        }
    }
}

/// DOCX导出选项
#[derive(Debug, Clone)]
pub struct DocxOptions {
    /// 行高
    pub line_height: f32,
    /// 文件路径
    pub filepath: String,
    /// 配置
    pub config: Conf,
    /// 解析结果
    pub parsed: Option<ParseOutput>,
    /// 打印配置
    pub print_profile: PrintProfile,
    /// 字体
    pub font: String,
    /// 导出配置
    pub exportconfig: Option<ExportConfig>,
    /// 斜体字体
    pub font_italic: String,
    /// 粗体字体
    pub font_bold: String,
    /// 粗斜体字体
    pub font_bold_italic: String,
    /// 右列样式缓存
    pub stash_style_right_column: Option<StyleStash>,
    /// 左列样式缓存
    pub stash_style_left_column: Option<StyleStash>,
    /// 全局样式缓存
    pub stash_style_global_column: Option<StyleStash>,
    /// 全局斜体
    pub italic_global: bool,
    /// 动态斜体
    pub italic_dynamic: bool,
    /// 是否找到斜体字体
    pub found_font_italic: bool,
    /// 是否找到粗体字体
    pub found_font_bold: bool,
    /// 是否找到粗斜体字体
    pub found_font_bold_italic: bool,
    /// 元数据
    pub metadata: Option<HashMap<String, String>>,
    /// 是否为预览
    pub for_preview: bool,
    /// 标题页是否已处理
    pub title_page_processed: bool,
}

impl Default for DocxOptions {
    fn default() -> Self {
        Self {
            line_height: 1.0,
            filepath: String::new(),
            config: Conf::default(),
            parsed: None,
            print_profile: PrintProfile::default(),
            font: "Courier Prime".to_string(),
            exportconfig: None,
            font_italic: String::new(),
            font_bold: String::new(),
            font_bold_italic: String::new(),
            stash_style_right_column: None,
            stash_style_left_column: None,
            stash_style_global_column: None,
            italic_global: false,
            italic_dynamic: false,
            found_font_italic: false,
            found_font_bold: false,
            found_font_bold_italic: false,
            metadata: None,
            for_preview: false,
            title_page_processed: false,
        }
    }
}

/// 导出配置
#[derive(Debug, Clone)]
pub struct ExportConfig {
    /// 字体名称
    pub font_family: String,
}

impl Default for ExportConfig {
    fn default() -> Self {
        Self {
            font_family: "Courier Prime".to_string(),
        }
    }
}

/// 缓存的对话组
#[derive(Debug, Clone)]
pub struct CachedDialogueGroup {
    pub style: String,
    pub children: Vec<TextRun>,
    pub indent_left: Option<i32>,
    pub indent_right: Option<i32>,
}

/// 文档上下文
pub struct DocxContext {
    pub options: DocxOptions,
    pub font_names: HashMap<String, String>,
    pub format_state: StyleStash,
    pub run_normal: RunProps,
    pub run_bold: RunProps,
    pub run_italic: RunProps,
    pub run_bold_italic: RunProps,
    pub run_notes: RunProps,
    pub rm_blank_line: i32,
    pub china_format: i32,
    pub cache_triangle: bool,
    pub force_note_orig: bool,
    pub current_note: CurrentNote,
    pub notes_len: usize,
    pub doc: Document,
    // china_format 缓存变量
    pub last_dial_gr: Option<CachedDialogueGroup>,
    pub last_dial_gr_left: Option<CachedDialogueGroup>,
    pub last_dial_gr_right: Option<CachedDialogueGroup>,
    // 全局双对话表格缓存 - 参考原项目的lastDialTableLeft和lastDialTableRight
    pub last_dial_table_left: Vec<crate::docx::adapter::docx::Paragraph>,
    pub last_dial_table_right: Vec<crate::docx::adapter::docx::Paragraph>,
}

impl DocxContext {
    /// 创建新的文档上下文
    pub fn new(options: DocxOptions) -> Self {
        let mut font_names = HashMap::new();
        font_names.insert("normal".to_string(), "Courier Prime".to_string());

        if options.font != "Courier Prime" && !options.font.is_empty() {
            font_names.insert("normal".to_string(), options.font.clone());
        }

        if !options.font_italic.is_empty() {
            font_names.insert("italic".to_string(), options.font_italic.clone());
            font_names.insert("bold_italic".to_string(), options.font_italic.clone());
        }

        if !options.font_bold.is_empty() {
            font_names.insert("bold".to_string(), options.font_bold.clone());
        }

        if !options.font_bold_italic.is_empty() {
            font_names.insert("bold_italic".to_string(), options.font_bold_italic.clone());
        }

        let font_size = (options.print_profile.font_size) as usize;
        let note_font_size = (options.print_profile.note_font_size) as usize;

        let run_normal = RunProps {
            size: Some(font_size),
            font: Some(
                font_names
                    .get("normal")
                    .unwrap_or(&"Courier Prime".to_string())
                    .clone(),
            ),
            bold: Some(false),
            italic: Some(false),
            ..Default::default()
        };

        let run_bold = RunProps {
            size: Some(font_size),
            font: Some(
                font_names
                    .get("bold")
                    .unwrap_or(
                        font_names
                            .get("normal")
                            .unwrap_or(&"Courier Prime".to_string()),
                    )
                    .clone(),
            ),
            bold: Some(!options.found_font_bold),
            italic: Some(false),
            ..Default::default()
        };

        let run_italic = RunProps {
            size: Some(font_size),
            font: Some(
                font_names
                    .get("italic")
                    .unwrap_or(
                        font_names
                            .get("normal")
                            .unwrap_or(&"Courier Prime".to_string()),
                    )
                    .clone(),
            ),
            bold: Some(false),
            italic: Some(!options.found_font_italic),
            ..Default::default()
        };

        let run_bold_italic = RunProps {
            size: Some(font_size),
            font: Some(
                font_names
                    .get("bold_italic")
                    .unwrap_or(
                        font_names
                            .get("normal")
                            .unwrap_or(&"Courier Prime".to_string()),
                    )
                    .clone(),
            ),
            bold: Some(!options.found_font_bold_italic),
            italic: Some(!options.found_font_italic),
            ..Default::default()
        };

        let run_notes = RunProps {
            size: Some(note_font_size),
            font: Some(
                font_names
                    .get("normal")
                    .unwrap_or(&"Courier Prime".to_string())
                    .clone(),
            ),
            color: Some("868686".to_string()),
            ..Default::default()
        };

        let mut rm_blank_line = 0;
        let mut china_format = 0;

        if let Some(metadata) = &options.metadata {
            // 处理嵌套的 print 对象
            if metadata.contains_key("print") {
                if let Some(china_format_str) = metadata.get("print.chinaFormat") {
                    if let Ok(value) = china_format_str.parse::<i32>() {
                        china_format = value;
                    }
                }

                if let Some(rm_blank_line_str) = metadata.get("print.rmBlankLine") {
                    if let Ok(value) = rm_blank_line_str.parse::<i32>() {
                        rm_blank_line = value;
                    }
                }
            }
        }

        DocxContext {
            options,
            font_names,
            format_state: StyleStash::default(),
            run_normal,
            run_bold,
            run_italic,
            run_bold_italic,
            run_notes,
            rm_blank_line,
            china_format,
            cache_triangle: false,
            force_note_orig: false,
            current_note: CurrentNote::default(),
            notes_len: 0,
            doc: Document::new(),
            // 初始化 china_format 缓存变量
            last_dial_gr: None,
            last_dial_gr_left: None,
            last_dial_gr_right: None,
            // 初始化全局双对话表格缓存
            last_dial_table_left: Vec::new(),
            last_dial_table_right: Vec::new(),
        }
    }

    /// 重置格式状态
    pub fn reset_format(&mut self) {
        // 完全重置格式状态，与原始项目保持一致
        self.format_state = StyleStash::default();

        // 重置斜体状态
        self.options.italic_global = false;
        self.options.italic_dynamic = false;
    }

    /// 创建页码运行
    pub fn create_page_number_runs(
        &self,
        show_page_numbers: &str,
    ) -> Vec<crate::docx::adapter::docx::RunType> {
        let mut runs = Vec::new();

        // 分割页码格式字符串，处理 {n} 占位符
        let parts: Vec<&str> = show_page_numbers.split("{n}").collect();

        for (i, part) in parts.iter().enumerate() {
            // 添加文本部分
            if !part.is_empty() {
                let text_run = crate::docx::adapter::docx::TextRun::new(part);
                runs.push(crate::docx::adapter::docx::RunType::Text(text_run));
            }

            // 如果不是最后一个部分，添加页码
            if i < parts.len() - 1 {
                let mut page_run = crate::docx::adapter::docx::PageNumberRun::new();
                page_run.add_page_number();
                runs.push(crate::docx::adapter::docx::RunType::PageNumber(page_run));
            }
        }

        runs
    }

    /// 保存全局样式
    pub fn global_stash(&mut self) {
        self.options.stash_style_global_column = Some(StyleStash {
            bold_italic: self.format_state.bold_italic,
            bold: self.format_state.bold,
            italic: self.format_state.italic,
            underline: self.format_state.underline,
            override_color: self.format_state.override_color.clone(),
            italic_global: self.options.italic_global,
            italic_dynamic: self.options.italic_dynamic,
            current_color: self.format_state.current_color.clone(),
        });
        self.reset_format();
    }

    /// 恢复全局样式
    pub fn global_pop(&mut self) {
        if let Some(stash) = &self.options.stash_style_global_column {
            self.format_state.bold_italic = stash.bold_italic;
            self.format_state.bold = stash.bold;
            self.format_state.italic = stash.italic;
            self.format_state.underline = stash.underline;
            self.format_state.override_color = stash.override_color.clone();
            self.format_state.current_color = stash.current_color.clone();
            self.options.italic_global = stash.italic_global;
            self.options.italic_dynamic = stash.italic_dynamic;
            // println!("【text2】====设置脚注颜色: override_color={:?}", self.format_state.override_color);
        }
    }

    /// 保存左列样式
    pub fn left_stash(&mut self) {
        self.options.stash_style_left_column = Some(StyleStash {
            bold_italic: self.format_state.bold_italic,
            bold: self.format_state.bold,
            italic: self.format_state.italic,
            underline: self.format_state.underline,
            override_color: self.format_state.override_color.clone(),
            italic_global: self.options.italic_global,
            italic_dynamic: self.options.italic_dynamic,
            current_color: self.format_state.current_color.clone(),
        });
        self.reset_format();
    }

    /// 恢复左列样式
    pub fn left_pop(&mut self) {
        if let Some(stash) = &self.options.stash_style_left_column {
            self.format_state.bold_italic = stash.bold_italic;
            self.format_state.bold = stash.bold;
            self.format_state.italic = stash.italic;
            self.format_state.underline = stash.underline;
            self.format_state.override_color = stash.override_color.clone();
            self.format_state.current_color = stash.current_color.clone();
            self.options.italic_global = stash.italic_global;
            self.options.italic_dynamic = stash.italic_dynamic;
            println!(
                "【text2】====1设置脚注颜色: override_color={:?}",
                self.format_state.override_color
            );
        }
    }

    /// 保存右列样式
    pub fn right_stash(&mut self) {
        self.options.stash_style_right_column = Some(StyleStash {
            bold_italic: self.format_state.bold_italic,
            bold: self.format_state.bold,
            italic: self.format_state.italic,
            underline: self.format_state.underline,
            override_color: self.format_state.override_color.clone(),
            italic_global: self.options.italic_global,
            italic_dynamic: self.options.italic_dynamic,
            current_color: self.format_state.current_color.clone(),
        });
        self.reset_format();
    }

    /// 恢复右列样式
    pub fn right_pop(&mut self) {
        if let Some(stash) = &self.options.stash_style_right_column {
            self.format_state.bold_italic = stash.bold_italic;
            self.format_state.bold = stash.bold;
            self.format_state.italic = stash.italic;
            self.format_state.underline = stash.underline;
            self.format_state.override_color = stash.override_color.clone();
            self.format_state.current_color = stash.current_color.clone();
            self.options.italic_global = stash.italic_global;
            self.options.italic_dynamic = stash.italic_dynamic;
            println!(
                "【text2】====4设置脚注颜色: override_color={:?}",
                self.format_state.override_color
            );
        }
    }

    /// 添加文档样式定义
    pub fn add_document_styles(&mut self, spacing: &ParagraphSpacing) {
        // 获取打印配置
        let print = &self.options.print_profile;

        // 使用adapter中的统一函数
        // docx-rs fix

        // 计算各种缩进值（参考原 TypeScript 项目）
        let action_indent = convert_inches_to_twip(print.action.feed - print.left_margin);
        let scene_indent = convert_inches_to_twip(print.scene_heading.feed - print.left_margin);
        let character_indent = convert_inches_to_twip(print.character.feed - print.left_margin);
        let dialogue_indent = convert_inches_to_twip(print.dialogue.feed - print.left_margin);
        let parenthetical_indent =
            convert_inches_to_twip(print.parenthetical.feed - print.left_margin);

        // 创建样式定义
        let mut styles = crate::docx::adapter::docx::Styles::new();

        // 添加段落样式
        // section 样式
        let mut section_style = crate::docx::adapter::docx::ParagraphStyle::new();
        section_style.id = Some("section".to_string());
        section_style.name = Some("Section".to_string());
        section_style.based_on = Some("Normal".to_string());
        section_style.next = Some("Normal".to_string());

        // 设置 section 样式的运行属性
        let mut section_run = crate::docx::adapter::docx::RunStyle::new();
        if let Some(color) = &print.section.color {
            section_run.color = Some(color.clone());
        }
        section_style.run = Some(section_run);
        section_style.spacing = Some(spacing.clone());

        styles.paragraph_styles.push(section_style);

        // scene 样式（场景头）
        let mut scene_style = crate::docx::adapter::docx::ParagraphStyle::new();
        scene_style.id = Some("scene".to_string());
        scene_style.name = Some("Scene".to_string());
        scene_style.based_on = Some("Normal".to_string());
        scene_style.next = Some("Normal".to_string());
        scene_style.indent = Some(crate::docx::adapter::docx::ParagraphIndent {
            left: Some(scene_indent),
            right: Some(scene_indent),
            first_line: None,
        });
        scene_style.spacing = Some(spacing.clone());
        styles.paragraph_styles.push(scene_style);

        // action 样式
        let mut action_style = crate::docx::adapter::docx::ParagraphStyle::new();
        action_style.id = Some("action".to_string());
        action_style.name = Some("Action".to_string());
        action_style.based_on = Some("Normal".to_string());
        action_style.next = Some("Normal".to_string());
        action_style.indent = Some(crate::docx::adapter::docx::ParagraphIndent {
            left: Some(action_indent),
            right: Some(action_indent),
            first_line: None,
        });
         let mut action_run = crate::docx::adapter::docx::RunStyle::new();
        // action_style.color = Some(print.note.color.clone());
        // action_run.size = Some((print.font_size) as usize);
        action_run.font = self.run_normal.font.clone();
        action_style.run = Some(action_run);
        action_style.spacing = Some(spacing.clone());
        styles.paragraph_styles.push(action_style);

        // character 样式
        let mut character_style = crate::docx::adapter::docx::ParagraphStyle::new();
        character_style.id = Some("character".to_string());
        character_style.name = Some("Character".to_string());
        character_style.based_on = Some("Normal".to_string());
        character_style.next = Some("Normal".to_string());
        character_style.indent = Some(crate::docx::adapter::docx::ParagraphIndent {
            left: Some(character_indent),
            right: Some(character_indent),
            first_line: None,
        });
        character_style.spacing = Some(spacing.clone());
        styles.paragraph_styles.push(character_style);

        // dial 样式（对话）
        let mut dial_style = crate::docx::adapter::docx::ParagraphStyle::new();
        dial_style.id = Some("dial".to_string());
        dial_style.name = Some("Dialogue".to_string());
        dial_style.based_on = Some("Normal".to_string());
        dial_style.next = Some("Normal".to_string());
        dial_style.indent = Some(crate::docx::adapter::docx::ParagraphIndent {
            left: Some(dialogue_indent),
            right: Some(dialogue_indent),
            first_line: None,
        });
        dial_style.spacing = Some(spacing.clone());
        styles.paragraph_styles.push(dial_style);

        // parenthetical 样式（括号动作）
        let mut parenthetical_style = crate::docx::adapter::docx::ParagraphStyle::new();
        parenthetical_style.id = Some("parenthetical".to_string());
        parenthetical_style.name = Some("Parenthetical".to_string());
        parenthetical_style.based_on = Some("Normal".to_string());
        parenthetical_style.next = Some("Normal".to_string());
        parenthetical_style.indent = Some(crate::docx::adapter::docx::ParagraphIndent {
            left: Some(parenthetical_indent),
            right: Some(parenthetical_indent),
            first_line: None,
        });
        parenthetical_style.spacing = Some(spacing.clone());
        styles.paragraph_styles.push(parenthetical_style);

        // notes 样式（注释）
        let mut notes_style = crate::docx::adapter::docx::ParagraphStyle::new();
        notes_style.id = Some("notes".to_string());
        notes_style.name = Some("Notes".to_string());
        notes_style.based_on = Some("Normal".to_string());
        notes_style.next = Some("Normal".to_string());

        // 设置 notes 样式的运行属性
        let mut notes_run = crate::docx::adapter::docx::RunStyle::new();
        // notes_run.color = Some(print.note.color.clone());
        notes_run.size = Some((print.font_size) as usize);
        if print.note.italic {
            // 注意：这里需要在 RunStyle 中添加 italic 字段
            // 暂时跳过斜体设置
        }
        notes_style.run = Some(notes_run);
        notes_style.indent = Some(crate::docx::adapter::docx::ParagraphIndent {
            left: Some(action_indent),
            right: Some(action_indent),
            // first_line: Some(convert_inches_to_twip(2.0 * print.font_width)),
            first_line: None,
        });
        // 注释使用相同的算法，但基于注释字体大小

        notes_style.spacing = Some(
            spacing
                .clone()
                .line(convert_inches_to_twip(print.note_line_height))
                .line_rule(crate::docx::adapter::LineRuleType::AtLeast),
        );
        styles.paragraph_styles.push(notes_style);

        //
        let mut notes_ref_style = crate::docx::adapter::docx::CharacterStyle::new();
        notes_ref_style.id = Some("FootnoteReference".to_string());
        notes_ref_style.name = Some("FootnoteReference".to_string());
        notes_ref_style.based_on = Some("Normal".to_string());
        let mut notes_ref_run = crate::docx::adapter::docx::RunStyle::new();
        // notes_ref_run.color = Some(print.note.color.clone());
        notes_ref_run.size = Some((print.font_size * 2.0 * 1.45) as usize);
        // notes_ref_run.superScript = Some(false);
        notes_ref_style.run = Some(notes_ref_run);
        styles.character_styles.push(notes_ref_style);

        // 将样式添加到文档
        self.doc.options.styles = Some(styles);
    }

    /// 完成双对话处理 - 直接使用全局缓存
    pub fn finish_double_dial(
        &mut self,
        section_main: &mut crate::docx::adapter::docx::Section,
        print: &PrintProfile,
        spacing: &ParagraphSpacing,
    ) {
        println!("【finish_double_dial】开始处理双对话，左侧缓存: {}, 右侧缓存: {}, 全局左侧: {}, 全局右侧: {}",
            self.last_dial_gr_left.is_some(), self.last_dial_gr_right.is_some(),
            self.last_dial_table_left.len(), self.last_dial_table_right.len());

        // 处理左侧对话缓存，添加到全局表格缓存
        if let Some(dial_gr_left) = self.last_dial_gr_left.take() {
            let mut paragraph =
                crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
            paragraph.style(&dial_gr_left.style);

            if let Some(left) = dial_gr_left.indent_left {
                paragraph.indent(left);
            }
            if let Some(right) = dial_gr_left.indent_right {
                paragraph.indent_right(right);
            }

            for run in dial_gr_left.children {
                paragraph.add_text_run(run);
            }

            self.last_dial_table_left.push(paragraph);
        }

        // 处理右侧对话缓存，添加到全局表格缓存
        if let Some(dial_gr_right) = self.last_dial_gr_right.take() {
            let mut paragraph =
                crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
            paragraph.style(&dial_gr_right.style);

            if let Some(left) = dial_gr_right.indent_left {
                paragraph.indent(left);
            }
            if let Some(right) = dial_gr_right.indent_right {
                paragraph.indent_right(right);
            }

            for run in dial_gr_right.children {
                paragraph.add_text_run(run);
            }

            self.last_dial_table_right.push(paragraph);
        };

        // 如果有双对话内容，创建表格 - 使用局部变量，恢复原项目逻辑
        if !self.last_dial_table_left.is_empty() || !self.last_dial_table_right.is_empty() {
            println!(
                "【finish_double_dial】创建双对话表格，左侧段落: {}, 右侧段落: {}",
                self.last_dial_table_left.len(),
                self.last_dial_table_right.len()
            );
            // 计算双对话表格的列宽（参考原项目逻辑）
            let inner_width_twip =
                convert_inches_to_twip(print.page_width - print.left_margin - print.right_margin);
            let action_indent = convert_inches_to_twip(print.action.feed - print.left_margin);
            let dial_double_tab_column_width =
                (inner_width_twip - action_indent - action_indent) / 2;

            // 创建表格
            let mut table = crate::docx::adapter::docx::Table::new();
            table.without_borders(true);
            table.columnWidths(vec![
                dial_double_tab_column_width as usize,
                dial_double_tab_column_width as usize,
            ]);

            // 设置表格缩进
            table.indent = Some(crate::docx::adapter::docx::TableIndent {
                width_type: crate::docx::adapter::WidthType::DXA,
                size: action_indent,
            });

            // 设置无边框
            table.borders = Some(crate::docx::adapter::docx::TableBorders {
                top: Some(crate::docx::adapter::docx::TableBorder {
                    size: 0,
                    color: "auto".to_string(),
                    style: "none".to_string(),
                }),
                bottom: Some(crate::docx::adapter::docx::TableBorder {
                    size: 0,
                    color: "auto".to_string(),
                    style: "none".to_string(),
                }),
                left: Some(crate::docx::adapter::docx::TableBorder {
                    size: 0,
                    color: "auto".to_string(),
                    style: "none".to_string(),
                }),
                right: Some(crate::docx::adapter::docx::TableBorder {
                    size: 0,
                    color: "auto".to_string(),
                    style: "none".to_string(),
                }),
                inside_h: Some(crate::docx::adapter::docx::TableBorder {
                    size: 0,
                    color: "auto".to_string(),
                    style: "none".to_string(),
                }),
                inside_v: Some(crate::docx::adapter::docx::TableBorder {
                    size: 0,
                    color: "auto".to_string(),
                    style: "none".to_string(),
                }),
            });

            // 创建表格行
            let mut row = crate::docx::adapter::docx::TableRow::new();

            // 左列
            let mut left_cell = crate::docx::adapter::docx::TableCell::new();
            left_cell.width = Some(crate::docx::adapter::docx::TableWidth {
                width_type: crate::docx::adapter::WidthType::DXA,
                size: dial_double_tab_column_width,
            });
            left_cell.children = self.last_dial_table_left.clone();

            // 右列
            let mut right_cell = crate::docx::adapter::docx::TableCell::new();
            right_cell.width = Some(crate::docx::adapter::docx::TableWidth {
                width_type: crate::docx::adapter::WidthType::DXA,
                size: dial_double_tab_column_width,
            });
            right_cell.children = self.last_dial_table_right.clone();

            // 添加单元格到行
            row.cells = vec![left_cell, right_cell];

            // 添加行到表格
            table.rows.push(row);

            // 添加表格到section
            section_main
                .children
                .push(crate::docx::adapter::docx::SectionChild::Table(table));

            // 如果只有左侧有对话，表格后补一个空行 - 参考原项目逻辑
            if self.last_dial_table_right.is_empty() {
                let mut empty_paragraph =
                    crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
                empty_paragraph.add_text_run(crate::docx::adapter::docx::TextRun::new(""));
                section_main
                    .children
                    .push(crate::docx::adapter::docx::SectionChild::Paragraph(
                        empty_paragraph,
                    ));
            }

            // 清空局部缓存 - 参考原项目逻辑
            self.last_dial_table_left.clear();
            self.last_dial_table_right.clear();
        }
    }

    /// 完成中文格式对话的第一行处理 - 直接操作全局缓存
    pub fn finish_china_dial_first(
        &mut self,
        section_main: &mut crate::docx::adapter::docx::Section,
        spacing: &ParagraphSpacing,
    ) {
        // 处理普通对话缓存
        if let Some(dial_gr) = self.last_dial_gr.take() {
            let mut paragraph =
                crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
            paragraph.style(&dial_gr.style);

            if let Some(left) = dial_gr.indent_left {
                paragraph.indent(left);
            }
            if let Some(right) = dial_gr.indent_right {
                paragraph.indent_right(right);
            }

            for run in dial_gr.children {
                paragraph.add_text_run(run);
            }

            section_main
                .children
                .push(crate::docx::adapter::docx::SectionChild::Paragraph(
                    paragraph,
                ));
        }

        // 处理左侧对话缓存，添加到全局表格缓存 - 修复关键问题
        if let Some(dial_gr_left) = self.last_dial_gr_left.take() {
            println!("【finish_china_dial_first】处理左侧对话缓存，添加到全局表格缓存");
            let mut paragraph =
                crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
            paragraph.style(&dial_gr_left.style);

            if let Some(left) = dial_gr_left.indent_left {
                paragraph.indent(left);
            }
            if let Some(right) = dial_gr_left.indent_right {
                paragraph.indent_right(right);
            }

            for run in dial_gr_left.children {
                paragraph.add_text_run(run);
            }

            // 添加到全局表格缓存 - 修复关键问题
            self.last_dial_table_left.push(paragraph);
        }

        // 处理右侧对话缓存，添加到全局表格缓存 - 修复关键问题
        if let Some(dial_gr_right) = self.last_dial_gr_right.take() {
            println!("【finish_china_dial_first】处理右侧对话缓存，添加到全局表格缓存");
            let mut paragraph =
                crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
            paragraph.style(&dial_gr_right.style);

            if let Some(left) = dial_gr_right.indent_left {
                paragraph.indent(left);
            }
            if let Some(right) = dial_gr_right.indent_right {
                paragraph.indent_right(right);
            }

            for run in dial_gr_right.children {
                paragraph.add_text_run(run);
            }

            // 添加到全局表格缓存 - 修复关键问题
            self.last_dial_table_right.push(paragraph);
        }
    }
    /// 格式化文本
    pub fn format_text(&mut self, text: &str, options: &HashMap<String, String>) -> Vec<TextRun> {
        // 保存当前格式状态
        self.global_stash();
        // 处理文本
        let result = self.text2(text, options, None, None);
        // 不恢复格式状态，让特殊字符的格式状态保持
        self.global_pop();
        result
    }

    /// 处理文本
    pub fn text2(
        &mut self,
        text: &str,
        options: &HashMap<String, String>,
        mut current_line_notes: Option<&mut Vec<Note>>,
        mut notes_page: Option<&mut Vec<Vec<Vec<Note>>>>,
    ) -> Vec<TextRun> {
        // 使用静态变量记录调用次数
        use std::sync::atomic::{AtomicUsize, Ordering};
        static CALL_COUNT: AtomicUsize = AtomicUsize::new(0);

        let current_count = CALL_COUNT.fetch_add(1, Ordering::SeqCst);

        // 只打印前30次调用的信息
        let _should_print = current_count < 30;

        // 调试：打印脚注状态
        // if should_print && (self.current_note.page_idx > -1 || text.contains("↺") || text.contains("↻") || text.contains("இ") || text.contains("晚霞") || text.contains("序言")) {
        //     println!("【text2】调用 #{}: 输入=\"{}\", 脚注状态: page_idx={}, force_note_orig={}, override_color={:?}",
        //         current_count + 1, text, self.current_note.page_idx, self.force_note_orig, self.format_state.override_color);
        // }

        // if should_print && (text.contains("晚霞") || text.contains("序言") || text.contains("@JANE") || text.contains("水淀粉去")) {
        //     println!("【text2】调用次数 #{}: 输入文本: \"{}\"", current_count + 1, text);
        // }

        // 使用 fountain_constants.rs 中定义的样式标记字符
        use crate::utils::fountain_constants::FountainConstants;
        let style_chars = FountainConstants::style_chars();

        let char_note_begin_ext = style_chars.get("note_begin_ext").unwrap();
        let char_note_begin = style_chars.get("note_begin").unwrap();
        let char_note_end = style_chars.get("note_end").unwrap();
        let char_italic = style_chars.get("italic").unwrap();
        let char_bold = style_chars.get("bold").unwrap();
        let char_bold_italic = style_chars.get("bold_italic").unwrap();
        let char_underline = style_chars.get("underline").unwrap();
        let char_style_left_stash = style_chars.get("style_left_stash").unwrap();
        let char_style_left_pop = style_chars.get("style_left_pop").unwrap();
        let char_style_right_stash = style_chars.get("style_right_stash").unwrap();
        let char_style_right_pop = style_chars.get("style_right_pop").unwrap();
        let char_style_global_stash = style_chars.get("style_global_stash").unwrap();
        let char_style_global_pop = style_chars.get("style_global_pop").unwrap();
        let char_style_global_clean = style_chars.get("style_global_clean").unwrap();
        let char_italic_global_begin = style_chars.get("italic_global_begin").unwrap();
        let char_italic_global_end = style_chars.get("italic_global_end").unwrap();
        let char_all = style_chars.get("all").unwrap();

        // 获取颜色，优先使用覆盖颜色
        let mut color = options
            .get("color")
            .unwrap_or(&"#000000".to_string())
            .clone();
        if let Some(override_color) = &self.format_state.override_color {
            color = override_color.clone();
        }

        // println!("【text2】初始化颜色: override_color={:?}, color={}", self.format_state.override_color, color);

        // 更新当前颜色
        self.format_state.current_color = color.clone();

        // 处理中文格式的三角符号
        let mut text = text.to_string();

        // 页面底部notes打印模式
        let mut catch_notes = false;
        let mut pushed = false;

        if current_line_notes.is_some() && notes_page.is_some() {
            catch_notes = true; // 页面底部notes打印模式
                                // if should_print && (text.contains("晚霞") || text.contains("序言") || text.contains("@JANE") || text.contains("水淀粉去")) {
                                //     println!("【text2】设置 catch_notes = true，当前脚注状态: page_idx={}", self.current_note.page_idx);
                                // }
            if self.current_note.page_idx >= 0 {
                // 如果正在处理notes，并且收集到底部，本行为notes开始内容
                if (self.china_format == 1 || self.china_format == 3) && text.starts_with('△') {
                    // 安全地移除第一个字符，处理 Unicode 字符
                    if let Some(first_char) = text.chars().next() {
                        text = text[first_char.len_utf8()..].to_string();
                    }
                    self.cache_triangle = true;
                }
            } else {
                self.cache_triangle = false;
            }
        }

        if (self.china_format == 1 || self.china_format == 3)
            && text.starts_with('△')
            && self.force_note_orig
        {
            // 安全地移除第一个字符，处理 Unicode 字符
            if let Some(first_char) = text.chars().next() {
                text = text[first_char.len_utf8()..].to_string();
            }
        }

        // 处理注释斜体
        let note_config = &self.options.print_profile.note;
        if note_config.italic {
            text = text
                .replace(
                    char_note_begin_ext,
                    &format!("{}{}", char_italic, char_note_begin_ext),
                )
                .replace(
                    char_note_begin,
                    &format!("{}{}", char_italic, char_note_begin),
                )
                .replace(char_note_end, &format!("{}{}", char_note_end, char_italic));
        }

        // 处理链接
        // 定义链接结构体
        #[derive(Debug, Clone)]
        struct Link {
            start: usize,
            length: usize,
            url: String,
        }

        let links: Vec<Link> = Vec::new();

        // 处理链接
        if let Some(links_option) = options.get("links") {
            if links_option == "true" {
                // 在完整实现中，这里应该使用正则表达式处理链接
                // 例如：regex.link.exec(text)
                // 但在简化版中，我们省略这部分
            }
        }

        // 分割文本以处理格式化
        let mut split_for_formatting = Vec::new();

        // 根据链接分割文本或直接添加整个文本
        if links.is_empty() {
            // 如果没有链接，直接添加整个文本
            split_for_formatting.push(text.clone());
        } else {
            // 根据链接分割文本
            // "This is a link: google.com and this is after"
            // |--------------|----------| - - - - - - - |
            let mut prevlink = 0;
            for link in &links {
                split_for_formatting.push(text[prevlink..link.start].to_string());
                split_for_formatting.push(text[link.start..link.start + link.length].to_string());
                prevlink = link.start + link.length;
            }

            // 添加剩余的文本
            // "This is a link: google.com and this is after"
            // | - - - - - - -| - - - - -|----------------|
            let leftover = text[prevlink..].to_string();
            if !leftover.is_empty() {
                split_for_formatting.push(leftover);
            }
        }

        // 进一步分割以处理粗体、斜体、下划线等
        let mut i = 0;
        while i < split_for_formatting.len() {
            let elem = split_for_formatting[i].clone();

            // 直接按字符处理，确保每个特殊字符单独成为一个元素
            let mut parts = Vec::new();
            let mut current_text = String::new();

            // 遍历字符串中的每个字符
            for c in elem.chars() {
                // 检查当前字符是否是特殊字符
                let is_special = char_all.contains(c);

                if is_special {
                    // 如果当前积累的普通文本不为空，先添加它
                    if !current_text.is_empty() {
                        parts.push(current_text);
                        current_text = String::new();
                    }
                    // 将特殊字符作为单独的元素添加
                    parts.push(c.to_string());
                } else {
                    // 普通字符，添加到当前文本
                    current_text.push(c);
                }
            }

            // 添加最后一段普通文本（如果有）
            if !current_text.is_empty() {
                parts.push(current_text);
            }

            // 如果有分割出的部分
            if parts.len() > 1
                || (parts.len() == 1 && char_all.contains(parts[0].chars().next().unwrap_or(' ')))
            {
                // 替换当前元素
                split_for_formatting.remove(i);
                for (j, part) in parts.iter().enumerate() {
                    if !part.is_empty() {
                        split_for_formatting.insert(i + j, part.clone());
                    }
                }
                i += parts.len();
            } else {
                i += 1;
            }
        }

        // 打印分割后的文本，用于调试
        // if should_print {
        //     println!("【text2】分割后的文本片段:");
        //     for (idx, part) in split_for_formatting.iter().enumerate() {
        //         println!("【text2】片段 #{}: '{}'", idx + 1, part);
        //     }
        // }

        // 处理分割后的文本
        let mut text_objects = Vec::new();
        // current_index 用于跟踪处理文本的当前位置
        let mut current_index = 0;

        for elem in split_for_formatting {
            if elem == *char_style_global_clean {
                self.reset_format();
                // 重置颜色
                color = options
                    .get("color")
                    .unwrap_or(&"#000000".to_string())
                    .clone();
                self.format_state.override_color = None;
                // println!("【text2】清空note颜色1: override_color={:?}, color={}", self.format_state.override_color, color);
            } else if elem == *char_style_global_stash {
                self.global_stash();
                // 重置颜色
                color = options
                    .get("color")
                    .unwrap_or(&"#000000".to_string())
                    .clone();
                self.format_state.override_color = None;
                // println!("【text2】清空note颜色2: override_color={:?}, color={}", self.format_state.override_color, color);
            } else if elem == *char_style_global_pop {
                self.global_pop();
            } else if elem == *char_style_left_stash {
                self.left_stash();
                // 重置颜色
                color = options
                    .get("color")
                    .unwrap_or(&"#000000".to_string())
                    .clone();
                self.format_state.override_color = None;
                // println!("【text2】清空note颜色3: override_color={:?}, color={}", self.format_state.override_color, color);
            } else if elem == *char_style_left_pop {
                self.left_pop();
            } else if elem == *char_style_right_stash {
                self.right_stash();
                // 重置颜色
                color = options
                    .get("color")
                    .unwrap_or(&"#000000".to_string())
                    .clone();
                self.format_state.override_color = None;
                // println!("【text2】清空note颜色4: override_color={:?}, color={}", self.format_state.override_color, color);
            } else if elem == *char_style_right_pop {
                self.right_pop();
            } else if elem == *char_italic_global_begin {
                self.options.italic_dynamic = self.format_state.italic;
                self.options.italic_global = true;
                self.format_state.italic = true;
            } else if elem == *char_italic_global_end {
                self.format_state.italic = self.options.italic_dynamic;
                self.options.italic_global = false;
            } else if elem == *char_bold_italic {
                if catch_notes && self.current_note.page_idx > -1 {
                    if !pushed {
                        self.current_note.note.text.push(elem.clone());
                        pushed = true;
                    } else if let Some(last) = self.current_note.note.text.last_mut() {
                        *last += &elem;
                    }
                } else {
                    self.format_state.bold_italic = !self.format_state.bold_italic;
                }
            } else if elem == *char_bold {
                if catch_notes && self.current_note.page_idx > -1 {
                    if !pushed {
                        self.current_note.note.text.push(elem.clone());
                        pushed = true;
                    } else if let Some(last) = self.current_note.note.text.last_mut() {
                        *last += &elem;
                    }
                } else {
                    // 切换粗体状态
                    self.format_state.bold = !self.format_state.bold;
                    // println!("【text2】切换粗体状态: {}", self.format_state.bold);
                }
            } else if elem == *char_italic {
                if catch_notes && self.current_note.page_idx > -1 {
                    if !pushed {
                        self.current_note.note.text.push(elem.clone());
                        pushed = true;
                    } else if let Some(last) = self.current_note.note.text.last_mut() {
                        *last += &elem;
                    }
                } else {
                    if self.options.italic_global {
                        self.options.italic_dynamic = !self.options.italic_dynamic;
                    } else {
                        self.format_state.italic = !self.format_state.italic;
                    }
                }
            } else if elem == *char_underline {
                if catch_notes && self.current_note.page_idx > -1 {
                    if !pushed {
                        self.current_note.note.text.push(elem.clone());
                        pushed = true;
                    } else if let Some(last) = self.current_note.note.text.last_mut() {
                        *last += &elem;
                    }
                } else {
                    // 切换下划线状态
                    self.format_state.underline = !self.format_state.underline;
                    // println!("【text2】切换下划线状态: {}", self.format_state.underline);
                }
            } else if elem == *char_note_end {
                // println!("【text2】处理 note_end 符号 ↻");
                // println!("【text2】重置前的颜色: override_color={:?}, color={}", self.format_state.override_color, color);

                if catch_notes && !self.force_note_orig {
                    if let Some(current_line_notes_ref) = &mut current_line_notes {
                        if !current_line_notes_ref.is_empty() {
                            let last_index = current_line_notes_ref.len() - 1;
                            current_line_notes_ref[last_index].text =
                                self.current_note.note.text.clone();
                        }
                    }

                    if let Some(notes_page_ref) = &mut notes_page {
                        if self.current_note.page_idx >= 0
                            && self.current_note.page_idx < notes_page_ref.len() as i32
                        {
                            let page_idx = self.current_note.page_idx as usize;
                            if !notes_page_ref[page_idx].is_empty() {
                                let token_row = notes_page_ref[page_idx].len() - 1;
                                if !notes_page_ref[page_idx][token_row].is_empty() {
                                    let note_idx = notes_page_ref[page_idx][token_row].len() - 1;
                                    notes_page_ref[page_idx][token_row][note_idx].text =
                                        self.current_note.note.text.clone();
                                }
                            }
                        }
                    }
                }

                // 无论是否收集脚注，都要重置脚注状态
                self.current_note.page_idx = -1;

                // 清除脚注状态 - 参考原项目第597行
                self.force_note_orig = false;
                self.format_state.override_color = None;
                // println!("【text2】清空note颜色5: override_color={:?}, color={}", self.format_state.override_color, color);

                // 重置颜色
                color = options
                    .get("color")
                    .unwrap_or(&"#000000".to_string())
                    .clone();

                // println!("【text2】重置后的颜色: override_color={:?}, color={}", self.format_state.override_color, color);
            } else if elem == *char_note_begin_ext {
                // 强制在原位置打印 note
                let note_config = &self.options.print_profile.note;
                // println!("【text2】处理 note_begin_ext 符号 இ");
                // println!("【text2】设置前的颜色: override_color={:?}, color={}", self.format_state.override_color, color);

                // 设置覆盖颜色，与原始项目保持一致
                self.format_state.override_color = Some(note_config.color.clone());
                // 同时更新当前颜色变量，确保它被应用到后续的所有文本中
                color = note_config.color.clone();

                // println!("【text2】设置后的颜色: override_color={:?}, color={}", self.format_state.override_color, color);
                self.force_note_orig = true;
                // println!("【text2】=======2设置脚注颜色: override_color={:?}, color={}", self.format_state.override_color, color);
            } else {
                // 处理普通文本和注释开始标记
                if elem == *char_note_begin {
                    let note_config = &self.options.print_profile.note;
                    // println!("【text2】处理 note_begin 符号 ↺");
                    // println!("【text2】设置前的颜色: override_color={:?}, color={}", self.format_state.override_color, color);

                    // 设置覆盖颜色，与原始项目保持一致
                    self.format_state.override_color = Some(note_config.color.clone());
                    // 同时更新当前颜色变量，确保它被应用到后续的所有文本中
                    color = note_config.color.clone();

                    // println!("【text2】=======设置脚注颜色: override_color={:?}, color={}", self.format_state.override_color, color);

                    if catch_notes {
                        self.notes_len += 1;

                        self.current_note = CurrentNote {
                            page_idx: 0,
                            note: Note {
                                no: self.notes_len,
                                text: vec!["".to_string()],
                            },
                        };

                        if let Some(notes_page_ref) = &mut notes_page {
                            if notes_page_ref.len() <= self.current_note.page_idx as usize {
                                notes_page_ref
                                    .resize_with(self.current_note.page_idx as usize + 1, Vec::new);
                            }

                            let page_idx = self.current_note.page_idx as usize;
                            if notes_page_ref[page_idx].is_empty() {
                                notes_page_ref[page_idx].push(Vec::new());
                            }

                            let token_row = notes_page_ref[page_idx].len() - 1;
                            notes_page_ref[page_idx][token_row].push(Note {
                                no: self.notes_len,
                                text: vec!["".to_string()],
                            });
                        }

                        if let Some(current_line_notes_ref) = &mut current_line_notes {
                            current_line_notes_ref.push(Note {
                                no: self.notes_len,
                                text: vec!["".to_string()],
                            });
                        }

                        pushed = true;
                    }

                    // 添加脚注引用 - 参考原项目逻辑
                    if catch_notes {
                        // 当收集脚注时，创建真正的脚注引用

                        // 将脚注文本转换为格式化的TextRun
                        let footnote_runs = Vec::new();

                        let footnote_ref = crate::docx::adapter::docx::TextRun::footnote_reference(
                            self.notes_len,
                            footnote_runs,
                            self.run_notes.clone(),
                        );
                        text_objects.push(footnote_ref);
                    }
                    // 当不收集脚注时（catch_notes = false），脚注开始标记不显示，但脚注内容会在原位置显示
                    // 注意：无论是否收集脚注，都需要设置override_color以确保脚注内容有正确的样式
                } else if !elem.is_empty() {
                    // 处理普通文本
                    // 参考原项目逻辑：先判断是否需要处理脚注相关逻辑
                    let mut draw = true;
                    let mut elem_to_draw = elem.clone();

                    if catch_notes {
                        // if should_print && (elem.contains("晚霞") || elem.contains("序言") || elem.contains("@JANE") || elem.contains("水淀粉去")) {
                        //     println!("【text2】catch_notes=true, page_idx={}, elem=\"{}\"", self.current_note.page_idx, elem);
                        // }
                        if self.current_note.page_idx >= 0 {
                            // 在脚注中，收集脚注内容
                            if !pushed {
                                self.current_note.note.text.push(elem.clone());
                                pushed = true;
                                // if should_print {
                                //     println!("【text2】收集脚注内容（新行）: \"{}\"", elem);
                                // }
                            } else if let Some(last) = self.current_note.note.text.last_mut() {
                                *last += &elem;
                                // if should_print {
                                //     println!("【text2】收集脚注内容（追加）: \"{}\"，当前行内容: \"{}\"", elem, last);
                                // }
                            }

                            // 根据原项目逻辑：当收集脚注时，设置 elem = '' 和 draw = false
                            // 只有当 force_note_orig=true 时才在原位置显示脚注内容
                            if !self.force_note_orig {
                                elem_to_draw = String::new();
                                draw = false;
                                // if should_print {
                                //     println!("【text2】脚注内容不在原位置显示（force_note_orig=false）: \"{}\"", elem);
                                // }
                            } else {
                                // if should_print {
                                //     println!("【text2】脚注内容在原位置显示（force_note_orig=true）: \"{}\"", elem);
                                // }
                            }
                        } else {
                            // 不在脚注中，正常处理
                            // if should_print && (elem.contains("晚霞") || elem.contains("序言") || elem.contains("@JANE") || elem.contains("水淀粉去")) {
                            //     println!("【text2】不在脚注中，正常处理: \"{}\"", elem);
                            // }
                            // 处理中文格式的三角符号缓存
                            if self.cache_triangle {
                                elem_to_draw = format!("△{}", elem);
                                self.cache_triangle = false;
                            }
                        }
                    } else {
                        // if should_print && (elem.contains("晚霞") || elem.contains("序言") || elem.contains("@JANE") || elem.contains("水淀粉去")) {
                        //     println!("【text2】catch_notes=false，正常处理: \"{}\"", elem);
                        // }
                    }

                    // 只有当 draw = true 时才处理文本
                    if draw {
                        let character_spacing =
                            if let Some(characterSpacing) = options.get("characterSpacing") {
                                characterSpacing.parse::<f32>().unwrap_or(1.0)
                            } else {
                                self.options.print_profile.character_spacing
                            };
                        let mut font_size = if let Some(font_size) = options.get("fontSize") {
                            font_size.parse::<usize>().unwrap_or(12)
                        } else {
                            (self.options.print_profile.font_size) as usize
                        };

                        // 参考原项目逻辑：如果在脚注中（override_color存在），使用脚注字体大小
                        if self.format_state.override_color.is_some() {
                            font_size = (self.options.print_profile.note_font_size) as usize;
                        }

                        // 检查 options 中的粗体设置
                        let options_bold =
                            options.get("bold").map(|v| v == "true").unwrap_or(false);

                        // 检查 options 中的粗体设置
                        let options_italic =
                            options.get("italic").map(|v| v == "true").unwrap_or(false);

                        // 根据当前格式状态选择预定义的样式，与原项目逻辑保持一致
                        let mut run_props = if self.format_state.bold_italic {
                            // 粗体斜体状态：使用 run_bold_italic
                            self.run_bold_italic.clone()
                        } else if self.format_state.bold || options_bold {
                            // 粗体状态：使用 run_bold
                            self.run_bold.clone()
                        } else if self.format_state.italic
                            || self.options.italic_global
                            || self.options.italic_dynamic
                            || options_italic
                        {
                            // 斜体状态：使用 run_italic
                            self.run_italic.clone()
                        } else if self.format_state.override_color.is_some() {
                            // 脚注状态：使用 run_notes
                            self.run_notes.clone()
                        } else {
                            // 正常状态：使用 run_normal
                            self.run_normal.clone()
                        };

                        // 应用动态属性
                        run_props.size = Some(font_size);
                        run_props.character_spacing =
                            Some(convert_point_to_twip(character_spacing));

                        // 应用下划线
                        if self.format_state.underline {
                            run_props.underline = Some(UnderlineTypeConst::SINGLE);
                        }

                        // 检查是否有链接
                        let mut link_url = None;
                        for link in &links {
                            if link.start <= current_index
                                && current_index < link.start + link.length
                            {
                                link_url = Some(link.url.clone());
                            }
                        }

                        // 如果有链接，添加下划线
                        if link_url.is_some() && run_props.underline.is_none() {
                            run_props.underline = Some(UnderlineTypeConst::SINGLE);
                        }

                        // 与原始项目保持一致的颜色处理逻辑
                        if self.force_note_orig {
                            // 如果 force_note_orig 为 true，表示我们在一个注释块内，应该使用灰色
                            let note_config = &self.options.print_profile.note;
                            color = note_config.color.clone();
                        } else if let Some(override_color) = &self.format_state.override_color {
                            // 如果有覆盖颜色，使用覆盖颜色
                            color = override_color.clone();
                        }

                        // 设置 TextRun 的颜色
                        if !color.is_empty() && color != "#000000" {
                            run_props.color = Some(color.clone());
                        }

                        // 处理普通文本中的换行符
                        let lines: Vec<&str> = elem_to_draw.split('\n').collect();

                        if !lines.is_empty() {
                            // 第一行不添加换行符
                            let run = TextRun::with_props(lines[0], run_props.clone());
                            text_objects.push(run);

                            // 后续行添加换行符（使用 break: 1 属性，与 TypeScript 版本一致）
                            for i in 1..lines.len() {
                                // 创建一个带有 break_before 属性的 TextRun 对象
                                let mut break_props = run_props.clone();
                                break_props.break_before = Some(true);

                                // 添加带有换行符和文本的 TextRun
                                let run = TextRun::with_props(lines[i], break_props);
                                text_objects.push(run);
                            }
                        }
                    }
                }
            }

            current_index += elem.len();
        }

        // if should_print && (text.contains("晚霞") || text.contains("序言") || text.contains("@JANE") || text.contains("水淀粉去")) {
        //     println!("【text2】最终返回 {} 个 TextRun，输入文本: \"{}\"", text_objects.len(), text);
        // }

        text_objects
    }
}

/// 初始化文档
pub async fn init_doc(options: DocxOptions) -> DocxContext {
    // 创建文档上下文
    let mut context = DocxContext::new(options.clone());

    // 设置文档属性
    context.doc.options.creator = "Arming".to_string();
    context.doc.options.description = "My screenplay document".to_string();
    context.doc.options.title = "My Screenplay".to_string();

    // 设置中文格式和空行处理
    let mut china_format = 0;
    let mut rm_blank_line = 0;

    if let Some(metadata) = &options.metadata {
        if metadata.contains_key("print") {
            if let Some(china_format_str) = metadata.get("print.chinaFormat") {
                if let Ok(cf) = china_format_str.parse::<i32>() {
                    china_format = cf;
                }
            }
            if let Some(rm_blank_line_str) = metadata.get("print.rmBlankLine") {
                if let Ok(rbl) = rm_blank_line_str.parse::<i32>() {
                    rm_blank_line = rbl;
                }
            }
        }
    }

    // 更新上下文
    context.china_format = china_format;
    context.rm_blank_line = rm_blank_line;

    // 重置格式状态
    context.reset_format();

    context
}
/// 清理文本中的格式标记
fn if_reset_format(input: String, line: &Line) -> String {
    // 检查行类型是否需要重置格式
    if line.token_type == "character"
        || line.token_type == "scene_heading"
        || line.token_type == "synopsis"
        || line.token_type == "centered"
        || line.token_type == "section"
        || line.token_type == "transition"
        || line.token_type == "lyric"
    {
        // 在 Rust 实现中，我们假设所有这些类型的行都需要重置格式
        // 因为当前的 Line 结构体没有 is_wrap 字段
        // 如果将来需要更精确的控制，可以在 Line 结构体中添加 is_wrap 字段
        use crate::utils::fountain_constants::FountainConstants;
        let style_chars = FountainConstants::style_chars();
        let style_global_clean = style_chars.get("style_global_clean").unwrap().to_string(); // "⇜"

        return add_tag_after_broken_note(input, style_global_clean);
    }

    input
}
fn add_tag_after_broken_note(input: String, tag: String) -> String {
    use crate::utils::fountain_constants::FountainConstants;
    let style_chars = FountainConstants::style_chars();

    let note_begin = style_chars.get("note_begin").unwrap(); // "↺"
    let note_end = style_chars.get("note_end").unwrap(); // "↻"

    // 查找 note_end 的位置
    if let Some(iend) = input.find(note_end) {
        // 查找 note_begin 的位置
        if let Some(istart) = input.find(note_begin) {
            // 如果 note_end 在 note_begin 之前，说明是破损的注释
            if iend < istart {
                // 在 note_end 之后插入 tag
                let mut result = String::new();
                result.push_str(&input[..iend + note_end.len()]);
                result.push_str(&tag);
                result.push_str(&input[iend + note_end.len()..]);
                return result;
            }
        } else {
            // 没有找到 note_begin，但找到了 note_end，说明是破损的注释
            // 在 note_end 之后插入 tag
            let mut result = String::new();
            result.push_str(&input[..iend + note_end.len()]);
            result.push_str(&tag);
            result.push_str(&input[iend + note_end.len()..]);
            return result;
        }
    }

    // 如果没有找到 note_end 或者注释结构正常，在开头添加 tag
    format!("{}{}", tag, input)
}
fn clear_formatting(text: &str) -> String {
    // 清除所有的格式化标记，如 *粗体*, _斜体_, **粗体**, __斜体__
    let mut result = text.to_string();

    // 清除粗体标记
    let re = regex::Regex::new(r"\*\*(.+?)\*\*").unwrap();
    result = re.replace_all(&result, "$1").to_string();

    let re = regex::Regex::new(r"\*(.+?)\*").unwrap();
    result = re.replace_all(&result, "$1").to_string();

    // 清除斜体标记
    let re = regex::Regex::new(r"__(.+?)__").unwrap();
    result = re.replace_all(&result, "$1").to_string();

    let re = regex::Regex::new(r"_(.+?)_").unwrap();
    result = re.replace_all(&result, "$1").to_string();

    // 清除下划线标记
    let re = regex::Regex::new(r"~~(.+?)~~").unwrap();
    result = re.replace_all(&result, "$1").to_string();

    // 移除内联注释 [[...]] 或 /* ... */
    result = result
        .replace("[[", "")
        .replace("]]", "")
        .replace("/*", "")
        .replace("*/", "");

    result.trim().to_string()
}

/// 将文本转换为内联文本(移除换行符)
fn inline(text: &str) -> String {
    // 将多行文本合并为一行，去除换行符
    let mut result = text.to_string();

    // 替换换行符为空格
    result = result.replace('\n', " ");

    // 替换多个空格为一个空格
    let re = regex::Regex::new(r"\s+").unwrap();
    result = re.replace_all(&result, " ").to_string();

    // 去除首尾空格
    result.trim().to_string()
}

/// 页面尺寸计算结果
#[derive(Debug, Clone)]
struct PageDimensions {
    page_width: i32,
    page_height: i32,
    left_margin: i32,
    right_margin: i32,
    top_margin: i32,
    bottom_margin: i32,
    line_height: i32,
    inner_width: i32,
}

/// 计算页面尺寸和边距
fn calculate_page_dimensions(print: &PrintProfile, line_height: f32) -> PageDimensions {
    let page_width = convert_inches_to_twip(print.page_width);
    let page_height = convert_inches_to_twip(print.page_height);
    let left_margin = convert_inches_to_twip(print.left_margin);
    let right_margin = convert_inches_to_twip(print.right_margin);
    let top_margin = convert_inches_to_twip(print.top_margin);
    let bottom_margin = convert_inches_to_twip(print.bottom_margin);
    let line_h = convert_inches_to_twip(line_height);
    let inner_width = page_width - left_margin - right_margin;

    PageDimensions {
        page_width,
        page_height,
        left_margin,
        right_margin,
        top_margin,
        bottom_margin,
        line_height: line_h,
        inner_width,
    }
}

/// 创建标题页框架配置 - 支持独立重叠的 frame
/// 通过使用不同的锚点和微小的位置偏移来确保每个 frame 真正独立
fn create_title_frame(
    position: &str,
    dimensions: &PageDimensions,
) -> crate::docx::adapter::ParagraphFrame {
    match position {
        "tl" => crate::docx::adapter::ParagraphFrame {
            // 左上角：使用页面锚点，确保独立性
            width: Some(dimensions.inner_width / 3),
            height: Some(0),
            anchor_horizontal: Some(crate::docx::adapter::FrameAnchorType::Page),
            anchor_vertical: Some(crate::docx::adapter::FrameAnchorType::Page),
            x_align: Some(crate::docx::adapter::HorizontalPositionAlign::Left),
            y_align: Some(crate::docx::adapter::VerticalPositionAlign::Top),
            page_height: Some(dimensions.page_height),
            page_width: Some(dimensions.page_width),
            left_margin: Some(dimensions.left_margin),
            right_margin: Some(dimensions.right_margin),
            top_margin: Some(dimensions.top_margin),
            bottom_margin: Some(dimensions.bottom_margin),
            line_height: Some(dimensions.line_height),
        },
        "tc" => crate::docx::adapter::ParagraphFrame {
            // 顶部中央：使用边距锚点，与 tl 区分
            width: Some(dimensions.inner_width / 3),
            height: Some(0),
            anchor_horizontal: Some(crate::docx::adapter::FrameAnchorType::Margin),
            anchor_vertical: Some(crate::docx::adapter::FrameAnchorType::Margin),
            x_align: Some(crate::docx::adapter::HorizontalPositionAlign::Center),
            y_align: Some(crate::docx::adapter::VerticalPositionAlign::Top),
            page_height: Some(dimensions.page_height),
            page_width: Some(dimensions.page_width),
            left_margin: Some(dimensions.left_margin),
            right_margin: Some(dimensions.right_margin),
            top_margin: Some(dimensions.top_margin),
            bottom_margin: Some(dimensions.bottom_margin),
            line_height: Some(dimensions.line_height),
        },
        "tr" => crate::docx::adapter::ParagraphFrame {
            // 右上角：使用文本锚点，与 tl 和 tc 区分
            width: Some(dimensions.inner_width / 3),
            height: Some(0),
            anchor_horizontal: Some(crate::docx::adapter::FrameAnchorType::Text),
            anchor_vertical: Some(crate::docx::adapter::FrameAnchorType::Text),
            x_align: Some(crate::docx::adapter::HorizontalPositionAlign::Right),
            y_align: Some(crate::docx::adapter::VerticalPositionAlign::Top),
            page_height: Some(dimensions.page_height),
            page_width: Some(dimensions.page_width),
            left_margin: Some(dimensions.left_margin),
            right_margin: Some(dimensions.right_margin),
            top_margin: Some(dimensions.top_margin),
            bottom_margin: Some(dimensions.bottom_margin),
            line_height: Some(dimensions.line_height),
        },
        "cc" => crate::docx::adapter::ParagraphFrame {
            // 中央：使用页面锚点，独立于其他位置
            width: Some(dimensions.inner_width),
            height: Some(0),
            anchor_horizontal: Some(crate::docx::adapter::FrameAnchorType::Page),
            anchor_vertical: Some(crate::docx::adapter::FrameAnchorType::Page),
            x_align: Some(crate::docx::adapter::HorizontalPositionAlign::Center),
            y_align: Some(crate::docx::adapter::VerticalPositionAlign::Center),
            page_height: Some(dimensions.page_height),
            page_width: Some(dimensions.page_width),
            left_margin: Some(dimensions.left_margin),
            right_margin: Some(dimensions.right_margin),
            top_margin: Some(dimensions.top_margin),
            bottom_margin: Some(dimensions.bottom_margin),
            line_height: Some(dimensions.line_height),
        },
        "bl" => crate::docx::adapter::ParagraphFrame {
            // 左下角：使用边距锚点，与 br 区分
            width: Some(dimensions.inner_width / 2),
            height: Some(0),
            anchor_horizontal: Some(crate::docx::adapter::FrameAnchorType::Margin),
            anchor_vertical: Some(crate::docx::adapter::FrameAnchorType::Margin),
            x_align: Some(crate::docx::adapter::HorizontalPositionAlign::Left),
            y_align: Some(crate::docx::adapter::VerticalPositionAlign::Bottom),
            page_height: Some(dimensions.page_height),
            page_width: Some(dimensions.page_width),
            left_margin: Some(dimensions.left_margin),
            right_margin: Some(dimensions.right_margin),
            top_margin: Some(dimensions.top_margin),
            bottom_margin: Some(dimensions.bottom_margin),
            line_height: Some(dimensions.line_height),
        },
        "br" => crate::docx::adapter::ParagraphFrame {
            // 右下角：使用文本锚点，与 bl 区分
            width: Some(dimensions.inner_width / 2),
            height: Some(0),
            anchor_horizontal: Some(crate::docx::adapter::FrameAnchorType::Text),
            anchor_vertical: Some(crate::docx::adapter::FrameAnchorType::Text),
            x_align: Some(crate::docx::adapter::HorizontalPositionAlign::Right),
            y_align: Some(crate::docx::adapter::VerticalPositionAlign::Bottom),
            page_height: Some(dimensions.page_height),
            page_width: Some(dimensions.page_width),
            left_margin: Some(dimensions.left_margin),
            right_margin: Some(dimensions.right_margin),
            top_margin: Some(dimensions.top_margin),
            bottom_margin: Some(dimensions.bottom_margin),
            line_height: Some(dimensions.line_height),
        },
        _ => crate::docx::adapter::ParagraphFrame {
            width: Some(dimensions.inner_width),
            height: Some(0),
            anchor_horizontal: Some(crate::docx::adapter::FrameAnchorType::Margin),
            anchor_vertical: Some(crate::docx::adapter::FrameAnchorType::Margin),
            x_align: Some(crate::docx::adapter::HorizontalPositionAlign::Left),
            y_align: Some(crate::docx::adapter::VerticalPositionAlign::Top),
            page_height: Some(dimensions.page_height),
            page_width: Some(dimensions.page_width),
            left_margin: Some(dimensions.left_margin),
            right_margin: Some(dimensions.right_margin),
            top_margin: Some(dimensions.top_margin),
            bottom_margin: Some(dimensions.bottom_margin),
            line_height: Some(dimensions.line_height),
        },
    }
}

/// 获取标题页位置的对齐方式
fn get_title_alignment(position: &str) -> Option<crate::docx::adapter::AlignmentType> {
    match position {
        "tc" | "cc" => Some(crate::docx::adapter::AlignmentType::Center),
        "tr" | "br" => Some(crate::docx::adapter::AlignmentType::Right),
        _ => None, // 默认左对齐
    }
}

/// 创建基础选项映射
fn create_basic_options_map(color: &str) -> HashMap<String, String> {
    let mut options_map = HashMap::new();
    options_map.insert("color".to_string(), color.to_string());
    options_map
}

/// 完成中文格式对话和双对话处理的辅助函数
fn finish_dialogue_processing(
    doc: &mut DocxContext,
    china_format: i32,
    token_type: &str,
    scene_or_section_or_tran_started: bool,
    section_main: &mut crate::docx::adapter::docx::Section,
    section_main_no_page_num: &mut crate::docx::adapter::docx::Section,
    print: &PrintProfile,
    spacing: &ParagraphSpacing,
) {
    // 完成中文格式对话处理
    if china_format > 0 {
        doc.finish_china_dial_first(
            if scene_or_section_or_tran_started {
                section_main
            } else {
                section_main_no_page_num
            },
            spacing,
        );
    }

    // 检查是否需要完成双对话处理
    let has_right_table_content = !doc.last_dial_table_right.is_empty();
    let has_left_or_right_cache =
        doc.last_dial_gr_left.is_some() || doc.last_dial_gr_right.is_some();
    let has_global_table_content =
        !doc.last_dial_table_left.is_empty() || !doc.last_dial_table_right.is_empty();

    if has_right_table_content
        || has_left_or_right_cache
        || has_global_table_content
        || token_type != "separator"
    {
        doc.finish_double_dial(
            if scene_or_section_or_tran_started {
                section_main
            } else {
                section_main_no_page_num
            },
            print,
            spacing,
        );
    }
}

/// 添加段落到相应section并更新行映射的辅助函数
fn add_paragraph_and_update_line_map(
    child: crate::docx::adapter::docx::SectionChild,
    line: &Line,
    scene_or_section_or_tran_started: bool,
    section_main: &mut crate::docx::adapter::docx::Section,
    section_main_no_page_num: &mut crate::docx::adapter::docx::Section,
    line_map: &mut Option<&mut HashMap<usize, LineStruct>>,
    current_sections: &[String],
    current_scene: &str,
    current_page: usize,
    current_duration: f64,
) {
    // 添加段落到相应的 section
    if scene_or_section_or_tran_started {
        section_main.children.push(child);
    } else {
        section_main_no_page_num.children.push(child);
    }

    // 更新行映射
    if let Some(token_line) = line.token {
        if let Some(ref mut lm) = line_map {
            lm.insert(
                token_line,
                LineStruct {
                    sections: current_sections.to_vec(),
                    scene: current_scene.to_string(),
                    page: current_page,
                    cumulative_duration: current_duration as f32,
                },
            );
        }
    }
}

/// 创建文本运行并添加到段落的辅助函数
fn add_text_runs_to_paragraph(
    doc: &mut DocxContext,
    paragraph: &mut crate::docx::adapter::docx::Paragraph,
    text: &str,
    options_map: &HashMap<String, String>,
) {
    let text_runs = doc.text2(text, options_map, None, None);
    for run in text_runs {
        paragraph.add_text_run(run);
    }
}

/// 处理场景编号的辅助函数
fn process_scene_number(scene_number: &str, scenes_numbers: &str) -> (String, String) {
    let scene_text_length = scene_number.chars().count();

    let left_char = if scene_text_length < 3 {
        3 - scene_text_length
    } else {
        0
    };

    let left_scene_number = if scenes_numbers == "both" || scenes_numbers == "left" {
        format!("{}{}", " ".repeat(left_char), scene_number)
    } else {
        String::new()
    };

    let right_scene_number = if scenes_numbers == "both" || scenes_numbers == "right" {
        scene_number.to_string()
    } else {
        String::new()
    };

    (left_scene_number, right_scene_number)
}

/// 生成文档
/// 生成 DOCX 文档
///
/// 此函数是 DOCX 生成的核心部分，它将解析后的剧本转换为 DOCX 文档。
///
/// # 参数
///
/// * `doc` - DOCX 上下文，包含文档对象和相关状态
/// * `options` - DOCX 选项，包含配置信息
/// * `line_map` - 可选的行映射，用于记录每行的页码、场景和章节信息
///
/// # 返回值
///
/// 返回总页数
pub fn generate(
    doc: &mut DocxContext,
    options: &DocxOptions,
    mut line_map: Option<&mut HashMap<usize, LineStruct>>,
) -> usize {
    // 确保有解析结果
    if options.parsed.is_none() {
        return 0;
    }

    let parsed = options.parsed.as_ref().unwrap();

    // 获取配置
    let cfg = &options.config;
    let print = &options.print_profile;

    // 初始化变量
    let mut china_format = 0; // 是否国内剧本格式
    let line_height = options.line_height;
    let mut print_title_page = cfg.print_title_page;
    let mut print_preface_page = cfg.print_preface_page;
    let mut scenes_numbers = cfg.scenes_numbers.clone();

    // 创建行间距配置
    // 使用adapter中的统一函数

    // 从元数据中获取配置
    if let Some(metadata) = &options.metadata {
        if metadata.contains_key("print") {
            if let Some(china_format_str) = metadata.get("print.chinaFormat") {
                china_format = china_format_str.parse::<i32>().unwrap_or(0);
            }
            if let Some(print_title_page_str) = metadata.get("print.print_title_page") {
                print_title_page = print_title_page_str != "0";
            }
            if let Some(print_preface_page_str) = metadata.get("print.print_preface_page") {
                print_preface_page = print_preface_page_str != "0";
            }
            if let Some(scenes_numbers_str) = metadata.get("print.scenes_numbers") {
                scenes_numbers = scenes_numbers_str.clone();
            }
        }
    }

    // 设置文档属性
    let title_token = parsed.tokens.iter().find(|t| t.token_type == "title");
    let author_token = parsed
        .tokens
        .iter()
        .find(|t| t.token_type == "author")
        .or_else(|| parsed.tokens.iter().find(|t| t.token_type == "authors"));

    if let Some(author) = author_token {
        let author_text = clear_formatting(&author.text);
        let author_text = inline(&author_text);
        doc.doc.custom_property("creator", &author_text);
    }

    if let Some(title) = title_token {
        let title_text = clear_formatting(&title.text);
        let title_text = inline(&title_text);
        doc.doc.custom_property("title", &title_text);
    }

    let section_props = crate::docx::adapter::docx::SectionProperties {
        page: Some(crate::docx::adapter::docx::PageProperties {
            size: Some(crate::docx::adapter::docx::PageSize::new(
                convert_inches_to_twip(print.page_height),
                convert_inches_to_twip(print.page_width),
            )),
            margin: Some(crate::docx::adapter::docx::PageMargin::new(
                convert_inches_to_twip(print.top_margin),
                convert_inches_to_twip(print.right_margin),
                convert_inches_to_twip(print.bottom_margin),
                convert_inches_to_twip(print.left_margin),
                convert_inches_to_twip(print.page_number_top_margin),
                convert_inches_to_twip(if (print.page_number_top_margin - line_height < 0.2) {
                    0.2
                } else {
                    print.page_number_top_margin - line_height
                }),
            )),
            page_numbers: Some(crate::docx::adapter::docx::PageNumbers::new(1)),
        }),
    };

    // 处理标题页
    println!(
        "【generate】开始处理标题页，print_title_page={}, 标题页元素数量={}, 标题页是否已处理={}",
        print_title_page,
        parsed.title_page.len(),
        doc.options.title_page_processed
    );
    println!("【generate】DocxContext 对象 ID: {:p}", doc);

    let mut section_title_page: Option<crate::docx::adapter::docx::Section> = None;

    if print_title_page && !parsed.title_page.is_empty() && !doc.options.title_page_processed {
        // 检查是否有标题页内容
        let has_title_content = parsed
            .title_page
            .iter()
            .any(|(key, tokens)| match key.as_str() {
                "tl" | "tc" | "tr" | "bl" | "cc" | "br" => !tokens.is_empty(),
                _ => false,
            });

        println!("【generate】标题页内容检查结果: {}", has_title_content);

        if has_title_content {
            // 标记标题页已处理
            doc.options.title_page_processed = true;
            println!("【generate】标题页处理标志已设置为 true");

            // 创建标题页 section
            let mut title_section = crate::docx::adapter::docx::Section::new();
            title_section.properties = section_props.clone();

            // 计算页面尺寸（一次性计算，避免重复）
            let dimensions = calculate_page_dimensions(print, convert_point_to_inches(12.0)); //标题页固定单倍行距s所以用240twip（12磅）,参数传入的单位需要的是 英寸

            // 处理标题页内容（按固定顺序：tl | tc | tr | cc | bl | br）
            for key in ["tl", "tc", "tr", "cc", "bl", "br"] {
                if let Some(tokens) = parsed.title_page.get(key) {
                    println!(
                        "【generate】处理标题页元素: {} (包含 {} 个 token)",
                        key,
                        tokens.len()
                    );
                    if !tokens.is_empty() {
                        // 按索引排序并连接文本
                        let mut sorted_tokens = tokens.clone();
                        sorted_tokens.sort_by(|a, b| {
                            if a.index == -1 {
                                std::cmp::Ordering::Equal
                            } else {
                                a.index.cmp(&b.index)
                            }
                        });

                        let mut text = String::new();
                        for token in sorted_tokens {
                            if !text.is_empty() {
                                text.push_str("\n\n");
                            }
                            text.push_str(&token.text);
                        }

                        println!("【generate】标题页元素 {} 文本内容: {}", key, text);

                        // 创建段落
                        if !text.is_empty() {
                            let mut paragraph = crate::docx::adapter::docx::Paragraph::new();

                            // 设置对齐方式
                            if let Some(alignment) = get_title_alignment(key) {
                                paragraph.align(alignment);
                            }

                            // 设置框架属性
                            paragraph.frame(create_title_frame(key, &dimensions));

                            // 处理文本格式化
                            let options_map = create_basic_options_map("#000000");
                            let text_runs = doc.text2(&text, &options_map, None, None);

                            // 添加文本运行到段落
                            for run in text_runs {
                                paragraph.add_text_run(run);
                            }

                            // 添加段落到标题页 section
                            title_section.children.push(
                                crate::docx::adapter::docx::SectionChild::Paragraph(paragraph),
                            );

                            println!("【generate】已添加标题页元素 {} 到标题页 section", key);
                        }
                    }
                }
            }
            section_title_page = Some(title_section);
        }
        println!("【generate】标题页处理完成");
    } else {
        println!(
            "【generate】跳过标题页处理，print_title_page={}, 标题页元素数量={}",
            print_title_page,
            parsed.title_page.len()
        );
    }

    // 创建序言页 section（无页码）
    let mut section_main_no_page_num = crate::docx::adapter::docx::Section::new();
    section_main_no_page_num.properties = section_props.clone();

    // 序言页不设置页眉页脚
    // 根据您的要求：序言页应该不显示页码，也不显示页眉页脚
    // 这样适配器层就不会检测到序言页有页眉页脚，从而不会设置全局页眉页脚
    println!("【generate】序言页不设置页眉页脚，确保序言页不显示页码");

    // 预计算常用的选项映射（避免重复创建）
    let header_footer_options = create_basic_options_map("#777777");

    // 创建主要内容 section（有页码）
    let mut section_main = crate::docx::adapter::docx::Section::new();
    section_main.properties = section_props.clone();

    // 添加页眉（有页码）
    if !cfg.print_header.is_empty() {
        let mut header_paragraph = crate::docx::adapter::docx::Paragraph::new();
        header_paragraph.align(crate::docx::adapter::AlignmentType::Center);

        // 使用 text2 方法格式化页眉文本，支持特殊字符
        let header_runs = doc.format_text(&cfg.print_header, &header_footer_options);
        for run in header_runs {
            header_paragraph.add_text_run(run);
        }

        let header = crate::docx::adapter::docx::Header::new();
        let mut headers = crate::docx::adapter::docx::Headers::new(header);
        headers.default.children.push(header_paragraph);
        section_main.headers = Some(headers.clone());
        section_main_no_page_num.headers = Some(headers);
    }

    // 添加页脚（有页码）
    let mut footer_paragraphs = Vec::new();

    // 添加页脚文本
    if !cfg.print_footer.is_empty() {
        let mut footer_paragraph = crate::docx::adapter::docx::Paragraph::new();
        footer_paragraph.align(crate::docx::adapter::AlignmentType::Center);

        let footer_runs = doc.format_text(&cfg.print_footer, &header_footer_options);
        for run in footer_runs {
            footer_paragraph.add_text_run(run);
        }

        footer_paragraphs.push(footer_paragraph);
    }
    if !footer_paragraphs.is_empty() {
        let footer = crate::docx::adapter::docx::Footer::new();
        let mut footers = crate::docx::adapter::docx::Footers::new(footer);
        footers.default.children = footer_paragraphs.clone();
        section_main_no_page_num.footers = Some(footers);
    }

    // 添加页码
    if !cfg.show_page_numbers.is_empty() {
        let mut page_number_paragraph = crate::docx::adapter::docx::Paragraph::new();
        page_number_paragraph.align(crate::docx::adapter::AlignmentType::Right);

        // 创建页码运行
        let page_runs = doc.create_page_number_runs(&cfg.show_page_numbers);
        for run in page_runs {
            page_number_paragraph.add_run(run);
        }

        footer_paragraphs.push(page_number_paragraph);
    }

    if !footer_paragraphs.is_empty() {
        let footer = crate::docx::adapter::docx::Footer::new();
        let mut footers = crate::docx::adapter::docx::Footers::new(footer);
        footers.default.children = footer_paragraphs;
        section_main.footers = Some(footers);
    }

    // 处理主要内容
    let mut scene_or_section_or_tran_started = false; // 第一个场景头出现之前的内容，不打印页码
    let mut scene_started = false; // 第一个场景头出现之前的内容，不打印三角形

    // Outline 相关变量
    let mut outline_depth = 0; // 当前大纲深度
    let mut current_section_level = 0; // 当前章节层级

    let _bottom_notes = cfg.note_position_bottom;

    // 获取内宽不再需要，因为我们不再使用框架

    // 缩进计算
    let shift_scene_number = if scenes_numbers == "both" || scenes_numbers == "left" {
        convert_inches_to_twip(5.0 * print.font_width)
    } else {
        0
    };

    let scene_indent =
        convert_inches_to_twip(print.scene_heading.feed - print.left_margin) - shift_scene_number;
    let action_indent = convert_inches_to_twip(print.action.feed - print.left_margin);
    let shot_cut_indent = action_indent - convert_inches_to_twip(4.0 * print.font_width); // 镜头交切标志缩进

    // 行间距设置 - 使用合理的固定行距
    // 问题根源：options.line_height 基于 lines_per_page=20 计算，产生过大的行距 (451 twips = 1.9倍)
    // 解决方案：使用基于字体大小的合理倍数，而不是页面布局计算的动态行距
    let line_spacing_twips = convert_inches_to_twip(options.line_height); // 转换为 twips (1pt = 20 twips)

    println!("【行距设置日志】正文样式 - 使用合理的固定行距:");
    println!("  字体大小: {} pt", print.font_size);
    println!("  行距倍数: 1.2 (120%)");
    println!("  转换为 twips: {} twips", line_spacing_twips);
    println!(
        "  相对单倍行距: {:.1} 倍",
        line_spacing_twips as f32 / 240.0
    );
    println!(
        "  (避免使用页面布局的 options.line_height = {} 英寸 = {} twips)",
        options.line_height,
        convert_inches_to_twip(options.line_height)
    );

    let spacing = crate::docx::adapter::docx::ParagraphSpacing {
        line: Some(line_spacing_twips), // 276 twips (1.15倍单倍行距)
        line_rule: Some(crate::docx::adapter::LineRuleType::Exact), // 使用精确行距
        before: None,
        after: None,
        before_lines: None,
        after_lines: None,
    };

    // 添加样式定义
    doc.add_document_styles(&spacing);

    // 预计算常用的选项映射（避免重复创建）
    let default_text_options = create_basic_options_map("#000000");

    // 预计算双对话缩进值（避免重复计算）
    let dial_indent_out = convert_inches_to_twip(3.0 * print.font_width);
    let dial_indent_in = dial_indent_out / 2;

    // 预计算各种对话类型的缩进值（避免重复计算）
    let character_indent = convert_inches_to_twip(print.character.feed - print.left_margin);
    let dialogue_indent = convert_inches_to_twip(print.dialogue.feed - print.left_margin);
    let parenthetical_indent = convert_inches_to_twip(print.parenthetical.feed - print.left_margin);

    // 预计算双对话的各种缩进组合
    let character_indent_out = dial_indent_out * 3;
    let character_indent_in = character_indent_out - dial_indent_in;
    let parenthetical_indent_out = dial_indent_out * 2;
    let parenthetical_indent_in = parenthetical_indent_out - dial_indent_in;

    // 初始化脚注页面数据结构 - 参考原项目 docxmaker.ts 中的 notesPage
    let mut notes_page: Vec<Vec<Vec<Note>>> = Vec::new();
    let mut current_line_notes: Vec<Note> = Vec::new(); // 当前行的脚注列表
    let bottom_notes = cfg.note_position_bottom; // 是否将脚注放在页面底部

    println!("【generate】脚注配置: bottom_notes = {}", bottom_notes);

    // 处理每一行
    let mut current_page = 0;
    let mut current_scene = String::new();
    let mut current_sections: Vec<String> = Vec::new();
    let current_duration = 0.0;
    let mut curr_type = String::new();
    let mut page_started = false;
    let mut after_section = false; // 跟踪是否在 section 之后

    // 判断是否应该跳过空行
    fn should_del_blank_line(
        lines: &[Line],
        idx: usize,
        rm_blank_line: i32,
        curr_type: &mut String,
    ) -> bool {
        if rm_blank_line == 0 {
            return false;
        }

        let curr = &lines[idx];
        let pre_type = curr_type.clone();

        if curr.token_type == "page_break"
            || curr.token_type == "page_switch"
            || curr.token_type == "redraw"
        {
            return false;
        }

        *curr_type = curr.token_type.clone();

        if !is_blank_line_after_style(&curr.text) {
            return false;
        }

        if idx + 1 < lines.len() {
            let next = &lines[idx + 1];
            let next_type = next.token_type.clone();

            if is_blank_line_after_style(&next.text)
                && next_type != "page_break"
                && next_type != "page_switch"
                && next_type != "redraw"
            {
                *curr_type = pre_type;
                return true;
            }

            if next_type == "scene_heading" {
                return false;
            }

            if rm_blank_line == 2 {
                if next_type == "character" {
                    return false;
                }

                if next_type != "parenthetical"
                    && next_type != "dialogue"
                    && (pre_type == "character"
                        || pre_type == "parenthetical"
                        || pre_type == "dialogue")
                {
                    return false;
                }
            }
        }

        *curr_type = pre_type;
        return true;
    }

    // 如果有处理过的行，则使用处理过的行
    if !parsed.lines.is_empty() {
        for (ii, line) in parsed.lines.iter().enumerate() {
            // 检查是否需要跳过空行
            if should_del_blank_line(&parsed.lines, ii, doc.rm_blank_line, &mut curr_type) {
                // 只绘制样式，再跳过
                doc.text2(
                    &line.text,
                    &default_text_options,
                    if bottom_notes {
                        Some(&mut current_line_notes)
                    } else {
                        None
                    },
                    Some(&mut notes_page),
                );
                continue;
            }

            // 去除页面前面的空行
            if !page_started {
                if line.token_type == "page_break" {
                    // 跳过空行
                    continue;
                }

                if line.token_type != "redraw" {
                    if line.text.trim().is_empty() {
                        // 跳过空行
                        continue;
                    }

                    if is_blank_line_after_style(&line.text) {
                        // 只含有样式字符
                        // 只绘制样式，再跳过
                        doc.text2(
                            &line.text,
                            &default_text_options,
                            if bottom_notes {
                                Some(&mut current_line_notes)
                            } else {
                                None
                            },
                            Some(&mut notes_page),
                        );
                        continue;
                    }
                }

                // 当页第一个非空行
                page_started = true;
            }

            // 根据行类型处理
            let token_type = line.token_type.as_str();
            if token_type == "scene_heading" {
                // 对话块结束，额外处理 (双对话 / 国内剧本对话) - 参考原项目逻辑
                finish_dialogue_processing(
                    doc,
                    china_format,
                    token_type,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &print,
                    &spacing,
                );

                if !scene_or_section_or_tran_started {
                    scene_or_section_or_tran_started = true;
                }
                if !scene_started {
                    scene_started = true;
                }

                // 更新当前场景
                current_scene = line.text.clone();

                // 创建段落
                let mut paragraph =
                    crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
                paragraph.style("scene");

                // 设置缩进
                paragraph.indent(scene_indent);

                // 设置 outline level（场景头使用层级）
                if cfg.create_bookmarks {
                    let do_outline = if cfg.print_sections { 8 } else { 0 }; // 固定最高级或最低级
                    paragraph.outline_level(do_outline);
                }

                // 处理场景编号
                let mut text = line.text.clone();
                if let Some(number) = &line.number {
                    if scenes_numbers == "both" || scenes_numbers == "left" {
                        let scene_number = number.to_string();
                        let scene_text_length = scene_number.chars().count(); // 使用字符数而不是字节数

                        let left_char = if scene_text_length < 3 {
                            3 - scene_text_length
                        } else {
                            0
                        };

                        let left_spaces = if left_char > 0 {
                            " ".repeat(left_char)
                        } else {
                            "".to_string()
                        };

                        text = format!("{}{}  {}", left_spaces, scene_number, text);
                    }
                }

                // 处理样式
                use crate::utils::fountain_constants::FountainConstants;
                let style_chars = FountainConstants::style_chars();

                if cfg.embolden_scene_headers {
                    // 使用特殊字符 ↭ 表示粗体
                    text = add_tag_after_broken_note(text, style_chars["bold"].to_string())
                        + &style_chars["bold"].to_string();
                }
                if cfg.underline_scene_headers {
                    // 使用特殊字符 ☄ 表示下划线
                    text = add_tag_after_broken_note(text, style_chars["underline"].to_string())
                        + &style_chars["underline"].to_string();
                }

                text = if_reset_format(text, line);

                // 创建文本运行
                let mut scene_options = default_text_options.clone();
                scene_options.insert("characterSpacing".to_string(), "0".to_string());

                let text_runs = doc.text2(
                    &text,
                    &scene_options,
                    if bottom_notes {
                        Some(&mut current_line_notes)
                    } else {
                        None
                    },
                    Some(&mut notes_page),
                );

                // 添加文本运行到段落
                for run in text_runs {
                    paragraph.add_text_run(run);
                }

                // 添加段落到相应section并更新行映射
                add_paragraph_and_update_line_map(
                    crate::docx::adapter::docx::SectionChild::Paragraph(paragraph),
                    line,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &mut line_map,
                    &current_sections,
                    &current_scene,
                    current_page,
                    current_duration,
                );
            } else if token_type == "action" {
                // 非对话元素：完成缓存的中文格式对话和双对话
                finish_dialogue_processing(
                    doc,
                    china_format,
                    token_type,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &print,
                    &spacing,
                );

                // 创建段落
                let mut paragraph =
                    crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
                paragraph.style("action");

                // 设置缩进
                paragraph.indent(action_indent);

                // 处理文本
                let mut text = line.text.clone();
                text = if_reset_format(text, line);

                // 添加三角形（国内剧本格式）
                if (china_format == 1 || china_format == 3) && scene_started {
                    text = format!("△ {}", text);
                }

                // 创建文本运行
                let text_runs = doc.text2(
                    &text,
                    &default_text_options,
                    if bottom_notes {
                        Some(&mut current_line_notes)
                    } else {
                        None
                    },
                    Some(&mut notes_page),
                );

                // 添加文本运行到段落
                for run in text_runs {
                    paragraph.add_text_run(run);
                }

                // 添加段落到相应section并更新行映射
                add_paragraph_and_update_line_map(
                    crate::docx::adapter::docx::SectionChild::Paragraph(paragraph),
                    line,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &mut line_map,
                    &current_sections,
                    &current_scene,
                    current_page,
                    current_duration,
                );
            } else if token_type == "centered" {
                // 非对话元素：完成缓存的中文格式对话和双对话
                finish_dialogue_processing(
                    doc,
                    china_format,
                    token_type,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &print,
                    &spacing,
                );

                // 创建段落
                let mut paragraph =
                    crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
                paragraph.style("action");

                // 设置缩进
                paragraph.indent(action_indent);
                paragraph.align(crate::docx::adapter::AlignmentType::Center);

                // 处理文本
                let mut text = line.text.clone();
                text = if_reset_format(text, line);

                // 创建文本运行
                let text_runs = doc.text2(
                    &text,
                    &default_text_options,
                    if bottom_notes {
                        Some(&mut current_line_notes)
                    } else {
                        None
                    },
                    Some(&mut notes_page),
                );

                // 添加文本运行到段落
                for run in text_runs {
                    paragraph.add_text_run(run);
                }

                // 添加段落到相应section并更新行映射
                add_paragraph_and_update_line_map(
                    crate::docx::adapter::docx::SectionChild::Paragraph(paragraph),
                    line,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &mut line_map,
                    &current_sections,
                    &current_scene,
                    current_page,
                    current_duration,
                );
            } else if token_type == "synopsis" {
                // 对话块结束，额外处理 (双对话 / 国内剧本对话) - 参考原项目逻辑
                finish_dialogue_processing(
                    doc,
                    china_format,
                    token_type,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &print,
                    &spacing,
                );

                // 处理 synopsis - 参考原项目逻辑
                let mut feed = print.synopsis.feed.unwrap_or(print.action.feed);

                // 检查是否需要根据最后一个 section 调整缩进
                if print.synopsis.feed_with_last_section && after_section {
                    feed += current_section_level as f32 * print.section.level_indent;
                } else {
                    feed = print.action.feed;
                }

                // 添加 synopsis 的 padding
                feed += print.synopsis.padding.unwrap_or(0.0);

                // 计算缩进（相对于 action 的缩进）
                let section_indent = convert_inches_to_twip(feed - print.action.feed);

                // 创建段落
                let mut paragraph =
                    crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
                paragraph.style("action"); // synopsis 使用 action 样式

                // 设置左右缩进
                paragraph.indent(section_indent);
                paragraph.indent_right(section_indent);

                let mut synopsis_options = if let Some(color) = print.synopsis.color.as_ref() {
                    create_basic_options_map(color)
                } else {
                    default_text_options.clone()
                };

                if print.synopsis.italic {
                    synopsis_options.insert("italic".to_string(), "true".to_string());
                }

                // 处理文本
                let mut text = line.text.clone();
                text = if_reset_format(text, line);

                // 创建文本运行
                let text_runs = doc.text2(
                    &text,
                    &synopsis_options,
                    if bottom_notes {
                        Some(&mut current_line_notes)
                    } else {
                        None
                    },
                    Some(&mut notes_page),
                );

                // 添加文本运行到段落
                for run in text_runs {
                    paragraph.add_text_run(run);
                }

                // 添加段落到相应section并更新行映射
                add_paragraph_and_update_line_map(
                    crate::docx::adapter::docx::SectionChild::Paragraph(paragraph),
                    line,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &mut line_map,
                    &current_sections,
                    &current_scene,
                    current_page,
                    current_duration,
                );
            } else if token_type == "dialogue"
                || token_type == "character"
                || token_type == "parenthetical"
            {
                println!(
                    "【generate】处理对话类型: {}, dual: {:?}, text: {}",
                    token_type, line.dual, line.text
                );

                // 检查是否需要处理双对话结束
                if line.token_type == "character" {
                    // 如果当前角色不是右侧对话，完成之前的双对话
                    if line.dual.as_deref() != Some("right") {
                        doc.finish_double_dial(
                            if scene_or_section_or_tran_started {
                                &mut section_main
                            } else {
                                &mut section_main_no_page_num
                            },
                            &print,
                            &spacing,
                        );
                    }
                }

                // 处理文本
                let mut text = line.text.clone();

                // 处理角色名
                if line.token_type == "character" {
                    // 添加粗体标记（如果启用）
                    if cfg.embolden_character_names {
                        // 使用样式字符常量
                        use crate::utils::fountain_constants::FountainConstants;
                        let style_chars = FountainConstants::style_chars();

                        // 检查是否以 text_contd 结尾
                        if text.ends_with(&cfg.text_contd) {
                            let base_text = &text[..text.len() - cfg.text_contd.len()];
                            text = add_tag_after_broken_note(
                                base_text.to_string(),
                                style_chars["bold"].to_string(),
                            ) + &style_chars["bold"].to_string()
                                + cfg.text_contd.as_str();
                        } else {
                            text = add_tag_after_broken_note(text, style_chars["bold"].to_string())
                                + &style_chars["bold"].to_string();
                        }
                    }

                    // 中文格式处理
                    if china_format > 0 {
                        // 先完成之前的缓存对话
                        doc.finish_china_dial_first(
                            if scene_or_section_or_tran_started {
                                &mut section_main
                            } else {
                                &mut section_main_no_page_num
                            },
                            &spacing,
                        );

                        // 添加冒号
                        text = format!("{}: ", text);
                    }
                }

                text = if_reset_format(text, line);

                // 创建文本运行
                let text_runs = doc.text2(
                    &text,
                    &default_text_options,
                    if bottom_notes {
                        Some(&mut current_line_notes)
                    } else {
                        None
                    },
                    Some(&mut notes_page),
                );

                // 根据 china_format、dual 属性和类型处理
                if china_format > 0 {
                    // 中文格式：使用缓存拼接机制
                    println!("【generate】中文格式处理，china_format = {}", china_format);
                    if line.token_type == "character" {
                        // 根据 dual 属性决定缓存位置
                        if line.dual.as_deref() == Some("left") {
                            println!("【generate】缓存左侧双对话角色: {}", text);
                            // 左侧对话
                            doc.last_dial_gr_left = Some(CachedDialogueGroup {
                                style: "dial".to_string(),
                                children: text_runs,
                                indent_left: Some(dial_indent_out),
                                indent_right: Some(dial_indent_in),
                            });
                        } else if line.dual.as_deref() == Some("right") {
                            println!("【generate】缓存右侧双对话角色: {}", text);
                            // 右侧对话
                            doc.last_dial_gr_right = Some(CachedDialogueGroup {
                                style: "dial".to_string(),
                                children: text_runs,
                                indent_left: Some(dial_indent_in),
                                indent_right: Some(dial_indent_out),
                            });
                        } else {
                            // 普通单对话：缓存到 last_dial_gr
                            doc.last_dial_gr = Some(CachedDialogueGroup {
                                style: "action".to_string(),
                                children: text_runs,
                                indent_left: None,
                                indent_right: None,
                            });
                        }
                    } else if line.token_type == "parenthetical" {
                        // 括号内容：根据 dual 属性追加到相应缓存
                        if line.dual.as_deref() == Some("left") {
                            // 左侧对话的括号内容
                            if let Some(ref mut dial_gr_left) = doc.last_dial_gr_left {
                                dial_gr_left.children.extend(text_runs);
                            } else {
                                // 如果没有缓存的左侧角色名，直接添加到表格 - 参考原项目逻辑
                                let mut paragraph =
                                    crate::docx::adapter::docx::Paragraph::new_with_spacing(
                                        spacing.clone(),
                                    );
                                paragraph.style("dial");
                                paragraph.indent(dial_indent_out);
                                paragraph.indent_right(dial_indent_in);

                                for run in text_runs {
                                    paragraph.add_text_run(run);
                                }

                                // 添加到全局表格缓存 - 修复关键问题
                                println!(
                                    "【generate】中文格式左侧括号内容（无缓存）添加到全局表格缓存"
                                );
                                doc.last_dial_table_left.push(paragraph);
                            }
                        } else if line.dual.as_deref() == Some("right") {
                            // 右侧对话的括号内容
                            if let Some(ref mut dial_gr_right) = doc.last_dial_gr_right {
                                dial_gr_right.children.extend(text_runs);
                            } else {
                                // 如果没有缓存的右侧角色名，直接添加到表格 - 参考原项目逻辑
                                let mut paragraph =
                                    crate::docx::adapter::docx::Paragraph::new_with_spacing(
                                        spacing.clone(),
                                    );
                                paragraph.style("dial");
                                paragraph.indent(dial_indent_in);
                                paragraph.indent_right(dial_indent_out);

                                for run in text_runs {
                                    paragraph.add_text_run(run);
                                }

                                // 添加到全局表格缓存 - 修复关键问题
                                println!(
                                    "【generate】中文格式右侧括号内容（无缓存）添加到全局表格缓存"
                                );
                                doc.last_dial_table_right.push(paragraph);
                            }
                        } else {
                            // 普通单对话的括号内容
                            if let Some(ref mut dial_gr) = doc.last_dial_gr {
                                dial_gr.children.extend(text_runs);
                            } else {
                                // 如果没有缓存的角色名，直接输出
                                let mut paragraph =
                                    crate::docx::adapter::docx::Paragraph::new_with_spacing(
                                        spacing.clone(),
                                    );
                                paragraph.style("action");

                                for run in text_runs {
                                    paragraph.add_text_run(run);
                                }

                                if scene_or_section_or_tran_started {
                                    section_main.children.push(
                                        crate::docx::adapter::docx::SectionChild::Paragraph(
                                            paragraph,
                                        ),
                                    );
                                } else {
                                    section_main_no_page_num.children.push(
                                        crate::docx::adapter::docx::SectionChild::Paragraph(
                                            paragraph,
                                        ),
                                    );
                                }
                            }
                        }
                    } else if line.token_type == "dialogue" {
                        // 对白：根据 dual 属性追加到相应缓存
                        if line.dual.as_deref() == Some("left") {
                            // 左侧对话
                            if let Some(ref mut dial_gr_left) = doc.last_dial_gr_left {
                                dial_gr_left.children.extend(text_runs);

                                // 根据 china_format 值决定是否立即输出到左侧表格
                                if china_format == 1 || china_format == 2 {
                                    // 立即输出并清空缓存 - 参考原项目逻辑
                                    let cached_group = doc.last_dial_gr_left.clone().unwrap();
                                    let mut paragraph =
                                        crate::docx::adapter::docx::Paragraph::new_with_spacing(
                                            spacing.clone(),
                                        );
                                    paragraph.style(&cached_group.style);

                                    if let Some(left) = cached_group.indent_left {
                                        paragraph.indent(left);
                                    }
                                    if let Some(right) = cached_group.indent_right {
                                        paragraph.indent_right(right);
                                    }

                                    for run in cached_group.children {
                                        paragraph.add_text_run(run);
                                    }

                                    // 添加到全局表格缓存
                                    println!("【generate】中文格式左侧对话添加到全局表格缓存");
                                    doc.last_dial_table_left.push(paragraph);

                                    // 清空缓存，准备下一个对话行 - 关键修复
                                    doc.last_dial_gr_left = None;
                                }
                                // china_format == 3 || china_format == 4 时继续缓存，等待更多对白
                            } else {
                                // 如果没有缓存的左侧角色名，直接输出到全局表格缓存
                                let mut paragraph =
                                    crate::docx::adapter::docx::Paragraph::new_with_spacing(
                                        spacing.clone(),
                                    );
                                paragraph.style("dial");
                                paragraph.indent(dial_indent_out);
                                paragraph.indent_right(dial_indent_in);

                                for run in text_runs {
                                    paragraph.add_text_run(run);
                                }

                                // 添加到全局表格缓存 - 修复关键问题
                                println!(
                                    "【generate】中文格式左侧对话（无缓存）添加到全局表格缓存"
                                );
                                doc.last_dial_table_left.push(paragraph);
                            }
                        } else if line.dual.as_deref() == Some("right") {
                            // 右侧对话
                            if let Some(ref mut dial_gr_right) = doc.last_dial_gr_right {
                                dial_gr_right.children.extend(text_runs);

                                // 根据 china_format 值决定是否立即输出到右侧表格
                                if china_format == 1 || china_format == 2 {
                                    // 立即输出并清空缓存 - 参考原项目逻辑
                                    let cached_group = doc.last_dial_gr_right.clone().unwrap();
                                    let mut paragraph =
                                        crate::docx::adapter::docx::Paragraph::new_with_spacing(
                                            spacing.clone(),
                                        );
                                    paragraph.style(&cached_group.style);

                                    if let Some(left) = cached_group.indent_left {
                                        paragraph.indent(left);
                                    }
                                    if let Some(right) = cached_group.indent_right {
                                        paragraph.indent_right(right);
                                    }

                                    for run in cached_group.children {
                                        paragraph.add_text_run(run);
                                    }

                                    // 添加到全局表格缓存
                                    println!("【generate】中文格式右侧对话添加到全局表格缓存");
                                    doc.last_dial_table_right.push(paragraph);

                                    // 清空缓存，准备下一个对话行 - 关键修复
                                    doc.last_dial_gr_right = None;
                                }
                                // china_format == 3 || china_format == 4 时继续缓存，等待更多对白
                            } else {
                                // 如果没有缓存的右侧角色名，直接输出到全局表格缓存
                                let mut paragraph =
                                    crate::docx::adapter::docx::Paragraph::new_with_spacing(
                                        spacing.clone(),
                                    );
                                paragraph.style("dial");
                                paragraph.indent(dial_indent_in);
                                paragraph.indent_right(dial_indent_out);

                                for run in text_runs {
                                    paragraph.add_text_run(run);
                                }

                                // 添加到全局表格缓存 - 修复关键问题
                                println!(
                                    "【generate】中文格式右侧对话（无缓存）添加到全局表格缓存"
                                );
                                doc.last_dial_table_right.push(paragraph);
                            }
                        } else {
                            // 普通单对话
                            if let Some(ref mut dial_gr) = doc.last_dial_gr {
                                dial_gr.children.extend(text_runs);

                                // 根据 china_format 值决定是否立即输出
                                if china_format == 1 || china_format == 2 {
                                    // 立即输出并清空缓存
                                    let cached_group = doc.last_dial_gr.take().unwrap();
                                    let mut paragraph =
                                        crate::docx::adapter::docx::Paragraph::new_with_spacing(
                                            spacing.clone(),
                                        );
                                    paragraph.style(&cached_group.style);

                                    for run in cached_group.children {
                                        paragraph.add_text_run(run);
                                    }

                                    if scene_or_section_or_tran_started {
                                        section_main.children.push(
                                            crate::docx::adapter::docx::SectionChild::Paragraph(
                                                paragraph,
                                            ),
                                        );
                                    } else {
                                        section_main_no_page_num.children.push(
                                            crate::docx::adapter::docx::SectionChild::Paragraph(
                                                paragraph,
                                            ),
                                        );
                                    }
                                }
                                // china_format == 3 时继续缓存，等待更多对白
                            } else {
                                // 如果没有缓存的角色名，直接输出
                                // 在 china_format 下，没有缓存的对白使用 "action" 样式，不设置特殊缩进
                                let mut paragraph =
                                    crate::docx::adapter::docx::Paragraph::new_with_spacing(
                                        spacing.clone(),
                                    );
                                paragraph.style("action");

                                for run in text_runs {
                                    paragraph.add_text_run(run);
                                }

                                if scene_or_section_or_tran_started {
                                    section_main.children.push(
                                        crate::docx::adapter::docx::SectionChild::Paragraph(
                                            paragraph,
                                        ),
                                    );
                                } else {
                                    section_main_no_page_num.children.push(
                                        crate::docx::adapter::docx::SectionChild::Paragraph(
                                            paragraph,
                                        ),
                                    );
                                }
                            }
                        }
                    }
                } else {
                    // 国际格式：根据 dual 属性处理
                    if line.dual.as_deref() == Some("left") || line.dual.as_deref() == Some("right")
                    {
                        // 双对话：添加到全局表格缓存
                        println!("【generate】国际格式双对话: {} - {}", line.token_type, text);

                        let mut paragraph = crate::docx::adapter::docx::Paragraph::new_with_spacing(
                            spacing.clone(),
                        );

                        if line.token_type == "dialogue" {
                            paragraph.style("dial");
                            if line.dual.as_deref() == Some("left") {
                                paragraph.indent(dial_indent_out);
                                paragraph.indent_right(dial_indent_in);
                            } else {
                                paragraph.indent(dial_indent_in);
                                paragraph.indent_right(dial_indent_out);
                            }
                        } else if line.token_type == "character" {
                            paragraph.style("character");
                            if line.dual.as_deref() == Some("left") {
                                paragraph.indent(character_indent_out);
                                paragraph.indent_right(character_indent_in);
                            } else {
                                paragraph.indent(character_indent_in);
                                paragraph.indent_right(character_indent_out);
                            }
                        } else if line.token_type == "parenthetical" {
                            paragraph.style("parenthetical");
                            if line.dual.as_deref() == Some("left") {
                                paragraph.indent(parenthetical_indent_out);
                                paragraph.indent_right(parenthetical_indent_in);
                            } else {
                                paragraph.indent(parenthetical_indent_in);
                                paragraph.indent_right(parenthetical_indent_out);
                            }
                        }

                        for run in text_runs {
                            paragraph.add_text_run(run);
                        }

                        // 添加到全局表格缓存 - 修复关键问题
                        if line.dual.as_deref() == Some("left") {
                            println!("【generate】添加到左侧全局表格缓存");
                            doc.last_dial_table_left.push(paragraph);
                        } else {
                            println!("【generate】添加到右侧全局表格缓存");
                            doc.last_dial_table_right.push(paragraph);
                        }
                    } else {
                        // 普通单对话：直接输出
                        let mut paragraph = crate::docx::adapter::docx::Paragraph::new_with_spacing(
                            spacing.clone(),
                        );

                        if line.token_type == "dialogue" {
                            paragraph.style("dial");
                            paragraph.indent(dialogue_indent);
                        } else if line.token_type == "character" {
                            paragraph.style("character");
                            paragraph.indent(character_indent);
                        } else if line.token_type == "parenthetical" {
                            paragraph.style("parenthetical");
                            paragraph.indent(parenthetical_indent);
                        }

                        for run in text_runs {
                            paragraph.add_text_run(run);
                        }

                        if scene_or_section_or_tran_started {
                            section_main.children.push(
                                crate::docx::adapter::docx::SectionChild::Paragraph(paragraph),
                            );
                        } else {
                            section_main_no_page_num.children.push(
                                crate::docx::adapter::docx::SectionChild::Paragraph(paragraph),
                            );
                        }
                    }
                }

                // 更新行映射（对话类型不需要添加段落，因为已经在对话处理逻辑中处理）
                if let Some(token_line) = line.token {
                    // 在 Line 结构体中没有 time 字段，暂时忽略时间计算
                    // if let Some(time) = line.time {
                    //     current_duration += time as f32;
                    // }
                    if let Some(ref mut lm) = line_map {
                        lm.insert(
                            token_line,
                            LineStruct {
                                sections: current_sections.clone(),
                                scene: current_scene.clone(),
                                page: current_page,
                                cumulative_duration: current_duration as f32,
                            },
                        );
                    }
                }
            } else if token_type == "section" {
                // 对话块结束，额外处理 (双对话 / 国内剧本对话) - 参考原项目逻辑
                finish_dialogue_processing(
                    doc,
                    china_format,
                    token_type,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &print,
                    &spacing,
                );

                if !scene_or_section_or_tran_started {
                    scene_or_section_or_tran_started = true;
                }

                // 处理 section 层级信息 - 参考原项目的 processSection 逻辑
                if let Some(level) = line.level {
                    // 更新当前章节层级
                    current_section_level = level as usize;

                    // 截断 current_sections 到当前层级-1
                    current_sections.truncate(if level > 0 { level as usize - 1 } else { 0 });

                    // 添加当前章节文本到 current_sections
                    current_sections.push(line.text.clone());

                    // 更新 outline_depth
                    if cfg.create_bookmarks {
                        outline_depth = level as usize;
                    }
                }

                // 创建段落
                let mut paragraph =
                    crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
                paragraph.style("section");

                // 计算章节缩进 - 参考原项目使用 current_section_level
                let feed =
                    print.section.feed + current_section_level as f32 * print.section.level_indent;
                let section_indent = convert_inches_to_twip(feed - print.left_margin);

                // 设置缩进
                paragraph.indent(section_indent);

                // 设置 outline level（章节使用当前章节层级）
                if cfg.create_bookmarks {
                    paragraph.outline_level(outline_depth);
                }

                // 处理文本
                let mut text = line.text.clone();
                text = if_reset_format(text, line);

                // 创建文本运行
                let section_options = if let Some(color) = print.section.color.as_ref() {
                    create_basic_options_map(color)
                } else {
                    default_text_options.clone()
                };

                let text_runs = doc.text2(
                    &text,
                    &section_options,
                    if bottom_notes {
                        Some(&mut current_line_notes)
                    } else {
                        None
                    },
                    Some(&mut notes_page),
                );

                // 添加文本运行到段落
                for run in text_runs {
                    paragraph.add_text_run(run);
                }

                // 添加段落到相应section并更新行映射
                add_paragraph_and_update_line_map(
                    crate::docx::adapter::docx::SectionChild::Paragraph(paragraph),
                    line,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &mut line_map,
                    &current_sections,
                    &current_scene,
                    current_page,
                    current_duration,
                );
            } else if token_type == "transition" {
                // 对话块结束，额外处理 (双对话 / 国内剧本对话) - 参考原项目逻辑
                finish_dialogue_processing(
                    doc,
                    china_format,
                    token_type,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &print,
                    &spacing,
                );

                if !scene_or_section_or_tran_started {
                    scene_or_section_or_tran_started = true;
                }

                // 处理文本
                let mut text = line.text.clone();
                let mut is_shot_cut = false;
                let mut style = "action";
                let mut indent = action_indent;

                // 检查是否是镜头交切特殊字符
                if (text.starts_with("{=") && text.ends_with("=} ↓"))
                    || (text.starts_with("{#") && text.ends_with("#} ↓"))
                    || (text.starts_with("{+") && text.ends_with("+} ↓"))
                    || (text.starts_with("{-") && text.ends_with("-} ↑"))
                {
                    // 镜头交切标志
                    is_shot_cut = true;
                    style = "shotCut";
                    indent = shot_cut_indent;

                    // 转换文本：去掉 {+ 和 +}，保留中间内容和箭头，用括号包围
                    // 例如：{+镜头交切+} ↓ -> (镜头交切 ↓)
                    // 使用字符索引而不是字节索引来处理 Unicode 字符
                    let chars: Vec<char> = text.chars().collect();
                    if chars.len() >= 6 {
                        // 至少需要 {+x+} ↓ 这样的格式
                        let inner_text: String = chars[2..chars.len() - 4].iter().collect(); // 去掉 {+ 和 +}
                        let arrow: String = chars[chars.len() - 2..].iter().collect(); // 获取箭头部分
                        text = format!("({}{})", inner_text, arrow);
                    }
                }

                // 创建段落
                let mut paragraph =
                    crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
                paragraph.style(style);

                // 设置缩进
                paragraph.indent(indent);

                // 设置行间距 - 使用行高减去字体大小的算法
                let final_line_spacing_twips = convert_inches_to_twip(options.line_height);

                // 处理转场对齐
                if is_shot_cut {
                    // 镜头交切：左对齐
                    paragraph.align(crate::docx::adapter::AlignmentType::Left);
                } else {
                    // 普通转场
                    if china_format > 0 {
                        text = format!("({})", text);
                        paragraph.align(crate::docx::adapter::AlignmentType::Left);
                    } else {
                        paragraph.align(crate::docx::adapter::AlignmentType::Right);
                    }
                }

                // 创建文本运行
                let mut transition_options = default_text_options.clone();

                // 如果是镜头交切，设置粗体
                if is_shot_cut {
                    transition_options.insert("bold".to_string(), "true".to_string());
                }

                text = if_reset_format(text, line);

                let text_runs = doc.text2(
                    &text,
                    &transition_options,
                    if bottom_notes {
                        Some(&mut current_line_notes)
                    } else {
                        None
                    },
                    Some(&mut notes_page),
                );

                // 添加文本运行到段落
                for run in text_runs {
                    paragraph.add_text_run(run);
                }

                // 添加段落到相应section并更新行映射
                add_paragraph_and_update_line_map(
                    crate::docx::adapter::docx::SectionChild::Paragraph(paragraph),
                    line,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &mut line_map,
                    &current_sections,
                    &current_scene,
                    current_page,
                    current_duration,
                );
            } else if token_type == "page_break" {
                // 对话块结束，额外处理 (双对话 / 国内剧本对话) - 参考原项目逻辑
                finish_dialogue_processing(
                    doc,
                    china_format,
                    token_type,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &print,
                    &spacing,
                );

                // 更新页码
                current_page += 1;

                // 添加段落到相应section并更新行映射
                add_paragraph_and_update_line_map(
                    crate::docx::adapter::docx::SectionChild::PageBreak,
                    line,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &mut line_map,
                    &current_sections,
                    &current_scene,
                    current_page,
                    current_duration,
                );
            } else {
                // 对话块结束，额外处理 (双对话 / 国内剧本对话) - 参考原项目逻辑
                finish_dialogue_processing(
                    doc,
                    china_format,
                    token_type,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &print,
                    &spacing,
                );

                // 处理其他类型的token
                let mut paragraph =
                    crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone());
                paragraph.style("action");

                // 设置缩进
                paragraph.indent(action_indent);

                // 创建文本运行
                let text_runs = doc.text2(
                    &line.text,
                    &default_text_options,
                    if bottom_notes {
                        Some(&mut current_line_notes)
                    } else {
                        None
                    },
                    Some(&mut notes_page),
                );

                // 添加文本运行到段落
                for run in text_runs {
                    paragraph.add_text_run(run);
                }

                // 添加段落到相应section并更新行映射
                add_paragraph_and_update_line_map(
                    crate::docx::adapter::docx::SectionChild::Paragraph(paragraph),
                    line,
                    scene_or_section_or_tran_started,
                    &mut section_main,
                    &mut section_main_no_page_num,
                    &mut line_map,
                    &current_sections,
                    &current_scene,
                    current_page,
                    current_duration,
                );
            }

            // 更新 after_section 状态 - 参考原项目逻辑
            if page_started {
                if token_type == "section" {
                    after_section = true;
                } else if token_type == "scene_heading" {
                    after_section = false;
                }
            }

            // 每行处理结束后清空当前行的脚注列表 - 参考原项目逻辑
            current_line_notes.clear();
        }
    }

    // 完成所有缓存的中文格式对话和双对话 - 参考原项目逻辑
    finish_dialogue_processing(
        doc,
        china_format,
        "document_end", // 特殊标记表示文档结束
        scene_or_section_or_tran_started,
        &mut section_main,
        &mut section_main_no_page_num,
        &print,
        &spacing,
    );

    // 处理脚注 - 参考原项目 docxmaker.ts 中的脚注处理逻辑
    // 只有当 bottom_notes = true 时才处理页面底部的脚注
    if bottom_notes && !notes_page.is_empty() && !notes_page[0].is_empty() {
        println!(
            "【generate】开始处理页面底部脚注，脚注数量: {}",
            notes_page[0].len()
        );

        let notes = &notes_page[0];
        for (i, token_row) in notes.iter().enumerate() {
            println!("【generate】处理脚注行 #{}: {} 个脚注", i, token_row.len());

            for (j, note) in token_row.iter().enumerate() {
                println!(
                    "【generate】处理脚注 #{}: 编号={}, 文本行数={}, 文本内容: {:?}",
                    j,
                    note.no,
                    note.text.len(),
                    note.text
                );

                let mut children = Vec::new();

                for (k, text_line) in note.text.iter().enumerate() {
                    let mut text = text_line.clone();

                    // 去掉第一个字符（脚注开始标记）
                    if k == 0 && !text.is_empty() {
                        if let Some(first_char) = text.chars().next() {
                            text = text[first_char.len_utf8()..].to_string();
                        }
                    }

                    // 去掉最后一个字符（脚注结束标记）
                    if k == note.text.len() - 1 && !text.is_empty() {
                        if let Some(last_char) = text.chars().last() {
                            let last_char_start = text.len() - last_char.len_utf8();
                            text = text[..last_char_start].to_string();
                        }
                    }

                    // 创建脚注段落
                    let mut paragraph = crate::docx::adapter::docx::Paragraph::new(); //底部脚注内容使用单倍行距
                                                                                      // let mut paragraph =
                                                                                      //     crate::docx::adapter::docx::Paragraph::new_with_spacing(spacing.clone().line(convert_inches_to_twip(options.print_profile.note_line_height)).line_rule(crate::docx::adapter::LineRuleType::AtLeast));
                    paragraph.style("notes");
                    if k == 0 {
                        paragraph.indent_first_line(convert_inches_to_twip(2.0 * print.font_width));
                    }

                    // 创建文本运行 - 参考原项目使用固定颜色 #868686
                    let mut footnote_options = create_basic_options_map("#868686");
                    footnote_options
                        .insert("fontSize".to_string(), print.note_font_size.to_string());
                    footnote_options.insert("characterSpacing".to_string(), "0".to_string());

                    let text_runs = doc.format_text(&text, &footnote_options);

                    // 添加文本运行到段落
                    for run in text_runs {
                        paragraph.add_text_run(run);
                    }

                    children.push(paragraph);
                }

                // 创建脚注对象
                let mut footnote = crate::docx::adapter::docx::Footnote::new();
                footnote.children = children;

                // 将脚注添加到文档
                doc.doc.options.footnotes.insert(note.no, footnote);
                println!(
                    "【generate】已添加脚注 #{}: {} 个段落",
                    note.no,
                    doc.doc
                        .options
                        .footnotes
                        .get(&note.no)
                        .unwrap()
                        .children
                        .len()
                );
            }
        }

        println!(
            "【generate】脚注处理完成，总脚注数: {}",
            doc.doc.options.footnotes.len()
        );
    } else if !bottom_notes {
        println!("【generate】脚注配置为原位置显示，不处理页面底部脚注");
    } else {
        println!("【generate】没有脚注需要处理");
    }

    // 创建 section 属性
    // 注意：在这里我们只是记录了 section 属性，但实际上没有使用它
    // 这是因为在 Rust 版本中，section 属性的设置方式与 TypeScript 版本不同
    // 在 TypeScript 版本中，section 属性是通过 sesctionProps 设置的
    // 在 Rust 版本中，section 属性是通过 doc.doc.set_section_properties 设置的
    // 但是，我们仍然需要记录 line_height 的使用，以便与 TypeScript 版本保持一致
    let _footer_margin = convert_inches_to_twip(print.page_number_top_margin - options.line_height);

    // 在 Rust 版本中，我们不使用 sections 的方式来组织文档
    // 但是，我们仍然需要记录 print_preface_page 的使用，以便与 TypeScript 版本保持一致
    // 在 TypeScript 版本中，print_preface_page 用于决定是否将 sectionMainNoPageNum 添加到文档的 sections 中
    // 在 Rust 版本中，我们可以通过其他方式来实现相同的功能
    // 例如，我们可以在生成文档时，根据 print_preface_page 的值来决定是否生成前言页
    // 但是，由于我们已经生成了文档，所以这里只是记录 print_preface_page 的使用
    let _print_preface_page_used = print_preface_page;

    // 将 sections 添加到文档
    doc.doc.options.sections.clear();

    if let Some(title_section) = section_title_page {
        doc.doc.options.sections.push(title_section);
        println!("【generate】已添加标题页 section");
    }

    if !section_main_no_page_num.children.is_empty() && print_preface_page {
        doc.doc.options.sections.push(section_main_no_page_num);
        println!("【generate】已添加序言页 section");
    }

    if !section_main.children.is_empty() {
        doc.doc.options.sections.push(section_main);
        println!("【generate】已添加主要内容 section");
    }

    // 重新创建文档以使用 sections
    doc.doc.docx = doc.doc.create_document();
    println!(
        "【generate】已重新创建文档以使用 sections，总 section 数: {}",
        doc.doc.options.sections.len()
    );

    // 返回页数
    current_page + 1
}

/// 完成文档生成并保存
pub fn finish_doc(doc: Document, filepath: &str) -> DocxResult<()> {
    doc.save(filepath).map_err(|e| DocxError::AdapterError(e))
}

/// 生成DOCX文档
pub async fn generate_docx(options: DocxOptions, parsed_document: &ParseOutput) -> DocxResult<()> {
    // 创建一个可变的解析结果副本
    let mut parsed_document_copy = parsed_document.clone();

    // 处理行
    crate::docx::line_processor::process_document_lines(&mut parsed_document_copy, &options.config);

    // 更新选项中的解析结果
    let mut options_with_lines = options.clone();
    options_with_lines.parsed = Some(parsed_document_copy);

    // 生成文档
    let mut doc = init_doc(options_with_lines.clone()).await;
    // 确保标题页处理标志被正确设置
    doc.options.title_page_processed = options_with_lines.title_page_processed;
    generate(&mut doc, &options_with_lines, None);
    finish_doc(doc.doc, &options_with_lines.filepath)
}

/// 获取DOCX文档
pub async fn get_docx(options: DocxOptions) -> DocxResult<()> {
    println!("【get_docx】开始获取 DOCX 文档 - docx_maker::get_docx 函数");
    println!(
        "【get_docx】title_page_processed = {}",
        options.title_page_processed
    );

    // 如果没有解析结果，则返回错误
    if options.parsed.is_none() {
        println!("【get_docx】错误：没有解析结果");
        return Err(DocxError::InvalidConfig("没有解析结果".to_string()));
    }

    // 获取解析结果
    let mut parsed_document_copy = options.parsed.as_ref().unwrap().clone();

    println!(
        "【get_docx】标题页元素数量: {}",
        parsed_document_copy.title_page.len()
    );
    for (key, tokens) in &parsed_document_copy.title_page {
        println!("【get_docx】标题页元素 {}: {} 个 token", key, tokens.len());
    }

    // 处理行
    println!("【get_docx】开始处理行");
    crate::docx::line_processor::process_document_lines(&mut parsed_document_copy, &options.config);
    println!("【get_docx】行处理完成");

    // 更新选项中的解析结果
    let mut options_with_lines = options.clone();
    options_with_lines.parsed = Some(parsed_document_copy);
    println!(
        "【get_docx】标题页是否已处理: {}",
        options_with_lines.title_page_processed
    );

    // 生成文档
    println!("【get_docx】开始初始化文档");
    let mut doc = init_doc(options_with_lines.clone()).await;
    // 确保标题页处理标志被正确设置
    let old_value = doc.options.title_page_processed;
    doc.options.title_page_processed = options_with_lines.title_page_processed;
    println!(
        "【get_docx】文档初始化完成，标题页是否已处理: {} (原值: {})",
        doc.options.title_page_processed, old_value
    );
    println!("【get_docx】DocxContext 对象 ID: {:p}", &doc);

    // 生成文档内容
    println!("【get_docx】开始生成文档内容");
    let page_count = generate(&mut doc, &options_with_lines, None);
    println!("【get_docx】文档内容生成完成，页数: {}", page_count);

    // 保存文档
    println!("【get_docx】开始保存文档");
    let result = finish_doc(doc.doc, &options_with_lines.filepath);
    println!("【get_docx】文档保存完成");

    result
}

/// 获取DOCX统计信息
pub async fn get_docx_stats(options: DocxOptions) -> DocxResult<DocxStats> {
    println!("【get_docx_stats】开始获取 DOCX 统计信息 - docx_maker::get_docx_stats 函数");
    println!(
        "【get_docx_stats】title_page_processed = {}",
        options.title_page_processed
    );

    // 如果没有解析结果，则返回错误
    if options.parsed.is_none() {
        println!("【get_docx_stats】错误：没有解析结果");
        return Err(DocxError::InvalidConfig("没有解析结果".to_string()));
    }

    // 获取解析结果
    let mut parsed_document_copy = options.parsed.as_ref().unwrap().clone();

    println!(
        "【get_docx_stats】标题页元素数量: {}",
        parsed_document_copy.title_page.len()
    );
    for (key, tokens) in &parsed_document_copy.title_page {
        println!(
            "【get_docx_stats】标题页元素 {}: {} 个 token",
            key,
            tokens.len()
        );
    }

    // 处理行
    println!("【get_docx_stats】开始处理行");
    crate::docx::line_processor::process_document_lines(&mut parsed_document_copy, &options.config);
    println!("【get_docx_stats】行处理完成");

    // 更新选项中的解析结果
    let mut options_with_lines = options.clone();
    options_with_lines.parsed = Some(parsed_document_copy);
    println!(
        "【get_docx_stats】标题页是否已处理: {}",
        options_with_lines.title_page_processed
    );

    // 生成文档
    println!("【get_docx_stats】开始初始化文档");
    let mut doc = init_doc(options_with_lines.clone()).await;
    // 确保标题页处理标志被正确设置
    let old_value = doc.options.title_page_processed;
    doc.options.title_page_processed = options_with_lines.title_page_processed;
    println!(
        "【get_docx_stats】文档初始化完成，标题页是否已处理: {} (原值: {})",
        doc.options.title_page_processed, old_value
    );
    println!("【get_docx_stats】DocxContext 对象 ID: {:p}", &doc);

    println!("【get_docx_stats】开始生成文档内容");
    let mut line_map = HashMap::new();
    let page_count = generate(&mut doc, &options_with_lines, Some(&mut line_map));
    println!("【get_docx_stats】文档内容生成完成，页数: {}", page_count);

    println!("【get_docx_stats】统计信息获取完成");
    Ok(DocxStats {
        page_count,
        page_count_real: page_count,
        line_map,
    })
}

/// 获取DOCX文档的Base64编码
pub async fn get_docx_base64(options: DocxOptions) -> DocxResult<DocxAsBase64> {
    println!("开始获取 DOCX 文档的 Base64 编码 - docx_maker::get_docx_base64 函数");

    // 如果没有解析结果，则返回错误
    if options.parsed.is_none() {
        println!("错误：没有解析结果");
        return Err(DocxError::InvalidConfig("没有解析结果".to_string()));
    }

    // 获取解析结果
    let mut parsed_document_copy = options.parsed.as_ref().unwrap().clone();

    println!("标题页元素数量: {}", parsed_document_copy.title_page.len());
    for (key, tokens) in &parsed_document_copy.title_page {
        println!("标题页元素 {}: {} 个 token", key, tokens.len());
    }

    // 处理行
    println!("开始处理行");
    crate::docx::line_processor::process_document_lines(&mut parsed_document_copy, &options.config);
    println!("行处理完成");

    // 更新选项中的解析结果
    let mut options_with_lines = options.clone();
    options_with_lines.parsed = Some(parsed_document_copy);

    // 确保标题页处理标志被正确传递
    println!(
        "标题页是否已处理: {}",
        options_with_lines.title_page_processed
    );

    // 生成文档
    println!("开始初始化文档");
    let mut doc = init_doc(options_with_lines.clone()).await;
    // 确保标题页处理标志被正确设置
    let old_value = doc.options.title_page_processed;
    doc.options.title_page_processed = options_with_lines.title_page_processed;
    println!(
        "文档初始化完成，标题页是否已处理: {} (原值: {})",
        doc.options.title_page_processed, old_value
    );
    println!("DocxContext 对象 ID: {:p}", &doc);

    println!("开始生成文档内容");
    let mut line_map = HashMap::new();
    let page_count = generate(&mut doc, &options_with_lines, Some(&mut line_map));
    println!("文档内容生成完成，页数: {}", page_count);

    println!("开始转换为 Base64");
    let data = doc
        .doc
        .to_base64()
        .map_err(|e| DocxError::AdapterError(e))?;
    println!("Base64 转换完成");

    println!("docx_maker::get_docx_base64 函数处理完成");
    Ok(DocxAsBase64 {
        data,
        stats: DocxStats {
            page_count,
            page_count_real: page_count,
            line_map,
        },
    })
}
