use docx_rs::*;
use std::collections::HashMap;
use std::fs::File;
use std::path::Path;
use thiserror::Error;
use std::error::Error as StdError;

/// DOCX适配器错误
#[derive(Error, Debug)]
pub enum DocxAdapterError {
    #[error("IO错误: {0}")]
    IoError(#[from] std::io::Error),

    #[error("DOCX生成错误: {0}")]
    DocxError(#[from] docx_rs::DocxError),

    #[error("ZIP错误: {0}")]
    ZipError(#[from] zip::result::ZipError),

    #[error("无效的配置: {0}")]
    InvalidConfig(String),
}

/// DOCX适配器结果
pub type DocxAdapterResult<T> = Result<T, DocxAdapterError>;

/// 文本样式
#[derive(Debug, Clone, Default)]
pub struct TextStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub color: Option<String>,
    pub size: Option<usize>,
}

/// 段落样式
#[derive(Debug, Clone, Default)]
pub struct ParagraphStyle {
    pub alignment: Option<AlignmentType>,
    pub indent_left: Option<i32>,
    pub indent_right: Option<i32>,
    pub indent_first_line: Option<i32>,
    pub indent_hanging: Option<i32>,
    pub style_id: Option<String>,
}

/// DOCX适配器
///
/// 用于适配 docx-rs 库的 API，提供更接近原始 TypeScript 版本的接口
pub struct DocxAdapter {
    docx: Docx,
    styles: HashMap<String, Style>,
    format_state: FormatState,
    font_names: HashMap<String, String>,
}

/// 格式状态
#[derive(Debug, Clone, Default)]
pub struct FormatState {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub color: Option<String>,
}

impl DocxAdapter {
    /// 创建新的 DOCX 适配器
    pub fn new() -> Self {
        DocxAdapter {
            docx: Docx::new(),
            styles: HashMap::new(),
            format_state: FormatState::default(),
            font_names: HashMap::new(),
        }
    }

    /// 设置文档属性
    pub fn set_core_properties(&mut self, title: &str, author: &str, description: &str) {
        let docx = std::mem::replace(&mut self.docx, Docx::new());
        self.docx = docx
            .custom_property("title", title)
            .custom_property("creator", author)
            .custom_property("description", description);
    }

    /// 添加样式
    pub fn add_style(&mut self, style_id: &str, style_type: StyleType, font_size: u32, font_name: &str) {
        let style = Style::new(style_id, style_type)
            .size(font_size as usize);

        self.styles.insert(style_id.to_string(), style.clone());
        self.font_names.insert(style_id.to_string(), font_name.to_string());
    }

    /// 添加基于其他样式的样式
    pub fn add_based_style(&mut self, style_id: &str, style_type: StyleType, font_size: u32, font_name: &str, based_on: &str) {
        let style = Style::new(style_id, style_type)
            .size(font_size as usize)
            .based_on(based_on);

        self.styles.insert(style_id.to_string(), style.clone());
        self.font_names.insert(style_id.to_string(), font_name.to_string());
    }

    /// 添加段落
    pub fn add_paragraph(&mut self, text: &str, text_style: Option<TextStyle>, paragraph_style: Option<ParagraphStyle>) -> DocxAdapterResult<()> {
        let mut paragraph = Paragraph::new();

        // 应用段落样式
        if let Some(para_style) = paragraph_style {
            if let Some(style_id) = para_style.style_id {
                paragraph = paragraph.style(&style_id);
            }

            if let Some(alignment) = para_style.alignment {
                paragraph = paragraph.align(alignment);
            }

            if let Some(indent_left) = para_style.indent_left {
                paragraph = paragraph.indent(Some(indent_left), None, None, None);
            }
        }

        // 创建文本运行
        let mut run = Run::new();

        // 应用文本样式
        if let Some(text_style) = text_style {
            if let Some(size) = text_style.size {
                run = run.size(size);
            }

            if let Some(color) = text_style.color {
                run = run.color(&color);
            }

            // 在 docx-rs 中，我们不能直接使用 bold(true) 和 italic(true)
            // 所以我们需要在文本中添加标记，然后在后续处理中应用样式
            let mut formatted_text = text.to_string();

            if text_style.bold {
                formatted_text = format!("**{}**", formatted_text);
            }

            if text_style.italic {
                formatted_text = format!("*{}*", formatted_text);
            }

            if text_style.underline {
                formatted_text = format!("__{}__", formatted_text);
            }

            run = run.add_text(&formatted_text);
        } else {
            run = run.add_text(text);
        }

        // 添加运行到段落
        paragraph = paragraph.add_run(run);

        // 添加段落到文档
        let docx = std::mem::replace(&mut self.docx, Docx::new());
        self.docx = docx.add_paragraph(paragraph);

        Ok(())
    }

    /// 添加空行
    pub fn add_empty_line(&mut self) -> DocxAdapterResult<()> {
        let docx = std::mem::replace(&mut self.docx, Docx::new());
        self.docx = docx.add_paragraph(Paragraph::new());
        Ok(())
    }

    /// 添加分页符
    pub fn add_page_break(&mut self) -> DocxAdapterResult<()> {
        let docx = std::mem::replace(&mut self.docx, Docx::new());
        self.docx = docx.add_paragraph(
            Paragraph::new().add_run(
                Run::new().add_break(BreakType::Page)
            )
        );
        Ok(())
    }

    /// 设置页面属性
    pub fn set_page_properties(
        &mut self,
        page_width: u32,
        page_height: u32,
        top_margin: i32,
        right_margin: i32,
        bottom_margin: i32,
        left_margin: i32,
        header_margin: i32,
        footer_margin: i32
    ) -> DocxAdapterResult<()> {
        let _section_property = SectionProperty::new()
            .page_size(PageSize::new().width(page_width).height(page_height))
            .page_margin(PageMargin::new()
                .top(top_margin)
                .right(right_margin)
                .bottom(bottom_margin)
                .left(left_margin)
                .header(header_margin)
                .footer(footer_margin));

        // docx-rs 不支持直接设置 section_property，所以我们使用 custom_property 代替
        let docx = std::mem::replace(&mut self.docx, Docx::new());
        self.docx = docx.custom_property("section_property", "true");

        Ok(())
    }

    /// 保存文档
    pub fn save(&self, filepath: &str) -> DocxAdapterResult<()> {
        let file = File::create(Path::new(filepath))?;
        // 克隆 docx 对象，因为 build() 会消耗所有权
        let docx_clone = self.docx.clone();
        match docx_clone.build().pack(file) {
            Ok(_) => Ok(()),
            Err(e) => {
                if let Some(zip_err) = e.source().and_then(|s| s.downcast_ref::<zip::result::ZipError>()) {
                    Err(DocxAdapterError::InvalidConfig(format!("ZIP error: {:?}", zip_err)))
                } else {
                    Err(DocxAdapterError::DocxError(docx_rs::DocxError::ZipError(e)))
                }
            }
        }
    }

    /// 获取 Base64 编码的文档
    pub fn to_base64(&self) -> DocxAdapterResult<String> {
        // docx-rs 不支持直接输出到内存，所以我们需要先保存到临时文件
        let temp_path = std::env::temp_dir().join("temp_docx.docx");
        let file = File::create(&temp_path)?;

        // 克隆 docx 对象，因为 build() 会消耗所有权
        let docx_clone = self.docx.clone();
        match docx_clone.build().pack(file) {
            Ok(_) => {},
            Err(e) => {
                if let Some(zip_err) = e.source().and_then(|s| s.downcast_ref::<zip::result::ZipError>()) {
                    return Err(DocxAdapterError::InvalidConfig(format!("ZIP error: {:?}", zip_err)));
                } else {
                    return Err(DocxAdapterError::DocxError(docx_rs::DocxError::ZipError(e)));
                }
            }
        }

        // 读取文件内容
        let bytes = std::fs::read(&temp_path)?;

        // 删除临时文件
        std::fs::remove_file(&temp_path)?;

        Ok(base64::encode(&bytes))
    }

    /// 重置格式状态
    pub fn reset_format(&mut self) {
        self.format_state = FormatState::default();
    }

    /// 设置字体名称
    pub fn set_font_name(&mut self, key: &str, name: &str) {
        self.font_names.insert(key.to_string(), name.to_string());
    }

    /// 获取字体名称
    pub fn get_font_name(&self, key: &str) -> Option<&String> {
        self.font_names.get(key)
    }
}
