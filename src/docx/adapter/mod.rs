//! Docx 适配器模块
//!
//! 该模块提供了与原始 TypeScript 版本 docxmaker.ts 兼容的 API

use docx_rs;
use std::collections::HashMap;
use thiserror::Error;
use std::error::Error as StdError;

/// 下划线类型
#[derive(Debug, Clone, Copy)]
pub enum UnderlineType {
    /// 单线下划线
    Single,
    /// 双线下划线
    Double,
    /// 虚线下划线
    Dash,
    /// 点线下划线
    Dotted,
    /// 波浪线下划线
    Wave,
    /// 无下划线
    None,
}

/// 下划线类型常量
pub struct UnderlineTypeConst;

impl UnderlineTypeConst {
    /// 单线下划线
    pub const SINGLE: UnderlineType = UnderlineType::Single;
    /// 双线下划线
    pub const DOUBLE: UnderlineType = UnderlineType::Double;
    /// 虚线下划线
    pub const DASH: UnderlineType = UnderlineType::Dash;
    /// 点线下划线
    pub const DOTTED: UnderlineType = UnderlineType::Dotted;
    /// 波浪线下划线
    pub const WAVE: UnderlineType = UnderlineType::Wave;
    /// 无下划线
    pub const NONE: UnderlineType = UnderlineType::None;
}

/// 对齐方式
#[derive(Debug, Clone, Copy)]
pub enum AlignmentType {
    /// 左对齐
    Left,
    /// 居中对齐
    Center,
    /// 右对齐
    Right,
    /// 两端对齐
    Justify,
}

impl AlignmentType {
    /// 转换为 docx-rs 的 AlignmentType
    pub fn to_docx_alignment(&self) -> docx_rs::AlignmentType {
        match self {
            AlignmentType::Left => docx_rs::AlignmentType::Left,
            AlignmentType::Center => docx_rs::AlignmentType::Center,
            AlignmentType::Right => docx_rs::AlignmentType::Right,
            AlignmentType::Justify => docx_rs::AlignmentType::Justified,
        }
    }
}

/// 对齐方式常量
pub struct AlignmentTypeConst;

impl AlignmentTypeConst {
    /// 左对齐
    pub const LEFT: AlignmentType = AlignmentType::Left;
    /// 居中对齐
    pub const CENTER: AlignmentType = AlignmentType::Center;
    /// 右对齐
    pub const RIGHT: AlignmentType = AlignmentType::Right;
    /// 两端对齐
    pub const JUSTIFY: AlignmentType = AlignmentType::Justify;
}

/// 分页符类型
#[derive(Debug, Clone, Copy)]
pub enum BreakType {
    /// 页面分页符
    Page,
    /// 列分页符
    Column,
    /// 文本换行符
    TextWrapping,
}

impl BreakType {
    /// 转换为 docx-rs 的 BreakType
    pub fn to_docx_break(&self) -> docx_rs::BreakType {
        match self {
            BreakType::Page => docx_rs::BreakType::Page,
            BreakType::Column => docx_rs::BreakType::Column,
            BreakType::TextWrapping => docx_rs::BreakType::TextWrapping,
        }
    }
}

/// 分页符类型常量
pub struct BreakTypeConst;

impl BreakTypeConst {
    /// 页面分页符
    pub const PAGE: BreakType = BreakType::Page;
    /// 列分页符
    pub const COLUMN: BreakType = BreakType::Column;
    /// 文本换行符
    pub const TEXT_WRAPPING: BreakType = BreakType::TextWrapping;
}

/// 字体类型
#[derive(Debug, Clone, Copy)]
pub enum FontType {
    /// 普通字体
    Normal,
    /// 粗体
    Bold,
    /// 斜体
    Italic,
    /// 粗斜体
    BoldItalic,
}

/// 字体类型常量
pub struct FontTypeConst;

impl FontTypeConst {
    /// 普通字体
    pub const NORMAL: FontType = FontType::Normal;
    /// 粗体
    pub const BOLD: FontType = FontType::Bold;
    /// 斜体
    pub const ITALIC: FontType = FontType::Italic;
    /// 粗斜体
    pub const BOLD_ITALIC: FontType = FontType::BoldItalic;
}

/// 行规则类型
#[derive(Debug, Clone, Copy)]
pub enum LineRuleType {
    /// 精确值
    Exact,
    /// 至少值
    AtLeast,
    /// 自动值
    Auto,
}

impl LineRuleType {
    /// 转换为 docx-rs 的 LineSpacingType
    pub fn to_docx_line_spacing_type(&self) -> docx_rs::LineSpacingType {
        match self {
            LineRuleType::Exact => docx_rs::LineSpacingType::Exact,
            LineRuleType::AtLeast => docx_rs::LineSpacingType::AtLeast,
            LineRuleType::Auto => docx_rs::LineSpacingType::Auto,
        }
    }
}

/// 行规则类型常量
pub struct LineRuleTypeConst;

impl LineRuleTypeConst {
    /// 精确值
    pub const EXACT: LineRuleType = LineRuleType::Exact;
    /// 至少值
    pub const AT_LEAST: LineRuleType = LineRuleType::AtLeast;
    /// 自动值
    pub const AUTO: LineRuleType = LineRuleType::Auto;
}

/// 框架锚点类型
#[derive(Debug, Clone, Copy)]
pub enum FrameAnchorType {
    /// 页面
    Page,
    /// 边距
    Margin,
    /// 文本
    Text,
}

/// 框架锚点类型常量
pub struct FrameAnchorTypeConst;

impl FrameAnchorTypeConst {
    /// 页面
    pub const PAGE: FrameAnchorType = FrameAnchorType::Page;
    /// 边距
    pub const MARGIN: FrameAnchorType = FrameAnchorType::Margin;
    /// 文本
    pub const TEXT: FrameAnchorType = FrameAnchorType::Text;
}

/// 水平位置对齐
#[derive(Debug, Clone, Copy)]
pub enum HorizontalPositionAlign {
    /// 左对齐
    Left,
    /// 居中对齐
    Center,
    /// 右对齐
    Right,
    /// 内部对齐
    Inside,
    /// 外部对齐
    Outside,
}

/// 水平位置对齐常量
pub struct HorizontalPositionAlignConst;

impl HorizontalPositionAlignConst {
    /// 左对齐
    pub const LEFT: HorizontalPositionAlign = HorizontalPositionAlign::Left;
    /// 居中对齐
    pub const CENTER: HorizontalPositionAlign = HorizontalPositionAlign::Center;
    /// 右对齐
    pub const RIGHT: HorizontalPositionAlign = HorizontalPositionAlign::Right;
    /// 内部对齐
    pub const INSIDE: HorizontalPositionAlign = HorizontalPositionAlign::Inside;
    /// 外部对齐
    pub const OUTSIDE: HorizontalPositionAlign = HorizontalPositionAlign::Outside;
}

/// 垂直位置对齐
#[derive(Debug, Clone, Copy)]
pub enum VerticalPositionAlign {
    /// 顶部对齐
    Top,
    /// 居中对齐
    Center,
    /// 底部对齐
    Bottom,
    /// 内部对齐
    Inside,
    /// 外部对齐
    Outside,
}

/// 垂直位置对齐常量
pub struct VerticalPositionAlignConst;

impl VerticalPositionAlignConst {
    /// 顶部对齐
    pub const TOP: VerticalPositionAlign = VerticalPositionAlign::Top;
    /// 居中对齐
    pub const CENTER: VerticalPositionAlign = VerticalPositionAlign::Center;
    /// 底部对齐
    pub const BOTTOM: VerticalPositionAlign = VerticalPositionAlign::Bottom;
    /// 内部对齐
    pub const INSIDE: VerticalPositionAlign = VerticalPositionAlign::Inside;
    /// 外部对齐
    pub const OUTSIDE: VerticalPositionAlign = VerticalPositionAlign::Outside;
}

/// 宽度类型
#[derive(Debug, Clone, Copy)]
pub enum WidthType {
    /// 自动
    Auto,
    /// DXA
    DXA,
    /// 百分比
    Percentage,
}

/// 宽度类型常量
pub struct WidthTypeConst;

impl WidthTypeConst {
    /// 自动
    pub const AUTO: WidthType = WidthType::Auto;
    /// DXA
    pub const DXA: WidthType = WidthType::DXA;
    /// 百分比
    pub const PERCENTAGE: WidthType = WidthType::Percentage;
}

/// 页码
pub struct PageNumber;

impl PageNumber {
    /// 当前页码
    pub const CURRENT: &'static str = "{n}";
    /// 总页数
    pub const TOTAL: &'static str = "{total}";
}

/// 将英寸转换为 twip
pub fn convert_inches_to_twip(inches: f32) -> i32 {
    (inches * 1440.0) as i32
}

/// 将英寸转换为 磅
pub fn convert_inches_to_point(inches: f32) -> f32 {
    inches * 72.0
}

/// 将 磅 转为 twip (常用于字体)
pub fn convert_point_to_inches(point: f32) -> f32 {
    point / 72.0
}


/// 将 磅 转为 twip (常用于字体)
pub fn convert_point_to_twip(point: f32) -> i32 {
    convert_inches_to_twip(convert_point_to_inches(point))
}

/// 将英寸转换为 twip (驼峰命名，兼容 TypeScript 版本)
#[allow(non_snake_case)]
pub fn convertInchesToTwip(inches: f32) -> i32 {
    convert_inches_to_twip(inches)
}

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

/// 运行属性
#[derive(Debug, Clone, Default)]
pub struct RunProps {
    pub size: Option<usize>,
    pub font: Option<String>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<UnderlineType>,
    pub color: Option<String>,
    pub character_spacing: Option<i32>,
    pub superscript: Option<bool>,
    pub subscript: Option<bool>,
    pub break_type: Option<BreakType>,
    pub break_before: Option<bool>, // 添加一个新属性，表示在文本之前添加换行符
}

impl RunProps {
    /// 创建新的运行属性
    pub fn new() -> Self {
        Self::default()
    }

    /// 设置字体大小
    pub fn size(mut self, size: usize) -> Self {
        self.size = Some(size);
        self
    }

    /// 设置字体
    pub fn font(mut self, font: &str) -> Self {
        self.font = Some(font.to_string());
        self
    }

    /// 设置粗体
    pub fn bold(mut self) -> Self {
        self.bold = Some(true);
        self
    }

    /// 设置斜体
    pub fn italic(mut self) -> Self {
        self.italic = Some(true);
        self
    }

    /// 设置下划线
    pub fn underline(mut self, underline: UnderlineType) -> Self {
        self.underline = Some(underline);
        self
    }

    /// 设置颜色
    pub fn color(mut self, color: &str) -> Self {
        self.color = Some(color.to_string());
        self
    }

    /// 设置上标
    pub fn superscript(mut self) -> Self {
        self.superscript = Some(true);
        self
    }

    /// 设置下标
    pub fn subscript(mut self) -> Self {
        self.subscript = Some(true);
        self
    }

    /// 转换为 docx-rs 的 RunProperty
    pub fn to_run_property(&self) -> docx_rs::RunProperty {
        let mut property = docx_rs::RunProperty::new();

        if let Some(size) = self.size {
            property = property.size(size);
        }

        if let Some(font) = &self.font {
            let run_fonts = docx_rs::RunFonts::new()
                .east_asia(font)
                .ascii(font)
                .hi_ansi(font);
            property = property.fonts(run_fonts);
        }

        if let Some(true) = self.bold {
            property = property.bold();
        }

        if let Some(true) = self.italic {
            property = property.italic();
        }

        if let Some(underline) = &self.underline {
            match underline {
                UnderlineType::Single => property = property.underline("single"),
                UnderlineType::Double => property = property.underline("double"),
                UnderlineType::Dash => property = property.underline("dash"),
                UnderlineType::Dotted => property = property.underline("dotted"),
                UnderlineType::Wave => property = property.underline("wave"),
                UnderlineType::None => {}
            }
        }

        if let Some(color) = &self.color {
            property = property.color(color);
        }

        if let Some(spacing) = self.character_spacing {
            property = property.spacing(spacing);
        }

        property
    }
}

/// 样式缓存
#[derive(Debug, Clone)]
pub struct StyleStash {
    pub bold_italic: bool,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub override_color: Option<String>,
    pub italic_global: bool,
    pub italic_dynamic: bool,
    /// 当前文本颜色
    pub current_color: String,
}

impl Default for StyleStash {
    fn default() -> Self {
        Self {
            bold_italic: false,
            bold: false,
            italic: false,
            underline: false,
            override_color: None,
            italic_global: false,
            italic_dynamic: false,
            current_color: "#000000".to_string(),
        }
    }
}

/// 当前注释
#[derive(Debug, Clone, Default)]
pub struct CurrentNote {
    pub page_idx: i32,
    pub note: Vec<String>,
}

/// 行结构
#[derive(Debug, Clone)]
pub struct LineStruct {
    pub sections: Vec<String>,
    pub scene: String,
    pub page: usize,
    pub cumulative_duration: f32,
}

/// DOCX统计信息
#[derive(Debug, Clone)]
pub struct DocxStats {
    pub page_count: usize,
    pub page_count_real: usize,
    pub line_map: HashMap<usize, LineStruct>,
}

/// DOCX Base64 结果
#[derive(Debug, Clone)]
pub struct DocxAsBase64 {
    pub data: String,
    pub stats: DocxStats,
}

/// 段落框架
#[derive(Debug, Clone, Default)]
pub struct ParagraphFrame {
    /// 宽度
    pub width: Option<i32>,
    /// 高度
    pub height: Option<i32>,
    /// 水平锚点
    pub anchor_horizontal: Option<FrameAnchorType>,
    /// 垂直锚点
    pub anchor_vertical: Option<FrameAnchorType>,
    /// 水平对齐
    pub x_align: Option<HorizontalPositionAlign>,
    /// 垂直对齐
    pub y_align: Option<VerticalPositionAlign>,
    /// 页面高度
    pub page_height: Option<i32>,
    /// 顶部边距
    pub top_margin: Option<i32>,
    /// 底部边距
    pub bottom_margin: Option<i32>,
    pub line_height: Option<i32>,
}

/// 导出 docx 子模块
pub mod docx;
