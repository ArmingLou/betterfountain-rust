//! Docx 命名空间
//!
//! 该模块提供了与原始 TypeScript 版本 docxmaker.ts 中 Docx 命名空间兼容的 API

use super::*;
use base64;
use docx_rs;
use std::fs::File;
use std::path::Path;

/// 段落间距
#[derive(Debug, Clone)]
pub struct ParagraphSpacing {
    pub before: Option<i32>,
    pub after: Option<i32>,
    pub before_lines: Option<i32>,
    pub after_lines: Option<i32>,
    pub line: Option<i32>,
    pub line_rule: Option<super::LineRuleType>,
}

impl ParagraphSpacing {
    /// 创建新的段落间距
    pub fn new() -> Self {
        Self {
            before: None,
            after: None,
            before_lines: None,
            after_lines: None,
            line: None,
            line_rule: None,
        }
    }

    /// 设置段前间距（twips）
    pub fn before(mut self, before: i32) -> Self {
        self.before = Some(before);
        self
    }

    /// 设置段后间距（twips）
    pub fn after(mut self, after: i32) -> Self {
        self.after = Some(after);
        self
    }

    /// 设置段前间距（行数）
    pub fn before_lines(mut self, before_lines: i32) -> Self {
        self.before_lines = Some(before_lines);
        self
    }

    /// 设置段后间距（行数）
    pub fn after_lines(mut self, after_lines: i32) -> Self {
        self.after_lines = Some(after_lines);
        self
    }

    /// 设置行间距（twips）
    pub fn line(mut self, line: i32) -> Self {
        self.line = Some(line);
        self
    }

    /// 设置行间距规则
    pub fn line_rule(mut self, line_rule: super::LineRuleType) -> Self {
        self.line_rule = Some(line_rule);
        self
    }

    /// 转换为 docx-rs 的 LineSpacing
    pub fn to_docx_line_spacing(&self) -> docx_rs::LineSpacing {
        let mut spacing = docx_rs::LineSpacing::new();

        if let Some(before) = self.before {
            println!("【适配器日志】设置段前间距: {} twips", before);
            spacing = spacing.before(before as u32);
        }

        if let Some(after) = self.after {
            println!("【适配器日志】设置段后间距: {} twips", after);
            spacing = spacing.after(after as u32);
        }

        if let Some(before_lines) = self.before_lines {
            println!("【适配器日志】设置段前行数: {}", before_lines);
            spacing = spacing.before_lines(before_lines as u32);
        }

        if let Some(after_lines) = self.after_lines {
            println!("【适配器日志】设置段后行数: {}", after_lines);
            spacing = spacing.after_lines(after_lines as u32);
        }

        if let Some(line) = self.line {
            println!(
                "【适配器日志】设置行距: {} twips ({:.1}倍单倍行距)",
                line,
                line as f32 / 240.0
            );
            spacing = spacing.line(line);
        }

        if let Some(line_rule) = self.line_rule {
            let rule_type = line_rule.to_docx_line_spacing_type();
            println!("【适配器日志】设置行距类型: {:?}", rule_type);
            spacing = spacing.line_rule(rule_type);
        }

        spacing
    }
}

/// 框架锚点
#[derive(Debug, Clone)]
pub struct FrameAnchor {
    pub horizontal: FrameAnchorType,
    pub vertical: FrameAnchorType,
}

impl FrameAnchor {
    /// 创建新的框架锚点
    pub fn new(horizontal: FrameAnchorType, vertical: FrameAnchorType) -> Self {
        Self {
            horizontal,
            vertical,
        }
    }
}

/// 框架对齐
#[derive(Debug, Clone)]
pub struct FrameAlignment {
    pub x: HorizontalPositionAlign,
    pub y: VerticalPositionAlign,
}

impl FrameAlignment {
    /// 创建新的框架对齐
    pub fn new(x: HorizontalPositionAlign, y: VerticalPositionAlign) -> Self {
        Self { x, y }
    }
}

/// 框架
#[derive(Debug, Clone)]
pub struct Frame {
    pub frame_type: String,
    pub width: i32,
    pub height: i32,
    pub anchor: FrameAnchor,
    pub alignment: FrameAlignment,
}

impl Frame {
    /// 创建新的框架
    pub fn new(
        frame_type: &str,
        width: i32,
        height: i32,
        anchor: FrameAnchor,
        alignment: FrameAlignment,
    ) -> Self {
        Self {
            frame_type: frame_type.to_string(),
            width,
            height,
            anchor,
            alignment,
        }
    }
}

/// 样式
#[derive(Debug, Clone)]
pub struct Styles {
    pub default: DefaultStyles,
    pub paragraph_styles: Vec<ParagraphStyle>,
    pub character_styles: Vec<CharacterStyle>,
}

impl Styles {
    /// 创建新的样式
    pub fn new() -> Self {
        Self {
            default: DefaultStyles::new(),
            paragraph_styles: Vec::new(),
            character_styles: Vec::new(),
        }
    }
}

/// 默认样式
#[derive(Debug, Clone)]
pub struct DefaultStyles {
    pub document: DocumentStyle,
}

impl DefaultStyles {
    /// 创建新的默认样式
    pub fn new() -> Self {
        Self {
            document: DocumentStyle::new(),
        }
    }
}

/// 文档样式
#[derive(Debug, Clone)]
pub struct DocumentStyle {
    pub run: RunStyle,
    pub paragraph: ParagraphStyle,
}

impl DocumentStyle {
    /// 创建新的文档样式
    pub fn new() -> Self {
        Self {
            run: RunStyle::new(),
            paragraph: ParagraphStyle::new(),
        }
    }
}

/// 运行样式
#[derive(Debug, Clone)]
pub struct RunStyle {
    pub font: Option<String>,
    pub size: Option<usize>,
    pub character_spacing: Option<i32>,
    pub color: Option<String>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
}

impl RunStyle {
    /// 创建新的运行样式
    pub fn new() -> Self {
        Self {
            font: None,
            size: None,
            character_spacing: None,
            color: None,
            bold: None,
            italic: None,
        }
    }
}

/// 段落样式
#[derive(Debug, Clone)]
pub struct ParagraphStyle {
    pub id: Option<String>,
    pub name: Option<String>,
    pub based_on: Option<String>,
    pub next: Option<String>,
    pub run: Option<RunStyle>,
    pub paragraph: Option<Box<ParagraphStyle>>,
    pub spacing: Option<ParagraphSpacing>,
    pub indent: Option<ParagraphIndent>,
}

impl ParagraphStyle {
    /// 创建新的段落样式
    pub fn new() -> Self {
        Self {
            id: None,
            name: None,
            based_on: None,
            next: None,
            run: None,
            paragraph: None,
            spacing: None,
            indent: None,
        }
    }
}

/// 段落缩进
#[derive(Debug, Clone)]
pub struct ParagraphIndent {
    pub left: Option<i32>,
    pub right: Option<i32>,
}

impl ParagraphIndent {
    /// 创建新的段落缩进
    pub fn new() -> Self {
        Self {
            left: None,
            right: None,
        }
    }
}

/// 字符样式
#[derive(Debug, Clone)]
pub struct CharacterStyle {
    pub id: Option<String>,
    pub name: Option<String>,
    pub based_on: Option<String>,
    pub quick_format: Option<bool>,
    pub run: Option<RunStyle>,
}

impl CharacterStyle {
    /// 创建新的字符样式
    pub fn new() -> Self {
        Self {
            id: None,
            name: None,
            based_on: None,
            quick_format: None,
            run: None,
        }
    }
}

/// 节
#[derive(Debug, Clone)]
pub struct Section {
    pub properties: SectionProperties,
    pub headers: Option<Headers>,
    pub footers: Option<Footers>,
    pub children: Vec<SectionChild>,
}

impl Section {
    /// 创建新的节
    pub fn new() -> Self {
        Self {
            properties: SectionProperties::new(),
            headers: None,
            footers: None,
            children: Vec::new(),
        }
    }
}

/// 节子元素
#[derive(Debug, Clone)]
pub enum SectionChild {
    Paragraph(Paragraph),
    Table(Table),
    PageBreak,
}

/// 节属性
#[derive(Debug, Clone)]
pub struct SectionProperties {
    pub page: Option<PageProperties>,
}

impl SectionProperties {
    /// 创建新的节属性
    pub fn new() -> Self {
        Self { page: None }
    }
}

/// 页面属性
#[derive(Debug, Clone)]
pub struct PageProperties {
    pub size: Option<PageSize>,
    pub margin: Option<PageMargin>,
    pub page_numbers: Option<PageNumbers>,
}

impl PageProperties {
    /// 创建新的页面属性
    pub fn new() -> Self {
        Self {
            size: None,
            margin: None,
            page_numbers: None,
        }
    }
}

/// 页面大小
#[derive(Debug, Clone)]
pub struct PageSize {
    pub height: i32,
    pub width: i32,
}

impl PageSize {
    /// 创建新的页面大小
    pub fn new(height: i32, width: i32) -> Self {
        Self { height, width }
    }
}

/// 页面边距
#[derive(Debug, Clone)]
pub struct PageMargin {
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
    pub left: i32,
    pub header: i32,
    pub footer: i32,
}

impl PageMargin {
    /// 创建新的页面边距
    pub fn new(top: i32, right: i32, bottom: i32, left: i32, header: i32, footer: i32) -> Self {
        Self {
            top,
            right,
            bottom,
            left,
            header,
            footer,
        }
    }
}

/// 页码
#[derive(Debug, Clone)]
pub struct PageNumbers {
    pub start: i32,
}

impl PageNumbers {
    /// 创建新的页码
    pub fn new(start: i32) -> Self {
        Self { start }
    }
}

/// 页眉
#[derive(Debug, Clone)]
pub struct Headers {
    pub default: Header,
}

impl Headers {
    /// 创建新的页眉
    pub fn new(default: Header) -> Self {
        Self { default }
    }
}

/// 页眉
#[derive(Debug, Clone)]
pub struct Header {
    pub children: Vec<Paragraph>,
}

impl Header {
    /// 创建新的页眉
    pub fn new() -> Self {
        Self {
            children: Vec::new(),
        }
    }
}

/// 页脚
#[derive(Debug, Clone)]
pub struct Footers {
    pub default: Footer,
}

impl Footers {
    /// 创建新的页脚
    pub fn new(default: Footer) -> Self {
        Self { default }
    }
}

/// 页脚
#[derive(Debug, Clone)]
pub struct Footer {
    pub children: Vec<Paragraph>,
}

impl Footer {
    /// 创建新的页脚
    pub fn new() -> Self {
        Self {
            children: Vec::new(),
        }
    }
}

/// 脚注
#[derive(Debug, Clone)]
pub struct Footnote {
    pub children: Vec<Paragraph>,
}

impl Footnote {
    /// 创建新的脚注
    pub fn new() -> Self {
        Self {
            children: Vec::new(),
        }
    }
}

/// 文档选项
#[derive(Debug, Clone)]
pub struct DocumentOptions {
    pub creator: String,
    pub description: String,
    pub title: String,
    pub styles: Option<Styles>,
    pub sections: Vec<Section>,
    pub footnotes: HashMap<usize, Footnote>,
}

impl DocumentOptions {
    /// 创建新的文档选项
    pub fn new() -> Self {
        Self {
            creator: String::new(),
            description: String::new(),
            title: String::new(),
            styles: None,
            sections: Vec::new(),
            footnotes: HashMap::new(),
        }
    }
}

/// 文档
pub struct Document {
    pub docx: docx_rs::Docx,
    pub options: DocumentOptions,
}

impl Document {
    /// 创建新的文档
    pub fn new() -> Self {
        Self {
            docx: docx_rs::Docx::new(),
            options: DocumentOptions::new(),
        }
    }

    /// 添加段落
    pub fn add_paragraph(&mut self, paragraph: Paragraph) -> &mut Self {
        self.docx = self.docx.clone().add_paragraph(
            paragraph
                .to_docx_paragraph(self.options.styles.clone(), self.options.footnotes.clone()),
        );
        self
    }

    /// 添加超链接
    pub fn add_hyperlink(&mut self, paragraph: Paragraph, _url: &str) -> &mut Self {
        // 创建一个带有超链接的段落
        let para = paragraph
            .to_docx_paragraph(self.options.styles.clone(), self.options.footnotes.clone());

        // 添加超链接属性
        // 注意：docx-rs 不直接支持超链接，这里只是一个占位符
        // 实际实现需要根据 docx-rs 的 API 进行调整

        // 添加段落到文档
        self.docx = self.docx.clone().add_paragraph(para);
        self
    }

    /// 添加自定义属性
    pub fn custom_property(&mut self, name: &str, value: &str) -> &mut Self {
        self.docx = self.docx.clone().custom_property(name, value);
        self
    }

    /// 添加脚注
    pub fn add_footnote(&mut self, id: usize, paragraphs: Vec<Paragraph>) -> usize {
        // 创建脚注
        let mut footnote = Footnote::new();

        // 添加段落到脚注
        for paragraph in paragraphs {
            footnote.children.push(paragraph);
        }

        // 添加脚注到选项
        self.options.footnotes.insert(id, footnote);

        id
    }

    /// 添加样式到 docx 文档
    fn add_styles_to_docx(&self, mut docx: docx_rs::Docx, styles: &Styles) -> docx_rs::Docx {
        // 添加段落样式
        for paragraph_style in &styles.paragraph_styles {
            if let Some(style_id) = &paragraph_style.id {
                // 创建 docx-rs 的段落样式
                let mut docx_style = docx_rs::Style::new(style_id, docx_rs::StyleType::Paragraph);

                if let Some(name) = &paragraph_style.name {
                    docx_style = docx_style.name(name);
                }

                if let Some(based_on) = &paragraph_style.based_on {
                    docx_style = docx_style.based_on(based_on);
                }

                // 添加运行属性
                if let Some(run_style) = &paragraph_style.run {
                    let mut run_property = docx_rs::RunProperty::new();

                    if let Some(color) = &run_style.color {
                        run_property = run_property.color(color);
                    }

                    if let Some(true) = run_style.bold {
                        run_property = run_property.bold();
                    }

                    if let Some(true) = run_style.italic {
                        run_property = run_property.italic();
                    }

                    // 直接设置字段而不是调用方法
                    docx_style.run_property = run_property;
                }

                // 添加段落属性
                let mut paragraph_property = docx_rs::ParagraphProperty::new();

                // 设置缩进
                if let Some(indent) = &paragraph_style.indent {
                    if let (Some(left), Some(right)) = (indent.left, indent.right) {
                        paragraph_property =
                            paragraph_property.indent(Some(left), None, Some(right), None);
                    }
                }

                // 设置间距
                if let Some(spacing) = &paragraph_style.spacing {
                    if let Some(_line) = spacing.line {
                        paragraph_property =
                            paragraph_property.line_spacing(spacing.to_docx_line_spacing());
                    }
                }

                // 直接设置字段而不是调用方法
                docx_style.paragraph_property = paragraph_property;

                // 添加样式到文档
                docx = docx.add_style(docx_style);
            }
        }

        // 添加字符样式
        for character_style in &styles.character_styles {
            if let Some(style_id) = &character_style.id {
                let mut docx_style = docx_rs::Style::new(style_id, docx_rs::StyleType::Character);

                if let Some(name) = &character_style.name {
                    docx_style = docx_style.name(name);
                }

                if let Some(based_on) = &character_style.based_on {
                    docx_style = docx_style.based_on(based_on);
                }

                // 添加运行属性
                if let Some(run_style) = &character_style.run {
                    let mut run_property = docx_rs::RunProperty::new();

                    if let Some(color) = &run_style.color {
                        run_property = run_property.color(color);
                    }

                    if let Some(true) = run_style.bold {
                        run_property = run_property.bold();
                    }

                    if let Some(true) = run_style.italic {
                        run_property = run_property.italic();
                    }

                    // 直接设置字段而不是调用方法
                    docx_style.run_property = run_property;
                }

                // 添加样式到文档
                docx = docx.add_style(docx_style);
            }
        }

        docx
    }

    /// 创建 Docx.Document 实例
    pub fn create_document(&self) -> docx_rs::Docx {
        let mut docx = self.docx.clone();

        // 设置文档属性
        // 注意：docx-rs 0.4.18-rc44 分支可能不支持 core_properties 方法
        // 直接使用 docx 对象，不设置核心属性

        // 添加样式定义
        if let Some(styles) = &self.options.styles.clone() {
            docx = self.add_styles_to_docx(docx, styles);
        }

        // 处理页眉和页脚
        let mut has_header = false;
        let mut has_footer = false;

        // 检查是否有任何section包含页眉或页脚
        for section in &self.options.sections {
            if section.headers.is_some() {
                has_header = true;
            }
            if section.footers.is_some() {
                has_footer = true;
            }
        }

        // 实现"只有正文页显示页码"的策略：
        // 使用docx-rs的section properties来控制页眉页脚
        // 为前面的页面设置空的first_header和first_footer
        // 为包含页码的section设置正常的页眉页脚

        // 查找有页码的section（通常是最后一个section，即主要内容section）
        let mut main_section_with_page_numbers = None;
        let mut sections_before_main = 0;

        for (index, section) in self.options.sections.iter().enumerate() {
            if let Some(footers) = &section.footers {
                // 检查页脚是否包含页码
                let has_page_numbers = footers.default.children.iter().any(|paragraph| {
                    paragraph
                        .runs
                        .iter()
                        .any(|run| matches!(run, RunType::PageNumber(_)))
                });

                if has_page_numbers {
                    main_section_with_page_numbers = Some(section);
                    sections_before_main = index;
                    break;
                }
            }
        }

        // 新策略：由业务层决定每个section是否有页眉页脚
        // 适配器只负责根据section配置来设置页眉页脚
        if main_section_with_page_numbers.is_some() {
            println!(
                "【create_document】找到包含页码的section（索引: {}），正文页从第{}页开始",
                sections_before_main,
                sections_before_main + 1
            );
            println!("【create_document】页眉页脚的显示由业务层通过section配置决定");
        } else {
            println!("【create_document】未找到包含页码的section");
        }

        // 新策略：检查是否有任何section需要页眉页脚
        // 如果有，则设置全局页眉页脚，但使用section break来控制显示
        let mut has_any_headers_footers = false;
        let mut first_section_with_headers_footers: Option<(usize, &Section)> = None;

        for (section_index, section) in self.options.sections.iter().enumerate() {
            let has_headers = section.headers.is_some();
            let has_footers = section.footers.is_some();
            if has_headers || has_footers {
                has_any_headers_footers = true;
                if first_section_with_headers_footers.is_none() {
                    first_section_with_headers_footers = Some((section_index, section));
                }
            }
        }

        if let Some((first_index, first_section)) = first_section_with_headers_footers {
            println!(
                "【create_document】找到第一个有页眉页脚的section（索引: {}），设置全局页眉页脚",
                first_index
            );

            // 设置全局页眉页脚
            if let Some(headers) = &first_section.headers {
                let mut docx_header = docx_rs::Header::new();
                for paragraph in &headers.default.children {
                    docx_header = docx_header.add_paragraph(paragraph.to_docx_paragraph(
                        self.options.styles.clone(),
                        self.options.footnotes.clone(),
                    ));
                }
                docx = docx.header(docx_header);
                println!(
                    "【create_document】已设置全局页眉，包含 {} 个段落",
                    headers.default.children.len()
                );
            }

            if let Some(footers) = &first_section.footers {
                let mut docx_footer = docx_rs::Footer::new();
                for paragraph in &footers.default.children {
                    docx_footer = docx_footer.add_paragraph(paragraph.to_docx_paragraph(
                        self.options.styles.clone(),
                        self.options.footnotes.clone(),
                    ));
                }
                docx = docx.footer(docx_footer);
                println!(
                    "【create_document】已设置全局页脚，包含 {} 个段落",
                    footers.default.children.len()
                );
            }

            // 关键：如果第一个有页眉页脚的section不是第一个section，
            // 则使用空的first_header和first_footer来覆盖前面的页面
            if first_index > 0 {
                let empty_header = docx_rs::Header::new();
                let empty_footer = docx_rs::Footer::new();
                docx = docx.first_header(empty_header).first_footer(empty_footer);
                println!("【create_document】已设置空的first_header和first_footer，前{}个section不显示页眉页脚", first_index);
            }
        } else {
            println!("【create_document】未找到有页眉页脚的section");
        }

        // 应用页面属性（从第一个section获取）
        if let Some(first_section) = self.options.sections.first() {
            if let Some(page_properties) = &first_section.properties.page {
                // 尝试使用docx-rs的页面设置方法

                // 设置页面大小
                if let Some(page_size) = &page_properties.size {
                    // 尝试使用page_size方法
                    docx = docx.page_size(page_size.width as u32, page_size.height as u32);
                    println!(
                        "【create_document】已应用页面大小: {}x{} twip",
                        page_size.width, page_size.height
                    );
                }

                // 设置页面边距
                if let Some(page_margin) = &page_properties.margin {
                    // 尝试使用page_margin方法
                    docx = docx.page_margin(
                        docx_rs::PageMargin::new()
                            .top(page_margin.top)
                            .right(page_margin.right)
                            .bottom(page_margin.bottom)
                            .left(page_margin.left)
                            .header(page_margin.header)
                            .footer(page_margin.footer),
                    );
                    println!("【create_document】已应用页面边距: top={}, right={}, bottom={}, left={}, header={}, footer={}",
                        page_margin.top, page_margin.right, page_margin.bottom,
                        page_margin.left, page_margin.header, page_margin.footer);
                }

                println!("【create_document】页面属性已应用到文档");
            }
        }

        // 添加节
        for (section_index, section) in self.options.sections.iter().enumerate() {
            // 检查是否是包含页码的section（正文页section）
            let is_main_section = if let Some(main_section) = main_section_with_page_numbers {
                std::ptr::eq(section, main_section)
            } else {
                false
            };

            // 检查当前section是否有页眉页脚配置
            let has_headers = section.headers.is_some();
            let has_footers = section.footers.is_some();

            println!("【create_document】section #{}: has_headers={}, has_footers={}, is_main_section={}",
                section_index, has_headers, has_footers, is_main_section);

            // 如果是包含页码的section，设置页码重新开始计数
            if is_main_section && sections_before_main > 0 {
                // 设置页码从0开始，这样：
                // - 标题页：页码0（隐藏）
                // - 序言页：页码1（隐藏）
                // - 正文页：页码2（显示，但需要在显示时减去前置页数来显示为"第1页"）
                let page_num_type = docx_rs::PageNumType::new().start(0);
                docx = docx.page_num_type(page_num_type);
                println!("【create_document】已设置页码从0开始重新计数（section #{}），前{}个section的页码将被隐藏", section_index, sections_before_main);
            }

            // 如果不是第一个 section，在开始前添加分页符
            if section_index > 0 {
                docx = docx.add_paragraph(
                    docx_rs::Paragraph::new()
                        .add_run(docx_rs::Run::new().add_break(docx_rs::BreakType::Page)),
                );
                println!(
                    "【create_document】在 section #{} 前添加分页符",
                    section_index
                );
            }

            println!(
                "【create_document】处理 section #{}, 包含 {} 个子元素",
                section_index,
                section.children.len()
            );

            // 检查是否需要跳过最后的分页符或空段落
            let children_to_process = if !section.children.is_empty() {
                let mut end_index = section.children.len();

                // 从后往前检查，跳过所有的分页符和空段落
                while end_index > 0 {
                    let current_index = end_index - 1;
                    let should_skip = match &section.children[current_index] {
                        SectionChild::PageBreak => {
                            println!(
                                "【create_document】section #{} 的元素 #{} 是分页符，将被跳过",
                                section_index, current_index
                            );
                            true
                        }
                        SectionChild::Paragraph(paragraph) => {
                            // 检查段落是否为空（没有运行或只有空的运行）
                            let is_empty = paragraph.runs.is_empty()
                                || paragraph.runs.iter().all(|run| {
                                    match run {
                                        RunType::Text(text_run) => {
                                            // 检查文本是否为空、只包含空白字符、或只包含换行符
                                            let text = &text_run.text;
                                            let is_text_empty = text.is_empty()
                                                || text.trim().is_empty()
                                                || text.chars().all(|c| c.is_whitespace());

                                            // 如果文本为空，还要检查是否只是因为 break_before 设置
                                            let is_only_break = text.is_empty()
                                                && (text_run.break_before
                                                    || text_run
                                                        .props
                                                        .break_before
                                                        .unwrap_or(false));

                                            is_text_empty || is_only_break
                                        }
                                        RunType::Break(_) => true, // 分页符运行也算空
                                        _ => false,
                                    }
                                });

                            if is_empty {
                                println!(
                                    "【create_document】section #{} 的元素 #{} 是空段落，将被跳过",
                                    section_index, current_index
                                );
                                // 打印段落内容以便调试
                                for (run_index, run) in paragraph.runs.iter().enumerate() {
                                    match run {
                                        RunType::Text(text_run) => {
                                            println!("【create_document】  运行 #{}: 文本='{}', break_before={}, props.break_before={:?}",
                                                run_index, text_run.text.replace('\n', "\\n"), text_run.break_before, text_run.props.break_before);
                                        }
                                        RunType::Break(_) => {
                                            println!(
                                                "【create_document】  运行 #{}: 分页符",
                                                run_index
                                            );
                                        }
                                        _ => {
                                            println!(
                                                "【create_document】  运行 #{}: 其他类型",
                                                run_index
                                            );
                                        }
                                    }
                                }
                            }
                            is_empty
                        }
                        _ => false,
                    };

                    if should_skip {
                        end_index -= 1;
                    } else {
                        break;
                    }
                }

                &section.children[..end_index]
            } else {
                &section.children[..]
            };

            println!(
                "【create_document】section #{} 实际处理 {} 个子元素",
                section_index,
                children_to_process.len()
            );

            // 添加节的子元素
            for child in children_to_process {
                match child {
                    SectionChild::Paragraph(paragraph) => {
                        docx = docx.add_paragraph(paragraph.to_docx_paragraph(
                            self.options.styles.clone(),
                            self.options.footnotes.clone(),
                        ));
                    }
                    SectionChild::Table(table) => {
                        docx = docx.add_table(table.to_docx_table(
                            self.options.styles.clone(),
                            self.options.footnotes.clone(),
                        ));
                    }
                    SectionChild::PageBreak => {
                        docx = docx.add_paragraph(
                            docx_rs::Paragraph::new()
                                .add_run(docx_rs::Run::new().add_break(docx_rs::BreakType::Page)),
                        );
                    }
                }
            }
        }

        // 添加脚注 - 使用docx-rs的真正脚注功能
        // 注意：脚注引用已经在段落中通过FootnoteReferenceRun添加了，这里不需要再处理引用
        println!("【create_document】脚注处理完成，文档中包含的脚注引用将自动关联到相应的脚注内容");

        docx
    }

    /// 保存文档
    pub fn save(&self, filepath: &str) -> DocxAdapterResult<()> {
        let file = File::create(Path::new(filepath))?;
        match self.docx.clone().build().pack(file) {
            Ok(_) => Ok(()),
            Err(e) => {
                if let Some(zip_err) = e
                    .source()
                    .and_then(|s| s.downcast_ref::<zip::result::ZipError>())
                {
                    Err(DocxAdapterError::InvalidConfig(format!(
                        "ZIP error: {:?}",
                        zip_err
                    )))
                } else {
                    Err(DocxAdapterError::DocxError(docx_rs::DocxError::ZipError(e)))
                }
            }
        }
    }

    /// 获取 Base64 编码的文档
    pub fn to_base64(&self) -> DocxAdapterResult<String> {
        // 保存到临时文件
        let temp_path = std::env::temp_dir().join("temp_docx.docx");
        let file = File::create(&temp_path)?;
        match self.docx.clone().build().pack(file) {
            Ok(_) => {}
            Err(e) => {
                if let Some(zip_err) = e
                    .source()
                    .and_then(|s| s.downcast_ref::<zip::result::ZipError>())
                {
                    return Err(DocxAdapterError::InvalidConfig(format!(
                        "ZIP error: {:?}",
                        zip_err
                    )));
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
}

/// 段落
#[derive(Debug, Clone)]
pub struct Paragraph {
    pub runs: Vec<RunType>,
    pub alignment: Option<AlignmentType>,
    pub indent: Option<ParagraphIndent>,
    pub style: Option<String>,
    pub spacing: Option<ParagraphSpacing>,
    pub frame: Option<ParagraphFrame>,
    pub outline_level: Option<usize>, // 添加 outline 层级支持
}

impl Paragraph {
    /// 创建新的段落
    pub fn new() -> Self {
        Self {
            runs: Vec::new(),
            alignment: None,
            indent: None,
            style: None,
            spacing: None,
            frame: None,
            outline_level: None,
        }
    }
    pub fn new_with_spacing(spacing: ParagraphSpacing) -> Self {
        Self {
            runs: Vec::new(),
            alignment: None,
            indent: None,
            style: None,
            spacing: Some(spacing),
            frame: None,
            outline_level: None,
        }
    }

    /// 添加文本运行
    pub fn add_text_run(&mut self, run: TextRun) -> &mut Self {
        self.runs.push(RunType::Text(run));
        self
    }

    /// 添加分页符运行
    pub fn add_break_run(&mut self, run: BreakRun) -> &mut Self {
        self.runs.push(RunType::Break(run));
        self
    }

    /// 添加超链接运行
    pub fn add_hyperlink_run(&mut self, run: HyperlinkRun) -> &mut Self {
        self.runs.push(RunType::Hyperlink(run));
        self
    }

    /// 添加页码运行
    pub fn add_page_number_run(&mut self, run: PageNumberRun) -> &mut Self {
        self.runs.push(RunType::PageNumber(run));
        self
    }

    /// 添加运行（通用方法）
    pub fn add_run(&mut self, run: RunType) -> &mut Self {
        self.runs.push(run);
        self
    }
    pub fn add_run_before(&mut self, run: RunType) -> &mut Self {
        self.runs.insert(0, run); // 在开头插入
        self
    }

    /// 设置对齐方式
    pub fn align(&mut self, alignment: AlignmentType) -> &mut Self {
        self.alignment = Some(alignment);
        self
    }

    /// 设置缩进（只设置左缩进）
    pub fn indent(&mut self, left_indent: i32) -> &mut Self {
        if let Some(ref mut indent) = self.indent {
            indent.left = Some(left_indent);
        } else {
            self.indent = Some(ParagraphIndent {
                left: Some(left_indent),
                right: None,
            });
        }
        self
    }

    /// 设置右缩进
    pub fn indent_right(&mut self, right_indent: i32) -> &mut Self {
        if let Some(ref mut indent) = self.indent {
            indent.right = Some(right_indent);
        } else {
            self.indent = Some(ParagraphIndent {
                left: None,
                right: Some(right_indent),
            });
        }
        self
    }

    /// 设置完整的缩进信息
    pub fn indent_full(&mut self, indent: ParagraphIndent) -> &mut Self {
        self.indent = Some(indent);
        self
    }

    /// 设置样式
    pub fn style(&mut self, style: &str) -> &mut Self {
        self.style = Some(style.to_string());
        self
    }

    /// 设置段落间距
    pub fn spacing(&mut self, spacing: ParagraphSpacing) -> &mut Self {
        self.spacing = Some(spacing);
        self
    }

    /// 设置行距（便捷方法）
    pub fn line_spacing(&mut self, spacing: ParagraphSpacing) -> &mut Self {
        self.spacing = Some(spacing);
        self
    }

    /// 设置单倍行距
    pub fn single_line_spacing(&mut self) -> &mut Self {
        self.spacing = Some(
            ParagraphSpacing::new()
                .line_rule(super::LineRuleType::Auto)
                .line(240), // 240 twips = 12pt = 单倍行距
        );
        self
    }

    /// 设置1.5倍行距
    pub fn one_and_half_line_spacing(&mut self) -> &mut Self {
        self.spacing = Some(
            ParagraphSpacing::new()
                .line_rule(super::LineRuleType::Auto)
                .line(360), // 360 twips = 18pt = 1.5倍行距
        );
        self
    }

    /// 设置双倍行距
    pub fn double_line_spacing(&mut self) -> &mut Self {
        self.spacing = Some(
            ParagraphSpacing::new()
                .line_rule(super::LineRuleType::Auto)
                .line(480), // 480 twips = 24pt = 双倍行距
        );
        self
    }

    /// 设置固定行距
    pub fn exact_line_spacing(&mut self, twips: i32) -> &mut Self {
        self.spacing = Some(
            ParagraphSpacing::new()
                .line_rule(super::LineRuleType::Exact)
                .line(twips),
        );
        self
    }

    /// 设置最小行距
    pub fn at_least_line_spacing(&mut self, twips: i32) -> &mut Self {
        self.spacing = Some(
            ParagraphSpacing::new()
                .line_rule(super::LineRuleType::AtLeast)
                .line(twips),
        );
        self
    }

    /// 设置段前间距
    pub fn space_before(&mut self, twips: i32) -> &mut Self {
        if let Some(ref mut spacing) = self.spacing {
            *spacing = spacing.clone().before(twips);
        } else {
            self.spacing = Some(ParagraphSpacing::new().before(twips));
        }
        self
    }

    /// 设置段后间距
    pub fn space_after(&mut self, twips: i32) -> &mut Self {
        if let Some(ref mut spacing) = self.spacing {
            *spacing = spacing.clone().after(twips);
        } else {
            self.spacing = Some(ParagraphSpacing::new().after(twips));
        }
        self
    }

    /// 设置框架
    pub fn frame(&mut self, frame: ParagraphFrame) -> &mut Self {
        self.frame = Some(frame);
        self
    }

    /// 设置大纲层级
    pub fn outline_level(&mut self, level: usize) -> &mut Self {
        self.outline_level = Some(level);
        self
    }

    /// 转换为 docx-rs 的 Paragraph
    pub fn to_docx_paragraph(
        &self,
        mstyles: Option<Styles>,
        footnotes: HashMap<usize, Footnote>,
    ) -> docx_rs::Paragraph {
        let mut paragraph = docx_rs::Paragraph::new();

        let mut left_i: Option<i32> = None;
        let mut right_i: Option<i32> = None;

        if let Some(style) = &self.style {
            if let Some(ref styles) = mstyles {
                for paragraph_style in &styles.paragraph_styles {
                    if let Some(style_id) = &paragraph_style.id {
                        if style_id == style {
                            // 设置缩进
                            if let Some(indent) = &paragraph_style.indent {
                                left_i = indent.left;
                                right_i = indent.right;
                            }
                            if let Some(sp) = &paragraph_style.spacing {
                                paragraph = paragraph.line_spacing(sp.to_docx_line_spacing());
                            }
                        }
                    }
                }
            }
        }

        if let Some(alignment) = &self.alignment {
            paragraph = paragraph.align(alignment.to_docx_alignment());
        }

        if let Some(indent) = &self.indent {
            if let Some(left) = indent.left {
                left_i = Some(left);
            }
            if let Some(right) = indent.right {
                right_i = Some(right);
            }
        }
        paragraph = paragraph.indent(
            left_i,  // 使用 or 方法替代 || 运算符
            None,    // 首行缩进
            right_i, // 使用 or 方法替代缺失的 rightI
            None,
        );

        if let Some(frame) = &self.frame {
            // 应用框架属性
            if let Some(width) = frame.width {
                // 设置框架宽度
                // docx-rs 的 size 方法只接受一个参数，表示宽度
                paragraph = paragraph.size(width as usize);
            }

            // 设置水平位置
            if let Some(x_align) = &frame.x_align {
                match x_align {
                    HorizontalPositionAlign::Left => {
                        // 左对齐，使用 0 表示左边缘
                        paragraph = paragraph.frame_x(0);
                    }
                    HorizontalPositionAlign::Center => {
                        // 居中对齐，使用 0 表示居中（依赖于段落的居中对齐）
                        paragraph = paragraph.frame_x(0);
                    }
                    HorizontalPositionAlign::Right => {
                        // 右对齐，使用 0 表示右边缘（依赖于段落的右对齐）
                        paragraph = paragraph.frame_x(0);
                    }
                    _ => {}
                }
            }

            // 设置垂直位置
            if let Some(y_align) = &frame.y_align {
                match y_align {
                    VerticalPositionAlign::Top => {
                        // 顶部对齐，使用 0 表示顶部
                        paragraph = paragraph.frame_y(0);
                    }
                    VerticalPositionAlign::Center => {
                        // 居中对齐，使用 5000 表示居中（docx 中的标准值）
                        paragraph = paragraph.frame_y(5000);
                    }
                    VerticalPositionAlign::Bottom => {
                        // 底部对齐
                        // 尝试使用更大的值来实现底部对齐
                        // 在 docx-rs 中，frame_y 的值范围是 0-9999
                        // 但实际上 Word 可能会接受更大的值

                        // 如果页面高度和底部边距信息可用，尝试使用它们来计算
                        // 一个更精确的底部位置
                        if let (Some(page_height), Some(bottom_margin), Some(line_height)) =
                            (frame.page_height, frame.bottom_margin, frame.line_height)
                        {
                            // 动态计算段落的行数
                            let mut lines = 1; // 至少有一行

                            // 遍历所有文本运行，计算换行符数量
                            for run in &self.runs {
                                if let RunType::Text(text_run) = run {
                                    // 计算文本中的换行符数量
                                    let line_breaks = text_run.text.matches('\n').count();
                                    lines += line_breaks;

                                    // 如果设置了 break_before，则增加一行
                                    if text_run.break_before {
                                        lines += 1;
                                    }

                                    // 如果 RunProps 中设置了 break_before，则增加一行
                                    if let Some(true) = text_run.props.break_before {
                                        lines += 1;
                                    }

                                    // 如果设置了 break_type，则增加一行
                                    if text_run.props.break_type.is_some() {
                                        lines += 1;
                                    }
                                }
                            }

                            // 确保至少有一行
                            lines = lines.max(1);

                            let bottom_pos =
                                page_height - bottom_margin - (line_height * lines as i32);

                            // 使用一个比例因子，将 bottom_pos 转换为更大的值
                            // 这个因子可以根据实际效果进行调整
                            let scale_factor = 0.80;
                            let adjusted_pos = (bottom_pos as f32 * scale_factor) as i32;

                            paragraph = paragraph.frame_y(adjusted_pos);
                        } else {
                            // 如果没有提供页面信息，则使用一个非常大的值
                            // 这个值可以根据实际效果进行调整
                            paragraph = paragraph.frame_y(20000);
                        }
                    }
                    _ => {}
                }
            }

            // docx-rs 没有 frame_wrap 方法，暂时不设置 wrap 属性
        }

        // 设置大纲层级
        if let Some(level) = self.outline_level {
            paragraph = paragraph.outline_lvl(level);
        }

        for run in &self.runs {
            paragraph = paragraph.add_run(run.to_docx_run(mstyles.clone(), footnotes.clone()));
        }

        if let Some(spacing) = &self.spacing {
            // 应用行距和段落间距设置
            paragraph = paragraph.line_spacing(spacing.to_docx_line_spacing());
        }

        paragraph
    }
}

/// 运行特性
pub trait RunTrait {
    fn to_docx_run(
        &self,
        mstyles: Option<Styles>,
        footnotes: HashMap<usize, Footnote>,
    ) -> docx_rs::Run;
}

/// 运行类型枚举
#[derive(Debug, Clone)]
pub enum RunType {
    Text(TextRun),
    Break(BreakRun),
    Hyperlink(HyperlinkRun),
    PageNumber(PageNumberRun),
}

impl RunTrait for RunType {
    fn to_docx_run(
        &self,
        mstyles: Option<Styles>,
        footnotes: HashMap<usize, Footnote>,
    ) -> docx_rs::Run {
        match self {
            RunType::Text(run) => run.to_docx_run(mstyles, footnotes),
            RunType::Break(run) => run.to_docx_run(mstyles, footnotes),
            RunType::Hyperlink(run) => run.to_docx_run(mstyles, footnotes),
            RunType::PageNumber(run) => run.to_docx_run(mstyles, footnotes),
        }
    }
}

/// 文本运行
#[derive(Debug, Clone)]
pub struct TextRun {
    pub text: String,
    pub props: RunProps,
    pub break_before: bool,
    pub children: Vec<RunType>,
    pub footnote_id: Option<usize>, // 脚注ID，如果是脚注引用则设置此值
    pub footnote_content: Option<Vec<Paragraph>>, // 脚注内容（已格式化的运行）
}

impl TextRun {
    /// 创建新的文本运行
    pub fn new(text: &str) -> Self {
        Self {
            text: text.to_string(),
            props: RunProps::default(),
            break_before: false,
            children: Vec::new(),
            footnote_id: None,
            footnote_content: None,
        }
    }

    /// 使用指定的属性创建新的文本运行
    pub fn with_props(text: &str, props: RunProps) -> Self {
        Self {
            text: text.to_string(),
            props,
            break_before: false,
            children: Vec::new(),
            footnote_id: None,
            footnote_content: None,
        }
    }

    /// 创建脚注引用运行
    pub fn footnote_reference(footnote_id: usize, footnote_content: Vec<Paragraph>, props: RunProps) -> Self {
        Self {
            text: format!("[{}]", footnote_id), // 显示脚注编号
            props: props,
            break_before: false,
            children: Vec::new(),
            footnote_id: Some(footnote_id),
            footnote_content: Some(footnote_content),
        }
    }

    /// 设置在文本前添加换行
    pub fn break_before(mut self, break_before: bool) -> Self {
        self.break_before = break_before;
        self
    }

    /// 设置字体大小
    pub fn size(mut self, size: usize) -> Self {
        self.props.size = Some(size);
        self
    }

    /// 设置字体
    pub fn font(mut self, font: &str) -> Self {
        self.props.font = Some(font.to_string());
        self
    }

    /// 设置粗体
    pub fn bold(mut self) -> Self {
        self.props.bold = Some(true);
        self
    }

    /// 设置斜体
    pub fn italic(mut self) -> Self {
        self.props.italic = Some(true);
        self
    }

    /// 设置下划线
    pub fn underline(mut self, underline: UnderlineType) -> Self {
        self.props.underline = Some(underline);
        self
    }

    /// 设置颜色
    pub fn color(mut self, color: &str) -> Self {
        self.props.color = Some(color.to_string());
        self
    }

    /// 设置上标
    pub fn superscript(mut self) -> Self {
        self.props.superscript = Some(true);
        self
    }

    /// 设置下标
    pub fn subscript(mut self) -> Self {
        self.props.subscript = Some(true);
        self
    }

    /// 添加文本子运行
    pub fn add_text_child(&mut self, run: TextRun) -> &mut Self {
        self.children.push(RunType::Text(run));
        self
    }

    /// 添加分页符子运行
    pub fn add_break_child(&mut self, run: BreakRun) -> &mut Self {
        self.children.push(RunType::Break(run));
        self
    }

    /// 添加超链接子运行
    pub fn add_hyperlink_child(&mut self, run: HyperlinkRun) -> &mut Self {
        self.children.push(RunType::Hyperlink(run));
        self
    }

    /// 添加子运行
    pub fn add_child(&mut self, run: RunType) -> &mut Self {
        self.children.push(run);
        self
    }
}

impl RunTrait for TextRun {
    fn to_docx_run(
        &self,
        mstyles: Option<Styles>,
        footnotes: HashMap<usize, Footnote>,
    ) -> docx_rs::Run {
        // 如果这是一个脚注引用，创建真正的脚注引用
        if let (Some(footnote_id), Some(footnote_content)) =
            (&self.footnote_id, &self.footnote_content)
        {
            println!(
                "【to_docx_run】创建脚注引用，ID: {}, 内容段落数: {}, 收集到脚注：{}",
                footnote_id,
                footnote_content.len(),
                footnotes.len()
            );
            let mut footnote = docx_rs::Footnote::new();

            // 添加脚注内容 - 使用已经格式化好的TextRun
            if !footnote_content.is_empty() {
                // 创建脚注段落，包含所有格式化好的运行
                for (i, pg) in footnote_content.iter().enumerate() {
                    // 使用TextRun的to_docx_run方法，保持原有的格式
                    if i == 0 {
                        // 克隆并插入脚注编号
                        let mut pg_with_no = pg.clone();
                        pg_with_no.add_run_before(RunType::Text(TextRun::with_props(
                            format!("[{}] ", footnote_id).as_str(), self.props.clone().color("#000000")
                        )));
                        footnote = footnote.add_content(
                            pg_with_no.to_docx_paragraph(mstyles.clone(), footnotes.clone()),
                        );
                    } else {
                        footnote = footnote
                            .add_content(pg.to_docx_paragraph(mstyles.clone(), footnotes.clone()));
                    }
                }
            } else {
                // 如果没有内容，添加默认内容（不设置样式，由外层控制）
                // 从footnotes找到与 footnote_id 匹配的脚注
                if let Some(fnote) = footnotes.get(footnote_id) {
                    println!(
                        "【to_docx_run】基q全局收集创建脚注引用，ID: {}, 内容段落数: {}",
                        footnote_id,
                        fnote.children.len()
                    );
                    // 使用脚注的段落创建默认内容
                    for (i, pg) in fnote.children.iter().enumerate() {
                        // 使用TextRun的to_docx_run方法，保持原有的格式
                        if i == 0 {
                            // 克隆并插入脚注编号
                            let mut pg_with_no = pg.clone();
                            pg_with_no.add_run_before(RunType::Text(TextRun::with_props(
                                format!("[{}] ", footnote_id).as_str(), self.props.clone().color("#000000")
                            )));
                            footnote = footnote.add_content(
                                pg_with_no.to_docx_paragraph(mstyles.clone(), footnotes.clone()),
                            );
                        } else {
                            footnote = footnote.add_content(
                                pg.to_docx_paragraph(mstyles.clone(), footnotes.clone()),
                            );
                        }
                    }
                }
            }

            // 创建带有脚注引用的运行
            return docx_rs::Run::new().add_footnote_reference_with_size(
                footnote,
                (self.props.size.unwrap_or(9) as f64 * 2.0 * 1.9) as usize, // 适配，字体磅数都要*2 ，比如实际的12磅实际传入参数需要24
            );
        }

        // 普通文本运行
        let mut run = docx_rs::Run::new();

        // 如果需要在文本前添加换行符（通过 TextRun 的 break_before 属性）
        if self.break_before {
            run = run.add_break(crate::docx::adapter::BreakType::TextWrapping.to_docx_break());
        }

        // 如果需要在文本前添加换行符（通过 RunProps 的 break_before 属性）
        if let Some(true) = self.props.break_before {
            run = run.add_break(crate::docx::adapter::BreakType::TextWrapping.to_docx_break());
        }

        // 添加文本
        run = run.add_text(&self.text);

        // 如果需要在文本后添加换行符
        if self.props.break_type.is_some() {
            let break_type = self
                .props
                .break_type
                .unwrap_or(crate::docx::adapter::BreakType::TextWrapping);
            run = run.add_break(break_type.to_docx_break());
        }

        if let Some(size) = self.props.size {
            run = run.size(size*2);// 适配，字体磅数都要*2 ，比如实际的12磅实际传入参数需要24
        }

        if let Some(font) = &self.props.font {
            let run_fonts = docx_rs::RunFonts::new()
                .east_asia(font)
                .ascii(font)
                .hi_ansi(font);
            run = run.fonts(run_fonts);
        }

        if let Some(true) = self.props.bold {
            run = run.bold();
        }

        if let Some(true) = self.props.italic {
            run = run.italic();
        }

        if let Some(underline) = &self.props.underline {
            match underline {
                UnderlineType::Single => run = run.underline("single"),
                UnderlineType::Double => run = run.underline("double"),
                UnderlineType::Dash => run = run.underline("dash"),
                UnderlineType::Dotted => run = run.underline("dotted"),
                UnderlineType::Wave => run = run.underline("wave"),
                UnderlineType::None => {}
            }
        }

        if let Some(color) = &self.props.color {
            run = run.color(color);
        }

        if let Some(_spacing) = self.props.character_spacing {
            // docx-rs 0.4.18-rc44 分支可能不支持 spacing 方法
            // 暂时跳过字符间距处理
            // run = run.spacing(spacing);
        }

        if let Some(true) = self.props.superscript {
            // docx-rs 支持上标
            run = run.vanish(); // 暂时使用 vanish 代替，实际应该使用 vertAlign
        }

        if let Some(true) = self.props.subscript {
            // docx-rs 支持下标
            run = run.vanish(); // 暂时使用 vanish 代替，实际应该使用 vertAlign
        }

        // 处理子运行
        // 注意：docx-rs 不直接支持子运行，这里只是一个占位符
        // 实际实现需要根据 docx-rs 的 API 进行调整
        // 在 docx-rs 中，可能需要将子运行添加到同一个段落中，而不是添加到当前运行中

        run
    }
}

/// 分页符运行
#[derive(Debug, Clone)]
pub struct BreakRun {
    pub break_type: BreakType,
}

impl BreakRun {
    /// 创建新的分页符运行
    pub fn new(break_type: BreakType) -> Self {
        Self { break_type }
    }
}

impl RunTrait for BreakRun {
    fn to_docx_run(
        &self,
        _mstyles: Option<Styles>,
        _footnotes: HashMap<usize, Footnote>,
    ) -> docx_rs::Run {
        docx_rs::Run::new().add_break(self.break_type.to_docx_break())
    }
}

/// 超链接运行
#[derive(Debug, Clone)]
pub struct HyperlinkRun {
    pub text: String,
    pub url: String,
    pub props: RunProps,
}

impl HyperlinkRun {
    /// 创建新的超链接运行
    pub fn new(text: &str, url: &str) -> Self {
        Self {
            text: text.to_string(),
            url: url.to_string(),
            props: RunProps::default(),
        }
    }

    /// 使用指定的属性创建新的超链接运行
    pub fn with_props(text: &str, url: &str, props: RunProps) -> Self {
        Self {
            text: text.to_string(),
            url: url.to_string(),
            props,
        }
    }

    /// 设置字体大小
    pub fn size(mut self, size: usize) -> Self {
        self.props.size = Some(size);
        self
    }

    /// 设置字体
    pub fn font(mut self, font: &str) -> Self {
        self.props.font = Some(font.to_string());
        self
    }

    /// 设置粗体
    pub fn bold(mut self, bold: bool) -> Self {
        self.props.bold = Some(bold);
        self
    }

    /// 设置斜体
    pub fn italic(mut self, italic: bool) -> Self {
        self.props.italic = Some(italic);
        self
    }

    /// 设置下划线
    pub fn underline(mut self, underline: UnderlineType) -> Self {
        self.props.underline = Some(underline);
        self
    }

    /// 设置颜色
    pub fn color(mut self, color: &str) -> Self {
        self.props.color = Some(color.to_string());
        self
    }
}

impl RunTrait for HyperlinkRun {
    fn to_docx_run(
        &self,
        _mstyles: Option<Styles>,
        _footnotes: HashMap<usize, Footnote>,
    ) -> docx_rs::Run {
        // 创建一个带有超链接样式的文本运行
        let mut run = docx_rs::Run::new();

        // 添加文本
        run = run.add_text(&self.text);

        // 添加样式
        if let Some(size) = self.props.size {
            run = run.size(size * 2); // 适配，字体磅数都要*2 ，比如实际的12磅实际传入参数需要24
        }

        if let Some(font) = &self.props.font {
            let run_fonts = docx_rs::RunFonts::new()
                .east_asia(font)
                .ascii(font)
                .hi_ansi(font);
            run = run.fonts(run_fonts);
        }

        if let Some(true) = self.props.bold {
            run = run.bold();
        }

        if let Some(true) = self.props.italic {
            run = run.italic();
        }

        if let Some(underline) = &self.props.underline {
            match underline {
                UnderlineType::Single => run = run.underline("single"),
                UnderlineType::Double => run = run.underline("double"),
                UnderlineType::Dash => run = run.underline("dash"),
                UnderlineType::Dotted => run = run.underline("dotted"),
                UnderlineType::Wave => run = run.underline("wave"),
                UnderlineType::None => {}
            }
        }

        // 注意：不在适配器层设置默认的超链接样式
        // 超链接的颜色和下划线应该由外层通过props传入

        run
    }
}

/// 页码运行
#[derive(Debug, Clone)]
pub struct PageNumberRun {
    pub children: Vec<PageNumberChild>,
    pub props: RunProps,
    pub main_content_start_page: Option<usize>,
}

/// 页码子元素
#[derive(Debug, Clone)]
pub enum PageNumberChild {
    Text(String),
    PageNumber,
}

impl PageNumberRun {
    /// 创建新的页码运行
    pub fn new() -> Self {
        Self {
            children: Vec::new(),
            props: RunProps::default(),
            main_content_start_page: None,
        }
    }

    /// 添加文本
    pub fn add_text(&mut self, text: &str) -> &mut Self {
        self.children.push(PageNumberChild::Text(text.to_string()));
        self
    }

    /// 添加页码
    pub fn add_page_number(&mut self) -> &mut Self {
        self.children.push(PageNumberChild::PageNumber);
        self
    }

    /// 添加条件页码（只在正文页显示，并从1开始计数）
    pub fn add_conditional_page_number(&mut self, main_content_start_page: usize) -> &mut Self {
        self.main_content_start_page = Some(main_content_start_page);
        self.children.push(PageNumberChild::PageNumber);
        self
    }

    /// 设置字体大小
    pub fn size(mut self, size: usize) -> Self {
        self.props.size = Some(size);
        self
    }

    /// 设置字体
    pub fn font(mut self, font: &str) -> Self {
        self.props.font = Some(font.to_string());
        self
    }

    /// 设置粗体
    pub fn bold(mut self) -> Self {
        self.props.bold = Some(true);
        self
    }

    /// 设置斜体
    pub fn italic(mut self) -> Self {
        self.props.italic = Some(true);
        self
    }

    /// 设置颜色
    pub fn color(mut self, color: &str) -> Self {
        self.props.color = Some(color.to_string());
        self
    }
}

impl RunTrait for PageNumberRun {
    fn to_docx_run(
        &self,
        _mstyles: Option<Styles>,
        _footnotes: HashMap<usize, Footnote>,
    ) -> docx_rs::Run {
        let mut run = docx_rs::Run::new();

        // 处理子元素
        for child in &self.children {
            match child {
                PageNumberChild::Text(text) => {
                    run = run.add_text(text);
                }
                PageNumberChild::PageNumber => {
                    // 使用简单的PAGE字段
                    // 页码的显示控制通过first_header/first_footer来实现
                    run = run
                        .add_field_char(docx_rs::FieldCharType::Begin, false)
                        .add_instr_text(docx_rs::InstrText::PAGE(docx_rs::InstrPAGE::new()))
                        .add_field_char(docx_rs::FieldCharType::End, false);
                }
            }
        }

        // 应用样式
        if let Some(size) = self.props.size {
            run = run.size(size * 2); // 适配，字体磅数都要*2 ，比如实际的12磅实际传入参数需要24
        }

        if let Some(font) = &self.props.font {
            let run_fonts = docx_rs::RunFonts::new()
                .east_asia(font)
                .ascii(font)
                .hi_ansi(font);
            run = run.fonts(run_fonts);
        }

        if let Some(true) = self.props.bold {
            run = run.bold();
        }

        if let Some(true) = self.props.italic {
            run = run.italic();
        }

        if let Some(color) = &self.props.color {
            run = run.color(color);
        }

        run
    }
}

/// 表格宽度
#[derive(Debug, Clone)]
pub struct TableWidth {
    pub size: i32,
    pub width_type: WidthType,
}

impl TableWidth {
    /// 创建新的表格宽度
    pub fn new(size: i32, width_type: WidthType) -> Self {
        Self { size, width_type }
    }

    /// 转换为 docx-rs 的 TableWidth
    pub fn to_docx_width_type(&self) -> docx_rs::TableWidth {
        match self.width_type {
            WidthType::Auto => docx_rs::TableWidth::new(0, docx_rs::WidthType::Auto),
            WidthType::DXA => docx_rs::TableWidth::new(self.size as usize, docx_rs::WidthType::Dxa),
            WidthType::Percentage => {
                docx_rs::TableWidth::new(self.size as usize, docx_rs::WidthType::Pct)
            }
        }
    }
}

/// 表格缩进
#[derive(Debug, Clone)]
pub struct TableIndent {
    pub size: i32,
    pub width_type: WidthType,
}

impl TableIndent {
    /// 创建新的表格缩进
    pub fn new(size: i32, width_type: WidthType) -> Self {
        Self { size, width_type }
    }
}

/// 表格边框
#[derive(Debug, Clone)]
pub struct TableBorders {
    pub top: Option<TableBorder>,
    pub bottom: Option<TableBorder>,
    pub left: Option<TableBorder>,
    pub right: Option<TableBorder>,
    pub inside_h: Option<TableBorder>,
    pub inside_v: Option<TableBorder>,
}

impl TableBorders {
    /// 创建新的表格边框
    pub fn new() -> Self {
        Self {
            top: None,
            bottom: None,
            left: None,
            right: None,
            inside_h: None,
            inside_v: None,
        }
    }
}

/// 表格边框
#[derive(Debug, Clone)]
pub struct TableBorder {
    pub size: i32,
    pub color: String,
    pub style: String,
}

impl TableBorder {
    /// 创建新的表格边框
    pub fn new(size: i32, color: &str, style: &str) -> Self {
        Self {
            size,
            color: color.to_string(),
            style: style.to_string(),
        }
    }
}

/// 表格配置
#[derive(Debug, Clone)]
pub struct TableConfig {
    pub rows: Option<Vec<TableRow>>,
    pub indent: Option<TableIndent>,
    pub borders: Option<TableBorders>,
}

impl TableConfig {
    /// 创建新的表格配置
    pub fn new() -> Self {
        Self {
            rows: None,
            indent: None,
            borders: None,
        }
    }
}

/// 表格
#[derive(Debug, Clone)]
pub struct Table {
    pub rows: Vec<TableRow>,
    pub indent: Option<TableIndent>,
    pub borders: Option<TableBorders>,
}

impl Table {
    /// 创建新的表格
    pub fn new() -> Self {
        Self {
            rows: Vec::new(),
            indent: None,
            borders: None,
        }
    }

    /// 使用配置创建新的表格
    pub fn with_config(config: TableConfig) -> Self {
        let mut table = Self::new();

        if let Some(indent) = config.indent {
            table.indent = Some(indent);
        }

        if let Some(borders) = config.borders {
            table.borders = Some(borders);
        }

        if let Some(rows) = config.rows {
            for row in rows {
                table.rows.push(row);
            }
        }

        table
    }

    /// 添加行
    pub fn add_row(&mut self, row: TableRow) -> &mut Self {
        self.rows.push(row);
        self
    }

    /// 设置缩进
    pub fn indent(&mut self, indent: TableIndent) -> &mut Self {
        self.indent = Some(indent);
        self
    }

    /// 设置边框
    pub fn borders(&mut self, borders: TableBorders) -> &mut Self {
        self.borders = Some(borders);
        self
    }

    /// 转换为 docx-rs 的 Table
    pub fn to_docx_table(
        &self,
        mstyles: Option<Styles>,
        footnotes: HashMap<usize, Footnote>,
    ) -> docx_rs::Table {
        let mut table = docx_rs::Table::without_borders(Vec::new());

        // 设置缩进
        if let Some(indent) = &self.indent {
            match indent.width_type {
                WidthType::DXA => {
                    table = table.indent(indent.size);
                }
                _ => {}
            }
        }

        // 设置边框 - 暂时跳过边框设置，docx-rs默认无边框
        // 注意：docx-rs的表格默认是无边框的，所以我们不需要显式设置
        // 如果需要设置边框，需要根据具体的docx-rs版本API来实现

        // 添加行
        for row in &self.rows {
            table = table.add_row(row.to_docx_table_row(mstyles.clone(), footnotes.clone()));
        }

        table
    }
}

/// 表格行配置
#[derive(Debug, Clone)]
pub struct TableRowConfig {
    pub children: Option<Vec<TableCell>>,
}

impl TableRowConfig {
    /// 创建新的表格行配置
    pub fn new() -> Self {
        Self { children: None }
    }
}

/// 表格行
#[derive(Debug, Clone)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

impl TableRow {
    /// 创建新的表格行
    pub fn new() -> Self {
        Self { cells: Vec::new() }
    }

    /// 使用配置创建新的表格行
    pub fn with_config(config: TableRowConfig) -> Self {
        let mut row = Self::new();

        if let Some(cells) = config.children {
            for cell in cells {
                row.cells.push(cell);
            }
        }

        row
    }

    /// 添加单元格
    pub fn add_cell(&mut self, cell: TableCell) -> &mut Self {
        self.cells.push(cell);
        self
    }

    /// 转换为 docx-rs 的 TableRow
    pub fn to_docx_table_row(
        &self,
        mstyles: Option<Styles>,
        footnotes: HashMap<usize, Footnote>,
    ) -> docx_rs::TableRow {
        // 创建单元格向量
        let mut cells = Vec::new();

        // 添加单元格
        for cell in &self.cells {
            cells.push(cell.to_docx_table_cell(mstyles.clone(), footnotes.clone()));
        }

        // 创建行
        docx_rs::TableRow::new(cells)
    }
}

/// 表格单元格配置
#[derive(Debug, Clone)]
pub struct TableCellConfig {
    pub width: Option<TableWidth>,
    pub children: Option<Vec<Paragraph>>,
}

impl TableCellConfig {
    /// 创建新的表格单元格配置
    pub fn new() -> Self {
        Self {
            width: None,
            children: None,
        }
    }
}

/// 表格单元格
#[derive(Debug, Clone)]
pub struct TableCell {
    pub width: Option<TableWidth>,
    pub children: Vec<Paragraph>,
}

impl TableCell {
    /// 创建新的表格单元格
    pub fn new() -> Self {
        Self {
            width: None,
            children: Vec::new(),
        }
    }

    /// 使用配置创建新的表格单元格
    pub fn with_config(config: TableCellConfig) -> Self {
        let mut cell = Self::new();

        if let Some(width) = config.width {
            cell.width = Some(width);
        }

        if let Some(children) = config.children {
            for paragraph in children {
                cell.children.push(paragraph);
            }
        }

        cell
    }

    /// 设置宽度
    pub fn width(&mut self, width: TableWidth) -> &mut Self {
        self.width = Some(width);
        self
    }

    /// 添加段落
    pub fn add_paragraph(&mut self, paragraph: Paragraph) -> &mut Self {
        self.children.push(paragraph);
        self
    }

    /// 转换为 docx-rs 的 TableCell
    pub fn to_docx_table_cell(
        &self,
        mstyles: Option<Styles>,
        footnotes: HashMap<usize, Footnote>,
    ) -> docx_rs::TableCell {
        let mut cell = docx_rs::TableCell::new();

        // 设置宽度
        if let Some(width) = &self.width {
            // 注意：docx-rs 0.4.18-rc44 分支的 width 方法需要 usize 和 WidthType 参数
            match width.width_type {
                WidthType::Auto => cell = cell.width(0, docx_rs::WidthType::Auto),
                WidthType::DXA => cell = cell.width(width.size as usize, docx_rs::WidthType::Dxa),
                WidthType::Percentage => {
                    cell = cell.width(width.size as usize, docx_rs::WidthType::Pct)
                }
            }
        }

        // 添加段落
        for paragraph in &self.children {
            cell =
                cell.add_paragraph(paragraph.to_docx_paragraph(mstyles.clone(), footnotes.clone()));
        }

        cell
    }
}
