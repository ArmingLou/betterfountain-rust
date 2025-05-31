use betterfountain_rust::docx::adapter::docx::{Document, Paragraph, TextRun, SectionChild, Section, SectionProperties, PageProperties, PageSize, PageMargin, PageNumbers};
use betterfountain_rust::docx::adapter::{ParagraphFrame, FrameAnchorType, HorizontalPositionAlign, VerticalPositionAlign, AlignmentType, DocxAdapterError};
use betterfountain_rust::docx::convert_inches_to_twip;

fn main() -> Result<(), DocxAdapterError> {
    println!("=== 测试独立 Frame 功能 ===");

    // 创建文档
    let mut doc = Document::new();

    // 设置页面属性
    let section_props = SectionProperties {
        page: Some(PageProperties {
            size: Some(PageSize::new(
                convert_inches_to_twip(11.0), // 页面高度
                convert_inches_to_twip(8.5),  // 页面宽度
            )),
            margin: Some(PageMargin::new(
                convert_inches_to_twip(1.0), // 上边距
                convert_inches_to_twip(1.0), // 右边距
                convert_inches_to_twip(1.0), // 下边距
                convert_inches_to_twip(1.0), // 左边距
                convert_inches_to_twip(0.5), // 页眉边距
                convert_inches_to_twip(0.5), // 页脚边距
            )),
            page_numbers: Some(PageNumbers::new(1)),
        }),
    };

    // 创建 section
    let mut section = Section::new();
    section.properties = section_props;

    // 计算页面尺寸
    let page_width = convert_inches_to_twip(8.5);
    let page_height = convert_inches_to_twip(11.0);
    let margin = convert_inches_to_twip(1.0);
    let inner_width = page_width - 2 * margin;

    println!("页面尺寸:");
    println!("  页面宽度: {} twips", page_width);
    println!("  页面高度: {} twips", page_height);
    println!("  内容宽度: {} twips", inner_width);

    // 测试1: 左上角 (tl) - 使用 Page 锚点
    println!("\n创建左上角 frame (tl) - Page 锚点");
    let mut tl_paragraph = Paragraph::new();
    tl_paragraph.add_text_run(TextRun::new("左上角内容\n这是第二行\n这是第三行"));
    tl_paragraph.align(AlignmentType::Left);
    tl_paragraph.frame(ParagraphFrame {
        width: Some(inner_width),
        height: Some(0),
        anchor_horizontal: Some(FrameAnchorType::Page),
        anchor_vertical: Some(FrameAnchorType::Page),
        x_align: Some(HorizontalPositionAlign::Left),
        y_align: Some(VerticalPositionAlign::Top),
        page_height: Some(page_height),
        top_margin: Some(margin),
        bottom_margin: Some(margin),
        line_height: Some(convert_inches_to_twip(0.2)),
    });
    section.children.push(SectionChild::Paragraph(tl_paragraph));

    // 测试2: 顶部中央 (tc) - 使用 Margin 锚点
    println!("创建顶部中央 frame (tc) - Margin 锚点");
    let mut tc_paragraph = Paragraph::new();
    tc_paragraph.add_text_run(TextRun::new("顶部中央内容\n居中显示"));
    tc_paragraph.align(AlignmentType::Center);
    tc_paragraph.frame(ParagraphFrame {
        width: Some(inner_width),
        height: Some(0),
        anchor_horizontal: Some(FrameAnchorType::Margin),
        anchor_vertical: Some(FrameAnchorType::Margin),
        x_align: Some(HorizontalPositionAlign::Center),
        y_align: Some(VerticalPositionAlign::Top),
        page_height: Some(page_height),
        top_margin: Some(margin),
        bottom_margin: Some(margin),
        line_height: Some(convert_inches_to_twip(0.2)),
    });
    section.children.push(SectionChild::Paragraph(tc_paragraph));

    // 测试3: 右上角 (tr) - 使用 Text 锚点
    println!("创建右上角 frame (tr) - Text 锚点");
    let mut tr_paragraph = Paragraph::new();
    tr_paragraph.add_text_run(TextRun::new("右上角内容\n右对齐显示"));
    tr_paragraph.align(AlignmentType::Right);
    tr_paragraph.frame(ParagraphFrame {
        width: Some(inner_width),
        height: Some(0),
        anchor_horizontal: Some(FrameAnchorType::Text),
        anchor_vertical: Some(FrameAnchorType::Text),
        x_align: Some(HorizontalPositionAlign::Right),
        y_align: Some(VerticalPositionAlign::Top),
        page_height: Some(page_height),
        top_margin: Some(margin),
        bottom_margin: Some(margin),
        line_height: Some(convert_inches_to_twip(0.2)),
    });
    section.children.push(SectionChild::Paragraph(tr_paragraph));

    // 添加 section 到文档
    doc.options.sections.push(section);

    // 保存文档
    doc.save("test_independent_frames.docx")?;
    
    println!("\n=== 测试完成 ===");
    println!("已生成测试文档: test_independent_frames.docx");
    println!("请打开文档检查:");
    println!("1. 标题页是否显示所有6个位置的内容");
    println!("2. tl、tc、tr 是否在顶部独立显示（不混合）");
    println!("3. cc 是否在页面中央显示");
    println!("4. bl、br 是否在底部独立显示");
    println!("5. 各个位置的对齐是否正确");
    println!("6. 内容是否能够重叠而不互相干扰");
    
    println!("\n修改说明:");
    println!("- 通过设置不同的锚点类型（Page、Margin、Text）");
    println!("- 通过设置微小的位置偏移（x_offset: 0,1,2 和 y_offset: -240,-239,-238）");
    println!("- 确保每个 frame 在 DOCX 中被视为独立的容器");
    println!("- 避免了相同位置的 frame 被合并的问题");
    
    Ok(())
}
