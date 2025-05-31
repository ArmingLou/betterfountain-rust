use betterfountain_rust::docx::adapter::docx::{Document, Paragraph, TextRun, SectionChild, Section, SectionProperties, PageProperties, PageSize, PageMargin, PageNumbers};
use betterfountain_rust::docx::adapter::{ParagraphFrame, FrameAnchorType, HorizontalPositionAlign, VerticalPositionAlign, AlignmentType, DocxAdapterError};
use betterfountain_rust::docx::convert_inches_to_twip;

fn main() -> Result<(), DocxAdapterError> {
    println!("=== 测试重叠 Frame 功能 ===");
    
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
    let inner_height = page_height - 2 * margin;
    
    println!("页面尺寸:");
    println!("  页面宽度: {} twips", page_width);
    println!("  页面高度: {} twips", page_height);
    println!("  内容宽度: {} twips", inner_width);
    println!("  内容高度: {} twips", inner_height);
    
    // 测试1: 左上角 (tl) - 独立 frame
    println!("\n创建左上角 frame (tl)");
    let mut tl_paragraph = Paragraph::new();
    tl_paragraph.add_text_run(TextRun::new("左上角内容\n这是第二行\n这是第三行"));
    tl_paragraph.align(AlignmentType::Left);
    tl_paragraph.frame(ParagraphFrame {
        width: Some(inner_width),
        height: Some(0),
        anchor_horizontal: Some(FrameAnchorType::Margin),
        anchor_vertical: Some(FrameAnchorType::Margin),
        x_align: Some(HorizontalPositionAlign::Left),
        y_align: Some(VerticalPositionAlign::Top),
        page_height: Some(page_height),
        top_margin: Some(margin),
        bottom_margin: Some(margin),
        line_height: Some(convert_inches_to_twip(0.2)),
    });
    section.children.push(SectionChild::Paragraph(tl_paragraph));
    
    // 测试2: 顶部中央 (tc) - 独立 frame，与 tl 重叠
    println!("创建顶部中央 frame (tc)");
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
    
    // 测试3: 右上角 (tr) - 独立 frame，与 tl 和 tc 重叠
    println!("创建右上角 frame (tr)");
    let mut tr_paragraph = Paragraph::new();
    tr_paragraph.add_text_run(TextRun::new("右上角内容\n右对齐显示"));
    tr_paragraph.align(AlignmentType::Right);
    tr_paragraph.frame(ParagraphFrame {
        width: Some(inner_width),
        height: Some(0),
        anchor_horizontal: Some(FrameAnchorType::Margin),
        anchor_vertical: Some(FrameAnchorType::Margin),
        x_align: Some(HorizontalPositionAlign::Right),
        y_align: Some(VerticalPositionAlign::Top),
        page_height: Some(page_height),
        top_margin: Some(margin),
        bottom_margin: Some(margin),
        line_height: Some(convert_inches_to_twip(0.2)),
    });
    section.children.push(SectionChild::Paragraph(tr_paragraph));
    
    // 测试4: 中央 (cc) - 独立 frame
    println!("创建中央 frame (cc)");
    let mut cc_paragraph = Paragraph::new();
    cc_paragraph.add_text_run(TextRun::new("中央内容\n这是页面中央的文本\n应该在页面中心显示"));
    cc_paragraph.align(AlignmentType::Center);
    cc_paragraph.frame(ParagraphFrame {
        width: Some(inner_width),
        height: Some(0),
        anchor_horizontal: Some(FrameAnchorType::Margin),
        anchor_vertical: Some(FrameAnchorType::Margin),
        x_align: Some(HorizontalPositionAlign::Center),
        y_align: Some(VerticalPositionAlign::Center),
        page_height: Some(page_height),
        top_margin: Some(margin),
        bottom_margin: Some(margin),
        line_height: Some(convert_inches_to_twip(0.2)),
    });
    section.children.push(SectionChild::Paragraph(cc_paragraph));
    
    // 测试5: 左下角 (bl) - 独立 frame
    println!("创建左下角 frame (bl)");
    let mut bl_paragraph = Paragraph::new();
    bl_paragraph.add_text_run(TextRun::new("左下角内容\n底部左对齐"));
    bl_paragraph.align(AlignmentType::Left);
    bl_paragraph.frame(ParagraphFrame {
        width: Some(inner_width),
        height: Some(0),
        anchor_horizontal: Some(FrameAnchorType::Page),
        anchor_vertical: Some(FrameAnchorType::Page),
        x_align: Some(HorizontalPositionAlign::Left),
        y_align: Some(VerticalPositionAlign::Bottom),
        page_height: Some(page_height),
        top_margin: Some(margin),
        bottom_margin: Some(margin),
        line_height: Some(convert_inches_to_twip(0.2)),
    });
    section.children.push(SectionChild::Paragraph(bl_paragraph));
    
    // 测试6: 右下角 (br) - 独立 frame，与 bl 重叠
    println!("创建右下角 frame (br)");
    let mut br_paragraph = Paragraph::new();
    br_paragraph.add_text_run(TextRun::new("右下角内容\n底部右对齐"));
    br_paragraph.align(AlignmentType::Right);
    br_paragraph.frame(ParagraphFrame {
        width: Some(inner_width),
        height: Some(0),
        anchor_horizontal: Some(FrameAnchorType::Page),
        anchor_vertical: Some(FrameAnchorType::Page),
        x_align: Some(HorizontalPositionAlign::Right),
        y_align: Some(VerticalPositionAlign::Bottom),
        page_height: Some(page_height),
        top_margin: Some(margin),
        bottom_margin: Some(margin),
        line_height: Some(convert_inches_to_twip(0.2)),
    });
    section.children.push(SectionChild::Paragraph(br_paragraph));
    
    // 添加 section 到文档
    doc.options.sections.push(section);
    
    // 保存文档
    doc.save("test_overlapping_frames.docx")?;
    
    println!("\n=== 测试完成 ===");
    println!("已生成测试文档: test_overlapping_frames.docx");
    println!("请打开文档检查:");
    println!("1. 左上角、顶部中央、右上角是否能在同一行重叠显示");
    println!("2. 中央内容是否在页面中心");
    println!("3. 左下角、右下角是否能在底部重叠显示");
    println!("4. 各个位置的对齐是否正确");
    println!("5. 内容是否能够重叠而不互相干扰");
    
    Ok(())
}
