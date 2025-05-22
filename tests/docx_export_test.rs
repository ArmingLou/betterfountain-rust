use betterfountain_rust::models::Conf;
use betterfountain_rust::parser::fountain_parser::FountainParser;
use betterfountain_rust::docx::docx::generate_docx_document;
use std::fs;
use std::path::Path;

#[tokio::test]
async fn test_docx_export() {
    eprintln!("测试开始");

    // 创建解析器
    let mut parser = FountainParser::new();
    eprintln!("解析器创建成功");

    // 读取中文测试文件
    let script_path = Path::new("tests/test_data/黑色爱情诗.fountain");
    let script = fs::read_to_string(script_path).expect("无法读取测试文件");
    eprintln!("读取测试文件成功，长度: {}", script.len());

    // 解析剧本
    let mut conf = Conf::default();
    // conf.note_position_bottom = false;
    // conf.use_dual_dialogue = false;
    // conf.print_sections = false;
    // conf.print_synopsis = false;

    eprintln!("配置创建成功");

    let result = parser.parse(&script, &conf, true);
    eprintln!("解析剧本成功");

    eprintln!("标题页元素数量: {}", result.title_page.len());
    for (key, tokens) in &result.title_page {
        eprintln!("标题页元素 {}: {} 个 token", key, tokens.len());
    }

    // 确保输出目录存在
    let output_dir = Path::new("tests/test_data_out");
    if !output_dir.exists() {
        std::fs::create_dir_all(output_dir).expect("无法创建输出目录");
    }
    eprintln!("输出目录确认存在");

    // 输出路径
    let output_path = "tests/test_data_out/黑色爱情诗.docx";
    eprintln!("输出路径: {}", output_path);

    // 导出 DOCX
    eprintln!("开始导出 DOCX");
    let docx_result = generate_docx_document(output_path, &conf, &result).await;
    eprintln!("DOCX 导出结果: {:?}", docx_result);

    // 验证导出结果
    assert!(docx_result.is_ok(), "DOCX 导出应该成功");

    // 验证文件是否存在
    let output_file = Path::new(output_path);
    eprintln!("检查文件是否存在: {}", output_file.display());
    eprintln!("文件存在: {}", output_file.exists());
    assert!(output_file.exists(), "导出的 DOCX 文件应该存在");

    eprintln!("DOCX 导出成功: {}", output_path);
}
