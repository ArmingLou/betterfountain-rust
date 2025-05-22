use betterfountain_rust::models::Conf;
use betterfountain_rust::parser::fountain_parser::FountainParser;
use std::fs;
use std::path::Path;

#[test]
fn test_chinese_script_parsing() {
    // 创建解析器
    let mut parser = FountainParser::new();

    // 读取中文测试文件
    let script_path = Path::new("tests/test_data/黑色爱情诗.fountain");
    let script = fs::read_to_string(script_path).expect("无法读取测试文件");

    // 解析剧本
    let mut conf = Conf::default();
    conf.print_notes = true;
    conf.merge_empty_lines = true;
    conf.print_dialogue_numbers = false;
    conf.each_scene_on_new_page = false;
    conf.use_dual_dialogue = true;
    conf.dialogue_foldable = true;
    conf.emitalic_dialog = true;
    conf.embolden_character_names = false;
    conf.text_contd = "(CONT'D)".to_string();
    conf.dial_sec_per_char = 0.3;
    conf.dial_sec_per_punc_short = 0.3;
    conf.dial_sec_per_punc_long = 0.75;
    conf.action_sec_per_char = 0.4;

    let result = parser.parse(&script, &conf, true);

    // 打印详细结果
    println!("=== 解析结果 ===");
    println!("剧本内容:\n{}", script);
    println!("\n解析出的标记:");
    for token in &result.tokens {
        println!("- {}: {}", token.token_type, token.text);
        if let Some(dual) = &token.dual {
            if !dual.is_empty() {
                println!("  双对话: {}", dual);
            }
        }
    }

    // 统计有时间的token
    let mut action_tokens = 0;
    let mut dialogue_tokens = 0;
    let mut total_action_time = 0.0;
    let mut total_dialogue_time = 0.0;

    // 检查特定行的token的text_valid内容
    for token in &result.tokens {
        if token.line >= 60 && token.line <= 65 {
            println!("DEBUG: 行{} - 类型: '{}', text: '{}', text_no_notes: '{:?}', time: {:?}",
                     token.line, token.token_type, token.text, token.text_no_notes, token.time);
        }

        if let Some(time) = token.time {
            if token.token_type == "action" {
                action_tokens += 1;
                total_action_time += time;
            } else if token.token_type == "dialogue" {
                dialogue_tokens += 1;
                total_dialogue_time += time;
            }
        }
    }

    println!("\n统计信息:");
    println!("- 动作长度: {}", result.length_action);
    println!("- 对话长度: {}", result.length_dialogue);
    println!("- 解析耗时: {}ms", result.parse_time);
    println!("- 场景数量: {}", result.properties.scenes.len());
    println!("- Token数量: {}", result.tokens.len());
    println!("- 动作Token数量: {}, 总时间: {}", action_tokens, total_action_time);
    println!("- 对话Token数量: {}, 总时间: {}", dialogue_tokens, total_dialogue_time);
    println!("- 角色列表: {:?}", result.properties.characters.keys().collect::<Vec<_>>());
    println!("- 地点列表: {:?}", result.properties.locations.keys().collect::<Vec<_>>());

    // 验证新添加的时长计算参数
    println!("- 对话每字符时长: {}", result.dial_sec_per_char);
    println!("- 对话短标点时长: {}", result.dial_sec_per_punc_short);
    println!("- 对话长标点时长: {}", result.dial_sec_per_punc_long);
    println!("- 动作每字符时长: {}", result.action_sec_per_char);

    // 验证结果
    assert!(result.tokens.len() > 0, "应该解析出至少一个标记");
    assert!(result.properties.scenes.len() > 0, "应该解析出至少一个场景");
    assert!(!result.properties.characters.is_empty(), "应该解析出角色");

    // 验证新添加的时长计算参数
    assert_eq!(result.dial_sec_per_char, 0.1, "对话每字符时长应该为0.1");
    assert_eq!(result.dial_sec_per_punc_short, 0.2, "对话短标点时长应该为0.2");
    assert_eq!(result.dial_sec_per_punc_long, 0.6, "对话长标点时长应该为0.6");
    assert_eq!(result.action_sec_per_char, 0.7, "动作每字符时长应该为0.7");

    // 验证特定角色是否存在
    assert!(
        result.properties.characters.contains_key("顾清"),
        "应该包含角色'顾清'"
    );
    assert!(
        result.properties.characters.contains_key("林静怡"),
        "应该包含角色'林静怡'"
    );
    assert!(
        result.properties.characters.contains_key("JOHN"),
        "应该包含角色'JOHN'"
    );
    assert!(
        result.properties.characters.contains_key("JANE"),
        "应该包含角色'JANE'"
    );

    // 验证特定地点是否存在
    assert!(!result.properties.locations.is_empty(), "应该解析出地点");

    // 检查特定地点
    let locations = result.properties.locations.keys().collect::<Vec<_>>();
    println!("所有地点: {:?}", locations);

    // 验证是否包含特定地点
    let has_park = result.properties.locations.keys().any(|k| k == "公园一角");
    let has_house = result.properties.locations.keys().any(|k| k == "顾清住处");
    let has_eng_house = result.properties.locations.keys().any(|k| k == "HOUSE");
    let has_eng_park = result.properties.locations.keys().any(|k| k == "PARK");
    let has_combined = result
        .properties
        .locations
        .keys()
        .any(|k| k == "公园一角 / 顾清住处");

    assert!(has_park, "应该包含地点'公园一角'");
    assert!(has_house, "应该包含地点'顾清住处'");
    assert!(has_eng_house, "应该包含地点'HOUSE'");
    assert!(has_eng_park, "应该包含地点'PARK'");
    assert!(has_combined, "应该包含地点'公园一角 / 顾清住处'");
}

#[test]
fn test_time_calculation() {
    let parser = FountainParser::new();

    // 测试实际的对话文本
    let dialogue_text = "晚霞如画，绚丽如金，";
    let dialogue_time = parser.calculate_dialogue_duration(dialogue_text, None, None, None);
    println!("对话文本: '{}', 时间: {}", dialogue_text, dialogue_time);

    // 测试字符计算
    let chars = parser.calculate_chars(dialogue_text);
    println!("原文本: '{}', 有效字符: '{}', 长度: {}", dialogue_text, chars, chars.len());

    // 测试动作时间计算
    let action_text = "这是一个测试动作";
    let action_time = parser.calculate_action_duration(action_text, None);
    println!("动作文本: '{}', 时间: {}", action_text, action_time);
}

#[test]
fn test_duration_parameters() {
    let mut parser = FountainParser::new();

    // 测试自定义配置参数
    let mut custom_conf = Conf::default();
    custom_conf.print_notes = true;
    custom_conf.merge_empty_lines = true;
    custom_conf.print_dialogue_numbers = false;
    custom_conf.each_scene_on_new_page = false;
    custom_conf.use_dual_dialogue = true;
    custom_conf.dialogue_foldable = true;
    custom_conf.emitalic_dialog = true;
    custom_conf.embolden_character_names = false;
    custom_conf.text_contd = "(CONT'D)".to_string();
    custom_conf.dial_sec_per_char = 0.5;  // 自定义值
    custom_conf.dial_sec_per_punc_short = 0.4;  // 自定义值
    custom_conf.dial_sec_per_punc_long = 1.0;  // 自定义值
    custom_conf.action_sec_per_char = 0.6;  // 自定义值

    // 简单的测试脚本
    let script = r#"
INT. ROOM - DAY

这是一个动作描述。

JOHN
你好，世界！这是一句对话。
"#;

    let result = parser.parse(script, &custom_conf, false);

    // 验证配置参数被正确传递
    assert_eq!(result.dial_sec_per_char, 0.5, "对话每字符时长应该为自定义值0.5");
    assert_eq!(result.dial_sec_per_punc_short, 0.4, "对话短标点时长应该为自定义值0.4");
    assert_eq!(result.dial_sec_per_punc_long, 1.0, "对话长标点时长应该为自定义值1.0");
    assert_eq!(result.action_sec_per_char, 0.6, "动作每字符时长应该为自定义值0.6");

    // 验证解析结果
    assert!(result.tokens.len() > 0, "应该解析出标记");
    assert!(result.properties.scenes.len() > 0, "应该解析出场景");

    println!("自定义参数测试通过！");
    println!("- 对话每字符时长: {}", result.dial_sec_per_char);
    println!("- 对话短标点时长: {}", result.dial_sec_per_punc_short);
    println!("- 对话长标点时长: {}", result.dial_sec_per_punc_long);
    println!("- 动作每字符时长: {}", result.action_sec_per_char);
}
