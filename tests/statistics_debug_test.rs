use betterfountain_rust::models::Conf;
use betterfountain_rust::parser::fountain_parser::FountainParser;
use std::fs;
use std::path::Path;

#[test]
fn test_statistics_debug() {
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
    conf.dial_sec_per_char = 0.1;
    conf.dial_sec_per_punc_short = 0.2;
    conf.dial_sec_per_punc_long = 0.6;
    conf.action_sec_per_char = 0.7;

    let result = parser.parse(&script, &conf, true, Some(true));

    println!("\n========== 统计调试信息 ==========");
    
    // 1. 基本统计信息
    println!("\n【1. 基本统计】");
    println!("- 动作总时长: {}", result.length_action);
    println!("- 对话总时长: {}", result.length_dialogue);
    println!("- 总时长: {}", result.length_action + result.length_dialogue);
    println!("- 场景数量: {}", result.properties.scenes.len());
    
    // 2. 打印每个场景的详细信息
    println!("\n【2. 场景详情】");
    for (i, scene) in result.properties.scenes.iter().enumerate() {
        let text = scene.get("text").and_then(|v| v.as_str()).unwrap_or("N/A");
        let number = scene.get("number").and_then(|v| v.as_str()).unwrap_or("N/A");
        let start = scene.get("startPlaySec").and_then(|v| v.as_f64()).unwrap_or(0.0);
        let end = scene.get("endPlaySec").and_then(|v| v.as_f64()).unwrap_or(0.0);
        let action_len = scene.get("actionLength").and_then(|v| v.as_f64()).unwrap_or(0.0);
        let dialogue_len = scene.get("dialogueLength").and_then(|v| v.as_f64()).unwrap_or(0.0);
        
        println!("  场景{}: {} (编号: {})", i + 1, text, number);
        println!("    - 开始时间: {}", start);
        println!("    - 结束时间: {}", end);
        println!("    - 动作时长: {}", action_len);
        println!("    - 对话时长: {}", dialogue_len);
        println!("    - 总时长: {}", action_len + dialogue_len);
    }
    
    // 3. 打印所有 token 中有时间的
    println!("\n【3. 有时长的Token】");
    let mut action_count = 0;
    let mut dialogue_count = 0;
    for token in &result.tokens {
        if let Some(time) = token.time {
            if token.token_type == "action" {
                action_count += 1;
                if action_count <= 10 {
                    let text_preview: String = token.text.chars().take(30).collect();
                    println!("  Action[行{}]: '{}' -> time={}", token.line, text_preview, time);
                }
            } else if token.token_type == "dialogue" {
                dialogue_count += 1;
                if dialogue_count <= 10 {
                    let char_name = token.character.as_deref().unwrap_or("N/A");
                    let text_preview: String = token.text.chars().take(20).collect();
                    println!("  Dialogue[行{}] {}: '{}' -> time={}", token.line, char_name, text_preview, time);
                }
            }
        }
    }
    println!("  总计: {} 个 action, {} 个 dialogue", action_count, dialogue_count);
    
    // 4. 打印统计结果
    println!("\n【4. 统计结果】");
    if let Some(ref stats) = result.statistics {
        // 角色统计
        println!("\n  角色统计:");
        for char_stat in &stats.character_stats.characters {
            println!("    - {}: 说话{}次, {}词, 对白{}秒, 总时长{}秒, 出现{}个场景, {}次独白", 
                char_stat.name, 
                char_stat.speaking_parts, 
                char_stat.words_spoken, 
                char_stat.seconds_spoken,
                char_stat.seconds_total,
                char_stat.number_of_scenes,
                char_stat.monologues);
        }
        
        // 打印顾清的所有对话文本
        println!("\n  顾清对话详情:");
        let mut total_chars = 0usize;
        for token in &result.tokens {
            if token.token_type == "dialogue" && token.character.as_deref() == Some("顾清") {
                let text = token.text_no_notes.as_deref().unwrap_or(&token.text);
                let chars = text.chars().filter(|c| !c.is_whitespace() && !c.is_ascii_punctuation()).count();
                total_chars += chars;
                let time = token.time.unwrap_or(0.0);
                println!("      行{}: '{}' 有效字符={}, time={}", token.line, text, chars, time);
            }
        }
        println!("    顾清总有效字符数: {}", total_chars);
        if total_chars > 0 {
            println!("    按0.1s/字计算: {}秒 (对比VSCode: 36秒)", total_chars as f64 * 0.1);
        }
        
        // 检查 action 中的 characters_action
        println!("\n  包含角色名的Action:");
        for token in &result.tokens {
            if token.token_type == "action" {
                if let Some(chars) = &token.characters_action {
                    if !chars.is_empty() {
                        let text_preview: String = token.text.chars().take(30).collect();
                        println!("    行{}: '{}' -> 角色: {:?}, time: {:?}", 
                            token.line, text_preview, chars, token.time);
                    }
                }
            }
        }
        
        // 再检查顾清的总时长是否包含action时间
        if let Some(ref stats) = result.statistics {
            if let Some(guqing) = stats.character_stats.characters.iter().find(|c| c.name == "顾清") {
                println!("\n  顾清详细: 对白={}s, 总时长={}s, 差异={}s",
                    guqing.seconds_spoken, guqing.seconds_total, 
                    guqing.seconds_total - guqing.seconds_spoken);
            }
        }
        
        // 时长统计
        println!("\n  时长统计:");
        println!("    - 对话时长: {}秒", stats.duration_stats.dialogue);
        println!("    - 动作时长: {}秒", stats.duration_stats.action);
        println!("    - 总时长: {}秒", stats.duration_stats.total);
        
        // 场景时长分布
        println!("\n  场景时长分布 (durationBySceneProp):");
        for prop in &stats.duration_stats.duration_by_scene_prop {
            println!("    - {}: {}秒", prop.prop, prop.duration);
        }
        
        // 场景详情
        println!("\n  场景列表:");
        for scene_item in &stats.duration_stats.scenes {
            println!("    - {} (type: {}, time: {})", scene_item.scene, scene_item.scene_type, scene_item.time);
        }
    } else {
        println!("  [警告] statistics 字段为 None!");
    }
    
    // 5. 打印 location 信息
    println!("\n【5. 地点统计】");
    for (loc_name, loc_refs) in &result.properties.locations {
        println!("  地点: {}", loc_name);
        for loc in loc_refs {
            println!("    - 场景号: {}, 内景: {}, 外景: {}, 时间: {}", 
                loc.scene_number, loc.interior, loc.exterior, loc.time_of_day);
        }
    }
    
    println!("\n========== 调试结束 ==========\n");
    
    // 验证基本数据
    assert!(result.properties.scenes.len() > 0, "应该有场景");
}