use crate::models::{Conf, ScriptToken};
use crate::parser::ParseOutput;
use std::collections::HashMap;

use super::docx_maker::{DocxOptions, DocxError, DocxResult, PrintProfile, generate_docx};
use super::metadata_extractor::{extract_metadata_from_parsed_document, ExtractedMetadata};



/// DOCX生成错误
pub type DocxGenerateError = DocxError;

/// DOCX生成结果
pub type DocxGenerateResult<T> = DocxResult<T>;

/// DOCX统计信息
#[derive(Debug)]
pub struct DocxStats {
    pub page_count: u32,
    pub page_count_real: u32,
    pub line_map: HashMap<u32, LineStruct>,
}

/// 行结构信息
#[derive(Debug)]
pub struct LineStruct {
    pub sections: Vec<String>,
    pub scene: String,
    pub page: u32,
    pub cumulative_duration: f32,
}

/// DOCX Base64编码结果
#[derive(Debug)]
pub struct DocxAsBase64 {
    pub data: String,
    pub stats: DocxStats,
}

/// 生成DOCX文档
///
/// # 参数
///
/// * `output_path` - 输出文件路径，如果为"$STATS$"则返回统计信息，如果为"$PREVIEW$"则返回Base64编码的文档
/// * `config` - 配置信息
/// * `parsed_document` - 解析后的文档
///
/// # 返回值
///
/// 如果`output_path`为"$STATS$"，则返回统计信息
/// 如果`output_path`为"$PREVIEW$"，则返回Base64编码的文档
/// 否则，生成DOCX文件并返回Ok(())
pub async fn generate_docx_document(
    output_path: &str,
    config: &Conf,
    parsed_document: &ParseOutput,
) -> DocxGenerateResult<Option<DocxStats>> {
    println!("【generate_docx_document】开始生成 DOCX 文档");
    println!("【generate_docx_document】文件路径: {}", output_path);

    // 创建一个可变的解析结果副本
    let mut parsed_document_copy = parsed_document.clone();

    println!("【generate_docx_document】标题页元素数量: {}", parsed_document_copy.title_page.len());
    for (key, tokens) in &parsed_document_copy.title_page {
        println!("【generate_docx_document】标题页元素 {}: {} 个 token", key, tokens.len());
        if key == "cc" {
            for (i, token) in tokens.iter().enumerate() {
                println!("【generate_docx_document】cc token {}: {}", i, token.text);
            }
        }
    }

    // 提取元数据
    println!("【generate_docx_document】开始从标题页提取元数据");
    let extracted_metadata = extract_metadata_from_parsed_document(&parsed_document_copy, &config.font_family);
    println!("【generate_docx_document】元数据提取完成");

    let metadata = extracted_metadata.metadata;
    let watermark = extracted_metadata.watermark;
    let header = extracted_metadata.header;
    let footer = extracted_metadata.footer;
    let font = extracted_metadata.font;
    let font_bold = extracted_metadata.font_bold;
    let font_italic = extracted_metadata.font_italic;
    let font_bold_italic = extracted_metadata.font_bold_italic;
    
    let mut print_sections = config.print_sections.clone();
    let mut print_synopsis = config.print_synopsis.clone();
    
    if let Some(print_sections_str) = metadata.get("print.print_sections") {
        if let Ok(_value) = print_sections_str.parse::<f32>() {
            print_sections = print_sections_str != "0";
        }
    }
    
    if let Some(print_synopsis_str) = metadata.get("print.print_synopsis") {
        if let Ok(_value) = print_synopsis_str.parse::<f32>() {
            print_synopsis = print_synopsis_str != "0";
        }
    }
    

    // 预处理 tokens
    let mut current_index = 0;
    let mut previous_type = String::new();
    let mut invisible_sections = Vec::new();

    while current_index < parsed_document_copy.tokens.len() {
        // 克隆当前标记，避免借用问题
        let current_token = parsed_document_copy.tokens[current_index].clone();

        // 检查是否需要跳过当前标记
        let skip = match current_token.token_type.as_str() {
            "dual_dialogue_begin" | "dialogue_begin" | "dialogue_end" | "dual_dialogue_end" => true,
            "action" | "transition" | "centered" | "shot" => !config.print_actions,
            "note" => !config.print_notes,
            "scene_heading" => !config.print_headers,
            "section" => !print_sections,
            "synopsis" => !print_synopsis,
            "separator" => config.merge_empty_lines && previous_type == "separator",
            _ => {
                // 检查是否是对话
                if current_token.token_type == "dialogue" {
                    !config.print_dialogues
                } else {
                    false
                }
            }
        };

        if skip {
            if current_token.token_type == "section" {
                // 在下一个场景标题处添加不可见的章节（用于创建书签和生成docx侧边栏）
                invisible_sections.push(current_token.clone());
            }

            // 移除当前标记
            parsed_document_copy.tokens.remove(current_index);
            continue;
        }

        // 处理场景标题
        if current_token.token_type == "scene_heading" {
            if !invisible_sections.is_empty() {
                // 添加不可见的章节到场景标题
                let mut token = current_token.clone();
                token.invisible_sections = Some(invisible_sections.clone());
                parsed_document_copy.tokens[current_index] = token;
                invisible_sections.clear();
            }
        }

        // 在场景之间添加额外的分隔符
        if config.double_space_between_scenes && current_token.token_type == "scene_heading" && current_token.number.as_deref() != Some("1") {
            // 创建额外的分隔符
            let separator = ScriptToken {
                token_type: "separator".to_string(),
                text: String::new(),
                line: current_token.line,
                start: current_token.start,
                end: current_token.end,
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
            parsed_document_copy.tokens.insert(current_index, separator);
            current_index += 1;
        }

        previous_type = current_token.token_type.clone();
        current_index += 1;
    }

    // 清理末尾的分隔符
    while !parsed_document_copy.tokens.is_empty() && parsed_document_copy.tokens.last().unwrap().token_type == "separator" {
        parsed_document_copy.tokens.pop();
    }

    // 设置配置
    let mut config_copy = config.clone();

    if let Some(wm) = watermark {
        config_copy.print_watermark = wm;
    }

    if let Some(h) = header {
        config_copy.print_header = h;
    }

    if let Some(f) = footer {
        config_copy.print_footer = f;
    }
    
    config_copy.print_sections = print_sections.clone();
    config_copy.print_synopsis = print_synopsis.clone();

    // 处理行
    // 这里应该调用 line2 函数，但我们使用 process_document_lines 代替
    crate::docx::line_processor::process_document_lines(&mut parsed_document_copy, &config_copy);

    // 设置打印配置
    let mut print_profile = config.print_profile.clone();

    // 根据配置设置打印配置
    if config.page_size == "Letter" {
        print_profile.page_width = 8.5;
        print_profile.page_height = 11.0;
        print_profile.paper_size = "Letter".to_string();
    }

    // 使用默认值
    print_profile.top_margin = 1.0;
    print_profile.bottom_margin = 1.0;
    print_profile.left_margin = 1.5;
    print_profile.right_margin = 1.0;

    // 不需要额外设置元数据，保持与原始 TypeScript 版本一致

    // 从元数据中更新打印配置
    // 检查是否存在 print 对象
    let has_print_object = metadata.contains_key("print");

    // 处理嵌套的 print 对象中的配置
    if has_print_object {
        // 处理行数配置
        if let Some(lines_per_page) = metadata.get("print.lines_per_page") {
            if let Ok(value) = lines_per_page.parse::<usize>() {
                print_profile.lines_per_page = value;
            }
        }

        // 处理边距配置
        if let Some(top_margin) = metadata.get("print.top_margin") {
            if let Ok(value) = top_margin.parse::<f32>() {
                print_profile.top_margin = value;
            }
        }

        if let Some(bottom_margin) = metadata.get("print.bottom_margin") {
            if let Ok(value) = bottom_margin.parse::<f32>() {
                print_profile.bottom_margin = value;
            }
        }

        if let Some(left_margin) = metadata.get("print.left_margin") {
            if let Ok(value) = left_margin.parse::<f32>() {
                print_profile.left_margin = value;
            }
        }

        if let Some(right_margin) = metadata.get("print.right_margin") {
            if let Ok(value) = right_margin.parse::<f32>() {
                print_profile.right_margin = value;
            }
        }

        // 处理页面尺寸配置
        if let Some(page_height) = metadata.get("print.page_height") {
            if let Ok(value) = page_height.parse::<f32>() {
                print_profile.page_height = value;
            }
        }

        if let Some(page_width) = metadata.get("print.page_width") {
            if let Ok(value) = page_width.parse::<f32>() {
                print_profile.page_width = value;
            }
        }

        if let Some(page_number_top_margin) = metadata.get("print.page_number_top_margin") {
            if let Ok(value) = page_number_top_margin.parse::<f32>() {
                print_profile.page_number_top_margin = value;
            }
        }

        // 处理纸张大小配置
        if let Some(paper_size) = metadata.get("print.paper_size") {
            print_profile.paper_size = paper_size.clone();
        }

        // 处理字体配置
        if let Some(font_size) = metadata.get("print.font_size") {
            if let Ok(value) = font_size.parse::<f32>() {
                print_profile.font_size = value;
            }
        }

        if let Some(note_font_size) = metadata.get("print.note_font_size") {
            if let Ok(value) = note_font_size.parse::<f32>() {
                print_profile.note_font_size = value;
            }
        }

        if let Some(font_width) = metadata.get("print.font_width") {
            if let Ok(value) = font_width.parse::<f32>() {
                print_profile.font_width = value;
            }
        }

        if let Some(note_line_height) = metadata.get("print.note_line_height") {
            if let Ok(value) = note_line_height.parse::<f32>() {
                print_profile.note_line_height = value;
            }
        }
    }

    // 从元数据中读取 action 和 scene_heading 的 feed 值，如果没有则使用默认值
    if let Some(action_feed) = metadata.get("print.action.feed") {
        if let Ok(value) = action_feed.parse::<f32>() {
            print_profile.action.feed = value;
        }
    }

    if let Some(scene_heading_feed) = metadata.get("print.scene_heading.feed") {
        if let Ok(value) = scene_heading_feed.parse::<f32>() {
            print_profile.scene_heading.feed = value;
        }
    }

    // 如果没有从元数据中读取到值，则使用默认值（保持与原项目一致）
    // 注意：不要强制设置为 left_margin，这会覆盖从元数据中读取的值

    let inner_width = print_profile.page_width - print_profile.left_margin - print_profile.right_margin;
    let indent = print_profile.action.feed - print_profile.left_margin;
    let available_width = inner_width - indent - indent;

    print_profile.character.feed = (available_width / 2.0) + print_profile.action.feed - print_profile.font_width * 7.0;
    print_profile.dialogue.feed = (print_profile.character.feed - print_profile.action.feed) / 2.0 + print_profile.action.feed;
    print_profile.parenthetical.feed = (print_profile.character.feed - print_profile.dialogue.feed) / 2.0 + print_profile.dialogue.feed;

    // 计算行高
    let line_height = (print_profile.page_height - print_profile.top_margin - print_profile.bottom_margin) / print_profile.lines_per_page as f32;
    let line_height = (line_height * 100.0).round() / 100.0;

    // 调整底部边距
    print_profile.bottom_margin = ((print_profile.page_height - print_profile.top_margin - (print_profile.lines_per_page as f32 * line_height)) * 100.0).round() / 100.0;
    
    println!("【bottom_margin】:");
    println!("  bottom_margin: {} pt", print_profile.bottom_margin);

    // 创建DOCX选项
    let mut docx_options = DocxOptions::default();
    docx_options.filepath = output_path.to_string();
    docx_options.config = config_copy;
    docx_options.parsed = Some(parsed_document_copy);
    docx_options.print_profile = print_profile;
    docx_options.font = font;
    docx_options.font_italic = font_italic;
    docx_options.font_bold = font_bold;
    docx_options.font_bold_italic = font_bold_italic;
    docx_options.line_height = line_height;
    docx_options.metadata = Some(metadata);
    docx_options.for_preview = output_path == "$PREVIEW$";

    // 根据输出路径处理不同的情况
    if output_path == "$STATS$" {
        // 返回统计信息
        println!("【generate_docx_document】开始获取统计信息 - get_docx_stats 分支");
        println!("【generate_docx_document】title_page_processed = {}", docx_options.title_page_processed);
        let stats = super::docx_maker::get_docx_stats(docx_options).await?;
        Ok(Some(DocxStats {
            page_count: stats.page_count as u32,
            page_count_real: stats.page_count_real as u32,
            line_map: stats.line_map.into_iter()
                .map(|(k, v)| (k as u32, LineStruct {
                    sections: v.sections,
                    scene: v.scene,
                    page: v.page as u32,
                    cumulative_duration: v.cumulative_duration,
                }))
                .collect(),
        }))
    } else if output_path == "$PREVIEW$" {
        // 返回Base64编码的文档
        println!("【generate_docx_document】开始生成预览 - get_docx_base64 分支");
        println!("【generate_docx_document】title_page_processed = {}", docx_options.title_page_processed);
        let base64_result = super::docx_maker::get_docx_base64(docx_options).await?;
        let stats = DocxStats {
            page_count: base64_result.stats.page_count as u32,
            page_count_real: base64_result.stats.page_count_real as u32,
            line_map: base64_result.stats.line_map.into_iter()
                .map(|(k, v)| (k as u32, LineStruct {
                    sections: v.sections,
                    scene: v.scene,
                    page: v.page as u32,
                    cumulative_duration: v.cumulative_duration,
                }))
                .collect(),
        };

        Ok(Some(stats))
    } else {
        // 生成DOCX文件
        println!("【generate_docx_document】开始生成 DOCX 文件 - get_docx 分支");
        println!("【generate_docx_document】title_page_processed = {}", docx_options.title_page_processed);
        super::docx_maker::get_docx(docx_options).await?;
        println!("【generate_docx_document】DOCX 文件生成完成");
        Ok(None)
    }
}

/// 获取DOCX统计信息
pub async fn get_docx_stats(
    config: &Conf,
    parsed_document: &ParseOutput,
) -> DocxGenerateResult<DocxStats> {
    // 创建一个可变的解析结果副本
    let mut parsed_document_copy = parsed_document.clone();

    // 处理行
    crate::docx::line_processor::process_document_lines(&mut parsed_document_copy, config);

    // 创建DOCX选项
    let mut docx_options = DocxOptions::default();
    docx_options.filepath = "$STATS$".to_string();
    docx_options.config = config.clone();
    docx_options.parsed = Some(parsed_document_copy);
    docx_options.print_profile = config.print_profile.clone();
    docx_options.line_height = 1.0;

    // 直接调用 docx_maker::get_docx_stats 而不是通过 generate_docx_document
    let stats = super::docx_maker::get_docx_stats(docx_options).await?;

    Ok(DocxStats {
        page_count: stats.page_count as u32,
        page_count_real: stats.page_count_real as u32,
        line_map: stats.line_map.into_iter()
            .map(|(k, v)| (k as u32, LineStruct {
                sections: v.sections,
                scene: v.scene,
                page: v.page as u32,
                cumulative_duration: v.cumulative_duration,
            }))
            .collect(),
    })
}

/// 获取DOCX Base64编码
pub async fn get_docx_base64(
    config: &Conf,
    parsed_document: &ParseOutput,
) -> DocxGenerateResult<DocxAsBase64> {
    // 创建一个可变的解析结果副本
    let mut parsed_document_copy = parsed_document.clone();

    // 提取元数据
    let extracted_metadata = extract_metadata_from_parsed_document(&parsed_document_copy, &config.font_family);

    let metadata = extracted_metadata.metadata;
    let watermark = extracted_metadata.watermark;
    let header = extracted_metadata.header;
    let footer = extracted_metadata.footer;
    let font = extracted_metadata.font;
    let font_bold = extracted_metadata.font_bold;
    let font_italic = extracted_metadata.font_italic;
    let font_bold_italic = extracted_metadata.font_bold_italic;
    
    let mut print_sections = config.print_sections.clone();
    let mut print_synopsis = config.print_synopsis.clone();
    
    if let Some(print_sections_str) = metadata.get("print.print_sections") {
        if let Ok(_value) = print_sections_str.parse::<f32>() {
            print_sections = print_sections_str != "0";
        }
    }
    
    if let Some(print_synopsis_str) = metadata.get("print.print_synopsis") {
        if let Ok(_value) = print_synopsis_str.parse::<f32>() {
            print_synopsis = print_synopsis_str != "0";
        }
    }

    // 预处理 tokens
    let mut current_index = 0;
    let mut previous_type = String::new();
    let mut invisible_sections = Vec::new();

    while current_index < parsed_document_copy.tokens.len() {
        // 克隆当前标记，避免借用问题
        let current_token = parsed_document_copy.tokens[current_index].clone();

        // 检查是否需要跳过当前标记
        let skip = match current_token.token_type.as_str() {
            "dual_dialogue_begin" | "dialogue_begin" | "dialogue_end" | "dual_dialogue_end" => true,
            "action" | "transition" | "centered" | "shot" => !config.print_actions,
            "note" => !config.print_notes,
            "scene_heading" => !config.print_headers,
            "section" => !print_sections,
            "synopsis" => !print_synopsis,
            "separator" => config.merge_empty_lines && previous_type == "separator",
            _ => {
                // 检查是否是对话
                if current_token.token_type == "dialogue" {
                    !config.print_dialogues
                } else {
                    false
                }
            }
        };

        if skip {
            if current_token.token_type == "section" {
                // 在下一个场景标题处添加不可见的章节（用于创建书签和生成docx侧边栏）
                invisible_sections.push(current_token.clone());
            }

            // 移除当前标记
            parsed_document_copy.tokens.remove(current_index);
            continue;
        }

        // 处理场景标题
        if current_token.token_type == "scene_heading" {
            if !invisible_sections.is_empty() {
                // 添加不可见的章节到场景标题
                let mut token = current_token.clone();
                token.invisible_sections = Some(invisible_sections.clone());
                parsed_document_copy.tokens[current_index] = token;
                invisible_sections.clear();
            }
        }

        // 在场景之间添加额外的分隔符
        if config.double_space_between_scenes && current_token.token_type == "scene_heading" && current_token.number.as_deref() != Some("1") {
            // 创建额外的分隔符
            let separator = ScriptToken {
                token_type: "separator".to_string(),
                text: String::new(),
                line: current_token.line,
                start: current_token.start,
                end: current_token.end,
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
            parsed_document_copy.tokens.insert(current_index, separator);
            current_index += 1;
        }

        previous_type = current_token.token_type.clone();
        current_index += 1;
    }

    // 清理末尾的分隔符
    while !parsed_document_copy.tokens.is_empty() && parsed_document_copy.tokens.last().unwrap().token_type == "separator" {
        parsed_document_copy.tokens.pop();
    }

    // 设置配置
    let mut config_copy = config.clone();

    if let Some(wm) = watermark {
        config_copy.print_watermark = wm;
    }

    if let Some(h) = header {
        config_copy.print_header = h;
    }

    if let Some(f) = footer {
        config_copy.print_footer = f;
    }
    
    config_copy.print_sections = print_sections.clone();
    config_copy.print_synopsis = print_synopsis.clone();

    // 处理行
    crate::docx::line_processor::process_document_lines(&mut parsed_document_copy, &config_copy);

    // 设置打印配置
    let mut print_profile = config.print_profile.clone();

    // 根据配置设置打印配置
    if config.page_size == "Letter" {
        print_profile.page_width = 8.5;
        print_profile.page_height = 11.0;
        print_profile.paper_size = "Letter".to_string();
    }

    // 使用默认值
    print_profile.top_margin = 1.0;
    print_profile.bottom_margin = 1.0;
    print_profile.left_margin = 1.5;
    print_profile.right_margin = 1.0;

    // 从元数据中更新打印配置
    // 检查是否存在 print 对象
    let has_print_object = metadata.contains_key("print");

    // 处理嵌套的 print 对象中的配置
    if has_print_object {
        // 处理行数配置
        if let Some(lines_per_page) = metadata.get("print.lines_per_page") {
            if let Ok(value) = lines_per_page.parse::<usize>() {
                print_profile.lines_per_page = value;
            }
        }

        // 处理边距配置
        if let Some(top_margin) = metadata.get("print.top_margin") {
            if let Ok(value) = top_margin.parse::<f32>() {
                print_profile.top_margin = value;
            }
        }

        if let Some(bottom_margin) = metadata.get("print.bottom_margin") {
            if let Ok(value) = bottom_margin.parse::<f32>() {
                print_profile.bottom_margin = value;
            }
        }

        if let Some(left_margin) = metadata.get("print.left_margin") {
            if let Ok(value) = left_margin.parse::<f32>() {
                print_profile.left_margin = value;
            }
        }

        if let Some(right_margin) = metadata.get("print.right_margin") {
            if let Ok(value) = right_margin.parse::<f32>() {
                print_profile.right_margin = value;
            }
        }

        // 处理页面尺寸配置
        if let Some(page_height) = metadata.get("print.page_height") {
            if let Ok(value) = page_height.parse::<f32>() {
                print_profile.page_height = value;
            }
        }

        if let Some(page_width) = metadata.get("print.page_width") {
            if let Ok(value) = page_width.parse::<f32>() {
                print_profile.page_width = value;
            }
        }

        if let Some(page_number_top_margin) = metadata.get("print.page_number_top_margin") {
            if let Ok(value) = page_number_top_margin.parse::<f32>() {
                print_profile.page_number_top_margin = value;
            }
        }

        // 处理纸张大小配置
        if let Some(paper_size) = metadata.get("print.paper_size") {
            print_profile.paper_size = paper_size.clone();
        }

        // 处理字体配置
        if let Some(font_size) = metadata.get("print.font_size") {
            if let Ok(value) = font_size.parse::<f32>() {
                print_profile.font_size = value;
            }
        }

        if let Some(note_font_size) = metadata.get("print.note_font_size") {
            if let Ok(value) = note_font_size.parse::<f32>() {
                print_profile.note_font_size = value;
            }
        }

        if let Some(font_width) = metadata.get("print.font_width") {
            if let Ok(value) = font_width.parse::<f32>() {
                print_profile.font_width = value;
            }
        }

        if let Some(note_line_height) = metadata.get("print.note_line_height") {
            if let Ok(value) = note_line_height.parse::<f32>() {
                print_profile.note_line_height = value;
            }
        }
    }

    // 从元数据中读取 action 和 scene_heading 的 feed 值，如果没有则使用默认值
    if let Some(action_feed) = metadata.get("print.action.feed") {
        if let Ok(value) = action_feed.parse::<f32>() {
            print_profile.action.feed = value;
        }
    }

    if let Some(scene_heading_feed) = metadata.get("print.scene_heading.feed") {
        if let Ok(value) = scene_heading_feed.parse::<f32>() {
            print_profile.scene_heading.feed = value;
        }
    }

    // 如果没有从元数据中读取到值，则使用默认值（保持与原项目一致）
    // 注意：不要强制设置为 left_margin，这会覆盖从元数据中读取的值

    let inner_width = print_profile.page_width - print_profile.left_margin - print_profile.right_margin;
    let indent = print_profile.action.feed - print_profile.left_margin;
    let available_width = inner_width - indent - indent;

    print_profile.character.feed = (available_width / 2.0) + print_profile.action.feed - print_profile.font_width * 7.0;
    print_profile.dialogue.feed = (print_profile.character.feed - print_profile.action.feed) / 2.0 + print_profile.action.feed;
    print_profile.parenthetical.feed = (print_profile.character.feed - print_profile.dialogue.feed) / 2.0 + print_profile.dialogue.feed;

    // 计算行高
    let line_height = (print_profile.page_height - print_profile.top_margin - print_profile.bottom_margin) / print_profile.lines_per_page as f32;
    let line_height = (line_height * 100.0).round() / 100.0;

    // 调整底部边距
    print_profile.bottom_margin = ((print_profile.page_height - print_profile.top_margin - (print_profile.lines_per_page as f32 * line_height)) * 100.0).round() / 100.0;

    // 创建DOCX选项
    let mut docx_options = DocxOptions::default();
    docx_options.filepath = "$PREVIEW$".to_string();
    docx_options.config = config_copy;
    docx_options.parsed = Some(parsed_document_copy);
    docx_options.print_profile = print_profile;
    docx_options.font = font;
    docx_options.font_italic = font_italic;
    docx_options.font_bold = font_bold;
    docx_options.font_bold_italic = font_bold_italic;
    docx_options.line_height = line_height;
    docx_options.metadata = Some(metadata);
    docx_options.for_preview = true;

    // 获取 Base64 编码
    let base64_result = super::docx_maker::get_docx_base64(docx_options).await?;

    // 返回结果
    Ok(DocxAsBase64 {
        data: base64_result.data,
        stats: DocxStats {
            page_count: base64_result.stats.page_count as u32,
            page_count_real: base64_result.stats.page_count_real as u32,
            line_map: base64_result.stats.line_map.into_iter()
                .map(|(k, v)| (k as u32, LineStruct {
                    sections: v.sections,
                    scene: v.scene,
                    page: v.page as u32,
                    cumulative_duration: v.cumulative_duration,
                }))
                .collect(),
        },
    })
}