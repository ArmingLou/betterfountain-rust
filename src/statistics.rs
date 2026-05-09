use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};

/// 完整统计数据结构
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Statistics {
    /// 角色统计
    pub character_stats: CharacterStatistics,
    /// 地点统计
    pub location_stats: LocationStatistics,
    /// 场景统计
    pub scene_stats: SceneStatistics,
    /// 时长统计
    pub duration_stats: DurationStatistics,
}

/// 角色统计数据
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CharacterStatistics {
    /// 每个角色的详细统计
    pub characters: Vec<CharacterStat>,
    /// 整体复杂度（中位数）
    pub complexity: f64,
    /// 角色总数
    pub character_count: usize,
    /// 独白总数
    pub monologues: usize,
}

/// 单个角色统计
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CharacterStat {
    /// 角色名称
    pub name: String,
    /// 说话次数（台词段数）
    pub speaking_parts: usize,
    /// 说话总词数
    pub words_spoken: usize,
    /// 对白时长（秒）- 仅对白
    pub seconds_spoken: f64,
    /// 总时长（秒）- 对白+角色动作，用于图表
    pub seconds_total: f64,
    /// 该角色台词平均复杂度
    pub average_complexity: f64,
    /// 该角色独白次数
    pub monologues: usize,
    /// 出现场景数
    pub number_of_scenes: usize,
    /// 颜色（用于图表）
    pub color: String,
}

/// 地点统计数据
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LocationStatistics {
    /// 每个地点的详细统计
    pub locations: Vec<LocationStat>,
    /// 地点总数
    pub locations_count: usize,
}

/// 单个地点统计
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LocationStat {
    /// 地点名称
    pub name: String,
    /// 场景编号列表
    pub scene_numbers: Vec<String>,
    /// 场景开始播放秒数
    pub scene_lines: Vec<f64>,
    /// 场景数量（去重后）
    pub number_of_scenes: usize,
    /// 出现的时间（day/night/morning/dusk/dawn）
    pub times_of_day: Vec<String>,
    /// 内/外景类型
    pub interior_exterior: String,
    /// 颜色（用于图表）
    pub color: String,
}

/// 场景统计数据
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SceneStatistics {
    /// 场景列表
    pub scenes: Vec<SceneStat>,
}

/// 单个场景统计
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SceneStat {
    /// 场景标题
    pub title: String,
}

/// 时长统计数据
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DurationStatistics {
    /// 对话总时长（秒）
    pub dialogue: f64,
    /// 动作总时长（秒）
    pub action: f64,
    /// 总时长（秒）
    pub total: f64,
    /// 动作时长图表数据
    pub lengthchart_action: Vec<LengthChartItem>,
    /// 对话时长图表数据
    pub lengthchart_dialogue: Vec<LengthChartItem>,
    /// 按场景属性分类的时长
    pub duration_by_scene_prop: Vec<DurationByProp>,
    /// 场景时长数据
    pub scenes: Vec<SceneItem>,
    /// 独白总数
    pub monologues: usize,
    /// 每个角色的时间线数据
    pub characters: Vec<Vec<CharacterDialogueItem>>,
    /// 角色名称列表
    pub characternames: Vec<String>,
}

/// 角色对话时间线数据项
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CharacterDialogueItem {
    /// 行号
    pub line: usize,
    /// 播放时间（秒）
    pub play_time_sec: f64,
    /// 场景名称
    pub scene: String,
    /// 累计时长（全局）
    pub length_time_global: f64,
    /// 累计词数（全局）
    pub length_words_global: f64,
    /// 是否独白
    pub monologue: bool,
    /// 本次时长
    pub length_time: f64,
    /// 本次词数
    pub length_words: f64,
}

/// 时长图表数据项
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LengthChartItem {
    /// 行号
    pub line: usize,
    /// 播放时间（秒）
    pub play_time_sec: f64,
    /// 场景名称
    pub scene: String,
    /// 累计时长
    pub length: f64,
}

/// 场景数据项
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SceneItem {
    /// 开始行号
    pub line: usize,
    /// 结束行号
    pub endline: usize,
    /// 场景标题
    pub scene: String,
    /// 场景类型
    #[serde(rename = "type")]
    pub scene_type: String,
    /// 时间
    pub time: String,
}

/// 按属性分类的时长
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DurationByProp {
    /// 属性名
    pub prop: String,
    /// 时长（秒）
    pub duration: f64,
}

/// 计算统计数据
pub fn calculate_statistics(
    tokens: &[crate::models::ScriptToken],
    properties: &crate::models::ScreenplayProperties,
    length_dialogue: f64,
    length_action: f64,
    dial_sec_per_char: f64,
    dial_sec_per_punc_short: f64,
    dial_sec_per_punc_long: f64,
) -> Statistics {
    Statistics {
        character_stats: calculate_character_statistics(
            tokens,
            properties,
            dial_sec_per_char,
            dial_sec_per_punc_short,
            dial_sec_per_punc_long,
        ),
        location_stats: calculate_location_statistics(properties),
        scene_stats: calculate_scene_statistics(properties),
        duration_stats: calculate_duration_statistics(
            tokens,
            properties,
            length_dialogue,
            length_action,
        ),
    }
}

/// 计算角色统计
fn calculate_character_statistics(
    tokens: &[crate::models::ScriptToken],
    properties: &crate::models::ScreenplayProperties,
    dial_sec_per_char: f64,
    dial_sec_per_punc_short: f64,
    dial_sec_per_punc_long: f64,
) -> CharacterStatistics {
    let mut dialogue_per_character: HashMap<String, Vec<String>> = HashMap::new();
    // 记录每个角色的动作时长
    let mut action_time_per_character: HashMap<String, f64> = HashMap::new();
    let mut first_scene_started = false;

    let mut i = 0;
    while i < tokens.len() {
        if tokens[i].token_type == "scene_heading" {
            first_scene_started = true;
        }

        if tokens[i].token_type == "character" {
            let character = tokens[i]
                .character
                .as_ref()
                .map(|s| s.replace(" (CONT'D)", "").replace(" (V.O.)", "").trim().to_string())
                .unwrap_or_else(|| "UNKNOWN".to_string());

            let mut speech = String::new();
            i += 1;

            while i < tokens.len() {
                if tokens[i].token_type == "dialogue" {
                    if first_scene_started {
                        let text = tokens[i].text_no_notes.as_deref().unwrap_or(&tokens[i].text);
                        speech.push_str(text);
                        speech.push(' ');
                    }
                    i += 1;
                } else if tokens[i].token_type == "character" {
                    break;
                } else if tokens[i].token_type == "action" && first_scene_started {
                    // 内层循环中也处理 action 的角色动作时间
                    if let Some(chars) = &tokens[i].characters_action {
                        if let Some(time) = tokens[i].time {
                            for char_name in chars {
                                *action_time_per_character.entry(char_name.clone()).or_insert(0.0) += time;
                            }
                        }
                    }
                    i += 1;
                } else if tokens[i].token_type == "scene_heading" {
                    first_scene_started = true;
                    i += 1;
                } else {
                    i += 1;
                }
            }

            let speech = speech.trim().to_string();
            if !speech.is_empty() {
                dialogue_per_character
                    .entry(character.clone())
                    .or_insert_with(Vec::new)
                    .push(speech);
            }
        } else if tokens[i].token_type == "action" {
            // 统计角色动作时间
            if let Some(characters_action) = &tokens[i].characters_action {
                if let Some(time) = tokens[i].time {
                    if !characters_action.is_empty() {
                        for char_name in characters_action {
                            let entry = action_time_per_character.entry(char_name.clone()).or_insert(0.0);
                            *entry += time;
                        }
                    }
                }
            }
            i += 1;
        } else {
            i += 1;
        }
    }

    let mut character_stats: Vec<CharacterStat> = Vec::new();
    let mut monologue_counter = 0;

    for (character_name, speeches) in dialogue_per_character.iter() {
        let speaking_parts = speeches.iter().filter(|s| !s.is_empty()).count();

        let mut seconds_spoken = 0.0;
        let mut monologues = 0;
        let mut all_dialogue_combined = String::new();

        for speech in speeches {
            let time = calculate_dialogue_duration(
                speech,
                dial_sec_per_char,
                dial_sec_per_punc_short,
                dial_sec_per_punc_long,
            );
            seconds_spoken += time;
            all_dialogue_combined.push_str(speech);
            all_dialogue_combined.push(' ');

            if time > 30.0 {
                monologues += 1;
            }
        }

        monologue_counter += monologues;

        let words_spoken = count_words(&all_dialogue_combined);

        // 获取角色动作时间
        let seconds_action = action_time_per_character.get(character_name).copied().unwrap_or(0.0);
        // 总时长 = 对白时长 + 动作时长
        let seconds_total = seconds_spoken + seconds_action;

        let number_of_scenes = properties
            .character_scene_number
            .as_ref()
            .and_then(|map| map.get(character_name))
            .map(|set| set.len())
            .unwrap_or(0);

        character_stats.push(CharacterStat {
            name: character_name.clone(),
            color: word_to_hex_color(character_name),
            speaking_parts,
            seconds_spoken,
            seconds_total,
            average_complexity: 0.0,
            monologues,
            words_spoken,
            number_of_scenes,
        });
    }

    // 补充只有场景出现但无对白的角色（如人设中定义的角色）
    for char_name in properties.characters.keys() {
        if !dialogue_per_character.contains_key(char_name) {
            let seconds_action = action_time_per_character.get(char_name).copied().unwrap_or(0.0);
            let number_of_scenes = properties
                .character_scene_number
                .as_ref()
                .and_then(|map| map.get(char_name))
                .map(|set| set.len())
                .unwrap_or(0);
            character_stats.push(CharacterStat {
                name: char_name.clone(),
                color: word_to_hex_color(char_name),
                speaking_parts: 0,
                seconds_spoken: 0.0,
                seconds_total: seconds_action,
                average_complexity: 0.0,
                monologues: 0,
                words_spoken: 0,
                number_of_scenes,
            });
        }
    }

    character_stats.sort_by(|a, b| {
        if b.speaking_parts != a.speaking_parts {
            b.speaking_parts.cmp(&a.speaking_parts)
        } else {
            b.words_spoken.cmp(&a.words_spoken)
        }
    });

    let character_count = character_stats.len();
    
    CharacterStatistics {
        characters: character_stats,
        complexity: 0.0,
        character_count,
        monologues: monologue_counter,
    }
}

/// 计算地点统计
fn calculate_location_statistics(
    properties: &crate::models::ScreenplayProperties,
) -> LocationStatistics {
    let mut locations: Vec<LocationStat> = Vec::new();

    for (location_slug, references) in properties.locations.iter() {
        let times_of_day = references
            .iter()
            .map(|loc| normalize_time(&loc.time_of_day))
            .collect::<HashSet<_>>()
            .into_iter()
            .collect::<Vec<_>>();

        let has_both = references.iter().any(|r| r.interior && r.exterior);
        let has_interior = references.iter().any(|r| r.interior && !r.exterior);
        let has_exterior = references.iter().any(|r| r.exterior && !r.interior);

        let mut interior_exterior = "unknown".to_string();
        let mut count = 0;

        if has_both {
            count += 1;
            interior_exterior = "ie".to_string();
        }
        if has_interior {
            count += 1;
            interior_exterior = "int".to_string();
        }
        if has_exterior {
            count += 1;
            interior_exterior = "ext".to_string();
        }
        if count > 1 {
            interior_exterior = "multiple".to_string();
        }

        let unique_scene_numbers: Vec<String> = references
            .iter()
            .map(|r| r.scene_number.clone())
            .filter(|s| !s.is_empty())
            .collect::<HashSet<_>>()
            .into_iter()
            .collect();

        let number_of_scenes = unique_scene_numbers.len();

        locations.push(LocationStat {
            name: location_slug.clone(),
            color: word_to_hex_color(location_slug),
            scene_numbers: references.iter().map(|r| r.scene_number.clone()).collect(),
            scene_lines: references.iter().map(|r| r.start_play_sec).collect(),
            number_of_scenes,
            times_of_day,
            interior_exterior,
        });
    }

    LocationStatistics {
        locations_count: locations.len(),
        locations,
    }
}

/// 计算场景统计
fn calculate_scene_statistics(
    properties: &crate::models::ScreenplayProperties,
) -> SceneStatistics {
    let scene_numbers: Vec<String> = properties
        .scenes
        .iter()
        .filter_map(|scene| {
            scene
                .get("number")
                .and_then(|v| v.as_str())
                .map(|s| s.to_string())
                .filter(|s| !s.trim().is_empty())
        })
        .collect::<HashSet<_>>()
        .into_iter()
        .collect();

    SceneStatistics {
        scenes: scene_numbers
            .into_iter()
            .map(|n| SceneStat { title: n })
            .collect(),
    }
}

/// 插入角色时间线数据点，跨度大时插入凹凸点（参考 VSCode 实现）
fn insert_character_timeline_point(
    timeline: &mut Vec<CharacterDialogueItem>,
    element: &crate::models::ScriptToken,
    current_scene: &str,
    time: f64,
    gap: f64,
    is_monologue: bool,
    word_count: f64,
) {
    let prev_global = timeline.last().map(|x| x.length_time_global).unwrap_or(0.0);
    let prev_words = timeline.last().map(|x| x.length_words_global).unwrap_or(0.0);
    let insert_pre = if let Some(last) = timeline.last() {
        last.play_time_sec + gap < element.play_time_sec - time
    } else {
        true // 第一个元素也要插入前置零点
    };

    if insert_pre {
        // 前一个点之后立即归零
        if let Some(last) = timeline.last() {
            timeline.push(CharacterDialogueItem {
                line: last.line + 1,
                play_time_sec: last.play_time_sec + 1.0,
                scene: current_scene.to_string(),
                length_time_global: 0.0,
                length_words_global: 0.0,
                monologue: false,
                length_time: 0.0,
                length_words: 0.0,
            });
        }
        // 当前点前两个时间单位，归零
        timeline.push(CharacterDialogueItem {
            line: element.line.saturating_sub(2),
            play_time_sec: element.play_time_sec - time - 1.0,
            scene: current_scene.to_string(),
            length_time_global: 0.0,
            length_words_global: 0.0,
            monologue: false,
            length_time: 0.0,
            length_words: 0.0,
        });
        // 当前点前一个时间单位，恢复到之前累计值（起桥）
        timeline.push(CharacterDialogueItem {
            line: element.line.saturating_sub(1),
            play_time_sec: element.play_time_sec - time,
            scene: current_scene.to_string(),
            length_time_global: prev_global,
            length_words_global: prev_words,
            monologue: false,
            length_time: 0.0,
            length_words: 0.0,
        });
    }

    // 当前数据点
    timeline.push(CharacterDialogueItem {
        line: element.line,
        play_time_sec: element.play_time_sec,
        scene: current_scene.to_string(),
        length_time_global: prev_global + time,
        length_words_global: prev_words + word_count,
        monologue: is_monologue,
        length_time: time,
        length_words: word_count,
    });
}

/// 计算时长统计
fn calculate_duration_statistics(
    tokens: &[crate::models::ScriptToken],
    properties: &crate::models::ScreenplayProperties,
    length_dialogue: f64,
    length_action: f64,
) -> DurationStatistics {
    let mut action_chart: Vec<LengthChartItem> = vec![LengthChartItem {
        line: 0,
        length: 0.0,
        scene: String::new(),
        play_time_sec: 0.0,
    }];
    let mut dialogue_chart: Vec<LengthChartItem> = vec![LengthChartItem {
        line: 0,
        length: 0.0,
        scene: String::new(),
        play_time_sec: 0.0,
    }];
    let mut scenes: Vec<SceneItem> = Vec::new();
    let mut duration_by_prop: Vec<DurationByProp> = Vec::new();
    let mut previous_length_action = 0.0;
    let mut previous_length_dialogue = 0.0;
    let mut current_scene = String::new();
    let mut monologues = 0;
    let mut scene_prop_durations: HashMap<String, f64> = HashMap::new();
    
    // 角色时间线数据
    let mut characters_timeline: HashMap<String, Vec<CharacterDialogueItem>> = HashMap::new();
    
    let gap = if length_action + length_dialogue > 380.0 {
        (length_action + length_dialogue) / 10.0
    } else {
        40.0
    };

    for element in tokens.iter() {
        if element.token_type == "action" || element.token_type == "dialogue" {
            let time = element.time.unwrap_or(0.0);

            if element.token_type == "action" {
                previous_length_action += time;
                action_chart.push(LengthChartItem {
                    line: element.line,
                    length: previous_length_action,
                    scene: current_scene.clone(),
                    play_time_sec: element.play_time_sec,
                });
                
                // 角色动作时间线
                if let Some(characters_action) = &element.characters_action {
                    if !characters_action.is_empty() && time > 0.0 {
                        for char_name in characters_action {
                            let char_timeline = characters_timeline.entry(char_name.clone()).or_insert_with(Vec::new);
                            insert_character_timeline_point(char_timeline, element, &current_scene, time, gap, false, 0.0);
                        }
                    }
                }
            } else if element.token_type == "dialogue" {
                previous_length_dialogue += time;
                dialogue_chart.push(LengthChartItem {
                    line: element.line,
                    length: previous_length_dialogue,
                    scene: current_scene.clone(),
                    play_time_sec: element.play_time_sec,
                });

                if time > 30.0 {
                    monologues += 1;
                }
                
                // 角色对话时间线
                if let Some(char_name) = &element.character {
                    let char_timeline = characters_timeline.entry(char_name.clone()).or_insert_with(Vec::new);
                    let word_count = element.text.split_whitespace().count() as f64;
                    let is_monologue = time > 30.0;
                    insert_character_timeline_point(char_timeline, element, &current_scene, time, gap, is_monologue, word_count);
                }
            }
        }
    }

    for scene in properties.scenes.iter() {
        if let Some(text) = scene.get("text").and_then(|v| v.as_str()) {
            current_scene = text.to_string();

            let (scene_type, scene_time) = parse_scene_heading(text);

            let start_play_sec = scene
                .get("startPlaySec")
                .and_then(|v| v.as_f64())
                .unwrap_or(0.0) as usize;
            let end_play_sec = scene
                .get("endPlaySec")
                .and_then(|v| v.as_f64())
                .unwrap_or(0.0) as usize;
            let action_length = scene
                .get("actionLength")
                .and_then(|v| v.as_f64())
                .unwrap_or(0.0);
            let dialogue_length = scene
                .get("dialogueLength")
                .and_then(|v| v.as_f64())
                .unwrap_or(0.0);

            scenes.push(SceneItem {
                line: start_play_sec,
                endline: end_play_sec,
                scene: text.to_string(),
                scene_type: scene_type.clone(),
                time: scene_time.clone(),
            });

            let type_key = format!("type_{}", scene_type);
            let current_type_duration = scene_prop_durations.get(&type_key).unwrap_or(&0.0);
            scene_prop_durations.insert(
                type_key.clone(),
                current_type_duration + action_length + dialogue_length,
            );

            let time_key = format!("time_{}", scene_time);
            let current_time_duration = scene_prop_durations.get(&time_key).unwrap_or(&0.0);
            scene_prop_durations.insert(
                time_key.clone(),
                current_time_duration + action_length + dialogue_length,
            );
        }
    }

    for (prop, duration) in scene_prop_durations.into_iter() {
        duration_by_prop.push(DurationByProp { prop, duration });
    }
    
    // 转换为角色时间线数据
    let mut characternames: Vec<String> = Vec::new();
    let mut characters: Vec<Vec<CharacterDialogueItem>> = Vec::new();
    for (name, mut timeline) in characters_timeline.into_iter() {
        timeline.sort_by(|a, b| a.play_time_sec.partial_cmp(&b.play_time_sec).unwrap_or(std::cmp::Ordering::Equal));
        characternames.push(name);
        characters.push(timeline);
    }

    DurationStatistics {
        dialogue: length_dialogue,
        action: length_action,
        total: length_dialogue + length_action,
        lengthchart_action: action_chart,
        lengthchart_dialogue: dialogue_chart,
        duration_by_scene_prop: duration_by_prop,
        scenes,
        monologues,
        characternames,
        characters,
    }
}

/// 计算对话时长
fn calculate_dialogue_duration(
    dialogue: &str,
    dial_sec_per_char: f64,
    dial_sec_per_punc_short: f64,
    dial_sec_per_punc_long: f64,
) -> f64 {
    use regex::Regex;

    let mut duration = 0.0;

    let sanitized = Regex::new(r"\s|\p{P}|\p{S}")
        .unwrap()
        .replace_all(dialogue, "");
    duration += sanitized.chars().count() as f64 * dial_sec_per_char;

    let rec = Regex::new(r"(\.|\?|\!|\:|。|？|！|：)|(\,|，|;|；|、)").unwrap();
    for cap in rec.captures_iter(dialogue) {
        if cap.get(1).is_some() {
            duration += dial_sec_per_punc_long;
        }
        if cap.get(2).is_some() {
            duration += dial_sec_per_punc_short;
        }
    }

    duration
}

/// 统计单词数
fn count_words(text: &str) -> usize {
    text.split_whitespace().count()
}

/// 字符串转颜色（十六进制）
fn word_to_hex_color(word: &str) -> String {
    let rgb = word_to_rgb_color(word);
    format!(
        "#{:02x}{:02x}{:02x}",
        rgb[0] as u8, rgb[1] as u8, rgb[2] as u8
    )
}

/// 字符串转 RGB 颜色
fn word_to_rgb_color(word: &str) -> [f64; 3] {
    let s = 0.7;
    let v = 0.7;
    let hash = hash_string(word);

    let s_diff = (1.0 - s) * 100.0;
    let s = if s_diff > 0.0 {
        ((hash * 19) % s_diff as i64) as f64 / 100.0 + s
    } else {
        s
    };

    let v_diff = (1.0 - v) * 100.0;
    let v = if v_diff > 0.0 {
        ((hash.wrapping_mul(11)) % v_diff as i64) as f64 / 100.0 + v
    } else {
        v
    };

    let h = ((hash.wrapping_mul(157)) % 360) as f64 / 360.0;

    hsv_to_rgb(h, s, v)
}

/// 字符串哈希
fn hash_string(str: &str) -> i64 {
    let mut hash: i64 = 0;
    let mut ls: Vec<i64> = Vec::new();

    for ch in str.chars() {
        let code_point = ch as i64;
        let mod_val = code_point % 256;
        let hash_temp = ((hash << 5) - hash + mod_val) | 0;

        if hash_temp < 0 {
            ls.push(hash);
            hash = mod_val;
        } else {
            hash = hash_temp;
        }
    }

    if !ls.is_empty() {
        let tot = (ls.len() + 1) as f64;
        hash = (hash as f64 / tot) as i64;
        for h in ls.iter() {
            hash += (*h as f64 / tot) as i64;
        }
    }

    if hash < 0 {
        hash = -hash;
    }
    hash
}

/// HSV 转 RGB
fn hsv_to_rgb(h: f64, s: f64, v: f64) -> [f64; 3] {
    let i = (h * 6.0).floor() as i32;
    let f = h * 6.0 - i as f64;
    let p = v * (1.0 - s);
    let q = v * (1.0 - f * s);
    let t = v * (1.0 - (1.0 - f) * s);

    let (r, g, b) = match i % 6 {
        0 => (v, t, p),
        1 => (q, v, p),
        2 => (p, v, t),
        3 => (p, q, v),
        4 => (t, p, v),
        5 => (v, p, q),
        _ => (v, t, p),
    };

    [
        (r * 255.0).round(),
        (g * 255.0).round(),
        (b * 255.0).round(),
    ]
}

/// 标准化时间描述（支持中英文）
fn normalize_time(val: &str) -> String {
    if val.is_empty() {
        return "unspecified".to_string();
    }

    let mut time = val.to_lowercase();
    time = time.replace("  ", " ");
    time = time.trim_end_matches('.').to_string();
    time = time.trim().to_string();

    let re_the = regex::Regex::new(r"^(the)?\s*(next|following)\b").unwrap();
    time = re_the.replace(&time, "").to_string();

    let re_early_late = regex::Regex::new(r"^(early|late)\b").unwrap();
    time = re_early_late.replace(&time, "").to_string();

    time = time.trim().to_string();

    if time.is_empty() {
        return "unspecified".to_string();
    }

    // 中文时间词映射
    let time_lower = time.to_lowercase();
    if time_lower.contains("正午") || time_lower.contains("上午") || time_lower.contains("午后") 
        || time_lower.contains("下午") || time_lower.contains("日") || time_lower.contains("白天") 
        || time_lower.contains("day") || time_lower.contains("中午") {
        return "day".to_string();
    }
    if time_lower.contains("夜") || time_lower.contains("深夜") || time_lower.contains("子夜") 
        || time_lower.contains("午夜") || time_lower.contains("夜晚") || time_lower.contains("晚上") 
        || time_lower.contains("night") || time_lower.contains("黑夜") || time_lower.contains("凌晨"){
        return "night".to_string();
    }
    if time_lower.contains("傍晚") || time_lower.contains("黄昏") || time_lower.contains("dusk") 
        || time_lower.contains("evening") || time_lower.contains("暮") {
        return "dusk".to_string();
    }
    if time_lower.contains("拂晓") || time_lower.contains("黎明") || time_lower.contains("dawn") {
        return "dawn".to_string();
    }
    if time_lower.contains("清晨") || time_lower.contains("早晨") || time_lower.contains("清早") 
        || time_lower.contains("早上") || time_lower.contains("morning") || time_lower.contains("早") {
        return "morning".to_string();
    }

    time
}

/// 解析场景标题（支持中英文）
fn parse_scene_heading(text: &str) -> (String, String) {
    let mut scene_type = "unknown".to_string();
    let mut scene_time = "unspecified".to_string();

    let text_trimmed = text.trim();

    // 首先检查中文场景类型：直接以括号开头的场景（如 "(内景)"、"（内景）"）
    if text_trimmed.starts_with("(内景)") || text_trimmed.starts_with("（内景）") {
        scene_type = "int".to_string();
    } else if text_trimmed.starts_with("(外景)") || text_trimmed.starts_with("（外景）") {
        scene_type = "ext".to_string();
    } else if text_trimmed.starts_with("(内外景)") || text_trimmed.starts_with("（内外景）") {
        scene_type = "ie".to_string();
    } else {
        // 尝试英文正则匹配
        if let Ok(re) = regex::Regex::new(r"^(?:\* *)?(?:((?:INT|EXT|I|E)\.?(?:\/(?:INT|EXT|I|E)\.?)?)[ ]+)?(.+?)(?:[-–—−](.+))?$") {
            if let Some(caps) = re.captures(text) {
                if let Some(type_part) = caps.get(1) {
                    scene_type = location_type(type_part.as_str());
                }

                if let Some(time_part) = caps.get(3) {
                    let time_str = after_dash(time_part.as_str());
                    if let Some(t) = time_str {
                        scene_time = normalize_time(&t);
                    }
                }
            }
        }
    }

    // 如果场景类型已确定，尝试从场景文本中提取时间
    if scene_time == "unspecified" {
        scene_time = extract_time_from_scene_text(text);
    } else {
        // 标准化时间分类
        let time_categories = ["day", "night", "dawn", "dusk", "morning", "evening"];
        let time_lower = scene_time.to_lowercase();

        for category in time_categories.iter() {
            if time_lower.contains(category) {
                scene_time = category.to_string();
                break;
            }
        }
    }

    (scene_type, scene_time)
}

/// 从场景文本中提取时间信息（支持中英文）
fn extract_time_from_scene_text(text: &str) -> String {
    let text_lower = text.to_lowercase();
    
    // 检查中文时间词
    if text_lower.contains("正午") || text_lower.contains("上午") || text_lower.contains("午后") 
        || text_lower.contains("下午") || text_lower.contains("日") || text_lower.contains("白天") 
        || text_lower.contains("day") || text_lower.contains("中午") {
        return "day".to_string();
    }
    if text_lower.contains("夜") || text_lower.contains("深夜") || text_lower.contains("子夜") 
        || text_lower.contains("午夜") || text_lower.contains("夜晚") || text_lower.contains("晚上") 
        || text_lower.contains("night") || text_lower.contains("黑夜") 
        || text_lower.contains("凌晨") {
        return "night".to_string();
    }
    if text_lower.contains("傍晚") || text_lower.contains("黄昏") || text_lower.contains("dusk") 
        || text_lower.contains("evening") || text_lower.contains("暮") {
        return "dusk".to_string();
    }
    if text_lower.contains("拂晓") || text_lower.contains("黎明") || text_lower.contains("dawn") {
        return "dawn".to_string();
    }
    if text_lower.contains("清晨") || text_lower.contains("早晨") || text_lower.contains("清早")
        || text_lower.contains("早上") || text_lower.contains("morning") || text_lower.contains("早") {
        return "morning".to_string();
    }
    
    // 英文时间词
    let time_categories = ["day", "night", "dawn", "dusk", "morning", "evening"];
    for category in time_categories.iter() {
        if text_lower.contains(category) {
            return category.to_string();
        }
    }
    
    "unspecified".to_string()
}

/// 判断场景类型（支持中英文）
fn location_type(val: &str) -> String {
    if val.is_empty() {
        return "unknown".to_string();
    }

    let val_lower = val.to_lowercase();

    if regex::Regex::new(r"i(nt)?\.?/e(xt)?\.?")
        .unwrap()
        .is_match(&val_lower)
    {
        return "ie".to_string();
    } else if regex::Regex::new(r"i(nt)?\.?")
        .unwrap()
        .is_match(&val_lower)
    {
        return "int".to_string();
    } else if regex::Regex::new(r"e(xt)?\.?")
        .unwrap()
        .is_match(&val_lower)
    {
        return "ext".to_string();
    }

    "unknown".to_string()
}

/// 提取破折号后的时间信息
fn after_dash(val: &str) -> Option<String> {
    if val.is_empty() {
        return None;
    }

    let dashes = ['-', '–', '—', '−'];
    for dash in dashes.iter() {
        if let Some(idx) = val.find(*dash) {
            let n = val[idx + 1..].trim();
            if !n.is_empty() {
                return Some(n.to_string());
            }
        }
    }

    None
}
