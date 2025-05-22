use serde::{Deserialize, Serialize};
use crate::docx::docx_maker::PrintProfile;

/// 页面边距
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Margins {
    pub top: f32,
    pub bottom: f32,
    pub left: f32,
    pub right: f32,
}

impl Default for Margins {
    fn default() -> Self {
        Margins {
            top: 1.0,
            bottom: 1.0,
            left: 1.5,
            right: 1.0,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Conf {
    pub print_notes: bool,
    pub merge_empty_lines: bool,
    pub each_scene_on_new_page: bool,
    pub use_dual_dialogue: bool,
    pub dialogue_foldable: bool,
    pub print_dialogue_numbers: bool,
    pub emitalic_dialog: bool,
    pub embolden_character_names: bool,
    pub text_contd: String,
    /// 对白中每字符耗时预估(不含标点)
    pub dial_sec_per_char: f64,
    /// 对白中每个短标点耗时预估(逗号顿号等)
    pub dial_sec_per_punc_short: f64,
    /// 对白中每个长标点耗时预估(句号问号等)
    pub dial_sec_per_punc_long: f64,
    /// action文本中每字符转化成影片时长预估(不含标点)
    pub action_sec_per_char: f64,
    /// 是否打印标题页
    pub print_title_page: bool,
    /// 是否打印前言页
    pub print_preface_page: bool,
    /// 场景编号位置
    pub scenes_numbers: String,
    /// 是否显示页码
    pub show_page_numbers: String,
    /// 是否加粗场景标题
    pub embolden_scene_headers: bool,
    /// 是否为场景标题添加下划线
    pub underline_scene_headers: bool,
    /// 页眉
    pub print_header: String,
    /// 页脚
    pub print_footer: String,
    /// 是否为章节添加编号
    pub number_sections: bool,
    /// 是否创建书签
    pub create_bookmarks: bool,
    /// 注释位置是否在底部
    pub note_position_bottom: bool,
    /// 是否打印动作
    pub print_actions: bool,
    /// 是否打印对话
    pub print_dialogues: bool,
    /// 是否打印场景标题
    pub print_headers: bool,
    /// 是否打印章节
    pub print_sections: bool,
    /// 是否打印概要
    pub print_synopsis: bool,
    /// 是否在场景之间添加双倍空格
    pub double_space_between_scenes: bool,
    /// 页面大小
    pub page_size: String,
    /// 字体名称
    pub font_family: String,
    /// 是否为对话添加斜体
    pub emitalic_dialogue: bool,
    /// 打印配置
    pub print_profile: PrintProfile,
    /// 水印
    pub print_watermark: String,
}

impl Default for Conf {
    fn default() -> Self {
        Conf {
            print_notes: true,
            merge_empty_lines: true,
            each_scene_on_new_page: false,
            use_dual_dialogue: true,
            dialogue_foldable: false,
            print_dialogue_numbers: false,
            emitalic_dialog: true,
            embolden_character_names: true,
            text_contd: "(CONT'D)".to_string(),
            dial_sec_per_char: 0.3,
            dial_sec_per_punc_short: 0.3,
            dial_sec_per_punc_long: 0.75,
            action_sec_per_char: 0.4,
            print_title_page: true,
            print_preface_page: true,
            scenes_numbers: "both".to_string(),
            show_page_numbers: "(第{n}页)".to_string(),
            embolden_scene_headers: true,
            underline_scene_headers: false,
            print_header: "".to_string(),
            print_footer: "".to_string(),
            number_sections: true,
            create_bookmarks: true,
            note_position_bottom: true,
            print_actions: true,
            print_dialogues: true,
            print_headers: true,
            print_sections: false,
            print_synopsis: true,
            double_space_between_scenes: false,
            page_size: "A4".to_string(),
            font_family: "Courier Prime".to_string(),
            emitalic_dialogue: false,
            print_profile: PrintProfile::default(),
            print_watermark: "".to_string(),
        }
    }
}
