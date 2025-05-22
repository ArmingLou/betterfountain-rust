use std::collections::HashMap;
use lazy_static::lazy_static;
use regex::Regex;

pub struct FountainConstants;

impl FountainConstants {
    // 样式标记字符映射
    pub fn style_chars() -> HashMap<&'static str, &'static str> {
        let mut map = HashMap::new();
        map.insert("note_begin_ext", "இ");
        map.insert("note_begin", "↺");
        map.insert("note_end", "↻");
        map.insert("italic", "☈");
        map.insert("bold", "↭");
        map.insert("bold_italic", "↯");
        map.insert("underline", "☄");
        map.insert("italic_underline", "⇀");
        map.insert("bold_underline", "☍");
        map.insert("bold_italic_underline", "☋");
        map.insert("link", "𓆡");
        map.insert("style_left_stash", "↷");
        map.insert("style_left_pop", "↶");
        map.insert("style_right_stash", "↝");
        map.insert("style_right_pop", "↜");
        map.insert("style_global_stash", "↬");
        map.insert("style_global_pop", "↫");
        map.insert("style_global_clean", "⇜");
        map.insert("italic_global_begin", "↾");
        map.insert("italic_global_end", "↿");
        map.insert("all", "☄☈↭↯↺↻↬↫☍☋↷↶↾↿↝↜⇀𓆡⇜இ");
        map
    }
}

lazy_static! {
    // 块级元素正则
    pub static ref BLOCK_REGEX: HashMap<&'static str, Regex> = {
        let mut map = HashMap::new();
        map.insert("block_dialogue_begin", Regex::new(r"^[ \t]*((\p{Lu}[^\p{Ll}\r\n@]*)|(@[^\r\n\(（\^]*))(\(.*\)|（.*）)?(\s*\^)?\s*$").unwrap());
        map.insert("block_except_dialogue_begin", Regex::new(r"^\s*[^\s]+.*$").unwrap());
        map.insert("block_end", Regex::new(r"^\s*$").unwrap());
        map.insert("line_break", Regex::new(r"^\s{2,}$").unwrap());
        map.insert("action_force", Regex::new(r"^(\s*)(\!)(.*)").unwrap());
        map.insert("lyric", Regex::new(r"^(\s*)(\~)(\s*)(.*)").unwrap());
        map
    };

    // Token解析正则
    pub static ref TOKEN_REGEX: HashMap<&'static str, Regex> = {
        let mut map = HashMap::new();
        map.insert("note_inline", Regex::new(r"(?:↺|இ)([\s\S]+?)(?:↻)").unwrap());
        map.insert("underline", Regex::new(r"(☄(?=.+☄))(.+?)(☄)").unwrap());
        map.insert("italic", Regex::new(r"(☈(?=.+☈))(.+?)(☈)").unwrap());
        map.insert("italic_global", Regex::new(r"(↾)([^↿]*)(↿)").unwrap());
        map.insert("bold", Regex::new(r"(↭(?=.+↭))(.+?)(↭)").unwrap());
        map.insert("bold_italic", Regex::new(r"(↯(?=.+↯))(.+?)(↯)").unwrap());
        map.insert("italic_underline", Regex::new(r"(?:☄☈(?=.+☈☄)|☈☄(?=.+☄☈))(.+?)(☈☄|☄☈)").unwrap());
        map.insert("bold_italic_underline", Regex::new(r"(☄↯(?=.+↯☄)|↯☄(?=.+☄↯))(.+?)(↯☄|☄↯)").unwrap());
        map.insert("bold_underline", Regex::new(r"(☄↭(?=.+↭☄)|↭☄(?=.+☄↭))(.+?)(↭☄|☄↭)").unwrap());
        map
    };
}
