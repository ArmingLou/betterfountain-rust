use std::collections::{HashMap, HashSet};
use serde::{Deserialize, Serialize};
use crate::models::location::Location;
use crate::models::struct_token::StructToken;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScreenplayProperties {
    pub scenes: Vec<HashMap<String, serde_json::Value>>,
    pub scene_lines: Vec<usize>,
    pub scene_names: Vec<String>,
    pub title_keys: Vec<String>,
    pub font_line: i32,
    pub characters: HashMap<String, Vec<usize>>,
    pub locations: HashMap<String, Vec<Location>>,
    pub structure: Vec<StructToken>,
    pub length_action: usize,
    pub length_dialogue: usize,
    pub first_scene_line: Option<usize>,
    pub first_token_line: Option<usize>,
    pub character_lines: Option<HashMap<usize, String>>,
    pub character_first_line: Option<HashMap<String, usize>>,
    pub character_describe: Option<HashMap<String, String>>,
    pub character_scene_number: Option<HashMap<String, HashSet<String>>>,
    pub scene_number_vars: Option<HashSet<String>>,
}

impl ScreenplayProperties {
    pub fn new() -> Self {
        ScreenplayProperties {
            scenes: Vec::new(),
            scene_lines: Vec::new(),
            scene_names: Vec::new(),
            title_keys: Vec::new(),
            font_line: -1,
            characters: HashMap::new(),
            locations: HashMap::new(),
            structure: Vec::new(),
            length_action: 0,
            length_dialogue: 0,
            first_scene_line: None,
            first_token_line: None,
            character_lines: Some(HashMap::new()),
            character_first_line: Some(HashMap::new()),
            character_describe: Some(HashMap::new()),
            character_scene_number: Some(HashMap::new()),
            scene_number_vars: Some(HashSet::new()),
        }
    }
}

impl Default for ScreenplayProperties {
    fn default() -> Self {
        Self::new()
    }
}
