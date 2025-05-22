use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Synopsis {
    pub synopsis: String,
    pub line: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Note {
    pub note: String,
    pub line: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Position {
    pub line: usize,
    pub character: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Range {
    pub start: Position,
    pub end: Position,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StructToken {
    pub text: String,
    pub isnote: bool,
    pub id: Option<String>,
    pub children: Vec<StructToken>,
    pub range: Option<Range>,
    pub level: usize,
    pub section: bool,
    pub synopses: Vec<Synopsis>,
    pub notes: Vec<Note>,
    pub isscene: bool,
    pub ischartor: bool,
    pub dialogue_end_line: usize,
    pub duration_sec: f64,
    pub play_sec: f64,
    pub structs: Vec<StructToken>,
    pub duration: f64,
}

impl StructToken {
    pub fn new(
        text: String,
        level: usize,
        section: bool,
        isscene: bool,
        ischartor: bool,
    ) -> Self {
        StructToken {
            text,
            isnote: false,
            id: None,
            children: Vec::new(),
            range: None,
            level,
            section,
            synopses: Vec::new(),
            notes: Vec::new(),
            isscene,
            ischartor,
            dialogue_end_line: 0,
            duration_sec: 0.0,
            play_sec: 0.0,
            structs: Vec::new(),
            duration: 0.0,
        }
    }
}
