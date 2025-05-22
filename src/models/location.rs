use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Location {
    pub name: String,
    pub interior: bool,
    pub exterior: bool,
    pub time_of_day: String,
    pub scene_number: String,
    pub line: usize,
    pub start_play_sec: f64,
}

impl Location {
    pub fn new(
        name: String, 
        interior: bool, 
        exterior: bool, 
        time_of_day: String
    ) -> Self {
        Location {
            name,
            interior,
            exterior,
            time_of_day,
            scene_number: String::new(),
            line: 0,
            start_play_sec: 0.0,
        }
    }
}
