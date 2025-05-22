use regex::Regex;

fn main() {
    let scene_heading_regex = Regex::new(r"^[ \t]*([.][\w\(（\p{L}]|(?:int|ext|est|int[.]?\/ext|i[.]?\/e)[. ])([^#]*)(#\s*[^\s].*#)?\s*$").unwrap();
    
    let test_cases = vec![
        "INT. HOUSE - DAY #1#",
        "EXT. PARK - NIGHT",
        ".INT. HOUSE - DAY #1#",
        ".EXT. PARK - NIGHT",
        ".(内景) 顾清住处 - 日 #1#",
        ".（外景） 公园一角 - 傍晚",
    ];
    
    for test in test_cases {
        println!("Testing: '{}'", test);
        if scene_heading_regex.is_match(test) {
            println!("  ✓ Matches scene_heading");
            if let Some(captures) = scene_heading_regex.captures(test) {
                println!("  Groups: {:?}", captures.iter().map(|m| m.map(|m| m.as_str())).collect::<Vec<_>>());
            }
        } else {
            println!("  ✗ Does NOT match scene_heading");
        }
        println!();
    }
}
