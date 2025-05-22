use betterfountain_rust::{Conf, parse};
use std::fs;
use std::env;

fn main() {
    let args: Vec<String> = env::args().collect();
    
    if args.len() < 2 {
        println!("Usage: {} <fountain_file>", args[0]);
        return;
    }
    
    let file_path = &args[1];
    
    match fs::read_to_string(file_path) {
        Ok(content) => {
            let config = Conf::default();
            let result = parse(&content, &config, true);
            
            println!("解析完成！");
            println!("解析时间: {}ms", result.parse_time);
            println!("Token数量: {}", result.tokens.len());
            println!("场景数量: {}", result.properties.scenes.len());
            println!("角色数量: {}", result.properties.characters.len());
            
            if let Some(html) = result.script_html {
                let html_path = format!("{}.html", file_path);
                fs::write(&html_path, html).unwrap();
                println!("HTML输出已保存到: {}", html_path);
            }
        },
        Err(e) => {
            println!("读取文件失败: {}", e);
        }
    }
}
