use crate::parser::ParseOutput;
use crate::models::Conf;
use crate::pdf::liner::Liner;

/// 处理文档行
pub fn process_document_lines(parsed_document: &mut ParseOutput, config: &Conf) {
    // 如果已经有处理过的行，则不再处理
    if !parsed_document.lines.is_empty() {
        return;
    }

    // 创建行处理器
    let liner = Liner::new(config.print_dialogue_numbers);

    // 处理行
    let lines = liner.line2(&parsed_document.tokens, config);

    // 更新解析结果
    parsed_document.lines = lines;
}
