pub mod docx_maker;
pub mod docx_adapter;
pub mod docx;
pub mod adapter;
pub mod line_processor;
pub mod metadata_extractor;

// 从 docx_maker 导出
pub use docx_maker::{
    generate_docx, DocxOptions, DocxResult, PrintProfile,
    DocxContext, CurrentNote
};

// 从 docx_adapter 导出
pub use docx_adapter::{DocxAdapter, DocxAdapterError, DocxAdapterResult, TextStyle, ParagraphStyle, FormatState};

// 从 docx 导出
pub use docx::{
    DocxGenerateError,
    DocxGenerateResult,
    generate_docx_document
};
pub use docx_maker::ExportConfig;

// 从 adapter 导出
pub use adapter::{
    DocxAdapterError as AdapterError,
    DocxAdapterResult as AdapterResult,
    DocxStats, DocxAsBase64, LineStruct,
    UnderlineType, AlignmentType, BreakType, FontType,
    LineRuleType, FrameAnchorType, HorizontalPositionAlign, VerticalPositionAlign, WidthType,
    UnderlineTypeConst, AlignmentTypeConst, BreakTypeConst, FontTypeConst,
    LineRuleTypeConst, FrameAnchorTypeConst, HorizontalPositionAlignConst, VerticalPositionAlignConst, WidthTypeConst,
    PageNumber, convert_inches_to_twip, convertInchesToTwip
};

// 从 adapter::docx 导出
pub use adapter::docx::{
    Document, Paragraph, TextRun, BreakRun, RunTrait
};

// 从 metadata_extractor 导出
pub use metadata_extractor::{
    ExtractedMetadata, extract_metadata_from_parsed_document
};
