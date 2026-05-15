use std::collections::BTreeMap;

use rhwp::wasm_api::HwpDocument;
use serde_json::{json, Value};

use crate::ops::write_document;
use crate::options::{optional_usize, required};

pub(crate) fn insert_text(
    doc: &mut HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let output = required(options, "--output")?;
    let value = required(options, "--value")?;
    let section = optional_usize(options, "--section")?.unwrap_or(0);
    let paragraph = optional_usize(options, "--paragraph")?
        .or(optional_usize(options, "--para")?)
        .unwrap_or(0);
    let offset = optional_usize(options, "--offset")?.unwrap_or(0);

    let insert_json = doc
        .insert_text_native(section, paragraph, offset, value)
        .map_err(|e| format!("text insert failed: {e}"))?;
    write_document(doc, format, output)?;

    let insert_result: Value =
        serde_json::from_str(&insert_json).unwrap_or_else(|_| json!({ "raw": insert_json }));
    println!(
        "{}",
        json!({
            "inserted": true,
            "operation": "insert-text",
            "output": output,
            "format": format,
            "section": section,
            "paragraph": paragraph,
            "offset": offset,
            "insert": insert_result,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "warnings": ["experimental text insertion; verify with Hancom before production use"]
        })
    );
    Ok(())
}
