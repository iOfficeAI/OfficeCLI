use std::collections::BTreeMap;

use rhwp::wasm_api::HwpDocument;
use serde_json::{json, Value};

use crate::ops::write_document;
use crate::ops_native_header_footer::try_run_native_header_footer_op;
use crate::ops_native_objects::try_run_native_object_op;
use crate::ops_native_style::try_run_native_style_op;
use crate::ops_native_table::try_run_native_table_op;
use crate::ops_native_text::try_run_native_text_op;
use crate::options::required;

type NativeOpRunner =
    fn(&mut HwpDocument, &str, &BTreeMap<String, String>) -> Result<Option<Value>, String>;

const NATIVE_OP_RUNNERS: [NativeOpRunner; 5] = [
    try_run_native_text_op,
    try_run_native_table_op,
    try_run_native_style_op,
    try_run_native_header_footer_op,
    try_run_native_object_op,
];

pub(crate) fn native_op(
    doc: &mut HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let op = required(options, "--op")?;
    let output = options.get("--output").map(String::as_str);
    let result = run_native_op(doc, op, options)?;

    if let Some(output_path) = output {
        let output_format = options
            .get("--output-format")
            .map(String::as_str)
            .unwrap_or(format);
        write_document(doc, output_format, output_path)?;
    }

    println!(
        "{}",
        json!({
            "operation": op,
            "result": result,
            "output": output,
            "outputFormat": output.map(|_| options.get("--output-format").map(String::as_str).unwrap_or(format)),
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": ["experimental native rhwp operation; use output mode and verify with readback/Hancom before production use"]
        })
    );
    Ok(())
}

fn run_native_op(
    doc: &mut HwpDocument,
    op: &str,
    options: &BTreeMap<String, String>,
) -> Result<Value, String> {
    for runner in NATIVE_OP_RUNNERS {
        if let Some(value) = runner(doc, op, options)? {
            return Ok(value);
        }
    }
    Err(format!("unsupported native op: {op}"))
}
