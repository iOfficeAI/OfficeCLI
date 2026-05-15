use std::{collections::BTreeMap, fs};

use rhwp::wasm_api::HwpDocument;
use serde_json::{json, Value};

use crate::options::{optional_usize, required, required_usize};

pub(crate) fn insert_picture(
    doc: &mut HwpDocument,
    options: &BTreeMap<String, String>,
) -> Result<Value, String> {
    let image = required(options, "--image")?;
    let image_data = fs::read(image).map_err(|e| format!("image read failed: {e}"))?;
    let extension = options
        .get("--extension")
        .map(String::as_str)
        .or_else(|| image.rsplit('.').next())
        .unwrap_or("png");
    json_call(
        doc.insert_picture_native(
            section(options)?,
            paragraph(options)?,
            offset(options)?,
            &image_data,
            req_u32(options, "--width")?,
            req_u32(options, "--height")?,
            req_u32(options, "--natural-width")?,
            req_u32(options, "--natural-height")?,
            extension,
            options
                .get("--description")
                .map(String::as_str)
                .unwrap_or(""),
        ),
    )
}

pub(crate) fn json_call(result: Result<String, impl std::fmt::Display>) -> Result<Value, String> {
    let raw = result.map_err(native_err)?;
    json_parse(raw)
}

pub(crate) fn json_parse(raw: String) -> Result<Value, String> {
    serde_json::from_str(&raw).or_else(|_| Ok(json!({ "raw": raw })))
}

pub(crate) fn bool_call(ok: bool) -> Result<Value, String> {
    Ok(json!({ "ok": ok }))
}

pub(crate) fn native_err(err: impl std::fmt::Display) -> String {
    format!("native operation failed: {err}")
}

pub(crate) fn section(options: &BTreeMap<String, String>) -> Result<usize, String> {
    optional_usize(options, "--section").map(|value| value.unwrap_or(0))
}

pub(crate) fn paragraph(options: &BTreeMap<String, String>) -> Result<usize, String> {
    optional_usize(options, "--paragraph")?
        .or(optional_usize(options, "--para")?)
        .ok_or_else(|| "missing required option: --paragraph".to_string())
}

pub(crate) fn parent_para(options: &BTreeMap<String, String>) -> Result<usize, String> {
    optional_usize(options, "--parent-para")?
        .or(optional_usize(options, "--paragraph")?)
        .ok_or_else(|| "missing required option: --parent-para".to_string())
}

pub(crate) fn control(options: &BTreeMap<String, String>) -> Result<usize, String> {
    required_usize(options, "--control")
}

pub(crate) fn cell(options: &BTreeMap<String, String>) -> Result<usize, String> {
    required_usize(options, "--cell")
}

pub(crate) fn cell_para(options: &BTreeMap<String, String>) -> Result<usize, String> {
    required_usize(options, "--cell-para")
}

pub(crate) fn offset(options: &BTreeMap<String, String>) -> Result<usize, String> {
    optional_usize(options, "--offset").map(|value| value.unwrap_or(0))
}

pub(crate) fn count(options: &BTreeMap<String, String>) -> Result<usize, String> {
    required_usize(options, "--count")
}

pub(crate) fn opt_usize(
    options: &BTreeMap<String, String>,
    key: &str,
) -> Result<Option<usize>, String> {
    optional_usize(options, key)
}

pub(crate) fn req_usize(options: &BTreeMap<String, String>, key: &str) -> Result<usize, String> {
    required_usize(options, key)
}

pub(crate) fn req_u8(options: &BTreeMap<String, String>, key: &str) -> Result<u8, String> {
    required(options, key)?
        .parse::<u8>()
        .map_err(|e| format!("invalid {key} value: {e}"))
}

pub(crate) fn req_u16(options: &BTreeMap<String, String>, key: &str) -> Result<u16, String> {
    required(options, key)?
        .parse::<u16>()
        .map_err(|e| format!("invalid {key} value: {e}"))
}

pub(crate) fn req_u32(options: &BTreeMap<String, String>, key: &str) -> Result<u32, String> {
    required(options, key)?
        .parse::<u32>()
        .map_err(|e| format!("invalid {key} value: {e}"))
}

pub(crate) fn req_i16(options: &BTreeMap<String, String>, key: &str) -> Result<i16, String> {
    required(options, key)?
        .parse::<i16>()
        .map_err(|e| format!("invalid {key} value: {e}"))
}

pub(crate) fn bool_opt(options: &BTreeMap<String, String>, key: &str, default: bool) -> bool {
    options
        .get(key)
        .map(|value| value.eq_ignore_ascii_case("true") || value == "1")
        .unwrap_or(default)
}

pub(crate) fn parse_u32_list(value: Option<&str>) -> Result<Option<Vec<u32>>, String> {
    let Some(value) = value else {
        return Ok(None);
    };
    let mut output = Vec::new();
    for item in value.split(',') {
        let item = item.trim();
        if item.is_empty() {
            continue;
        }
        output.push(
            item.parse::<u32>()
                .map_err(|e| format!("invalid --col-widths item '{item}': {e}"))?,
        );
    }
    Ok((!output.is_empty()).then_some(output))
}

pub(crate) fn is_header(options: &BTreeMap<String, String>) -> bool {
    options
        .get("--kind")
        .map(|value| value.eq_ignore_ascii_case("header"))
        .unwrap_or_else(|| bool_opt(options, "--is-header", true))
}

pub(crate) fn apply_to(options: &BTreeMap<String, String>) -> Result<u8, String> {
    options
        .get("--apply-to")
        .map(|value| {
            value
                .parse::<u8>()
                .map_err(|e| format!("invalid --apply-to value: {e}"))
        })
        .unwrap_or(Ok(0))
}

pub(crate) fn props_json(options: &BTreeMap<String, String>) -> &str {
    options
        .get("--props-json")
        .map(String::as_str)
        .unwrap_or("{}")
}
