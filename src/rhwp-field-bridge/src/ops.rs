use std::{collections::BTreeMap, fs, path::Path};

use rhwp::wasm_api::HwpDocument;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};

use crate::options::{optional_usize, required, required_usize, selected_pages};

pub(crate) fn create_blank(options: &BTreeMap<String, String>) -> Result<(), String> {
    let output = required(options, "--output")?;
    let mut doc = HwpDocument::create_empty();
    let info_json = doc
        .create_blank_document_native()
        .map_err(|e| format!("blank HWP creation failed: {e}"))?;
    write_document(&mut doc, "hwp", output)?;
    let document_info: Value =
        serde_json::from_str(&info_json).unwrap_or_else(|_| json!({ "raw": info_json }));
    println!(
        "{}",
        json!({
            "created": true,
            "operation": "create-blank",
            "output": output,
            "format": "hwp",
            "documentInfo": document_info,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "warnings": ["experimental blank HWP creation; verify with provider readback before production use"]
        })
    );
    Ok(())
}

pub(crate) fn read_text(
    doc: &HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let page_count = doc.page_count();
    let pages = selected_pages(options, page_count)?;
    let mut page_payload = Vec::new();
    let mut text = String::new();
    for page_idx in pages {
        let mut page_text = doc
            .extract_page_text_native(page_idx)
            .map_err(|e| format!("page {} text extraction failed: {e}", page_idx + 1))?;
        if !page_text.ends_with('\n') {
            page_text.push('\n');
        }
        text.push_str(&page_text);
        page_payload.push(json!({
            "page": page_idx + 1,
            "text": page_text
        }));
    }
    println!(
        "{}",
        json!({
            "text": text,
            "pages": page_payload,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": []
        })
    );
    Ok(())
}

pub(crate) fn render_svg(
    doc: &HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let out_dir = required(options, "--out-dir")?;
    fs::create_dir_all(out_dir).map_err(|e| format!("output directory create failed: {e}"))?;
    let page_count = doc.page_count();
    let pages = selected_pages(options, page_count)?;
    let mut page_payload = Vec::new();
    for page_idx in pages {
        let svg = doc
            .render_page_svg_native(page_idx)
            .map_err(|e| format!("page {} SVG render failed: {e}", page_idx + 1))?;
        let path = Path::new(out_dir).join(format!("page_{:03}.svg", page_idx + 1));
        fs::write(&path, svg.as_bytes()).map_err(|e| format!("SVG write failed: {e}"))?;
        page_payload.push(json!({
            "page": page_idx + 1,
            "path": path.to_string_lossy(),
            "sha256": sha256_bytes(svg.as_bytes())
        }));
    }
    let manifest = Path::new(out_dir).join("manifest.json");
    fs::write(
        &manifest,
        serde_json::to_vec_pretty(&json!({ "pages": page_payload }))
            .map_err(|e| format!("manifest JSON encode failed: {e}"))?,
    )
    .map_err(|e| format!("manifest write failed: {e}"))?;
    println!(
        "{}",
        json!({
            "pages": page_payload,
            "manifest": manifest.to_string_lossy(),
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": []
        })
    );
    Ok(())
}

pub(crate) fn list_fields(doc: &HwpDocument, format: &str) -> Result<(), String> {
    let fields: Value = serde_json::from_str(&doc.get_field_list_json())
        .map_err(|e| format!("field list JSON parse failed: {e}"))?;
    println!(
        "{}",
        json!({
            "fields": fields,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": []
        })
    );
    Ok(())
}

pub(crate) fn get_field(
    doc: &HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let value_json = if let Some(name) = options.get("--name") {
        doc.get_field_value_by_name(name)
            .map_err(|e| format!("field name lookup failed: {e}"))?
    } else if let Some(id) = options.get("--id") {
        let id = id
            .parse::<u32>()
            .map_err(|e| format!("invalid --id value: {e}"))?;
        doc.get_field_value_by_id(id)
            .map_err(|e| format!("field id lookup failed: {e}"))?
    } else {
        return Err("missing --name or --id".to_string());
    };

    let value: Value = serde_json::from_str(&value_json)
        .map_err(|e| format!("field value JSON parse failed: {e}"))?;
    println!(
        "{}",
        json!({
            "field": value,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": []
        })
    );
    Ok(())
}

pub(crate) fn set_field(
    doc: &mut HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let output = required(options, "--output")?;
    let value = required(options, "--value")?;
    let mutation_json = if let Some(name) = options.get("--name") {
        doc.set_field_value_by_name(name, value)
            .map_err(|e| format!("field name mutation failed: {e}"))?
    } else if let Some(id) = options.get("--id") {
        let id = id
            .parse::<u32>()
            .map_err(|e| format!("invalid --id value: {e}"))?;
        doc.set_field_value_by_id(id, value)
            .map_err(|e| format!("field id mutation failed: {e}"))?
    } else {
        return Err("missing --name or --id".to_string());
    };

    write_document(doc, format, output)?;

    let field: Value = serde_json::from_str(&mutation_json)
        .map_err(|e| format!("field mutation JSON parse failed: {e}"))?;
    println!(
        "{}",
        json!({
            "field": field,
            "output": output,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": ["experimental field mutation; verify round-trip before production use"]
        })
    );
    Ok(())
}

pub(crate) fn replace_text(
    doc: &mut HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let output = required(options, "--output")?;
    let query = required(options, "--query")?;
    let value = required(options, "--value")?;
    let mode = options.get("--mode").map(String::as_str).unwrap_or("one");
    let case_sensitive = options
        .get("--case-sensitive")
        .map(|v| v.eq_ignore_ascii_case("true"))
        .unwrap_or(false);

    let replace_json = match mode {
        "one" => doc
            .replace_one_native(query, value, case_sensitive)
            .map_err(|e| format!("replace one failed: {e}"))?,
        "all" => doc
            .replace_all_native(query, value, case_sensitive)
            .map_err(|e| format!("replace all failed: {e}"))?,
        other => return Err(format!("unsupported --mode: {other}")),
    };
    write_document(doc, format, output)?;

    let replacement: Value = serde_json::from_str(&replace_json)
        .map_err(|e| format!("replace text JSON parse failed: {e}"))?;
    println!(
        "{}",
        json!({
            "replacement": replacement,
            "output": output,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": ["experimental text replacement; verify round-trip before production use"]
        })
    );
    Ok(())
}

pub(crate) fn get_cell_text(
    doc: &HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let section = required_usize(options, "--section")?;
    let parent_para = required_usize(options, "--parent-para")?;
    let control = required_usize(options, "--control")?;
    let cell = required_usize(options, "--cell")?;
    let cell_para = required_usize(options, "--cell-para")?;
    let offset = optional_usize(options, "--offset")?.unwrap_or(0);
    let count = optional_usize(options, "--count")?.unwrap_or(usize::MAX / 2);

    let text = doc
        .get_text_in_cell_native(
            section,
            parent_para,
            control,
            cell,
            cell_para,
            offset,
            count,
        )
        .map_err(|e| format!("cell text lookup failed: {e}"))?;
    println!(
        "{}",
        json!({
            "cell": {
                "section": section,
                "parentPara": parent_para,
                "control": control,
                "cell": cell,
                "cellPara": cell_para,
                "offset": offset,
                "count": count,
                "text": text
            },
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": []
        })
    );
    Ok(())
}

pub(crate) fn scan_cells(
    doc: &HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let section = optional_usize(options, "--section")?.unwrap_or(0);
    let max_parent_para = optional_usize(options, "--max-parent-para")?.unwrap_or(50);
    let max_control = optional_usize(options, "--max-control")?.unwrap_or(4);
    let max_cell = optional_usize(options, "--max-cell")?.unwrap_or(64);
    let max_cell_para = optional_usize(options, "--max-cell-para")?.unwrap_or(4);
    let count = optional_usize(options, "--count")?.unwrap_or(120);
    let include_empty = options
        .get("--include-empty")
        .map(|v| v.eq_ignore_ascii_case("true"))
        .unwrap_or(false);
    let mut cells = Vec::new();

    for parent_para in 0..=max_parent_para {
        for control in 0..=max_control {
            for cell in 0..=max_cell {
                for cell_para in 0..=max_cell_para {
                    let result = doc.get_text_in_cell_native(
                        section,
                        parent_para,
                        control,
                        cell,
                        cell_para,
                        0,
                        count,
                    );
                    if let Ok(text) = result {
                        if include_empty || !text.is_empty() {
                            cells.push(json!({
                                "section": section,
                                "parentPara": parent_para,
                                "control": control,
                                "cell": cell,
                                "cellPara": cell_para,
                                "text": text
                            }));
                        }
                    }
                }
            }
        }
    }

    println!(
        "{}",
        json!({
            "cells": cells,
            "count": cells.len(),
            "limits": {
                "section": section,
                "maxParentPara": max_parent_para,
                "maxControl": max_control,
                "maxCell": max_cell,
                "maxCellPara": max_cell_para
            },
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": ["bounded scan; absence from results is not proof that no table cell exists"]
        })
    );
    Ok(())
}

pub(crate) fn set_cell_text(
    doc: &mut HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let output = required(options, "--output")?;
    let output_format = options
        .get("--output-format")
        .map(String::as_str)
        .unwrap_or(format);
    if format == "hwpx" && output_format != "hwp" {
        return Err("HWPX table cell mutation must use --output-format hwp".to_string());
    }

    let value = required(options, "--value")?;
    let section = required_usize(options, "--section")?;
    let parent_para = required_usize(options, "--parent-para")?;
    let control = required_usize(options, "--control")?;
    let cell = required_usize(options, "--cell")?;
    let cell_para = required_usize(options, "--cell-para")?;
    let offset = optional_usize(options, "--offset")?.unwrap_or(0);
    let existing = doc
        .get_text_in_cell_native(
            section,
            parent_para,
            control,
            cell,
            cell_para,
            offset,
            usize::MAX / 2,
        )
        .map_err(|e| format!("cell text lookup failed before mutation: {e}"))?;
    let delete_count =
        optional_usize(options, "--count")?.unwrap_or_else(|| existing.chars().count());

    let delete_json = doc
        .delete_text_in_cell_native(
            section,
            parent_para,
            control,
            cell,
            cell_para,
            offset,
            delete_count,
        )
        .map_err(|e| format!("cell text delete failed: {e}"))?;
    let insert_json = doc
        .insert_text_in_cell_native(
            section,
            parent_para,
            control,
            cell,
            cell_para,
            offset,
            value,
        )
        .map_err(|e| format!("cell text insert failed: {e}"))?;
    write_document(doc, output_format, output)?;

    let delete_result: Value = serde_json::from_str(&delete_json)
        .map_err(|e| format!("delete cell text JSON parse failed: {e}"))?;
    let insert_result: Value = serde_json::from_str(&insert_json)
        .map_err(|e| format!("insert cell text JSON parse failed: {e}"))?;
    println!(
        "{}",
        json!({
            "cell": {
                "section": section,
                "parentPara": parent_para,
                "control": control,
                "cell": cell,
                "cellPara": cell_para,
                "offset": offset,
                "deleteCount": delete_count,
                "oldText": existing,
                "newText": value
            },
            "delete": delete_result,
            "insert": insert_result,
            "output": output,
            "outputFormat": output_format,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": ["experimental table cell mutation; HWPX inputs are saved as HWP for verification"]
        })
    );
    Ok(())
}

pub(crate) fn save_as_hwp(
    doc: &mut HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let output = required(options, "--output")?;
    write_document(doc, "hwp", output)?;
    println!(
        "{}",
        json!({
            "saved": true,
            "operation": "save-as-hwp",
            "output": output,
            "outputFormat": "hwp",
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": ["experimental HWP export; verify round-trip before production use"]
        })
    );
    Ok(())
}

fn write_document(doc: &mut HwpDocument, format: &str, output: &str) -> Result<(), String> {
    let bytes = match format {
        "hwp" => doc
            .export_hwp_with_adapter()
            .map_err(|e| format!("HWP export failed: {e}"))?,
        "hwpx" => doc
            .export_hwpx_native()
            .map_err(|e| format!("HWPX export failed: {e}"))?,
        other => return Err(format!("unsupported --format: {other}")),
    };
    fs::write(output, bytes).map_err(|e| format!("output write failed: {e}"))
}

fn sha256_bytes(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    format!("{digest:x}")
}
