use std::{collections::BTreeMap, env, fs, process};

use rhwp::wasm_api::HwpDocument;
use serde_json::{json, Value};

fn main() {
    if let Err(err) = run() {
        eprintln!("{err}");
        println!(
            "{}",
            json!({
                "success": false,
                "error": {
                    "message": err,
                    "code": "rhwp_field_bridge_error"
                }
            })
        );
        process::exit(1);
    }
}

fn run() -> Result<(), String> {
    let args: Vec<String> = env::args().skip(1).collect();
    if args.is_empty() || args[0] == "--help" || args[0] == "-h" {
        print_help();
        return Ok(());
    }

    let command = args[0].as_str();
    let options = parse_options(&args[1..])?;
    let input = required(&options, "--input")?;
    let format = required(&options, "--format")?;
    let bytes = fs::read(input).map_err(|e| format!("input read failed: {e}"))?;
    let mut doc = HwpDocument::from_bytes(&bytes).map_err(|e| format!("rhwp parse failed: {e}"))?;

    match command {
        "list-fields" => list_fields(&doc, format),
        "get-field" => get_field(&doc, format, &options),
        "set-field" => set_field(&mut doc, format, &options),
        "replace-text" => replace_text(&mut doc, format, &options),
        "get-cell-text" => get_cell_text(&doc, format, &options),
        "scan-cells" => scan_cells(&doc, format, &options),
        "set-cell-text" => set_cell_text(&mut doc, format, &options),
        _ => Err(format!("unsupported command: {command}")),
    }
}

fn list_fields(doc: &HwpDocument, format: &str) -> Result<(), String> {
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

fn get_field(
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

fn set_field(
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

fn replace_text(
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

fn get_cell_text(
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
        .get_text_in_cell_native(section, parent_para, control, cell, cell_para, offset, count)
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

fn scan_cells(
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

fn set_cell_text(
    doc: &mut HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let output = required(options, "--output")?;
    let output_format = options.get("--output-format").map(String::as_str).unwrap_or(format);
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
    let delete_count = optional_usize(options, "--count")?.unwrap_or_else(|| existing.chars().count());

    let delete_json = doc
        .delete_text_in_cell_native(section, parent_para, control, cell, cell_para, offset, delete_count)
        .map_err(|e| format!("cell text delete failed: {e}"))?;
    let insert_json = doc
        .insert_text_in_cell_native(section, parent_para, control, cell, cell_para, offset, value)
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

fn parse_options(args: &[String]) -> Result<BTreeMap<String, String>, String> {
    let mut options = BTreeMap::new();
    let mut index = 0;
    while index < args.len() {
        let arg = &args[index];
        if !arg.starts_with("--") {
            index += 1;
            continue;
        }
        if arg == "--json" {
            options.insert(arg.clone(), "true".to_string());
            index += 1;
            continue;
        }
        if index + 1 >= args.len() || args[index + 1].starts_with("--") {
            return Err(format!("missing value for {arg}"));
        }
        options.insert(arg.clone(), args[index + 1].clone());
        index += 2;
    }
    Ok(options)
}

fn required<'a>(options: &'a BTreeMap<String, String>, key: &str) -> Result<&'a str, String> {
    options
        .get(key)
        .map(String::as_str)
        .filter(|value| !value.trim().is_empty())
        .ok_or_else(|| format!("missing required option: {key}"))
}

fn required_usize(options: &BTreeMap<String, String>, key: &str) -> Result<usize, String> {
    required(options, key)?
        .parse::<usize>()
        .map_err(|e| format!("invalid {key} value: {e}"))
}

fn optional_usize(options: &BTreeMap<String, String>, key: &str) -> Result<Option<usize>, String> {
    match options.get(key) {
        Some(value) => value
            .parse::<usize>()
            .map(Some)
            .map_err(|e| format!("invalid {key} value: {e}")),
        None => Ok(None),
    }
}

fn print_help() {
    println!(
        "rhwp-field-bridge list-fields|get-field|set-field|replace-text|get-cell-text|scan-cells|set-cell-text --format hwp|hwpx --input <path> [--output <path>] [--output-format hwp|hwpx] [--name <field>] [--id <fieldId>] [--query <text>] [--value <text>] [--mode one|all] [--section N --parent-para N --control N --cell N --cell-para N --offset N --count N] --json"
    );
}
