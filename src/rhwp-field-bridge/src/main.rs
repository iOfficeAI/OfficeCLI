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

    let bytes = match format {
        "hwp" => doc
            .export_hwp_with_adapter()
            .map_err(|e| format!("HWP export failed: {e}"))?,
        "hwpx" => doc
            .export_hwpx_native()
            .map_err(|e| format!("HWPX export failed: {e}"))?,
        other => return Err(format!("unsupported --format: {other}")),
    };
    fs::write(output, bytes).map_err(|e| format!("output write failed: {e}"))?;

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

fn print_help() {
    println!(
        "rhwp-field-bridge list-fields|get-field|set-field --format hwp|hwpx --input <path> [--output <path>] [--name <field>] [--id <fieldId>] [--value <text>] --json"
    );
}
