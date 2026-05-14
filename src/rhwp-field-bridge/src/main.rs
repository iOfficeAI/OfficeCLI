mod ops;
mod options;

use std::{env, fs, process};

use rhwp::wasm_api::HwpDocument;
use serde_json::json;

use ops::{
    create_blank, get_cell_text, get_field, list_fields, read_text, render_svg, replace_text,
    save_as_hwp, scan_cells, set_cell_text, set_field,
};
use options::parse_options;

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

    if command == "create-blank" {
        return create_blank(&options);
    }

    let input = options::required(&options, "--input")?;
    let format = options::required(&options, "--format")?;
    let bytes = fs::read(input).map_err(|e| format!("input read failed: {e}"))?;
    let mut doc = HwpDocument::from_bytes(&bytes).map_err(|e| format!("rhwp parse failed: {e}"))?;

    match command {
        "read-text" => read_text(&doc, format, &options),
        "render-svg" => render_svg(&doc, format, &options),
        "list-fields" => list_fields(&doc, format),
        "get-field" => get_field(&doc, format, &options),
        "set-field" => set_field(&mut doc, format, &options),
        "replace-text" => replace_text(&mut doc, format, &options),
        "get-cell-text" => get_cell_text(&doc, format, &options),
        "scan-cells" => scan_cells(&doc, format, &options),
        "set-cell-text" => set_cell_text(&mut doc, format, &options),
        "save-as-hwp" => save_as_hwp(&mut doc, format, &options),
        _ => Err(format!("unsupported command: {command}")),
    }
}

fn print_help() {
    println!(
        "rhwp-field-bridge create-blank|read-text|render-svg|list-fields|get-field|set-field|replace-text|get-cell-text|scan-cells|set-cell-text|save-as-hwp --format hwp|hwpx --input <path> [--output <path>] [--out-dir <dir>] [--output-format hwp|hwpx] [--name <field>] [--id <fieldId>] [--query <text>] [--value <text>] [--mode one|all] [--section N --parent-para N --control N --cell N --cell-para N --offset N --count N] --json"
    );
}
