mod ops;
mod ops_native;
mod ops_native_support;
mod ops_text;
mod ops_view;
mod options;

use std::{env, fs, process};

use rhwp::wasm_api::HwpDocument;
use serde_json::json;

use ops::{
    convert_to_editable, create_blank, get_cell_text, get_field, list_fields, replace_text,
    save_as_hwp, scan_cells, set_cell_text, set_field,
};
use ops_native::native_op;
use ops_text::insert_text;
use ops_view::{
    diagnostics, document_info, dump_controls, dump_pages, export_markdown, export_pdf, read_text,
    render_png, render_svg, thumbnail,
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
        "render-png" => render_png(&doc, format, &options),
        "export-pdf" => export_pdf(&doc, format, &options),
        "export-markdown" => export_markdown(&doc, format, &options),
        "document-info" => document_info(&doc, format),
        "diagnostics" => diagnostics(&doc, format),
        "dump-controls" => dump_controls(&doc, format),
        "dump-pages" => dump_pages(&doc, format, &options),
        "thumbnail" => thumbnail(&bytes, format, &options),
        "list-fields" => list_fields(&doc, format),
        "get-field" => get_field(&doc, format, &options),
        "set-field" => set_field(&mut doc, format, &options),
        "replace-text" => replace_text(&mut doc, format, &options),
        "insert-text" => insert_text(&mut doc, format, &options),
        "get-cell-text" => get_cell_text(&doc, format, &options),
        "scan-cells" => scan_cells(&doc, format, &options),
        "set-cell-text" => set_cell_text(&mut doc, format, &options),
        "convert-to-editable" => convert_to_editable(&mut doc, format, &options),
        "native-op" => native_op(&mut doc, format, &options),
        "save-as-hwp" => save_as_hwp(&mut doc, format, &options),
        _ => Err(format!("unsupported command: {command}")),
    }
}

fn print_help() {
    let render_png = if cfg!(all(not(target_arch = "wasm32"), feature = "native-skia")) {
        "|render-png"
    } else {
        ""
    };
    println!(
        "rhwp-field-bridge create-blank|read-text|render-svg{render_png}|export-pdf|export-markdown|document-info|diagnostics|dump-controls|dump-pages|thumbnail|list-fields|get-field|set-field|replace-text|insert-text|get-cell-text|scan-cells|set-cell-text|convert-to-editable|native-op|save-as-hwp --format hwp|hwpx --input <path> [--op <native-op>] [--output <path>] [--out-dir <dir>] [--output-format hwp|hwpx] [--name <field>] [--id <fieldId>] [--query <text>] [--value <text>] [--mode one|all] [--section N --paragraph N --para N --parent-para N --control N --cell N --cell-para N --offset N --count N] --json"
    );
}
