use std::collections::BTreeMap;

use rhwp::wasm_api::HwpDocument;
use serde_json::{json, Value};

use crate::ops::write_document;
use crate::ops_native_support::*;
use crate::options::required;

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
    match op {
        "delete-text" => json_call(doc.delete_text_native(
            section(options)?,
            paragraph(options)?,
            offset(options)?,
            count(options)?,
        )),
        "split-paragraph" => json_call(doc.split_paragraph_native(
            section(options)?,
            paragraph(options)?,
            offset(options)?,
        )),
        "merge-paragraph" => {
            json_call(doc.merge_paragraph_native(section(options)?, paragraph(options)?))
        }
        "insert-paragraph" => {
            json_call(doc.insert_paragraph_native(section(options)?, paragraph(options)?))
        }
        "delete-paragraph" => {
            json_call(doc.delete_paragraph_native(section(options)?, paragraph(options)?))
        }
        "insert-page-break" => json_call(doc.insert_page_break_native(
            section(options)?,
            paragraph(options)?,
            offset(options)?,
        )),
        "insert-column-break" => json_call(doc.insert_column_break_native(
            section(options)?,
            paragraph(options)?,
            offset(options)?,
        )),
        "set-column-def" => json_call(doc.set_column_def_native(
            section(options)?,
            req_u16(options, "--columns")?,
            req_u8(options, "--column-type")?,
            bool_opt(options, "--same-width", true),
            req_i16(options, "--spacing")?,
        )),
        "insert-text-in-cell" => json_call(doc.insert_text_in_cell_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            cell(options)?,
            cell_para(options)?,
            offset(options)?,
            required(options, "--value")?,
        )),
        "delete-text-in-cell" => json_call(doc.delete_text_in_cell_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            cell(options)?,
            cell_para(options)?,
            offset(options)?,
            count(options)?,
        )),
        "split-paragraph-in-cell" => json_call(doc.split_paragraph_in_cell_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            cell(options)?,
            cell_para(options)?,
            offset(options)?,
        )),
        "merge-paragraph-in-cell" => json_call(doc.merge_paragraph_in_cell_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            cell(options)?,
            cell_para(options)?,
        )),
        "create-table" => json_call(doc.create_table_native(
            section(options)?,
            paragraph(options)?,
            offset(options)?,
            req_u16(options, "--rows")?,
            req_u16(options, "--cols")?,
        )),
        "create-table-ex" => {
            let widths = parse_u32_list(options.get("--col-widths").map(String::as_str))?;
            json_call(doc.create_table_ex_native(
                section(options)?,
                paragraph(options)?,
                offset(options)?,
                req_u16(options, "--rows")?,
                req_u16(options, "--cols")?,
                bool_opt(options, "--treat-as-char", false),
                widths.as_deref(),
            ))
        }
        "insert-table-row" => json_call(doc.insert_table_row_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_u16(options, "--row")?,
            bool_opt(options, "--below", true),
        )),
        "insert-table-column" => json_call(doc.insert_table_column_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_u16(options, "--col")?,
            bool_opt(options, "--right", true),
        )),
        "delete-table-row" => json_call(doc.delete_table_row_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_u16(options, "--row")?,
        )),
        "delete-table-column" => json_call(doc.delete_table_column_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_u16(options, "--col")?,
        )),
        "merge-table-cells" => json_call(doc.merge_table_cells_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_u16(options, "--start-row")?,
            req_u16(options, "--start-col")?,
            req_u16(options, "--end-row")?,
            req_u16(options, "--end-col")?,
        )),
        "split-table-cell" => json_call(doc.split_table_cell_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_u16(options, "--row")?,
            req_u16(options, "--col")?,
        )),
        "split-table-cell-into" => json_call(doc.split_table_cell_into_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_u16(options, "--row")?,
            req_u16(options, "--col")?,
            req_u16(options, "--rows")?,
            req_u16(options, "--cols")?,
            bool_opt(options, "--equal-row-height", true),
            bool_opt(options, "--merge-first", false),
        )),
        "split-table-cells-in-range" => json_call(doc.split_table_cells_in_range_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_u16(options, "--start-row")?,
            req_u16(options, "--start-col")?,
            req_u16(options, "--end-row")?,
            req_u16(options, "--end-col")?,
            req_u16(options, "--rows")?,
            req_u16(options, "--cols")?,
            bool_opt(options, "--equal-row-height", true),
        )),
        "delete-table-control" => json_call(doc.delete_table_control_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
        )),
        "get-paragraph-count" => Ok(
            json!({"count": doc.get_paragraph_count_native(section(options)?).map_err(native_err)?}),
        ),
        "get-paragraph-length" => Ok(
            json!({"length": doc.get_paragraph_length_native(section(options)?, paragraph(options)?).map_err(native_err)?}),
        ),
        "get-text-range" => Ok(
            json!({"text": doc.get_text_range_native(section(options)?, paragraph(options)?, offset(options)?, count(options)?).map_err(native_err)?}),
        ),
        "get-style-list" => json_parse(doc.get_style_list()),
        "get-style-detail" => json_parse(doc.get_style_detail(req_u32(options, "--style-id")?)),
        "update-style" => {
            bool_call(doc.update_style(req_u32(options, "--style-id")?, props_json(options)))
        }
        "update-style-shapes" => bool_call(
            doc.update_style_shapes(
                req_u32(options, "--style-id")?,
                options
                    .get("--char-json")
                    .map(String::as_str)
                    .unwrap_or("{}"),
                options
                    .get("--para-json")
                    .map(String::as_str)
                    .unwrap_or("{}"),
            ),
        ),
        "create-style" => Ok(json!({"styleId": doc.create_style(props_json(options))})),
        "delete-style" => bool_call(doc.delete_style(req_u32(options, "--style-id")?)),
        "get-numbering-list" => json_parse(doc.get_numbering_list()),
        "get-bullet-list" => json_parse(doc.get_bullet_list()),
        "ensure-default-numbering" => Ok(json!({"numberingId": doc.ensure_default_numbering()})),
        "apply-char-format" => json_call(doc.apply_char_format_native(
            section(options)?,
            paragraph(options)?,
            req_usize(options, "--start")?,
            req_usize(options, "--end")?,
            props_json(options),
        )),
        "apply-para-format" => json_call(doc.apply_para_format_native(
            section(options)?,
            paragraph(options)?,
            props_json(options),
        )),
        "apply-style" => json_call(doc.apply_style_native(
            section(options)?,
            paragraph(options)?,
            req_usize(options, "--style-id")?,
        )),
        "set-numbering-restart" => json_call(doc.set_numbering_restart_native(
            section(options)?,
            paragraph(options)?,
            req_u8(options, "--mode")?,
            req_u32(options, "--start-number")?,
        )),
        "set-page-hide" => json_call(doc.set_page_hide_native(
            section(options)?,
            paragraph(options)?,
            bool_opt(options, "--hide-header", false),
            bool_opt(options, "--hide-footer", false),
            bool_opt(options, "--hide-master-page", false),
            bool_opt(options, "--hide-border", false),
            bool_opt(options, "--hide-fill", false),
            bool_opt(options, "--hide-page-num", false),
        )),
        "get-page-hide" => {
            json_call(doc.get_page_hide_native(section(options)?, paragraph(options)?))
        }
        "get-header-footer" => json_call(doc.get_header_footer_native(
            section(options)?,
            is_header(options),
            apply_to(options)?,
        )),
        "create-header-footer" => json_call(doc.create_header_footer_native(
            section(options)?,
            is_header(options),
            apply_to(options)?,
        )),
        "delete-header-footer" => json_call(doc.delete_header_footer_native(
            section(options)?,
            is_header(options),
            apply_to(options)?,
        )),
        "get-header-footer-list" => json_call(doc.get_header_footer_list_native(
            section(options)?,
            is_header(options),
            apply_to(options)?,
        )),
        "insert-text-in-header-footer" => json_call(doc.insert_text_in_header_footer_native(
            section(options)?,
            is_header(options),
            apply_to(options)?,
            req_usize(options, "--hf-para")?,
            offset(options)?,
            required(options, "--value")?,
        )),
        "delete-text-in-header-footer" => json_call(doc.delete_text_in_header_footer_native(
            section(options)?,
            is_header(options),
            apply_to(options)?,
            req_usize(options, "--hf-para")?,
            offset(options)?,
            count(options)?,
        )),
        "split-paragraph-in-header-footer" => {
            json_call(doc.split_paragraph_in_header_footer_native(
                section(options)?,
                is_header(options),
                apply_to(options)?,
                req_usize(options, "--hf-para")?,
                offset(options)?,
            ))
        }
        "merge-paragraph-in-header-footer" => {
            json_call(doc.merge_paragraph_in_header_footer_native(
                section(options)?,
                is_header(options),
                apply_to(options)?,
                req_usize(options, "--hf-para")?,
            ))
        }
        "insert-picture" => insert_picture(doc, options),
        "get-picture-properties" => json_call(doc.get_picture_properties_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
        )),
        "set-picture-properties" => json_call(doc.set_picture_properties_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            props_json(options),
        )),
        "delete-picture-control" => json_call(doc.delete_picture_control_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
        )),
        "get-shape-properties" => json_call(doc.get_shape_properties_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
        )),
        "set-shape-properties" => json_call(doc.set_shape_properties_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            props_json(options),
        )),
        "delete-shape-control" => json_call(doc.delete_shape_control_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
        )),
        "get-equation-properties" => json_call(doc.get_equation_properties_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            opt_usize(options, "--cell")?,
            opt_usize(options, "--cell-para")?,
        )),
        "set-equation-properties" => json_call(doc.set_equation_properties_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            opt_usize(options, "--cell")?,
            opt_usize(options, "--cell-para")?,
            props_json(options),
        )),
        "insert-footnote" => json_call(doc.insert_footnote_native(
            section(options)?,
            paragraph(options)?,
            offset(options)?,
        )),
        "get-footnote-info" => json_call(doc.get_footnote_info_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
        )),
        "insert-text-in-footnote" => json_call(doc.insert_text_in_footnote_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_usize(options, "--footnote-para")?,
            offset(options)?,
            required(options, "--value")?,
        )),
        "delete-text-in-footnote" => json_call(doc.delete_text_in_footnote_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_usize(options, "--footnote-para")?,
            offset(options)?,
            count(options)?,
        )),
        "split-paragraph-in-footnote" => json_call(doc.split_paragraph_in_footnote_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_usize(options, "--footnote-para")?,
            offset(options)?,
        )),
        "merge-paragraph-in-footnote" => json_call(doc.merge_paragraph_in_footnote_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_usize(options, "--footnote-para")?,
        )),
        other => Err(format!("unsupported native op: {other}")),
    }
}
