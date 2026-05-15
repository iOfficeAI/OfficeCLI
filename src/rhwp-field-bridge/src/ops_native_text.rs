use std::collections::BTreeMap;

use rhwp::wasm_api::HwpDocument;
use serde_json::{json, Value};

use crate::ops_native_support::*;
use crate::options::required;

pub(crate) fn try_run_native_text_op(
    doc: &mut HwpDocument,
    op: &str,
    options: &BTreeMap<String, String>,
) -> Result<Option<Value>, String> {
    let value = match op {
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
        "get-paragraph-count" => Ok(json!({
            "count": doc.get_paragraph_count_native(section(options)?).map_err(native_err)?
        })),
        "get-paragraph-length" => Ok(json!({
            "length": doc.get_paragraph_length_native(
                section(options)?,
                paragraph(options)?,
            ).map_err(native_err)?
        })),
        "get-text-range" => Ok(json!({
            "text": doc.get_text_range_native(
                section(options)?,
                paragraph(options)?,
                offset(options)?,
                count(options)?,
            ).map_err(native_err)?
        })),
        "get-textbox-control-index" => Ok(json!({
            "control": doc.get_textbox_control_index_native(section(options)?, paragraph(options)?)
        })),
        "find-next-editable-control" => json_parse(
            doc.find_next_editable_control_native(
                section(options)?,
                paragraph(options)?,
                options
                    .get("--control")
                    .map(|_| req_i32(options, "--control"))
                    .unwrap_or(Ok(-1))?,
                req_i32(options, "--delta")?,
            ),
        ),
        "find-nearest-control-backward" => json_parse(doc.find_nearest_control_backward_native(
            section(options)?,
            paragraph(options)?,
            offset(options)?,
        )),
        "find-nearest-control-forward" => json_parse(doc.find_nearest_control_forward_native(
            section(options)?,
            paragraph(options)?,
            offset(options)?,
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
        "get-cell-paragraph-count" => Ok(json!({
            "count": doc.get_cell_paragraph_count_native(
                section(options)?,
                parent_para(options)?,
                control(options)?,
                cell(options)?,
            ).map_err(native_err)?
        })),
        "get-cell-paragraph-length" => Ok(json!({
            "length": doc.get_cell_paragraph_length_native(
                section(options)?,
                parent_para(options)?,
                control(options)?,
                cell(options)?,
                cell_para(options)?,
            ).map_err(native_err)?
        })),
        "get-text-in-cell" => Ok(json!({
            "text": doc.get_text_in_cell_native(
                section(options)?,
                parent_para(options)?,
                control(options)?,
                cell(options)?,
                cell_para(options)?,
                offset(options)?,
                count(options)?,
            ).map_err(native_err)?
        })),
        _ => return Ok(None),
    }?;
    Ok(Some(value))
}
