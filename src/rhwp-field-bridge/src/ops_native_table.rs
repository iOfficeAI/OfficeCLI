use std::collections::BTreeMap;

use rhwp::wasm_api::HwpDocument;
use serde_json::Value;

use crate::ops_native_support::*;

pub(crate) fn try_run_native_table_op(
    doc: &mut HwpDocument,
    op: &str,
    options: &BTreeMap<String, String>,
) -> Result<Option<Value>, String> {
    let value = match op {
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
        "get-cell-char-properties-at" => json_call(doc.get_cell_char_properties_at_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            cell(options)?,
            cell_para(options)?,
            offset(options)?,
        )),
        "get-cell-para-properties-at" => json_call(doc.get_cell_para_properties_at_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            cell(options)?,
            cell_para(options)?,
        )),
        _ => return Ok(None),
    }?;
    Ok(Some(value))
}
