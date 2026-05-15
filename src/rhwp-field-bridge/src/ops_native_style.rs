use std::collections::BTreeMap;

use rhwp::wasm_api::HwpDocument;
use serde_json::{json, Value};

use crate::ops_native_support::*;
use crate::options::required;

pub(crate) fn try_run_native_style_op(
    doc: &mut HwpDocument,
    op: &str,
    options: &BTreeMap<String, String>,
) -> Result<Option<Value>, String> {
    let value = match op {
        "get-char-properties-at" => json_call(doc.get_char_properties_at_native(
            section(options)?,
            paragraph(options)?,
            offset(options)?,
        )),
        "get-para-properties-at" => {
            json_call(doc.get_para_properties_at_native(section(options)?, paragraph(options)?))
        }
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
        "find-or-create-font-id" => Ok(json!({
            "fontId": doc.find_or_create_font_id_native(required(options, "--name")?)
        })),
        "apply-char-format" => json_call(doc.apply_char_format_native(
            section(options)?,
            paragraph(options)?,
            req_usize(options, "--start")?,
            req_usize(options, "--end")?,
            props_json(options),
        )),
        "apply-char-format-in-cell" => json_call(doc.apply_char_format_in_cell_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            cell(options)?,
            cell_para(options)?,
            req_usize(options, "--start")?,
            req_usize(options, "--end")?,
            props_json(options),
        )),
        "apply-para-format" => json_call(doc.apply_para_format_native(
            section(options)?,
            paragraph(options)?,
            props_json(options),
        )),
        "apply-para-format-in-cell" => json_call(doc.apply_para_format_in_cell_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            cell(options)?,
            cell_para(options)?,
            props_json(options),
        )),
        "apply-style" => json_call(doc.apply_style_native(
            section(options)?,
            paragraph(options)?,
            req_usize(options, "--style-id")?,
        )),
        "apply-cell-style" => json_call(doc.apply_cell_style_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            cell(options)?,
            cell_para(options)?,
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
        _ => return Ok(None),
    }?;
    Ok(Some(value))
}
