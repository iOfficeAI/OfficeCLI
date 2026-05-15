use std::collections::BTreeMap;

use rhwp::wasm_api::HwpDocument;
use serde_json::Value;

use crate::ops_native_support::*;
use crate::options::required;

pub(crate) fn try_run_native_header_footer_op(
    doc: &mut HwpDocument,
    op: &str,
    options: &BTreeMap<String, String>,
) -> Result<Option<Value>, String> {
    let value = match op {
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
        "get-header-footer-para-info" => json_call(doc.get_header_footer_para_info_native(
            section(options)?,
            is_header(options),
            apply_to(options)?,
            req_usize(options, "--hf-para")?,
        )),
        "navigate-header-footer-by-page" => json_call(doc.navigate_header_footer_by_page_native(
            req_u32(options, "--page")?,
            is_header(options),
            req_i32(options, "--direction")?,
        )),
        "toggle-hide-header-footer" => json_call(
            doc.toggle_hide_header_footer_native(req_u32(options, "--page")?, is_header(options)),
        ),
        "get-para-properties-in-hf" => json_call(doc.get_para_properties_in_hf_native(
            section(options)?,
            is_header(options),
            apply_to(options)?,
            req_usize(options, "--hf-para")?,
        )),
        "apply-para-format-in-hf" => json_call(doc.apply_para_format_in_hf_native(
            section(options)?,
            is_header(options),
            apply_to(options)?,
            req_usize(options, "--hf-para")?,
            props_json(options),
        )),
        "insert-field-in-hf" => json_call(doc.insert_field_in_hf_native(
            section(options)?,
            is_header(options),
            apply_to(options)?,
            req_usize(options, "--hf-para")?,
            offset(options)?,
            req_u8(options, "--field-type")?,
        )),
        "apply-hf-template" => json_call(doc.apply_hf_template_native(
            section(options)?,
            is_header(options),
            apply_to(options)?,
            req_u8(options, "--template-id")?,
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
        _ => return Ok(None),
    }?;
    Ok(Some(value))
}
