use std::collections::BTreeMap;

use rhwp::wasm_api::HwpDocument;
use serde_json::Value;

use crate::ops_native_support::*;
use crate::options::required;

pub(crate) fn try_run_native_object_op(
    doc: &mut HwpDocument,
    op: &str,
    options: &BTreeMap<String, String>,
) -> Result<Option<Value>, String> {
    let value = match op {
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
        "create-shape-control" => json_call(
            doc.create_shape_control_native(
                section(options)?,
                paragraph(options)?,
                offset(options)?,
                req_u32(options, "--width")?,
                req_u32(options, "--height")?,
                req_u32(options, "--horz-offset")?,
                req_u32(options, "--vert-offset")?,
                bool_opt(options, "--treat-as-char", false),
                options
                    .get("--text-wrap")
                    .map(String::as_str)
                    .unwrap_or("InFrontOfText"),
                options
                    .get("--shape-type")
                    .map(String::as_str)
                    .unwrap_or("rectangle"),
                bool_opt(options, "--line-flip-x", false),
                bool_opt(options, "--line-flip-y", false),
                &[],
            ),
        ),
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
        "change-shape-z-order" => json_call(doc.change_shape_z_order_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            required(options, "--operation")?,
        )),
        "move-line-endpoint" => json_call(doc.move_line_endpoint_native(
            section(options)?,
            parent_para(options)?,
            control(options)?,
            req_i32(options, "--start-x")?,
            req_i32(options, "--start-y")?,
            req_i32(options, "--end-x")?,
            req_i32(options, "--end-y")?,
        )),
        "group-shapes" => {
            let targets = parse_targets(required(options, "--targets")?)?;
            json_call(doc.group_shapes_native(section(options)?, &targets))
        }
        "ungroup-shape" => json_call(doc.ungroup_shape_native(
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
        "render-equation-preview" => json_call(doc.render_equation_preview_native(
            required(options, "--script")?,
            req_u32(options, "--font-size")?,
            req_u32(options, "--color")?,
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
        _ => return Ok(None),
    }?;
    Ok(Some(value))
}
