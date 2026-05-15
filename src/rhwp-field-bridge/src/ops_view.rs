use std::{collections::BTreeMap, fs, path::Path};

use rhwp::wasm_api::HwpDocument;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};

use crate::options::{optional_usize, required, selected_pages};

pub(crate) fn read_text(
    doc: &HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let page_count = doc.page_count();
    let pages = selected_pages(options, page_count)?;
    let mut page_payload = Vec::new();
    let mut text = String::new();
    for page_idx in pages {
        let mut page_text = doc
            .extract_page_text_native(page_idx)
            .map_err(|e| format!("page {} text extraction failed: {e}", page_idx + 1))?;
        if !page_text.ends_with('\n') {
            page_text.push('\n');
        }
        text.push_str(&page_text);
        page_payload.push(json!({
            "page": page_idx + 1,
            "text": page_text
        }));
    }
    println!(
        "{}",
        json!({
            "text": text,
            "pages": page_payload,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": []
        })
    );
    Ok(())
}

pub(crate) fn render_svg(
    doc: &HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let out_dir = required(options, "--out-dir")?;
    fs::create_dir_all(out_dir).map_err(|e| format!("output directory create failed: {e}"))?;
    let page_count = doc.page_count();
    let pages = selected_pages(options, page_count)?;
    let mut page_payload = Vec::new();
    for page_idx in pages {
        let svg = doc
            .render_page_svg_native(page_idx)
            .map_err(|e| format!("page {} SVG render failed: {e}", page_idx + 1))?;
        let path = Path::new(out_dir).join(format!("page_{:03}.svg", page_idx + 1));
        fs::write(&path, svg.as_bytes()).map_err(|e| format!("SVG write failed: {e}"))?;
        page_payload.push(json!({
            "page": page_idx + 1,
            "path": path.to_string_lossy(),
            "sha256": sha256_bytes(svg.as_bytes())
        }));
    }
    write_manifest_and_print(out_dir, page_payload, format, &[])
}

pub(crate) fn export_pdf(
    doc: &HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let output = required(options, "--output")?;
    if let Some(parent) = Path::new(output).parent() {
        if !parent.as_os_str().is_empty() {
            fs::create_dir_all(parent)
                .map_err(|e| format!("output directory create failed: {e}"))?;
        }
    }

    let page_count = doc.page_count();
    let pages = selected_pages(options, page_count)?;
    let mut svg_pages = Vec::new();
    let mut page_payload = Vec::new();
    for page_idx in pages {
        let svg = doc
            .render_page_svg_native(page_idx)
            .map_err(|e| format!("page {} PDF render failed: {e}", page_idx + 1))?;
        svg_pages.push(svg);
        page_payload.push(json!({ "page": page_idx + 1 }));
    }

    let pdf = rhwp::renderer::pdf::svgs_to_pdf(&svg_pages)
        .map_err(|e| format!("PDF export failed: {e}"))?;
    fs::write(output, &pdf).map_err(|e| format!("PDF write failed: {e}"))?;

    println!(
        "{}",
        json!({
            "pdf": {
                "path": output,
                "bytes": pdf.len(),
                "sha256": sha256_bytes(&pdf)
            },
            "pages": page_payload,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": ["experimental PDF export; verify visual output before production use"]
        })
    );
    Ok(())
}

#[cfg(all(not(target_arch = "wasm32"), feature = "native-skia"))]
pub(crate) fn render_png(
    doc: &HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let out_dir = required(options, "--out-dir")?;
    fs::create_dir_all(out_dir).map_err(|e| format!("output directory create failed: {e}"))?;
    let page_count = doc.page_count();
    let pages = selected_pages(options, page_count)?;
    let mut page_payload = Vec::new();
    for page_idx in pages {
        let png = doc
            .render_page_png_native(page_idx)
            .map_err(|e| format!("page {} PNG render failed: {e}", page_idx + 1))?;
        let path = Path::new(out_dir).join(format!("page_{:03}.png", page_idx + 1));
        fs::write(&path, &png).map_err(|e| format!("PNG write failed: {e}"))?;
        page_payload.push(json!({
            "page": page_idx + 1,
            "path": path.to_string_lossy(),
            "sha256": sha256_bytes(&png),
            "bytes": png.len()
        }));
    }
    write_manifest_and_print(
        out_dir,
        page_payload,
        format,
        &["experimental PNG render; verify visual output before production use"],
    )
}

#[cfg(any(target_arch = "wasm32", not(feature = "native-skia")))]
pub(crate) fn render_png(
    _doc: &HwpDocument,
    _format: &str,
    _options: &BTreeMap<String, String>,
) -> Result<(), String> {
    Err("render-png requires rhwp-field-bridge built with --features native-skia".to_string())
}

pub(crate) fn export_markdown(
    doc: &HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let page_count = doc.page_count();
    let pages = selected_pages(options, page_count)?;
    let mut page_payload = Vec::new();
    let mut markdown = String::new();
    for page_idx in pages {
        let page_markdown = doc
            .extract_page_markdown_native(page_idx)
            .map_err(|e| format!("page {} markdown export failed: {e}", page_idx + 1))?;
        markdown.push_str(&page_markdown);
        if !markdown.ends_with('\n') {
            markdown.push('\n');
        }
        page_payload.push(json!({
            "page": page_idx + 1,
            "markdown": page_markdown
        }));
    }
    println!(
        "{}",
        json!({
            "markdown": markdown,
            "pages": page_payload,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": ["experimental markdown export; image assets are not yet materialized by OfficeCLI"]
        })
    );
    Ok(())
}

pub(crate) fn document_info(doc: &HwpDocument, format: &str) -> Result<(), String> {
    let raw = doc.get_document_info();
    let info: Value = serde_json::from_str(&raw).unwrap_or_else(|_| json!({ "raw": raw }));
    println!(
        "{}",
        json!({
            "info": info,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": []
        })
    );
    Ok(())
}

pub(crate) fn diagnostics(doc: &HwpDocument, format: &str) -> Result<(), String> {
    let raw = doc.get_validation_warnings();
    let diagnostics: Value = serde_json::from_str(&raw).unwrap_or_else(|_| json!({ "raw": raw }));
    println!(
        "{}",
        json!({
            "diagnostics": diagnostics,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": ["rhwp diagnostics are provider warnings, not full OfficeCLI package validation"]
        })
    );
    Ok(())
}

pub(crate) fn dump_pages(
    doc: &HwpDocument,
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let page = optional_usize(options, "--page")?.map(|value| value.saturating_sub(1) as u32);
    let dump = doc.dump_page_items(page);
    println!(
        "{}",
        json!({
            "dump": dump,
            "page": page.map(|value| value + 1),
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": ["diagnostic page dump; output format is not a stable editing contract"]
        })
    );
    Ok(())
}

pub(crate) fn dump_controls(doc: &HwpDocument, format: &str) -> Result<(), String> {
    let dump = format!("{:#?}", doc.document());
    println!(
        "{}",
        json!({
            "dump": dump,
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": ["diagnostic full document/control dump; output format is not a stable editing contract"]
        })
    );
    Ok(())
}

pub(crate) fn thumbnail(
    bytes: &[u8],
    format: &str,
    options: &BTreeMap<String, String>,
) -> Result<(), String> {
    let output = required(options, "--output")?;
    let result = rhwp::parser::extract_thumbnail_only(bytes)
        .ok_or_else(|| "document does not contain a thumbnail preview image".to_string())?;
    fs::write(output, &result.data).map_err(|e| format!("thumbnail write failed: {e}"))?;
    println!(
        "{}",
        json!({
            "thumbnail": {
                "path": output,
                "format": result.format,
                "width": result.width,
                "height": result.height,
                "bytes": result.data.len(),
                "sha256": sha256_bytes(&result.data)
            },
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": []
        })
    );
    Ok(())
}

fn write_manifest_and_print(
    out_dir: &str,
    page_payload: Vec<Value>,
    format: &str,
    warnings: &[&str],
) -> Result<(), String> {
    let manifest = Path::new(out_dir).join("manifest.json");
    fs::write(
        &manifest,
        serde_json::to_vec_pretty(&json!({ "pages": page_payload }))
            .map_err(|e| format!("manifest JSON encode failed: {e}"))?,
    )
    .map_err(|e| format!("manifest write failed: {e}"))?;
    println!(
        "{}",
        json!({
            "pages": page_payload,
            "manifest": manifest.to_string_lossy(),
            "engineVersion": concat!("rhwp-api ", env!("CARGO_PKG_VERSION")),
            "format": format,
            "warnings": warnings
        })
    );
    Ok(())
}

fn sha256_bytes(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    format!("{digest:x}")
}
