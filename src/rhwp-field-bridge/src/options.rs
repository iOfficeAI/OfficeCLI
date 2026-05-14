use std::collections::BTreeMap;

pub(crate) fn parse_options(args: &[String]) -> Result<BTreeMap<String, String>, String> {
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

pub(crate) fn required<'a>(
    options: &'a BTreeMap<String, String>,
    key: &str,
) -> Result<&'a str, String> {
    options
        .get(key)
        .map(String::as_str)
        .filter(|value| !value.trim().is_empty())
        .ok_or_else(|| format!("missing required option: {key}"))
}

pub(crate) fn required_usize(
    options: &BTreeMap<String, String>,
    key: &str,
) -> Result<usize, String> {
    required(options, key)?
        .parse::<usize>()
        .map_err(|e| format!("invalid {key} value: {e}"))
}

pub(crate) fn optional_usize(
    options: &BTreeMap<String, String>,
    key: &str,
) -> Result<Option<usize>, String> {
    match options.get(key) {
        Some(value) => value
            .parse::<usize>()
            .map(Some)
            .map_err(|e| format!("invalid {key} value: {e}")),
        None => Ok(None),
    }
}

pub(crate) fn selected_pages(
    options: &BTreeMap<String, String>,
    page_count: u32,
) -> Result<Vec<u32>, String> {
    if page_count == 0 {
        return Err("document has no pages".to_string());
    }
    let selector = options.get("--page").map(String::as_str).unwrap_or("all");
    if selector.eq_ignore_ascii_case("all") {
        return Ok((0..page_count).collect());
    }
    let one_based = selector
        .parse::<u32>()
        .map_err(|e| format!("invalid --page value: {e}"))?;
    if one_based == 0 || one_based > page_count {
        return Err(format!(
            "--page out of range: {one_based}; valid range is 1..={page_count}"
        ));
    }
    Ok(vec![one_based - 1])
}
