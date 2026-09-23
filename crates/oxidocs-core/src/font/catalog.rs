// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Portable metrics for installed faces, generated without font programs.
//!
//! Calibrated registry entries remain authoritative. This catalog supplies
//! missing faces before disk resolution, so Linux and WASM do not depend on
//! the fonts installed on the Windows measurement machine. Each face is a
//! separate gzip member and is decoded only when requested.

use std::collections::HashMap;
use std::sync::{OnceLock, RwLock};

use serde::Deserialize;

use super::{FontMetrics, RawFontMetrics};

const DATA: &[u8] = include_bytes!("data/font_catalog_metrics.gz");

#[derive(Deserialize)]
struct Face {
    key: String,
    offset: usize,
    length: usize,
    families: Vec<String>,
    full_names: Vec<String>,
    bold: bool,
    italic: bool,
    weight: u16,
    width_class: u16,
    priority: u16,
    #[serde(default)]
    codepage_range1: Option<u32>,
}

#[derive(Deserialize)]
struct Index {
    schema: u32,
    faces: Vec<Face>,
}

struct Catalog {
    faces: Vec<Face>,
    names: HashMap<String, Vec<usize>>,
    metrics: Vec<OnceLock<FontMetrics>>,
}

impl Catalog {
    fn load() -> Self {
        let index: Index = serde_json::from_str(include_str!("data/font_catalog_index.json"))
            .expect("generated font catalog index must be valid");
        assert_eq!(index.schema, 1, "unsupported font catalog schema");
        let mut names: HashMap<String, Vec<usize>> = HashMap::new();
        for (i, face) in index.faces.iter().enumerate() {
            assert!(
                face.offset
                    .checked_add(face.length)
                    .is_some_and(|end| end <= DATA.len()),
                "font catalog member is outside the metric data: {}",
                face.key
            );
            for name in face.families.iter().chain(&face.full_names) {
                let entries = names.entry(name.clone()).or_default();
                if entries.last() != Some(&i) {
                    entries.push(i);
                }
            }
        }
        let metrics = (0..index.faces.len()).map(|_| OnceLock::new()).collect();
        Self {
            faces: index.faces,
            names,
            metrics,
        }
    }

    fn select(&self, name: &str, bold: bool, italic: bool) -> Option<usize> {
        self.names
            .get(name)?
            .iter()
            .copied()
            .filter(|&i| {
                let face = &self.faces[i];
                if face.families.iter().any(|n| n == name) {
                    // A family name must select the requested style, not whichever
                    // face happened to be enumerated first.
                    face.bold == bold && face.italic == italic
                } else {
                    // A full name can explicitly name a styled face. PostScript
                    // names alone are metadata, not Word family aliases.
                    (!bold || face.bold) && (!italic || face.italic)
                }
            })
            .min_by_key(|&i| {
                let face = &self.faces[i];
                (
                    face.priority,
                    face.weight.abs_diff(if bold { 700 } else { 400 }),
                    face.width_class.abs_diff(5),
                    i,
                )
            })
    }

    fn metrics(&self, i: usize) -> &FontMetrics {
        self.metrics[i].get_or_init(|| {
            let face = &self.faces[i];
            let decoder =
                flate2::read::GzDecoder::new(&DATA[face.offset..face.offset + face.length]);
            let raw: RawFontMetrics = serde_json::from_reader(decoder)
                .expect("generated font catalog member must contain metrics");
            let em = f32::from(raw.units_per_em);
            assert!(em > 0.0, "font catalog contains zero units per em");
            FontMetrics {
                synthetic_bold_advance: 0.0,
                family: raw.family,
                units_per_em: raw.units_per_em,
                ascent: f32::from(raw.ascender) / em,
                descent: -f32::from(raw.descender) / em,
                line_gap: f32::from(raw.line_gap) / em,
                win_ascent: f32::from(raw.win_ascent) / em,
                win_descent: f32::from(raw.win_descent) / em,
                typo_ascent: f32::from(raw.typo_ascender) / em,
                typo_descent: -f32::from(raw.typo_descender) / em,
                typo_line_gap: f32::from(raw.typo_line_gap) / em,
                use_typo_metrics: raw.use_typo_metrics,
                codepage_range1: face.codepage_range1,
                sym_coverage: Vec::new(),
                char_widths: raw
                    .widths
                    .into_iter()
                    .filter_map(|(cp, advance)| {
                        char::from_u32(cp).map(|c| (c, f32::from(advance) / em))
                    })
                    .collect(),
            }
        })
    }
}

fn catalog() -> &'static Catalog {
    static CATALOG: OnceLock<Catalog> = OnceLock::new();
    CATALOG.get_or_init(Catalog::load)
}

/// Resolve a real styled face without sharing the regular face's calibrated
/// width-table key. A regular-family key would silently replace these advances.
pub(super) fn resolve_styled(family: &str, bold: bool, italic: bool) -> Option<&'static FontMetrics> {
    type Cache = HashMap<(String, bool, bool), &'static FontMetrics>;
    static CACHE: OnceLock<RwLock<Cache>> = OnceLock::new();
    let key = (family.trim().to_lowercase(), bold, italic);
    let cache = CACHE.get_or_init(|| RwLock::new(HashMap::new()));
    if let Some(metrics) = cache.read().expect("font catalog cache poisoned").get(&key) {
        return Some(*metrics);
    }
    let catalog = catalog();
    let index = catalog.select(&key.0, bold, italic)?;
    let face = &catalog.faces[index];
    let mut cache = cache.write().expect("font catalog cache poisoned");
    Some(*cache.entry(key.clone()).or_insert_with(|| {
        let mut metrics = catalog.metrics(index).clone();
        metrics.family = if face.families.iter().any(|name| name == &key.0) {
            let suffix = match (face.bold, face.italic) {
                (true, true) => " Bold Italic",
                (true, false) => " Bold",
                (false, true) => " Italic",
                (false, false) => "",
            };
            format!("{family}{suffix}")
        } else {
            family.to_owned()
        };
        Box::leak(Box::new(metrics))
    }))
}

pub(super) fn codepage_range1(family: &str) -> Option<u32> {
    let name = family.trim().to_lowercase();
    let catalog = catalog();
    let index = catalog.select(&name, false, false).or_else(|| {
        catalog
            .names
            .get(&name)
            .and_then(|faces| faces.first().copied())
    })?;
    catalog.faces[index].codepage_range1
}

/// Preserve the requested name, as disk resolution does, because existing
/// calibrated spacing rules distinguish localized family names.
pub(super) fn resolve(family: &str, bold: bool, italic: bool) -> Option<&'static FontMetrics> {
    type Cache = HashMap<(String, bool, bool), &'static FontMetrics>;
    static CACHE: OnceLock<RwLock<Cache>> = OnceLock::new();
    let key = (family.trim().to_lowercase(), bold, italic);
    let cache = CACHE.get_or_init(|| RwLock::new(HashMap::new()));
    if let Some(metrics) = cache.read().expect("font catalog cache poisoned").get(&key) {
        return Some(*metrics);
    }
    let catalog = catalog();
    let index = catalog.select(&key.0, bold, italic)?;
    let mut cache = cache.write().expect("font catalog cache poisoned");
    Some(*cache.entry(key).or_insert_with(|| {
        let mut metrics = catalog.metrics(index).clone();
        metrics.family = family.to_owned();
        Box::leak(Box::new(metrics))
    }))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_catalog_member_has_valid_metrics() {
        let catalog = catalog();
        assert!(!catalog.faces.is_empty());
        for i in 0..catalog.faces.len() {
            let face = &catalog.faces[i];
            let metrics = catalog.metrics(i);
            assert!(metrics.units_per_em > 0, "{}", face.key);
            assert!(
                !metrics.char_widths.is_empty(),
                "{} has no character map",
                face.key
            );
            assert!(metrics
                .char_widths
                .values()
                .all(|w| w.is_finite() && *w >= 0.0));
        }
    }

    #[test]
    fn localized_names_share_widths_without_an_installed_font() {
        let en = resolve("HGPSoeiPresenceEB", false, false).unwrap();
        let ja = resolve("HGP創英ﾌﾟﾚｾﾞﾝｽEB", false, false).unwrap();
        assert_eq!(en.units_per_em, 256);
        assert_eq!(en.char_widths, ja.char_widths);
        assert_eq!(en.char_widths[&'A'], 170.0 / 256.0);
        assert_eq!(en.char_widths.len(), 7484);
    }

    #[test]
    fn styles_select_distinct_faces_and_missing_names_do_not_substitute() {
        let regular = resolve("Gill Sans Nova", false, false).unwrap();
        let italic = resolve("Gill Sans Nova", false, true).unwrap();
        assert_ne!(regular.char_widths, italic.char_widths);
        assert!(resolve("Oxi nonexistent catalog fixture", false, false).is_none());
        assert!(resolve("MS-Mincho", false, false).is_none());
    }
}
