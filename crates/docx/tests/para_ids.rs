//! Paragraph ids (`w14:paraId`, `w14:textId`, [MS-DOCX] 2.6.2): every paragraph keeps its own
//! through an edit and a save, and paragraphs without a usable one get a fresh, unique id below
//! 0x80000000. All markup here is synthetic.

use std::collections::HashSet;
use std::io::{Read, Write};

use wordcraft_doc::Pos;
use wordcraft_doc::props::CharProps;

const NS: &str =
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;
const ROOT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;

fn package(body: &str) -> Vec<u8> {
    let doc = format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#);
    let mut zw = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for (name, bytes) in [("_rels/.rels", ROOT_RELS.as_bytes()), ("word/document.xml", doc.as_bytes())] {
        zw.start_file(name, zip::write::SimpleFileOptions::default()).unwrap();
        zw.write_all(bytes).unwrap();
    }
    zw.finish().unwrap().into_inner()
}

/// A paragraph with ids (`None`: the attribute is left out).
fn p(para_id: Option<&str>, text_id: Option<&str>, text: &str) -> String {
    let mut attrs = String::new();
    if let Some(id) = para_id {
        attrs.push_str(&format!(r#" w14:paraId="{id}""#));
    }
    if let Some(id) = text_id {
        attrs.push_str(&format!(r#" w14:textId="{id}""#));
    }
    format!(r#"<w:p{attrs}><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#)
}

/// Each `w:p` in the saved `word/document.xml`, in order: (paraId, textId, text).
fn saved_paras(bytes: &[u8]) -> Vec<(Option<String>, Option<String>, String)> {
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    zip.by_name("word/document.xml").unwrap().read_to_string(&mut xml).unwrap();
    let attr = |tag: &str, name: &str| {
        let at = tag.find(&format!(r#"{name}=""#))? + name.len() + 2;
        tag[at..].split('"').next().map(str::to_string)
    };
    let mut out = Vec::new();
    for (start, _) in xml.match_indices("<w:p") {
        let rest = &xml[start..];
        if !matches!(rest.as_bytes().get(4), Some(b' ' | b'>' | b'/')) {
            continue;
        }
        let tag = &rest[..rest.find('>').unwrap()];
        let end = rest.find("</w:p>").unwrap_or(rest.len());
        let text: String =
            rest[..end].split("<w:t").skip(1).filter_map(|t| t.split_once('>').map(|(_, r)| r.split('<').next().unwrap_or(""))).collect();
        out.push((attr(tag, "w14:paraId"), attr(tag, "w14:textId"), text));
    }
    out
}

/// A usable id: eight hex digits, below 0x80000000.
fn usable(id: &Option<String>) -> bool {
    id.as_deref().is_some_and(|s| s.len() == 8 && u32::from_str_radix(s, 16).is_ok_and(|v| v < 0x8000_0000))
}

#[test]
fn paragraphs_keep_their_ids_through_an_edit_and_new_ones_get_fresh_ids() {
    let cell = format!("<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>", p(Some("0C0FFEE0"), Some("1BADB002"), "Cell"));
    let body = [
        p(Some("1A2B3C4D"), Some("77777777"), "Alpha beta"),
        p(Some("00000ABC"), Some("2468ACE0"), "Gamma"),
        cell,
        p(Some("7FFFFFFE"), Some("13579BDF"), "Delta"),
    ]
    .concat();
    let mut doc = wordcraft_docx::read(&package(&body)).unwrap();
    // Enter after "Alpha " makes a new paragraph; "Delta" is edited in place.
    doc.split_paragraph(&Pos::body(0, "Alpha ".len())).unwrap();
    doc.insert_text(&Pos::body(4, "Delta".len()), " and more", &CharProps::default()).unwrap();
    let saved = wordcraft_docx::write(&doc).unwrap();

    let paras = saved_paras(&saved);
    let ids: Vec<(Option<&str>, Option<&str>, &str)> = paras.iter().map(|(p, t, x)| (p.as_deref(), t.as_deref(), x.as_str())).collect();
    assert_eq!(ids.len(), 5, "{ids:?}");
    // Untouched paragraphs keep both ids; edited ones keep their paraId and get a new textId.
    assert_eq!((ids[0].0, ids[0].2), (Some("1A2B3C4D"), "Alpha "));
    assert_ne!(ids[0].1, Some("77777777"), "its text changed");
    assert_eq!(ids[1].2, "beta");
    assert_eq!(ids[2], (Some("00000ABC"), Some("2468ACE0"), "Gamma"));
    assert_eq!(ids[3], (Some("0C0FFEE0"), Some("1BADB002"), "Cell"));
    assert_eq!((ids[4].0, ids[4].2), (Some("7FFFFFFE"), "Delta and more"));
    assert_ne!(ids[4].1, Some("13579BDF"), "its text changed");
    // The new paragraph has an id of its own.
    assert!(usable(&paras[1].0) && usable(&paras[1].1), "{:?}", paras[1]);
    let unique: HashSet<_> = paras.iter().map(|p| p.0.clone()).collect();
    assert_eq!(unique.len(), paras.len(), "{ids:?}");

    // Reopened, every paragraph still has the id it was saved with.
    let again = wordcraft_docx::write(&wordcraft_docx::read(&saved).unwrap()).unwrap();
    assert_eq!(saved_paras(&again), paras);
}

#[test]
fn missing_duplicate_and_out_of_range_ids_are_replaced() {
    let body = [
        p(Some("12345678"), Some("77777777"), "First"),
        p(Some("12345678"), Some("77777777"), "Copy"),
        p(Some("80000000"), Some("FFFFFFFF"), "Too big"),
        p(Some("not hex"), None, "Junk"),
        p(None, None, "None"),
    ]
    .concat();
    let saved = wordcraft_docx::write(&wordcraft_docx::read(&package(&body)).unwrap()).unwrap();
    let paras = saved_paras(&saved);
    assert_eq!(paras.len(), 5);
    // The first holder of an id keeps it.
    assert_eq!(paras[0].0.as_deref(), Some("12345678"));
    assert!(paras.iter().all(|p| usable(&p.0) && usable(&p.1)), "{paras:?}");
    let unique: HashSet<_> = paras.iter().map(|p| p.0.clone()).collect();
    assert_eq!(unique.len(), 5, "{paras:?}");
}
