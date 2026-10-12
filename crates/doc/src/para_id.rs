//! Paragraph ids as Word keeps them ([MS-DOCX] 2.6.2 `w14:paraId` and `w14:textId`): every
//! paragraph is named by an id that stays with it through edits and saves, so other tools (and
//! agents) can refer to it across versions of the file.

use std::collections::HashSet;
use std::hash::{BuildHasher, RandomState};
use std::sync::Arc;

use serde::{Deserialize, Serialize};

use crate::{Block, Blocks, Document, Paragraph};

/// A paragraph's ids: `para` names the paragraph for as long as it exists and is unique in the
/// document; `text` names a version of its text. Word reads only values below [`ParaId::LIMIT`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct ParaId {
    pub para: u32,
    pub text: u32,
}

impl ParaId {
    /// Both ids are below this ([MS-DOCX] 2.6.2).
    pub const LIMIT: u32 = 0x8000_0000;
}

/// Paragraphs nested deeper than this in table cells are left as they are.
const MAX_DEPTH: usize = 32;

/// One paragraph that needs ids: its index in document order, and whether its `para` id must
/// change too (missing, out of range or held by an earlier paragraph) or only its `text` id.
type Fix = (usize, bool);

impl Document {
    /// Whether every paragraph has usable ids of its own (see [`Document::assign_para_ids`]).
    pub fn para_ids_complete(&self) -> bool {
        self.para_id_fixes().0.is_empty()
    }

    /// Give every paragraph without usable ids fresh ones: a paragraph with none, ids out of range,
    /// or the `para` id of a paragraph earlier in the document (the body, then each part — headers,
    /// footers, notes, comments, text boxes — then the note separators). The first holder of an id
    /// keeps it, so a paragraph split in two keeps its id in the first half. Fresh `para` ids are
    /// unique in the document. Returns how many paragraphs got new ids.
    pub fn assign_para_ids(&mut self) -> usize {
        let (fixes, mut taken) = self.para_id_fixes();
        if fixes.is_empty() {
            return 0;
        }
        let mut fresh = Fresh::default();
        let mut at = 0;
        for blocks in self.stories_mut() {
            fix_blocks(blocks, 0, &mut at, &fixes, &mut taken, &mut fresh);
        }
        fixes.len()
    }

    /// The paragraphs that need ids, in document order, and the `para` ids kept.
    fn para_id_fixes(&self) -> (Vec<Fix>, HashSet<u32>) {
        let mut taken = HashSet::new();
        let mut fixes = Vec::new();
        let mut at = 0;
        for blocks in self.stories() {
            visit(blocks, 0, &mut |p| {
                let keep_para = p.id.is_some_and(|id| id.para < ParaId::LIMIT && taken.insert(id.para));
                if !keep_para || p.id.is_none_or(|id| id.text >= ParaId::LIMIT) {
                    fixes.push((at, !keep_para));
                }
                at += 1;
            });
        }
        (fixes, taken)
    }

    /// Every story's blocks, in document order.
    fn stories(&self) -> Vec<&Blocks> {
        let mut v = vec![&self.body];
        v.extend(self.parts.values().map(|p| &p.blocks));
        for s in [&self.footnote_separators, &self.endnote_separators] {
            v.extend([&s.separator, &s.continuation_separator, &s.continuation_notice].into_iter().flatten());
        }
        v
    }

    /// [`Document::stories`], to change.
    fn stories_mut(&mut self) -> Vec<&mut Blocks> {
        let mut v = vec![&mut self.body];
        v.extend(self.parts.values_mut().map(|p| &mut p.blocks));
        for s in [&mut self.footnote_separators, &mut self.endnote_separators] {
            v.extend([&mut s.separator, &mut s.continuation_separator, &mut s.continuation_notice].into_iter().flatten());
        }
        v
    }
}

/// Forget the ids of every paragraph in `blocks` (pasted copies are new paragraphs).
pub fn clear_para_ids(blocks: &mut [Block]) {
    fn clear(b: &mut Block, depth: usize) {
        match b {
            Block::Para(p) => p.id = None,
            Block::Table(t) if depth < MAX_DEPTH => {
                for cell in t.rows.iter_mut().flat_map(|r| r.cells.iter_mut()) {
                    for b in &mut cell.blocks {
                        clear(Arc::make_mut(b), depth + 1);
                    }
                }
            }
            Block::Table(_) => {}
        }
    }
    for b in blocks {
        clear(b, 0);
    }
}

fn visit(blocks: &Blocks, depth: usize, f: &mut dyn FnMut(&Paragraph)) {
    for b in blocks {
        match &**b {
            Block::Para(p) => f(p),
            Block::Table(t) if depth < MAX_DEPTH => {
                for cell in t.rows.iter().flat_map(|r| &r.cells) {
                    visit(&cell.blocks, depth + 1, f);
                }
            }
            Block::Table(_) => {}
        }
    }
}

/// The paragraphs [`visit`] reaches in `b`.
fn count(b: &Block, depth: usize) -> usize {
    match b {
        Block::Para(_) => 1,
        Block::Table(t) if depth < MAX_DEPTH => {
            t.rows.iter().flat_map(|r| &r.cells).flat_map(|c| &c.blocks).map(|b| count(b, depth + 1)).fold(0, usize::saturating_add)
        }
        Block::Table(_) => 0,
    }
}

/// Give the paragraphs [`Document::para_id_fixes`] found new ids, copying only the blocks that
/// hold one (the rest stay shared with undo history).
fn fix_blocks(blocks: &mut Blocks, depth: usize, at: &mut usize, fixes: &[Fix], taken: &mut HashSet<u32>, fresh: &mut Fresh) {
    for b in blocks {
        let n = count(b, depth);
        let first = fixes.partition_point(|(i, _)| *i < *at);
        if fixes.get(first).is_none_or(|(i, _)| *i >= at.saturating_add(n)) {
            *at = at.saturating_add(n);
            continue;
        }
        match Arc::make_mut(b) {
            Block::Para(p) => {
                let new_para = fixes.get(first).is_some_and(|(_, new_para)| *new_para);
                let para = match p.id {
                    Some(id) if !new_para => id.para,
                    _ => fresh.unique(taken),
                };
                p.id = Some(ParaId { para, text: fresh.any() });
                *at += 1;
            }
            Block::Table(t) => {
                for cell in t.rows.iter_mut().flat_map(|r| r.cells.iter_mut()) {
                    fix_blocks(&mut cell.blocks, depth + 1, at, fixes, taken, fresh);
                }
            }
        }
    }
}

/// Fresh ids: random (as Word's are), so a paragraph added later never takes the id of one
/// deleted earlier.
#[derive(Default)]
struct Fresh {
    seed: RandomState,
    n: u64,
}

impl Fresh {
    /// An id in range, never 0.
    fn any(&mut self) -> u32 {
        for _ in 0..64 {
            self.n = self.n.wrapping_add(1);
            let v = (self.seed.hash_one(self.n) as u32) & (ParaId::LIMIT - 1);
            if v != 0 {
                return v;
            }
        }
        1
    }

    /// An id in range that no paragraph holds yet (recorded in `taken`).
    fn unique(&mut self, taken: &mut HashSet<u32>) -> u32 {
        for _ in 0..64 {
            let v = self.any();
            if taken.insert(v) {
                return v;
            }
        }
        // A document with most of the 2³¹ ids in use: the first free one.
        (1..ParaId::LIMIT).find(|v| taken.insert(*v)).unwrap_or(1)
    }
}
