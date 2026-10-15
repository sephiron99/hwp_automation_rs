//! 팝업에서 확정한 후보의 선택 이력.
//!
//! 이 모듈은 Win32·HWP에 의존하지 않으므로 Windows가 아닌 환경에서도
//! `rustc --edition 2024 --test src/history.rs`로 단위 테스트를 실행할 수 있다.

use std::collections::BTreeMap;

#[derive(Debug, Clone, Copy)]
struct Entry {
    /// 확정한 횟수.
    count: u32,
    /// 마지막으로 확정한 순번. 클수록 최근이다.
    last: u64,
}

/// 확정한 단어별 횟수와 최근 순번을 기억하고, 후보 목록을 그 순서로 정렬한다.
///
/// 애드온이 로드되어 있는 동안만 유지하며 파일에 저장하지 않는다. 문서를
/// 구분하지 않으므로 한 문서에서 확정한 단어는 다른 문서에서도 우선한다.
/// 후보 자체는 매번 활성 문서에서 수집하므로, 이력에만 있고 문서에 없는
/// 단어가 후보로 나타나지는 않는다.
#[derive(Debug)]
pub(crate) struct SelectionHistory {
    entries: BTreeMap<String, Entry>,
    seq: u64,
}

impl SelectionHistory {
    /// 정적 플러그인 인스턴스에서 쓸 수 있도록 `const`로 만든다.
    pub(crate) const fn new() -> Self {
        Self {
            entries: BTreeMap::new(),
            seq: 0,
        }
    }

    /// `word`를 확정했다고 기록한다.
    pub(crate) fn record(&mut self, word: &str) {
        self.seq += 1;
        let seq = self.seq;
        self.entries
            .entry(word.to_string())
            .and_modify(|e| {
                e.count = e.count.saturating_add(1);
                e.last = seq;
            })
            .or_insert(Entry {
                count: 1,
                last: seq,
            });
    }

    /// 선택 이력이 있는 후보를 앞으로 옮긴다.
    ///
    /// 확정 횟수가 많은 후보가 먼저 오고, 횟수가 같으면 최근에 확정한 후보가
    /// 먼저 온다. 이력이 없는 후보는 그 뒤에 원래 순서(문서 등장 순서)를
    /// 유지한다.
    pub(crate) fn prioritize(&self, candidates: &mut [String]) {
        // sort_by_key는 안정 정렬이므로 같은 키의 후보는 원래 순서를 유지한다.
        // 이력이 없는 후보는 모두 (0, 0)이라 서로의 순서가 바뀌지 않는다.
        candidates.sort_by_key(|word| {
            let (count, last) = self.entries.get(word).map_or((0, 0), |e| (e.count, e.last));
            std::cmp::Reverse((count, last))
        });
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn words(list: &[&str]) -> Vec<String> {
        list.iter().map(|s| s.to_string()).collect()
    }

    #[test]
    fn no_history_keeps_document_order() {
        let history = SelectionHistory::new();
        let mut c = words(&["자동화", "자동완성", "자동차"]);
        history.prioritize(&mut c);
        assert_eq!(c, words(&["자동화", "자동완성", "자동차"]));
    }

    #[test]
    fn chosen_candidate_comes_first() {
        let mut history = SelectionHistory::new();
        history.record("자동완성");
        let mut c = words(&["자동화", "자동완성", "자동차"]);
        history.prioritize(&mut c);
        assert_eq!(c, words(&["자동완성", "자동화", "자동차"]));
    }

    #[test]
    fn more_frequent_before_less_frequent() {
        let mut history = SelectionHistory::new();
        history.record("자동차");
        history.record("자동차");
        history.record("자동완성");
        let mut c = words(&["자동화", "자동완성", "자동차"]);
        history.prioritize(&mut c);
        assert_eq!(c, words(&["자동차", "자동완성", "자동화"]));
    }

    #[test]
    fn same_count_prefers_most_recent() {
        let mut history = SelectionHistory::new();
        history.record("자동차");
        history.record("자동완성");
        let mut c = words(&["자동화", "자동차", "자동완성"]);
        history.prioritize(&mut c);
        assert_eq!(c, words(&["자동완성", "자동차", "자동화"]));
    }

    #[test]
    fn history_of_absent_word_does_not_add_candidate() {
        let mut history = SelectionHistory::new();
        history.record("rustc");
        let mut c = words(&["rustup", "rust-analyzer"]);
        history.prioritize(&mut c);
        assert_eq!(c, words(&["rustup", "rust-analyzer"]));
    }

    #[test]
    fn unchosen_candidates_keep_relative_order() {
        let mut history = SelectionHistory::new();
        history.record("d");
        let mut c = words(&["a", "b", "c", "d", "e"]);
        history.prioritize(&mut c);
        assert_eq!(c, words(&["d", "a", "b", "c", "e"]));
    }
}
