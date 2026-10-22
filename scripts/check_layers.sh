#!/usr/bin/env bash
# hwp_core의 층 의존 방향을 검사한다: 기반 ← pset ← method ← actions ← ext
# 아래 층이 위 층을 참조하면 실패한다. 규칙은 hwp_core.md 참고.
set -u
cd "$(dirname "$0")/../crates/hwp_core/src" || exit 2

fail=0
check() { # check <dir...> -- <금지 모듈 정규식>
    local dirs=()
    while [ "$1" != "--" ]; do dirs+=("$1"); shift; done
    shift
    # 주석(문서 링크 포함)은 의존이 아니므로 제외한다.
    if rg -n "(crate|hwp_core)::($1)\b" "${dirs[@]}" | rg -v '^[^:]+:[0-9]+:\s*//'; then
        echo "^ 층 위반: ${dirs[*]} 에서 $1 참조" >&2
        fail=1
    fi
}

base=(com_util.rs disp_obj.rs error.rs hwp_obj.rs hwp_ver.rs pos.rs variant.rs)
check "${base[@]}" -- 'pset|method|actions|ext'
check pset -- 'method|actions|ext'
check method -- 'actions|ext'
check actions -- 'ext'

[ "$fail" -eq 0 ] && echo "층 검사 통과"
exit "$fail"
