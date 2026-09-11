from __future__ import annotations

import json
import math
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field, replace
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from .semantic import (
    ProtectedText,
    TranslationConstraints,
    TranslationContext,
    TranslationQualityReport,
    TranslationService,
    TranslationUnit,
)

_WORD_RE = re.compile(r"[^\W_]+", re.UNICODE)
_DEFAULT_CORPUS = Path(__file__).with_name("evaluation_corpus.json")


def _norm_text(text: str) -> str:
    value = unicodedata.normalize("NFKC", str(text or "")).casefold()
    value = re.sub(r"\s+", " ", value).strip()
    # Typography should not dominate a linguistic regression score.
    value = value.translate(str.maketrans({
        "’": "'",
        "‘": "'",
        "“": '"',
        "”": '"',
        "–": "-",
        "—": "-",
        "…": "...",
    }))
    return value


def _word_tokens(text: str) -> list[str]:
    return _WORD_RE.findall(_norm_text(text))


def _counter_f1(a: Counter[str], b: Counter[str], *, beta: float = 1.0) -> float:
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    overlap = sum((a & b).values())
    precision = overlap / sum(a.values())
    recall = overlap / sum(b.values())
    if precision <= 0.0 or recall <= 0.0:
        return 0.0
    beta2 = beta * beta
    return (1.0 + beta2) * precision * recall / (beta2 * precision + recall)


def token_f1(candidate: str, reference: str) -> float:
    return _counter_f1(Counter(_word_tokens(candidate)), Counter(_word_tokens(reference)))


def character_ngram_f_score(
    candidate: str,
    reference: str,
    *,
    max_order: int = 6,
    beta: float = 2.0,
) -> float:
    """Return a small dependency-free chrF-like character n-gram score.

    This is intentionally named descriptively rather than claiming bit-for-bit
    compatibility with sacreBLEU chrF. It is stable enough for local regression
    comparisons and requires only the Python standard library.
    """

    cand = _norm_text(candidate)
    ref = _norm_text(reference)
    if cand == ref:
        return 1.0
    if not cand or not ref:
        return 0.0

    order_scores: list[float] = []
    usable_orders = min(max_order, max(len(cand), len(ref)))
    for order in range(1, usable_orders + 1):
        cand_grams = Counter(cand[i : i + order] for i in range(max(0, len(cand) - order + 1)))
        ref_grams = Counter(ref[i : i + order] for i in range(max(0, len(ref) - order + 1)))
        if not cand_grams and not ref_grams:
            continue
        order_scores.append(_counter_f1(cand_grams, ref_grams, beta=beta))
    return sum(order_scores) / len(order_scores) if order_scores else 0.0


def sequence_similarity(candidate: str, reference: str) -> float:
    return SequenceMatcher(None, _norm_text(candidate), _norm_text(reference)).ratio()


@dataclass(frozen=True)
class EvaluationCase:
    id: str
    source: str
    source_lang: str
    target_lang: str
    references: tuple[str, ...]
    role: str = "body"
    mode: str = "natural"
    constraints: TranslationConstraints = field(default_factory=TranslationConstraints)
    context: TranslationContext = field(default_factory=TranslationContext)
    protected: ProtectedText | None = None
    tags: tuple[str, ...] = ()
    notes: str | None = None

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "EvaluationCase":
        case_id = str(data.get("id") or "").strip()
        source = str(data.get("source") or "")
        source_lang = str(data.get("source_lang") or "auto").strip() or "auto"
        target_lang = str(data.get("target_lang") or "").strip()
        references = tuple(str(item) for item in data.get("references", ()) if str(item).strip())
        if not case_id:
            raise ValueError("evaluation case is missing id")
        if not source.strip():
            raise ValueError(f"evaluation case {case_id!r} has empty source")
        if not target_lang:
            raise ValueError(f"evaluation case {case_id!r} is missing target_lang")
        if not references:
            raise ValueError(f"evaluation case {case_id!r} has no references")

        raw_constraints = data.get("constraints") or {}
        constraints = TranslationConstraints(
            max_chars=raw_constraints.get("max_chars"),
            max_lines=raw_constraints.get("max_lines"),
            source_visible_chars=raw_constraints.get("source_visible_chars"),
            source_line_count=raw_constraints.get("source_line_count"),
        )
        raw_context = data.get("context") or {}
        context = TranslationContext.from_values(
            domain=raw_context.get("domain"),
            section=raw_context.get("section"),
            before=raw_context.get("before") or (),
            after=raw_context.get("after") or (),
            metadata=raw_context.get("metadata") or {},
            terminology=raw_context.get("terminology") or {},
        )

        protected: ProtectedText | None = None
        raw_protected = data.get("protected")
        if isinstance(raw_protected, Mapping):
            masked = str(raw_protected.get("text") or source)
            token_map = {
                str(key): str(value)
                for key, value in (raw_protected.get("token_map") or {}).items()
            }
            protected = ProtectedText(masked, token_map)

        return cls(
            id=case_id,
            source=source,
            source_lang=source_lang,
            target_lang=target_lang,
            references=references,
            role=str(data.get("role") or "body"),
            mode=str(data.get("mode") or "natural"),
            constraints=constraints,
            context=context,
            protected=protected,
            tags=tuple(str(tag) for tag in data.get("tags", ()) if str(tag).strip()),
            notes=(str(data.get("notes")).strip() if data.get("notes") else None),
        )

    def to_unit(self) -> TranslationUnit:
        return TranslationUnit(
            text=self.source,
            source_lang=self.source_lang,
            role=self.role,
            mode=self.mode,  # type: ignore[arg-type]
            constraints=self.constraints,
            context=self.context,
            protected=self.protected,
        )


@dataclass(frozen=True)
class CaseScore:
    case_id: str
    tags: tuple[str, ...]
    candidate: str | None
    best_reference: str
    exact_match: bool
    reference_similarity: float
    char_similarity: float
    token_similarity: float
    sequence_similarity: float
    regression_score: float
    quality: TranslationQualityReport

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["quality"] = asdict(self.quality)
        return data


@dataclass(frozen=True)
class EvaluationSummary:
    cases: int
    scored_cases: int
    exact_matches: int
    mean_reference_similarity: float
    mean_regression_score: float
    terminology_score: float | None
    hard_failure_rate: float
    constraint_warning_rate: float


@dataclass(frozen=True)
class EvaluationReport:
    name: str
    scores: tuple[CaseScore, ...]
    summary: EvaluationSummary
    service_metrics: Mapping[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "summary": asdict(self.summary),
            "service_metrics": dict(self.service_metrics),
            "scores": [score.to_dict() for score in self.scores],
        }

    def to_json(self, *, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)


@dataclass(frozen=True)
class ComparisonRow:
    name: str
    mean_regression_score: float
    mean_reference_similarity: float
    exact_matches: int
    hard_failure_rate: float
    terminology_score: float | None


def load_evaluation_corpus(path: str | Path | None = None) -> list[EvaluationCase]:
    corpus_path = Path(path) if path is not None else _DEFAULT_CORPUS
    payload = json.loads(corpus_path.read_text(encoding="utf-8"))
    raw_cases = payload.get("cases") if isinstance(payload, Mapping) else payload
    if not isinstance(raw_cases, list):
        raise ValueError("evaluation corpus must contain a 'cases' list")
    cases = [EvaluationCase.from_dict(item) for item in raw_cases]
    ids = [case.id for case in cases]
    if len(ids) != len(set(ids)):
        raise ValueError("evaluation corpus contains duplicate case ids")
    return cases


def _score_against_reference(candidate: str, reference: str) -> tuple[float, float, float, float]:
    char_score = character_ngram_f_score(candidate, reference)
    tok_score = token_f1(candidate, reference)
    seq_score = sequence_similarity(candidate, reference)
    # Character overlap is most robust for morphology/no-space languages; token
    # and sequence similarity provide secondary signals. This is a regression
    # heuristic, not a human semantic-quality score.
    combined = 0.60 * char_score + 0.25 * tok_score + 0.15 * seq_score
    return combined, char_score, tok_score, seq_score


def score_case(case: EvaluationCase, candidate: str | None) -> CaseScore:
    text = "" if candidate is None else str(candidate)
    candidates: list[tuple[float, float, float, float, str]] = []
    for reference in case.references:
        combined, char_score, tok_score, seq_score = _score_against_reference(text, reference)
        candidates.append((combined, char_score, tok_score, seq_score, reference))
    combined, char_score, tok_score, seq_score, best_reference = max(candidates, key=lambda item: item[0])

    quality_unit = case.to_unit()
    # TranslationService returns restored visible text. Raw protected-token parity
    # is already validated inside the service and cannot be reconstructed from a
    # saved final-output file, so corpus scoring evaluates the visible form.
    if quality_unit.protected is not None:
        quality_unit = replace(quality_unit, protected=None)
    quality = quality_unit.evaluate_translation(
        candidate,
        case.target_lang,
        enforce_terminology=bool(case.context.terminology),
        detect_source_echo=True,
    )
    exact = any(_norm_text(text) == _norm_text(reference) for reference in case.references)

    invariant_factor = 1.0
    if quality.hard_failures:
        invariant_factor = 0.0
    elif any(item in {"max-chars-exceeded", "max-lines-exceeded"} for item in quality.warnings):
        invariant_factor = 0.90
    regression_score = 100.0 * combined * invariant_factor

    return CaseScore(
        case_id=case.id,
        tags=case.tags,
        candidate=candidate,
        best_reference=best_reference,
        exact_match=exact,
        reference_similarity=combined,
        char_similarity=char_score,
        token_similarity=tok_score,
        sequence_similarity=seq_score,
        regression_score=regression_score,
        quality=quality,
    )


def _summarize(scores: Sequence[CaseScore]) -> EvaluationSummary:
    if not scores:
        return EvaluationSummary(0, 0, 0, 0.0, 0.0, None, 0.0, 0.0)
    scored = [score for score in scores if score.candidate is not None]
    count = len(scores)
    terminology_required = sum(score.quality.terminology_required for score in scores)
    terminology_matched = sum(score.quality.terminology_matched for score in scores)
    terminology_score = (
        terminology_matched / terminology_required if terminology_required else None
    )
    hard_failures = sum(bool(score.quality.hard_failures) for score in scores)
    constraint_warnings = sum(
        any(item in {"max-chars-exceeded", "max-lines-exceeded"} for item in score.quality.warnings)
        for score in scores
    )
    return EvaluationSummary(
        cases=count,
        scored_cases=len(scored),
        exact_matches=sum(score.exact_match for score in scores),
        mean_reference_similarity=(
            sum(score.reference_similarity for score in scores) / count
        ),
        mean_regression_score=sum(score.regression_score for score in scores) / count,
        terminology_score=terminology_score,
        hard_failure_rate=hard_failures / count,
        constraint_warning_rate=constraint_warnings / count,
    )


def evaluate_outputs(
    cases: Sequence[EvaluationCase],
    outputs: Mapping[str, str | None] | Sequence[str | None],
    *,
    name: str = "candidate",
    service_metrics: Mapping[str, Any] | None = None,
) -> EvaluationReport:
    if isinstance(outputs, Mapping):
        candidates = [outputs.get(case.id) for case in cases]
    else:
        candidates = list(outputs)
        if len(candidates) != len(cases):
            raise ValueError("candidate output count does not match evaluation case count")
    scores = tuple(score_case(case, candidate) for case, candidate in zip(cases, candidates))
    return EvaluationReport(
        name=name,
        scores=scores,
        summary=_summarize(scores),
        service_metrics=dict(service_metrics or {}),
    )


def run_translator_evaluation(
    translator: Any,
    cases: Sequence[EvaluationCase],
    *,
    name: str | None = None,
    fallback_translator: Any | None = None,
    cancel_event: Any | None = None,
) -> EvaluationReport:
    outputs: list[str | None] = [None] * len(cases)
    aggregate_metrics: Counter[str] = Counter()
    by_target: dict[str, list[int]] = defaultdict(list)
    for idx, case in enumerate(cases):
        by_target[case.target_lang].append(idx)

    # A fresh service per target language avoids leaking target-specific backend
    # cache assumptions while preserving batching/dedupe within each language.
    for target_lang, indices in by_target.items():
        service = TranslationService(translator, fallback_translator=fallback_translator)
        result = service.translate_units(
            [cases[idx].to_unit() for idx in indices],
            target_lang,
            cancel_event=cancel_event,
        )
        for idx, text in zip(indices, result.texts):
            outputs[idx] = text
        for key, value in asdict(result.metrics).items():
            if isinstance(value, (int, float)):
                aggregate_metrics[key] += value

    report_name = name or getattr(translator, "_engine_name", None) or type(translator).__name__
    return evaluate_outputs(
        cases,
        outputs,
        name=str(report_name),
        service_metrics=dict(aggregate_metrics),
    )



@dataclass(frozen=True)
class EvaluationThresholds:
    min_mean_regression_score: float = 0.0
    max_hard_failure_rate: float = 1.0
    min_terminology_score: float | None = None
    max_constraint_warning_rate: float = 1.0


@dataclass(frozen=True)
class EvaluationGateResult:
    passed: bool
    failures: tuple[str, ...] = ()


def check_quality_gate(
    report: EvaluationReport,
    thresholds: EvaluationThresholds,
) -> EvaluationGateResult:
    failures: list[str] = []
    summary = report.summary
    if summary.mean_regression_score < thresholds.min_mean_regression_score:
        failures.append(
            f"mean-regression-score {summary.mean_regression_score:.2f} < "
            f"{thresholds.min_mean_regression_score:.2f}"
        )
    if summary.hard_failure_rate > thresholds.max_hard_failure_rate:
        failures.append(
            f"hard-failure-rate {summary.hard_failure_rate:.4f} > "
            f"{thresholds.max_hard_failure_rate:.4f}"
        )
    if thresholds.min_terminology_score is not None:
        score = summary.terminology_score
        if score is None or score < thresholds.min_terminology_score:
            rendered = "none" if score is None else f"{score:.4f}"
            failures.append(
                f"terminology-score {rendered} < {thresholds.min_terminology_score:.4f}"
            )
    if summary.constraint_warning_rate > thresholds.max_constraint_warning_rate:
        failures.append(
            f"constraint-warning-rate {summary.constraint_warning_rate:.4f} > "
            f"{thresholds.max_constraint_warning_rate:.4f}"
        )
    return EvaluationGateResult(passed=not failures, failures=tuple(failures))


def summarize_by_tag(report: EvaluationReport) -> dict[str, EvaluationSummary]:
    grouped: dict[str, list[CaseScore]] = defaultdict(list)
    for score in report.scores:
        for tag in score.tags:
            grouped[tag].append(score)
    return {tag: _summarize(scores) for tag, scores in sorted(grouped.items())}

def compare_reports(reports: Iterable[EvaluationReport]) -> list[ComparisonRow]:
    rows = [
        ComparisonRow(
            name=report.name,
            mean_regression_score=report.summary.mean_regression_score,
            mean_reference_similarity=report.summary.mean_reference_similarity,
            exact_matches=report.summary.exact_matches,
            hard_failure_rate=report.summary.hard_failure_rate,
            terminology_score=report.summary.terminology_score,
        )
        for report in reports
    ]
    return sorted(
        rows,
        key=lambda row: (
            -row.mean_regression_score,
            row.hard_failure_rate,
            -row.exact_matches,
            row.name,
        ),
    )


def filter_cases(
    cases: Iterable[EvaluationCase],
    *,
    tags: Iterable[str] = (),
    source_lang: str | None = None,
    target_lang: str | None = None,
) -> list[EvaluationCase]:
    wanted_tags = {str(tag) for tag in tags if str(tag)}
    out: list[EvaluationCase] = []
    for case in cases:
        if source_lang and case.source_lang != source_lang:
            continue
        if target_lang and case.target_lang != target_lang:
            continue
        if wanted_tags and not wanted_tags.issubset(set(case.tags)):
            continue
        out.append(case)
    return out
