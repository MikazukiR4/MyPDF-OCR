"""
Character Error Rate (CER) and Word Error Rate (WER) Evaluation Script
======================================================================
Provides benchmark metrics to quantitatively measure Thai OCR accuracy.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List, Tuple


def levenshtein_distance(ref: str | list[str], hyp: str | list[str]) -> int:
    """Compute the minimum edit distance (insertions, deletions, substitutions)."""
    n, m = len(ref), len(hyp)
    if n == 0:
        return m
    if m == 0:
        return n

    # Dynamic programming with O(min(n, m)) memory
    prev_row = list(range(m + 1))
    for i, r_char in enumerate(ref):
        curr_row = [i + 1] + [0] * m
        for j, h_char in enumerate(hyp):
            cost = 0 if r_char == h_char else 1
            curr_row[j + 1] = min(
                curr_row[j] + 1,       # insertion
                prev_row[j + 1] + 1,   # deletion
                prev_row[j] + cost,    # substitution
            )
        prev_row = curr_row

    return prev_row[m]


def calculate_cer(reference: str, hypothesis: str) -> float:
    """
    Calculate Character Error Rate (CER).
    CER = (Substitutions + Deletions + Insertions) / len(reference)
    Returns a float between 0.0 and infinity (0.0 = perfect match).
    """
    ref_chars = list(reference.strip())
    hyp_chars = list(hypothesis.strip())
    if not ref_chars:
        return 0.0 if not hyp_chars else 1.0

    distance = levenshtein_distance(ref_chars, hyp_chars)
    return distance / len(ref_chars)


def calculate_wer(reference: str, hypothesis: str) -> float:
    """
    Calculate Word Error Rate (WER).
    Tokenizes by whitespace (or pythainlp if available).
    """
    try:
        from pythainlp.tokenize import word_tokenize
        ref_words = word_tokenize(reference.strip())
        hyp_words = word_tokenize(hypothesis.strip())
    except ImportError:
        ref_words = reference.strip().split()
        hyp_words = hypothesis.strip().split()

    if not ref_words:
        return 0.0 if not hyp_words else 1.0

    distance = levenshtein_distance(ref_words, hyp_words)
    return distance / len(ref_words)


def evaluate_pairs(pairs: List[Tuple[str, str, str]]) -> Dict[str, object]:
    """
    Evaluate list of (sample_id, reference_text, hypothesis_text).
    """
    results = []
    total_ref_chars = 0
    total_char_edits = 0

    for sample_id, ref, hyp in pairs:
        cer = calculate_cer(ref, hyp)
        wer = calculate_wer(ref, hyp)
        ref_len = len(ref.strip())
        char_edits = round(cer * ref_len)

        total_ref_chars += ref_len
        total_char_edits += char_edits

        results.append({
            "id": sample_id,
            "reference_length": ref_len,
            "hypothesis_length": len(hyp.strip()),
            "cer": round(cer, 4),
            "wer": round(wer, 4),
        })

    overall_cer = total_char_edits / max(1, total_ref_chars)
    return {
        "overall_cer": round(overall_cer, 4),
        "total_samples": len(pairs),
        "total_reference_characters": total_ref_chars,
        "samples": results,
    }


def main():
    parser = argparse.ArgumentParser(description="Evaluate OCR CER and WER against Ground Truth")
    parser.add_argument("--ref", type=str, help="Path to ground truth text file or directory")
    parser.add_argument("--hyp", type=str, help="Path to OCR output text file or directory")
    args = parser.parse_args()

    if not args.ref or not args.hyp:
        print("Usage: python evaluate_cer.py --ref <ground_truth.txt> --hyp <ocr_result.txt>")
        sys.exit(1)

    ref_path = Path(args.ref)
    hyp_path = Path(args.hyp)

    if ref_path.is_file() and hyp_path.is_file():
        ref_text = ref_path.read_text(encoding="utf-8")
        hyp_text = hyp_path.read_text(encoding="utf-8")
        cer = calculate_cer(ref_text, hyp_text)
        wer = calculate_wer(ref_text, hyp_text)
        print(f"Sample: {ref_path.name}")
        print(f"  CER: {cer * 100:.2f}%")
        print(f"  WER: {wer * 100:.2f}%")


if __name__ == "__main__":
    main()
