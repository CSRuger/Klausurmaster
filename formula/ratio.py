"""Calculations that aggregate statistics for entire rows of cards."""

from __future__ import annotations

from typing import Dict, List

from cards.utils import CardDict

CardRow = Dict[str, List[CardDict]]
CardMatrix = Dict[str, CardRow]


def _effective_grade_value(raw_value: float) -> float:
    return 6.0 if abs(raw_value - 5.0) < 1e-6 else raw_value


# Calculates how well a row performs by weighting cards per column.
def calculate_ratio(row_name: str, cards: CardMatrix, columns: list[str]) -> float:
    if row_name not in cards:
        return 0.0

    total_cards = sum(len(cards[row_name][col]) for col in columns)
    if total_cards == 0:
        return 0.0

    sum_factors = 0.0
    for col in columns:
        try:
            col_val = _effective_grade_value(float(col))
        except (TypeError, ValueError):
            continue
        count = len(cards[row_name][col])
        normalized = (5.0 - col_val) / 3.0
        sum_factors += count * normalized

    return sum_factors / total_cards


def calculate_expected_grade(row_name: str, cards: CardMatrix, columns: list[str]) -> float:
    if row_name not in cards:
        return 0.0

    weighted_sum = 0.0
    total_weight = 0.0
    for col in columns:
        try:
            col_value = _effective_grade_value(float(col))
        except (TypeError, ValueError):
            continue
        for card in cards[row_name][col]:
            try:
                weight = float(card.get("weight", 100.0))
            except (TypeError, ValueError):
                weight = 100.0
            if weight <= 0:
                continue
            normalized_weight = weight / 100.0
            total_weight += normalized_weight
            weighted_sum += col_value * normalized_weight

    if total_weight == 0:
        return 0.0
    return weighted_sum / total_weight
