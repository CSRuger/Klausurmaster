"""Utility helpers that keep card data consistent throughout the app."""

from __future__ import annotations

from typing import Dict, List

DEFAULT_CARD_WEIGHT = 100.0

CardDict = Dict[str, object]


# Creates a normalized card dictionary that is JSON serializable.
def create_card(front: str, back: str = "", marked: bool = False, weight: float = DEFAULT_CARD_WEIGHT) -> CardDict:
    return {"front": front, "back": back, "marked": marked, "weight": float(weight)}


# Ensures any legacy card entries contain the expected keys.
def normalize_card_entry(card: object) -> CardDict:
    if isinstance(card, str):
        return create_card(card)
    if isinstance(card, dict):
        weight = card.get("weight", DEFAULT_CARD_WEIGHT)
        try:
            weight_value = float(weight)
        except (TypeError, ValueError):
            weight_value = DEFAULT_CARD_WEIGHT
        return create_card(
            front=str(card.get("front", "")),
            back=str(card.get("back", "")),
            marked=bool(card.get("marked", False)),
            weight=weight_value,
        )
    return create_card(str(card))


# Walks the nested row/column structure and normalizes every card entry.
def normalize_cards_tree(cards: Dict[str, Dict[str, List[CardDict]]]) -> None:
    for row_data in cards.values():
        for column_name, card_list in row_data.items():
            for idx, card in enumerate(card_list):
                card_list[idx] = normalize_card_entry(card)


# Finds and returns the first card dictionary that matches the provided front text.
def find_card(card_list: List[CardDict], front_text: str) -> CardDict | None:
    for card in card_list:
        if str(card.get("front")) == front_text:
            return card
    return None
