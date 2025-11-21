"""Helpers that control how the table should be structured."""


# Builds the column values used for the Leitner-like progression table.
def generate_columns(num_columns: int) -> list[str]:
    start_value = 5.0
    end_value = 1.0
    diff = start_value - end_value
    steps = num_columns - 1 if num_columns > 1 else 1
    step = diff / steps

    columns: list[str] = []
    for idx in range(num_columns):
        val = start_value - idx * step
        columns.append(f"{val:.1f}")
    return columns
