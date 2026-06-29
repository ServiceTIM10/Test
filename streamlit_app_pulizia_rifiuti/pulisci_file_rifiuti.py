from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


REQUIRED_COLUMNS = ("Nr. Doc.", "MC", "Kg")

# Excel salva i formati numerici in convenzione internazionale.
# In Excel con impostazioni italiane questi formati vengono visualizzati come:
#   #,##0    -> 1.000
#   #,##0.00 -> 1.125,00
NR_DOC_NUMBER_FORMAT = "#,##0"
DECIMAL_NUMBER_FORMAT = "#,##0.00"


@dataclass(frozen=True)
class DocCorrection:
    first_excel_row: int
    last_excel_row: int
    original_value: int
    corrected_value: int
    previous_distinct_value: int | None
    next_distinct_value: int | None
    rule: str
    affected_rows: int


def normalize_header(value: Any) -> str:
    return "" if value is None else str(value).strip()


@dataclass(frozen=True)
class HeaderDetectionResult:
    columns: dict[str, int]
    header_row: int


def scan_header_row(ws: Worksheet, header_row: int) -> dict[str, int]:
    found: dict[str, int] = {}

    for cell in ws[header_row]:
        header = normalize_header(cell.value)
        if header in REQUIRED_COLUMNS:
            if header in found:
                raise ValueError(
                    f"Colonna duplicata trovata nella riga {header_row}: '{header}'. "
                    "Il file deve contenerla una sola volta."
                )
            found[header] = cell.column

    return found


def find_required_columns(ws: Worksheet) -> HeaderDetectionResult:
    """
    Cerca le intestazioni obbligatorie nella riga 1.
    Se la riga 1 non contiene tutte le intestazioni richieste, controlla la riga 2.
    Se non le trova neanche nella riga 2, restituisce errore.
    """
    attempts: list[tuple[int, dict[str, int], list[str]]] = []

    for header_row in (1, 2):
        found = scan_header_row(ws, header_row)
        missing = [col for col in REQUIRED_COLUMNS if col not in found]
        attempts.append((header_row, found, missing))

        if not missing:
            return HeaderDetectionResult(columns=found, header_row=header_row)

    details = []
    for header_row, _found, missing in attempts:
        details.append(f"riga {header_row}: mancanti {', '.join(missing)}")

    raise ValueError(
        "Colonne obbligatorie mancanti. "
        "Il codice cerca le intestazioni solo nella riga 1 o nella riga 2. "
        + "; ".join(details)
        + "."
    )


def parse_decimal_value(value: Any, *, row: int, column_name: str) -> Decimal:
    if value is None or (isinstance(value, str) and value.strip() == ""):
        raise ValueError(f"Valore vuoto in riga {row}, colonna '{column_name}'.")

    if isinstance(value, bool):
        raise ValueError(f"Valore booleano non valido in riga {row}, colonna '{column_name}': {value!r}")

    if isinstance(value, int):
        return Decimal(value)

    if isinstance(value, float):
        return Decimal(str(value))

    text = str(value).strip().replace("\xa0", "").replace(" ", "")

    if text.startswith("="):
        raise ValueError(
            f"Formula non convertibile in riga {row}, colonna '{column_name}': {value!r}. "
            "La colonna deve contenere valori numerici, non formule."
        )

    text = text.replace("€", "").replace("'", "")

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]

    if text.startswith("+"):
        text = text[1:]
    elif text.startswith("-"):
        negative = True
        text = text[1:]

    if not re.fullmatch(r"\d+(?:[\.,]\d+)*", text):
        raise ValueError(
            f"Impossibile convertire il valore numerico in riga {row}, colonna '{column_name}': {value!r}"
        )

    comma_positions = [m.start() for m in re.finditer(",", text)]
    dot_positions = [m.start() for m in re.finditer(r"\.", text)]

    if comma_positions and dot_positions:
        last_comma = comma_positions[-1]
        last_dot = dot_positions[-1]
        decimal_pos = max(last_comma, last_dot)
        integer_part = text[:decimal_pos].replace(",", "").replace(".", "")
        decimal_part = text[decimal_pos + 1 :]
        normalized = f"{integer_part}.{decimal_part}"

    elif comma_positions or dot_positions:
        separator = "," if comma_positions else "."
        parts = text.split(separator)

        if len(parts) == 2:
            left, right = parts
            if len(right) <= 2:
                normalized = f"{left}.{right}"
            elif len(right) == 3 and len(left) <= 3:
                normalized = f"{left}{right}"
            else:
                normalized = f"{left}.{right}"
        else:
            last_part = parts[-1]
            if len(last_part) <= 2:
                normalized = "".join(parts[:-1]) + "." + last_part
            elif all(len(part) == 3 for part in parts[1:]):
                normalized = "".join(parts)
            else:
                normalized = "".join(parts[:-1]) + "." + last_part

    else:
        normalized = text

    if negative:
        normalized = "-" + normalized

    try:
        return Decimal(normalized)
    except InvalidOperation as exc:
        raise ValueError(
            f"Impossibile convertire il valore numerico in riga {row}, colonna '{column_name}': {value!r}"
        ) from exc


def parse_doc_number_value(value: Any, *, row: int, column_name: str = "Nr. Doc.") -> int:
    """
    Converte il numero documento in intero.

    La colonna Nr. Doc. può arrivare in due modi:
      1) già come intero: 1001, 161, 25, 1;
      2) in formato inglese delle migliaia letto male da Excel/locale IT:
         5,645 può diventare 5.645; 5,640 può diventare 5.64.

    Per questa colonna i separatori . e , NON sono decimali: sono trattati
    come separatori di migliaia / gruppo numerico e quindi rimossi.
    Esempi:
      "5,645" -> 5645
      5.645   -> 5645
      "5,64" -> 564   poi la funzione di correzione zeri ricostruisce 5640
      1       -> 1     poi, se coerente con i vicini, viene ricostruito 1000
    """
    if value is None or (isinstance(value, str) and value.strip() == ""):
        raise ValueError(f"Valore vuoto in riga {row}, colonna '{column_name}'.")

    if isinstance(value, bool):
        raise ValueError(f"Valore booleano non valido in riga {row}, colonna '{column_name}': {value!r}")

    if isinstance(value, int):
        return value

    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        text = format(value, ".15g")
    else:
        text = str(value).strip().replace("\xa0", "").replace(" ", "")
        if text.startswith("="):
            raise ValueError(
                f"Formula non convertibile in riga {row}, colonna '{column_name}': {value!r}. "
                "La colonna deve contenere valori numerici, non formule."
            )
        text = text.replace("'", "")

    negative = False
    if text.startswith("+"):
        text = text[1:]
    elif text.startswith("-"):
        negative = True
        text = text[1:]

    if "e" in text.lower():
        # Gestione prudenziale di eventuali notazioni scientifiche.
        try:
            number = Decimal(text)
        except InvalidOperation as exc:
            raise ValueError(
                f"Impossibile convertire il Nr. Doc. in riga {row}: {value!r}"
            ) from exc

        if number == number.to_integral_value():
            result = int(number)
            return -result if negative else result

        text = format(float(number), ".15f").rstrip("0").rstrip(".")

    if not re.fullmatch(r"\d+(?:[\.,]\d+)*", text):
        raise ValueError(
            f"Impossibile convertire il Nr. Doc. in riga {row}: {value!r}"
        )

    digits_only = re.sub(r"[\.,]", "", text)
    result = int(digits_only)
    return -result if negative else result


def round_to_2_decimals(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def build_runs(values: list[int]) -> list[tuple[int, int, int]]:
    runs: list[tuple[int, int, int]] = []
    start = 0

    while start < len(values):
        end = start + 1
        while end < len(values) and values[end] == values[start]:
            end += 1
        runs.append((start, end, values[start]))
        start = end

    return runs


def trailing_zero_candidates(value: int, max_zeroes: int = 6) -> list[int]:
    if value < 0:
        return []
    text = str(value)
    return [int(text + ("0" * zeroes)) for zeroes in range(1, max_zeroes + 1)]


def correct_truncated_doc_numbers(
    doc_numbers: list[int],
    *,
    first_data_row: int,
    max_neighbor_gap: int = 200,
) -> tuple[list[int], list[DocCorrection]]:
    """
    Corregge i numeri documento apparentemente troncati per perdita degli zeri finali.

    Esempi gestiti:
      1001, 1, 999     -> 1001, 1000, 999
      1611, 161, 1609  -> 1611, 1610, 1609
      2501, 25, 2499   -> 2501, 2500, 2499

    La correzione viene fatta solo quando il valore ricostruito con zeri finali
    è coerente con il valore distinto precedente e/o successivo. Questo evita di
    trasformare numeri bassi validi, come 94 o 252, in 940 o 2520 senza evidenza.
    """
    corrected = doc_numbers[:]
    corrections: list[DocCorrection] = []
    runs = build_runs(doc_numbers)

    for idx, (start, end, raw_value) in enumerate(runs):
        previous_value = runs[idx - 1][2] if idx > 0 else None
        next_value = runs[idx + 1][2] if idx + 1 < len(runs) else None

        candidates: list[tuple[int, Decimal, int, str]] = []

        if previous_value is not None and next_value is not None:
            low = min(previous_value, next_value)
            high = max(previous_value, next_value)
            raw_is_between_neighbors = low < raw_value < high

            if not raw_is_between_neighbors and abs(previous_value - next_value) <= max_neighbor_gap:
                expected_midpoint = Decimal(previous_value + next_value) / Decimal(2)
                for candidate in trailing_zero_candidates(raw_value):
                    if low < candidate < high:
                        distance = abs(Decimal(candidate) - expected_midpoint)
                        candidates.append((0, distance, candidate, "between_previous_and_next"))

        if previous_value is not None:
            for candidate in trailing_zero_candidates(raw_value):
                if abs(candidate - previous_value) == 1 and abs(raw_value - previous_value) > 50:
                    candidates.append((1, Decimal(0), candidate, "adjacent_to_previous"))

        if next_value is not None:
            for candidate in trailing_zero_candidates(raw_value):
                if abs(candidate - next_value) == 1 and abs(raw_value - next_value) > 50:
                    candidates.append((1, Decimal(0), candidate, "adjacent_to_next"))

        if not candidates:
            continue

        candidates.sort(key=lambda item: (item[0], item[1], item[2]))
        _, _, corrected_value, rule = candidates[0]

        for pos in range(start, end):
            corrected[pos] = corrected_value

        corrections.append(
            DocCorrection(
                first_excel_row=first_data_row + start,
                last_excel_row=first_data_row + end - 1,
                original_value=raw_value,
                corrected_value=corrected_value,
                previous_distinct_value=previous_value,
                next_distinct_value=next_value,
                rule=rule,
                affected_rows=end - start,
            )
        )

    return corrected, corrections


def sort_data_rows_by_doc(ws: Worksheet, *, nr_doc_col: int, first_data_row: int = 2) -> None:
    max_row = ws.max_row
    max_col = ws.max_column

    rows: list[tuple[int, int, list[Any]]] = []
    for row in range(first_data_row, max_row + 1):
        doc_value = ws.cell(row=row, column=nr_doc_col).value
        if not isinstance(doc_value, int):
            doc_value = int(doc_value)
        row_values = [ws.cell(row=row, column=col).value for col in range(1, max_col + 1)]
        rows.append((doc_value, row, row_values))

    rows.sort(key=lambda item: (item[0], item[1]))

    for output_row, (_, _, row_values) in enumerate(rows, start=first_data_row):
        for col, value in enumerate(row_values, start=1):
            ws.cell(row=output_row, column=col).value = value


def clean_excel_file(input_path: str | Path, output_path: str | Path, sheet_name: str | None = None) -> list[DocCorrection]:
    input_path = Path(input_path)
    output_path = Path(output_path)

    if not input_path.exists():
        raise FileNotFoundError(f"File non trovato: {input_path}")

    workbook = load_workbook(input_path)
    ws = workbook[sheet_name] if sheet_name else workbook.worksheets[0]

    header_detection = find_required_columns(ws)
    columns = header_detection.columns
    header_row = header_detection.header_row

    nr_doc_col = columns["Nr. Doc."]
    mc_col = columns["MC"]
    kg_col = columns["Kg"]

    first_data_row = header_row + 1
    if ws.max_row < first_data_row:
        raise ValueError("Il file non contiene righe dati sotto l'intestazione.")

    doc_numbers: list[int] = []

    for row in range(first_data_row, ws.max_row + 1):
        doc_number = parse_doc_number_value(ws.cell(row=row, column=nr_doc_col).value, row=row, column_name="Nr. Doc.")
        doc_numbers.append(doc_number)

        mc_value = parse_decimal_value(ws.cell(row=row, column=mc_col).value, row=row, column_name="MC")
        kg_value = parse_decimal_value(ws.cell(row=row, column=kg_col).value, row=row, column_name="Kg")

        ws.cell(row=row, column=mc_col).value = float(round_to_2_decimals(mc_value))
        ws.cell(row=row, column=kg_col).value = float(round_to_2_decimals(kg_value))

    corrected_doc_numbers, corrections = correct_truncated_doc_numbers(
        doc_numbers,
        first_data_row=first_data_row,
    )

    for offset, corrected_doc in enumerate(corrected_doc_numbers, start=first_data_row):
        ws.cell(row=offset, column=nr_doc_col).value = corrected_doc

    sort_data_rows_by_doc(ws, nr_doc_col=nr_doc_col, first_data_row=first_data_row)

    for row in range(first_data_row, ws.max_row + 1):
        ws.cell(row=row, column=nr_doc_col).number_format = NR_DOC_NUMBER_FORMAT
        ws.cell(row=row, column=mc_col).number_format = DECIMAL_NUMBER_FORMAT
        ws.cell(row=row, column=kg_col).number_format = DECIMAL_NUMBER_FORMAT

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return corrections


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_pulito{input_path.suffix}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Pulisce le colonne 'Nr. Doc.', 'MC' e 'Kg' di un file Excel rifiuti."
    )
    parser.add_argument("input", help="Percorso del file Excel di input (.xlsx)")
    parser.add_argument("output", nargs="?", help="Percorso del file Excel pulito di output (.xlsx)")
    parser.add_argument("--sheet", dest="sheet_name", default=None, help="Nome del foglio da pulire. Se omesso, usa il primo foglio.")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else default_output_path(input_path)

    corrections = clean_excel_file(input_path, output_path, sheet_name=args.sheet_name)

    print(f"File pulito creato: {output_path}")
    print(f"Gruppi Nr. Doc. corretti: {len(corrections)}")
    print(f"Righe Nr. Doc. corrette: {sum(item.affected_rows for item in corrections)}")

    if corrections:
        print("Prime correzioni Nr. Doc.:")
        for item in corrections[:10]:
            if item.first_excel_row == item.last_excel_row:
                rows = str(item.first_excel_row)
            else:
                rows = f"{item.first_excel_row}-{item.last_excel_row}"
            print(f"  righe {rows}: {item.original_value} -> {item.corrected_value} ({item.rule})")


if __name__ == "__main__":
    main()
