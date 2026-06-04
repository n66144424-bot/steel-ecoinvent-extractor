from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from lxml import etree
from openpyxl import Workbook
from tqdm import tqdm

# 請改成你的實際資料夾路徑（可放多個批次）
DATASETS_FOLDERS = [
    r"D:\\ecoinvent\\batch_1\\datasets",
    r"D:\\ecoinvent\\batch_2\\datasets",
]

OUTPUT_XLSX = "steel_technosphere_inputs.xlsx"
DEDUPLICATE_PROCESSES = True

# 鋼鐵關鍵字（可自行擴充）
STEEL_KEYWORDS = [
    "steel",
    "stainless",
    "hot rolled",
    "cold rolled",
    "pig iron",
    "blast furnace",
    "electric arc furnace",
    "鐵",
    "鋼",
]


def _text(element: Optional[etree._Element]) -> str:
    if element is None or element.text is None:
        return ""
    return element.text.strip()


def _first_by_localname(root: etree._Element, name: str) -> Optional[etree._Element]:
    return root.find(f".//*[local-name()='{name}']")


def _all_by_localname(root: etree._Element, name: str) -> List[etree._Element]:
    return root.findall(f".//*[local-name()='{name}']")


def _child_text_by_localname(element: etree._Element, name: str) -> str:
    child = element.find(f"./*[local-name()='{name}']")
    return _text(child)


def _is_steel_process(process_name: str, reference_product: str) -> bool:
    haystack = f"{process_name} {reference_product}".lower()
    return any(keyword in haystack for keyword in STEEL_KEYWORDS)


def _is_technosphere_exchange(exchange: etree._Element) -> bool:
    input_group = (exchange.get("inputGroup") or "").strip()
    if input_group == "5":
        return True

    # 部分資料可能沒有 inputGroup，使用 activityLinkId 當備援判定
    if exchange.get("activityLinkId"):
        return True

    group_text = _child_text_by_localname(exchange, "inputGroup")
    return group_text == "5"


def _parse_process(spold_path: Path, batch_name: str) -> Tuple[Dict[str, str], List[Dict[str, str]], List[str]]:
    logs: List[str] = []
    process_meta: Dict[str, str] = {}
    input_rows: List[Dict[str, str]] = []

    try:
        tree = etree.parse(str(spold_path))
        root = tree.getroot()
    except Exception as exc:
        logs.append(f"[ERROR] 無法解析 XML：{spold_path} ({exc})")
        return process_meta, input_rows, logs

    activity = _first_by_localname(root, "activity")
    process_name = _child_text_by_localname(activity, "activityName") if activity is not None else ""
    geography = _child_text_by_localname(activity, "geography") if activity is not None else ""

    reference_function = _first_by_localname(root, "referenceFunction")
    reference_product = _child_text_by_localname(reference_function, "name") if reference_function is not None else ""

    process_uuid = ""
    if activity is not None:
        process_uuid = (activity.get("id") or activity.get("activityId") or "").strip()

    if not _is_steel_process(process_name, reference_product):
        return process_meta, input_rows, logs

    process_meta = {
        "batch": batch_name,
        "source_file": spold_path.name,
        "process_name": process_name,
        "reference_product": reference_product,
        "geography": geography,
        "process_uuid": process_uuid,
    }

    exchanges = _all_by_localname(root, "exchange")
    for exchange in exchanges:
        if not _is_technosphere_exchange(exchange):
            continue

        input_name = _child_text_by_localname(exchange, "name")
        if not input_name:
            input_name = exchange.get("name", "").strip()

        amount = exchange.get("amount", "")
        unit = exchange.get("unitName", "")

        if not unit:
            unit = _child_text_by_localname(exchange, "unitName")
        if not unit:
            unit = _child_text_by_localname(exchange, "unit")

        input_rows.append(
            {
                **process_meta,
                "exchange_id": (exchange.get("id") or "").strip(),
                "input_name": input_name,
                "amount": amount,
                "unit": unit,
                "linked_activity_id": (exchange.get("activityLinkId") or "").strip(),
            }
        )

    return process_meta, input_rows, logs


def _iter_spold_files(dataset_folders: Iterable[str]) -> Tuple[List[Tuple[Path, str]], List[str]]:
    files: List[Tuple[Path, str]] = []
    logs: List[str] = []

    for folder in dataset_folders:
        folder_path = Path(folder)
        if not folder_path.exists():
            logs.append(f"[WARN] 找不到資料夾：{folder_path}")
            continue

        batch_name = folder_path.parent.name or folder_path.name
        spold_files = sorted(folder_path.glob("*.spold"))
        if not spold_files:
            logs.append(f"[WARN] 資料夾內沒有 .spold 檔：{folder_path}")
            continue

        files.extend((spold_file, batch_name) for spold_file in spold_files)
        logs.append(f"[INFO] 批次 {batch_name}：找到 {len(spold_files)} 個 .spold")

    return files, logs


def _build_workbook(
    execution_logs: List[Tuple[str, str]],
    process_rows: List[Dict[str, str]],
    input_rows: List[Dict[str, str]],
    material_rows: List[Tuple[str, str, float]],
) -> Workbook:
    workbook = Workbook()

    ws_log = workbook.active
    ws_log.title = "execution_log"
    ws_log.append(["timestamp", "level", "message"])
    for timestamp, message in execution_logs:
        level = message[1:5] if message.startswith("[") else "INFO"
        ws_log.append([timestamp, level, message])

    ws_process = workbook.create_sheet("process_overview")
    ws_process.append(["batch", "process_name", "reference_product", "geography", "process_uuid", "source_file"])
    for row in process_rows:
        ws_process.append(
            [
                row["batch"],
                row["process_name"],
                row["reference_product"],
                row["geography"],
                row["process_uuid"],
                row["source_file"],
            ]
        )

    ws_input = workbook.create_sheet("input_details")
    ws_input.append(
        [
            "batch",
            "process_name",
            "reference_product",
            "geography",
            "process_uuid",
            "source_file",
            "exchange_id",
            "input_name",
            "amount",
            "unit",
            "linked_activity_id",
        ]
    )
    for row in input_rows:
        ws_input.append(
            [
                row["batch"],
                row["process_name"],
                row["reference_product"],
                row["geography"],
                row["process_uuid"],
                row["source_file"],
                row["exchange_id"],
                row["input_name"],
                row["amount"],
                row["unit"],
                row["linked_activity_id"],
            ]
        )

    ws_agg = workbook.create_sheet("aggregated_materials")
    ws_agg.append(["batch", "input_name", "total_amount"])
    for batch, input_name, total_amount in material_rows:
        ws_agg.append([batch, input_name, total_amount])

    return workbook


def main() -> None:
    now = datetime.now
    execution_logs: List[Tuple[str, str]] = []

    def log(message: str) -> None:
        timestamp = now().strftime("%Y-%m-%d %H:%M:%S")
        execution_logs.append((timestamp, message))
        print(message)

    all_files, scan_logs = _iter_spold_files(DATASETS_FOLDERS)
    for msg in scan_logs:
        log(msg)

    process_rows: List[Dict[str, str]] = []
    input_rows: List[Dict[str, str]] = []
    dedupe_seen: set[Tuple[str, str, str, str]] = set()
    duplicate_count = 0

    for spold_file, batch_name in tqdm(all_files, desc="掃描 SPOLD", unit="file"):
        process_meta, inputs, parse_logs = _parse_process(spold_file, batch_name)
        for msg in parse_logs:
            log(msg)

        if not process_meta:
            continue

        dedupe_key = (
            process_meta["process_name"],
            process_meta["reference_product"],
            process_meta["geography"],
            process_meta["process_uuid"],
        )
        if DEDUPLICATE_PROCESSES and dedupe_key in dedupe_seen:
            duplicate_count += 1
            continue

        dedupe_seen.add(dedupe_key)
        process_rows.append(process_meta)
        input_rows.extend(inputs)

    aggregated: Dict[Tuple[str, str], float] = defaultdict(float)
    for row in input_rows:
        amount_raw = str(row.get("amount", "")).strip()
        try:
            amount_value = float(amount_raw) if amount_raw else 0.0
        except ValueError:
            amount_value = 0.0
        key = (row["batch"], row["input_name"])
        aggregated[key] += amount_value

    material_rows = sorted((batch, name, total) for (batch, name), total in aggregated.items())

    log(f"[INFO] 鋼鐵製程數：{len(process_rows)}")
    log(f"[INFO] 技術投入筆數：{len(input_rows)}")
    log(f"[INFO] 去重複略過製程數：{duplicate_count}")

    workbook = _build_workbook(execution_logs, process_rows, input_rows, material_rows)
    workbook.save(OUTPUT_XLSX)
    log(f"[DONE] 完成！Excel 已輸出：{Path(OUTPUT_XLSX).resolve()}")


if __name__ == "__main__":
    main()
