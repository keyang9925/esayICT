"""Utility to extract interface information from network device text dumps.
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, Iterable, List

import pandas as pd

DEFAULT_VALUE = "未捕获相关数据"
DEFAULT_INPUT_PATH = Path(r"C:\Users\yangk\SC-CD-BSL-CE-21.CDMA-2025.10.24.txt")


def _first_match(pattern: str, text: str, flags: int = 0, default: str = DEFAULT_VALUE) -> str:
    match = re.search(pattern, text, flags)
    if not match:
        return default

    if match.lastindex:
        for index in range(1, match.lastindex + 1):
            group_value = match.group(index)
            if group_value:
                return group_value.strip()

    return match.group(0).strip()


def parse_display_current_configuration(text: str) -> pd.DataFrame:
    section_match = re.search(
        r"display current-configuration[\s\S]+?(?=^<)", text, re.MULTILINE
    )
    if not section_match:
        return pd.DataFrame(
            columns=[
                "接口名",
                "接口IPv4地址",
                "接口VLAN ID",
                "接口VPN",
                "接口描述",
            ]
        )

    section = section_match.group(0)
    interface_chunks = re.findall(r"^interface [\s\S]+?^#", section, re.MULTILINE)

    rows: List[Dict[str, str]] = []
    for chunk in interface_chunks:
        interface_name = _first_match(r"^interface (\S+)", chunk, re.MULTILINE)
        ipv4_address = _first_match(r"^\s*ip address ([\S ]+)", chunk, re.MULTILINE)
        vlan_id = _first_match(
            r"^\s*port default ([\S ]+)|^\s*vlan-type dot1q ([\S ]+)|^\s*port trunk allow-pass ([\S ]+)",
            chunk,
            re.MULTILINE,
        )
        vpn = _first_match(r"^\s*ip binding vpn-instance ([\S ]+)", chunk, re.MULTILINE)
        description = _first_match(r"^\s*description ([\S ]+)", chunk, re.MULTILINE)

        rows.append(
            {
                "接口名": interface_name,
                "接口IPv4地址": ipv4_address,
                "接口VLAN ID": vlan_id,
                "接口VPN": vpn,
                "接口描述": description,
            }
        )

    return pd.DataFrame(rows)


def parse_display_interface(text: str) -> pd.DataFrame:
    section_match = re.search(r"display interface[\s\S]+?(?=^<)", text, re.MULTILINE)
    if not section_match:
        return pd.DataFrame(
            columns=[
                "接口名",
                "接口当前状态",
                "接口链路状态",
                "接口速率",
                "接口模块类型",
                "接口收光",
                "接口发光",
                "模块波长",
                "传输距离",
                "接口当前CRC",
            ]
        )

    section = section_match.group(0)
    interface_chunks = re.findall(
        r"\S+ current state :[\s\S]+?(?=\r?\n\r?\n)", section
    )

    rows: List[Dict[str, str]] = []
    for chunk in interface_chunks:
        interface_name = _first_match(r"^(\S+)", chunk)
        current_state = _first_match(
            r"^\S{1,40} current state : ([\S ]+)",
            chunk,
            re.MULTILINE,
        )
        link_state = _first_match(
            r"Line protocol current state : ([\S ]+)",
            chunk,
        )
        speed = _first_match(
            r"Port BW: (\S+?)(?=,)|Current BW: ?(\S+?)(?=,)",
            chunk,
        )
        module_type = _first_match(
            r"Transceiver Mode: (\S+)|Media type: (\S+)",
            chunk,
        )
        rx_power = _first_match(r"Rx Power: (\S+)", chunk)
        tx_power = _first_match(r"Tx Power: (\S+)", chunk)
        wavelength = _first_match(r"WaveLength: (\S+)(?=,)", chunk)
        distance = _first_match(r"Transmission Distance: (\S+)", chunk)
        crc = _first_match(r"CRC: (\S+)", chunk)

        rows.append(
            {
                "接口名": interface_name,
                "接口当前状态": current_state,
                "接口链路状态": link_state,
                "接口速率": speed,
                "接口模块类型": module_type,
                "接口收光": rx_power,
                "接口发光": tx_power,
                "模块波长": wavelength,
                "传输距离": distance,
                "接口当前CRC": crc,
            }
        )

    return pd.DataFrame(rows)


def merge_datasets(dataset1: pd.DataFrame, dataset2: pd.DataFrame) -> pd.DataFrame:
    if dataset1.empty and dataset2.empty:
        return pd.DataFrame()

    if dataset1.empty:
        return dataset2
    if dataset2.empty:
        return dataset1

    return pd.merge(dataset1, dataset2, on="接口名", how="outer")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def write_excel(df: pd.DataFrame, output_path: Path) -> None:
    if df.empty:
        # Still create an empty excel file with headers if possible
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
    else:
        df.to_excel(output_path, index=False)


def parse_arguments(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从display命令采集的文本中提取接口信息，并输出Excel。"
    )
    parser.add_argument(
        "-i",
        "--input",
        type=Path,
        default=DEFAULT_INPUT_PATH,
        help="display命令的原始txt文件",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="输出Excel文件路径，例如 result.xlsx",
    )

    # 使用 parse_known_args 以兼容 Jupyter/ IPython 额外注入的参数（例如 -f）。
    args, _ = parser.parse_known_args(argv)
    return args


def _prompt_for_path(prompt: str) -> Path:
    """Prompt the user for a path when CLI arguments are missing."""

    try:
        raw_value = input(prompt)
    except EOFError:  # pragma: no cover - defensive: non-interactive environments
        raise SystemExit("未提供必要的文件路径，程序终止。") from None

    raw_value = raw_value.strip()
    if not raw_value:
        raise SystemExit("未提供必要的文件路径，程序终止。")

    return Path(raw_value)


def _resolve_paths(args: argparse.Namespace) -> tuple[Path, Path]:
    """Resolve input/output paths from CLI args or interactive prompts."""

    input_path = args.input
    output_path = args.output

    # 如果缺少参数，在交互式环境下提示用户输入。
    if input_path is None:
        input_path = _prompt_for_path("请输入display命令txt文件路径：")
    if output_path is None:
        output_path = _prompt_for_path("请输入输出Excel文件路径，例如 result.xlsx：")

    return input_path, output_path


def process_file(input_path: Path, output_path: Path) -> None:
    text = read_text(input_path)

    dataset1 = parse_display_current_configuration(text)
    dataset2 = parse_display_interface(text)
    merged = merge_datasets(dataset1, dataset2)

    write_excel(merged, output_path)


def main(argv: Iterable[str] | None = None) -> None:
    args = parse_arguments(argv)
    input_path, output_path = _resolve_paths(args)

    if not input_path.exists():
        raise SystemExit(f"找不到输入文件：{input_path}")

    process_file(input_path, output_path)


if __name__ == "__main__":
    main()
