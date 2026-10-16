"""生成一套示例数据，用来练习 / 自测计划工作台。

用法：
    pip install -r requirements.txt
    python make_sample_data.py            # 默认输出到桌面 fcst-示例数据
    python make_sample_data.py 输出目录    # 也可以自己指定目录

生成的文件名和网页上上传框的顺序一一对应（1-PSI文件、2-当月提货跟进表……），
每个工具一个子目录，另外附带一份「生成结果示例.xlsx」方便对照。
文件是随机造的假数据，不含任何真实业务信息，可以直接删除。
"""

from __future__ import annotations

import importlib.util
import sys
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
BRANDS = ["MOVA", "追觅", "光子跃迁", "Arçelik", "Beko"]
BUS = ["吹风机", "洗地机", "造型器", "脱毛仪", "扫地机"]
CUSTOMER_TYPES = ["TO-B", "TO-C", "TO-MKT"]
ROWS = 8
RUN_DATE = date.today()

TOOLS = [
    ("01-常规总表", "app.py", False, False),
    ("02-月末总表", "month_end_app.py", True, False),
    ("03-常规-客户类型", "Separation of TOB and TOC dmd app.py", False, True),
    ("04-月末-客户类型", "Separation of TOB and TOC dmd app month_end.py", True, True),
]


def build_frames(month_end: bool, split: bool) -> dict[str, pd.DataFrame]:
    psi_rows, delivery_rows = [], []
    for i in range(1, ROWS + 1):
        brand = BRANDS[i % len(BRANDS)]
        sku = f"SKU{i:04d}"
        psi_rows.append(
            {"SKU": sku, "品牌": brand, **{f"m_{m}": 100 * m + i * 3 for m in range(1, 6)}}
        )
        row = {
            "SKU": sku, "品牌": brand,
            "m_0": 500 + i, "m_0安全库存": 50 + i,
            "m_1": 600 + i, "m_1安全库存": 60 + i,
        }
        if split:
            row["客户类型"] = CUSTOMER_TYPES[i % len(CUSTOMER_TYPES)]
        delivery_rows.append(row)

    psi = pd.DataFrame(psi_rows)
    delivery = pd.DataFrame(delivery_rows)
    master = pd.DataFrame([
        {"SKU": f"SKU{i:04d}", "品牌": BRANDS[i % len(BRANDS)], "BU": BUS[i % len(BUS)],
         "项目号": f"P{i:03d}", "物料编码": f"MAT{i:05d}", "系列": f"S{i}", "版本": "V1"}
        for i in range(1, ROWS + 1)
    ])
    keys = [f"{BRANDS[i % len(BRANDS)]}SKU{i:04d}" for i in range(1, ROWS + 1)]
    template = pd.DataFrame({"Key": keys})
    # 模板里带一列「上次版本的需求总和」，用来算差异列。
    # 两个「客户类型」工具要求上次版本必须是「同月、更早的一天」才会自动识别，所以取上一天。
    prev_day = RUN_DATE.day - 1 if RUN_DATE.day > 1 else 1
    last_mmdd = f"{RUN_DATE.month:02d}{prev_day:02d}"
    template[f"{RUN_DATE.month}月需求总和(包含安全库存)({last_mmdd})"] = [
        500 + i * 10 for i in range(1, ROWS + 1)
    ]

    files = {
        "1-PSI文件.xlsx": psi,
        "4-FCST总表模板.xlsx": template,
        "5-主数据.xlsx": master,
    }
    if month_end:
        files["2-当月提货跟进表.xlsx"] = delivery.drop(columns=["m_1", "m_1安全库存"])
        files["3-次月提货跟进表.xlsx"] = delivery.drop(columns=["m_0", "m_0安全库存"])
    else:
        files["2-当月提货跟进表.xlsx"] = delivery.drop(columns=["m_1", "m_1安全库存"])
    return files


def load_processor(script_name: str):
    spec = importlib.util.spec_from_file_location(
        "sample_" + script_name.replace(" ", "_").replace(".py", ""), ROOT / script_name
    )
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def main() -> None:
    out_root = Path(sys.argv[1]) if len(sys.argv) > 1 else Path.home() / "Desktop" / "fcst-示例数据"
    out_root.mkdir(parents=True, exist_ok=True)
    print(f"输出目录：{out_root}\n")

    for folder, script, month_end, split in TOOLS:
        target = out_root / folder
        target.mkdir(exist_ok=True)
        frames = build_frames(month_end, split)

        paths: dict[str, Path] = {}
        for name, frame in frames.items():
            path = target / name
            frame.to_excel(path, index=False)
            paths[name] = path

        # 直接调用各脚本里的计算函数，验证一遍并产出结果示例
        streams = {_core_key(name, month_end, split): BytesIO(p.read_bytes()) for name, p in paths.items()}
        module = load_processor(script)
        result, output, warnings = module.build_forecast_workbook(streams, RUN_DATE, None)
        (target / "生成结果示例.xlsx").write_bytes(output.getvalue())

        status = f"  ✓ {folder}：{len(paths)} 个输入文件，生成 {len(result)} 行 × {len(result.columns)} 列"
        print(status)
        for warning in warnings:
            print(f"      ⚠ {warning}")

    print("\n完成。打开网页后，按文件名里的数字顺序依次上传即可。")


def _core_key(file_name: str, month_end: bool, split: bool) -> str:
    """把展示用的中文文件名映射回各脚本期望的内部键名。

    工具 02 的键是 delivery_m0 / delivery_m1，工具 04 的键是 delivery / delivery_1，
    这是两个脚本本来就不一致的地方，这里跟着它们走。
    """
    if file_name.startswith("1-"):
        return "psi"
    if file_name.startswith("2-"):
        # 只有「月末但不拆客户」才是 delivery_m0，其余都是 delivery
        return "delivery_m0" if (month_end and not split) else "delivery"
    if file_name.startswith("3-"):
        return "delivery_1" if split else "delivery_m1"
    if file_name.startswith("4-"):
        return "template"
    return "master"


if __name__ == "__main__":
    main()
