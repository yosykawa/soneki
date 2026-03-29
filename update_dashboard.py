"""
update_dashboard.py
--------------------
使い方:
    python update_dashboard.py

  Excel ファイルを data/ フォルダに置いて実行するだけで
  dashboard/index.html が自動更新されます。

フォルダ構成:
    project/
    ├── update_dashboard.py   ← このファイル
    ├── template.html         ← HTMLテンプレート（フォーマット本体）
    ├── data/
    │   └── _損益管理工数分析_*.xlsx   ← ここにExcelを置く
    └── dashboard/
        └── index.html        ← 出力先（ブラウザで開くファイル）
"""

import json
import sys
import glob
import os
from pathlib import Path
from datetime import datetime
from typing import Optional, Tuple, List

# ── 設定 ──────────────────────────────────────────
BASE_DIR      = Path(__file__).parent
DATA_DIR      = BASE_DIR / "data"
TEMPLATE_FILE = BASE_DIR / "template.html"
OUTPUT_DIR    = BASE_DIR / "dashboard"
OUTPUT_FILE   = OUTPUT_DIR / "index.html"
PLACEHOLDER   = "/*__ALL_DATA__*/"

# 稼働率データのある列範囲（0-indexed）
DEPT_COL   = 10   # K列
NAME_COL   = 11   # L列
ID_COL     = 12   # M列
VAL_START  = 13   # N列（月次データ開始）
VAL_END    = 25   # Y列+1（月次データ終了）

VALID_DEPTS = {'統括部', '第一IT部', '第二IT部', '第三IT部'}
# ──────────────────────────────────────────────────


def find_excel() -> Path:
    """data/ フォルダ内の xlsx を探す（複数ある場合は最新更新日時を優先）"""
    files = list(DATA_DIR.glob("*.xlsx"))
    if not files:
        print(f"[ERROR] data/ フォルダに .xlsx が見つかりません: {DATA_DIR}")
        sys.exit(1)
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    if len(files) > 1:
        print(f"[INFO] 複数のExcelが見つかりました。最新を使用: {files[0].name}")
    return files[0]


def to_month_str(v) -> str:
    """セル値を YYYY/MM 形式に変換"""
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y/%m")
    s = str(v)
    if len(s) >= 7:
        return s[:7].replace("-", "/")
    return s


def safe_float(v) -> Optional[float]:
    """数値変換。失敗・日付・文字列は None"""
    if v is None:
        return None
    try:
        import pandas as pd
        if pd.isna(v):
            return None
    except Exception:
        pass
    if hasattr(v, "year"):   # datetime
        return None
    if isinstance(v, str):
        return None
    try:
        return round(float(v), 6)
    except (TypeError, ValueError):
        return None


def parse_members(df, header_row: int) -> Tuple[list, List[str]]:
    """稼働率メンバーデータをパース"""
    import pandas as pd

    months = [to_month_str(df.iloc[header_row, c]) for c in range(VAL_START, VAL_END)]

    members = []
    for r in range(header_row + 1, len(df)):
        row = df.iloc[r]
        dept   = row[DEPT_COL]
        name   = row[NAME_COL]
        emp_id = row[ID_COL]

        if pd.isna(dept) or pd.isna(name):
            continue
        if str(dept) not in VALID_DEPTS:
            continue

        vals = []
        for c in range(VAL_START, VAL_END):
            v = safe_float(row[c])
            vals.append(v if v is not None else 0.0)

        try:
            eid = int(float(str(emp_id).split(".")[0]))
        except (ValueError, TypeError):
            eid = 0

        members.append({
            "dept": str(dept),
            "name": str(name),
            "id":   eid,
            "v":    vals,
        })

    return members, months


def parse_projects(df) -> Tuple[list, List[str]]:
    """案件一覧をパース（B25～Z最終行相当）"""
    import pandas as pd

    # ヘッダー行を動的に検索
    proj_hr = None
    for r in range(20, min(len(df), 40)):
        if str(df.iloc[r, 1]) == "案件状況":
            proj_hr = r
            break

    if proj_hr is None:
        return [], []

    proj_months = [to_month_str(df.iloc[proj_hr, c]) for c in range(VAL_START, VAL_END)]

    projects = []
    cur = {k: None for k in ("status", "dept", "num", "biz", "client", "pj")}
    fill_cols = {1: "status", 2: "dept", 3: "num", 4: "biz", 6: "client", 7: "pj"}

    for r in range(proj_hr + 1, len(df)):
        row = df.iloc[r]

        for col, key in fill_cols.items():
            v = row[col]
            if pd.notna(v) and str(v) not in ("nan", ""):
                cur[key] = str(v)

        if cur["status"] is None:
            continue

        role  = str(row[8])  if pd.notna(row[8])  else ""
        rank  = str(row[9])  if pd.notna(row[9])  else ""
        affil = str(row[10]) if pd.notna(row[10]) else ""
        name  = str(row[11]) if pd.notna(row[11]) else ""
        eid   = str(row[12]).split(".")[0] if pd.notna(row[12]) else ""

        vals = [safe_float(row[c]) for c in range(VAL_START, VAL_END)]

        projects.append({
            "status": cur["status"] or "",
            "dept":   cur["dept"]   or "",
            "num":    cur["num"]    or "",
            "biz":    cur["biz"]    or "",
            "client": cur["client"] or "",
            "pj":     cur["pj"]     or "",
            "role":   role,
            "rank":   rank,
            "affil":  affil,
            "name":   name,
            "id":     eid,
            "v":      vals,
        })

    return projects, proj_months


def build_all_data(excel_path: Path) -> dict:
    """Excel全シートを読み込み ALL_DATA 辞書を構築"""
    import pandas as pd

    xl = pd.ExcelFile(excel_path)
    print(f"[INFO] シート: {xl.sheet_names}")

    result = {}
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, header=None)

        # 稼働率ヘッダー行
        header_row = next(
            (r for r in range(5) if df.iloc[r, DEPT_COL] == "部署"),
            None
        )
        if header_row is None:
            print(f"[WARN] {sheet}: 稼働率ヘッダーが見つかりません。スキップ。")
            continue

        members, months       = parse_members(df, header_row)
        projects, proj_months = parse_projects(df)

        result[sheet] = {
            "months":      months,
            "members":     members,
            "proj_months": proj_months if proj_months else months,
            "projects":    projects,
        }
        print(f"  {sheet}: {len(members)}名, {len(projects)}件")

    return result


def inject_data(all_data: dict) -> str:
    """テンプレートにデータを注入して完成HTMLを返す"""
    if not TEMPLATE_FILE.exists():
        print(f"[ERROR] template.html が見つかりません: {TEMPLATE_FILE}")
        sys.exit(1)

    with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
        template = f.read()

    if PLACEHOLDER not in template:
        print(f"[ERROR] template.html に {PLACEHOLDER} が見つかりません")
        sys.exit(1)

    data_js = "const ALL_DATA = " + json.dumps(all_data, ensure_ascii=False) + ";"
    return template.replace(PLACEHOLDER, data_js)


def main():
    print("=" * 50)
    print("  稼働率ダッシュボード 更新スクリプト")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    # 前提確認
    if not DATA_DIR.exists():
        DATA_DIR.mkdir(parents=True)
        print(f"[INFO] data/ フォルダを作成しました: {DATA_DIR}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Excel 検索
    excel_path = find_excel()
    print(f"[INFO] Excel: {excel_path.name}")

    # データ構築
    all_data = build_all_data(excel_path)
    if not all_data:
        print("[ERROR] 有効なシートが見つかりませんでした")
        sys.exit(1)

    # HTML 生成
    html = inject_data(all_data)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n[OK] 出力完了: {OUTPUT_FILE}")
    print(f"     ブラウザで開く: file://{OUTPUT_FILE.resolve()}")
    print("=" * 50)


if __name__ == "__main__":
    main()
