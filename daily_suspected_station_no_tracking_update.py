# pip install requests pandas openpyxl

import requests
import pandas as pd
import datetime

# ── 配置 ────────────────────────────────────────────────────
APP_ID = "cli_aaaa63af823a9ce3"
APP_SECRET = "yljUK3LoJCy4xznQRq9LMbuyd6fBnfv2"

SOURCE_SPREADSHEET_TOKEN = "XRdMw9RsuiEBz6kAB5Xc3utLnbf"
SOURCE_SHEET_ID = "ysztZJ"
TARGET_SPREADSHEET_TOKEN = "MMNmwGNeXitJk7kDxLochDXpnkj"

DATE_COLS = {"统计日期", "异常触发时间", "轨迹时间"}


# ── 认证 ────────────────────────────────────────────────────
def get_auth_headers() -> dict:
    """获取带 Bearer token 的请求头"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    resp = requests.post(url, json={"app_id": APP_ID, "app_secret": APP_SECRET})
    resp.raise_for_status()
    token = resp.json()["tenant_access_token"]
    return {"Authorization": f"Bearer {token}"}


# ── 数据读取 ─────────────────────────────────────────────────
def fetch_source_data(headers: dict) -> pd.DataFrame:
    """从飞书源表读取数据，返回 DataFrame"""
    url = (
        f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets"
        f"/{SOURCE_SPREADSHEET_TOKEN}/values/{SOURCE_SHEET_ID}"
    )
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    rows = resp.json()["data"]["valueRange"]["values"]
    return pd.DataFrame(rows[1:], columns=rows[0])


# ── 数据处理 ─────────────────────────────────────────────────
def build_filtered_df(df: pd.DataFrame) -> pd.DataFrame:
    """筛选滞更天数==2，并新增站点列"""
    df = df.copy()
    df["滞更天数"] = pd.to_numeric(df["滞更天数"], errors="coerce")
    df = df[df["滞更天数"] == 2].reset_index(drop=True)
    df["站点"] = df["派送方"].astype(str).str[:3]
    return df


def build_pivot_df(df_filtered: pd.DataFrame) -> pd.DataFrame:
    """生成带小计行和总计行的透视表"""
    pivot = pd.pivot_table(
        df_filtered,
        index=["站点", "派送方", "快递员区域名称"],
        columns="轨迹节点",
        values="运单号",
        aggfunc="count",
        fill_value=0,
    ).reset_index()

    node_cols = [c for c in pivot.columns if c not in ("站点", "派送方", "快递员区域名称")]
    pivot["Total"] = pivot[node_cols].sum(axis=1)

    result_rows = []
    for site, group in pivot.groupby("站点", sort=False):
        result_rows.extend(group.to_dict("records"))
        subtotal = {"站点": f"{site} Total", "派送方": "", "快递员区域名称": ""}
        subtotal.update({col: group[col].sum() for col in node_cols + ["Total"]})
        result_rows.append(subtotal)

    grand_total = {"站点": "Grand Total", "派送方": "", "快递员区域名称": ""}
    grand_total.update({col: pivot[col].sum() for col in node_cols + ["Total"]})
    result_rows.append(grand_total)

    final = pd.DataFrame(result_rows)
    for col in node_cols + ["Total"]:
        final[col] = final[col].fillna(0).astype(int)

    return final[["站点", "派送方", "快递员区域名称"] + node_cols + ["Total"]]


# ── 写入工具 ─────────────────────────────────────────────────
def convert_excel_date(value, col_name: str):
    """将飞书/Excel 序列号转换为日期字符串（仅对日期列生效）"""
    if col_name not in DATE_COLS:
        return value
    try:
        num = float(value)
        dt = pd.Timestamp("1899-12-30") + pd.Timedelta(days=num)
        return dt.strftime("%Y-%m-%d") if num % 1 == 0 else dt.strftime("%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError):
        return value


def safe_val(v):
    """将值转为飞书 API 可接受的格式"""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    if isinstance(v, (int, float)):
        return v
    return str(v)


def create_and_write_sheet(title: str, dataframe: pd.DataFrame, headers: dict) -> None:
    """在目标表中创建（或复用）sheet，并写入数据"""
    base_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{TARGET_SPREADSHEET_TOKEN}"

    # 查找或创建 sheet
    meta = requests.get(f"{base_url}/metainfo", headers=headers).json()
    existing = next((s for s in meta["data"]["sheets"] if s["title"] == title), None)

    if existing:
        sheet_id = existing["sheetId"]
        print(f"[复用] {title}  sheetId={sheet_id}")
    else:
        resp = requests.post(
            f"{base_url}/sheets_batch_update",
            headers=headers,
            json={"requests": [{"addSheet": {"properties": {"title": title}}}]},
        )
        sheet_id = resp.json()["data"]["replies"][0]["addSheet"]["properties"]["sheetId"]
        print(f"[新建] {title}  sheetId={sheet_id}")

    # 构造写入数据
    col_names = dataframe.columns.tolist()
    data_rows = [
        [safe_val(convert_excel_date(v, col_names[i])) for i, v in enumerate(row)]
        for row in dataframe.values.tolist()
    ]
    values = [col_names] + data_rows

    # 写入
    resp = requests.put(
        f"{base_url}/values",
        headers=headers,
        json={"valueRange": {"range": sheet_id, "values": values}},
    )
    print(f"[写入] {title}  状态: {resp.json().get('msg')}")


# ── 主流程 ───────────────────────────────────────────────────
def main():
    headers = get_auth_headers()

    # 读取 & 处理
    df_raw = fetch_source_data(headers)
    df_filtered = build_filtered_df(df_raw)
    df_pivot = build_pivot_df(df_filtered)

    # 写入
    date_str = datetime.datetime.now().strftime("%Y%m%d")
    create_and_write_sheet(date_str, df_filtered, headers)
    create_and_write_sheet(f"{date_str}_pivot table", df_pivot, headers)


if __name__ == "__main__":
    main()

