# =============================================================================
# 日期    : 20230810 CJ TSAI  (改善版 2025)
# 目的    : 讀入EXCEL"短期預示排程.xlsx" 中 分頁
#           "光陽短期預示訂單", "Ecount庫存餘量分析表", "生管BOM" 及 "鐵管BOM"
#           產生ECOUNT "採購單" EXCEL上傳格式
#
# 改善項目（相較原始版本）：
#   [MRP-1] MRP_採購件向下抵扣_v4 取代原 MRP_採購件向下抵扣
#           子件需求基準 = 原始0層建議採購量（非逐層傳遞）
#           中間半成品庫存只抵扣本身那層，不縮減底層物料需求
#   [MRP-2] 0層建議採購量分兩步計算
#           A1：原始母件庫存 → 需製造量（子層基準）
#           A2：焊接件自身庫存 → 焊接件實際採購量
#   [MRP-3] MRP 在加總之前執行（先算各路徑，加總後輸出採購單）
#           共用子件在不同母件路徑下各自獨立計算後才加總
#   [MRP-4] stock_df 統一傳入函數，子件庫存不與0層庫存混用
#   [MRP-5] mask 加入 來源原始母件號碼，避免跨路徑覆蓋
#   [MRP-6] bom_map key = (原始母件, 父件, 層級)，路徑隔離
#   [MRP-7] 累計使用量計算，使用量 round(6) 消除浮點誤差
#
# 20260713 修正 v7 → v7.3：
#   [BUGFIX] 品項自己同時是訂單來源、又被其他訂單當共用子件引用時，
#            其「自身訂單需求」在 v7 會遺失，導致建議採購量被錯誤歸零。
#            詳見 MRP_採購件向下抵扣_v7 函式內註解。
#
# 20260713 修正 v7.3 → v7.4：
#   [BUGFIX] 移除 STEP6「L0品項逐訂單分拆」機制。此機制與 STEP5 的
#            全域拓樸計算重複且矛盾：STEP5 已把「全域正確總量」寫入
#            同一品項編碼的每一列，若下游（主程式）對這些重複列做
#            groupby+sum，會把正確值依品項出現列數重複疊加、灌水。
#            改為 STEP5 一次算好全域正確值，下游僅需「過濾重複列」，
#            不可再加總。同步移除主程式中舊的 groupby+sum 加總邏輯。
#
# 20260713 修正 v7.4 → v7.5：
#   [格式修正] 「說明」欄改為逐階敘述：同一品項若被多張原始母件訂單
#   共用，依 original_order 順序逐筆呈現「這一筆訂單淨額後的需製造
#   量」與「處理到這裡時還剩多少可用庫存」的階梯式扣抵過程，而不是
#   每一列都印相同的全域文字。建議採購量／採購量仍以全域一次淨額
#   為準，不受此格式調整影響。
#
# 20260713 修正 v7.5 → v7.6（效能優化）：
#   [效能] STEP4/STEP5/主程式多處用「布林遮罩重新掃描整個df」的方式
#   逐一查找，在大量品項/大量原始母件下會是O(N×M)甚至O(N×M×K)的效能
#   陷阱，實測會直接卡死。改為預先用groupby建立查找索引，全部O(1)
#   查找，15,000列/300訂單規模下由原本估算數十億次比對降為3.57秒。
#
# 20260714 修正 v7.6 → v8（重大架構調整）：
#   [業務規則修正] 依使用者提供的完整人工試算表逆推確認：正確的扣抵
#   順序不是「拓樸全域一次淨額」，而是：
#     1. 每張原始母件訂單對0層品項的需求，先用該母件自身庫存淨額成
#        「需製造量」（品項自己也是訂單來源時例外，見下）。
#     2. 0層品項自身維護一個可用庫存池，依 original_order 給定順序，
#        逐張訂單扣抵「需製造量」，扣不夠才需採購（階梯式）。
#     3. 任何更深層品項（1層、2層...）一律參照其「0層祖先」在同一張
#        訂單當下的『採購』結果 × 使用量比例，作為自己的原始需求，
#        同樣用自己的庫存池、依同一訂單順序做階梯式扣抵——不是逐層
#        傳遞（不經過中間層的計算結果），而是統一參照0層的結果。
#     4. 品項若被多個不同「0層家族」共用，扣抵池不分家族，依
#        original_order 的全域順序連續扣抵。
#     5. 最終「建議採購量」＝該品項所有階梯步驟『採購』數字的加總
#        （寫入該品項每一列皆為同一總量，供下游過濾重複列使用，
#        不可再對重複列做二次加總）。
#   已用使用者提供的雙家族共用件案例（5012A-BLA7-900／5012A-PRB3-900
#   共用 50120-BAL7-900）完整驗證，10個品項的最終數字與人工試算表
#   逐一比對皆吻合（3079/3078/3079/3079/9699/6620/5820/6620/6610/6620）。
#
# 20260714 修正 v8 → v9（角色判斷改為逐列判斷，不再用全域品項編碼分類）：
#   [BUGFIX] 同一品項編碼在不同家族／不同訂單底下，可能扮演不同角色
#   （在A家族是0層直接子件，在B家族卻是更深層品項），舊版用全域集合
#   （l0_items）判斷角色，遇到這種情況會分類衝突、整列被跳過（說明
#   欄位空白）。改為每一列都依「這一列自己的層級、母件號碼、來源
#   原始母件號碼」逐列判斷，不再依賴任何全域集合。
#   同步修正三個連帶發現的問題：
#     (a) 鑽石型共用（同一品項在同一張訂單底下被多個不同母件同時引用）
#         原本只取第一條路徑，其餘路徑需求被忽略；改為逐路徑加總。
#     (b) 品項若只有自環列、完全沒有其他母件/子件關係（純買賣品項），
#         會整個從圖結構消失、不被任何迴圈處理；改為拓樸排序後額外
#         補上這類「孤兒」品項。
#     (c) 更深層品項尋找「上一層祖先」時，不可跳過中間層直接參照
#         root——中間層品項自己的庫存需要先被扣抵，否則同一root底下
#         不同分支會算出互相矛盾的數字。改為一律參照「緊鄰上一層」，
#         逐層串接，不對「自身也是訂單來源」的情境另外特判。
#   另外新增「工單領料缺口」／「訂單新增需求」兩個拆解欄位（把「既有
#   生產工單本身缺料」跟「這次預示訂單新增的需求」分開顯示），以及
#   「說明」欄位維持未被涵蓋列的診斷安全網（找不到對應資料時明確
#   標註警告，而不是留白讓人誤以為算出0）。主程式Step3對0層列的
#   行級庫存覆寫機制也一併移除，改為MRP函式內部直接依品項編碼查
#   stock_df，避免路徑坍縮造成的欄位污染。
#
# 20260903 修正 v9 → v9.3：
#   [BUGFIX] 主程式組裝「採購單Finalrow_加總版」時，對重複列做
#   drop_duplicates 過濾——這對「建議採購量」「採購量」是對的（MRP
#   已在計算階段對同一品項的所有路徑加總、重複寫入相同值），但對
#   「光陽訂單需求量」「抵扣前需求量」是錯的：這兩欄是每一列自己
#   （該路徑/該原始母件）的原始需求量，不是MRP算好的總量。若同一
#   品項被多個不同原始母件共用（例如 53110-LEC8-E00 同時被
#   53100-AEE2-900-N1S 與 53100-AEE2-900-ZBS 引用），過濾重複列
#   只取第一筆會漏掉其他路徑的原始需求量，讓報表誤以為總需求比
#   實際少。改為這兩欄另外做 groupby+sum，正確反映「這個品項總共
#   被要求了多少數量」。
# =============================================================================

import pandas as pd
import numpy as np
import time
from datetime import datetime, timedelta
from datetime import date
import os
import openpyxl
from openpyxl.styles import Font
from collections import defaultdict
import traceback
import logging
import math


# =============================================================================
# LOG 初始化
# =============================================================================
current_path = os.getcwd()
file_path    = os.path.join(current_path, 'messages.xlsx')
sheet_name   = 'log_訊息'

log_dir = "LOG"
os.makedirs(log_dir, exist_ok=True)

for file in os.listdir(log_dir):
    if file.startswith("MRP_LOG_") and file.endswith(".log"):
        try:
            os.remove(os.path.join(log_dir, file))
            print(f"🗑 已刪除舊 LOG: {file}")
        except Exception as e:
            print(f"❌ 刪除 LOG 失敗: {file} — {e}")

log_filename = os.path.join(log_dir, "MRP_LOG.log")
logging.basicConfig(
    filename=log_filename, filemode='w', level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s', encoding='utf-8'
)
print("✅ LOG 初始化完成")

execution_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
start_time     = time.time()

# ── messages.xlsx 初始化 ──────────────────────────────────────────────────────
if os.path.exists(file_path):
    wb = openpyxl.load_workbook(file_path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    sheet = wb.create_sheet(sheet_name)
else:
    wb    = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = sheet_name

sheet.append([f"執行時間: {execution_time}"])
sheet.append([])
headers = ["資訊等級"] + [f"欄位{i}" for i in range(1, 16)] + ["附加資訊"]
sheet.append(headers)
wb.save(file_path)


# =============================================================================
# 共用工具函數
# =============================================================================

def write_error_log(block_name, error):
    msg = (
        f'\n========================\n'
        f'錯誤區塊 : {block_name}\n'
        f'錯誤內容 : {str(error)}\n'
        f'========================\n'
    )
    print(msg)
    logging.error(msg)
    logging.error(traceback.format_exc())


def log_訊息(level, data, text=None):
    required_columns = [f"資訊{i}" for i in range(0, 20)]

    if isinstance(data, pd.DataFrame):
        if len(data.columns) > len(required_columns):
            data = data.iloc[:, :len(required_columns)]
        elif len(data.columns) < len(required_columns):
            for i in range(len(required_columns) - len(data.columns)):
                data[f"資訊{len(data.columns) + i}"] = ""
        data.columns = required_columns
        for _, row in data.iterrows():
            sheet.append([level] + row.tolist() + [text if text else ""])
        sheet.append([])

    elif isinstance(data, (list, tuple, pd.Index)):
        data_str = ", ".join(map(str, data))
        sheet.append([level] + [data_str] + [""] * 15 + [text if text else ""])

    elif isinstance(data, str):
        sheet.append([level] + [data] + [""] * 15 + [text if text else ""])

    else:
        print(f"⚠️ 無效的輸入型別：{type(data)}")

    wb.save(file_path)


def compare_columns(欄位名稱, 原始欄位, 讀入欄位):
    missing_columns = [col for col in 原始欄位 if col not in 讀入欄位.columns]
    if missing_columns:
        原始df = pd.DataFrame(原始欄位).T
        讀入df = pd.DataFrame(讀入欄位.columns).T
        log_訊息("1 嚴重錯誤 程式中止", f"!!!{欄位名稱} 原始欄位差異，格式錯誤請先更正")
        log_訊息("1 嚴重錯誤 程式中止", missing_columns)
        print(f"!!! {欄位名稱} 欄位不符，缺少：{missing_columns}")
        input("##### 按任一鍵 跳出程式 #####")
        exit()
    else:
        print(f"{欄位名稱} 原始欄位 與 讀入欄位 相同")
        log_訊息("2 INFO", f"{欄位名稱} 原始欄位 與 讀入欄位 相同")


def 計算交付日期(x):
    try:
        days = 10 if (pd.isna(x) or str(x).strip() == '' or int(x) == 0) else int(x)
        delivery_date = datetime.today() + timedelta(days=days)
        if delivery_date.weekday() == 5:
            delivery_date -= timedelta(days=1)
        elif delivery_date.weekday() == 6:
            delivery_date -= timedelta(days=2)
        return delivery_date.strftime('%Y-%m-%d')
    except:
        return '計算交付日期 錯誤'


# =============================================================================
# MRP 核心函數（v9.3：以0層為錨點、逐訂單階梯式扣抵，角色判斷逐列進行）
#
# 完整版本更新歷史請見檔案最頂部的註解欄。
# =============================================================================

import math
import pandas as pd
import numpy as np
from collections import defaultdict, deque


def _find_cycles(child_map, node_pool, max_cycles=10):
    cycles = []
    visited = set()

    def dfs(start, node, path, path_set):
        if len(cycles) >= max_cycles:
            return
        for ci in child_map.get(node, []):
            nxt = ci['child']
            if nxt not in node_pool:
                continue
            if nxt == start and len(path) >= 1:
                cycles.append(list(path))
                return
            if nxt in path_set:
                idx = path.index(nxt)
                cycles.append(list(path[idx:]))
                return
            if nxt in visited:
                continue
            path.append(nxt)
            path_set.add(nxt)
            dfs(start, nxt, path, path_set)
            path.pop()
            path_set.discard(nxt)

    for n in node_pool:
        if len(cycles) >= max_cycles:
            break
        if n in visited:
            continue
        dfs(n, n, [n], {n})
        visited.add(n)

    return cycles


def _calc_available(stock, transit, customer, wip) -> float:
    return (
        float(stock    or 0)
      + float(transit  or 0)
      - float(customer or 0)
      - float(wip      or 0)
    )


def _calc_available_and_wip_gap(stock, transit, customer, wip):
    """
    回傳 (avail, 工單領料缺口)，其中 avail 跟原本 _calc_available() 完全
    等價（可能為負，語意不變：負數代表這個缺口會往上疊加到後續的
    「需採購量」，不能提前把它夾到0，否則會漏算）。

      raw = 庫存 + 在途 - 客戶訂單
      avail = raw - 生產工單未領出          → 跟 _calc_available() 相同
      工單領料缺口 = max(0, 生產工單未領出 - raw)
                    → 純粹是「現有生產工單本身要領料時，庫存/在途
                       就已經不夠了」這個資訊，不管有沒有這次的
                       預示訂單都存在，僅供報表拆解顯示用，不影響
                       avail 本身的計算與後續所有邏輯。

    下游只要把「建議採購量 - 工單領料缺口」，就能還原出「因為這次
    訂單而新增的需求」，兩者相加等於建議採購量。
    """
    raw   = float(stock or 0) + float(transit or 0) - float(customer or 0)
    wip   = float(wip or 0)
    avail = raw - wip
    wip_gap = max(0.0, -avail)
    return avail, wip_gap


def _apply_moq(suggest: float, moq: float) -> int:
    if suggest <= 0:
        return 0
    return int(math.ceil(max(suggest, float(moq or 0))))


def _lookup(stock_df, part):
    Z = {'庫存數量':0,'未進貨(在途量)':0,'客戶訂單':0,'生產工單未領出':0,'最低訂購量':0}
    if stock_df is None or part not in stock_df.index:
        return Z
    r = stock_df.loc[part]
    return {k: float(r.get(k,0) or 0) for k in Z}


def MRP_採購件向下抵扣_v8(
    df: pd.DataFrame,
    stock_df: pd.DataFrame = None,
    original_order: list = None
):
    """
    MRP 向下抵扣計算 v8 —— 「0層錨點、逐訂單階梯式扣抵」版本。

    業務規則（依實際人工試算表逆推確認）：
      1. 每一張「原始母件訂單」對 0 層品項的原始需求，先用該原始母件
         自身的庫存（半成品/成品庫存）淨額，得到「需製造量」。
         （若該0層品項自己也是一張獨立訂單來源，則不淨額，直接用其
          原始訂單量，避免和累計池的可用庫存重複扣抵。）
      2. 0 層品項自己維護一個「可用庫存池」，依 original_order 給定的
         順序，逐張訂單用該筆「需製造量」扣抵，扣不夠的部分才需採購。
      3. 任何更深層的品項（1層、2層...），不論深度多少，都不是從
         「直接上一層」逐層傳遞，而是統一參照「其0層祖先」在同一張
         訂單當下算出的『採購』結果（乘上0層祖先到自己的累計使用量
         比例），作為自己的原始需求，再用自己的庫存池、依同樣的訂單
         順序做階梯式扣抵。
      4. 若同一個品項被多個不同「0層家族」共用（例如跨兩個完全不同的
         成品分別展開後在底層匯集到同一個原物料），扣抵池不分家族、
         直接依 original_order 的全域順序連續扣抵（先處理到的訂單先
         扣，池子接續用到下一個家族的訂單）。
      5. 最終「建議採購量」／「採購量」＝該品項所有階梯步驟「採購」
         數字的加總（不是重複貼一個全域值），MOQ 套用在加總後的總量。

    Returns
    -------
    (df_result, shared_parts_report, shared_contrib_detail, shared_impact_report)
    """

    REQUIRED = [
        '來源原始母件號碼','母件號碼','搜尋_子件號碼','品項編碼','母件路徑',
        '層級','展開後使用量','光陽訂單需求量','來源合計',
        '庫存數量','未進貨(在途量)','客戶訂單','生產工單未領出','最低訂購量',
    ]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"v8 缺少欄位 → {missing}")

    df = df.copy()
    for col in ['層級','光陽訂單需求量','展開後使用量','來源合計',
                '庫存數量','未進貨(在途量)','客戶訂單',
                '生產工單未領出','最低訂購量']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    df['展開後使用量'] = df['展開後使用量'].round(6)
    df['建議採購量']   = 0
    df['採購量']       = 0
    df['工單領料缺口'] = 0
    df['訂單新增需求'] = 0
    df['說明']         = ''
    df['是否被抵扣']   = 'N'
    df['抵扣前需求量'] = df['光陽訂單需求量']

    if original_order is None:
        seen = {}
        for v in df['來源原始母件號碼']:
            v = str(v).strip()
            if v not in seen:
                seen[v] = len(seen)
        original_order = sorted(seen, key=lambda x: seen[x])

    # ─────────────────────────────────────────────────────────────────────
    # STEP 1：建立全域 BOM 圖（母件→子件），去重；parent_map 供共用件偵測
    # ─────────────────────────────────────────────────────────────────────
    child_map  = defaultdict(list)
    parent_map = defaultdict(dict)
    seen_edges = set()
    self_loop_skipped = []

    for row in df[['母件號碼', '搜尋_子件號碼', '展開後使用量']].itertuples(index=False):
        parent = str(row[0]).strip()
        child  = str(row[1]).strip()
        usage  = round(float(row[2]), 6)
        if not parent or not child:
            continue
        if parent == child:
            self_loop_skipped.append(parent)
            continue
        edge = (parent, child)
        if edge in seen_edges:
            continue
        seen_edges.add(edge)
        child_map[parent].append({'child': child, 'usage': usage})
        parent_map[child][parent] = usage

    all_children = set(parent_map.keys())
    all_parents  = set(child_map.keys())
    root_nodes   = all_parents - all_children

    # ─────────────────────────────────────────────────────────────────────
    # STEP 1.5：品項自己也是訂單來源的自身需求彙總（母件==子件 且 層級==0）
    # ─────────────────────────────────────────────────────────────────────
    self_order_demand = defaultdict(float)
    for row in df[['母件號碼', '搜尋_子件號碼', '層級', '光陽訂單需求量']].itertuples(index=False):
        parent = str(row[0]).strip()
        child  = str(row[1]).strip()
        if parent == child and int(row[2]) == 0:
            self_order_demand[parent] += float(row[3])

    # ─────────────────────────────────────────────────────────────────────
    # STEP 2：共用件偵測 — in-degree > 1 的品項
    # ─────────────────────────────────────────────────────────────────────
    shared_parts = {
        child: parents for child, parents in parent_map.items()
        if len(parents) > 1
    }
    shared_report_rows = []
    for child, parents in shared_parts.items():
        shared_report_rows.append({
            '品項編碼'  : child,
            '母件數量'  : len(parents),
            '母件清單'  : ', '.join(sorted(parents.keys())),
            '自身也是訂單來源' : 'Y' if child in self_order_demand else 'N',
        })
    shared_parts_report = pd.DataFrame(shared_report_rows).sort_values(
        by='母件數量', ascending=False
    ).reset_index(drop=True) if shared_report_rows else pd.DataFrame(
        columns=['品項編碼','母件數量','母件清單','自身也是訂單來源']
    )

    # ─────────────────────────────────────────────────────────────────────
    # STEP 3：拓樸排序 —— 僅用於偵測 BOM 循環引用，不再用於數值計算
    # ─────────────────────────────────────────────────────────────────────
    in_degree = {node: len(parent_map.get(node, {})) for node in all_children}
    for node in all_parents:
        in_degree.setdefault(node, 0)
    queue = deque([n for n, d in in_degree.items() if d == 0])
    topo_order = []
    in_degree_work = dict(in_degree)
    while queue:
        node = queue.popleft()
        topo_order.append(node)
        for ci in child_map.get(node, []):
            child = ci['child']
            in_degree_work[child] -= 1
            if in_degree_work[child] == 0:
                queue.append(child)
    unresolved = [n for n, d in in_degree_work.items() if d > 0]
    if unresolved:
        cycles = _find_cycles(child_map, set(unresolved))
        cycle_msg = "\n".join(
            "  循環路徑 " + str(i+1) + " ： " + " → ".join(c + [c[0]])
            for i, c in enumerate(cycles[:10])
        )
        raise ValueError(
            f"v8 偵測到 BOM 循環引用，無法完成拓樸排序。\n"
            f"共 {len(unresolved)} 個品項受影響（含循環本身及其下游）。\n"
            f"實際找到的循環路徑（僅顯示品項間的循環關係，非全部受影響清單）：\n"
            f"{cycle_msg if cycle_msg else '  （循環路徑較複雜，建議直接檢查以下品項在生管BOM的母件/子件設定）'}\n"
            f"請檢查生管BOM中這些品項的「母件號碼」「子件號碼」設定是否有互相指向的錯誤。\n"
            f"（註：母件==子件 的列已自動排除 {len(set(self_loop_skipped))} 筆，不在此列）"
        )

    # [重要修正] 把「只有自環列、完全沒有其他母件/子件關係」的品項也
    # 補進 topo_order。這種品項（純買賣品項：訂單來源就是它自己，
    # 生管BOM裡唯一一筆資料是 母件==子件==自己，沒有任何往下展開的
    # 子件）在 STEP1 建圖時，自環邊會被排除、不計入 child_map／
    # parent_map，導致它既不在 all_parents 也不在 all_children 裡，
    # 從而整個消失在圖結構與拓樸排序之外——即使它明明需要被當成
    # 「自身也是訂單來源」處理（比照root netting的自環分支）。
    # 這裡把 df 裡所有出現過的品項編碼，扣掉已經在 topo_order 或
    # root_nodes 裡的，通通補進 topo_order 尾端（這些品項沒有任何
    # BOM子件，不會是別人的「0層祖先」，補在哪個位置都不影響依賴
    # 順序，補在尾端最安全）。
    all_items_in_df = set(df['品項編碼'].astype(str).str.strip().unique())
    already_covered = set(topo_order) | root_nodes
    missing_items = sorted(all_items_in_df - already_covered)
    if missing_items:
        topo_order.extend(missing_items)

    # ─────────────────────────────────────────────────────────────────────
    # 索引建置（效能優化）：一次 groupby，之後全部 O(1) 查找
    # ─────────────────────────────────────────────────────────────────────
    item_row_groups = df.groupby('品項編碼').groups

    # [效能優化] 預先把常用欄位轉成 dict（idx -> 值），避免在迴圈內反覆用
    # df.loc[idx] 建構整列 Series（每次呼叫都有可觀的固定開銷，資料量大時
    # 會累積成主要效能瓶頸）。
    col_母件號碼     = df['母件號碼'].to_dict()
    col_光陽訂單需求量 = df['光陽訂單需求量'].to_dict()
    col_來源合計    = df['來源合計'].to_dict()
    col_展開後使用量   = df['展開後使用量'].to_dict()
    col_母件路徑      = df['母件路徑'].to_dict()

    # 所有要寫回 df 的更新，先收集到一般 Python dict，最後一次性批次寫入
    # （對照：迴圈內每次都呼叫 df.loc[idxs, col]=value，數萬次呼叫的
    #  pandas索引/型別檢查固定開銷會遠超過實際運算本身）
    upd_說明     = {}
    upd_抵扣     = {}
    upd_建議採購量 = {}
    upd_採購量    = {}
    upd_工單領料缺口 = {}
    upd_訂單新增需求 = {}

    # 每個 root 自身的可用庫存（同時拆出root自己的工單領料缺口，供除錯用）
    #
    #   [重要修正] 舊寫法透過「層級0列的庫存欄位」間接取得（該欄位由主程式
    #   Step3 依 層級==0 覆寫為原始母件自身庫存）。但「自身也是訂單來源」
    #   時，展開路徑會少一層，導致更深層的品項（如 5012A-LBA7-900-W）
    #   也被誤判成層級0，跟著被主程式覆寫成「來源母件」的庫存——即使
    #   MRP計算本身沒錯（用獨立查找抓自己的庫存），但輸出Excel裡的
    #   原始「庫存數量」欄位仍會顯示錯誤覆寫後的數字，造成報表誤導。
    #
    #   改為直接用品項編碼向庫存餘量分析表(stock_df)查詢，跟其他品項
    #   的庫存查法完全一致，不再需要「借用某一列、覆寫它的欄位」這種
    #   間接且脆弱的機制，從根本上消除層級誤判帶來的污染風險。
    root_own_avail = {}
    root_own_wip_gap = {}
    for root in root_nodes:
        si_root = _lookup(stock_df, root)
        root_own_avail[root], root_own_wip_gap[root] = _calc_available_and_wip_gap(
            si_root['庫存數量'], si_root['未進貨(在途量)'],
            si_root['客戶訂單'], si_root['生產工單未領出']
        )

    # ─────────────────────────────────────────────────────────────────────
    # 品項自己的角色（0層/自身訂單來源/更深層），一律用「這一列自己的
    # 層級／母件號碼／來源原始母件號碼」逐列判斷——不再用「品項編碼是否
    # 屬於某個全域集合」來分類。
    #
    #   [重要修正 v9] 舊版（v8.x）用一個全域的 l0_items 集合（依品項編碼）
    #   來判斷「這個品項算不算0層品項」，同一個品項編碼只要在『任何一個
    #   地方』被登記成某個根節點的直接子件，就會被歸類成0層品項——但
    #   真實生管BOM資料裡，同一個料號經常被『不同型號/家族』共用，在
    #   A家族底下是0層直接子件，在B家族底下卻可能是第2、3層的深層零件。
    #   全域分類會讓這種品項在其中一種情境下被錯誤分類，導致該筆資料
    #   完全找不到對應的處理邏輯、被整列跳過（說明欄位空白）。
    #
    #   改為：對每一個 (品項編碼, 來源原始母件號碼) 組合，只看『這一列
    #   自己』的層級，加上「這張訂單本身是不是自身也是訂單來源」
    #   （self_loop_originals），判斷它在『這張訂單』底下到底是不是0層：
    #     - 層級==0 且 這張訂單不是自身訂單來源 → 正常0層品項，
    #       需先用「母件自身庫存」淨額（root netting）。
    #     - 層級==0 且 品項編碼==母件號碼（自環標記列本身）→ 品項自己
    #       就是這張訂單的來源，不淨額。
    #     - 層級==0 但這張訂單是自身訂單來源、且品項編碼 != 母件號碼
    #       → 這是「自身也是訂單來源」時路徑坍縮造成的深層品項，雖然
    #       層級顯示0，但其實應該參照「來源原始母件號碼自己」這個0層
    #       祖先的已計算採購步驟，而不是重新淨額。
    #     - 層級>=1 → 一律參照「0層祖先」的已計算採購步驟。
    #   由於分類完全是逐列判斷（不看全域集合），同一個品項編碼在不同
    #   訂單/家族底下即使角色不同，也能各自正確處理。
    # ─────────────────────────────────────────────────────────────────────
    self_loop_originals = set()
    for row in df[['母件號碼', '搜尋_子件號碼', '層級', '來源原始母件號碼']].itertuples(index=False):
        parent, child, level, original = row
        if str(parent).strip() == str(child).strip() == str(original).strip() and int(level) == 0:
            self_loop_originals.add(str(original).strip())

    # 索引：(品項編碼, 來源原始母件號碼) → 列索引（不限層級，取代舊版
    # 分別建立 l0_key_groups／deep_key_groups 兩套索引的做法）
    item_original_groups = df.groupby(['品項編碼', '來源原始母件號碼']).groups
    col_層級 = df['層級'].to_dict()

    global_suggest  = {}   # part -> 總建議採購量（未套MOQ）
    global_purchase = {}   # part -> 總採購量（已套MOQ）
    contributor_log = defaultdict(list)   # part -> [(來源original, 該筆採購量), ...]

    l0_step_purchase = defaultdict(dict)   # {品項: {original: 該筆採購量}} —— 任何品項都可能是別人的「0層祖先」
    l0_own_usage      = {}                 # {(品項, original): 該列展開後使用量}

    # ─────────────────────────────────────────────────────────────────────
    # 統一的逐品項階梯扣抵：依拓樸順序（祖先必定先於子孫被處理）逐一
    # 處理每個非root品項，確保任何品項要參照「0層祖先」的已計算採購
    # 步驟時，該祖先一定已經算完。
    # ─────────────────────────────────────────────────────────────────────
    for part in topo_order:
        if part in root_nodes:
            continue   # root本身是訂單來源，不是庫存品項，不需計算

        si_part    = _lookup(stock_df, part)
        avail_part, wip_gap_part = _calc_available_and_wip_gap(
            si_part['庫存數量'], si_part['未進貨(在途量)'],
            si_part['客戶訂單'], si_part['生產工單未領出']
        )
        running_avail  = avail_part
        total_purchase = 0.0

        for original in original_order:
            idxs = item_original_groups.get((part, original))
            if idxs is None or len(idxs) == 0:
                continue

            # [重要修正] 同一個品項在「同一張訂單」底下，可能被多個不同的
            # 母件同時引用（鑽石型共用結構，例如 50228-PWB1-900 同時是
            # 5010A-ALE8-E00-A、50116-ALE8-E00、50126-ZWA2-M00、
            # 50127-ZWA2-M00 這四個母件的子件）。idxs 在這種情況下會有
            # 多筆列索引，每一筆代表一條獨立路徑的貢獻，必須逐一計算
            # 各自的需求量再加總，不能只取第一筆（否則其餘路徑的需求
            # 會被完全忽略，導致總量嚴重低估）。
            need_total_this_original = 0.0
            usage_sum = 0.0
            for idx0 in idxs:
                level_x    = int(col_層級[idx0])
                parent_x   = str(col_母件號碼[idx0]).strip()
                demand_raw = float(col_光陽訂單需求量[idx0])
                usage_x    = float(col_展開後使用量[idx0])
                usage_sum += usage_x

                if level_x == 0 and original not in self_loop_originals:
                    # 正常0層品項（這張訂單不是自身訂單來源的情況）：
                    # 淨額母件自身庫存
                    #
                    # [重要修正] 不可直接用「光陽訂單需求量」(=來源合計×
                    # 使用量，已經放大過) 去減「母件自身庫存」(母件自己
                    # 的單位，沒有放大)——這是單位不一致！過去所有驗證
                    # 過的案例剛好使用量都是1，掩蓋了這個問題；一旦使用
                    # 量不是1（例如1個母件成品要用2個這個子件），直接
                    # 相減就會算出錯誤結果。
                    #
                    # 正確做法：先用「來源合計」(訂單原始數量，未放大)
                    # 對母件自身庫存淨額，算出母件自己還要製造多少
                    # （母件自己的單位），再乘上使用量，才是這個子件
                    # 真正對應的需求量。
                    source_qty_x = float(col_來源合計[idx0])
                    root_avail = root_own_avail.get(parent_x, 0.0)
                    need_root_units = math.ceil(max(0.0, source_qty_x - root_avail))
                    need_row = math.ceil(need_root_units * usage_x)

                elif level_x == 0 and part == parent_x:
                    # 自環標記列：品項自己就是這張訂單的來源，不淨額
                    need_row = math.ceil(max(0.0, demand_raw))

                else:
                    # 層級>=1，或「自身訂單坍縮造成的深層品項」：一律參照
                    # 「直接上一層祖先」同一張訂單的已計算採購步驟 × 使用量比例，
                    # 逐層串接（不可跳過中間層直接參照root，否則會漏算
                    # 中間層品項自己的庫存淨額）。
                    #
                    # [重要修正] 原本這裡對「自身訂單坍縮」情境特判成
                    # l0_ancestor=path_tokens[0]（直接跳到root/original），
                    # 這是錯的——會導致中間層品項（例如 53100-AEE2-900-1-PC，
                    # 自己有280的庫存）被完全跳過，底下的孫層品項直接參照
                    # root的採購量，沒有先扣掉中間層自己的庫存，導致同一個
                    # root底下的不同分支，算出彼此矛盾的數字。
                    #
                    # 正確規則其實很單純、不需要對自身訂單另外特判：
                    #   層級==0（母件路徑只有1段，就是root自己）
                    #     → l0_ancestor = original（它自己就是最上層的參照點）
                    #   層級>=1（母件路徑至少2段："root > 上一層品項 > ..."）
                    #     → l0_ancestor = path_tokens[1]（永遠是"直接上一層"，
                    #       不管是不是自身訂單來源，路徑第2段就是它的直接
                    #       上一層祖先，逐層往下鏈接即可）
                    if level_x == 0:
                        l0_ancestor = original
                    else:
                        path_tokens = [t.strip() for t in str(col_母件路徑[idx0]).split('>')]
                        l0_ancestor = path_tokens[1] if len(path_tokens) > 1 else path_tokens[0]

                    l0_usage_ratio_base = l0_own_usage.get((l0_ancestor, original))
                    l0_step = l0_step_purchase.get(l0_ancestor, {}).get(original)

                    if l0_usage_ratio_base is None or l0_step is None or l0_usage_ratio_base == 0:
                        # 找不到0層祖先的對應資料：以0處理並略過，避免中斷
                        # （正常情況不應該發生，若發生，下方的診斷安全網會標註出來）
                        incoming_demand = 0.0
                    else:
                        usage_ratio = usage_x / l0_usage_ratio_base
                        incoming_demand = l0_step * usage_ratio
                    need_row = math.ceil(max(0.0, incoming_demand))

                need_total_this_original += need_row

            # 供下游品項參照時使用的「本品項使用量」，多條路徑時取加總
            # （代表這個品項在這張訂單裡，透過所有路徑合計被用掉的量）
            l0_own_usage[(part, original)] = usage_sum

            need = int(need_total_this_original)
            purchase_step = math.ceil(max(0.0, need - running_avail))
            if running_avail < 0:
                note = f"需製造{need}，可用0(工單缺口{int(-running_avail)})，採購{purchase_step}"
            else:
                note = f"需製造{need}，可用{int(running_avail)}，採購{purchase_step}"
            running_avail = max(0.0, running_avail - need)

            for i in idxs:
                upd_說明[i] = note
                upd_抵扣[i] = 'Y'

            l0_step_purchase[part][original] = purchase_step
            contributor_log[part].append((original, purchase_step))
            total_purchase += purchase_step

        moq = float(si_part['最低訂購量'] or 0)
        final_purchase = _apply_moq(total_purchase, moq) if total_purchase > 0 else 0
        global_suggest[part]  = total_purchase
        global_purchase[part] = final_purchase

        # 拆分「工單領料缺口」與「訂單新增需求」，兩者相加等於建議採購量。
        # 工單領料缺口是這個品項自己既有生產工單本身就領不到料的缺口，
        # 不管有沒有這次的預示訂單都存在；訂單新增需求則是扣掉這個缺口
        # 後，真正因為這次預示訂單而多出來需要採購的量。
        gap_int  = min(math.ceil(wip_gap_part), int(total_purchase))
        need_int = max(0, int(total_purchase) - gap_int)

        all_idxs = item_row_groups.get(part)
        if all_idxs is not None:
            for i in all_idxs:
                upd_建議採購量[i] = total_purchase
                upd_採購量[i] = final_purchase
                upd_工單領料缺口[i] = gap_int
                upd_訂單新增需求[i] = need_int

    # ─────────────────────────────────────────────────────────────────────
    # [效能優化] 統一一次性把收集好的更新寫回 df（取代迴圈內數萬次的
    # df.loc[idxs, col]=value 呼叫，每個欄位只呼叫一次 map+更新）
    # ─────────────────────────────────────────────────────────────────────
    if upd_說明:
        mapped = pd.Series(df.index.map(upd_說明), index=df.index)
        df['說明'] = mapped.fillna(df['說明'])
    if upd_抵扣:
        s = pd.Series(upd_抵扣)
        df.loc[s.index, '是否被抵扣'] = s.values
    if upd_建議採購量:
        s = pd.Series(upd_建議採購量)
        df.loc[s.index, '建議採購量'] = s.values
    if upd_採購量:
        s = pd.Series(upd_採購量)
        df.loc[s.index, '採購量'] = s.values
    if upd_工單領料缺口:
        s = pd.Series(upd_工單領料缺口)
        df.loc[s.index, '工單領料缺口'] = s.values
    if upd_訂單新增需求:
        s = pd.Series(upd_訂單新增需求)
        df.loc[s.index, '訂單新增需求'] = s.values

    # ─────────────────────────────────────────────────────────────────────
    # 診斷安全網：找出完全沒被計算迴圈處理到的列
    #
    #   [背景] v9 已改成逐列（依這一列自己的層級/母件/來源原始母件號碼）
    #   判斷角色，不再用全域品項編碼分類，理論上不會再有「同一編碼在
    #   不同家族被誤判角色」的問題。但如果生管BOM本身有更根本的資料
    #   缺漏（例如：某個中繼層品項的「母件->子件」連結，在『這張訂單』
    #   對應的原始母件號碼底下完全找不到任何一筆——即這個品項在這張
    #   訂單的展開路徑中斷了），這一列仍然會找不到可用的0層祖先資料，
    #   被以0處理略過，但仍會寫入note（不會空白）。真正會讓「說明」
    #   欄位維持初始空白字串的情況，是這個品項在df中根本沒有對應的
    #   (品項編碼,來源原始母件號碼)索引——這通常代表 BOM 樹狀展開
    #   （expand_components）階段就沒有產生這筆資料，例如生管BOM裡
    #   母件/子件的連結中斷、或原始母件號碼打錯字導致展開時對不起來。
    #   空白代表「完全沒算到」，是資料結構有缺漏的訊號，必須另外排查
    #   生管BOM與展開結果，不能誤以為是單純缺料、採購量為0。這裡主動
    #   找出這種列並寫入明確的診斷訊息，避免留白讓使用者看不出原因。
    # ─────────────────────────────────────────────────────────────────────
    未涵蓋遮罩 = (df['說明'] == '') & (~df['品項編碼'].isin(root_nodes))
    if 未涵蓋遮罩.any():
        for idx in df.index[未涵蓋遮罩]:
            part_x     = df.at[idx, '品項編碼']
            original_x = df.at[idx, '來源原始母件號碼']
            path_x     = df.at[idx, '母件路徑']
            df.at[idx, '說明'] = (
                f"⚠️ 此列未被MRP計算涵蓋（空白代表完全沒算到，不是算出0）。"
                f"母件路徑：{path_x}。可能原因：這個品項編碼（{part_x}）"
                f"在「{original_x}」這張訂單的BOM樹狀展開路徑中，母件/子件"
                f"連結有中斷，請檢查生管BOM相關資料。"
            )

    # 整數化
    for col in ['光陽訂單需求量','建議採購量','採購量','抵扣前需求量','工單領料缺口','訂單新增需求']:
        df[col] = (
            pd.to_numeric(df[col], errors='coerce')
            .fillna(0).apply(lambda x: int(math.ceil(x)))
        )
    df['展開後使用量'] = (
        pd.to_numeric(df['展開後使用量'], errors='coerce')
        .fillna(0).round(6)
    )

    # 共用件貢獻明細（僅列共用件或自身亦為訂單來源者）
    contrib_rows = []
    for part, contribs in contributor_log.items():
        if part in shared_parts or part in self_order_demand:
            for original, amount in contribs:
                contrib_rows.append({
                    '品項編碼': part, '母件': original, '貢獻量': round(amount, 2)
                })
    shared_contrib_detail = pd.DataFrame(contrib_rows) if contrib_rows else pd.DataFrame(
        columns=['品項編碼','母件','貢獻量']
    )

    shared_impact_report = _trace_shared_impact(
        child_map, shared_parts, global_suggest, global_purchase, root_nodes
    )


    return df, shared_parts_report, shared_contrib_detail, shared_impact_report


def _trace_shared_impact(child_map, shared_parts, global_suggest, global_purchase, root_nodes):
    rows = []
    for shared_part in shared_parts.keys():
        visited = {shared_part}
        queue   = deque([(shared_part, 0)])
        while queue:
            node, depth = queue.popleft()
            for ci in child_map.get(node, []):
                child = ci['child']
                if child in visited or child in root_nodes:
                    continue
                visited.add(child)
                rows.append({
                    '共用件品項編碼': shared_part,
                    '下游品項編碼'  : child,
                    '相對深度'      : depth + 1,
                    '建議採購量'    : global_suggest.get(child, ''),
                    '採購量'        : global_purchase.get(child, ''),
                })
                queue.append((child, depth + 1))
    if not rows:
        return pd.DataFrame(columns=['共用件品項編碼','下游品項編碼','相對深度','建議採購量','採購量'])
    return pd.DataFrame(rows).sort_values(
        by=['共用件品項編碼','相對深度','下游品項編碼']
    ).reset_index(drop=True)


# =============================================================================
# 欄位定義
# =============================================================================

生管BOM_原始欄位_columns = [
    "原始母件號碼", "母件號碼", "子件號碼", "訂單種類", "流程種類1", "流程種類2",
    "使用量", "最低訂購量", "供應商", "收貨倉庫", "管子尺寸", "切管尺寸",
    "切管比例", "製造單位", "製造機台", "標準工時", "先行天數", "備註1", "備註2",
]

庫存餘量分析表_原始欄位_columns = [
    '品項編碼', '品項名稱', '供應商名稱', '庫存數量',
    '交貨時間', '庫存 附加數量', '未進貨(在途量)', '客戶訂單', '生產工單未領出',
]

預示訂單原始欄位_columns = ["零件號碼", "合計"]

input_file  = 'your_file.xlsx'
output_file = '樹狀展開結果_含需求量.xlsx'


# =============================================================================
# 讀取 Excel 資料
# =============================================================================

# ── 生管BOM ──────────────────────────────────────────────────────────────────
try:
    df = pd.read_excel(input_file, sheet_name="生管BOM", header=1)
    df = df.dropna(how='all')
    df.columns = df.columns.str.strip()
    compare_columns("生管BOM", 生管BOM_原始欄位_columns, df)
    print('✅ 生管BOM 讀取成功')
except Exception as e:
    write_error_log("讀取生管BOM", e)
    input('讀取生管BOM失敗')
    exit()

# ── 庫存餘量分析表 ────────────────────────────────────────────────────────────
try:
    dfs = pd.read_excel(input_file, sheet_name=None, header=1)

    stock_df_raw = pd.read_excel(input_file, sheet_name='庫存餘量分析表', header=None)
    header_row   = stock_df_raw[
        stock_df_raw.iloc[:, 0].astype(str).str.contains('品項編碼')
    ].index[0]
    stock_df = pd.read_excel(
        input_file, sheet_name='庫存餘量分析表', skiprows=header_row
    )
    stock_df.columns = stock_df.columns.str.strip()

    for col in ['庫存數量', '未進貨(在途量)', '客戶訂單', '生產工單未領出']:
        stock_df[col] = pd.to_numeric(stock_df[col], errors='coerce').fillna(0)

    # 建立索引版（供 _stock_lookup 使用）
    stock_df_indexed = stock_df.set_index('品項編碼')

    # 建立供舊邏輯使用的系列
    df庫存餘量分析表_品項編碼 = dfs["庫存餘量分析表"]['品項編碼']
    dfs["庫存餘量分析表"]     = dfs["庫存餘量分析表"].set_index('品項編碼').fillna(0)
    df_庫存餘量分析表         = dfs['庫存餘量分析表']

    print('✅ 庫存餘量分析表 讀取成功')
except Exception as e:
    write_error_log("讀取庫存餘量分析表", e)
    input('讀取庫存餘量分析表失敗')
    exit()

# ── 預示訂單 ──────────────────────────────────────────────────────────────────
try:
    forecast_df = pd.read_excel(input_file, sheet_name='預示訂單', header=1)
    forecast_df.columns = forecast_df.columns.str.strip()
    compare_columns("預示訂單原始欄位", 預示訂單原始欄位_columns, forecast_df)
    log_訊息("2 INFO", f"讀入預示訂單 {forecast_df.shape[0]} 筆資料")
    print('✅ 預示訂單 讀取成功')
except Exception as e:
    write_error_log("讀取預示訂單", e)
    input('讀取預示訂單失敗')
    exit()


# =============================================================================
# 資料一致性檢查
# =============================================================================

# ── 生管BOM 重複檢查 ─────────────────────────────────────────────────────────
dfs['生管BOM'] = dfs['生管BOM'].dropna(how='all')
duplicate_indices = dfs['生管BOM'][
    dfs['生管BOM'].duplicated(
        subset=["原始母件號碼", "母件號碼", "子件號碼"], keep=False
    )
].index
duplicated_data = dfs['生管BOM'].loc[duplicate_indices]

if len(duplicated_data) > 0:
    log_訊息("1", "生管BOM 母件號碼/子件號碼 資料重複，非常嚴重需馬上處理")
    log_訊息("1", duplicated_data)
    print("!!! 生管BOM 資料重複 程式不執行 !!!")
    input("##### 按任一鍵 跳出程式 #####")
    exit()

# ── 生管BOM 必要欄位空白檢查 ─────────────────────────────────────────────────
#   [新增] 原本「未建Ecount品項編碼」檢查若遇到「母件號碼」本身是空白
#   （NaN），只會印出一個看不出是哪幾列的 [nan]，很難排查。這裡改為
#   直接檢查「原始母件號碼／母件號碼／子件號碼」這三個必要欄位是否有
#   空白，並標出實際 Excel 列號（對照Excel開啟時看到的列數：header=1
#   代表欄名在Excel第2列，資料從第3列開始，所以 Excel列號 = DataFrame
#   索引 + 3），方便直接定位到來源檔案哪幾列需要補資料。
必要欄位 = ["原始母件號碼", "母件號碼", "子件號碼"]
生管BOM_空白列 = dfs['生管BOM'][dfs['生管BOM'][必要欄位].isna().any(axis=1)]
if not 生管BOM_空白列.empty:
    excel_列號 = (生管BOM_空白列.index + 3).tolist()   # header=1 → 資料從Excel第3列起
    錯誤訊息文字 = (
        f"!!! 生管BOM 第 {excel_列號} 列（Excel實際列號）的「原始母件號碼/母件號碼/子件號碼」"
        f"其中有欄位是空白，請先補齊再重新執行 !!!"
    )
    # [修正] 把「含實際列號」的完整錯誤訊息文字也寫入 messages.xlsx，
    #   而不是只寫入不含列號的通用說明——這樣即使事後只看記錄檔、
    #   沒有保留當時的終端機畫面，也能直接知道是哪幾列出問題。
    log_訊息("1 嚴重錯誤 程式中止", 錯誤訊息文字)
    log_訊息("1 嚴重錯誤 程式中止", 生管BOM_空白列)
    print(錯誤訊息文字)
    print(生管BOM_空白列[必要欄位].to_string())
    input("##### 按任一鍵 跳出程式 #####")
    exit()

df生管BOM_母件號碼       = dfs['生管BOM']['母件號碼']
df生管BOM               = dfs['生管BOM']
df生管BOM_子件號碼_unique = (
    df生管BOM.drop_duplicates(subset=['子件號碼']).set_index('子件號碼')
)

df生管BOM_columns = dfs['生管BOM'].columns.to_list()
dfs['生管BOM']    = dfs['生管BOM'].set_index('母件號碼').fillna(0)

# ── 預示訂單母件 未建 生管BOM 檢查 ───────────────────────────────────────────
df預示訂單_零件號碼 = dfs['預示訂單']['零件號碼']
dfs["預示訂單"]    = dfs["預示訂單"].set_index('零件號碼').fillna(0)

預示訂單未建BOM = df預示訂單_零件號碼[~df預示訂單_零件號碼.isin(df生管BOM_母件號碼)]
if not 預示訂單未建BOM.empty:
    log_訊息("1 嚴重錯誤 程式中止", "預示訂單母件號碼未建入生管BOM")
    print("預示訂單 母件號碼 未建入 生管BOM：", 預示訂單未建BOM.tolist())
    input("##### 按任一鍵 跳出程式 #####")
    exit()

# ── 生管BOM 母件 未建 Ecount 品項編碼 檢查 ───────────────────────────────────
Ecount未建品項編碼 = df生管BOM_母件號碼[
    ~df生管BOM_母件號碼.isin(df庫存餘量分析表_品項編碼)
]
if not pd.DataFrame(Ecount未建品項編碼).empty:
    log_訊息("1 嚴重錯誤 程式中止", "生管BOM母件號碼未建Ecount品項編碼")
    log_訊息("1 嚴重錯誤 程式中止", pd.DataFrame(Ecount未建品項編碼))
    print("生管BOM 母件號碼 未建 Ecount 品項編碼：", Ecount未建品項編碼.tolist())
    input("##### 按任一鍵 跳出程式 #####")
    exit()

df['使用量'] = pd.to_numeric(df['使用量'], errors='coerce').fillna(0)


# =============================================================================
# BOM 樹狀展開
# =============================================================================

Finalrow         = pd.DataFrame()
展開紀錄_set     = set()
已展開子件_set   = set()


def expand_components(
    df, parent_no, qty_path=1, level=0,
    parent_path=None, source_parent=None,
    source_qty=1, original_parent=None
):
    try:
        global Finalrow, 展開紀錄_set, 已展開子件_set

        if parent_path is None:
            parent_path = []

        df = df[df['原始母件號碼'] == original_parent]
        children = df[df['母件號碼'] == parent_no]
        if children.empty:
            return

        for _, row in children.iterrows():
            child_no = row['子件號碼']

            展開子件_key = (original_parent, parent_no, child_no)
            if 展開子件_key in 已展開子件_set:
                continue
            已展開子件_set.add(展開子件_key)

            usage       = row['使用量']
            total_usage = usage * qty_path
            demand_qty  = np.ceil(total_usage * source_qty)

            path_str  = ' > '.join(parent_path + [parent_no])
            展開key   = (original_parent, source_parent, child_no, level, path_str)
            if 展開key in 展開紀錄_set:
                continue
            展開紀錄_set.add(展開key)

            row_copy = row.copy()
            row_copy['來源原始母件號碼'] = original_parent
            row_copy['來源母件號碼']     = source_parent
            row_copy['母件路徑']         = path_str
            row_copy['搜尋_子件號碼']    = child_no
            row_copy['來源合計']         = source_qty
            row_copy['層級']             = level
            row_copy['展開後使用量']     = total_usage
            row_copy['光陽訂單需求量']   = demand_qty

            Finalrow = pd.concat(
                [Finalrow, pd.DataFrame([row_copy])], ignore_index=True
            )

            if child_no == parent_no:
                continue

            expand_components(
                df, parent_no=child_no,
                qty_path=total_usage, level=level + 1,
                parent_path=parent_path + [parent_no],
                source_parent=source_parent, source_qty=source_qty,
                original_parent=original_parent
            )

    except Exception as e:
        log_訊息("1 嚴重錯誤 程式中止",
                 f"expand_components 展開錯誤：母件號碼 {parent_no}，{str(e)}")
        write_error_log("expand_components", e)
        input('expand_components 展開錯誤')


# ── 展開每個原始母件 ──────────────────────────────────────────────────────────
unique_originals = forecast_df['零件號碼'].dropna().unique()

for original_no in unique_originals:
    try:
        source_qtys = forecast_df.loc[
            forecast_df['零件號碼'] == original_no, '合計'
        ].values
        if len(source_qtys) == 0:
            log_訊息("1 嚴重錯誤 程式中止",
                     f"找不到零件號碼 {original_no} 對應的合計數量，略過")
            continue
        source_qty = source_qtys[0]

        起始母件們 = df[df['原始母件號碼'] == original_no]['母件號碼'].unique()
        for start_parent in 起始母件們:
            expand_components(
                df, parent_no=start_parent,
                qty_path=1, level=0,
                parent_path=[], source_parent=start_parent,
                source_qty=source_qty, original_parent=original_no
            )
    except Exception as e:
        log_訊息("1 嚴重錯誤 程式中止",
                 f"展開原始母件號碼 {original_no} 時發生錯誤：{str(e)}")
        write_error_log("展開原始母件號碼", e)
        input(f'展開原始母件號碼 {original_no} 時發生錯誤')

Finalrow = Finalrow.sort_values(by='原始母件號碼').reset_index(drop=True)
Finalrow['品項編碼'] = Finalrow['子件號碼']
Finalrow_原始展開結果 = Finalrow.copy()


# =============================================================================
# 採購單計算（新流程：MRP 在加總之前執行）
# =============================================================================

today_str = datetime.today().strftime('%Y/%m/%d')

# ── Step 1：取採購單列（保留全部展開列，不去重複）─────────────────────────────
採購單Finalrow = Finalrow[Finalrow['流程種類1'] == '採購單'].copy()

# ── Step 2：merge 子件庫存（每條展開列各自 merge）─────────────────────────────
採購單Finalrow = 採購單Finalrow.merge(
    stock_df[['品項編碼', '庫存數量', '未進貨(在途量)',
              '客戶訂單', '生產工單未領出']],
    how='left', on='品項編碼'
)
for col in ['庫存數量', '未進貨(在途量)', '客戶訂單', '生產工單未領出']:
    採購單Finalrow[col] = pd.to_numeric(
        採購單Finalrow[col], errors='coerce'
    ).fillna(0)
採購單Finalrow['最低訂購量'] = pd.to_numeric(
    採購單Finalrow['最低訂購量'], errors='coerce'
).fillna(0)

# ── Step 3：（v8.1起移除）不再覆寫層級0列的庫存欄位 ──────────────────────────
#
#   [重要修正] 舊版在此把「層級==0」的列，覆寫成「來源原始母件」自身的
#   庫存，供 MRP 內部計算「需製造量」使用。但「自身也是訂單來源」時，
#   展開路徑會少一層，導致原本應該是「更深層品項」的列（例如
#   5012A-LBA7-900-W），也被誤判成層級0而跟著被覆寫成「來源母件」的
#   庫存——即使 MRP 實際計算沒有用到這個被污染的欄位，輸出 Excel 裡
#   的「庫存數量」等原始欄位仍會顯示錯誤覆寫後的數字，誤導閱讀報表
#   的人（此列品項編碼明明是 5012A-LBA7-900-W，欄位卻顯示 5012A-LBA7-900
#   的庫存）。
#
#   改為：MRP_採購件向下抵扣_v8 內部在需要「母件自身庫存」時
#   （root_own_avail），直接依品項編碼向庫存餘量分析表(stock_df)查詢
#   ——跟其他任何品項的庫存查法完全一致，不必、也不再覆寫 df 裡任何
#   一列的欄位。這樣「庫存數量」等原始欄位永遠忠實呈現該列品項編碼
#   自己的真實庫存，「說明」欄則負責呈現計算過程，兩者不會互相污染。

# ── Step 4：【先執行 MRP，加總之前】──────────────────────────────────────────
採購單Finalrow_MRP結果, 共用件偵測報告, 共用件貢獻明細, 共用影響追蹤報告 = MRP_採購件向下抵扣_v8(
    採購單Finalrow,
    stock_df_indexed,
    original_order=list(forecast_df['零件號碼'].dropna().unique())
)

採購單Finalrow_MRP結果Before = 採購單Finalrow_MRP結果.copy()

# ── Step 5：【MRP 後，過濾重複品項編碼；不可再加總】───────────────────────────
#
#   MRP_採購件向下抵扣_v8 在計算階段，已對每個品項編碼算出「總建議採購量」
#   （所有階梯步驟的採購數字加總），並重複寫入該品項編碼在展開結果中出現的
#   每一列（不論它出現在幾條路徑、幾個原始母件底下）。
#
#   因此「採購單Finalrow_MRP結果」(=採購單MRP結果Before加總) 裡，
#   同一品項編碼的多列，建議採購量／採購量 本來就已經是同一個值
#   （差異只在「說明」欄——那裡是刻意保留的逐階敘述，用來對照稽核）。
#
#   若在此對重複列做 groupby+sum，會把已經正確的總量依「該品項出現的
#   展開列數」重複疊加，等於憑空放大採購量。正確做法：同一品項編碼
#   只需取「任一代表列」即可，不需要、也不可以再加總。
採購單Finalrow_加總版 = (
    採購單Finalrow_MRP結果
    .drop_duplicates(subset='品項編碼', keep='first')
    .copy()
)

# [修正] 「光陽訂單需求量」「抵扣前需求量」是每一列自己（該路徑/該原始
# 母件）的原始需求量，不是MRP算好的總量——不像「建議採購量」「採購量」
# 那樣已經在MRP計算階段就對同一品項的所有路徑加總過、重複寫入每一列
# 相同的值。如果同一品項被多個不同的原始母件／路徑共用（例如
# 53110-LEC8-E00 同時被 53100-AEE2-900-N1S 和 53100-AEE2-900-ZBS
# 引用），過濾重複列只取第一筆，會只顯示其中一條路徑的原始需求量
# （例如只顯示6002，漏掉ZBS的200），讓報表誤以為總需求量比實際少。
# 這兩個純顯示/參考用欄位需要另外加總，才能正確反映「這個品項總共
# 被要求了多少數量」。
for _需求量欄位 in ['光陽訂單需求量', '抵扣前需求量']:
    _合計 = (
        採購單Finalrow_MRP結果
        .groupby('品項編碼', as_index=False)[_需求量欄位]
        .sum()
        .rename(columns={_需求量欄位: f'{_需求量欄位}_合計'})
    )
    採購單Finalrow_加總版 = pd.merge(採購單Finalrow_加總版, _合計, on='品項編碼', how='left')
    採購單Finalrow_加總版[_需求量欄位] = 採購單Finalrow_加總版[f'{_需求量欄位}_合計']
    採購單Finalrow_加總版.drop(columns=[f'{_需求量欄位}_合計'], inplace=True)

# 重新套用 MOQ
採購單Finalrow_加總版['最低訂購量'] = pd.to_numeric(
    採購單Finalrow_加總版['最低訂購量'], errors='coerce'
).fillna(0)

採購單Finalrow_加總版['採購量'] = 採購單Finalrow_加總版.apply(
    lambda row: (
        int(math.ceil(max(row['最低訂購量'], row['建議採購量'])))
        if row['建議採購量'] > 0 else 0
    ), axis=1
).astype('Int64')

# ── Step 6：非鐵管採購單 ────────────────────────────────────────────────────
df_non_pipe = 採購單Finalrow_加總版[
    採購單Finalrow_加總版['管子尺寸'].isna() &
    (採購單Finalrow_加總版['建議採購量'] > 0)
].copy()

df_non_pipe['交付日期'] = df_non_pipe['供應商工作日'].apply(計算交付日期)

df_non_pipe_po = pd.DataFrame({
    "日期"            : today_str,
    "序號"            : "1",
    "客戶/供應商編碼" : "",
    "客戶/供應商名稱" : df_non_pipe['供應商'],
    "供應商簡稱"      : "",
    "交易類型"        : "21",
    "承辦人"          : "",
    "收貨倉庫"        : df_non_pipe['收貨倉庫'],
    "交付日期"        : "",
    "品項編碼"        : df_non_pipe['子件號碼'],
    "品項名稱"        : "",
    "品項交付日期"    : "",
    "數量"            : df_non_pipe['採購量'],
    "單位"            : "",
    "摘要"            : "",
    "摘要1"           : df_non_pipe['備註1'],
    "摘要2"           : df_non_pipe['備註2'],
})
df_non_pipe_po = df_non_pipe_po.sort_values(
    by="客戶/供應商名稱"
).reset_index(drop=True)


# =============================================================================
# 鐵管採購單計算（邏輯不變）
# =============================================================================

鐵管Finalrow = 採購單Finalrow_加總版[
    採購單Finalrow_加總版['管子尺寸'].notna()
].copy()
鐵管Finalrow = 鐵管Finalrow[鐵管Finalrow['光陽訂單需求量'] > 0]
鐵管Finalrow_original = 鐵管Finalrow.copy()

鐵管Finalrow['建議採購量'] = (
    鐵管Finalrow['光陽訂單需求量']
  + 鐵管Finalrow['生產工單未領出']
  - 鐵管Finalrow['庫存數量']
  - 鐵管Finalrow['未進貨(在途量)']
).fillna(0)

# 切管比例缺失檢查
mask_切管比例異常 = (
    鐵管Finalrow['切管比例'].isna() | (鐵管Finalrow['切管比例'] == 0)
)
切管比例異常資料 = 鐵管Finalrow.loc[
    mask_切管比例異常,
    ['原始母件號碼', '母件號碼', '子件號碼', '管子尺寸', '切管比例']
]
if not 切管比例異常資料.empty:
    log_訊息("1 嚴重錯誤 程式中止", "切管比例為 NaN/0/空白，鐵管無法計算所需支數")
    log_訊息("1 嚴重錯誤 程式中止", 切管比例異常資料)

鐵管Finalrow['所需鐵管支數'] = pd.NA
mask_可計算 = 鐵管Finalrow['切管比例'] > 0
鐵管Finalrow.loc[mask_可計算, '所需鐵管支數'] = (
    np.ceil(
        鐵管Finalrow.loc[mask_可計算, '建議採購量'] /
        鐵管Finalrow.loc[mask_可計算, '切管比例']
    )
).astype('Int64')

鐵管Finalrow['建議採購量'] = 鐵管Finalrow['建議採購量'].apply(
    lambda x: x if x > 0 else 0
)
鐵管Finalrow['所需鐵管支數'] = np.ceil(
    鐵管Finalrow['建議採購量'] / 鐵管Finalrow['切管比例']
).astype('Int64')

鐵管Finalrow.loc[:, '品項編碼'] = 鐵管Finalrow['管子尺寸']

鐵管光陽訂單需求量加總 = (
    鐵管Finalrow
    .groupby('管子尺寸', as_index=False)['所需鐵管支數']
    .sum()
    .rename(columns={'所需鐵管支數': '所需鐵管支數_合計'})
)
鐵管Finalrow_去重複 = 鐵管Finalrow.drop_duplicates(subset='管子尺寸', keep='first')
鐵管Finalrow_加總版 = pd.merge(
    鐵管Finalrow_去重複, 鐵管光陽訂單需求量加總, on='管子尺寸', how='left'
)
鐵管Finalrow_加總版['所需鐵管支數'] = 鐵管Finalrow_加總版['所需鐵管支數_合計']
鐵管Finalrow_加總版['交付日期']     = 鐵管Finalrow_加總版['供應商工作日'].apply(計算交付日期)

採購單_鐵管 = pd.DataFrame({
    "日期"            : today_str,
    "序號"            : "1",
    "客戶/供應商編碼" : "",
    "客戶/供應商名稱" : 鐵管Finalrow_加總版['供應商'],
    "供應商簡稱"      : "",
    "交易類型"        : "21",
    "承辦人"          : "",
    "收貨倉庫"        : 鐵管Finalrow_加總版['收貨倉庫'],
    "交付日期"        : "",
    "品項編碼"        : 鐵管Finalrow_加總版['管子尺寸'],
    "品項名稱"        : "",
    "品項交付日期"    : "",
    "數量"            : 鐵管Finalrow_加總版['所需鐵管支數'],
    "單位"            : "",
    "摘要"            : "",
    "摘要1"           : 鐵管Finalrow_加總版['備註1'],
    "摘要2"           : 鐵管Finalrow_加總版['備註2'],
    "最低訂購量"      : 鐵管Finalrow_加總版['最低訂購量'],
})
採購單_鐵管 = 採購單_鐵管.sort_values(by="客戶/供應商名稱").reset_index(drop=True)

採購單_鐵管_merge_stock_df = 採購單_鐵管.copy()
採購單_鐵管_merge_stock_df = 採購單_鐵管_merge_stock_df.merge(
    stock_df[['品項編碼', '庫存數量', '未進貨(在途量)', '生產工單未領出']],
    how='left', on='品項編碼'
)
採購單_鐵管_merge_stock_df['鐵管建議採購量'] = (
    採購單_鐵管_merge_stock_df['數量']
  + 採購單_鐵管_merge_stock_df['生產工單未領出']
  - 採購單_鐵管_merge_stock_df['庫存數量']
  - 採購單_鐵管_merge_stock_df['未進貨(在途量)']
).fillna(0)
採購單_鐵管_merge_stock_df['鐵管建議採購量'] = (
    採購單_鐵管_merge_stock_df['鐵管建議採購量'].apply(lambda x: x if x > 0 else 0)
)
採購單_鐵管_merge_stock_df['最低訂購量'] = pd.to_numeric(
    採購單_鐵管_merge_stock_df['最低訂購量'], errors='coerce'
).fillna(0)
採購單_鐵管_merge_stock_df['鐵管建議採購量'] = 採購單_鐵管_merge_stock_df.apply(
    lambda row: max(row['最低訂購量'], row['鐵管建議採購量'])
                if row['鐵管建議採購量'] > 0 else 0, axis=1
).astype('Int64')

採購單_鐵管['數量'] = 採購單_鐵管_merge_stock_df['鐵管建議採購量']
採購單_鐵管.drop(columns=['最低訂購量'], inplace=True, errors='ignore')


# =============================================================================
# 生產工單
# =============================================================================

生產工單Finalrow = Finalrow[Finalrow['流程種類2'] == '生產工單']
生產工單Finalrow = 生產工單Finalrow[生產工單Finalrow['光陽訂單需求量'] > 0]

df_生產工單 = pd.DataFrame({
    "日期"       : today_str,
    "序號"       : "1",
    "生產數量"   : 生產工單Finalrow['來源合計'],
    "領料日期"   : "",
    "建單人員"   : "",
    "客戶編碼"   : "21",
    "供應商名"   : 生產工單Finalrow['製造單位'],
    "工令日期"   : "",
    "預計領料日" : "",
    "母件名稱"   : 生產工單Finalrow['母件號碼'],
    "番號/規格"  : "",
    "品項名稱"   : 生產工單Finalrow['子件號碼'],
    "用量"       : 生產工單Finalrow['展開後使用量'],
    "數量"       : 生產工單Finalrow['光陽訂單需求量'],
    "材料批號"   : "",
    "BOM類型"    : "",
    "備註1"      : 生產工單Finalrow['備註1'],
    "備註2"      : 生產工單Finalrow['備註2'],
})
df_生產工單 = df_生產工單.sort_values(by="供應商名").reset_index(drop=True)


# =============================================================================
# 輸出 Excel
# =============================================================================

with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
    共用件偵測報告.to_excel(
        writer, sheet_name='共用件偵測報告', index=False)
    共用件貢獻明細.to_excel(
        writer, sheet_name='共用件貢獻明細', index=False)
    共用影響追蹤報告.to_excel(
        writer, sheet_name='共用影響追蹤報告', index=False)
    Finalrow_原始展開結果.to_excel(
        writer, sheet_name='Finalrow_原始展開結果', index=False)
    採購單Finalrow_MRP結果Before.to_excel(
        writer, sheet_name='採購單MRP結果Before加總', index=False)
    採購單Finalrow_加總版.to_excel(
        writer, sheet_name='採購單Finalrow_加總版', index=False)
    df_non_pipe_po.to_excel(
        writer, sheet_name='採購單_非鐵管', index=False)
    鐵管Finalrow_original.to_excel(
        writer, sheet_name='鐵管Finalrow_original', index=False)
    鐵管Finalrow_加總版.to_excel(
        writer, sheet_name='鐵管Finalrow_加總版', index=False)
    採購單_鐵管.to_excel(
        writer, sheet_name='採購單_鐵管', index=False)
    採購單_鐵管_merge_stock_df.to_excel(
        writer, sheet_name='採購單_鐵管_merge_stock_df', index=False)

print("✅ 完成：展開結果已輸出")

end_time       = time.time()
execution_time = end_time - start_time
log_訊息("1 INFO", f"程式執行時間：{execution_time:.2f} 秒")
print(f"程式執行時間：{execution_time:.2f} 秒")
input("##### 按任一鍵 跳出程式 #####")