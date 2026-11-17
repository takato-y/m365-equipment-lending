#!/usr/bin/env python3
"""flows/*.json から、人が読むためのフロー設計書（docs/flows/FLW-00n.md）を作る。

式や設定の全文は載せない（flows/*.json が正本）。全体図は大きな流れだけ、ステップ一覧は
「No・ステップ・種類・場所・実行条件・実行の記録」の6列にする。手で書く欄（目的・関連要件など）は INFO に持つ。

使い方:
  python3 tools/flow_doc.py            # docs/flows/ に6本を書き出す
"""
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
TODAY = "2026-11-18"

INFO = {
    "FLW-001": dict(trig="Power Apps から呼ぶ（引数: 操作の種類・備品 ID・開始日・返却予定日・目的・記録 ID）",
                    purpose="Power Apps から呼ばれ、申請・一覧・返却・取消を、確認してから所有者の接続で書き込む", req="REQ-002、REQ-003、REQ-007、REQ-104",
                    conc="同時の申請は、記録を「確認中」で作った後の再確認で、ID が小さい方を優先して解決する（基本設計 2章）", vol="1回 1件。1日 数十回",
                    rerun="利用者が画面からやり直す。確認中のまま残った記録は FLW-003 が1時間後にエラーにする",
                    dup="申請を2回送ると2件目は重なりで却下される。返却・取消の2回目は状態が合わず却下される"),
    "FLW-002": dict(trig="貸出記録のアイテムが変更されたとき（1分ごとに確認。状態＝承認待ち、かつ承認開始日時が空のときだけ）",
                    purpose="貸出記録が承認待ちになったら承認を依頼し、結果を記録して申請者に知らせる", req="REQ-004、REQ-005、REQ-006",
                    conc="承認開始日時を ETag の照合つきで先に書くため、同じ記録で2本目が動いても承認依頼は1つだけ", vol="1回 1件。1日 数十件",
                    rerun="承認の途中で止まった記録は、FLW-003 が8日後に期限切れにする。利用者はもう一度申請する", dup="承認開始日時が入った記録では動かない"),
    "FLW-003": dict(trig="毎朝 9:00（日本時間）",
                    purpose="毎朝、開始日を迎えた予約を貸出中にし、返却期限のお知らせ・督促と、止まった記録の後始末をする", req="REQ-005、REQ-006、REQ-008",
                    conc="1日1回だけ動く", vol="1回 数件〜数十件",
                    rerun="詳細画面の「実行」で手動で動かす。同じ日の2回目は最終通知日で送らない", dup="同じ日に2回動いても、メール・督促回数は増えない（TC-019）"),
    "FLW-004": dict(trig="Forms「故障・不具合の報告」に回答が送信されたとき",
                    purpose="Forms の故障の報告を故障報告リストに確認待ちで登録し、Planner にタスクを作り、Teams に知らせる", req="REQ-006、REQ-009",
                    conc="回答ごとに1回動く。同時に動いても互いに影響しない", vol="1回 1件。月に数件",
                    rerun="失敗した回答は、処理ログの時刻から Forms の回答を見て、管理者が故障報告リストに手で登録する",
                    dup="同じ回答で2回動くと、故障報告とタスクが2件できる（管理者が片方を「対象外」にする）"),
    "FLW-005": dict(trig="毎月1日 9:00（日本時間）",
                    purpose="前月の貸出件数・延滞・故障を集計し、明細の CSV を月次レポートに保存して管理者に知らせる", req="REQ-010",
                    conc="月に1回だけ動く", vol="1回 前月の記録の数（数十〜数百件）",
                    rerun="詳細画面の「実行」で手動で動かす。同じ月の CSV は上書きされる", dup="CSV は上書き、メールは2通届く"),
    "FLW-006": dict(trig="故障報告のアイテムが変更されたとき（1分ごとに確認。対応状況が 修理中／完了／対象外 のときだけ）",
                    purpose="管理者が故障報告の対応状況を変えたら、備品台帳の状態と Planner のタスクをそろえる", req="REQ-009",
                    conc="備品の状態は修理中の報告の件数で決めるため、どの順で動いても最後は同じ結果になる", vol="1回 1件。月に数件",
                    rerun="対応状況をもう一度保存し直す（同じ値でも動く）", dup="何度動いても同じ結果になる（TC-024）"),
}

# 全体図は主な流れを手で書く（自動で全手順を描くと長すぎて読めないため）。
# CI はステップ一覧が定義と一致することだけを確かめ、この図の中身は確かめない。gen_flows.py の処理を変えたら、ここも直す。
DIAGRAMS = {
    "FLW-001": """flowchart TD
    T(["Power Apps から呼ぶ"]) --> A["呼び出した人を確定<br>（要求ヘッダー → SharePoint のユーザー ID）"]
    A --> B{"操作の種類"}
    B -- 申請 --> C["入力の検査<br>備品が利用可か・期間の重なりを確認"]
    C --> D["確認中で記録を作成し、もう一度重なりを確認"]
    D --> E["予約表に追加し、承認待ちにする"]
    B -- 一覧 --> F["本人の記録だけを返す"]
    B -- 返却 --> G["本人の貸出中の記録か確認し、返却済みにする"]
    B -- 取消 --> H["本人の開始前の記録か確認し、取消にする"]
    E & F & G & H --> R["アプリへ結果を返す"]
    X["失敗時：作りかけの記録をエラーにし、処理ログに記録・管理者に知らせ、アプリへ応答する"]""",
    "FLW-002": """flowchart TD
    T(["貸出記録が承認待ちになった"]) --> A["記録を読み直し、承認待ちで承認未開始か確認"]
    A --> B["承認開始日時を記録（ETag の照合つき）"]
    B --> C["承認を開始して待つ（最長7日）"]
    C --> D{"応答あり／期限切れ"}
    D -- 応答あり --> W{"記録はまだ承認待ちか"}
    W -- いいえ（取消済みなど） --> Z["結果を書かずに終わる"]
    W -- はい --> R{"結果"}
    R -- 承認 --> E{"応答者が申請者本人で、<br>自己承認を許可が「はい」以外"}
    E -- いいえ --> F["予約済み／貸出中にし、申請者にメール・Teams にカード"]
    E -- はい --> G["却下にし、申請者にメール"]
    R -- 却下・応答者が分からない --> G
    D -- 期限切れ --> H["期限切れにし、申請者にメール"]
    X["失敗時：処理ログに記録し、管理者に知らせる"]""",
    "FLW-003": """flowchart TD
    T(["毎朝 9:00"]) --> A["開始日になった予約を貸出中にする"]
    A --> B["返却予定日が明日の申請者にお知らせのメール"]
    B --> C["返却予定日を過ぎた申請者に督促のメール<br>上限に達したら管理者にメール・Teams にカード"]
    C --> D["承認の途中で止まった承認待ちを期限切れにする"]
    D --> E["作りかけの記録（確認中のまま1時間）をエラーにする"]
    N["5つの処理は別々の範囲で動き、<br>1つが失敗しても残りは動く"]
    X["失敗時：処理ログに記録し、管理者に知らせる"]""",
    "FLW-004": """flowchart TD
    T(["Forms に故障の報告が届いた"]) --> A["回答の詳細を読む<br>管理番号を整える（空白除去・大文字）"]
    A --> B["備品台帳から備品を探す<br>回答者を SharePoint のユーザーとして確定"]
    B --> C["故障報告リストに「確認待ち」で登録"]
    C --> D["Planner の「確認待ち」にタスクを作り、<br>タスクの ID を故障報告に記録"]
    D --> E["Teams にカード（管理番号だけ。内容・報告者は載せない）"]
    N["備品台帳は変えない（管理者が確かめてから対応状況を変える）"]
    X["失敗時：処理ログに記録し、管理者に知らせる"]""",
    "FLW-005": """flowchart TD
    T(["毎月1日 9:00"]) --> A["前月（日本時間）の範囲を決める"]
    A --> B["前月の貸出記録と故障報告を読む"]
    B --> C["明細を作る（記録ID・管理番号・期間・状態・督促回数）<br>数式として動く値は先頭に ' を付ける"]
    C --> D["CSV（UTF-8・BOM 付き）を月次レポートに保存<br>（同じ月は上書き。TC-026 で確認）"]
    D --> E["申請・貸出・延滞・故障の件数を管理者にメール"]
    X["失敗時：処理ログに記録し、管理者に知らせる"]""",
    "FLW-006": """flowchart TD
    T(["管理者が故障報告の対応状況を変えた"]) --> A["故障報告を読み直す"]
    A --> B{"対応状況"}
    B -- 修理中 --> C["Planner のタスクを進行中にする<br>（タスクの ID があるとき）"]
    B -- 完了・対象外 --> D["Planner のタスクを完了にする<br>（タスクの ID があるとき）"]
    C & D --> E["その備品の「修理中」の報告の件数を数える"]
    E --> F{"件数と、備品台帳の今の状態"}
    F -- 1件以上で今が利用可 --> G["備品台帳を修理中にする"]
    F -- 0件で今が修理中 --> H["備品台帳を利用可に戻す"]
    F -- それ以外（廃棄など） --> I["変えない"]
    X["失敗時：処理ログに記録し、管理者に知らせる"]""",
}

CONNECTORS = {
    "shared_sharepointonline": "SharePoint", "shared_office365": "Office 365 Outlook", "shared_teams": "Microsoft Teams",
    "shared_approvals": "承認", "shared_microsoftforms": "Microsoft Forms", "shared_planner": "Planner",
}
OPERATIONS = {
    "HttpRequest": "SharePoint に HTTP 要求を送信", "SendEmailV2": "メールの送信", "PostCardToConversation": "Teams にカードを投稿",
    "StartAndWaitForAnApproval": "承認を開始して待つ", "GetFormResponseById": "Forms の回答の詳細を取得", "CreateTask_V3": "Planner のタスクを作成",
    "UpdateTask_V2": "Planner のタスクを更新", "CreateFile": "SharePoint にファイルを作成",
}
TYPES = {
    "InitializeVariable": "変数の初期化", "SetVariable": "変数の設定", "Compose": "作成", "If": "条件", "Switch": "分岐", "Scope": "範囲",
    "Foreach": "繰り返し", "Query": "絞り込み", "Select": "選択", "Table": "CSV の表の作成", "Response": "アプリへの応答",
}
STATUS = {"Succeeded": "成功", "Failed": "失敗", "Skipped": "スキップ", "TimedOut": "タイムアウト"}


def kind(action):
    t = action.get("type", "")
    if t in ("OpenApiConnection", "OpenApiConnectionWebhook"):
        op = action.get("inputs", {}).get("host", {}).get("operationId", "")
        return OPERATIONS.get(op, op)
    return TYPES.get(t, t)


def ordered(actions):
    """runAfter に従った実行順に並べる"""
    done, out = set(), []
    names = list(actions)
    while len(out) < len(names):
        progressed = False
        for n in names:
            if n in done:
                continue
            deps = [d for d in actions[n].get("runAfter", {}) if d in actions]
            if all(d in done for d in deps):
                done.add(n)
                out.append(n)
                progressed = True
        if not progressed:  # 循環はないはずだが、念のため残りを定義順で出す
            out += [n for n in names if n not in done]
            break
    return out


def condition(action):
    """成功以外でも動く手順だけ、実行条件を書く"""
    parts = []
    for dep, sts in action.get("runAfter", {}).items():
        if sts and sts != ["Succeeded"]:
            parts.append(f"{dep} が{'・'.join(STATUS.get(s, s) for s in sts)}のとき")
    return "、".join(parts)


def secured(action):
    props = action.get("runtimeConfiguration", {}).get("secureData", {}).get("properties", [])
    if not props:
        return ""
    return "入力・出力を隠す" if "outputs" in props else "入力を隠す"


def children(action):
    """(見出し, 子の手順) の組を返す"""
    t = action.get("type")
    if t in ("Scope", "Foreach"):
        return [(None, action.get("actions", {}))]
    if t == "If":
        return [("条件が成り立つとき", action.get("actions", {})), ("成り立たないとき", action.get("else", {}).get("actions", {}))]
    if t == "Switch":
        out = [(f"「{c.get('case')}」のとき", c.get("actions", {})) for c in action.get("cases", {}).values()]
        return out + [("それ以外のとき", action.get("default", {}).get("actions", {}))]
    return []


def rows(actions, where="", counter=None):
    """通し番号の行を作る。どの範囲・条件・分岐の中にあるかは「場所」の列に書く"""
    counter = counter if counter is not None else [0]
    out = []
    for name in ordered(actions):
        a = actions[name]
        counter[0] += 1
        out.append(f"| {counter[0]} | {name} | {kind(a)} | {where} | {condition(a)} | {secured(a)} |")
        for label, sub in children(a):
            if sub:
                out += rows(sub, f"{name}の中" if label is None else f"{name}：{label}", counter)
    return out


def trigger_row(trig, info):
    t = trig.get("type")
    if t == "Request":
        return "Power Apps から呼ぶ"
    if t == "Recurrence":
        return "スケジュール"
    op = trig.get("inputs", {}).get("host", {}).get("operationId", "")
    return {"GetOnUpdatedItems": "SharePoint のアイテムの変更を確認", "CreateFormWebhook": "Forms の回答の送信"}.get(op, op)


def build(flow_id):
    spec = json.loads((ROOT / "flows" / f"{flow_id}.json").read_text(encoding="utf-8"))
    defn, info = spec["definition"], INFO[flow_id]
    trig = next(iter(defn["triggers"].values()))
    used = sorted({c for c in spec["connectors"]})
    step_rows = rows(defn["actions"])
    md = f"""# {spec['displayName']} フロー設計書

## 文書情報

| 項目 | 内容 |
|---|---|
| 文書ID | DD-{flow_id} |
| システム名 | 備品・PC の貸出管理 |
| 版 | 1.0 |
| 作成日 | {TODAY} |

## 改版履歴

| 版 | 日付 | 内容 |
|---|---|---|
| 1.0 | {TODAY} | `flows/{flow_id}.json` から `tools/flow_doc.py` で作成 |

## 1. 概要

| 項目 | 内容 |
|---|---|
| フロー名 | {spec['displayName']} |
| 目的 | {info['purpose']} |
| 関連要件 | {info['req']} |
| トリガー | {info['trig']} |
| 同時実行 | {info['conc']} |
| 想定件数 | {info['vol']} |
| 手順の数 | {len(step_rows)} |

式・設定の全文は [`flows/{flow_id}.json`](../../flows/{flow_id}.json)（生成元は `tools/gen_flows.py`）を参照する。

## 2. 全体図

大きな流れだけを示す。条件・分岐・繰り返しの中の手順は、4章のステップ一覧に書く。

```mermaid
{DIAGRAMS[flow_id]}
```

## 3. トリガー

| 種類 | 内容 |
|---|---|
| {trigger_row(trig, info)} | {info['trig']} |

## 4. ステップ一覧

- No は実行の順の通し番号。「場所」は、その手順がどの範囲・繰り返し・条件・分岐の中で動くかを表す（空欄はフローの直下）。
- 実行条件は、前の手順が成功したとき以外にも動く手順だけ書く。
- 実行の記録は、氏名・メールアドレス・利用目的などを実行の記録に残さない手順に印を付ける（詳細設計 2.4）。

| No | ステップ | 種類 | 場所 | 実行条件 | 実行の記録 |
|---|---|---|---|---|---|
""" + "\n".join(step_rows) + f"""

## 5. 接続

| 接続 | 使うアカウント | 更新期限 |
|---|---|---|
""" + "\n".join(f"| {CONNECTORS.get(c, c)} | フローの所有者（管理者の1人） | なし（所有者のアカウントが有効な間。異動・退職のときに付け替える） |" for c in used) + f"""

## 6. エラー処理・再実行

| 項目 | 内容 |
|---|---|
| エラー時 | 「試行」の中で失敗したら「失敗時」で処理ログに記録し、管理者にメールする |
| 通知先 | 設定「管理者メール」（空ならメールは送らず、処理ログだけ） |
| 再実行の方法 | {info['rerun']} |
| 二重実行時の影響 | {info['dup']} |
"""
    out = ROOT / "docs" / "flows" / f"{flow_id}.md"
    out.write_text(md, encoding="utf-8")
    print(out.relative_to(ROOT))


if __name__ == "__main__":
    for n in range(1, 7):
        build(f"FLW-00{n}")
