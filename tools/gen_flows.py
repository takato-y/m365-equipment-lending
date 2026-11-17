#!/usr/bin/env python3
"""フローの定義（flows/*.json）を作る。

フローの定義（Logic Apps のワークフロー定義）を直接 JSON で書くと、括弧や参照の誤りが起きやすいため、
小さな部品（SharePoint への HTTP 要求・変数の設定など）を組み合わせて作る。
作った flows/*.json は tools/build_package.py でパッケージ（.zip）にして、Power Automate に取り込む。

SharePoint への読み書きは、すべて「SharePoint に HTTP 要求を送信します」で行う。
- 状態の書き換えに ETag（If-Match）を付けて、確認した後に別の操作が変えていたら上書きしない
- リストの ID（GUID）を定義に書かない（リストは表示名で指す）
サイトの URL は {{SITE_URL}} と書き、パッケージを作るときにリポジトリの外の設定ファイルから入れる。

使い方: python3 tools/gen_flows.py   → flows/FLW-001.json などを書き出す
"""
import json
import pathlib

SITE = "{{SITE_URL}}"
JSON_NOMETA = "application/json;odata=nometadata"
JSON_MINIMAL = "application/json;odata=minimalmetadata"
TZ = "Tokyo Standard Time"
SCHEMA = "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#"


# ---- 部品 ----
def after(*names, status=("Succeeded",)):
    return {n: list(status) for n in names}


def _quote(text):
    """Logic Apps の式の中の文字列リテラル（' は '' にする）"""
    return "'" + text.replace("'", "''") + "'"


def template_to_expr(value):
    """"文字 @{式} 文字" の形を、concat(...) の式にする。"@式" はそのまま、ただの文字は文字列リテラルにする"""
    if not isinstance(value, str):
        return json.dumps(value)  # 数値・真偽値
    if value.startswith("@") and not value.startswith("@{"):
        return value[1:]
    parts, i = [], 0
    while i < len(value):
        j = value.find("@{", i)
        if j < 0:
            parts.append(_quote(value[i:]))
            break
        if j > i:
            parts.append(_quote(value[i:j]))
        # 対応する } を探す（' で囲まれた文字列の中の { } は数えない。'' は ' の文字）
        depth, k, in_str = 0, j + 1, False
        while k < len(value):
            ch = value[k]
            if in_str:
                if ch == "'":
                    if k + 1 < len(value) and value[k + 1] == "'":
                        k += 1
                    else:
                        in_str = False
            elif ch == "'":
                in_str = True
            elif ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    break
            k += 1
        if k >= len(value):
            raise ValueError(f"@{{ が閉じていません: {value}")
        parts.append(("expr", value[j + 2:k]))
        i = k + 1
    if not parts:
        return "''"
    if len(parts) == 1 and isinstance(parts[0], tuple):
        return parts[0][1]  # 式1つだけなら型を保つ（数値の列を数値のまま送る）
    rendered = ["string(" + p[1] + ")" if isinstance(p, tuple) else p for p in parts]
    return rendered[0] if len(rendered) == 1 else "concat(" + ", ".join(rendered) + ")"


def value_expr(val):
    if isinstance(val, dict):
        return object_expr(val)
    if isinstance(val, list):
        return "createArray(" + ", ".join(value_expr(v) for v in val) + ")" if val else "json('[]')"
    return template_to_expr(val)


def object_expr(fields):
    """フィールドの辞書から、setProperty を重ねてオブジェクトを作る式を返す。
    値を JSON の文字列に埋め込まないので、利用者が入れた文字（" や改行）でデータの形が崩れない（注入を防ぐ）。"""
    expr = "json('{}')"
    for key, val in fields.items():
        expr = f"setProperty({expr}, {_quote(key)}, {value_expr(val)})"
    return expr


def html_safe(expr, limit=200):
    """メールの本文（HTML）に入れる自由入力の文: 特殊文字を無害化し、URL を崩し、長さを切る"""
    return (f"replace(replace(replace(replace(take(coalesce({expr}, ''), {limit}), '&', '&amp;'), '<', '&lt;'), '>', '&gt;'), '://', '[:]//')")


def md_safe(expr, limit=200):
    """承認依頼の本文（Markdown）・Teams のカードに入れる自由入力の文: リンクや書式にならないよう記号を全角にし、改行を空白にする"""
    e = f"take(coalesce({expr}, ''), {limit})"
    for a, b in (("[", "［"), ("]", "］"), ("(", "（"), (")", "）"), ("<", "＜"), (">", "＞"), ("*", "＊"), ("_", "＿"), ("#", "＃"), ("`", "｀"), ("|", "｜"), ("\\", "＼"), ("://", "[:]//"), ("www.", "www[.]"), ("WWW.", "WWW[.]")):
        e = f"replace({e}, {_quote(a)}, {_quote(b)})"
    return f"replace(replace({e}, decodeUriComponent('%0A'), ' '), decodeUriComponent('%0D'), ' ')"


SECURE = {"secureData": {"properties": ["inputs", "outputs"]}}
SECURE_INPUTS = {"secureData": {"properties": ["inputs"]}}  # 作成・選択などのデータ操作は入力だけを隠せる


def secure(action):
    """実行の記録に入力・出力を残さない（利用目的・利用者の情報を扱う手順に付ける）"""
    data_ops = {"Compose", "Select", "Query", "Table"}
    action["runtimeConfiguration"] = SECURE_INPUTS if action.get("type") in data_ops else SECURE
    return action


def sp_http(method, uri, body=None, headers=None, run_after=None, accept=JSON_NOMETA):
    """SharePoint に HTTP 要求を送信します（shared_sharepointonline / HttpRequest）"""
    h = {"Accept": accept}
    if body is not None:
        h["Content-Type"] = JSON_NOMETA
    if headers:
        h.update(headers)
    params = {"dataset": SITE, "parameters/method": method, "parameters/uri": uri, "parameters/headers": h}
    if body is not None:
        params["parameters/body"] = body if isinstance(body, str) else "@string(" + object_expr(body) + ")"
    return {
        "runAfter": run_after or {},
        "type": "OpenApiConnection",
        "inputs": {
            "host": {"connectionName": "shared_sharepointonline", "operationId": "HttpRequest", "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"},
            "parameters": params,
            "authentication": "@parameters('$authentication')",
        },
    }


def sp_merge(list_title, id_expr, etag_expr, fields, run_after=None):
    """アイテムの更新（ETag が一致したときだけ。一致しなければ 412 で失敗し、失敗時の処理に回る）"""
    return sp_http(
        "POST",
        f"_api/web/lists/getbytitle('{list_title}')/items(@{{{id_expr}}})",
        body=fields,
        headers={"IF-MATCH": f"@{{{etag_expr}}}", "X-HTTP-Method": "MERGE"},
        run_after=run_after,
    )


def compose(inputs, run_after=None):
    return {"runAfter": run_after or {}, "type": "Compose", "inputs": inputs}


def set_var(name, value, run_after=None):
    return {"runAfter": run_after or {}, "type": "SetVariable", "inputs": {"name": name, "value": value}}


def init_var(name, typ, value, run_after=None):
    return {"runAfter": run_after or {}, "type": "InitializeVariable", "inputs": {"variables": [{"name": name, "type": typ, "value": value}]}}


def cond(expression, yes, no, run_after=None):
    return {"runAfter": run_after or {}, "type": "If", "expression": expression, "actions": yes, "else": {"actions": no}}


def chain(*pairs):
    """(名前, アクション) を、前のアクションの成功の後に続くようにつなげる"""
    actions, prev = {}, None
    for name, act in pairs:
        if prev is not None and not act.get("runAfter"):
            act["runAfter"] = after(prev)
        actions[name] = act
        prev = name
    return actions


def odata_str(expr):
    """OData のフィルターに入れる文字列（' を '' にする）"""
    return f"replace({expr}, '''', '''''')"


def definition(trigger, actions):
    return {
        "$schema": SCHEMA,
        "contentVersion": "1.0.0.0",
        "parameters": {"$connections": {"defaultValue": {}, "type": "Object"}, "$authentication": {"defaultValue": {}, "type": "SecureObject"}},
        "triggers": {"manual": trigger},
        "actions": actions,
    }


# ---- 共通の式 ----
CALLER = "outputs('呼び出し元')"
TODAY = "outputs('今日')"
SETTING = "outputs('設定')"
LIVE_STATUSES = "(LoanStatus eq '確認中' or LoanStatus eq '承認待ち' or LoanStatus eq '予約済み' or LoanStatus eq '貸出中')"


def jst_midnight_utc(date_expr):
    """日本時間のその日の 0:00 を、UTC の日時の文字列にする（日付の列の書き込み・比較に使う）"""
    return f"convertToUtc(concat({date_expr}, 'T00:00:00'), '{TZ}', 'yyyy-MM-ddTHH:mm:ssZ')"


def prologue():
    """変数・呼び出し元・今日を用意する（失敗しない処理だけ。SharePoint を読む処理は「試行」の中に置く）"""
    return [
        ("結果", init_var("結果", "string", "エラー")),
        ("メッセージ", init_var("メッセージ", "string", "")),
        ("データ", init_var("データ", "string", "[]")),
        ("利用者ID", init_var("利用者ID", "integer", 0)),
        ("新しい記録ID", init_var("新しい記録ID", "integer", 0)),
        ("管理者メール", init_var("管理者メール", "string", "")),
        # 呼び出したユーザー（アプリから渡す値は使わない。Power Apps のトリガーが付ける要求の情報。要検証）
        ("呼び出し元", secure(compose("@toLower(coalesce(triggerOutputs()?['headers']?['x-ms-user-email'], ''))"))),
        ("今日", compose(f"@formatDateTime(convertFromUtc(utcNow(), '{TZ}'), 'yyyy-MM-dd')")),
    ]


def settings_steps():
    """設定リストを読み、{項目: 値} の形にする（「試行」の中で使う）"""
    return [
        ("設定の取得", sp_http("GET", "_api/web/lists/getbytitle('設定')/items?$select=Title,SettingValue&$top=100")),
        ("設定の選択", {
            "type": "Select",
            "inputs": {"from": "@body('設定の取得')?['value']", "select": "@substring(string(setProperty(json('{}'), coalesce(item()?['Title'], ''), coalesce(item()?['SettingValue'], ''))), 1, sub(length(string(setProperty(json('{}'), coalesce(item()?['Title'], ''), coalesce(item()?['SettingValue'], '')))), 2))"},
        }),
        ("設定", compose("@json(concat('{', join(body('設定の選択'), ','), '}'))")),
        ("管理者メールの設定", set_var("管理者メール", f"@{{coalesce({SETTING}?['管理者メール'], '')}}")),
    ]


def outlook_mail(to_expr, subject, body_html, run_after=None):
    """Office 365 Outlook「メールの送信 (V2)」。宛先・氏名・コメントを実行の記録に残さないよう、入力・出力を隠す"""
    return secure({
        "runAfter": run_after or {},
        "type": "OpenApiConnection",
        "inputs": {
            "host": {"connectionName": "shared_office365", "operationId": "SendEmailV2", "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365"},
            "parameters": {"emailMessage/To": to_expr, "emailMessage/Subject": subject, "emailMessage/Body": body_html, "emailMessage/Importance": "Normal"},
            "authentication": "@parameters('$authentication')",
        },
    })


ANY_END = ("Succeeded", "Failed", "Skipped", "TimedOut")


# ---- FLW-001 利用者の操作 ----
def flw001():
    inp = "outputs('入力')"
    period = "outputs('期間')"
    overlap_filter = (
        f"EquipmentId eq @{{{inp}?['equipmentId']}} and StartDate le datetime'@{{{period}?['due']}}' "
        f"and DueDate ge datetime'@{{{period}?['start']}}' and {LIVE_STATUSES}"
    )
    new_id = "body('記録の作成')?['Id']"
    new_etag = "outputs('記録の作成')?['headers']?['ETag']"

    apply_yes = chain(
        ("記録の作成", secure(sp_http("POST", "_api/web/lists/getbytitle('貸出記録')/items", body={
            "Title": f"@{{first(body('備品の取得')?['value'])?['Title']}} @{{{inp}?['start']}}〜@{{{inp}?['due']}}",
            "EquipmentId": f"@{{{inp}?['equipmentId']}}",
            "RequesterId": "@{variables('利用者ID')}",
            "Purpose": f"@{{{inp}?['purpose']}}",
            "StartDate": f"@{{{period}?['start']}}",
            "DueDate": f"@{{{period}?['due']}}",
            "LoanStatus": "確認中",
        }))),
        # 途中で失敗したときに、確認中のまま残さずエラーにするため、作った記録の ID を持っておく
        ("記録IDの保持", set_var("新しい記録ID", f"@{new_id}")),
        # 作った後にもう一度確かめる。自分より ID が小さい重なる記録があれば、先に作られた方を優先する
        ("先の重なりの取得", sp_http("GET", f"_api/web/lists/getbytitle('貸出記録')/items?$select=Id&$top=50&$filter={overlap_filter} and Id lt @{{{new_id}}}")),
        ("先の申請なし", cond(
            {"equals": ["@length(body('先の重なりの取得')?['value'])", 0]},
            chain(
                # 予約表に先に書いてから承認待ちにする（どちらで失敗しても、失敗時の処理で「記録はエラー・予約表なし」にそろう）
                ("予約表に追加", sp_http("POST", "_api/web/lists/getbytitle('予約表')/items", body={
                    "Title": f"@{{string({new_id})}}", "EquipmentId": f"@{{{inp}?['equipmentId']}}",
                    "StartDate": f"@{{{period}?['start']}}", "DueDate": f"@{{{period}?['due']}}", "LoanStatus": "承認待ち",
                })),
                ("承認待ちにする", sp_merge("貸出記録", new_id, new_etag, {"LoanStatus": "承認待ち"})),
                ("受付の結果", set_var("結果", "受付")),
                ("受付のメッセージ", set_var("メッセージ", "申請を受け付けました。承認の結果はメールでお知らせします。")),
            ),
            chain(
                ("先の申請に譲る", sp_merge("貸出記録", new_id, new_etag, {"LoanStatus": "却下", "ProcessNote": "同じ期間に先に申請がありました。期間を変えて、もう一度申請してください。"})),
                ("競合の結果", set_var("結果", "却下")),
                ("競合のメッセージ", set_var("メッセージ", "同じ期間に、ほかの申請があります。期間を変えて申請してください。")),
            ),
        )),
    )

    apply_ok = chain(
        ("期間", compose({"start": f"@{jst_midnight_utc(inp + '?[' + repr('start') + ']')}", "due": f"@{jst_midnight_utc(inp + '?[' + repr('due') + ']')}"})),
        ("備品の取得", sp_http("GET", f"_api/web/lists/getbytitle('備品台帳')/items?$select=Id,Title,EquipStatus&$filter=Id eq @{{{inp}?['equipmentId']}} and EquipStatus eq '利用可'")),
        ("重なりの取得", sp_http("GET", f"_api/web/lists/getbytitle('貸出記録')/items?$select=Id&$top=50&$filter={overlap_filter}")),
        ("受付できる", cond(
            {"and": [
                {"equals": ["@length(body('備品の取得')?['value'])", 1]},
                {"equals": ["@length(body('重なりの取得')?['value'])", 0]},
            ]},
            apply_yes,
            chain(
                ("不可の結果", set_var("結果", "却下")),
                ("不可のメッセージ", set_var("メッセージ", "@{if(equals(length(body('備品の取得')?['value']), 0), 'この備品は今は貸し出せません。', '同じ期間に、ほかの申請・予約があります。期間を変えて申請してください。')}")),
            ),
        )),
    )

    case_apply = chain(
        ("入力", secure(compose({
            "equipmentId": "@int(coalesce(triggerBody()?['number'], 0))",
            "start": "@{trim(coalesce(triggerBody()?['text_1'], ''))}",
            "due": "@{trim(coalesce(triggerBody()?['text_2'], ''))}",
            "purpose": "@{trim(coalesce(triggerBody()?['text_3'], ''))}",
        }))),
        # 日付は yyyy-MM-dd の10文字だけを受け付ける（形が違えば、ticks で失敗して失敗時の処理に回る）
        ("入力の検査", compose(
            "@and(greater(" + inp + "?['equipmentId'], 0), "
            "equals(formatDateTime(" + inp + "?['start'], 'yyyy-MM-dd'), " + inp + "?['start']), "
            "equals(formatDateTime(" + inp + "?['due'], 'yyyy-MM-dd'), " + inp + "?['due']), "
            "greaterOrEquals(" + inp + "?['start'], " + TODAY + "), "
            "greaterOrEquals(" + inp + "?['due'], " + inp + "?['start']), "
            "lessOrEquals(div(sub(ticks(" + inp + "?['due']), ticks(" + inp + "?['start'])), 864000000000), sub(int(" + SETTING + "?['貸出上限日数']), 1)), "
            "lessOrEquals(" + inp + "?['start'], formatDateTime(addDays(" + TODAY + ", int(" + SETTING + "?['予約受付日数'])), 'yyyy-MM-dd')), "
            "greater(length(" + inp + "?['purpose']), 0), lessOrEquals(length(" + inp + "?['purpose']), 500))"
        )),
        ("入力が正しい", cond(
            {"equals": ["@outputs('入力の検査')", True]},
            apply_ok,
            chain(
                ("入力の誤りの結果", set_var("結果", "却下")),
                ("入力の誤りのメッセージ", set_var("メッセージ", "@{concat('入力の内容（期間・目的）を確かめてください。貸出は最長 ', " + SETTING + "?['貸出上限日数'], ' 日、開始日は今日から ', " + SETTING + "?['予約受付日数'], ' 日先までです。')}")),
            ),
        )),
    )

    # 本人の判定は、SharePoint のユーザー ID（整数）で行う（メールと UPN の違いに左右されない）
    me_filter = "RequesterId eq @{variables('利用者ID')}"
    case_list = chain(
        ("自分の記録", secure(sp_http("GET", "_api/web/lists/getbytitle('貸出記録')/items?$select=Id,StartDate,DueDate,LoanStatus,ProcessNote,Equipment/Title&$expand=Equipment"
                                    f"&$filter={me_filter}&$orderby=Id desc&$top=100"))),
        ("一覧の整形", {
            "runtimeConfiguration": SECURE_INPUTS,
            "type": "Select",
            "inputs": {"from": "@body('自分の記録')?['value']", "select": {
                "id": "@item()?['Id']",
                "equipment": "@item()?['Equipment']?['Title']",
                "start": f"@formatDateTime(convertFromUtc(item()?['StartDate'], '{TZ}'), 'yyyy-MM-dd')",
                "due": f"@formatDateTime(convertFromUtc(item()?['DueDate'], '{TZ}'), 'yyyy-MM-dd')",
                "status": "@item()?['LoanStatus']",
                "note": "@item()?['ProcessNote']",
            }},
        }),
        ("一覧のデータ", set_var("データ", "@{string(body('一覧の整形'))}")),
        ("一覧の結果", set_var("結果", "ok")),
    )

    def target_case(prefix, allowed_expr, fields, ok_msg, ng_msg):
        target = f"first(body('{prefix}の対象')?['value'])"
        return chain(
            (f"{prefix}の対象", sp_http("GET", "_api/web/lists/getbytitle('貸出記録')/items?$select=Id,LoanStatus,StartDate"
                                         f"&$filter=Id eq @{{int(coalesce(triggerBody()?['number_1'], 0))}} and {me_filter}", accept=JSON_MINIMAL)),
            (f"{prefix}できる", cond(
                {"and": [{"equals": [f"@length(body('{prefix}の対象')?['value'])", 1]}, {"equals": [allowed_expr.format(t=target), True]}]},
                chain(
                    (f"{prefix}にする", sp_merge("貸出記録", f"{target}?['Id']", f"{target}?['odata.etag']", fields)),
                    (f"{prefix}の予約表", sp_http("GET", f"_api/web/lists/getbytitle('予約表')/items?$select=Id&$filter=Title eq '@{{{target}?['Id']}}'")),
                    (f"{prefix}の予約表から消す", {
                        "type": "Foreach",
                        "foreach": f"@body('{prefix}の予約表')?['value']",
                        "actions": {f"{prefix}の予約の削除": sp_http("POST", "_api/web/lists/getbytitle('予約表')/items(@{items('" + prefix + "の予約表から消す')?['Id']})",
                                                                headers={"IF-MATCH": "*", "X-HTTP-Method": "DELETE"})},
                    }),
                    (f"{prefix}の結果", set_var("結果", "ok")),
                    (f"{prefix}のメッセージ", set_var("メッセージ", ok_msg)),
                ),
                chain(
                    (f"{prefix}不可の結果", set_var("結果", "却下")),
                    (f"{prefix}不可のメッセージ", set_var("メッセージ", ng_msg)),
                ),
            )),
        )

    case_return = target_case(
        "返却", "@equals({t}?['LoanStatus'], '貸出中')",
        {"LoanStatus": "返却済み", "ReturnedDate": f"@{{{jst_midnight_utc(TODAY)}}}"},
        "返却を登録しました。", "この記録は返却できません。返却できるのは、本人の貸出中の記録だけです。",
    )
    case_cancel = target_case(
        "取消",
        "@and(or(equals({t}?['LoanStatus'], '承認待ち'), equals({t}?['LoanStatus'], '予約済み')), greater(formatDateTime(convertFromUtc({t}?['StartDate'], '" + TZ + "'), 'yyyy-MM-dd'), " + TODAY + "))",
        {"LoanStatus": "取消"},
        "取り消しました。", "取り消せません。取り消せるのは、本人の開始前の承認待ち・予約済みの記録だけです。",
    )

    switch = {
        "type": "Switch",
        "expression": "@triggerBody()?['text']",
        "cases": {
            "申請": {"case": "申請", "actions": case_apply},
            "一覧": {"case": "一覧", "actions": case_list},
            "返却": {"case": "返却", "actions": case_return},
            "取消": {"case": "取消", "actions": case_cancel},
        },
        "default": {"actions": chain(
            ("不明な操作の結果", set_var("結果", "却下")),
            ("不明な操作のメッセージ", set_var("メッセージ", "この操作は受け付けていません。")),
        )},
    }
    try_scope = {
        "type": "Scope",
        "runAfter": after("今日"),
        "actions": chain(
            *settings_steps(),
            ("呼び出し元あり", cond(
                {"not": {"equals": [f"@{CALLER}", ""]}},
                chain(
                    # 呼び出した人を SharePoint のユーザーとして確定し、以後はその ID（整数）で本人を判定する
                    ("利用者の確認", secure(sp_http("POST", "_api/web/ensureuser", body={"logonName": f"@{{concat('i:0#.f|membership|', {CALLER})}}"}))),
                    # 解決したユーザーのメール・UPN が、呼び出し元と一致するときだけ進める（別のアカウントに解決されるのを防ぐ）
                    ("利用者の一致", cond(
                        {"or": [
                            {"equals": ["@toLower(coalesce(body('利用者の確認')?['Email'], ''))", f"@{CALLER}"]},
                            {"equals": ["@toLower(coalesce(body('利用者の確認')?['UserPrincipalName'], ''))", f"@{CALLER}"]},
                        ]},
                        chain(
                            ("利用者IDの設定", set_var("利用者ID", "@int(body('利用者の確認')?['Id'])")),
                            ("操作", switch),
                        ),
                        chain(
                            ("利用者不一致の結果", set_var("結果", "却下")),
                            ("利用者不一致のメッセージ", set_var("メッセージ", "利用者を確認できませんでした。管理者に連絡してください。")),
                        ),
                    )),
                ),
                chain(
                    ("呼び出し元なしの結果", set_var("結果", "却下")),
                    ("呼び出し元なしのメッセージ", set_var("メッセージ", "利用者を確認できませんでした。会社のアカウントでアプリを開いてください。")),
                ),
            )),
        ),
    }
    catch_actions = {}
    # 試行の結果（result）は各手順の入力・出力を含みうるため、実行の記録に残さない
    catch_actions["失敗した手順"] = secure({"runAfter": {}, "type": "Query", "inputs": {"from": "@result('試行')", "where": "@equals(item()?['status'], 'Failed')"}})
    # 手順の名前とエラーの種類だけを取り出す（入力・出力の中身は入れない）
    catch_actions["失敗の要約"] = secure({"runAfter": after("失敗した手順"), "type": "Select", "inputs": {"from": "@body('失敗した手順')", "select": "@concat(item()?['name'], ':', coalesce(item()?['code'], ''))"}})
    # 作りかけの記録を「エラー」にする（確認中のまま残すと、その期間の申請が受け付けられなくなるため）
    catch_actions["作りかけの記録あり"] = cond(
        {"greater": ["@variables('新しい記録ID')", 0]},
        chain(
            # 状態を読み直し、まだ「確認中」のときだけ、読んだ ETag を付けてエラーにする（承認待ちに進んだ記録や、別の操作で変わった記録は上書きしない）
            ("作りかけの読み直し", sp_http("GET", "_api/web/lists/getbytitle('貸出記録')/items?$select=Id,LoanStatus&$filter=Id eq @{variables('新しい記録ID')}", accept=JSON_MINIMAL)),
            ("まだ確認中", cond(
                {"equals": ["@first(body('作りかけの読み直し')?['value'])?['LoanStatus']", "確認中"]},
                # 記録をエラーにできたときだけ、その記録の予約表の行を消す（承認待ちに進んだ記録の予約は残す）
                chain(
                    ("作りかけをエラーにする", sp_merge("貸出記録", "variables('新しい記録ID')", "first(body('作りかけの読み直し')?['value'])?['odata.etag']",
                                                  {"LoanStatus": "エラー", "ProcessNote": "処理中に問題が発生したため、申請を受け付けられませんでした。お手数ですが、もう一度申請してください。"})),
                    ("作りかけの予約表", sp_http("GET", "_api/web/lists/getbytitle('予約表')/items?$select=Id&$filter=Title eq '@{variables('新しい記録ID')}'")),
                    ("作りかけの予約表から消す", {
                        "type": "Foreach",
                        "foreach": "@body('作りかけの予約表')?['value']",
                        "actions": {"作りかけの予約の削除": sp_http("POST", "_api/web/lists/getbytitle('予約表')/items(@{items('作りかけの予約表から消す')?['Id']})",
                                                            headers={"IF-MATCH": "*", "X-HTTP-Method": "DELETE"})},
                    }),
                ),
                {},
            )),
        ),
        {},
        run_after={"失敗の要約": list(ANY_END)},
    )
    catch_actions["処理ログに記録"] = sp_http("POST", "_api/web/lists/getbytitle('処理ログ')/items", body={
        "Title": "FLW-001 利用者の操作",
        "Target": "@{concat('操作=', coalesce(triggerBody()?['text'], ''), ' 記録ID=', string(variables('新しい記録ID')))}",
        "Result": "失敗",
        # 申請の本文・個人情報は入れない（手順の名前とエラーの種類だけ）
        "Detail": "@{take(join(body('失敗の要約'), ', '), 1000)}",
        "RunLink": {"Url": "@{concat('https://make.powerautomate.com/environments/', workflow()?['tags']?['environmentName'], '/flows/', workflow()?['name'], '/runs/', workflow()?['run']?['name'])}", "Description": "実行の記録"},
    }, run_after={"作りかけの記録あり": list(ANY_END)})
    catch_actions["管理者に知らせる"] = cond(
        {"not": {"equals": ["@variables('管理者メール')", ""]}},
        {"管理者へのメール": outlook_mail("@variables('管理者メール')", "【備品貸出】処理に失敗しました（FLW-001）",
                                       "<p>FLW-001 利用者の操作で、処理に失敗しました。処理ログの最新の行と、実行の記録を確かめてください。</p>")},
        {},
        run_after={"処理ログに記録": list(ANY_END)},
    )
    catch_actions["失敗の結果"] = set_var("結果", "エラー", run_after={"管理者に知らせる": list(ANY_END)})
    catch_actions["失敗のメッセージ"] = set_var("メッセージ", "処理に失敗しました。時間をおいてもう一度お試しください。続く場合は管理者に連絡してください。", run_after=after("失敗の結果"))
    catch_scope = {"type": "Scope", "runAfter": {"試行": ["Failed", "TimedOut"]}, "actions": catch_actions}
    actions = chain(*prologue())
    actions["試行"] = try_scope
    actions["失敗時"] = catch_scope
    actions["アプリへの応答"] = {
        # 失敗時の処理そのものが失敗しても、アプリには必ず応答を返す
        "runAfter": {"失敗時": list(ANY_END)},
        "type": "Response",
        "kind": "PowerApp",
        "inputs": {
            "statusCode": 200,
            "body": {"result": "@variables('結果')", "message": "@variables('メッセージ')", "data": "@variables('データ')"},
            "schema": {"type": "object", "properties": {
                "result": {"title": "result", "type": "string", "x-ms-content-hint": "TEXT", "x-ms-dynamically-added": True},
                "message": {"title": "message", "type": "string", "x-ms-content-hint": "TEXT", "x-ms-dynamically-added": True},
                "data": {"title": "data", "type": "string", "x-ms-content-hint": "TEXT", "x-ms-dynamically-added": True},
            }},
        },
    }
    trigger = {
        "type": "Request",
        "kind": "PowerAppV2",
        # 実行の記録に、利用目的や要求の情報（メール）を残さない
        "runtimeConfiguration": {"secureData": {"properties": ["outputs"]}},
        "inputs": {"schema": {"type": "object", "properties": {
            "text": {"title": "operation", "type": "string", "description": "申請／一覧／返却／取消", "x-ms-content-hint": "TEXT", "x-ms-dynamically-added": True},
            "text_1": {"title": "startDate", "type": "string", "description": "貸出開始日 yyyy-MM-dd", "x-ms-content-hint": "TEXT", "x-ms-dynamically-added": True},
            "text_2": {"title": "dueDate", "type": "string", "description": "返却予定日 yyyy-MM-dd", "x-ms-content-hint": "TEXT", "x-ms-dynamically-added": True},
            "text_3": {"title": "purpose", "type": "string", "description": "利用目的（500文字まで）", "x-ms-content-hint": "TEXT", "x-ms-dynamically-added": True},
            "number": {"title": "equipmentId", "type": "number", "description": "備品台帳の ID", "x-ms-content-hint": "NUMBER", "x-ms-dynamically-added": True},
            "number_1": {"title": "loanId", "type": "number", "description": "貸出記録の ID（返却・取消）", "x-ms-content-hint": "NUMBER", "x-ms-dynamically-added": True},
        }, "required": ["text"]}},
    }
    return {
        "displayName": "FLW-001 利用者の操作",
        "description": "Power Apps から呼ぶ。申請・一覧・返却・取消を、確認してから所有者の接続で書き込む",
        "connectors": ["shared_sharepointonline", "shared_office365"],
        "definition": definition(trigger, actions),
    }


def teams_card(title_expr, lines, run_after=None):
    """Microsoft Teams「チャットやチャネルにカードを投稿する」（フロー ボットで、チーム「備品貸出管理（デモ）」の「貸出通知」へ）。
    カードは部品（オブジェクト）として組み立て、文字列の連結で JSON を作らない。lines は (見出し, 値の式) の組"""
    card = {
        "type": "AdaptiveCard", "version": "1.4",
        "body": [{"type": "TextBlock", "text": title_expr, "weight": "Bolder", "wrap": True}]
                + [{"type": "TextBlock", "text": f"@{{concat({_quote(label + '：')}, {expr})}}", "wrap": True} for label, expr in lines],
    }
    return secure({
        "runAfter": run_after or {},
        "type": "OpenApiConnection",
        "inputs": {
            "host": {"connectionName": "shared_teams", "operationId": "PostCardToConversation", "apiId": "/providers/Microsoft.PowerApps/apis/shared_teams"},
            "parameters": {
                "poster": "Flow bot", "location": "Channel",
                "body/recipient/groupId": "{{TEAM_ID}}", "body/recipient/channelId": "{{CHANNEL_ID}}",
                "body/messageBody": "@string(" + object_expr(card) + ")",
            },
            "authentication": "@parameters('$authentication')",
        },
    })


def catch_block(flow_label, extra_first=None, target_expr="@{concat('貸出記録ID=', string(coalesce(triggerBody()?['ID'], 0)))}"):
    """失敗時の処理（処理ログへの記録・管理者へのメール）。どの手順が失敗しても最後まで進む"""
    acts = {}
    # 試行の結果（result）は各手順の入力・出力を含みうるため、実行の記録に残さない
    acts["失敗した手順"] = secure({"runAfter": {}, "type": "Query", "inputs": {"from": "@result('試行')", "where": "@equals(item()?['status'], 'Failed')"}})
    acts["失敗の要約"] = secure({"runAfter": after("失敗した手順"), "type": "Select", "inputs": {"from": "@body('失敗した手順')", "select": "@concat(item()?['name'], ':', coalesce(item()?['code'], ''))"}})
    prev = "失敗の要約"
    for name, act in (extra_first or []):
        act["runAfter"] = {prev: list(ANY_END)}
        acts[name] = act
        prev = name
    acts["処理ログに記録"] = sp_http("POST", "_api/web/lists/getbytitle('処理ログ')/items", body={
        "Title": flow_label,
        "Target": target_expr,
        "Result": "失敗",
        "Detail": "@{take(join(body('失敗の要約'), ', '), 1000)}",
        "RunLink": {"Url": "@{concat('https://make.powerautomate.com/environments/', workflow()?['tags']?['environmentName'], '/flows/', workflow()?['name'], '/runs/', workflow()?['run']?['name'])}", "Description": "実行の記録"},
    }, run_after={prev: list(ANY_END)})
    acts["管理者に知らせる"] = cond(
        {"not": {"equals": ["@variables('管理者メール')", ""]}},
        {"管理者へのメール": outlook_mail("@variables('管理者メール')", f"【備品貸出】処理に失敗しました（{flow_label}）",
                                       f"<p>{flow_label}で、処理に失敗しました。処理ログの最新の行と、実行の記録を確かめてください。</p>")},
        {},
        run_after={"処理ログに記録": list(ANY_END)},
    )
    return {"type": "Scope", "runAfter": {"試行": ["Failed", "TimedOut"]}, "actions": acts}


# ---- FLW-002 承認 ----
def flw002():
    loan_id = "triggerBody()?['ID']"
    rec = "first(body('記録の読み直し')?['value'])"
    rec2 = "first(body('結果の前の読み直し')?['value'])"
    approval = "body('承認')"
    responder = f"toLower(coalesce(first({approval}?['responses'])?['responder']?['email'], ''))"
    comment = f"coalesce(first({approval}?['responses'])?['comments'], '')"
    requester_mail = f"{rec}?['Requester']?['EMail']"
    start_jst = f"formatDateTime(convertFromUtc({rec}?['StartDate'], '{TZ}'), 'yyyy-MM-dd')"
    due_jst = f"formatDateTime(convertFromUtc({rec}?['DueDate'], '{TZ}'), 'yyyy-MM-dd')"
    equip = f"{rec}?['Equipment']?['Title']"
    period = f"concat({start_jst}, '〜', {due_jst})"
    select = ("$select=Id,LoanStatus,StartDate,DueDate,Purpose,ApprovalStartedAt,RequesterId,Equipment/Title,Requester/EMail,Requester/Title"
              "&$expand=Equipment,Requester")

    def reservation_update(prefix, status=None):
        """予約表の行を、状態の変更（status あり）または削除（status なし）で、貸出記録とそろえる"""
        rows = f"{prefix}の予約表"
        if status:
            inner = {f"{prefix}の予約の更新": sp_http("POST", f"_api/web/lists/getbytitle('予約表')/items(@{{items('{prefix}の予約表をそろえる')?['Id']}})",
                                                  body={"LoanStatus": status}, headers={"IF-MATCH": "*", "X-HTTP-Method": "MERGE"})}
        else:
            inner = {f"{prefix}の予約の削除": sp_http("POST", f"_api/web/lists/getbytitle('予約表')/items(@{{items('{prefix}の予約表をそろえる')?['Id']}})",
                                                  headers={"IF-MATCH": "*", "X-HTTP-Method": "DELETE"})}
        return [
            (rows, sp_http("GET", f"_api/web/lists/getbytitle('予約表')/items?$select=Id&$filter=Title eq '@{{{loan_id}}}'")),
            (f"{prefix}の予約表をそろえる", {"type": "Foreach", "foreach": f"@body('{rows}')?['value']", "actions": inner}),
        ]

    def mail_requester(subject, html):
        return outlook_mail(f"@{{{requester_mail}}}", subject, html)

    # 自己承認の判定は、SharePoint のユーザー ID で比べる（メールの文字の違いに左右されない）
    self_ng = (f"and(equals(int(body('応答者の確認')?['Id']), int({rec}?['RequesterId'])), "
               f"not(equals({SETTING}?['自己承認を許可'], 'はい')))")

    approved = chain(
        ("承認後の状態", compose(f"@if(lessOrEquals({start_jst}, {TODAY}), '貸出中', '予約済み')")),
        ("承認を記録", secure(sp_merge("貸出記録", loan_id, f"{rec2}?['odata.etag']", {
            "LoanStatus": "@{outputs('承認後の状態')}",
            "ApproverId": "@{body('応答者の確認')?['Id']}",
            "ApprovalComment": f"@{{take({comment}, 500)}}",
            "ApprovedAt": "@{utcNow()}",
        }))),
        *reservation_update("承認", "@{outputs('承認後の状態')}"),
        ("承認のメール", mail_requester(
            f"@{{concat('【備品貸出】承認されました：', {equip}, ' ', {period})}}",
            f"<p>貸出の申請が承認されました。</p><p>備品：@{{{equip}}}<br>期間：@{{{period}}}<br>状態：@{{outputs('承認後の状態')}}</p>"
            f"<p>承認者のコメント：@{{{html_safe(comment)}}}</p>")),
        ("承認の通知", teams_card("@{concat('承認：', " + equip + ")}", [("期間", period), ("状態", "outputs('承認後の状態')")])),
    )

    def rejected(prefix, status, note_expr, subject_word, body_text, with_approver=False):
        fields = {"LoanStatus": status, "ApprovalComment": f"@{{take({note_expr}, 500)}}", "ApprovedAt": "@{utcNow()}"}
        if with_approver:  # 応答者を確定できた分岐の中でだけ、承認者（却下した人）を記録する
            fields["ApproverId"] = "@{body('応答者の確認')?['Id']}"
        return chain(
            (f"{prefix}を記録", secure(sp_merge("貸出記録", loan_id, f"{rec2}?['odata.etag']", fields))),
            *reservation_update(prefix),
            (f"{prefix}のメール", mail_requester(
                f"@{{concat('【備品貸出】{subject_word}：', {equip}, ' ', {period})}}",
                f"<p>{body_text}</p><p>備品：@{{{equip}}}<br>期間：@{{{period}}}</p><p>コメント：@{{{html_safe(note_expr)}}}</p>")),
        )

    after_approval = chain(
        ("結果の前の読み直し", sp_http("GET", f"_api/web/lists/getbytitle('貸出記録')/items?$select=Id,LoanStatus&$filter=Id eq @{{{loan_id}}}", accept=JSON_MINIMAL)),
        # 取消などで状態が変わっていたら、承認の結果は書かない
        ("まだ承認待ち", cond(
            {"equals": [f"@{rec2}?['LoanStatus']", "承認待ち"]},
            # 応答者が分からないときは、ユーザーの確定も判定もせずに却下にする（飛ばした手順を式から参照しないため）
            {"応答者あり": cond(
                {"not": {"equals": [f"@{responder}", ""]}},
                chain(
                    ("応答者の確認", secure(sp_http("POST", "_api/web/ensureuser", body={"logonName": f"@{{concat('i:0#.f|membership|', {responder})}}"}))),
                    ("結果の種類", {
                        "type": "Switch",
                        "expression": f"@if(equals({approval}?['outcome'], 'Approve'), if({self_ng}, '自己承認', '承認'), '却下')",
                        "cases": {
                            "承認": {"case": "承認", "actions": approved},
                            "自己承認": {"case": "自己承認", "actions": rejected("自己承認", "却下", "'申請者本人による承認は受け付けていません'", "却下されました",
                                                                            "申請者本人による承認は受け付けていないため、却下としました。別の承認者に依頼してください。", with_approver=True)},
                        },
                        "default": {"actions": rejected("却下", "却下", comment, "却下されました", "貸出の申請が却下されました。", with_approver=True)},
                    }),
                ),
                rejected("応答者不明", "却下", "'承認した人を確認できませんでした'", "却下されました",
                         "承認した人を確認できなかったため、却下としました。管理者に連絡してください。"),
            )},
            {},
        )),
    )

    expired = chain(
        ("期限後の読み直し", sp_http("GET", f"_api/web/lists/getbytitle('貸出記録')/items?$select=Id,LoanStatus&$filter=Id eq @{{{loan_id}}}", accept=JSON_MINIMAL)),
        ("期限時もまだ承認待ち", cond(
            {"equals": ["@first(body('期限後の読み直し')?['value'])?['LoanStatus']", "承認待ち"]},
            chain(
                ("期限切れを記録", sp_merge("貸出記録", loan_id, "first(body('期限後の読み直し')?['value'])?['odata.etag']",
                                     {"LoanStatus": "期限切れ", "ProcessNote": "承認の期限までに承認されなかったため、申請は期限切れになりました。必要な場合は、もう一度申請してください。"})),
                *reservation_update("期限切れ"),
                ("期限切れのメール", mail_requester(
                    f"@{{concat('【備品貸出】申請が期限切れになりました：', {equip}, ' ', {period})}}",
                    f"<p>承認の期限までに承認されなかったため、申請は期限切れになりました。必要な場合は、もう一度申請してください。</p><p>備品：@{{{equip}}}<br>期間：@{{{period}}}</p>")),
            ),
            {},
        )),
    )

    try_actions = chain(
        *settings_steps(),
        ("記録の読み直し", secure(sp_http("GET", f"_api/web/lists/getbytitle('貸出記録')/items?{select}&$filter=Id eq @{{{loan_id}}}", accept=JSON_MINIMAL))),
        # トリガーの後に状態が変わっていないか、承認依頼がまだ出ていないかを確かめる（二重の承認依頼を防ぐ）
        ("承認を始めてよい", cond(
            {"and": [
                {"equals": [f"@{rec}?['LoanStatus']", "承認待ち"]},
                {"equals": [f"@empty({rec}?['ApprovalStartedAt'])", True]},
            ]},
            chain(
                ("承認開始の記録", sp_merge("貸出記録", loan_id, f"{rec}?['odata.etag']", {"ApprovalStartedAt": "@{utcNow()}"})),
                ("承認", secure({
                    "type": "OpenApiConnectionWebhook",
                    "limit": {"timeout": "P7D"},  # 設定の「承認期限日数」の既定値と同じ。変える場合はここも変える
                    "inputs": {
                        "host": {"connectionName": "shared_approvals", "operationId": "StartAndWaitForAnApproval", "apiId": "/providers/Microsoft.PowerApps/apis/shared_approvals"},
                        "parameters": {
                            "approvalType": "Basic",
                            "WebhookApprovalCreationInput/title": f"@{{concat('【備品貸出】', {equip}, ' ', {period})}}",
                            "WebhookApprovalCreationInput/assignedTo": f"@{{{SETTING}?['承認者メール']}}",
                            "WebhookApprovalCreationInput/details": (
                                f"@{{concat('申請者：', {md_safe(f'{rec}?[' + repr('Requester') + ']?[' + repr('Title') + ']', 50)}, decodeUriComponent('%0A%0A'), "
                                f"'期間：', {period}, decodeUriComponent('%0A%0A'), '利用目的：', {md_safe(f'{rec}?[' + repr('Purpose') + ']')})}}"
                            ),
                            "WebhookApprovalCreationInput/enableNotifications": True,
                            "WebhookApprovalCreationInput/enableReassignment": False,
                        },
                        "authentication": "@parameters('$authentication')",
                    },
                })),
                # 承認の完了と期限切れを、1つの手順の中で分ける（期限切れを「失敗」として扱わないため）
                ("承認の結果の処理", cond(
                    {"equals": ["@actions('承認')?['status']", "TimedOut"]},
                    expired,
                    after_approval,
                    run_after={"承認": ["Succeeded", "TimedOut"]},
                )),
            ),
            {},
        )),
    )
    actions = chain(
        ("管理者メール", init_var("管理者メール", "string", "")),
        ("今日", compose(f"@formatDateTime(convertFromUtc(utcNow(), '{TZ}'), 'yyyy-MM-dd')")),
    )
    actions["試行"] = {"type": "Scope", "runAfter": after("今日"), "actions": try_actions}
    actions["失敗時"] = catch_block("FLW-002 承認")
    trigger = {
        "type": "OpenApiConnection",
        "runtimeConfiguration": {"secureData": {"properties": ["outputs"]}},  # 利用目的・申請者を実行の記録に残さない
        "recurrence": {"interval": 1, "frequency": "Minute"},
        "splitOn": "@triggerOutputs()?['body/value']",
        # 承認待ちで、承認依頼がまだ出ていない記録だけで動く（自分が書いた変更で何度も動かないように）
        "conditions": [{"expression": "@and(equals(triggerBody()?['LoanStatus']?['Value'], '承認待ち'), empty(triggerBody()?['ApprovalStartedAt']))"}],
        "inputs": {
            "host": {"connectionName": "shared_sharepointonline", "operationId": "GetOnUpdatedItems", "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"},
            "parameters": {"dataset": SITE, "table": "{{LIST_LOANS_ID}}"},
            "authentication": "@parameters('$authentication')",
        },
    }
    return {
        "displayName": "FLW-002 承認",
        "description": "貸出記録が承認待ちになったら承認を依頼し、結果を記録して申請者に知らせる",
        "connectors": ["shared_sharepointonline", "shared_approvals", "shared_office365", "shared_teams"],
        "definition": definition(trigger, actions),
    }


# ---- FLW-003 期限の通知・督促（毎朝） ----
def flw003():
    today_utc = jst_midnight_utc(TODAY)
    tomorrow = f"formatDateTime(addDays({TODAY}, 1), 'yyyy-MM-dd')"
    jst = lambda e: f"formatDateTime(convertFromUtc({e}, '{TZ}'), 'yyyy-MM-dd')"
    select = ("$select=Id,LoanStatus,StartDate,DueDate,ReminderCount,LastNotifiedDate,ApprovalStartedAt,Created,Equipment/Title,Requester/EMail,Requester/Title"
              "&$expand=Equipment,Requester&$top=500")
    it = "item()"

    def loop(name, source, item_actions):
        return (name, {"type": "Foreach", "foreach": f"@body('{source}')", "runtimeConfiguration": {"concurrency": {"repetitions": 1}}, "actions": item_actions})

    EPOCH = "'1900-01-01T00:00:00Z'"

    def notified_today(e):
        # 空の日時でも変換が失敗しないよう、既定の日時に置き換えてから比べる（and() の評価の順序に頼らない）
        return f"equals({jst(f'coalesce({e}?[' + repr('LastNotifiedDate') + f'], {EPOCH})')}, {TODAY})"

    def reservation_sync(prefix, cur, status=None):
        rows = f"{prefix}の予約表"
        if status:
            inner = {f"{prefix}の予約の更新": sp_http("POST", f"_api/web/lists/getbytitle('予約表')/items(@{{items('{prefix}の予約表をそろえる')?['Id']}})",
                                                  body={"LoanStatus": status}, headers={"IF-MATCH": "*", "X-HTTP-Method": "MERGE"})}
        else:
            inner = {f"{prefix}の予約の削除": sp_http("POST", f"_api/web/lists/getbytitle('予約表')/items(@{{items('{prefix}の予約表をそろえる')?['Id']}})",
                                                  headers={"IF-MATCH": "*", "X-HTTP-Method": "DELETE"})}
        return [
            (rows, sp_http("GET", f"_api/web/lists/getbytitle('予約表')/items?$select=Id&$filter=Title eq '@{{{cur}?['Id']}}'")),
            (f"{prefix}の予約表をそろえる", {"type": "Foreach", "foreach": f"@body('{rows}')?['value']", "actions": inner}),
        ]

    # A: 開始日になった予約を貸出中にする
    cur_a = "items('開始日の記録ごと')"
    part_a = [
        ("予約済みの取得", sp_http("GET", f"_api/web/lists/getbytitle('貸出記録')/items?{select}&$filter=LoanStatus eq '予約済み' and StartDate le datetime'@{{{today_utc}}}'", accept=JSON_MINIMAL)),
        ("開始日の記録", {"type": "Query", "inputs": {"from": "@body('予約済みの取得')?['value']", "where": "@true"}}),
        loop("開始日の記録ごと", "開始日の記録", chain(
            ("貸出中にする", sp_merge("貸出記録", f"{cur_a}?['Id']", f"{cur_a}?['odata.etag']", {"LoanStatus": "貸出中"})),
            *reservation_sync("開始", cur_a, "貸出中"),
        )),
    ]

    # B・C: 貸出中の記録を読み、翌日が期限のもの（お知らせ）と、期限を過ぎたもの（督促）に分ける
    lent = "body('貸出中の取得')?['value']"
    cur_b = "items('前日のお知らせごと')"
    cur_c = "items('督促ごと')"
    limit = f"int(if(empty({SETTING}?['督促上限回数']), '5', {SETTING}?['督促上限回数']))"
    lent_c = "body('督促対象の取得')?['value']"
    part_b = [
        ("貸出中の取得", secure(sp_http("GET", f"_api/web/lists/getbytitle('貸出記録')/items?{select}&$filter=LoanStatus eq '貸出中'", accept=JSON_MINIMAL))),
        ("期限が翌日", {"type": "Query", "inputs": {"from": f"@{lent}", "where": f"@and(equals({jst(it + '?[' + repr('DueDate') + ']')}, {tomorrow}), not({notified_today(it)}))"}}),
        loop("前日のお知らせごと", "期限が翌日", chain(
            ("お知らせのメール", outlook_mail(f"@{{{cur_b}?['Requester']?['EMail']}}",
                                       f"@{{concat('【備品貸出】返却期限は明日です：', {cur_b}?['Equipment']?['Title'])}}",
                                       f"<p>貸出中の備品の返却期限は明日です。期限までに返却し、アプリで「返却した」を押してください。</p>"
                                       f"<p>備品：@{{{cur_b}?['Equipment']?['Title']}}<br>返却予定日：@{{{jst(cur_b + '?[' + repr('DueDate') + ']')}}}</p>")),
            ("お知らせ済みにする", sp_merge("貸出記録", f"{cur_b}?['Id']", f"{cur_b}?['odata.etag']", {"LastNotifiedDate": f"@{{{today_utc}}}"})),
        )),
    ]
    # 督促はお知らせと別の範囲で、読み直してから行う（お知らせでつまずいても督促は動く）
    part_c = [
        ("督促対象の取得", secure(sp_http("GET", f"_api/web/lists/getbytitle('貸出記録')/items?{select}&$filter=LoanStatus eq '貸出中'", accept=JSON_MINIMAL))),
        ("期限切れ", {"type": "Query", "inputs": {"from": f"@{lent_c}", "where": f"@and(less({jst(it + '?[' + repr('DueDate') + ']')}, {TODAY}), not({notified_today(it)}), less(int(coalesce({it}?['ReminderCount'], 0)), {limit}))"}}),
        loop("督促ごと", "期限切れ", chain(
            ("督促の回数", compose(f"@add(int(coalesce({cur_c}?['ReminderCount'], 0)), 1)")),
            ("督促のメール", outlook_mail(f"@{{{cur_c}?['Requester']?['EMail']}}",
                                     f"@{{concat('【備品貸出】返却期限を過ぎています（督促', string(outputs('督促の回数')), '回目）：', {cur_c}?['Equipment']?['Title'])}}",
                                     f"<p>貸出中の備品の返却期限を過ぎています。至急返却し、アプリで「返却した」を押してください。</p>"
                                     f"<p>備品：@{{{cur_c}?['Equipment']?['Title']}}<br>返却予定日：@{{{jst(cur_c + '?[' + repr('DueDate') + ']')}}}</p>")),
            ("督促を記録", sp_merge("貸出記録", f"{cur_c}?['Id']", f"{cur_c}?['odata.etag']", {"ReminderCount": "@outputs('督促の回数')", "LastNotifiedDate": f"@{{{today_utc}}}"})),
            ("上限に達した", cond(
                {"greaterOrEquals": ["@outputs('督促の回数')", f"@{limit}"]},
                chain(
                    # 管理者のメールアドレスが設定されていないときは、メールは送らず、Teams の通知だけにする（Teams には氏名を載せない）
                    ("上限を管理者に知らせる", cond(
                        {"not": {"equals": ["@variables('管理者メール')", ""]}},
                        {"上限の管理者メール": outlook_mail("@variables('管理者メール')",
                                                    f"@{{concat('【備品貸出】督促が上限に達しました：', {cur_c}?['Equipment']?['Title'])}}",
                                                    f"<p>督促が上限の回数に達しました。これ以上は自動で督促しません。申請者に直接連絡してください。</p>"
                                                    f"<p>備品：@{{{cur_c}?['Equipment']?['Title']}}<br>申請者：@{{{html_safe(cur_c + '?[' + repr('Requester') + ']?[' + repr('Title') + ']', 50)}}}<br>"
                                                    f"返却予定日：@{{{jst(cur_c + '?[' + repr('DueDate') + ']')}}}</p>")},
                        {},
                    )),
                    ("上限の通知", teams_card(f"@{{concat('督促が上限に達しました：', {cur_c}?['Equipment']?['Title'])}}",
                                         [("貸出記録ID", f"string({cur_c}?['Id'])"), ("返却予定日", jst(cur_c + "?['DueDate']")), ("対応", "'貸出記録リスト（所有者だけが読める）で、この ID の申請者に連絡してください'")])),
                ),
                {},
            )),
        )),
    ]

    # D: 止まった記録の後始末（承認の途中で止まった承認待ち・作りかけの確認中）
    cur_d = "items('止まった承認待ちごと')"
    cur_e = "items('作りかけごと')"
    part_d1 = [
        ("承認待ちの取得", sp_http("GET", f"_api/web/lists/getbytitle('貸出記録')/items?{select}&$filter=LoanStatus eq '承認待ち'", accept=JSON_MINIMAL)),
        # 承認を始めてから8日（承認の期限 7日＋1日）過ぎたもの、または承認を始めないまま1日過ぎたもの
        ("止まった承認待ち", {"type": "Query", "inputs": {"from": "@body('承認待ちの取得')?['value']",
                                                         # 日時は ticks（数値）で比べる（文字列の比較は書式の違いでずれる）
                                                         # 空の日時は既定の日時に置き換えてから ticks にする（評価の順序に頼らない）
                                                         "where": "@or(and(not(empty(item()?['ApprovalStartedAt'])), less(ticks(coalesce(item()?['ApprovalStartedAt'], '2999-01-01T00:00:00Z')), ticks(addDays(utcNow(), -8)))), "
                                                                  "and(empty(item()?['ApprovalStartedAt']), less(ticks(item()?['Created']), ticks(addDays(utcNow(), -1)))))"}}),
        loop("止まった承認待ちごと", "止まった承認待ち", chain(
            ("止まった記録を期限切れにする", sp_merge("貸出記録", f"{cur_d}?['Id']", f"{cur_d}?['odata.etag']",
                                            {"LoanStatus": "期限切れ", "ProcessNote": "承認の期限までに承認されなかったため、申請は期限切れになりました。必要な場合は、もう一度申請してください。"})),
            *reservation_sync("止まった", cur_d),
            ("止まった記録のメール", outlook_mail(f"@{{{cur_d}?['Requester']?['EMail']}}",
                                         f"@{{concat('【備品貸出】申請が期限切れになりました：', {cur_d}?['Equipment']?['Title'])}}",
                                         "<p>承認の期限までに承認されなかったため、申請は期限切れになりました。必要な場合は、もう一度申請してください。</p>")),
        )),
    ]
    part_d2 = [
        ("確認中の取得", sp_http("GET", f"_api/web/lists/getbytitle('貸出記録')/items?{select}&$filter=LoanStatus eq '確認中'", accept=JSON_MINIMAL)),
        ("作りかけ", {"type": "Query", "inputs": {"from": "@body('確認中の取得')?['value']", "where": "@less(ticks(item()?['Created']), ticks(addHours(utcNow(), -1)))"}}),
        loop("作りかけごと", "作りかけ", chain(
            ("作りかけをエラーにする", sp_merge("貸出記録", f"{cur_e}?['Id']", f"{cur_e}?['odata.etag']",
                                          {"LoanStatus": "エラー", "ProcessNote": "申請の処理が完了しなかったため、申請を受け付けられませんでした。お手数ですが、もう一度申請してください。"})),
            *reservation_sync("作りかけ", cur_e),
        )),
    ]

    # 5つの処理を、それぞれ独立した範囲（Scope）にし、前の範囲が失敗しても次の範囲に進める
    # （どれか1つでつまずいても、残りは動く。失敗は「失敗時」で処理ログに残る）
    def secure_reads(part):
        for name, act in part:
            if name.endswith("の取得") and act.get("type") == "OpenApiConnection":
                secure(act)
        return part

    try_actions = chain(*settings_steps())
    prev = "管理者メールの設定"
    scopes = (("開始日の切り替え", part_a), ("前日のお知らせ", part_b), ("督促", part_c),
              ("止まった承認待ちの後始末", part_d1), ("作りかけの後始末", part_d2))
    for scope_name, part in scopes:
        try_actions[scope_name] = {"type": "Scope", "runAfter": {prev: list(ANY_END)}, "actions": chain(*secure_reads(part))}
        prev = scope_name
    actions = chain(
        ("管理者メール", init_var("管理者メール", "string", "")),
        ("今日", compose(f"@formatDateTime(convertFromUtc(utcNow(), '{TZ}'), 'yyyy-MM-dd')")),
    )
    actions["試行"] = {"type": "Scope", "runAfter": after("今日"), "actions": try_actions}
    actions["失敗時"] = catch_block("FLW-003 期限の通知・督促")
    # catch_block の Target は triggerBody()?['ID'] を使うが、この流れでは使えないので固定の文字にする
    actions["失敗時"]["actions"]["処理ログに記録"] = sp_http("POST", "_api/web/lists/getbytitle('処理ログ')/items", body={
        "Title": "FLW-003 期限の通知・督促", "Target": "毎朝の処理", "Result": "失敗",
        "Detail": "@{take(join(body('失敗の要約'), ', '), 1000)}",
        "RunLink": {"Url": "@{concat('https://make.powerautomate.com/environments/', workflow()?['tags']?['environmentName'], '/flows/', workflow()?['name'], '/runs/', workflow()?['run']?['name'])}", "Description": "実行の記録"},
    }, run_after={"失敗の要約": list(ANY_END)})
    # 範囲の状態は最後の手順で決まるため、「後始末」が成功すると、前の範囲が失敗していても「試行」は成功になる。
    # そこで「失敗時」は試行の後に必ず動かし、中のどれかが失敗していたときだけ、記録と通知をする
    inner = ["設定の取得"] + [name for name, _ in scopes]
    failed_any = {"or": [{"equals": [f"@actions('{n}')?['status']", "Failed"]} for n in inner]
                  + [{"equals": ["@actions('試行')?['status']", "Failed"]}]}
    catch = actions.pop("失敗時")
    actions["失敗時"] = {"type": "Scope", "runAfter": {"試行": list(ANY_END)}, "actions": {
        "失敗があった": cond(failed_any, catch["actions"], {}),
    }}
    trigger = {"type": "Recurrence", "recurrence": {"frequency": "Day", "interval": 1, "timeZone": TZ, "schedule": {"hours": ["9"], "minutes": [0]}}}
    return {
        "displayName": "FLW-003 期限の通知・督促",
        "description": "毎朝9時。開始日になった予約を貸出中にし、返却期限の前日に知らせ、期限を過ぎたら督促する。止まった申請を後始末する",
        "connectors": ["shared_sharepointonline", "shared_office365", "shared_teams"],
        "definition": definition(trigger, actions),
    }


def planner(op, params, run_after=None):
    """Planner の操作（CreateTask_V3 / UpdateTask_V2）"""
    return {"runAfter": run_after or {}, "type": "OpenApiConnection", "inputs": {
        "host": {"connectionName": "shared_planner", "operationId": op, "apiId": "/providers/Microsoft.PowerApps/apis/shared_planner"},
        "parameters": params, "authentication": "@parameters('$authentication')"}}


def base_actions():
    return chain(
        ("管理者メール", init_var("管理者メール", "string", "")),
        ("今日", compose(f"@formatDateTime(convertFromUtc(utcNow(), '{TZ}'), 'yyyy-MM-dd')")),
    )


# ---- FLW-004 故障の受付（Forms → 故障報告・Planner・Teams） ----
def flw004():
    resp = "body('回答の詳細')"
    code = "outputs('管理番号')"
    detail = f"take(coalesce({resp}?['{{{{Q_DETAIL}}}}'], ''), 1000)"
    equip = "body('備品の取得')?['value']"
    fault_id = "body('故障報告の作成')?['Id']"
    try_actions = chain(
        *settings_steps(),
        ("回答の詳細", secure({
            "type": "OpenApiConnection",
            "inputs": {
                "host": {"connectionName": "shared_microsoftforms", "operationId": "GetFormResponseById", "apiId": "/providers/Microsoft.PowerApps/apis/shared_microsoftforms"},
                "parameters": {"form_id": "{{FORM_ID}}", "response_id": "@triggerOutputs()?['body/resourceData/responseId']"},
                "authentication": "@parameters('$authentication')",
            },
        })),
        # 管理番号は、前後の空白を除いて大文字にし、20文字で切る（フィルターでは ' を重ねる）
        ("管理番号", compose(f"@toUpper(take(trim(coalesce({resp}?['{{{{Q_CODE}}}}'], '')), 20))")),
        # URL に入れる前に、' を重ねてからエンコードする（# % & を含む入力で要求が壊れない。%27 の形でも ' にならない）
        ("備品の取得", sp_http("GET", f"_api/web/lists/getbytitle('備品台帳')/items?$select=Id,Title&$filter=Title eq '@{{encodeUriComponent({odata_str(code)})}}'")),
        # 回答者（組織内の人だけ回答でき、名前を記録する設定のフォーム）。確定できなくても報告は保存する
        ("報告者の確認", secure(sp_http("POST", "_api/web/ensureuser", body={"logonName": f"@{{concat('i:0#.f|membership|', toLower(coalesce({resp}?['responder'], '')))}}"},
                                    run_after={"備品の取得": list(ANY_END)}))),
        # 備品台帳は自動では変えない（いたずらで貸出を止められないように）。管理者が確かめてから「修理中」にする
        # 備品が見つからない・報告者を確定できない場合も、その列を空にして報告は必ず保存する
        ("故障報告の作成", secure(sp_http("POST", "_api/web/lists/getbytitle('故障報告')/items", run_after={"報告者の確認": list(ANY_END)}, body={
            "Title": f"@{{concat({code}, ' ', {TODAY})}}",
            "EquipmentId": f"@if(and(equals(actions('備品の取得')?['status'], 'Succeeded'), equals(length(coalesce({equip}, json('[]'))), 1)), first({equip})?['Id'], null)",
            "ReporterId": "@if(equals(actions('報告者の確認')?['status'], 'Succeeded'), body('報告者の確認')?['Id'], null)",
            "Detail": f"@{{{detail}}}",
            "FaultStatus": "確認待ち",
        }))),
        ("確認のタスク", planner("CreateTask_V3", {
            "body/groupId": "{{TEAM_ID}}", "body/planId": "{{PLAN_ID}}", "body/bucketId": "{{BUCKET_ID}}",
            "body/title": f"@{{concat('【確認待ち】', {code}, ' 故障報告ID ', string({fault_id}))}}",
        })),
        ("タスクIDを記録", sp_merge("故障報告", fault_id, "'*'", {"PlannerTaskId": "@{body('確認のタスク')?['id']}"})),
        # チャンネルはメンバー全員が読むため、内容と報告者は載せない（故障報告リストは所有者だけが読める）
        ("故障の通知", teams_card(f"@{{concat('故障の報告を受け付けました：', {md_safe(code, 20)})}}",
                                [("対応", "'故障報告リストと Planner「備品の修理対応」で確認してください'")])),
    )
    actions = base_actions()
    actions["試行"] = {"type": "Scope", "runAfter": after("今日"), "actions": try_actions}
    actions["失敗時"] = catch_block("FLW-004 故障の受付", target_expr="Forms の回答")
    trigger = {
        "type": "OpenApiConnectionWebhook",
        "runtimeConfiguration": {"secureData": {"properties": ["outputs"]}},
        "splitOn": "@triggerOutputs()?['body/value']",
        "inputs": {
            "host": {"connectionName": "shared_microsoftforms", "operationId": "CreateFormWebhook", "apiId": "/providers/Microsoft.PowerApps/apis/shared_microsoftforms"},
            "parameters": {"form_id": "{{FORM_ID}}"},
            "authentication": "@parameters('$authentication')",
        },
    }
    return {
        "displayName": "FLW-004 故障の受付",
        "description": "Forms の故障・不具合の報告を、故障報告リストと Planner の確認のタスクにし、Teams に知らせる",
        "connectors": ["shared_sharepointonline", "shared_microsoftforms", "shared_planner", "shared_teams", "shared_office365"],
        "definition": definition(trigger, actions),
    }


# ---- FLW-006 故障の対応状況の反映（故障報告 → 備品台帳・Planner） ----
def flw006():
    """管理者が故障報告の「対応状況」を変えたら、備品台帳と Planner のタスクをそろえる。
    故障報告を編集できるのは所有者（管理者）だけなので、変更した人の確認は不要（Planner の完了のトリガーは完了者を返さないため、こちらを入口にした）"""
    fault = "first(body('故障報告の読み直し')?['value'])"
    eq = "first(body('備品の読み直し')?['value'])"

    def set_equip(from_status, to_status, prefix):
        return [
            (f"{prefix}備品の読み直し", sp_http("GET", f"_api/web/lists/getbytitle('備品台帳')/items?$select=Id,EquipStatus&$filter=Id eq @{{int(coalesce({fault}?['EquipmentId'], 0))}}", accept=JSON_MINIMAL)),
            (f"{prefix}備品を変える", cond(
                {"equals": [f"@first(body('{prefix}備品の読み直し')?['value'])?['EquipStatus']", from_status]},
                {f"{prefix}備品の更新": sp_merge("備品台帳", f"first(body('{prefix}備品の読み直し')?['value'])?['Id']",
                                             f"first(body('{prefix}備品の読み直し')?['value'])?['odata.etag']", {"EquipStatus": to_status})},
                {},
            )),
        ]

    def task(percent, prefix):
        return (f"{prefix}タスクの更新", cond(
            {"not": {"equals": [f"@coalesce({fault}?['PlannerTaskId'], '')", ""]}},
            {f"{prefix}タスクを{percent}": planner("UpdateTask_V2", {"id": f"@{fault}?['PlannerTaskId']", "body/percentComplete": percent})},
            {},
        ))

    equip_id = f"int(coalesce({fault}?['EquipmentId'], 0))"
    open_count = "length(body('修理中の報告')?['value'])"
    equip_status = "first(body('備品の読み直し')?['value'])?['EquipStatus']"
    try_actions = chain(
        ("故障報告の読み直し", sp_http("GET", "_api/web/lists/getbytitle('故障報告')/items?$select=Id,FaultStatus,EquipmentId,PlannerTaskId&$filter=Id eq @{triggerBody()?['ID']}", accept=JSON_MINIMAL)),
        # Planner のタスクを、この報告の対応状況に合わせる
        ("対応状況", {
            "type": "Switch",
            "expression": f"@{fault}?['FaultStatus']",
            "cases": {
                "修理中": {"case": "修理中", "actions": chain(task("In Progress", "修理中の"))},
                "完了": {"case": "完了", "actions": chain(task("Completed", "完了の"))},
                "対象外": {"case": "対象外", "actions": chain(task("Completed", "対象外の"))},
            },
            "default": {"actions": {}},
        }),
        # 備品の状態は、この報告だけでなく「その備品の修理中の故障報告が残っているか」で決める
        # （ほかの報告で修理中にしている備品を戻さない。何度動いても同じ結果になる）
        ("備品あり", cond(
            {"greater": [f"@{equip_id}", 0]},
            chain(
                ("修理中の報告", sp_http("GET", f"_api/web/lists/getbytitle('故障報告')/items?$select=Id&$filter=EquipmentId eq @{{{equip_id}}} and FaultStatus eq '修理中'")),
                ("備品の読み直し", sp_http("GET", f"_api/web/lists/getbytitle('備品台帳')/items?$select=Id,EquipStatus&$filter=Id eq @{{{equip_id}}}", accept=JSON_MINIMAL)),
                ("備品の状態をそろえる", cond(
                    {"or": [
                        {"and": [{"greater": [f"@{open_count}", 0]}, {"equals": [f"@{equip_status}", "利用可"]}]},
                        {"and": [{"equals": [f"@{open_count}", 0]}, {"equals": [f"@{equip_status}", "修理中"]}]},
                    ]},
                    {"備品の更新": sp_merge("備品台帳", "first(body('備品の読み直し')?['value'])?['Id']", "first(body('備品の読み直し')?['value'])?['odata.etag']",
                                            {"EquipStatus": f"@{{if(greater({open_count}, 0), '修理中', '利用可')}}"})},
                    {},
                )),
            ),
            {},
            run_after={"対応状況": list(ANY_END)},
        )),
    )
    actions = base_actions()
    actions["試行"] = {"type": "Scope", "runAfter": after("今日"), "actions": chain(*settings_steps(), *try_actions.items())}
    actions["失敗時"] = catch_block("FLW-006 故障の対応状況の反映", target_expr="@{concat('故障報告ID=', string(coalesce(triggerBody()?['ID'], 0)))}")
    trigger = {
        "type": "OpenApiConnection",
        "runtimeConfiguration": {"secureData": {"properties": ["outputs"]}},
        "recurrence": {"interval": 1, "frequency": "Minute"},
        "splitOn": "@triggerOutputs()?['body/value']",
        "conditions": [{"expression": "@or(equals(triggerBody()?['FaultStatus']?['Value'], '修理中'), equals(triggerBody()?['FaultStatus']?['Value'], '完了'), equals(triggerBody()?['FaultStatus']?['Value'], '対象外'))"}],
        "inputs": {
            "host": {"connectionName": "shared_sharepointonline", "operationId": "GetOnUpdatedItems", "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"},
            "parameters": {"dataset": SITE, "table": "{{LIST_FAULTS_ID}}"},
            "authentication": "@parameters('$authentication')",
        },
    }
    return {
        "displayName": "FLW-006 故障の対応状況の反映",
        "description": "管理者が故障報告の対応状況を変えたら、備品台帳の状態と Planner のタスクをそろえる",
        "connectors": ["shared_sharepointonline", "shared_planner", "shared_office365"],
        "definition": definition(trigger, actions),
    }


# ---- FLW-005 月次の集計 ----
def flw005():
    cur_first = f"formatDateTime(startOfMonth(convertFromUtc(utcNow(), '{TZ}')), 'yyyy-MM-dd')"
    prev_first = f"formatDateTime(addToTime(startOfMonth(convertFromUtc(utcNow(), '{TZ}')), -1, 'Month'), 'yyyy-MM-dd')"
    month = "outputs('対象の月')"
    loans = "body('前月の貸出')?['value']"
    jst = lambda e: f"formatDateTime(convertFromUtc({e}, '{TZ}'), 'yyyy-MM-dd')"

    def csv_safe(e):
        # Excel で数式として動かないよう、= + - @（全角を含む）とタブ・CR・LF で始まる値の先頭に ' を付ける（数式注入の対策）
        heads = ["'='", "'+'", "'-'", "'@'", "'＝'", "'＋'", "'－'", "'＠'",
                 "decodeUriComponent('%09')", "decodeUriComponent('%0D')", "decodeUriComponent('%0A')"]
        return "if(or(" + ", ".join(f"startsWith({e}, {h})" for h in heads) + f"), concat('''', {e}), {e})"

    try_actions = chain(
        *settings_steps(),
        ("対象の月", compose({"from": f"@{prev_first}", "to": f"@{cur_first}", "label": f"@{{substring({prev_first}, 0, 7)}}"})),
        ("前月の貸出", secure(sp_http("GET", "_api/web/lists/getbytitle('貸出記録')/items?$select=Id,LoanStatus,StartDate,DueDate,ReturnedDate,ReminderCount,Equipment/Title&$expand=Equipment&$top=5000"
                                       f"&$filter=StartDate ge datetime'@{{{jst_midnight_utc(month + '?[' + repr('from') + ']')}}}' and StartDate lt datetime'@{{{jst_midnight_utc(month + '?[' + repr('to') + ']')}}}'",
                                       accept=JSON_MINIMAL))),
        ("前月の故障", sp_http("GET", "_api/web/lists/getbytitle('故障報告')/items?$select=Id&$top=5000"
                                f"&$filter=Created ge datetime'@{{{jst_midnight_utc(month + '?[' + repr('from') + ']')}}}' and Created lt datetime'@{{{jst_midnight_utc(month + '?[' + repr('to') + ']')}}}'")),
        # 明細には自由入力（利用目的）と氏名を入れない（管理番号・期間・状態・督促回数だけ）
        ("明細", {"type": "Select", "inputs": {"from": f"@{loans}", "select": {
            "記録ID": "@item()?['Id']",
            "管理番号": "@" + csv_safe("coalesce(item()?['Equipment']?['Title'], '')"),
            "貸出開始日": f"@{jst('item()?[' + repr('StartDate') + ']')}",
            "返却予定日": f"@{jst('item()?[' + repr('DueDate') + ']')}",
            "状態": "@item()?['LoanStatus']",
            "督促回数": "@coalesce(item()?['ReminderCount'], 0)",
        }}}),
        ("CSV", {"type": "Table", "inputs": {"from": "@body('明細')", "format": "CSV"}}),
    )
    # 件数は Query で数える
    extra = chain(
        ("実際の貸出", {"type": "Query", "inputs": {"from": f"@{loans}", "where": "@or(equals(item()?['LoanStatus'], '貸出中'), equals(item()?['LoanStatus'], '返却済み'))"}}),
        ("延滞あり", {"type": "Query", "inputs": {"from": f"@{loans}", "where": "@greater(int(coalesce(item()?['ReminderCount'], 0)), 0)"}}),
        ("ファイルの保存", {
            "type": "OpenApiConnection",
            "runtimeConfiguration": {"contentTransfer": {"transferMode": "Chunked"}},
            "inputs": {
                "host": {"connectionName": "shared_sharepointonline", "operationId": "CreateFile", "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"},
                # Excel で開いたときに文字化けしないよう、先頭に BOM を付ける
                "parameters": {"dataset": SITE, "folderPath": "/MonthlyReports", "name": f"@{{concat('貸出_', {month}?['label'], '.csv')}}",
                               # 明細が0件のときも、列名の行は出す（CSV の表は行が無いと列名も出さないため）
                               "body": "@concat(decodeUriComponent('%EF%BB%BF'), if(empty(body('明細')), concat('記録ID,管理番号,貸出開始日,返却予定日,状態,督促回数', decodeUriComponent('%0D%0A')), body('CSV')))"},
                "authentication": "@parameters('$authentication')",
            },
        }),
        ("集計のメール", cond(
            {"not": {"equals": ["@variables('管理者メール')", ""]}},
            {"管理者への集計": outlook_mail("@variables('管理者メール')", f"@{{concat('【備品貸出】月次の集計（', {month}?['label'], '）')}}",
                                       f"<p>@{{{month}?['label']}} の集計です。</p><ul>"
                                       f"<li>申請の件数：@{{length({loans})}}</li>"
                                       "<li>貸出の件数（貸出中・返却済み）：@{length(body('実際の貸出'))}</li>"
                                       "<li>延滞の件数（督促あり）：@{length(body('延滞あり'))}</li>"
                                       "<li>故障の報告の件数：@{length(body('前月の故障')?['value'])}</li></ul>"
                                       "<p>明細は、チームのサイトの「月次レポート」に CSV で保存しました。</p>")},
            {},
        )),
    )
    last = list(try_actions.keys())[-1]
    first_extra = list(extra.keys())[0]
    extra[first_extra]["runAfter"] = after(last)
    try_actions.update(extra)
    actions = base_actions()
    actions["試行"] = {"type": "Scope", "runAfter": after("今日"), "actions": try_actions}
    actions["失敗時"] = catch_block("FLW-005 月次の集計", target_expr="毎月の処理")
    trigger = {"type": "Recurrence", "recurrence": {"frequency": "Month", "interval": 1, "timeZone": TZ, "schedule": {"monthDays": [1], "hours": ["9"], "minutes": [0]}}}
    return {
        "displayName": "FLW-005 月次の集計",
        "description": "毎月1日9時に前月の申請・貸出・延滞・故障の件数を集計し、明細の CSV を保存して管理者に知らせる",
        "connectors": ["shared_sharepointonline", "shared_office365"],
        "definition": definition(trigger, actions),
    }


def main():
    out = pathlib.Path(__file__).resolve().parent.parent / "flows"
    out.mkdir(exist_ok=True)
    for spec in (flw001(), flw002(), flw003(), flw004(), flw005(), flw006()):
        path = out / (spec["displayName"].split(" ")[0] + ".json")
        path.write_text(json.dumps(spec, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(path)


if __name__ == "__main__":
    main()
