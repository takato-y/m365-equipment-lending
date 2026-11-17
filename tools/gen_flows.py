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


def object_expr(fields):
    """フィールドの辞書から、setProperty を重ねてオブジェクトを作る式を返す。
    値を JSON の文字列に埋め込まないので、利用者が入れた文字（" や改行）でデータの形が崩れない（注入を防ぐ）。"""
    expr = "json('{}')"
    for key, val in fields.items():
        inner = object_expr(val) if isinstance(val, dict) else template_to_expr(val)
        expr = f"setProperty({expr}, {_quote(key)}, {inner})"
    return expr


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
        ("呼び出し元", compose("@toLower(coalesce(triggerOutputs()?['headers']?['x-ms-user-email'], ''))")),
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
    """Office 365 Outlook「メールの送信 (V2)」"""
    return {
        "runAfter": run_after or {},
        "type": "OpenApiConnection",
        "inputs": {
            "host": {"connectionName": "shared_office365", "operationId": "SendEmailV2", "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365"},
            "parameters": {"emailMessage/To": to_expr, "emailMessage/Subject": subject, "emailMessage/Body": body_html, "emailMessage/Importance": "Normal"},
            "authentication": "@parameters('$authentication')",
        },
    }


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
                ("先の申請に譲る", sp_merge("貸出記録", new_id, new_etag, {"LoanStatus": "却下", "ProcessNote": "同じ期間に先の申請があるため受け付けなかった"})),
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
                    ("利用者の確認", sp_http("POST", "_api/web/ensureuser", body={"logonName": f"@{{concat('i:0#.f|membership|', {CALLER})}}"})),
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
    catch_actions["失敗した手順"] = {"runAfter": {}, "type": "Query", "inputs": {"from": "@result('試行')", "where": "@equals(item()?['status'], 'Failed')"}}
    # 手順の名前とエラーの種類だけを取り出す（入力・出力の中身は入れない）
    catch_actions["失敗の要約"] = {"runAfter": after("失敗した手順"), "type": "Select", "inputs": {"from": "@body('失敗した手順')", "select": "@concat(item()?['name'], ':', coalesce(item()?['code'], ''))"}}
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
                                                  {"LoanStatus": "エラー", "ProcessNote": "処理の途中で失敗したため、エラーにした"})),
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


def main():
    out = pathlib.Path(__file__).resolve().parent.parent / "flows"
    out.mkdir(exist_ok=True)
    for spec in (flw001(),):
        path = out / (spec["displayName"].split(" ")[0] + ".json")
        path.write_text(json.dumps(spec, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(path)


if __name__ == "__main__":
    main()
