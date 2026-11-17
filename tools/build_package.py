#!/usr/bin/env python3
"""フローの定義（flows/*.json）から、Power Automate の「パッケージのインポート (レガシ)」で取り込める zip を作る。

ソリューションの書き出し・取り込みは、このテナントでは「Async operations are currently disabled」で使えないため、
パッケージで取り込む（docs/adr を参照）。接続は取り込みの画面で、既にある接続を選んで結び付ける。

使い方:
  python3 tools/build_package.py flows/FLW-001.json -o out/FLW-001.zip
flows/*.json の形:
  {"displayName": "...", "description": "...", "connectors": ["shared_sharepointonline", ...], "definition": { Logic Apps のワークフロー定義 }}
"""
import argparse
import json
import pathlib
import uuid
import zipfile

CONNECTOR_NAMES = {
    "shared_sharepointonline": "SharePoint",
    "shared_office365": "Office 365 Outlook",
    "shared_teams": "Microsoft Teams",
    "shared_approvals": "承認",
    "shared_microsoftforms": "Microsoft Forms",
    "shared_planner": "Planner",
}


def stable_guid(*parts):
    """同じ入力からは毎回同じ ID を作る（zip の中身を再現可能にするため）"""
    return str(uuid.uuid5(uuid.NAMESPACE_URL, "m365-equipment-lending/" + "/".join(parts)))


def fill_placeholders(text, config):
    """定義の中の {{KEY}} を、リポジトリの外に置いた設定ファイルの値で置き換える（サイトの URL などを公開しないため）"""
    for key, value in config.items():
        text = text.replace("{{" + key + "}}", value)
    left = [m for m in __import__("re").findall(r"\{\{([A-Z_]+)\}\}", text)]
    if left:
        raise SystemExit(f"設定ファイルに値がありません: {sorted(set(left))}")
    return text


def build(spec_path, out_path, config_path=None):
    raw = pathlib.Path(spec_path).read_text(encoding="utf-8")
    config = json.loads(pathlib.Path(config_path).read_text(encoding="utf-8")) if config_path else {}
    spec = json.loads(fill_placeholders(raw, config))
    name = spec["displayName"]
    flow_res = stable_guid("flow", name)
    resources = {
        flow_res: {
            "type": "Microsoft.Flow/flows",
            "suggestedCreationType": "New",
            "creationType": "Existing, New, Update",
            "details": {"displayName": name},
            "configurableBy": "User",
            "hierarchy": "Root",
            "dependsOn": [],
        }
    }
    apis_map, conns_map, conn_refs = {}, {}, {}
    for api in spec.get("connectors", []):
        api_res = stable_guid("api", api)
        conn_res = stable_guid("conn", api)
        display = CONNECTOR_NAMES.get(api, api)
        resources[api_res] = {
            "id": f"/providers/Microsoft.PowerApps/apis/{api}",
            "name": api,
            "type": "Microsoft.PowerApps/apis",
            "suggestedCreationType": "Existing",
            "details": {"displayName": display},
            "configurableBy": "System",
            "hierarchy": "Child",
            "dependsOn": [],
        }
        resources[conn_res] = {
            "type": "Microsoft.PowerApps/apis/connections",
            "suggestedCreationType": "Existing",
            "creationType": "Existing",
            "details": {"displayName": display},
            "configurableBy": "User",
            "hierarchy": "Child",
            "dependsOn": [api_res],
        }
        resources[flow_res]["dependsOn"] += [api_res, conn_res]
        apis_map[api] = api_res
        conns_map[api] = conn_res
        conn_refs[api] = {"connectionName": "", "source": "Embedded", "id": f"/providers/Microsoft.PowerApps/apis/{api}", "tier": "NotSpecified"}

    manifest = {
        "schema": "1.0",
        "details": {"displayName": name, "description": spec.get("description", ""), "createdTime": "2026-11-17T00:00:00Z", "packageTelemetryId": stable_guid("pkg", name), "creator": "", "sourceEnvironment": ""},
        "resources": resources,
    }
    definition = {
        "name": stable_guid("flowname", name),
        "id": f"/providers/Microsoft.Flow/flows/{stable_guid('flowname', name)}",
        "type": "Microsoft.Flow/flows",
        "properties": {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
            "displayName": name,
            "definition": spec["definition"],
            "connectionReferences": conn_refs,
            "flowFailureAlertSubscribed": False,
            "isManaged": False,
        },
    }
    out = pathlib.Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    base = f"Microsoft.Flow/flows/{flow_res}/"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        z.writestr("Microsoft.Flow/flows/manifest.json", json.dumps({"packageSchemaVersion": "1.0", "flowAssets": {"assetPaths": [flow_res]}}, indent=2))
        z.writestr(base + "definition.json", json.dumps(definition, ensure_ascii=False, indent=2))
        z.writestr(base + "apisMap.json", json.dumps(apis_map, indent=2))
        z.writestr(base + "connectionsMap.json", json.dumps(conns_map, indent=2))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("-c", "--config", help="{{KEY}} に入れる値の JSON（リポジトリの外に置く）")
    args = ap.parse_args()
    print(build(args.spec, args.out, args.config))


if __name__ == "__main__":
    main()
