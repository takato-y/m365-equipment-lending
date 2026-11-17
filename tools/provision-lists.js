/**
 * provision-lists.js
 * 備品貸出管理のリスト・列・権限・初期データを、SharePoint の REST API で作る。
 *
 * 使い方: 対象のチームサイトの「サイトの設定」ページ（<サイトの URL>/_layouts/15/settings.aspx）を所有者のアカウントで開き、
 *         ブラウザの開発者ツールのコンソールにこのファイルの中身を貼り付けて実行する。
 *         （サイトのホームなどのモダン ページでは _spPageContextInfo が無く動かない）
 *         何度実行してもよい（既にあるリスト・列・初期データの行は作り直さず、足りない物だけを足す。
 *         権限は毎回、設計どおりに付け直す。既にある設定の値は変えない）。
 * 結果:   最後に、作った物・飛ばした物の一覧をコンソールに表示し、オブジェクトとして返す。
 *
 * 内部名は英数字（フローから参照する名前）、表示名は日本語にする。
 * 承認者・管理者のメールアドレスは、実行したユーザーのアドレスを実行時に読んで設定リストに入れる
 * （リポジトリにアドレスを書かないため）。
 */
(async () => {
  const site = location.origin + _spPageContextInfo.webServerRelativeUrl;
  const log = [];
  const say = (msg) => { log.push(msg); console.log('[provision] ' + msg); };

  // ---- REST の共通処理 ----
  async function digest() {
    const r = await fetch(site + '/_api/contextinfo', { method: 'POST', headers: { Accept: 'application/json;odata=nometadata' } });
    if (!r.ok) throw new Error('contextinfo ' + r.status);
    return (await r.json()).FormDigestValue;
  }
  let token = await digest();
  async function api(path, { method = 'GET', body, merge = false } = {}) {
    const headers = { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': token };
    if (merge) { headers['X-HTTP-Method'] = 'MERGE'; headers['IF-MATCH'] = '*'; }
    const r = await fetch(site + '/_api/' + path, { method, headers, body: body === undefined ? undefined : JSON.stringify(body) });
    if (!r.ok) throw new Error(method + ' ' + path + ' → ' + r.status + ' ' + (await r.text()).slice(0, 300));
    return r.status === 204 ? null : r.json().catch(() => null);
  }
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  // ---- 定義 ----
  // type: text / note / choice / date / datetime / number / user / lookup / url
  const LISTS = [
    { key: 'Equipment', title: '備品台帳', template: 100, titleLabel: '管理番号', uniqueTitle: true, members: 'read', fields: [
      { name: 'ItemName', label: '品名', type: 'text', required: true },
      { name: 'Category', label: '区分', type: 'choice', required: true, choices: ['ノートPC', 'モニター', 'プロジェクター', 'タブレット', 'その他'] },
      { name: 'EquipStatus', label: '状態', type: 'choice', required: true, choices: ['利用可', '修理中', '廃棄'], def: '利用可' },
      { name: 'Location', label: '保管場所', type: 'text' },
      { name: 'Note', label: '備考', type: 'note' },
    ] },
    { key: 'Loans', title: '貸出記録', template: 100, titleLabel: '件名', members: 'none', fields: [
      { name: 'Equipment', label: '備品', type: 'lookup', list: 'Equipment', required: true },
      { name: 'Requester', label: '申請者', type: 'user', required: true },
      { name: 'Purpose', label: '利用目的', type: 'note', required: true },
      { name: 'StartDate', label: '貸出開始日', type: 'date', required: true },
      { name: 'DueDate', label: '返却予定日', type: 'date', required: true },
      { name: 'LoanStatus', label: '状態', type: 'choice', required: true, choices: ['確認中', '承認待ち', '予約済み', '貸出中', '返却済み', '却下', '取消', '期限切れ', 'エラー'], def: '確認中' },
      { name: 'Approver', label: '承認者', type: 'user' },
      { name: 'ApprovalComment', label: '承認コメント', type: 'note' },
      { name: 'ApprovedAt', label: '承認日時', type: 'datetime' },
      { name: 'ApprovalStartedAt', label: '承認開始日時', type: 'datetime' },
      { name: 'ReturnedDate', label: '返却日', type: 'date' },
      { name: 'ReminderCount', label: '督促回数', type: 'number', def: '0' },
      { name: 'LastNotifiedDate', label: '最終通知日', type: 'date' },
      { name: 'ProcessNote', label: '処理メモ', type: 'note' },
    ] },
    { key: 'Faults', title: '故障報告', template: 100, titleLabel: '件名', members: 'none', fields: [
      { name: 'Equipment', label: '備品', type: 'lookup', list: 'Equipment' },
      { name: 'Reporter', label: '報告者', type: 'user' },
      { name: 'Detail', label: '内容', type: 'note' },
      { name: 'FaultStatus', label: '対応状況', type: 'choice', choices: ['確認待ち', '修理中', '完了', '対象外'], def: '確認待ち' },
      { name: 'PlannerTaskId', label: 'Planner タスクID', type: 'text' },
    ] },
    { key: 'Settings', title: '設定', template: 100, titleLabel: '項目', members: 'read', fields: [
      { name: 'SettingValue', label: '値', type: 'text' },
      { name: 'Description', label: '内容', type: 'note' },
    ] },
    { key: 'ProcessLog', title: '処理ログ', template: 100, titleLabel: 'フロー名', members: 'none', fields: [
      { name: 'Target', label: '対象', type: 'text' },
      { name: 'Result', label: '結果', type: 'choice', choices: ['成功', '受け付けず', '失敗'] },
      { name: 'Detail', label: '詳細', type: 'note' },
      { name: 'RunLink', label: '実行の記録', type: 'url' },
    ] },
    { key: 'Reservations', title: '予約表', template: 100, titleLabel: '貸出記録ID', members: 'read', fields: [
      { name: 'Equipment', label: '備品', type: 'lookup', list: 'Equipment' },
      { name: 'StartDate', label: '貸出開始日', type: 'date' },
      { name: 'DueDate', label: '返却予定日', type: 'date' },
      { name: 'LoanStatus', label: '状態', type: 'choice', choices: ['確認中', '承認待ち', '予約済み', '貸出中'] },
    ] },
    { key: 'MonthlyReports', title: '月次レポート', template: 101, members: 'none', fields: [] },
  ];

  // ---- 権限（継承を切り、所有者＝フル コントロール、メンバー＝閲覧またはなし） ----
  // リストを作った直後に行う（途中で失敗しても、メンバーの編集権限を引き継いだリストを残さないため）。
  // 再実行のときも設計どおりに付け直す（所有者以外の付与をすべて外してから付ける）。
  const owners = await api('web/associatedownergroup?$select=Id,Title');
  const members = await api('web/associatedmembergroup?$select=Id,Title');
  const roleId = async (type) => (await api(`web/roledefinitions/getbytype(${type})?$select=Id`)).Id; // 5=フル コントロール, 2=閲覧
  const FULL = await roleId(5);
  const READ = await roleId(2);
  // 実行する人が、サイトの管理者か所有者のグループに入っていることを先に確かめる
  // （個人に直接付与された権限だけで操作していると、下で個人の付与を外した時点で自分が締め出されるため）
  const who = await api('web/currentuser?$select=Id,IsSiteAdmin');
  const inOwners = await api(`web/sitegroups/getbyid(${owners.Id})/users?$select=Id&$filter=Id eq ${who.Id}`).then((r) => r.value.length > 0).catch(() => false);
  if (!who.IsSiteAdmin && !inOwners) throw new Error('実行する人がサイトの管理者でも所有者のグループのメンバーでもないため、止めました');
  async function applyPermissions(def, id) {
    const base = `web/lists(guid'${id}')`;
    // copy=false で切ると、実行した本人が個人として付与される。下で外す（本人はサイトの所有者として引き続き操作できる）
    await api(`${base}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`, { method: 'POST' });
    const current = (await api(`${base}/roleassignments?$select=PrincipalId`)).value.map((x) => x.PrincipalId);
    if (!current.includes(owners.Id)) {
      await api(`${base}/roleassignments/addroleassignment(principalid=${owners.Id},roledefid=${FULL})`, { method: 'POST' });
    }
    // 所有者以外の付与を、付いている役割ごとにすべて外す（実行した本人の個人の付与・以前のメンバーの付与を含む）
    const bindings = (await api(`${base}/roleassignments?$select=PrincipalId,RoleDefinitionBindings/Id&$expand=RoleDefinitionBindings`)).value;
    for (const ra of bindings) {
      if (ra.PrincipalId === owners.Id) continue;
      for (const rd of ra.RoleDefinitionBindings) {
        await api(`${base}/roleassignments/removeroleassignment(principalid=${ra.PrincipalId},roledefid=${rd.Id})`, { method: 'POST' });
      }
    }
    if (def.members === 'read') {
      await api(`${base}/roleassignments/addroleassignment(principalid=${members.Id},roledefid=${READ})`, { method: 'POST' });
    }
    const after = (await api(`${base}/roleassignments?$select=PrincipalId`)).value.map((x) => x.PrincipalId).sort();
    const expected = (def.members === 'read' ? [owners.Id, members.Id] : [owners.Id]).sort();
    if (JSON.stringify(after) !== JSON.stringify(expected)) {
      throw new Error(`権限が設計と一致しません: ${def.title} 実際=${after} 期待=${expected}`);
    }
    say(`権限: ${def.title} → 所有者=フル コントロール、メンバー=${def.members === 'read' ? '閲覧' : 'なし'}（照合済み）`);
  }

  // ---- リストを用意する ----
  const ids = {};
  for (const def of LISTS) {
    let list = await api(`web/lists?$filter=EntityTypeName eq '${def.key}' or EntityTypeName eq '${def.key}List'&$select=Id,Title,EntityTypeName`).then((r) => r.value[0]);
    if (!list) {
      // URL を英数字にするため、英語名で作ってから表示名を日本語に変える
      list = await api('web/lists', { method: 'POST', body: { BaseTemplate: def.template, Title: def.key } });
      say(`作成: ${def.key}`);
    } else {
      say(`既にある: ${list.Title}（${def.key}）`);
    }
    ids[def.key] = list.Id;
    await applyPermissions(def, list.Id);
    if (list.Title !== def.title) {
      await api(`web/lists(guid'${list.Id}')`, { method: 'POST', merge: true, body: { Title: def.title } });
      say(`表示名を設定: ${def.key} → ${def.title}`);
    }
  }

  // ---- 列を用意する ----
  function schema(f) {
    const req = f.required ? 'TRUE' : 'FALSE';
    const base = `Name="${f.name}" StaticName="${f.name}" DisplayName="${esc(f.label)}" Required="${req}"`;
    const def = f.def !== undefined ? `<Default>${esc(f.def)}</Default>` : '';
    switch (f.type) {
      case 'text': return `<Field Type="Text" ${base} MaxLength="255" />`;
      case 'note': return `<Field Type="Note" ${base} NumLines="6" RichText="FALSE" />`;
      case 'choice': return `<Field Type="Choice" ${base} Format="Dropdown"><CHOICES>${f.choices.map((c) => `<CHOICE>${esc(c)}</CHOICE>`).join('')}</CHOICES>${def}</Field>`;
      case 'date': return `<Field Type="DateTime" ${base} Format="DateOnly" />`;
      case 'datetime': return `<Field Type="DateTime" ${base} Format="DateTime" />`;
      case 'number': return `<Field Type="Number" ${base} Decimals="0">${def}</Field>`;
      case 'user': return `<Field Type="User" ${base} UserSelectionMode="PeopleOnly" UserSelectionScope="0" />`;
      case 'lookup': return `<Field Type="Lookup" ${base} List="{${ids[f.list]}}" ShowField="Title" />`;
      case 'url': return `<Field Type="URL" ${base} Format="Hyperlink" />`;
      default: throw new Error('未知の列の種類: ' + f.type);
    }
  }
  for (const def of LISTS) {
    const id = ids[def.key];
    const existing = new Set((await api(`web/lists(guid'${id}')/fields?$select=InternalName`)).value.map((x) => x.InternalName));
    for (const f of def.fields) {
      if (existing.has(f.name)) { say(`  列は既にある: ${def.title}.${f.label}`); continue; }
      // Options 25 = AddToDefaultContentType(1) + AddFieldInternalNameHint(8, 内部名を Name のまま使う) + AddFieldToDefaultView(16)
      await api(`web/lists(guid'${id}')/fields/createfieldasxml`, { method: 'POST', body: { parameters: { SchemaXml: schema(f), Options: 25 } } });
      say(`  列を作成: ${def.title}.${f.label}`);
    }
    if (def.titleLabel) {
      const titleField = `web/lists(guid'${id}')/fields/getbyinternalnameortitle('Title')`;
      const body = { Title: def.titleLabel };
      if (def.uniqueTitle) { body.Indexed = true; body.EnforceUniqueValues = true; }
      await api(titleField, { method: 'POST', merge: true, body });
    }
  }

  // ---- 初期データ（Title で照合し、無い行だけ足す。途中で失敗しても再実行で揃う） ----
  async function addItems(key, rows) {
    const id = ids[key];
    const have = new Set((await api(`web/lists(guid'${id}')/items?$select=Title&$top=5000`)).value.map((x) => x.Title));
    let added = 0;
    for (const row of rows) {
      if (have.has(row.Title)) continue;
      await api(`web/lists(guid'${id}')/items`, { method: 'POST', body: row });
      added++;
    }
    say(`初期データ: ${key} に ${added} 件追加（既にあった ${rows.length - added} 件は変えない）`);
  }
  const me = await api('web/currentuser?$select=Email');
  await addItems('Settings', [
    { Title: '承認者メール', SettingValue: me.Email, Description: '承認依頼の宛先' },
    { Title: '管理者メール', SettingValue: me.Email, Description: 'エラー・督促の上限到達・月次の集計の宛先' },
    { Title: '貸出上限日数', SettingValue: '14', Description: 'これを超える期間の申請は受け付けない' },
    { Title: '予約受付日数', SettingValue: '60', Description: '開始日をこれより先にする申請は受け付けない' },
    { Title: '督促上限回数', SettingValue: '5', Description: '達したら管理者に知らせ、それ以上は督促しない' },
    { Title: '承認期限日数', SettingValue: '7', Description: '過ぎたら期限切れにする' },
    // 既定は「いいえ」（設計: 自己承認は既定で拒否）。ユーザーが1人のデモの環境だけ、実行後に設定リストで「はい」に変える
    { Title: '自己承認を許可', SettingValue: 'いいえ', Description: '「はい」のときだけ申請者本人の承認を受け付ける。既定は「いいえ」' },
  ]);
  const cats = [
    ['PC', 'ノートPC', 'ノートPC 14インチ', 8], ['MN', 'モニター', 'モニター 24インチ', 5],
    ['PJ', 'プロジェクター', 'プロジェクター（会議室用）', 3], ['TB', 'タブレット', 'タブレット 11インチ', 3], ['OT', 'その他', 'Web 会議用スピーカー', 1],
  ];
  const equipment = [];
  for (const [prefix, cat, name, n] of cats) {
    for (let i = 1; i <= n; i++) {
      equipment.push({ Title: `${prefix}-${String(i).padStart(3, '0')}`, ItemName: name, Category: cat, EquipStatus: '利用可', Location: '本社3F 総務部キャビネット' });
    }
  }
  await addItems('Equipment', equipment);

  say('完了');
  return { site: _spPageContextInfo.webServerRelativeUrl, lists: Object.keys(ids), log };
})();
