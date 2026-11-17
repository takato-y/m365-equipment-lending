/**
 * seed-test-loans.js
 * FLW-003（期限の通知・督促）のテスト用に、架空の貸出記録を作る。件名の先頭に【テスト】を付ける。
 * 使い方: チームサイトの「サイトの設定」ページを所有者のアカウントで開き、開発者ツールのコンソールに貼り付けて実行する。
 * 日付は、日本時間の今日を基準にする（サイトの地域設定が日本時間であることが前提）。
 * 申請者は実行したユーザー。承認を通さずにリストへ直接書くのは、テストのため（社員はリストに書き込めない）。
 */
(async () => {
  const site = location.origin + _spPageContextInfo.webServerRelativeUrl;
  const H = { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' };
  const digest = (await (await fetch(site + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const post = async (path, body) => {
    const r = await fetch(site + '/_api/' + path, { method: 'POST', headers: { ...H, 'X-RequestDigest': digest }, body: JSON.stringify(body) });
    if (!r.ok) throw new Error(path + ' → ' + r.status + ' ' + (await r.text()).slice(0, 300));
    return r.json();
  };
  // 日本時間のその日の 0:00 を UTC で表す（フローと同じ扱い）
  const jstDay = (offset) => {
    const now = new Date(Date.now() + 9 * 3600e3);
    const d = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate() + offset) - 9 * 3600e3);
    return d.toISOString().replace('.000Z', 'Z');
  };
  const me = await (await fetch(site + '/_api/web/currentuser?$select=Id', { headers: H })).json();
  const equip = async (code) => (await (await fetch(site + `/_api/web/lists/getbytitle('備品台帳')/items?$select=Id&$filter=Title eq '${code}'`, { headers: H })).json()).value[0].Id;
  const cases = [
    { key: 'A 開始日を迎えた予約', code: 'PC-008', LoanStatus: '予約済み', StartDate: jstDay(0), DueDate: jstDay(2) },
    { key: 'B 返却期限が明日', code: 'MN-001', LoanStatus: '貸出中', StartDate: jstDay(-2), DueDate: jstDay(1) },
    { key: 'C 期限を2日過ぎた', code: 'MN-002', LoanStatus: '貸出中', StartDate: jstDay(-5), DueDate: jstDay(-2), ReminderCount: 0 },
    { key: 'D 督促が上限の1回前', code: 'MN-003', LoanStatus: '貸出中', StartDate: jstDay(-9), DueDate: jstDay(-4), ReminderCount: 4 },
    // 督促が上限（5回）に達した後の記録。これ以上督促しないこと（督促回数が 5 のまま）を確かめる
    { key: 'F 督促が上限に達した後', code: 'MN-004', LoanStatus: '貸出中', StartDate: jstDay(-12), DueDate: jstDay(-6), ReminderCount: 5 },
    { key: 'E 承認の途中で止まった', code: 'PJ-001', LoanStatus: '承認待ち', StartDate: jstDay(20), DueDate: jstDay(21),
      ApprovalStartedAt: new Date(Date.now() - 9 * 86400e3).toISOString() },
  ];
  const made = [];
  for (const c of cases) {
    const body = { Title: `【テスト】${c.key}`, EquipmentId: await equip(c.code), RequesterId: me.Id, Purpose: 'FLW-003 のテスト（架空）',
      StartDate: c.StartDate, DueDate: c.DueDate, LoanStatus: c.LoanStatus };
    if (c.ReminderCount !== undefined) body.ReminderCount = c.ReminderCount;
    if (c.ApprovalStartedAt) body.ApprovalStartedAt = c.ApprovalStartedAt;
    const item = await post(`web/lists/getbytitle('貸出記録')/items`, body);
    if (['予約済み', '承認待ち'].includes(c.LoanStatus)) {
      await post(`web/lists/getbytitle('予約表')/items`, { Title: String(item.Id), EquipmentId: body.EquipmentId, StartDate: c.StartDate, DueDate: c.DueDate, LoanStatus: c.LoanStatus });
    }
    made.push({ case: c.key, id: item.Id });
  }
  console.log('[seed] ' + JSON.stringify(made));
  return made;
})();
