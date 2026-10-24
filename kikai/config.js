// 自社機械稼働表の設定（GitHub に公開される。ここには社員名・メール・共有リンクを書かないこと）
// clientId / tenantId は全体工程表と同じ Entra の「アプリの登録」（秘密の値ではない）
// folder は 稼働データ フォルダの場所。driveId は全体工程表の 工程データ と同じ OneDrive（北澤さん）。
//   itemId が空のあいだは path で探す（持ち主しか通らない）→ 見つかった itemId を「データ」の小窓で見て、ここへ入れる
window.KIKAI_CONFIG = {
  clientId: "9a4c8447-3e6a-437b-85d0-4662ded35f42",   // アプリケーション (クライアント) ID（全体工程表）
  tenantId: "acd238cd-b6aa-4b5a-8056-44ed0177515f",   // ディレクトリ (テナント) ID
  folder: {
    driveId: "b!fW3m73KhV0i7xf8c58hB1ha1ymrf1RZIq3ibCMmfIEMZJwRVBBcZSqMjvDbFdiXW",
    itemId: "01NJOWDESCMVESZPAWFJCK3KKPBSCGAJKG",   // 稼働データ フォルダ（2026-10-24 北澤さんのサインインで調べた）
    path: "社内共有フォルダ/03_自社機械稼働表/稼働データ"
  }
};
