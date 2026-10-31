// 常用日報の設定（GitHub に公開される。ここには社員名・メール・共有リンクを書かないこと）
// clientId / tenantId は全体工程表・工事日報と同じ Entra の「アプリの登録」（秘密の値ではない）
// マスタはほかのアプリと同じ所を読むだけ（書かない）：
//   kotei：全体工程表の 工程データ フォルダ → マスタ.json（配置表の名簿：社員・従業員・外注業者）
//   kikai：自社機械稼働表の 稼働データ フォルダ → 現場一覧.json（NAS の受注台帳などから毎朝作る：現場・元請）・稼働データ.json（自社機械）
window.JOUYOU_CONFIG = {
  clientId: "9a4c8447-3e6a-437b-85d0-4662ded35f42",
  tenantId: "acd238cd-b6aa-4b5a-8056-44ed0177515f",
  kotei: {
    driveId: "b!fW3m73KhV0i7xf8c58hB1ha1ymrf1RZIq3ibCMmfIEMZJwRVBBcZSqMjvDbFdiXW",
    itemId: "01NJOWDERZ6F2LIAJIHZEI7OHUJ44H2GUD"
  },
  kikai: {
    driveId: "b!fW3m73KhV0i7xf8c58hB1ha1ymrf1RZIq3ibCMmfIEMZJwRVBBcZSqMjvDbFdiXW",
    itemId: "01NJOWDESCMVESZPAWFJCK3KKPBSCGAJKG"
  }
};
