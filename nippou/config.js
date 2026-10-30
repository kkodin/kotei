// 工事日報の設定（GitHub に公開される。ここには社員名・メール・共有リンクを書かないこと）
// clientId / tenantId は全体工程表と同じ Entra の「アプリの登録」（秘密の値ではない）
// kotei：全体工程表の 工程データ フォルダ。工程（予定）とマスタを読み、日報はその中の「日報」フォルダに書く
//        （工程データ の共有がそのまま効くので、新しく共有を頼まなくてよい）
// kikai：自社機械稼働表の 稼働データ フォルダ。予約を「使用機材」の予定に使う（読むだけ。開けなければ機械の予定なしで動く）
window.NIPPOU_CONFIG = {
  clientId: "9a4c8447-3e6a-437b-85d0-4662ded35f42",
  tenantId: "acd238cd-b6aa-4b5a-8056-44ed0177515f",
  kotei: {
    driveId: "b!fW3m73KhV0i7xf8c58hB1ha1ymrf1RZIq3ibCMmfIEMZJwRVBBcZSqMjvDbFdiXW",
    itemId: "01NJOWDERZ6F2LIAJIHZEI7OHUJ44H2GUD"
  },
  kikai: {
    driveId: "b!fW3m73KhV0i7xf8c58hB1ha1ymrf1RZIq3ibCMmfIEMZJwRVBBcZSqMjvDbFdiXW",
    itemId: "01NJOWDESCMVESZPAWFJCK3KKPBSCGAJKG"
  },
  // 天候を自動で入れる場所（横浜市役所あたり）。open-meteo（登録不要の天気の API）に緯度経度だけを送る
  weather: { lat: 35.444, lon: 139.638 }
};
