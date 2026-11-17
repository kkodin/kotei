/* 全アプリ共通の設定【本番の置き場 kotei のもの。テストへ写さない】（2026-11-17 kotei-test の形に置き換えたときに作った）
   ・テスト（kotei-test）と本番（kotei）で違うのは、このファイルだけ。公開の道具はこのファイルを写さない（それぞれの置き場のものを残す）。
   ・これは本番の値：データは 社内共有フォルダ の 12_工程\工程データ・03_自社機械稼働表\稼働データ・06_配置表\配置表データ。
   ・値は前の本番の config.js（全体工程表・工事日報 nippou4・自社機械稼働表）と同じ。 */
window.HX_CONFIG = {
  clientId: "9a4c8447-3e6a-437b-85d0-4662ded35f42",   // アプリケーション (クライアント) ID
  tenantId: "acd238cd-b6aa-4b5a-8056-44ed0177515f",   // ディレクトリ (テナント) ID
  driveId: "b!fW3m73KhV0i7xf8c58hB1ha1ymrf1RZIq3ibCMmfIEMZJwRVBBcZSqMjvDbFdiXW",   // 社内共有フォルダのあるドライブ
  kotei: { itemId: "01NJOWDERZ6F2LIAJIHZEI7OHUJ44H2GUD" },   // 工程データ フォルダ（日報・詳細日報・マスタもこの下）
  kikai: { itemId: "01NJOWDESCMVESZPAWFJCK3KKPBSCGAJKG", path: "社内共有フォルダ/03_自社機械稼働表/稼働データ" },   // 稼働データ フォルダ
  haichi: { path: "社内共有フォルダ/06_配置表/配置表データ", oldPath: "社内共有フォルダ/02_配置表/配置表データ" },   // 配置表データ（oldPath＝前の名前の場所）
  weather: { lat: 35.444, lon: 139.638 },
  test: false   // 本番
};

/* 昔の名前（各アプリの config.js の形） */
(function (c) {
  const f = id => ({ driveId: c.driveId, itemId: id });
  window.KOTEI_CONFIG = { clientId: c.clientId, tenantId: c.tenantId, folder: f(c.kotei.itemId) };
  window.KIKAI_CONFIG = { clientId: c.clientId, tenantId: c.tenantId, folder: { ...f(c.kikai.itemId), path: c.kikai.path } };
  window.JOUYOU_CONFIG = { clientId: c.clientId, tenantId: c.tenantId, kotei: f(c.kotei.itemId), kikai: f(c.kikai.itemId), haichi: { driveId: c.driveId, ...c.haichi } };
  window.NIPPOU_CONFIG = { clientId: c.clientId, tenantId: c.tenantId, kotei: f(c.kotei.itemId), kikai: f(c.kikai.itemId), haichi: { driveId: c.driveId, ...c.haichi }, nippou: null, weather: c.weather };
})(window.HX_CONFIG);
