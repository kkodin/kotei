/* 全アプリ共通のサインイン（2026-11-16 開発環境の整理）
   ・戻り先（リダイレクト URI）は、いつも置き場のいちばん上（kotei/・kotei-test/。Entra に登録済み）。
     サインインのあと、MSAL が始めた画面（nippou/・nippou/shousai/ など）へ戻す（工事部会資料で使ってきたやり方）。
     → アプリを足しても、アドレスを変えても、Entra に戻り先を足さなくてよい。
   ・サインインの控えはブラウザ（localStorage）に置く。同じ置き場のアプリどうしで共有される（1回サインインすれば、ほかのアプリもそのまま開ける）。
   使い方：
     const s = await HX.signIn();   // { app, account, error }
     if (!s.account) { ボタン.onclick = () => HX.login(s.app); return; }
     const t = await HX.token(s.app);   // 読み書きの鍵 */
window.HX = window.HX || {};
(function (HX) {
  const me = document.currentScript && document.currentScript.src;
  HX.root = me ? new URL("../", me).href : new URL("/", location.href).href;   // 置き場のいちばん上（common/ の1つ上）
  HX.SCOPES = ["Files.ReadWrite.All", "User.Read"];
  HX.signIn = async function () {
    const c = window.HX_CONFIG || {};
    const app = new msal.PublicClientApplication({
      auth: { clientId: c.clientId, authority: "https://login.microsoftonline.com/" + c.tenantId, redirectUri: HX.root, navigateToLoginRequestUrl: true },
      cache: { cacheLocation: "localStorage" }
    });
    await app.initialize();
    let error = null;
    try { const r = await app.handleRedirectPromise(); if (r) app.setActiveAccount(r.account); } catch (e) { error = e; }
    const account = app.getActiveAccount() || app.getAllAccounts()[0] || null;
    if (account) app.setActiveAccount(account);
    return { app, account, error };
  };
  HX.login = (app, scopes) => app.loginRedirect({ scopes: scopes || HX.SCOPES, redirectStartPage: location.href });
  HX.token = async function (app, scopes) {
    const sc = scopes || HX.SCOPES;
    try { return (await app.acquireTokenSilent({ scopes: sc })).accessToken; }
    catch (e) { if (e instanceof msal.InteractionRequiredAuthError) await app.acquireTokenRedirect({ scopes: sc, redirectStartPage: location.href }); throw e; }
  };
})(window.HX);
