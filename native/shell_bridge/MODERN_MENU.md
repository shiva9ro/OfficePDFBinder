# Windows 11の右クリックメニュー

現在のユーザー専用です。全ユーザー向けインストールは廃止し、既存の管理者版がある場合は先に削除するよう案内します。
Windows 11の最初のメニューに1項目を登録し、Windows 10には登録しません。従来メニューは削除します。
全ファイルに対する1つの登録から、選択された全ファイルの拡張子が対応形式の場合だけ表示します。
COM配信サーバーもパッケージに登録し、旧LocalServer32への依存を廃止します。

公開証明書の信頼登録はCurrentUser/TrustedPeople、パッケージ登録も現在のユーザーだけです。
新規に追加した証明書の拇印を記録し、アンインストール時に不要な信頼登録だけ削除します。
既に信頼されていた証明書、旧版がLocalMachineに登録した証明書、署名用秘密鍵は削除しません。
通常はインストーラーが自動実行します。登録スクリプトを手動で呼ぶ場合:

```powershell
& "$env:LOCALAPPDATA\Programs\Office PDF Binder\shell-integration\register_sparse_package.ps1" `
  -ApplicationDirectory "$env:LOCALAPPDATA\Programs\Office PDF Binder"
```

解除には `-Unregister` を付けます。管理者として実行する必要はありません。
ログは `%TEMP%/OfficePDFBinder-shell.log`。解除が失敗した場合、アンインストーラーはファイルを消す前に中止します。
表示更新には再サインインが必要な場合があります。Explorerを強制終了する処理はありません。
本体と配信ヘルパーのIPCはログオンセッションごとに分け、別ユーザーの起動と干渉しないようにします。

実機試験: `docs/followup-regression-2026-10-01.md`。
根拠: https://learn.microsoft.com/en-us/windows/apps/desktop/modernize/grant-identity-to-nonpackaged-apps
