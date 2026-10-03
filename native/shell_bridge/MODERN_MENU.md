# Windows 11の右クリックメニュー

現在のユーザー専用です。全ユーザー向けインストールは廃止し、既存の管理者版がある場合は先に削除するよう案内します。
Windows 11の最初のメニューに1項目を登録し、Windows 10には登録しません。従来メニューは削除します。
全ファイルに対する1つの登録から、選択された全ファイルの拡張子が対応形式の場合だけ表示します。
COM配信サーバーもパッケージに登録し、旧LocalServer32への依存を廃止します。

公開証明書の信頼登録はLocalMachine/TrustedPeopleで、UACによる管理者権限の確認が必要です。パッケージ登録は現在のユーザーだけです。
アンインストール時は全ユーザーのパッケージを照会し、残っていなければ配布証明書と拇印が一致する証明書を削除します。署名用秘密鍵は削除しません。
照会・証明書削除の失敗やUAC拒否では証明書を残し、警告して本体削除を続行します。
通常はインストーラーが自動実行します。証明書登録済みの環境で、パッケージ登録だけを再実行する場合:

```powershell
& "$env:LOCALAPPDATA\Programs\Office PDF Binder\shell-integration\register_sparse_package.ps1" `
  -ApplicationDirectory "$env:LOCALAPPDATA\Programs\Office PDF Binder"
```

パッケージ解除には `-Unregister` を付けます。この操作には管理者権限は不要です。
証明書操作は管理者権限で `-InstallMachineCertificate` または `-UninstallMachineCertificate` を指定します。これらのスイッチは同時指定しません。
ログは操作を実行したユーザーの `%TEMP%/OfficePDFBinder-shell.log`。パッケージ自体の解除が失敗した場合、アンインストーラーはファイルを消す前に中止します。
表示更新には再サインインが必要な場合があります。Explorerを強制終了する処理はありません。
本体と配信ヘルパーのIPCはログオンセッションごとに分け、別ユーザーの起動と干渉しないようにします。

試験の範囲と手順: `TESTING.md`。
根拠: https://learn.microsoft.com/en-us/windows/apps/desktop/modernize/grant-identity-to-nonpackaged-apps
