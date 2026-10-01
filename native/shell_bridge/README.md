# エクスプローラーからの一括追加

Windows 11のIExplorerCommandから選択全件をローカルCOMサーバーへ渡す。
メニューとCOMサーバーをユーザー単位のSparse Packageに登録する。
従来のレジストリによる表示項目と全ユーザー向けインストールは廃止。既定のPDFアプリは変更しない。

## 構成

- `main.cpp`: IExecuteCommand / IObjectWithSelection、クラスファクトリー、STAの寿命管理。
- `delivery.cpp` / `delivery.h`: IShellItemArrayからのパス取得、本体起動、準備待ち、IPC送信。
- `test_driver.cpp`: 選択配列と別プロセスCOMの統合テスト専用。配布しない。
- `scripts/build_shell_bridge.ps1`: MSVC x64 + Windows SDKでビルドする。C++ランタイムは静的リンク。

`explorer_command.cpp` はWindows 11の新メニュー用IExplorerCommand DLL。
選択配列を既存のローカルCOMサーバーへ渡し、deliveryを共用する。
Explorerのメニュー構築中には本体の起動やPDF処理をしない。
自己署名Sparse Packageの導入・解除は [MODERN_MENU.md](MODERN_MENU.md) を参照。

## 通信と起動

インストール版専用かつログオンセッションごとのパイプを使う。本体が未起動ならヘルパーと同じディレクトリの
`OfficePDFBinder_Main.exe`を一度起動し、最大30秒受信準備を待つ。
起動済みなら本体EXEを新たに起動しない。

フレームは `OPB2` + little-endian uint32の本文バイト数 + UTF-8本文。
本文は改行区切りの絶対パス（上限16 MiB）。本体は全体を受信して追加要求を受け付けた
後に `OK\n` を返す。これは追加・変換の完了を意味しない。
読み書きにはタイムアウトを設け、結果が不明な送信を自動再送しない。
旧来のコマンドライン経由のEOF区切りIPCも引き続き受け付ける。

COMからの一操作は完成した一覧なので、従来の500msの集約待ちは不要。
別の処理中なら一覧をキューに保持し、スレッドが実際に終了してから追加する。
一覧内の自然順ソートと既存の重複確認は本体が担当する。

## ビルドと検証

プロジェクトルートのPowerShellで実行する。

```powershell
./scripts/build_shell_bridge.ps1
./scripts/build_shell_bridge.ps1 -Tests
./scripts/build_shell_bridge.ps1 -IntegrationServer
python -m pytest tests/test_ipc_regressions.py -q
```

PythonはOfficePDFBinderのビルド環境を使う。ネイティブのテスト成果物がない場合は
該当する統合テストをスキップする。本番CLSIDと異なるテスト専用CLSIDを一時登録し、
テスト終了時に削除する。既存の右クリック登録や起動済み本体は使わない。

インストーラー作成時にヘルパーをビルドし、単独のEXEとして同梱する。
**今回の変更では本体も再ビルドすること。古いdistのPackageだけでは新通信に対応しない。**
ポータブル版にはCOM登録を追加しない。

実機での最終確認: 本体の未起動/起動済み、PDF・Office・画像の混在、2/16/100件、
日本語・空白付きパス、連続する二操作、読み込み中の追加、キャンセル、アンインストール。
Windows 11の最初のメニューが対象。「その他のオプションを確認」に従来の項目を残さない。

最新の受入手順は `docs/followup-regression-2026-10-01.md`。以下は過去の検証記録。

## 2026-09-30の検証状況

MSVCビルドとInno Setupの出力なしコンパイルは成功。
エージェントの実行環境では全体のpytestは189件成功、1件失敗。
成功した範囲は、分割受信、日本語パス、2/16/100件の一括受信、実ワーカー終了直前の
追加要求、ネイティブの選択配列からの送信、起動済みのCOMサーバー越しの一括追加。

エージェント環境で失敗したのは、未起動COMサーバーの自動起動
`test_local_com_activation_and_selection[False]`（`0x80040154`）。
その後、ユーザーが通常のWindows環境で同じIPCテストを実行し、
**自動起動・起動済みの両方を含め11件すべて成功（6.52秒、終了コード0）**した。
実行環境による差は確認できたが、エージェント環境側の具体的な制約は未特定。
全体テストすべてを通常環境で再実行した結果ではない点に注意する。

その後ユーザーがFastビルド・インストールを行い、従来メニューでの複数ファイル追加に成功。
PowerShellのコンソール付き起動での異常終了は、タイムアウト接続をlambda経由に変更する対策を
本体ソースに反映済み。小さな検証EXEでは改善を確認。本体再ビルド後の実機確認はまだ必要。

2026-10-01: 新メニューDLLの寿命管理テストに成功。IPCと合わせた12件では11件成功、
1件は上記と同じテスト用COM自動起動の環境依存失敗。新メニューの登録・表示はまだ未確認。
