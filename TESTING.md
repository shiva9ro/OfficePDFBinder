# Office PDF Binder テストガイド

## 1. テストの位置づけ

pytest は、コード化された確認条件を繰り返し実行するために使用します。
全テストの成功は、対象とした機能の回帰が検出されなかったことを示しますが、
Office PDF Binder 全体の無欠陥を保証するものではありません。

実際の Microsoft Office、Windows のドラッグ&ドロップ、Nuitka、Inno Setup など、
GitHub Actionsで再現できない項目は手動で確認します。

## 2. 自動テスト

### 実行方法

```powershell
python --version
python -m pip install -r requirements.txt -r requirements-dev.txt
python -m pytest
```

表示文言を追加・変更した場合は、テスト前に翻訳カタログを更新します。

```powershell
pyside6-lupdate OfficePDFBinder_Main.py -ts translations/OfficePDFBinder_en.ts -source-language ja_JP -target-language en_US
```

CIでは同じコマンドを実行し、更新後の `.ts` に未コミット差分が生じた場合は失敗します。

開発・テストではPython 3.13系を使用します。CIでもPython 3.13を使用します。

テスト一覧だけを表示する場合:

```powershell
python -m pytest --collect-only -q
```

ページ番号4形式の目視確認用PDFを残す場合:

```powershell
$env:OFFICEPDFBINDER_KEEP_TEST_ARTIFACTS = "1"
python -m pytest tests/test_page_number_pdf.py
Remove-Item Env:OFFICEPDFBINDER_KEEP_TEST_ARTIFACTS
```

生成先は`tests/artifacts/page_numbers/`です。このフォルダーはGit管理対象外です。

### 内訳

| ファイル | 主な確認内容 |
|---|---|
| `tests/test_cli.py` | CLI引数、既存一括処理への設定引き渡し、汎用フォルダ名、CSVログ、終了コード |
| `tests/test_gui_operations.py` | ページ移動、回転、削除、Undo/Redo、しおり、重複判定、自然順、状態表示、キャンセル、安全な終了、一括処理ダイアログ、画像書き出しDPI |
| `tests/test_i18n.py` | 言語判定、英語翻訳、翻訳カタログ同期、README/HTMLマニュアル、ビルド補助ファイル |
| `tests/test_i18n_source_audit.py` | ユーザー表示文字列が翻訳対象になっていること |
| `tests/test_page_number_format.py` | ページ番号4形式 |
| `tests/test_page_number_pdf.py` | 4形式、回転4方向、横長・既存回転・CropBox・スキャンPDFを実PDFで確認 |
| `tests/test_settings.py` | 設定保存、読込、破損設定からの継続 |
| `tests/test_runtime_mode.py` | 通常版/ポータブル版の実行時パス、設定保存先、一時PDF作成先 |
| `tests/test_ui_baseline.py` | 日本語UIとヘッダー・フッター設定の現行仕様 |
| `tests/test_worker_pdf.py` | PDF読込、異常系、結合、既存PDFの安全な置換、回転、しおり、Office変換制御、Office所有権と一括処理リトライ、画像追加、PDF注釈除去、画像拡大抑制、画像出力、一括処理 |
| `tests/test_bookmark_hierarchy.py` | 階層しおりの読込・保存・削除、同一ページの親子、部分書出し、並べ替え、Undo/Redo |
| `tests/test_audit_regressions.py` | 保存後のページ間リンク、回転画像、TIFFしおり、部分失敗、中止と追加キュー、設定と履歴、一括処理の自己取込み防止 |

正確なテスト件数は、更新漏れを避けるため本文には固定せず、`python -m pytest --collect-only -q` で確認します。`tests/conftest.py` はテストデータとGUI環境を準備するファイルです。
テスト用PDFは実行時に一時フォルダーへ生成し、リポジトリには実データを置きません。

## 3. 手動テスト

公開前、Office処理変更後、ビルド設定変更後に実施します。
個人情報・内部情報を含まないダミーファイルを使用してください。

### 3.1 Microsoft Office変換

- [ ] Wordの1ページ・複数ページ文書を読み込める
- [ ] Excelの複数シート、印刷範囲、横向き設定を反映できる
- [ ] PowerPointの横長・縦長スライドを読み込める
- [ ] Word、Excel、PowerPoint、PDFを混在させて順番どおり保存できる
- [ ] 日本語フォント、図形、画像、表を含む文書の見た目を確認する
- [ ] Wordのコメント・変更履歴が、設定ONでPDFに出ないことを確認する
- [ ] Excelのコメント・メモが、設定ONで印刷/PDF出力されないことを確認する
- [ ] PowerPointは設定ONでも通常どおりPDF化され、保存確認ダイアログが出ない
- [ ] Officeファイルを開いた状態でのエラー表示を確認する
- [ ] Word、Excel、PowerPointをユーザーが開いた状態で変換し、開いていたOfficeが終了しないことを確認する
- [ ] 一括処理で一時的なOffice変換失敗が発生した場合、最大2回再試行されることを確認する
- [ ] 通常のGUI保存ではOffice変換が自動再試行されないことを確認する
- [ ] パスワード保護・破損ファイルで明確なエラーを表示する
- [ ] マクロ有効文書を開いてもAutoOpen／Document_Open／Workbook_Open等が実行されない
- [ ] Word・Excelの外部リンクが更新されず、保存済みの値でPDF化される
- [ ] Office変換後も元ファイルが更新されていない
- [ ] 変換後にWord、Excel、PowerPointのプロセスが残らない

### 3.2 GUI操作

- [ ] ファイル選択とドラッグ&ドロップの両方で追加できる
- [ ] 複数選択の移動、削除、回転が期待どおり動く
- [ ] Undo/Redoを15回程度繰り返して状態が崩れない
- [ ] ズーム5段階とCtrl+ホイールが動作する
- [ ] 長い日本語ファイル名、空白、記号を含むパスを扱える
- [ ] しおりの追加、名前変更、削除、移動後の追従を確認する
- [ ] ヘッダー、フッター、日付、ページ番号の位置を目視確認する
- [ ] 画像ファイルを追加し、横長はA4横、縦長はA4縦になることを確認する
- [ ] `小さい画像を拡大しない` のON/OFFで画像サイズが変わることを確認する
- [ ] 空白ページを挿入し、保存PDFに空白ページが含まれることを確認する
- [ ] PDF注釈付きPDFで、注釈除去ON/OFFの出力差を確認する
- [ ] 保存、画像書き出し、キャンセルの完了表示を確認する
- [ ] 進捗ダイアログの×、中止ボタン、Escで中止を要求しても、後始末が終わるまでダイアログが閉じない
- [ ] Office変換中に中止を要求し、変換元文書・Officeプロセス・一時PDFが残らない
- [ ] 既存PDFの保存中に終了を選んだ場合、既存PDFが置き換わらない
- [ ] 新規PDFの保存完了直前に終了を選んだ場合、完成したPDFだけが残る

### 3.3 単一起動・エクスプローラー連携

- [ ] 起動済みアプリへ別プロセスからファイルを追加できる
- [ ] 複数ファイルが自然順で追加される
- [ ] エクスプローラーの右クリックからPDF・Office文書を追加できる
- [ ] 既定アプリの関連付けを変更していない

### 3.4 ビルド・インストール

- [ ] `build.ps1 -Mode Release`が完了する
- [ ] クリーンなWindows 11 64bit環境でインストールできる
- [ ] Python未導入環境で起動できる
- [ ] README、LICENSE、NOTICE、source.zipが同梱される
- [ ] ユーザー専用でインストールし、旧管理者版との混在を防ぐ
- [ ] ネットワーク上のインストーラー実行を拒否する
- [ ] 旧バージョン検出時の案内が正しい
- [ ] アンインストール後にアプリ本体と右クリックメニューが残らない
- [ ] SmartScreen警告を含む導入手順がREADMEと一致する
- [ ] ポータブル版のフォルダーとZIPに`OfficePDFBinder.portable`が含まれる
- [ ] ポータブル版が「Office PDF Binder（ポータブル版）」と表示され、終了後もAppDataに設定を作らない
- [ ] ポータブル版のOffice変換用一時PDFが元ファイルと同じ場所に作られ、処理後に削除される
- [ ] マーカーなしの通常版は従来どおり設定を保存・復元できる

### 3.5 一括処理

- [ ] 親フォルダを選ぶと、出力フォルダにも同じパスが初期設定される
- [ ] サブフォルダごとに、直下の対応ファイルだけがファイル名順でPDF化される
- [ ] 未対応ファイルがCSVログに記録される
- [ ] 既存PDFがある場合、上書きOFFではスキップされる
- [ ] 既存PDFがある場合、上書きONではPDFが置き換わる
- [ ] 一括処理でも自動しおり設定、しおり表示設定、PDF注釈除去、画像拡大抑制、Word/Excelレビュー情報抑制が反映される

### 3.6 CLI一括処理

- [ ] PowerShellから`OfficePDFBinder_Main.exe --batch-subfolders`を実行できる
- [ ] CLI実行中にGUIウィンドウや新しいコンソールウィンドウが表示されない
- [ ] 標準出力に処理件数とログパスが表示され、`$LASTEXITCODE`が結果と一致する
- [ ] CLI一括処理でもOffice変換失敗時に最大2回再試行され、回復時は終了コード0になる
- [ ] CLIオプションで、しおり、注釈除去、画像拡大抑制、Word/Excelレビュー情報抑制を切り替えられる
- [ ] portable版とインストーラー版で同じCLI引数を使用できる

### 3.7 性能・大容量

- [ ] 100ページ程度のPDFを操作・保存できる
- [ ] 500ページ以上では操作時間とメモリ使用量を記録する
- [ ] 大きな画像を含むPDFでサムネイルと保存を確認する
- [ ] 保存先の空き容量不足・書込権限不足を確認する
- [ ] ネットワーク切断や元ファイル移動時のエラーを確認する

Python環境間の速度比較は、通常のpytestとは分けて実行します。

```powershell
python tools\benchmark_python_versions.py `
  --baseline "C:\Path\To\Python312\python.exe" `
  --candidate "C:\Path\To\Python313\python.exe" `
  --samples 3 `
  --warmups 1 `
  --rounds 3 `
  --json-output "benchmark_python_312_vs_313.json"
```

起動、Python処理、PDF結合、JPEG画像書き出しを交互に測定し、中央値と
候補環境の増減率を表示します。負の増減率は候補環境が高速だったことを示します。
Python以外の依存ライブラリのバージョンも結果に記録されるため、バージョンが
異なる場合はPython単体ではなく環境全体の比較として扱います。

ビルド時間を計測する場合は、Python 3.13環境を有効化してビルドします。
各コマンドの終了時に、結果が`Backup/build-metrics/`のJSONへ保存されます。

```powershell
conda activate OfficePDFBinder313
.\build.ps1 -Mode Release
```

正式なWindowsビルド環境はPython 3.13とVisual StudioのClang-clです。
Nuitkaのコンパイラキャッシュによって、同一環境での2回目以降の
`Release`ビルドも大幅に短縮される場合があります。

## 4. 手動テスト記録

リリースごとに次を記録します。

| 項目 | 記録内容 |
|---|---|
| 実施日 | YYYY-MM-DD |
| バージョン／コミット | 例: v1.3.0 / commit hash |
| Windows | Windows 10 または 11、ビルド番号 |
| Office | Microsoft 365 / Office 2021等、32/64bit |
| 結果 | 合格／条件付き合格／不合格 |
| 残課題 | 再現手順、対象ファイルの特徴、回避策 |

実データや内部文書はリポジトリへ追加せず、必要な場合は特徴だけを記録してください。
# エクスプローラー一括追加のネイティブテスト

自動テストはプロジェクト直下で `python -m pytest -q` を実行します。
手動確認では、対象のバージョンとインストール版／ソース実行の別を記録してください。

COMヘルパーのビルド・通信仕様・実機確認項目は
[native/shell_bridge/README.md](native/shell_bridge/README.md)を参照してください。
`build_shell_bridge.ps1` の通常、`-Tests`、`-IntegrationServer` の3種類をビルドすると、
`tests/test_ipc_regressions.py` が実際のネイティブコードとCOM通信も検証します。
COMテストはランダムなテスト用CLSIDを一時登録し、終了時に削除します。

GitHub ActionsではWindows x64用のテストドライバー、COMテストサーバー、
Explorer DLLとSparse Packageを作成してから、全pytestテストを実行します。
アプリ本体のNuitkaビルドは実行しません。署名には使い捨てのランナー内で
作成する自己署名証明書を使い、配布用の秘密鍵やGitHub Secretsは使いません。
パッケージの実インストールは行わず、登録・解除は模擬処理で検証します。
必要ファイルの欠落、およびIPC・Explorer・登録解除テストのスキップはCI失敗とします。
結果のXMLはActionsの `windows-test-results` からダウンロードできます。
ARM64での動作、実際の右クリック表示・前面化・インストールは別途手動確認します。
