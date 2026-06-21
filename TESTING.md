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
pip install -r requirements.txt -r requirements-dev.txt
python -m pytest
```

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

| ファイル | 件数 | 主な確認内容 |
|---|---:|---|
| `tests/test_gui_operations.py` | 13 | ページ移動、回転、削除、Undo/Redo、しおり、重複判定、自然順、状態表示、キャンセル |
| `tests/test_page_number_format.py` | 4 | ページ番号4形式 |
| `tests/test_page_number_pdf.py` | 12 | 4形式、回転4方向、横長・既存回転・CropBox・スキャンPDFを実PDFで確認 |
| `tests/test_settings.py` | 3 | 設定保存、読込、破損設定からの継続 |
| `tests/test_ui_baseline.py` | 2 | 日本語UIとヘッダー・フッター設定の現行仕様 |
| `tests/test_worker_pdf.py` | 13 | PDF読込、異常系、結合、回転、しおり、Office変換制御、画像出力 |
| **合計** | **47** | |

`tests/conftest.py` はテストデータとGUI環境を準備するファイルで、テスト件数には含みません。
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
- [ ] Officeファイルを開いた状態でのエラー表示を確認する
- [ ] パスワード保護・破損ファイルで明確なエラーを表示する
- [ ] 変換後にWord、Excel、PowerPointのプロセスが残らない

### 3.2 GUI操作

- [ ] ファイル選択とドラッグ&ドロップの両方で追加できる
- [ ] 複数選択の移動、削除、回転が期待どおり動く
- [ ] Undo/Redoを15回程度繰り返して状態が崩れない
- [ ] ズーム5段階とCtrl+ホイールが動作する
- [ ] 長い日本語ファイル名、空白、記号を含むパスを扱える
- [ ] しおりの追加、名前変更、削除、移動後の追従を確認する
- [ ] ヘッダー、フッター、日付、ページ番号の位置を目視確認する
- [ ] 保存、画像書き出し、キャンセルの完了表示を確認する

### 3.3 単一起動・エクスプローラー連携

- [ ] 起動済みアプリへ別プロセスからファイルを追加できる
- [ ] 複数ファイルが自然順で追加される
- [ ] エクスプローラーの右クリックからPDF・Office文書を追加できる
- [ ] 既定アプリの関連付けを変更していない

### 3.4 ビルド・インストール

- [ ] `build.ps1 -Mode Release`が完了する
- [ ] クリーンなWindows 10/11 64bit環境でインストールできる
- [ ] Python未導入環境で起動できる
- [ ] README、LICENSE、NOTICE、source.zipが同梱される
- [ ] 通常ユーザーと管理者の両方でインストール先が適切になる
- [ ] ネットワーク上のインストーラー実行を拒否する
- [ ] 旧バージョン検出時の案内が正しい
- [ ] アンインストール後にアプリ本体と右クリックメニューが残らない
- [ ] SmartScreen警告を含む導入手順がREADMEと一致する
- [ ] ポータブル版のフォルダーとZIPに`OfficePDFBinder.restricted-portable`が含まれる
- [ ] ポータブル版が「Office PDF Binder（ポータブル版）」と表示され、終了後もAppDataに設定を作らない
- [ ] ポータブル版のOffice変換用一時PDFが元ファイルと同じ場所に作られ、処理後に削除される
- [ ] マーカーなしの通常版は従来どおり設定を保存・復元できる

### 3.5 性能・大容量

- [ ] 100ページ程度のPDFを操作・保存できる
- [ ] 500ページ以上では操作時間とメモリ使用量を記録する
- [ ] 大きな画像を含むPDFでサムネイルと保存を確認する
- [ ] 保存先の空き容量不足・書込権限不足を確認する
- [ ] ネットワーク切断や元ファイル移動時のエラーを確認する

## 4. 手動テスト記録

リリースごとに次を記録します。

| 項目 | 記録内容 |
|---|---|
| 実施日 | YYYY-MM-DD |
| バージョン／コミット | 例: v1.2.0 / commit hash |
| Windows | Windows 10 または 11、ビルド番号 |
| Office | Microsoft 365 / Office 2021等、32/64bit |
| 結果 | 合格／条件付き合格／不合格 |
| 残課題 | 再現手順、対象ファイルの特徴、回避策 |

実データや内部文書はリポジトリへ追加せず、必要な場合は特徴だけを記録してください。
