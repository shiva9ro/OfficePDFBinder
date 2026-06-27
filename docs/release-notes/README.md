# Release notes management

GitHub Releaseの本文は、このフォルダの`vX.Y.Z.md`を正本として管理します。
GitHubのWeb画面で本文を直接作成せず、先にリリースノートをコミットしてから
GitHub CLI（`gh`）でReleaseを作成・更新します。

過去のリリースノートは、公開時点の内容を履歴として保存します。新しいリリースは
原則として英語・日本語の両方を記載します。

## 日英ページ内リンク

GitHub Releaseでは見出しに自動アンカーが付かないため、日本語部分へのリンクは
明示的なアンカーを使用します。

```markdown
[日本語](#user-content-japanese)

## English

...

<a id="japanese"></a>
## 日本語
```

## 新規リリース

以下はv1.4.0の例です。

```powershell
$version = "1.4.0"
$tag = "v$version"
$notes = "docs/release-notes/$tag.md"
$installer = "Output/OfficePDFBinder_Setup_$version.exe"
$portable = "Output/OfficePDFBinder_Portable_$version.zip"

git add -A
git commit -m "Release $tag"
git push origin main

git tag -a $tag -m "Office PDF Binder $tag"
git push origin $tag

gh release create $tag `
  $installer `
  $portable `
  --repo shiva9ro/OfficePDFBinder `
  --verify-tag `
  --title "Office PDF Binder $tag" `
  --notes-file $notes `
  --latest
```

## 公開済みReleaseの本文更新

```powershell
$tag = "v1.3.0"

gh release edit $tag `
  --repo shiva9ro/OfficePDFBinder `
  --notes-file "docs/release-notes/$tag.md"
```

## 公開済みReleaseの成果物差し替え

同名ファイルを明示的に置き換える場合だけ`--clobber`を使用します。

```powershell
$version = "1.3.0"
$tag = "v$version"

gh release upload $tag `
  "Output/OfficePDFBinder_Setup_$version.exe" `
  "Output/OfficePDFBinder_Portable_$version.zip" `
  --repo shiva9ro/OfficePDFBinder `
  --clobber
```

## 公開後の確認

```powershell
gh release view $tag `
  --repo shiva9ro/OfficePDFBinder `
  --json name,tagName,isDraft,isPrerelease,url,assets
```

ブラウザでも、本文、日英リンク、添付ファイル名、Latest表示を確認します。
