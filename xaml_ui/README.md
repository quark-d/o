# Select-Items.ps1 — 汎用アイテム選択ダイアログ

いろいろなアプリケーション(当面は Excel VBA)から呼び出して使う、
WPF(XAML)リストボックスの選択ダイアログ。
入出力はどちらも **UTF-8 (BOM 付き) の JSON ファイル**で受け渡す。

## ファイル構成

| ファイル | 説明 |
|---|---|
| `Select-Items.ps1` | PowerShell (WPF) 版ダイアログ本体。Windows PowerShell 5.1 用(UTF-8 BOM 付き) |
| `sample_input.json` | 入力 JSON のサンプル |
| **VBA 側** (すべて Shift_JIS / CRLF。VBE にそのままインポート可) | |
| `SelectItemsDialog.bas` | エントリポイント。`CreateItemSelector()` ファクトリとデモ |
| `IItemSelector.cls` | インターフェイス。PowerShell 版 / UserForm 版を共通の呼び出し方にする |
| `SelectItemsResult.cls` | 結果クラス(Status / Selected / Message / IsOk) |
| `PsItemSelector.cls` | IItemSelector 実装: Select-Items.ps1 を起動し JSON ファイルで受け渡し |
| `FormItemSelector.cls` | IItemSelector 実装: VBA UserForm を表示し Collection で受け渡し |
| `SelectItemsForm.frm` + `.frx` | UserForm 本体(コントロールは実行時生成。**2 ファイルセットでインポート**) |

## 起動方法

```
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File Select-Items.ps1 `
    -InputPath <入力JSONパス> -OutputPath <出力JSONパス>
```

- WPF 表示のため **`-STA` が必要**(powershell.exe の既定は STA だが明示推奨)。
- ダイアログは `Topmost` なので、コンソールを非表示で起動しても前面に出る。
- 引数は両方指定が必須。片方だけだとエラー(終了コード 2)。

## テストモード(引数なしで起動)

```
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File Select-Items.ps1
```

引数なしで起動すると、呼び出し元アプリなしで単体テストできる:

1. 組み込みのテストデータ(氏名・住所などの列名 8 件)をリストボックスに表示
2. OK / キャンセル後、選択結果をメッセージボックスで表示
3. 結果 JSON をスクリプトと同じフォルダの `test_output.json` に出力
4. 出力ファイルを既定のアプリで開く(.json に関連付けがなければメモ帳)

## 出力ファイルが既に存在する場合

**確認なしで無条件に上書きされる**(追記やバックアップはしない)。
他のプロセスがファイルをロックしていて書き込めない場合はエラー扱い
(終了コード 2)になる。VBA 側の例では、古い結果を誤読しないよう
起動前に出力ファイルを削除している。

## 入力 JSON

```json
{
    "title": "列を選択してください",
    "multiSelect": true,
    "items": ["氏名", "住所", "電話番号"]
}
```

| キー | 必須 | 既定値 | 説明 |
|---|---|---|---|
| `items` | ○ | — | 表示するアイテムの文字列配列(空はエラー) |
| `title` | — | `項目を選択してください` | ウィンドウタイトル |
| `multiSelect` | — | `true` | `true` = 複数選択可(Ctrl/Shift クリック) |

## 出力 JSON

出力ファイルは **OK・キャンセル・エラーのいずれでも必ず作成される**。

```json
{ "status": "ok",     "selected": ["氏名", "住所"] }
{ "status": "cancel", "selected": [] }
{ "status": "error",  "selected": [], "message": "入力ファイルが見つかりません: ..." }
```

- `status: "cancel"` … キャンセルボタン、または × でウィンドウを閉じた場合
- 何も選択せずに OK を押すと `status: "ok"` で `selected: []`

## 終了コード

| コード | 意味 |
|---|---|
| 0 | OK |
| 1 | キャンセル |
| 2 | エラー(入力ファイルなし、JSON 不正など) |

## 操作

- ダブルクリック / Enter = OK、Esc = キャンセル
- 複数選択は Ctrl+クリック、Shift+クリック

## XAML を変更するときの注意

スクリプト内の `$xaml` ヒアストリング(`@'...'@`)は自由に編集してよいが、
以下は**変更・削除しないこと**(スクリプト本体が参照している):

| 変更禁止 | 理由 |
|---|---|
| `x:Name="lstItems"` / `"btnOK"` / `"btnCancel"` | コード側が `FindName()` でこの名前を検索している。変えると起動時エラー |
| ルート要素 `<Window>` と 2 行の `xmlns` 宣言 | XamlReader での読み込みに必須 |

また、以下の点にも注意:

- **XAML に `Click="..."` などのイベント属性を書かない**。コードビハインドが
  ないため `XamlReader.Load()` が例外を投げる。イベントはコード側で
  `Add_Click` 等により登録している。
- **`x:Class` 属性を追加しない**(同じ理由で読み込みエラーになる)。
- `Window` の `Title` と `ListBox` の `SelectionMode` は、XAML に書いても
  **コード側が入力 JSON の値で上書きする**ので、XAML で設定しても無効。
- ボタンの `IsDefault="True"` / `IsCancel="True"` は Enter / Esc キーの
  動作を担っている。消すとキー操作が効かなくなる。
- ヒアストリングは単一引用符版(`@'...'@`)なので `$` はそのまま書けるが、
  **閉じの `'@` は必ず行頭**に置くこと(インデントすると構文エラー)。
- 色・フォント・サイズ・余白・レイアウト(Grid 構成含む)は自由に変更可。
  ボタンの表示文字列(`Content`)も自由に変えてよい。

## VBA からの使い方

VBA 側のファイル 6 つ(.bas ×1、.cls ×4、.frm+.frx)をすべて VBE に
インポートする。`PsItemSelector.cls` の `Class_Initialize` にある
`ScriptPath` の既定値を `Select-Items.ps1` の実際のパスに合わせる。

```vba
Dim items As New Collection
items.Add "氏名"
items.Add "住所"

' 実装をここで切り替える (呼び出し方は共通)
Dim selector As IItemSelector
Set selector = CreateItemSelector(skPowerShell)   ' または skUserForm

Dim result As SelectItemsResult
Set result = selector.SelectItems(items, "列を選択してください", True)

If result.IsOk Then
    ' result.Selected に選択されたアイテム (Collection)
ElseIf result.Status = "cancel" Then
    ' キャンセルされた
Else
    ' result.Message にエラー詳細
End If
```

### アーキテクチャ

- `IItemSelector` インターフェイスで実装を抽象化。呼び出し側は
  `CreateItemSelector(kind)` の引数を変えるだけで PowerShell 版 /
  UserForm 版を切り替えられる。
- **JSON はプロセス境界(VBA ⇔ PowerShell)の輸送手段としてのみ使う。**
  UserForm 版は同一プロセスなので Collection を直接受け渡す。
  JSON の生成・解析は `PsItemSelector` の内部に閉じており、
  利用側・インターフェイスには現れない。
- `SelectItemsForm` のコントロールは `UserForm_Initialize` で実行時に
  生成している(レイアウト変更はコードの定数で行う)。
- デモ: `Demo_SelectColumns_PowerShell` / `Demo_SelectColumns_UserForm`
  (アクティブセルが属するテーブルの列名を選択させる)。
- PowerShell 版の JSON ファイルは既定で `%TEMP%\select_items_in.json` /
  `select_items_out.json` に作られる(`PsItemSelector` のプロパティで変更可)。

### VBA ソースファイルの取り扱い注意

- 文字コードは **Shift_JIS、改行は CRLF** を維持すること
  (UTF-8 や LF のままだと VBE へのインポートで壊れる/失敗する)。
- `SelectItemsForm` は `.frm` と `.frx` の 2 ファイルセット。
  インポート時は同じフォルダに両方置くこと。
