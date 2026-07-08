# 汎用アイテム選択ダイアログ (Select-Items / Select-DualItems)

いろいろなアプリケーション(当面は Excel VBA)から呼び出して使う、
WPF(XAML)リストボックスの選択ダイアログ。2 種類ある:

- **Select-Items.ps1** … リストボックス 1 つ
- **Select-DualItems.ps1** … リストボックス 2 つ(左右に配置。常に同じ大きさ)

入出力はどちらも **UTF-8 (BOM 付き) の JSON ファイル**で受け渡す。
選択モード(複数選択/単数選択)は入力 JSON の `multiSelect` キーで制御する
(省略時は複数選択。2 リスト版は左右それぞれ個別に指定できる)。

## ファイル構成

| ファイル | 説明 |
|---|---|
| `Select-Items.ps1` | PowerShell (WPF) 版ダイアログ本体(1 リスト)。Windows PowerShell 5.1 用(UTF-8 BOM 付き) |
| `Select-DualItems.ps1` | 同・2 リスト(左右)版 |
| `sample_input.json` | 入力 JSON のサンプル(1 リスト版) |
| `sample_input_dual.json` | 入力 JSON のサンプル(2 リスト版) |
| **VBA 側** (すべて Shift_JIS / CRLF。VBE にそのままインポート可) | |
| `SelectItemsDialog.bas` | エントリポイント。`CreateItemSelector()` / `CreateDualItemSelector()` ファクトリとデモ |
| `JsonLite.bas` | 内部実装用: 簡易 JSON ヘルパー + UTF-8 ファイル I/O(Ps*Selector 専用) |
| **1 リスト版** | |
| `IItemSelector.cls` | インターフェイス。PowerShell 版 / UserForm 版を共通の呼び出し方にする |
| `SelectItemsResult.cls` | 結果クラス(Status / Selected / Message / IsOk) |
| `PsItemSelector.cls` | IItemSelector 実装: Select-Items.ps1 を起動し JSON ファイルで受け渡し |
| `FormItemSelector.cls` | IItemSelector 実装: VBA UserForm を表示し Collection で受け渡し |
| `SelectItemsForm.frm` + `.frx` | UserForm 本体(コントロールは実行時生成。**2 ファイルセットでインポート**) |
| **2 リスト版** | |
| `IDualItemSelector.cls` | インターフェイス(2 リスト版) |
| `ItemListSpec.cls` | リスト 1 つ分の仕様(Caption / MultiSelect / Items)。左右 1 つずつ渡す |
| `SelectDualItemsResult.cls` | 結果クラス(Status / LeftSelected / RightSelected / Message / IsOk) |
| `PsDualItemSelector.cls` | IDualItemSelector 実装: Select-DualItems.ps1 を起動し JSON ファイルで受け渡し |
| `FormDualItemSelector.cls` | IDualItemSelector 実装: VBA UserForm を表示しオブジェクトで受け渡し |
| `SelectDualItemsForm.frm` + `.frx` | UserForm 本体(コントロールは実行時生成。**2 ファイルセットでインポート**) |

## 起動方法

```
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File Select-Items.ps1 `
    -InputPath <入力JSONパス> -OutputPath <出力JSONパス>
```

`Select-DualItems.ps1` も引数・起動方法は同じ(入力 JSON の形式だけ異なる)。

- WPF 表示のため **`-STA` が必要**(powershell.exe の既定は STA だが明示推奨)。
- ダイアログは `Topmost` なので、コンソールを非表示で起動しても前面に出る。
- 引数は両方指定が必須。片方だけだとエラー(終了コード 2)。

## テストモード(引数なしで起動)

```
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File Select-Items.ps1
```

引数なしで起動すると、呼び出し元アプリなしで単体テストできる:

1. 組み込みのテストデータをリストボックスに表示
   (2 リスト版は 左=列名 8 件・複数選択 / 右=出力形式 3 件・単数選択)
2. OK / キャンセル後、選択結果をメッセージボックスで表示
3. 結果 JSON をスクリプトと同じフォルダに出力
   (1 リスト版: `test_output.json` / 2 リスト版: `test_output_dual.json`)
4. 出力ファイルを既定のアプリで開く(.json に関連付けがなければメモ帳)

## 出力ファイルが既に存在する場合

**確認なしで無条件に上書きされる**(追記やバックアップはしない)。
他のプロセスがファイルをロックしていて書き込めない場合はエラー扱い
(終了コード 2)になる。VBA 側の例では、古い結果を誤読しないよう
起動前に出力ファイルを削除している。

## 入力 JSON (1 リスト版: Select-Items.ps1)

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
| `multiSelect` | — | `true` | `true` = 複数選択可(Ctrl/Shift クリック) / `false` = 単数選択 |

## 出力 JSON (1 リスト版)

出力ファイルは **OK・キャンセル・エラーのいずれでも必ず作成される**。

```json
{ "status": "ok",     "selected": ["氏名", "住所"] }
{ "status": "cancel", "selected": [] }
{ "status": "error",  "selected": [], "message": "入力ファイルが見つかりません: ..." }
```

- `status: "cancel"` … キャンセルボタン、または × でウィンドウを閉じた場合
- 何も選択せずに OK を押すと `status: "ok"` で `selected: []`

## 入力 JSON (2 リスト版: Select-DualItems.ps1)

```json
{
    "title": "出力する列と形式を選択してください",
    "left": {
        "caption": "出力する列",
        "multiSelect": true,
        "items": ["氏名", "住所", "電話番号", "メールアドレス"]
    },
    "right": {
        "caption": "出力形式",
        "multiSelect": false,
        "items": ["新規シート", "CSV ファイル", "クリップボード"]
    }
}
```

| キー | 必須 | 既定値 | 説明 |
|---|---|---|---|
| `left` / `right` | ○ | — | 左右それぞれのリストの定義(どちらか欠けるとエラー) |
| `left.items` / `right.items` | ○ | — | 表示するアイテムの文字列配列(空はエラー) |
| `left.caption` など | — | (なし) | リストの上に出す見出し。省略・空なら見出し行を表示しない |
| `left.multiSelect` など | — | `true` | `true` = 複数選択可 / `false` = 単数選択。**左右で別々に指定できる** |
| `title` | — | `項目を選択してください` | ウィンドウタイトル |

## 出力 JSON (2 リスト版)

```json
{ "status": "ok",     "left": ["氏名", "住所"], "right": ["CSV ファイル"] }
{ "status": "cancel", "left": [], "right": [] }
{ "status": "error",  "left": [], "right": [], "message": "..." }
```

`status` と終了コードのルールは 1 リスト版と同じ。

## 終了コード (両版共通)

| コード | 意味 |
|---|---|
| 0 | OK |
| 1 | キャンセル |
| 2 | エラー(入力ファイルなし、JSON 不正など) |

## 操作

- Enter = OK、Esc = キャンセル
- 複数選択は Ctrl+クリック、Shift+クリック
- 1 リスト版のみ: ダブルクリック = 即 OK
  (2 リスト版は片方だけ選んで誤確定しやすいため付けていない)

## XAML を変更するときの注意

スクリプト内の `$xaml` ヒアストリング(`@'...'@`)は自由に編集してよいが、
以下は**変更・削除しないこと**(スクリプト本体が参照している):

| 変更禁止 | 理由 |
|---|---|
| 1 リスト版: `x:Name="lstItems"` / `"btnOK"` / `"btnCancel"` | コード側が `FindName()` でこの名前を検索している。変えると起動時エラー |
| 2 リスト版: `x:Name="lstLeft"` / `"lstRight"` / `"lblLeft"` / `"lblRight"` / `"btnOK"` / `"btnCancel"` | 同上 |
| ルート要素 `<Window>` と 2 行の `xmlns` 宣言 | XamlReader での読み込みに必須 |

また、以下の点にも注意:

- **XAML に `Click="..."` などのイベント属性を書かない**。コードビハインドが
  ないため `XamlReader.Load()` が例外を投げる。イベントはコード側で
  `Add_Click` 等により登録している。
- **`x:Class` 属性を追加しない**(同じ理由で読み込みエラーになる)。
- `Window` の `Title`、`ListBox` の `SelectionMode`、`TextBlock`(見出し)の
  `Text` / `Visibility` は、XAML に書いても**コード側が入力 JSON の値で
  上書きする**ので、XAML で設定しても無効。
- ボタンの `IsDefault="True"` / `IsCancel="True"` は Enter / Esc キーの
  動作を担っている。消すとキー操作が効かなくなる。
- 2 リスト版の左右同サイズは `Grid.ColumnDefinitions` の 2 つの
  `Width="*"` で実現している(中央の `Width="12"` は隙間)。
- ヒアストリングは単一引用符版(`@'...'@`)なので `$` はそのまま書けるが、
  **閉じの `'@` は必ず行頭**に置くこと(インデントすると構文エラー)。
- 色・フォント・サイズ・余白・レイアウト(Grid 構成含む)は自由に変更可。
  ボタンの表示文字列(`Content`)も自由に変えてよい。

## VBA からの使い方

VBA 側のファイル 13 個(.bas ×2、.cls ×8、.frm+.frx ×2 組)をすべて VBE に
インポートする。`PsItemSelector.cls` / `PsDualItemSelector.cls` の
`Class_Initialize` にある `ScriptPath` の既定値を、それぞれ
`Select-Items.ps1` / `Select-DualItems.ps1` の実際のパスに合わせる。

### 1 リスト版

```vba
Dim items As New Collection
items.Add "氏名"
items.Add "住所"

' 実装をここで切り替える (呼び出し方は共通)
Dim selector As IItemSelector
Set selector = CreateItemSelector(skPowerShell)   ' または skUserForm

Dim result As SelectItemsResult
Set result = selector.SelectItems(items, "列を選択してください", True)  ' True = 複数選択

If result.IsOk Then
    ' result.Selected に選択されたアイテム (Collection)
ElseIf result.Status = "cancel" Then
    ' キャンセルされた
Else
    ' result.Message にエラー詳細
End If
```

### 2 リスト版

左右それぞれの仕様を `ItemListSpec` に詰めて渡す。
`MultiSelect` の既定は `True`(複数選択)なので、単数選択にしたい側だけ
`False` を設定すればよい。

```vba
Dim leftList As New ItemListSpec
leftList.Caption = "出力する列"        ' 省略可 (空 = 見出しなし)
leftList.Add "氏名"                    ' Items.Add の省略形
leftList.Add "住所"

Dim rightList As New ItemListSpec
rightList.Caption = "出力形式"
rightList.MultiSelect = False          ' 右だけ単数選択にする
rightList.Add "新規シート"
rightList.Add "CSV ファイル"

' 実装をここで切り替える (呼び出し方は共通)
Dim selector As IDualItemSelector
Set selector = CreateDualItemSelector(skPowerShell)   ' または skUserForm

Dim result As SelectDualItemsResult
Set result = selector.SelectItems(leftList, rightList, "出力する列と形式を選択してください")

If result.IsOk Then
    ' result.LeftSelected / result.RightSelected に選択されたアイテム (Collection)
ElseIf result.Status = "cancel" Then
    ' キャンセルされた
Else
    ' result.Message にエラー詳細
End If
```

### アーキテクチャ

- `IItemSelector` / `IDualItemSelector` インターフェイスで実装を抽象化。
  呼び出し側は `CreateItemSelector(kind)` / `CreateDualItemSelector(kind)` の
  引数を変えるだけで PowerShell 版 / UserForm 版を切り替えられる。
- **JSON はプロセス境界(VBA ⇔ PowerShell)の輸送手段としてのみ使う。**
  UserForm 版は同一プロセスなので Collection / ItemListSpec を直接受け渡す。
  JSON の生成・解析は `PsItemSelector` / `PsDualItemSelector` と内部実装用
  モジュール `JsonLite.bas` に閉じており、利用側・インターフェイスには
  現れない(`JsonLite` を利用側コードから直接呼ばないこと)。
- `SelectItemsForm` / `SelectDualItemsForm` のコントロールは
  `UserForm_Initialize` で実行時に生成している(レイアウト変更はコードの
  定数で行う)。
- デモ(アクティブセルが属するテーブルを使う):
  - 1 リスト版: `Demo_SelectColumns_PowerShell` / `Demo_SelectColumns_UserForm`
  - 2 リスト版: `Demo_SelectDualColumns_PowerShell` / `Demo_SelectDualColumns_UserForm`
    (左=列名・複数選択、右=出力形式・単数選択)
- PowerShell 版の JSON ファイルは既定で `%TEMP%` に作られる
  (1 リスト版: `select_items_in.json` / `select_items_out.json`、
   2 リスト版: `select_dual_items_in.json` / `select_dual_items_out.json`。
   いずれも `Ps*Selector` のプロパティで変更可)。

### VBA ソースファイルの取り扱い注意

- 文字コードは **Shift_JIS、改行は CRLF** を維持すること
  (UTF-8 や LF のままだと VBE へのインポートで壊れる/失敗する)。
- `SelectItemsForm` / `SelectDualItemsForm` は `.frm` と `.frx` の
  2 ファイルセット。インポート時は同じフォルダに両方置くこと。
