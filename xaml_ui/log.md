# log

## QA
### 1. JSON 処理を JScript に任せるべきか? → 採用しませんでした

定番の MSScriptControl.ScriptControl は 32bit 専用で、現在主流の 64bit Office では動きません。
代替の htmlfile + JSON.parse トリックは動きますが、廃止方向の IE コンポーネントへの依存が増えるうえ、JScript オブジェクトから VBA へ値を取り出すヘルパーが結局必要で、行数はあまり減りません。
本件の JSON は自分で設計した固定フォーマットなので、依存ゼロの簡易パーサーを PsItemSelector クラスの Private メソッドに封じ込めました。
利用側からは JSON の存在自体が見えなくなったので、「シンプルさ」の目的はクラス化の方で達成しています。

### 2. クラス導入 → VBA は導入、ps1 は見送り

VBA 側は下記の構成に全面リファクタリングしました。
ps1 側は見送りました:PowerShell 5.1 の class は WPF イベントハンドラー内のスコープの癖、Add-Type より先に型解決されるパース順序の問題など制約が多く、150 行程度のスクリプトでは関数分割の現状の方が読みやすいためです。

### 3. ps1 / Form の切り替え → インターフェイス + ファクトリで実現

Set selector = CreateItemSelector(skPowerShell)   ' ← ここを skUserForm に変えるだけ
Set result = selector.SelectItems(items, "列を選択してください", True)

###4. Form の場合のデータ受け渡し → Collection にしました(判断理由つき)

JSON は「プロセス境界(VBA ⇔ PowerShell)を越えるための輸送手段」であり、同一プロセス内の UserForm にファイル経由の JSON を強いるのは失敗要因(ファイル I/O、パース)を無意味に増やすだけです。
そこで共通化はデータ形式ではなくインターフェイス(IItemSelector)のレベルで行いました。
呼び出し側はどちらの実装でも Collection を渡して SelectItemsResult を受け取るだけで、JSON は PsItemSelector の内部実装に完全に隠蔽されています。
つまり「Form なら Collection が良い、ただしインターフェイスを実装すべき」— 両方やりました。

## 新しいファイル構成(VBA 側は全て Shift_JIS / CRLF)
IItemSelector.cls — インターフェイス
SelectItemsResult.cls — 結果(Status / Selected / Message / IsOk)
PsItemSelector.cls — PowerShell 版実装(JSON 処理を内包。ScriptPath は要調整)
FormItemSelector.cls — UserForm 版実装
SelectItemsForm.frm + .frx — フォーム本体(コントロールは実行時生成。2 ファイルセットでインポート)
SelectItemsDialog.bas — ファクトリとデモ 2 種(Demo_SelectColumns_PowerShell / _UserForm)

## 検証結果(実 Excel 16 にインポートして実行)
- 6 ファイルすべてインポート成功、コンパイル成功
- ファクトリ・結果クラス・フォームの動的コントロール生成(日本語キャプション含む)すべて正常
- エンドツーエンド: VBA → PsItemSelector → WPF ダイアログで「氏名」「電話番号」を選択 → ok/氏名/電話番号 が Collection で返ることを確認

途中 2 つの罠を踏んで修正しました。
(1) VBA ソースが LF 改行だと全モジュールが不可視のコンパイルエラーになる(CRLF に修正済み)、
(2) frx なしの手書き .frm は VBE が受け付けないため、Excel 自身にフォームを生成させて正規の .frm/.frx ペアをエクスポートしました。
今後 VBA ファイルを編集する際は Shift_JIS + CRLF の維持が必須です(README に記載済み)。

UserForm 版の実際の画面操作(クリックして OK)だけは自動化が困難なため未検証です。Demo_SelectColumns_UserForm を一度実行して見た目をご確認
ください。
