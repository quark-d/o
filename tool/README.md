# 在庫管理台帳 入力ツール

在庫管理台帳.xlsx (イベントログ方式) への入力 UI。台帳ブック本体はデータ専用の
.xlsx のまま変更せず、**別ブック「台帳入力ツール.xlsm」** から開いて書き込む。

- UI は 2 実装: **VBA UserForms (基本)** / **PowerShell 5.1 + XAML/WPF (DPI 問題時のフォールバック)**
- 共通インターフェイス `ILedgerEventDialog` で切替 (設定シートの「UI 実装」)
- 検証・台帳行への展開・書込は UI 非依存の **LedgerCore** に一本化
- JSON (UTF-8 BOM 付き) は **VBA ⇔ PowerShell のプロセス境界でのみ** 使用。
  Forms 版はオブジェクトを直接受け渡す

## 画面構成 (設計合意 2026-07-20)

| 画面 | 生成される台帳行 |
|---|---|
| 搬入 (倉庫A→倉庫B) | 搬入 ×1 (品番プールへ。ロット指定なし) |
| 工程進行 / 発送 | 工程イベントまたは発送 ×1 (元状態は工程パターンから自動導出) |
| 不良発生 (+補填) | 不良発生 ×1 + 補填 ON なら充当 ×1 (未割当→当該ロット) |
| 例外操作 | 廃棄 / 倉庫A戻し / 余剰化 / 充当 ×1 |

「充当」「余剰化」は通常操作としては見せず、不良補填などの操作から自動生成する。
搬入時の割当行自動生成 (2026-07-15 決定) は **V_ シートの数式変更が前提のため今回は
未実装**。数式を触るフェーズで導入する。

## ファイル構成

| ファイル | 説明 |
|---|---|
| `build_xlsm.ps1` | 台帳入力ツール.xlsm を自動生成するビルドスクリプト |
| `Show-LedgerDialog.ps1` | PS/XAML 版ダイアログ本体 (PowerShell 5.1 用・UTF-8 BOM 付き) |
| `sample_input_defect.json` / `sample_output_defect.json` | JSON スキーマのサンプル |
| **VBA 側** (すべて Shift_JIS / CRLF。標準モジュール = `m_`、クラスモジュール = `c_` 接頭辞) | |
| `m_LedgerCore.bas` | 検証・展開・書込・再計算・保護/解除・保存 (中核・UI 非依存) |
| `m_LedgerDialogs.bas` | ボタン用エントリマクロ + ファクトリ + 設定読込 + SmokeTest |
| `m_LedgerFormUtil.bas` | UserForm 共通ヘルパー |
| `m_JsonLite.bas` | 内部実装用: 簡易 JSON ヘルパー + UTF-8 I/O (c_PsLedgerEventDialog 専用) |
| `c_ILedgerEventDialog.cls` | ダイアログ共通インターフェイス |
| `c_FormLedgerEventDialog.cls` / `c_PsLedgerEventDialog.cls` | 実装 (Forms 版 / PS 版) |
| `c_LedgerOperation.cls` | ダイアログの出力 (操作レベル) |
| `c_DialogContext.cls` / `c_ItemInfo.cls` / `c_LotInfo.cls` / `c_StockSnapshot.cls` | ダイアログへの入力 (選択肢+在庫スナップショット) |
| `c_LedgerDialogResult.cls` | ダイアログの戻り値 (Status / Message / Operation) |
| `*Form.formcode.txt` ×4 | UserForm のコード (コントロールはビルド時にデザイン生成、.frx 不要) |

## ビルド手順

1. Excel のオプションで「**VBA プロジェクト オブジェクト モデルへのアクセスを信頼する**」
   を有効にする (ファイル > オプション > トラスト センター > トラスト センターの設定 >
   マクロの設定)。**ビルド時のみ必要**。終わったら無効に戻してよい。
2. 実行:
   ```
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File build_xlsm.ps1
   ```
3. `台帳入力ツール.xlsm` が生成され、SmokeTest (フォーム生成+導出ロジックの疎通) まで
   自動確認される。

手動で作る場合: 空の .xlsm に .bas/.cls をインポートし、UserForm を 4 つ作成して
`*.formcode.txt` のコードを貼り付け、フォームコードのコメントにあるコントロール名で
コントロールを配置する (ビルドスクリプトの Layout 定義が座標の正)。

## 使い方

1. 「設定」シートで台帳パスを確認 (既定: E:\LocalAppsWorks\InventoryManager\在庫管理台帳.xlsx)。
   Forms の画面が DPI で縮小される PC では「UI 実装」を `PowerShell` に切り替える。
2. 「操作」シートのボタンから入力。流れ:
   **ボタン → (台帳を開く・ReadOnly 検知 → 再計算 → 選択肢スナップショット) →
   ダイアログ → 検証 → 確認 → 書込 → 再計算 → 負在庫チェック → 保護 → 保存**
3. 台帳の行を直接修正したいときだけ「保護解除」→ 修正 →「再保護 + 保存」。

- 記録者リストは台帳の **M_リスト G 列** で管理する (必須・選択式。前回値が既定)。
- 実施日の未来日は警告のうえ続行可。補填数量の上限は不良数量に固定。
- 台帳が読み取り専用で開いた場合 (他の人が編集中) は書き込まずに中断する。
- 保存時にシート保護+ブック構造保護+読み取り専用推奨フラグを自動で掛ける
  (パスワードは `LedgerCore.PROTECT_PWD`。既定 "ledger"。事故防止用であり
  意図的な改変は防げない。本命の保護は SharePoint の閲覧権限で行うこと)。
- 書込後に V_在庫内訳へ**新たな**マイナスが発生した場合 (他の入力との競合など) は
  警告し、その場で書込行を削除して取り消せる。既知のサンプル不整合
  (T_台帳 4〜7 行目によるロットA 未処理 −3) は「既存の負」として除外される。

## JSON スキーマ (schemaVersion 1.0)

受け渡しファイルの既定: `%TEMP%\ledger_dialog_in.json` / `ledger_dialog_out.json`
(UTF-8 BOM 付き。`PsLedgerEventDialog` のプロパティまたは設定シートの
「JSON フォルダ」で変更可)。出力は起動前に削除し、終了時に無条件上書きで必ず作成。
`schemaVersion` のメジャー不一致は PS 側がエラー (終了コード 2)。

### 入力 (Excel → PS): 選択肢スナップショット

| キー | 説明 |
|---|---|
| `schemaVersion` | "1.0" |
| `dialog` | `carryIn` / `progress` / `defect` / `exception` (表示する画面) |
| `defaults.actionDate` / `defaults.recorder` | 既定値 (日付は `yyyy-MM-dd`) |
| `recorders` / `states` | 記録者リスト / 状態リスト (表示順) |
| `items[]` | `itemNo` / `itemName` / `pattern` (1〜4) |
| `lots[]` | `lotId` / `itemNo` / `shipWeek` / `dest` / `required` / `planStatus` (キャンセル含む全ロット) |
| `stock.lots[]` | `lotId` + `byState` (非ゼロ状態のみ。負もあり得る) |
| `stock.unallocated[]` | `itemNo` + `byState` (同上) |

### 出力 (PS → Excel): 操作結果

`status` (`ok`/`cancel`/`error`) + `message` + `operation` (全キー常時出力):
`kind` / `eventKind` / `itemNo` / `lotId` / `fromState` / `qty` / `actionDate` /
`recorder` / `note` / `refill` / `refillFromState` / `refillQty` / `targetLotId`

**運ぶのは操作レベルのみ。** 台帳行への展開と正の検証は VBA 側 LedgerCore が行う
(PS 側の上限制御は入力ガイド)。

### 終了コード (xaml_ui と同じ)

| コード | 意味 |
|---|---|
| 0 | OK |
| 1 | キャンセル |
| 2 | エラー (入力ファイルなし、スキーマ不一致など) |

## テストモード (引数なしで起動)

```
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File Show-LedgerDialog.ps1 [-TestDialog defect]
```

組み込みテストデータ (サンプルブック相当。ロットA の未処理 −3 も再現) で単体テスト
できる。結果はメッセージボックス表示+ `ledger_dialog_test_out.json` に出力。
`-TestDialog` で `carryIn` / `progress` / `defect` / `exception` を指定 (既定 defect)。

## VBA コーディング規約

- モジュール名は接頭辞付き: 標準モジュール = `m_`、クラスモジュール = `c_`
  (インターフェイスのクラスモジュールも `c_ILedgerEventDialog`)。
- **Public メンバーの呼び出しは必ずモジュール名で修飾する**
  (例: `m_LedgerDialogs.RunLedgerDialog LedgerOpKind.opCarryIn`)。
  Private メンバーはモジュール外に公開されず修飾参照できないため、無修飾で呼ぶ。
- 列挙値は `LedgerOpKind.opCarryIn` のように列挙型名で修飾する。
- 初期化が単純な変数は `Dim x As T: Set x = ...` のワンライナーで宣言する。
- `Call` キーワードは使わない (冗長キーワードとして現在は非推奨のため)。

## 規約 (xaml_ui を踏襲)

- VBA ソースは **Shift_JIS / CRLF** を維持 (UTF-8 や LF のままだとインポートで壊れる)。
- PS は **5.1 + `-STA`** で起動。スクリプトは UTF-8 BOM 付き。
- **XAML にイベント属性 (`Click=` 等) や `x:Class` を書かない**
  (コードビハインドがないため `XamlReader.Load()` が例外を投げる。イベントは
  コード側で `Add_Click` 等により登録)。`x:Name` の変更も不可 (`FindName` が参照)。
- `JsonLite` を利用側コードから直接呼ばない (PsLedgerEventDialog の内部実装)。
- ボタンの `IsDefault` / `IsCancel` が Enter / Esc を担う。

## 検証ルール (LedgerCore.ValidateOperation)

- **エラー (書込拒否)**: 数量 < 1 / 記録者未選択 / 品番・ロット不明 /
  キャンセルロットへの工程進行・不良・充当 / 工程パターン外イベント /
  補填数量 > 不良数量 / **在庫負防止** (展開行をスナップショットに順次適用し、
  消費が可用量を超えたら拒否。既に負のバケットは可用 0 扱い)
- **警告 (確認のうえ続行可)**: 実施日が未来日
- 元状態は入力させず導出する: 工程進行 = 工程パターンの連鎖から / 発送 =
  発送準備完了 / 廃棄 = 不良。不良発生と例外操作は在庫のある状態のみ選択肢に出す。
