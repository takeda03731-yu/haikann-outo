# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 概要

Excel（.xlsm）の配管データから、ezdxf で DXF 形式の配管長略図を生成し、JWCAD で開く業務ツール（上水道配管工事の工事日報用）。テスト・lint・ビルドは存在しない。

## 実行

```
pip install -r requirements_windows.txt   # ezdxf, pandas, openpyxl
python main.py
```

- `main.py` はモジュールトップレベルで全処理を実行するスクリプト（関数・`__main__` ガードなし）。import するとそのまま実行される。
- 入力: `constants.py` の `EXCEL_FILE_NAME`。`os.path.join(data, EXCEL_FILE_NAME)` なので、絶対パスならそのパスが優先される（現状はデスクトップの絶対パス）。サンプルは `data/配管長略図.xlsm`。
- 出力: `C:\JWW\配管長略図\text_sample.dxf` に固定出力（`main.py` 末尾）。エンコーディングは `cp932`、DXF は R2010。
- 実行前にユーザー環境のExcelパスが存在するか確認すること。Excel は読み取り専用で扱う。

## アーキテクチャ

データの流れ: `components.excel_to_dict`（Excel4行目をヘッダーとして読込、材料No・測点・名称・方向・配管長・水平距離・摘要・施工日付の辞書リスト化。材料Noの次行の空白行にある配管長2/水平距離2/測点2も統合）→ `filter_by_construction_date_with_previous`（施工日付で絞込。直前の材料も含む。`継続元` フラグ）→ `main.py` のループが 1 材料ずつ描画。

- **`main.py`（約4500行）**: 材料ごとに `名称`（部分文字列一致: 「不断水割」「副弁付」「GX」「NS」「PE」「45°」「下がり」等）と `方向` コード（SF/FS/AA/AB/BA/BB など）の巨大な if/elif 連鎖で、`components` の部品関数を選んで呼ぶ。
  - 状態はグローバル変数で持つ: `x_count`（次の部品のX位置）, `y_count`, `y_offsetcount`（測点文字の段ずらし）, `flat_sum`/`haikann_sum`（水平距離・配管長の累計）, `today_count`。
  - 縦曲り（上がり/下がり）は直前1〜3件の名称（`prev1..3`）から `down_45`, `up_22` 等のフラグを作り、`st`（開始）/`ed`（終了）部品を切り替える。前後の材料に依存するので、ループ内の順序を崩さないこと。
  - 各ブロックはほぼ同じ後処理（寸法線の `add_line`、`flat_line1`、累計加算）をコピペで持つ。
- **`components.py`（約5400行）**: Excel読込関数 + 部品描画関数。シグネチャは `(msp, [item,] offset_x=0, offset_y=0)`。`item` を取るのは寸法（配管長）文字が必要な直管系（`*_kou_*`=甲切管, `*_otsu_*`=乙切管）。
  - 接頭辞で管種: `gx_`=GX形ダクタイル鋳鉄管, `ns_`=NS形, `pe_`=PE管, `meka_`=メカニカル, `soft_`=軟弁。
  - 語: `tyokkann`/`tankan`=直管/短管, `katauke`/`ryouuke`=片受/両受, `vb`=縦曲り, `hb`=横曲り, `st`/`ed`=縦曲り開始/終了, `raina`=ライナ。末尾は方向コード。
  - `ab`/`ba`、`up`/`down`、`gx`/`ns` で形が似た関数が大量にあり、重複が多い。
  - `choose_zairyou_no`、`name_sokutenn`、`flat_line*`、`haikann_line*` が材料No・測点・寸法線の補助描画。
- **`constants.py`**: Excelパスのみ。

## 注意点

- **座標は非常に大きな値**（例: `x_count = -160000000`、文字高さ `6000000`、部品幅 `39000000`）。ezdxf/JWCAD 側の単位換算を前提としており、部品関数の寸法も同じスケール。新規部品は既存関数の数値をコピーして合わせる。値を「きれいな mm」に直さないこと。
- 名称は Excel の固定語彙に依存する。未対応の名称は描画されない（または分岐の途中で抜ける）ので、新規材料は `components.py` に部品関数と `main.py` の分岐の両方が必要。
- `main.py` 冒頭の `haikann` / `haikann2` はテスト用の手書きデータで、現状は使われていない。
- リファクタ時は同じ入力で DXF を生成し、前後でエンティティ（種別・座標・文字）を比較して出力が変わらないことを確認する。
- ファイル削除は必ずユーザーに確認する。

## 専用エージェント

`.claude/agents/jwcad-zumen-improver.md`: このコードの改善（リファクタ・重複削減・バグ修正・部品追加）用サブエージェント。大きな変更はこれを使う。
