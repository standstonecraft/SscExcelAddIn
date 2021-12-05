# SscExcelAddIn

自分の業務効率化のための、小さなExcelアドインです。  

## 動作確認環境

- Windows 10 Home 64bit
- Microsoft Excel 2019 MSO Version 2110 Build 16.0.14527.20270 64bit
- .NET Framework 4.8

## 機能とTODO

- [x] シート編集
  - [x] 高度な置換 (.\*アイコン)
    - [x] 正規表現での置換
    - [x] 半角変換
    - [x] 連続置換
    - [x] 置換結果プレビュー
    - [x] 序列を表す文字列[^1]のインクリメント
    - [x] 序列を表す文字列の連番追加と振り直し
    - [x] 序列を表す文字列の文字種変換
  - [x] 行(列)交互選択 (シマウマアイコン)
  - [x] 空行(列)削除
- [x] 図形編集
  - [x] 図形テキスト編集 (図形アイコン)
    - [x] セルから図形テキストに書き込み
    - [x] セルを参照する数式として埋め込み
    - [x] 図形テキストのセルへの書き出し
    - [x] 図形テキスト検索
  - [x] 図形サイズ変更 (矢印アイコン)
- [ ] データ集計表生成
- [ ] 背景色置換
  - [ ] 白背景解除
- [ ] 目次シート生成
- [ ] 罫線コマンド
- その他
  - [x] 更新チェック機能

[^1]: (([全半]角|丸囲み|ローマ)数字|[大小]文字アルファベット|[全半]角カタカナ)を指す。
