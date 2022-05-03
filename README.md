_このプラグインは実験的です_

# WinMerge CbzDiffPlugin

## 説明

.CBZ/.CBR/.CB7を一つの画像として抽出するWinMergeプラグインです。

## インストール

- WinMergeをインストールした先の、`MergePlugins`ディレクトリ内にDLLをコピーしてください。
- WinMergeをインストールする際、アーカイブサポートと有効にしていなければ[`7z.dll`をダウンロード](https://sevenzip.osdn.jp/download.html)し、`CbzDiffPlugin.dll`と同じディレクトリにコピーしてください。

## 注意事項

- メモリーを大量に使用するため、環境により動作しない可能性があります。
- WinMergeには[バイナリ比較切替閾値]があります。それを超えるサイズのファイルはプラグイン展開されません。
- 64bit版しか対応しません。

## ライセンス

このソフトウェアを使用することに対して、あらゆることは無保証です。

[LICENSE](LICENSE) を参照してください。
