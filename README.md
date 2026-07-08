# errorcheck
対象データから異常値を判定し、出力用のシートへ出力する処理を自動化した異常値検出マクロ。

（errorcheck 概要）

対象データから異常と判定できる値が存在する場合、異常値が含まれる行を抽出し、

異常個所と異常内容を記載したデータを行単位で出力する。

異常判定対象のデータを1行単位で走査し、仮に同じ1行に複数の異常値が存在する場合は

異常値の数に相当する行数分のデータを出力する。

(errorcheck 実行手順)

1, このリポジトリに添付されている拡張子「.xlsx」で終わる各エクセルファイルをマクロ実行時に使用する 任意のフォルダに格納する。

2, このリポジトリに添付されている「errorcheck_GitHub.xlsm」ファイルをダウンロードし、名前をつけて 保存。

3, 2で保存したマクロファイルを開き、コンテンツ有効化してから「Alt + F8」

4, マクロ一覧が表示されるので、実行したいマクロを選択し実行。

「VBAコード詳細は別添付ファイル参照。」


（マクロ実行前後やその他スクリーンショット掲載）

エラー検出対象シート
<img width="1920" height="1080" alt="スクリーンショット (12)" src="https://github.com/user-attachments/assets/8ba9797f-7b96-45bc-9f99-9aa6158948bc" />


マクロ実行前の出力用シート
<img width="1920" height="1080" alt="スクリーンショット (10)" src="https://github.com/user-attachments/assets/cc6ebe70-af38-44a6-bfb1-cdc75c2cc6aa" />


マクロ実行画面
<img width="1920" height="1080" alt="スクリーンショット (11)" src="https://github.com/user-attachments/assets/b765f236-c166-4120-9959-9f975e3b1b1c" />


マクロ実行後の出力用シート
<img width="1920" height="1080" alt="スクリーンショット (13)" src="https://github.com/user-attachments/assets/e1d205d0-67c3-43bd-8180-33e97067198e" />
