# errorcheck
対象データから異常値を判定し、出力用のシートへ出力する処理を自動化した異常値検出マクロ。

（errorcheck 概要）

対象データから異常と判定できる値が存在する場合、異常値が含まれる行を抽出し、

異常個所と異常内容を記載したデータを行単位で出力する。

異常判定対象のデータを1行単位で走査し、仮に同じ1行に複数の異常値が存在する場合は

異常値の数に相当する行数分のデータを出力する。

（マクロ実行前と実行後のシートのスクリーンショット）

マクロ実行前の出力用シート
<img width="1920" height="1080" alt="取得したファイルのシート1" src="https://github.com/user-attachments/assets/f0ccea90-4e93-4492-bf2f-6a4d24070843" />

マクロ実行画面
<img width="1920" height="1080" alt="取得したファイルのシート2" src="https://github.com/user-attachments/assets/75ec32cb-830c-4b93-b0b3-010305a6bebe" />

マクロ実行後の出力用シート
<img width="1920" height="1080" alt="取得したファイルのシート3" src="https://github.com/user-attachments/assets/b676c577-0303-468d-a439-9357d0f49f72" />


<img width="1920" height="1080" alt="取得したファイルのシート4" src="https://github.com/user-attachments/assets/3f7c136f-cc7a-463e-998c-2d00555655df" />
