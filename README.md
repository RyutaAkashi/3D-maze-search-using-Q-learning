Implemented in 2021.06-2021.06
# 3D maze search using Q-learning
Find the shortest path in a 3D maze using reinforcement learning. The learning method is Q-learning, created using Excel VBA.

# Q学習を用いた3次元迷路の探索
## 目的
2次元の問題解決は比較的容易であるが、3次元空間における問題解決は困難を伴う。本プログラムは、人間が直感的に理解しにくい3次元空間での経路最適化を目指す。

## 手法
Excelシート上に3次元迷路を構築し、VBAを用いてQ学習アルゴリズムを実装することで、3次元迷路の最短経路を探索する。

## 実装詳細
### Excelシート構成
* **A1:U7（操作パネル）**: 図形に登録されたマクロにより、シートの更新状況を確認しながら各操作を実行できる。
* **A列:U列（3次元迷路）**: 21×21×17サイズの迷路であり、階層ごとに記述される。スタート地点は2階左上のマス、ゴール地点は11階右下のマスに設定。セル値0（灰色）は壁、1（白色）は通路を示す。
* **Y列:AE列（状態推移）**: 各部屋（マス）に隣接するマスの番号が入力される。
* **AG列:AM列（Q初期値）**: 移動先の各マスに対する学習前のQ値が入力される。
* **AO列:AU列（Q最終値）**: 移動先の各マスに対する学習後のQ値が入力される。
* **2807行:5807行（状態履歴）**: 学習の各エピソードにおけるステップ数と状態の履歴が記録される。

### VBAマクロ機能
1.  **`1_reset`**: シート上の全データ（迷路を含む）を初期化する。
2.  **`2_MazeGeneration`**: 3次元迷路を自動生成する。生成アルゴリズムは[棒倒し法](https://news.mynavi.jp/article/nadeshiko-61/)を参考にしている。
3.  **`3_InputRoomNumber`**: 迷路内の通路部分（空白マス）に番号を割り振る。
4.  **`4_InputTransition`**: 生成された迷路に基づき、各マスからの移動可能なマスを判定する。移動先が壁の場合は0、通路の場合は対応するマス番号を記録する。
5.  **`5_InputQInitialValue`**: 状態推移の情報に基づき、Q値の初期値を設定する。移動先が壁の場合は-100、通路の場合は1桁のランダムな整数を割り当てる。
6.  **`6_StartQlearning`**: 設定されたQ初期値と以下のパラメータに基づき、Q学習を実行する。各エピソードの学習結果は状態履歴に記録される。
    * 割引率 ($\gamma$): 0.7
    * 学習率 ($\alpha$): 0.5
    * 学習回数: 3000エピソード
7.  **`7_ShowRoute`**: 学習完了後、得られた最短経路を迷路シート上に表示する。経路上のにあたるセルには-1が入力され、Excelの機能により黄色で強調表示される。

## 結果
21×21×17マスの3次元迷路において、最短ステップ数である45ステップの経路を発見することに複数回成功した。壁が多い場合でも47~51ステップでゴールにたどり着いている。なお、学習結果の経路が最短経路である証明はできていない。

## 実行デモンストレーション
本プログラムの実際の動作や、操作パネル、Q学習のプロセス、迷路や重みの入力、最終的な結果表示については、以下の動画で確認できる。
* [Q学習を用いた3次元迷路探索デモンストレーション](https://youtu.be/dtFFJx3qzhI)

学習前
![image]([https://github.com/user-attachments/assets/a20b02ff-8c30-4d73-b34f-b795bec3a95e](https://github.com/RyutaAkashi/3D-maze-search-using-Q-learning/blob/main/result/before.png))

学習後
![image]([https://github.com/user-attachments/assets/a20b02ff-8c30-4d73-b34f-b795bec3a95e](https://github.com/RyutaAkashi/3D-maze-search-using-Q-learning/blob/main/result/after.png))
