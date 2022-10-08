# PI-Net 実査数入力

PI-Netの実査数入力を自動化します。



## 使用方法

1. main.pyと同じ階層にesedgedriver.exeを配置する
   [Microsoft Edge WebDriver - Microsoft Edge Developer](https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/)
2. main.pyと同じ階層に.envファイルを作成し、以下を記述する
   URL=対象のURL
   INIT_DIR=規定で開くフォルダ
   USER_NAME=pi-netのユーザー名
   PASSWORD=pi-netのパスワード
3. main.pyを実行する
   対象の実査登録確認リスト(xlsx)をダイアログで選択する



## 備考

- 有償 or 無償によって自動的に画面を切り替えます。

- 実査入力画面は99ページまでしか表示されません。99ページに到達したらいったん終了します。続きを入力するには再度main.pyを実行してください。

- PI-Netに数値を入力したらチェックボックスにチェックを入れます。

- PI-Netに数値を入力したら実査登録確認リスト(xlsx)の該当セルを黄色く塗ります。

  
