# pinet_inputの使い方



## 概要

PI-NETへの入力を補助するショートカットキーを提供します。

1. 詳細説明<br>
   <kbd>F1</kbd>
2. ⏯ PI-NETの図番をエクセルで検索します。<br>
   <kbd>無変換+f</kbd>
3. このショートカットキー機能を終了させます<br>
   <kbd>無変換+q</kbd>
4. 🅿️ チェックボックス全てオンはF1詳細キーで詳細を確認してください
<br>
<br>


## まず始めに

1. `pinet_input.exe` を実行する
2. タスクバーの右下にHアイコンが現れる<br>
   ![exeIcon](https://user-images.githubusercontent.com/69337126/91275802-9735af80-e7bb-11ea-9a33-1dc72ec7c4df.png)
<br>
<br>


## 作業が終了したら

作業が終了したら`pinet_input.exe`を終了させます。

終了させるには、２種類の方法があります。

1. <kbd>無変換とqを同時押し</kbd>する
2. Hアイコンを右クリックして<kbd>Exit</kbd>をクリックします。<br>
   ![exit](https://user-images.githubusercontent.com/69337126/91275803-97ce4600-e7bb-11ea-87ac-194ed9fe9444.png)
<br>
<br>


## ⏯ PI-NETに員数を半自動入力する手順
### エクセルの初期設定

まず初めにエクセルの検索設定を以下のように行ってください。<br>
![excelFind](https://user-images.githubusercontent.com/69337126/91275804-9866dc80-e7bb-11ea-8b93-0b05b888574c.png)


PI-NET → エクセルで検索 → PI-NETに入力する手順は以下のとおりです。

手順
1. 検索対象のエクセルファイルを開いておく。
   (対象以外のエクセルファイルは全て閉じておいて下さい)
2. PI-NETの支給部品列の部品番号をダブルクリックして版番号全体を選択状態にする。
   (このとき、空白文字も一緒に選択される場合がありますが問題ありません)
3. <kbd>無変換とfを同時押し</kbd>するとエクセルで検索します。
4. 検索がヒットしたら再度<kbd>無変換とfを同時押し</kbd>するとPI-NETに員数を入力します。
   1. エクセルの員数セルには色が塗られます
   2. PI-NETは部品番号をダブルクリックした位置を基準に自動操作するので、別のところをクリックしないよう注意してください
5. 上記1~4を繰り返す
<br>
<br>

## 🔍 選択した部品番号をエクセルで検索 →  PI-NETに入力


## 🅿️ PI-NETのチェックボックスのオンオフ

手順は以下のとおりです。
### 1. 以下のコードをブックマークレットにしてブックマークバーなどに登録してください。
```javascript
javascript:(function (){var inputs = document.getElementsByTagName('input');for(var i=0; ; i++){for (var j=0; j < inputs.length; j ++) {var e = inputs[j];if (e.type == 'checkbox')e.checked = true;}if(i < window.frames.length){try {inputs = window.frames[i].document.getElementsByTagName('input');}catch(e){}}else{break;}}})();
```
[チェックボックスオン](<a href="javascript
javascript:(function (){var inputs = document.getElementsByTagName('input');for(var i=0; ; i++){for (var j=0; j < inputs.length; j ++) {var e = inputs[j];if (e.type == 'checkbox')e.checked = true;}if(i < window.frames.length){try {inputs = window.frames[i].document.getElementsByTagName('input');}catch(e){}}else{break;}}})();">チェックボックスオン</a>)

![bookmarklet](https://user-images.githubusercontent.com/69337126/91371605-6e112f80-e84c-11ea-91ae-27543dffd045.png)

### 2. 登録したブックマークをクリックするとチェックボックスが全てオンになります

![bookmarkButton](https://user-images.githubusercontent.com/69337126/91371746-d2cc8a00-e84c-11ea-8c22-545aa57dee60.png)