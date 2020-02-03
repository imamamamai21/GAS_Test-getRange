# 「GASでスプシの行をスマートに指定する方法」サンプル用コード
[Qiita]()にて書いた記事のサンプルコードです。

## はじめに
GAS(google Action Script)はとても便利ですね。
大量のデータを扱った自動化などで日々お世話になっております。
でも足りない機能・頼りない機能がとても多く、自前実装で賄っている方も多いかと思います。

# GASの頼りない仕様
GASを使っていて最初「行を変えたらコード動かなくなるやん」と不安になった方は多くいるかと思います。

```js
// これらは全て行を変えたら動かなくなる
sheet.getRange('A2:A10').getValue();
sheet.getRange(0,0,1,1).getValue();
sheet.getRange('A1').getValue();
```

これを列名や数字などの指定ではなく、キーワードで指定したいというのが今回のお話です。

# アルファベットじゃなくてキーワードで行を検出したい
実装したファイルはこちらに上がっています
Git

テスト用のスプレッドシードファイルはこんな感じのものを作りました
<img width="589" alt="スクリーンショット 2020-01-31 16.23.44.png" src="https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/80770/b1594722-1a18-4e95-4785-e717433f5ba7.png">

## 1. 通常の取得の仕方と問題
`A列`というタイトルの1番目の情報を取得してみましょう。

```js
var MY_SHEET_ID = '123a-AAa1a1AAaAAAAaaaaaaaAaAAaaaAAAaaaa1aA1'; // シートのID

function myFunction() {
  var sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('シート1');
  var a1Data = sheet.getRange('A1').getValue();
  Logger.log(a1Data); // A2が取れる
}
```

ここで、シートのA列に１つ列を追加してみます
<img width="678" alt="スクリーンショット 2020-01-31 16.29.31.png" src="https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/80770/82c87abe-5001-1cd9-0bdd-7f3f344bf07b.png">

すると先程のコードで取得できる値が `A2`から`X2`に変わってしまいました。
コードを書いた人が予期せぬところで、こういった編集をされてしまうとコードはうまく動かなくなるかと思います。

こんな問題をなくすために、スプレッドシートの列を追加しまくってもきちんと動くコードを作りましょう。

## 2. KeyWordを使って、データの取得・書き換えを行う
### 1. シートのインスタンスを作成


```js:TestSheet.gs
var TestSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('シート1');
  this.values = this.sheet.getDataRange().getValues(); // 全てのデータを２次元配列取得し、保存
  this.titleRow = 0;
  this.index = {};
  
  /**
   * 一行目(タイトル行)の中身をチェックして、タイトルに一致する行数をobjectに保管する
   */
  this.createIndex = function() {
    const TITLE = 'A列';
    var me = this;
    var filterData = (function() {
      for(var i = 0; i < me.values.length; i++) {
        if (me.values[i].indexOf(TITLE) > -1) {
          me.titleRow = i + 1; // タイトルのある列数を念の為保存
          return me.values[i];
        }
      }
    }());
    if(!filterData || filterData.length === 0) { // 一致するタイトルがないときはエラー
      this.showTitleError(TITLE);
      return;
    }
    this.index = {
      a: filterData.indexOf(TITLE),
      b: filterData.indexOf('B列'),
      c: filterData.indexOf('C列'),
      d: filterData.indexOf('D列'),
      e: filterData.indexOf('E列'),
      f: filterData.indexOf('F列'),
      x: filterData.indexOf('X列')
    };
    return this.index;
  }
  
  function showTitleError(key) {
    Browser.msgBox('データが見つかりません', '表のタイトル名を変えていませんか？ : ' + key, Browser.Buttons.OK);
  }
}
  
TestSheet.prototype = { // 外部呼び出しのものは一応protptypeにしておく
  getRowKey: function(target) {
    var targetIndex = this.getIndex()[target];
    var alfabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
    var returnKey = (targetIndex > -1) ? alfabet[targetIndex] : '';
    if (!returnKey || returnKey === '') this.showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  }
};

var testSheet = new TestSheet();
```

### 2.testSheetの中身を取得する

```js:Main.gs
function myFunction() {
  /*var sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('シート1');
  var a1Data = sheet.getRange('A1').getValue();
  Logger.log(a1Data)
  ↑Before   ↓After */  
  var index = testSheet.getIndex();
  var a1Data = testSheet.values[1][index.a];
  Logger.log(a1Data); // A2が取得できる
}
```

### 3.testSheetの中身の書き換えをする
こんなかんじに最終行に特定の文字を書き込んでみます。
![スクリーンショット 2020-02-03 12.15.31.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/80770/f484ff68-b06c-8d88-45f4-8a6964335ace.png)

```js:Main.gs
function edit() {
  var key = testSheet.getRowKey('a'); // return 'B'
  var lastRow = testSheet.sheet.getRange(key + ':' + key).getValues().filter(String).length + 1; // return 15
  // `A列`列の最終行に特定の文字を書き込み
  testSheet.sheet.getRange(key + lastRow).setValue('A列の最後行に書き換え');
  // testSheet.values[lastRow - 1][testSheet.getIndex().a]) = 'A列の最後行に書き換え'になっている
}
```

# さいごに
方法というか私はこうやっていますといういち例です。
いろいろなやり方があるとは思うのですが、「ある程度無秩序にシートを編集してしまっても動く」ということに重点を置いた結果、一番きれいな方法としてこれに落ち着きました。

こんな方法あるよ〜
こここうしたらもっと良くなるよ
というご意見ございましたらお気軽にコメントくださいませm(_ _)m
