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
