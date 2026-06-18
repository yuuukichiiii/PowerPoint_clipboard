// ショートカットから呼ばれる関数を登録する
Office.onReady(function () {
  // 表示
  Office.actions.associate("showTaskpane", function () {
    return Office.addin
      .showAsTaskpane()
      .then(function () {
        return;
      })
      .catch(function (error) {
        return error.code;
      });
  });

  // 非表示
  Office.actions.associate("hideTaskpane", function () {
    return Office.addin
      .hide()
      .then(function () {
        return;
      })
      .catch(function (error) {
        return error.code;
      });
  });
});
