function createpwd() {
  //英数字を用意する
  var letters = 'abcdefghijklmnopqrstuvwxyz';
  var numbers = '0123456789';
  
  var string  = letters + letters.toUpperCase() + numbers;
  //toUpperCase()  小文字を大文字に変換
  
  var len = 12;　　　//8文字
  var password=''; //文字列が空っぽという定義をする
  
  for (var i = 0; i < len; i++) {
    password += string.charAt(Math.floor(Math.random() * string.length));
    // charAt メソッドを用いて文字列から指定した文字を返す。
  }
  
  return password
}
