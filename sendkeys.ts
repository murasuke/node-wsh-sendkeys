const clipboardy = require('clipboardy');
const sleep = require('sleep');
require('winax');

//  Shell関連の操作を提供するオブジェクトを取得
const sh = new ActiveXObject('WScript.Shell');
// メモ帳を起動
sh.Run('notepad.exe');
sleep.msleep(1000);

// 文字を送信
sh.SendKeys('abcde');     sleep.msleep(100);
sh.SendKeys('{ENTER}');   sleep.msleep(100);
sh.SendKeys('{1 5}');     sleep.msleep(100);
sh.SendKeys('{TAB 3}');   sleep.msleep(100);

// 日本語は文字化けしてしまうため、一旦クリップボードへコピーしてから貼り付ける
clipboardy.writeSync('あいうえお😵😎'); sleep.msleep(100);
sh.SendKeys('^v'); 

console.log('終了');
