// winaxをロード
require('winax');

//  Shell関連の操作を提供するオブジェクトを取得
const sh = new ActiveXObject('WScript.Shell');

//  メモ帳を起動
sh.Run( "notepad.exe");
WScript.Sleep( 1000 );

// 文字を送信
sh.SendKeys('abcde');     WScript.Sleep(100);
sh.SendKeys('{ENTER}');   WScript.Sleep(100);
sh.SendKeys('{1 5}');     WScript.Sleep(100);
sh.SendKeys('{TAB 3}');   WScript.Sleep(100);

// 日本語は文字化けしてしまうため、一旦クリップボードへコピーしてから
sh.Run('cmd.exe /c Echo あいうえお😵😎 | Clip', 0, true);
sh.SendKeys('^v'); 

WScript.Echo('終了');
