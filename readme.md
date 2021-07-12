# node.jsでsendkeys(キーコードの送信)の動作確認

## 前書き

昔々VBScript(Windows Scripting Host)では、比較的簡単に自動入力スクリプトを作ることができました。
```vbscript
Set objShell = WScript.CreateObject ("WScript.Shell")
sh.SendKeys("abcde")
```

今でも現役で使えますが、node.jsと比較すると環境面での非力さは否めません。
そこで、node.js環境側でSendKeysと同じことができるか確認してみます。

1. javascriptで、JScript(WSH)互換スクリプトを目指すパターン

1. TypeScriptで[`WScriptエミュレーション`](https://github.com/durs/node-activex#wscript)環境を利用しないパターン


## 1. javascriptで、JScript(WSH)互換スクリプトを目指すパターン
 　
  ～～**winaxの[`WScriptエミュレーションモード`](https://github.com/durs/node-activex#wscript)を利用**～～

* 利用したライブラリと実行方法

WScriptはランタイム環境のため、エミュレーションのために`nodewscript`から実行する必要があります。


package.json(抜粋)
```json
{
   "scripts": {
    "dev-js": "nodewscript sendkeys.js"
  },
  "dependencies": {
    "winax": "^3.1.5"
  }
}
```

* 動作確認ソース

  * メモ帳を起動して、sendkeysで書き込みを行います
  * 日本語の文字は直接入力できないため、clipboardを経由して貼り付けます(vbsでも同じ)
  * Unicodeの「顔文字」は文字化けします(cmd.exeのコピー機能を利用しているため。codepageを変えればうまくいくかもしれませんが、そこまで確認はしていません)

sendkeys.js
```javascript
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

// 日本語は文字化けしてしまうため、一旦クリップボードへコピーしてから貼り付け(顔文字は文字化けします)
sh.Run('cmd.exe /c Echo あいうえお😵😎 | Clip', 0, true);
sh.SendKeys('^v'); 

WScript.Echo('終了');
```

## 2. TypeScriptで`WScriptエミュレーション`環境を利用しないパターン
 　～～**ts-node上で実行(WScriptオブジェクト未使用)**～～
コヒペのため、clipboardyを利用しています。他いくつか試しましたが、顔文字まで正しく処理できたのはこれだけでした。
package.json(抜粋)
 ```json
 {

  "scripts": {
    "dev-ts": "ts-node sendkeys.ts"
  },
  "dependencies": {
    "@types/windows-script-host": "^5.8.3",
    "clipboardy": "^2.3.0",
    "sleep": "^6.3.0",
    "ts-node": "^10.1.0",
    "winax": "^3.1.5"
  },
  "devDependencies": {
    "@types/node": "^16.3.1",
    "typescript": "^4.3.5"
  }
}
```

* 動作確認ソース

  * メモ帳を起動して、sendkeysで書き込みを行います
  * 日本語の文字は直接入力できないため、clipboardを経由して貼り付けます(vbsでも同じ)
  * Unicodeの「顔文字」も利用できます(cmd.exeのコピー機能ではなく、[npm - clipboardy](https://www.npmjs.com/package/clipboardy)を利用)
  * `WScript`を利用できないため、[sleep](https://www.npmjs.com/package/sleep)を利用

```typescript
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

```
