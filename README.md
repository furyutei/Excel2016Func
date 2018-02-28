このアドオン（Excel2016Func）について
======================================
配布者：©2018 [風柳](http://furyu.hatenablog.com/about)  [@furyutei](https://twitter.com/furyutei)  


これはなに？
---
Excel 2013 以前ではサポートされていないワークシート関数の一部  

- IFS  
- SWITCH  
- CONCAT  
- TEXTJOIN  

を、ユーザー定義関数により、疑似的に使用できるようにするためのアドインです。  

なお、各関数の使い方や、ユーザー定義関数の実装については、  

> [関数リファレンス | 経理・会計事務所向けエクセルスピードアップ講座](https://www.excelspeedup.com/category/kansuu/)  

をご参照ください。


ファイルの内容
---

- Excel2016Func.xlam : アドイン本体  
- Install.vbs : インストール用のスクリプト  
- Uninstall.vbs : アンインストール用のスクリプト  
- README.md : このファイル


環境について
---

- Windows 10
- Excel 2010

においてのみ、動作を確認しております。  


インストール方法
---
1. Excel を起動している場合、いったん終了します。  
2. Install.vbs をダブルクリックし、指示に従います。  

Excel を起動し、「ファイル」→「オプション」→「アドイン」→[設定(G)」にて、「☑Excel2016Func」が存在し、チェックがついていれば、インストールは成功しています。  



アンインストール方法
---
1. Excel を起動している場合、いったん終了します。  
2. Uninstall.vbs をダブルクリックし、指示に従います。  

Excel を起動し、「ファイル」→「オプション」→「アドイン」にて、「Excel2016Func」が存在しなくなっていることを確認してください。  


免責事項など
---
無償で利用できますが、無保証です。ご利用の際には全て自己責任でお願いします。  
使用した結果等により何らかの不都合が発生した場合でも、一切関知いたしません。  

再配布の際には、  

> [関数リファレンス | 経理・会計事務所向けエクセルスピードアップ講座](https://www.excelspeedup.com/category/kansuu/)  

へのリンクをお願いします。  

また、当方(風柳)は、[こちらの各記事](https://www.excelspeedup.com/category/kansuu/)に掲載されたユーザ定義関数をまとめて利用しやすくしただけであり、各関数に関して動作検証などは実施しておりませんので、あしからず。  


謝辞
---
- [はけた(羽毛田　睦土) 様](https://www.excelspeedup.com/) [@excelspeedup](https://twitter.com/excelspeedup)  
    > [関数リファレンス | 経理・会計事務所向けエクセルスピードアップ講座](https://www.excelspeedup.com/category/kansuu/)  
    
    にあるユーザー定義関数を使用させていただきました。  



- [fnya 様](http://fnya.cocolog-nifty.com/blog/) [@fnya](https://twitter.com/fnya)

    > [VBScript で Excel にアドインを自動でインストール/アンインストールする方法: ある SE のつぶやき](http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html)

    にあるインストール／アンインストール用スクリプトを使用させていただきました。  
