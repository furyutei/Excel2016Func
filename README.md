このアドイン（Excel2016Func）について
======================================
配布者：©2018 [風柳](http://furyu.hatenablog.com/about)  [@furyutei](https://twitter.com/furyutei)  


これはなに？
---
[Excel 2013 以前ではサポートされていないワークシート関数](https://blogs.office.com/en-us/2016/02/23/6-new-excel-functions-that-simplify-your-formula-editing-experience/)  

- IFS  
- SWITCH  
- CONCAT  
- TEXTJOIN  
- MAXIFS  
- MINIFS  

を、ユーザー定義関数により、疑似的に使用できるようにするためのアドインです。  

なお、各関数の使い方や、ユーザー定義関数の実装については、  

- [関数リファレンス | 経理・会計事務所向けエクセルスピードアップ講座](https://www.excelspeedup.com/category/kansuu/)  
- [ユーザー定義関数：MAXIFS・MINIFS（Excel 2013以前向け）](https://gist.github.com/furyutei/ca02a52e564535e051f1d96eba390e8d)

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


Mac をお使いの場合
---
Mac の場合にはインストール／アンインストール用の VBScript は動作しません。  
適当なディレクトリにアドイン本体（Excel2016Func.xlam）をコピーした上で、  

> **プレインストールされている Excel のアドインを有効にするには**    
> 
> 1. [ツール] メニューの [アドイン] を選択します。
> 2. [有効なアドイン] ボックスで、有効にするアドインのチェック ボックスをオンにして、[OK] をクリックします。
> 
> **Excel アドインをインストールするには**   
> 
> - 一部の Excel アドインはコンピューターに保存されており、上記の [アドイン] ダイアログ ボックスの [参照] をクリックしてアドインを探し、[OK] をクリックすることでインストールまたは有効にすることができます。  
> 
> <cite>[Excel でアドインを追加または削除する - Office サポート](https://support.office.com/ja-jp/article/excel-%E3%81%A7%E3%82%A2%E3%83%89%E3%82%A4%E3%83%B3%E3%82%92%E8%BF%BD%E5%8A%A0%E3%81%BE%E3%81%9F%E3%81%AF%E5%89%8A%E9%99%A4%E3%81%99%E3%82%8B-0af570c4-5cf3-4fa9-9b88-403625a0b460#OfficeVersion=Mac)  </cite>

を参考に、直接インストールしてみてください。  


免責事項など
---
無償で利用できますが、無保証です。ご利用の際には全て自己責任でお願いします。  
使用した結果等により何らかの不都合が発生した場合でも、一切関知いたしません。  

再配布の際には、  

> [関数リファレンス | 経理・会計事務所向けエクセルスピードアップ講座](https://www.excelspeedup.com/category/kansuu/)  

へのリンクをお願いします。  

なお、当方(風柳)は、[こちらの各記事](https://www.excelspeedup.com/category/kansuu/)に掲載されたユーザ定義関数については、まとめて利用しやすくしただけであり、各関数に関して動作検証などは実施しておりません。  
[一部の関数（MAXIFS・MINIFS）については実装を行いましたが](https://gist.github.com/furyutei/ca02a52e564535e051f1d96eba390e8d)、簡単な動作確認を行ったのみです。  


謝辞
---
- [はけた(羽毛田　睦土) 様](https://www.excelspeedup.com/) [@excelspeedup](https://twitter.com/excelspeedup)  
    > [関数リファレンス | 経理・会計事務所向けエクセルスピードアップ講座](https://www.excelspeedup.com/category/kansuu/)  
    
    にある[ユーザー定義関数を使用させていただきました](https://twitter.com/excelspeedup/status/968806992029433857)。  



- [fnya 様](http://fnya.cocolog-nifty.com/blog/) [@fnya](https://twitter.com/fnya)

    > [VBScript で Excel にアドインを自動でインストール/アンインストールする方法: ある SE のつぶやき](http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html)

    にある[インストール／アンインストール用スクリプトを使用させていただきました](https://twitter.com/fnya/status/968810606793973760)。  
