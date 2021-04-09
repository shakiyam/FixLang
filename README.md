FixLang
=======

FixLangはPowerPointのプレゼンテーション資料の校正言語を統一します。

PowerPointでプレゼンテーション資料を作成する際、特に外国語の資料を和訳する際に、日本語の文が他の言語と判断されてスペルチェックに引っかかってしまうことがあります。
FixLangはこの問題を解決します。FixLangはPowerPointのプレゼンテーションの校正言語を日本語に修正します。

使い方
------

FixLangコマンドに引数としてPowerPointファイルを指定します。

```console
FixLang PPT_FILE...
```

* FixLangは指定されたファイルを順次バックアップしてから開き、言語を修正して保存します。(`sample.pptx` は `sample - backup.pptx` にバックアップされる)
* 修正状況は画面に出力されるだけではなく、PowerPointファイルと同名のログ・ファイルに記録されます。(`sample.pptx` の言語修正は `sample.log` に記録される)
* ログに記録されるLanguageIDの意味は以下のとおりです。詳細なリストは[MsoLanguageID enumeration (Office) | Microsoft Docs](https://docs.microsoft.com/en-us/office/vba/api/office.msolanguageid)をご参照ください。

  Value | Description
  ------|-------------------------
  2057  | The English UK language
  1033  | The English US language
  1036  | The French language

Tips
----

エクスプローラーの［送る］メニューに追加すると便利です。

参考: [［送る］メニューに項目を追加する方法（Windows 7／8.x／10編）：Tech TIPS - ＠IT](https://www.atmarkit.co.jp/ait/articles/1109/30/news131.html)

Author
------

[Shinichi Akiyama](https://github.com/shakiyam)

License
-------

[MIT License](https://opensource.org/licenses/MIT)
