# Visual Basic for Applications(VBA)コーディング規約
## 命名規則
### 命名方法の種類
次の記法を場面によって使い分けます。

|記法|例|詳細|
|---|---|---|
|スネークケース|find_by_item_code|単語の区切りをアンダースコアで区切る|
|パスカルケース(アッパーキャメルケース)|FindByItemCode|単語の区切りを大文字で書く|
|キャメルケース(ローワーキャメルケース)|findByItemCode|単語の区切りを大文字で書く(先頭は小文字)|

### プログラムコード
|種別|記法|プレフィックス|例|
|---|---|---|---|
|定数|スネークケース(大文字)|なし|MAX_ITEM|
|ローカル変数|キャメルケース|なし|currentItem|
|関数|パスカルケース|なし|GetLastItem|
|クラス|パスカルケース|Cls|ClsItemStorage|

クラスについてはプレフィックスを付けないのが一般的のようですが、VBAでは大文字小文字が区別されません。そのため、グローバル空間に影響があるため、プレフィックスに`Cls`を付けて区別することにします。。

### Excelオブジェクト
追加予定

### フォーム
フォーム自身と各コントロールにはキャメルケースを使い、種類に応じてプレフィックスを付けます。

|種別|プレフィックス|例|
|---|---|---
|フォームオブジェクト|fm|fmItemEdit

## 参考資料
### VBA関連
+ エクセルVBAコーディングガイドライン(いつも隣にITのお仕事)
  <https://tonari-it.com/excel-vba-coding-guide-line/>
+ Excel VBAコーディング ガイドライン案(Qiita)
  <https://qiita.com/mima_ita/items/8b0eec3b5a81f168822d>

### Visual Basic 6関連
+ Visual Studio 6.0 Constant and Variable Naming Conventions(MSDN)<br>
  <https://msdn.microsoft.com/en-us/library/aa240858(v=vs.60).aspx>
+ Visual Studio 6.0 Object Naming Conventions(MSDN)<br>
  <https://msdn.microsoft.com/en-us/library/aa263493(v=vs.60).aspx>
+ Visual Studio 6.0 Structured Coding Conventions(MSDN)<br>
  https://msdn.microsoft.com/en-us/library/aa733583(v=vs.60).aspx
