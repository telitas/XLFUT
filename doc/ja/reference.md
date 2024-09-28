# XLFUT リファレンス

[README](./README.md)に戻る

## Language

- [英語(English)](../reference.md)
- 日本語

## モジュールの構造

XLFUTは入れ子になった名前空間構造を持っています。

```txt
XLFUT┬(関数・数式群)
     ├ASSERT─(関数・数式群)
     └UTILS─(関数・数式群)
```

最上位の名前空間は、"."で区切られた前方の文字列で表現されます。
次点またはそれ以降の名前空間は、"_"で区切られた接頭辞で表現されます。

```txt
XLFUT.<関数/数式名>
XLFUT.ASSERT_<関数/数式名>
XLFUT.UTILS_<関数/数式名>
```

## 関数・数式リファレンス

### XLFUT

#### VERSION

XLFUTのバージョンを返します。

#### LICENSE

XLFUTのライセンス文を返します。

### XLFUT.ASSERT

XLFUT.ASSERTに所属する関数は表明（アサーション）する役割を持ちます。

これらの関数は以下に示す共通仕様を持ちます。

<dl>
  <dt>戻り値</dt>
  <dd>1×2配列:
    <dl>
      <dt>要素(1,1)</dt>
      <dd>LOGICAL: 表明が成功した場合、TRUE</dd>
      <dt>要素(1,2)</dt>
      <dd>TEXT: エラーメッセージ</dd>
    </dl>
  </dd>
</dl>

#### 比較関数

比較関数は2つの値を比較する表明関数です。

比較関数は以下に示す共通仕様を持ちます。

<dl>
  <dt>引数</dt>
  <dd>
    <dl>
      <dt>actual</dt>
      <dd>実際の値</dd>
      <dt>expect</dt>
      <dd>期待する値</dd>
    </dl>
  </dd>
  <dt>戻り値</dt>
  <dd>1×2配列: XLFUT_ASSERTの共通戻り値</dd>
  <dt>注記</dt>
  <dd>ASSERT_SAME関数を除き、エラー値間の比較は常に失敗します。</dd>
</dl>

|関数名|対応するExcel演算子/関数|詳細|
|----|----------------------------|-----------|
|ASSERT_EQUAL|=||
|ASSERT_NOTEQUAL|<>||
|ASSERT_GREATERTHAN|>||
|ASSERT_GREATEROREQUAL|>=||
|ASSERT_LESSTHAN|<||
|ASSERT_LESSOREQUAL|<=||
|ASSERT_EXACT|EXACT||
|ASSERT_SAME|-|expectとactualが"同じ"であることを表明します|

ASSERT_SAME関数が示す"同じ"の意味は、ASSERT_EQUAL関数が示す"等しい"と似ています。しかし、以下の点において"同じ"は"等しい"よりも厳密です。

1. 型とサイズを区別する  
  ASSERT_SAMEは単一の値と1×1配列を区別し、
  同様に*n*×*m*配列と1×*m*配列を区別します。
2. エラー値比較可能  
  ASSERT_SAMEはエラー値が同じであればTRUEを返します。

|比較|ASSERT_EQUAL|ASSERT_SAME|
|-|------------|-----------|
|1 と {1}|TRUE|FALSE|
|{1,2,3;1,2,3} と {1,2,3}|TRUE|FALSE|
|#VALUE! と #VALUE!|FALSE|TRUE|

#### IS/ISNOT関数

IS/ISNOT関数は、値が特定の性質を満たすことを検証する表明関数です。

IS/ISNOT関数は以下に示す共通仕様を持ちます。

<dl>
  <dt>引数</dt>
  <dd>
    <dl>
      <dt>actual</dt>
      <dd>検証する値</dd>
    </dl>
  </dd>
  <dt>戻り値</dt>
  <dd>1×2配列: XLFUT_ASSERTの共通戻り値</dd>
  <dt>注記</dt>
  <dd>全てのIS関数には対応するISNOT関数があります。</dd>
</dl>

|Name|Corresponding Excel function|Description|
|----|----------------------------|-----------|
|ASSERT_ISNUMBER|ISNUMBER||
|ASSERT_ISTEXT|ISTEXT||
|ASSERT_ISLOGICAL|ISLOGICAL||
|ASSERT_ISERROR|ISERROR||
|ASSERT_ISARRAY|-|valueの型が配列であることを表明します|
|ASSERT_ISCOMPOUNDDATA|-|valueの型が複合データであることを表明します|
|ASSERT_ISNULL|-|valueが#NULL!エラーであることを表明します|
|ASSERT_ISDIV0|-|valueが#DIV/0!エラーであることを表明します|
|ASSERT_ISVALUEERROR|-|valueが#VALUE!エラーであることを表明します|
|ASSERT_ISREFERROR|-|valueが#REF!エラーであることを表明します|
|ASSERT_ISNAMEERROR|-|valueが#NAME?エラーであることを表明します|
|ASSERT_ISNUMERROR|-|valueが#NUM!エラーであることを表明します|
|ASSERT_ISNA|ISNA||
|ASSERT_ISGETTING_DATA|-|valueが#GETTING_DATAエラーであることを表明します|
|ASSERT_ISSPILLERROR|-|valueが#SPILL!エラーであることを表明します|
|ASSERT_ISCONNECTERROR|-|valueが#CONNECT!エラーであることを表明します|
|ASSERT_ISBLOCKED|-|valueが#BLOCKED!エラーであることを表明します|
|ASSERT_ISUNKNOWN|-|valueが#UNKNOWN!エラーであることを表明します|
|ASSERT_ISFIELDERROR|-|valueが#FIELD!エラーであることを表明します|
|ASSERT_ISCALCERROR|-|valueが#CALC!エラーであることを表明します|
|ASSERT_ISBLANK|ISBLANK||
|ASSERT_ISREF|ISREF||

#### 論理関数

論理関数は単一のテストに複数の表明を適用するため、結果をまとめます。

論理関数は以下に示す共通仕様を持ちます。

<dl>
  <dt>引数</dt>
  <dd>
    <dl>
      <dt>results</dt>
      <dd>複数の表明結果</dd>
    </dl>
  </dd>
  <dt>戻り値</dt>
  <dd>1×2配列: XLFUT_ASSERTの共通戻り値</dd>
</dl>

|Name|Corresponding Excel function|Description|
|----|----------------------------|-----------|
|ASSERT_AND|AND|複数の表明の結果をAND演算を用いて集約し、エラーメッセージを"AND()"書式を用いて結合します。|
|ASSERT_OR|OR|複数の表明の結果をOR演算を用いて集約し、エラーメッセージを"OR()"書式を用いて結合します。|

#### ASSERT_IN/ASSERT_NOTIN

値が候補リストに含まれているか（または含まれていないか）を検証します。

<dl>
  <dt>引数</dt>
  <dd>
    <dl>
      <dt>actual</dt>
      <dd>実際の値</dd>
      <dt>expect</dt>
      <dd>候補となる値のリスト</dd>
    </dl>
  </dd>
  <dt>戻り値</dt>
  <dd>1×2配列: XLFUT_ASSERTの共通戻り値</dd>
</dl>

#### ASSERT_SIZE

配列のサイズが特定の値であることを表明します。

<dl>
  <dt>引数</dt>
  <dd>
    <dl>
      <dt>data</dt>
      <dd>サイズが表明されるデータ</dd>
      <dt>rows</dt>
      <dd>期待される行サイズ</dd>
      <dt>columns</dt>
      <dd>期待される列サイズ</dd>
    </dl>
  </dd>
  <dt>戻り値</dt>
  <dd>1×2配列: XLFUT_ASSERTの共通戻り値</dd>
</dl>

#### ASSERT_EVALUATE

カスタマイズされた条件とエラーメッセージで表明します。

<dl>
  <dt>引数</dt>
  <dd>
    <dl>
      <dt>logical_formula</dt>
      <dd>表明に使用する論理式</dd>
      <dt>message</dt>
      <dd>論理式がFALSEの際に表示するエラーメッセージ</dd>
    </dl>
  </dd>
  <dt>戻り値</dt>
  <dd>1×2配列: XLFUT_ASSERTの共通戻り値</dd>
</dl>

### XLFUT_UTILS

#### UTILS_EXPORTTESTS

テストをテキストファイルにエクスポートするための書式に変換します。

<dl>
  <dt>引数</dt>
  <dd>
    <dl>
      <dt>tests_range</dt>
      <dd>Range: ヘッダを含む変換したいテストの範囲</dd>
    </dl>
  </dd>
  <dt>戻り値</dt>
  <dd>Range: 変換されたテスト</dd>
  <dt>注記</dt>
  <dd>description列が空白の行は省略されます。</dd>
</dl>

#### UTILS_OUTPUTINTAP

テスト結果をTAP形式に変換します。

<dl>
  <dt>引数</dt>
  <dd>
    <dl>
      <dt>tests_range</dt>
      <dd>Range: ヘッダを含む変換したいテストの範囲</dd>
    </dl>
  </dd>
  <dt>戻り値</dt>
  <dd>TEXT: 変換されたテスト結果</dd>
  <dt>注記</dt>
  <dd>
    <p>TAPのバージョンは14です。</p>
    <p>TAPの詳細は<a href="https://testanything.org/">https://testanything.org/</a>をご覧ください。</p>
  </dd>
</dl>
