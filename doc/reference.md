# XLFUT Reference

Back to [README](../README.md)

## Language

- English
- [Japanese(日本語)](./ja/reference.md)

## Module structure

XLFUT has nested namespace structure.

```txt
XLFUT┬(functions and formulas)
     ├ASSERT─(functions and formulas)
     └UTILS─(functions and formulas)
```

Top level namespace is represented as front string separated by ".".
Second or subsequent level namespace is represented as prefix separated by "_".

```txt
XLFUT.<function or formula name>
XLFUT.ASSERT_<function or formula name>
XLFUT.UTILS_<function or formula name>
```

## Functions and formulas refereneces

### XLFUT

#### VERSION

Return the version of XLFUT.

#### LICENSE

Return the license term of XLFUT.

### XLFUT.ASSERT

Functions belong to XLFUT_ASSERT have a role in asserion.

The functions have the following common specifications.

<dl>
  <dt>return</dt>
  <dd>1×2 array:
    <dl>
      <dt>element of (1,1)</dt>
      <dd>LOGICAL: TRUE if the assertion was successful</dd>
      <dt>element of (1,2)</dt>
      <dd>TEXT: Error message</dd>
    </dl>
  </dd>
</dl>

#### Comparison functions

Comparison functions are assert functions that compares 2 values.

Comparison functions have the following common specifications.

<dl>
  <dt>parameters</dt>
  <dd>
    <dl>
      <dt>actual</dt>
      <dd>The value of actual</dd>
      <dt>expect</dt>
      <dd>The value of expect</dd>
    </dl>
  </dd>
  <dt>return</dt>
  <dd>1×2 array: Common return value of XLFUT_ASSERT</dd>
  <dt>notes</dt>
  <dd>Comparison between errors always fail excluding ASSERT_SAME function</dd>
</dl>

|Name|Corresponding Excel operator/function|Description|
|----|----------------------------|-----------|
|ASSERT_EQUAL|=||
|ASSERT_NOTEQUAL|<>||
|ASSERT_GREATERTHAN|>||
|ASSERT_GREATEROREQUAL|>=||
|ASSERT_LESSTHAN|<||
|ASSERT_LESSOREQUAL|<=||
|ASSERT_EXACT|EXACT||
|ASSERT_SAME|-|Assert that expect and actual are "SAME"|

Meaning of "SAME" indicated by the ASSERT_SAME is similar to "EQUAL" indicated by the ASSERT_EQUAL.But "SAME" is more strict than "EQUAL" in the following points:

1. Type and size sensitive  
  ASSERT_SAME distinguish between single value and 1×1 array,
  as well as between *n*×*m* array and 1×*m* array.
2. Error comparable  
  ASSERT_SAME returns TRUE if error values are same.

|comparison|ASSERT_EQUAL|ASSERT_SAME|
|-|------------|-----------|
|1 and {1}|TRUE|FALSE|
|{1,2,3;1,2,3} and {1,2,3}|TRUE|FALSE|
|compare #VALUE! and #VALUE!|FALSE|TRUE|

#### IS/ISNOT functions

IS/ISNOT functions are assert functions that verifies that the value satisfies(or not) a certain property.

IS/ISNOT functions have the following common specifications.

<dl>
  <dt>parameters</dt>
  <dd>
    <dl>
      <dt>actual</dt>
      <dd>The value to be verified</dd>
    </dl>
  </dd>
  <dt>return</dt>
  <dd>1×2 array: Common return value of XLFUT_ASSERT</dd>
  <dt>notes</dt>
  <dd>Every IS function has a pairwise ISNOT function</dd>
</dl>

|Name|Corresponding Excel function|Description|
|----|----------------------------|-----------|
|ASSERT_ISNUMBER|ISNUMBER||
|ASSERT_ISTEXT|ISTEXT||
|ASSERT_ISLOGICAL|ISLOGICAL||
|ASSERT_ISERROR|ISERROR||
|ASSERT_ISARRAY|-|Assert the value type is array|
|ASSERT_ISCOMPOUNDDATA|-|Assert the value type is compound data|
|ASSERT_ISNULL|-|Assert the value is #NULL! error|
|ASSERT_ISDIV0|-|Assert the value is #DIV/0! error|
|ASSERT_ISVALUEERROR|-|Assert the value is #VALUE! error|
|ASSERT_ISREFERROR|-|Assert the value is #REF! error|
|ASSERT_ISNAMEERROR|-|Assert the value is #NAME? error|
|ASSERT_ISNUMERROR|-|Assert the value is #NUM! error|
|ASSERT_ISNA|ISNA||
|ASSERT_ISGETTING_DATA|-|Assert the value is #GETTING_DATA error|
|ASSERT_ISSPILLERROR|-|Assert the value is #SPILL! error|
|ASSERT_ISCONNECTERROR|-|Assert the value is #CONNECT! error|
|ASSERT_ISBLOCKED|-|Assert the value is #BLOCKED! error|
|ASSERT_ISUNKNOWN|-|Assert the value is #UNKNOWN! error|
|ASSERT_ISFIELDERROR|-|Assert the value is #FIELD! error|
|ASSERT_ISCALCERROR|-|Assert the value is #CALC! error|
|ASSERT_ISBLANK|ISBLANK||
|ASSERT_ISREF|ISREF||

#### Logical functions

Logical functions reduce the results of the assertions for apply many assertion to the test.

Logical functions have the following common specifications.

<dl>
  <dt>parameters</dt>
  <dd>
    <dl>
      <dt>results</dt>
      <dd>many results of assertions</dd>
    </dl>
  </dd>
  <dt>return</dt>
  <dd>1×2 array: Common return value of XLFUT_ASSERT</dd>
</dl>

|Name|Corresponding Excel function|Description|
|----|----------------------------|-----------|
|ASSERT_AND|AND|The results of the assertions are summed with AND operator and the error messages are combined in the "AND()" format.|
|ASSERT_OR|OR|The results of the assertions are summed with OR operator and the error messages are combined in the "OR()" format.|

#### ASSERT_IN/ASSERT_NOTIN

Verify the value is contained(or not) in candidate list.

<dl>
  <dt>parameters</dt>
  <dd>
    <dl>
      <dt>actual</dt>
      <dd>the value of actual</dd>
      <dt>expect</dt>
      <dd>candidate list of expected values</dd>
    </dl>
  </dd>
  <dt>return</dt>
  <dd>1×2 array: Common return value of XLFUT_ASSERT</dd>
</dl>

#### ASSERT_SIZE

Assert the size of array is specified size.

<dl>
  <dt>parameters</dt>
  <dd>
    <dl>
      <dt>data</dt>
      <dd>The Data that the size is asserted</dd>
      <dt>rows</dt>
      <dd>Expected rows size</dd>
      <dt>columns</dt>
      <dd>Expected columns size</dd>
    </dl>
  </dd>
  <dt>return</dt>
  <dd>1×2 array: Common return value of XLFUT_ASSERT</dd>
</dl>

#### ASSERT_EVALUATE

Assert with the customized condition and error message.

<dl>
  <dt>parameters</dt>
  <dd>
    <dl>
      <dt>logical_formula</dt>
      <dd>logical expression used for assertion</dd>
      <dt>message</dt>
      <dd>The error message displayed if logical_formula is false</dd>
    </dl>
  </dd>
  <dt>return</dt>
  <dd>1×2 array: Common return value of XLFUT_ASSERT</dd>
</dl>

### XLFUT_UTILS

#### UTILS_EXPORTTESTS

Convert tests to the format for exporting to text.

<dl>
  <dt>parameters</dt>
  <dd>
    <dl>
      <dt>tests_range</dt>
      <dd>Range: Range of tests including the header for conversion</dd>
    </dl>
  </dd>
  <dt>return</dt>
  <dd>Range: Converted tests</dd>
  <dt>note</dt>
  <dd>Rows with empty description column are skipped</dd>
</dl>

#### UTILS_OUTPUTINTAP

Convert results of tests to TAP format.

<dl>
  <dt>parameters</dt>
  <dd>
    <dl>
      <dt>tests_range</dt>
      <dd>Range: Range of tests including the header for conversion</dd>
    </dl>
  </dd>
  <dt>return</dt>
  <dd>TEXT: Converted results of tests</dd>
  <dt>note</dt>
  <dd>
    <p>TAP version is 14.</p>
    <p>See <a href="https://testanything.org/">https://testanything.org/</a> for information on TAP.</p>
  </dd>
</dl>
