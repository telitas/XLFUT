# XLFUT TIPS

[README](./README.md)に戻る

## Language

- [英語(English)](../TIPS.md)
- 日本語

## テストをより読みやすくするために

`LET`関数を使用すると、より読みやすくなるかもしれません。

```excel
=LET(
    expect, "hello, world",
    actual, MYMODULE.HELLOWORLD(),
    XLFUT.ASSERT_EQUAL(expect, actual)
)
```

## 単一の式に複数のアサーションを適用するには

`VSTACK`関数と`ASSERT_AND`/`ASSERT_OR`アサーションを使用します。

```excel
=LET(
    actual, MYMODULE.HELLOWORLD(),
    XLFUT.ASSERT_AND(
        VSTACK(
            XLFUT.ASSERT_ISNOTARRAY(actual),
            XLFUT.ASSERT_ISTEXT(actual),
            XLFUT.ASSERT_EQUAL(actual, "hello, world")
        )
    )
)
```

## カスタマイズしたアサーションを作るには

`ASSERT_EVALUATE`関数を使用します。

```excel
=LET(
    expect, "hello, world",
    XLFUT.ASSERT_EVALUATE(EXACT(expect, LOWER(expect)), "The value must not contains any upper case characters.")
)
```
