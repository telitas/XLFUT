# XLFUT TIPS

Back to [README](../README.md)

## Language

- English
- [Japanese(日本語)](./ja/TIPS.md)

## How to make the test code more readable

Maybe `LET` function makes your tests more readable.

```excel
=LET(
    expect, "hello, world",
    actual, MYMODULE.HELLOWORLD(),
    XLFUT.ASSERT_EQUAL(expect, actual)
)
```

## How to run multiple assertions on a single formula

Use `VSTACK` function and `ASSERT_AND`/`ASSERT_OR` assertions.

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

## How to make customized assertions

Use `ASSERT_EVALUATE` function.

```excel
=LET(
    expect, "hello, world",
    XLFUT.ASSERT_EVALUATE(EXACT(expect, LOWER(expect)), "The value must not contains any upper case characters.")
)
```
