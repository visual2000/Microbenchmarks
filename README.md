# Benchmarking VB6

<figure>
  <blockquote cite="https://biblehub.com/1_corinthians/10-23.htm">
    All things are lawful for me, but all things are not expedient: all things are lawful for me, but all things edify not.
  </blockquote>
  <figcaption>1 Corinthians 10:23</figcaption>
</figure>

## Some hypotheses around potential speedups

What can we try to speed up about our current approach?

* Desugar structs: for example, a `Dim v As TVec3` should simply
  become `Dim v_x as Double, v_y as Double, v_z as Double`.

* don't use things like sqrt(), use John Carmack's numerical
  approximations

* Avoid function calls. This one might be harder, but declare global
  variables for the arguments and the return value.  Perhaps:

    ```VB6
    ' Inside main module:
    Dim value as Double
    For i = 1 To 10
        value = something(i * 1#, 20#)
    Next i
    ' .. the rest of the main module

    Public Function something(x as Double, y as Double) as Double
        something = x * y
    End Function
    ```

  might become:

```VB6
'' this used to be the guts of the function
Dim something_ret as Double
Dim something_x as Double, something_y as Double

Dim value as Double
For i = 1 To 10
    something_x = i * 1#
    something_y = 20#
    GoTo SomethingBody
CallSiteSomething1:
    value = something_ret
Next i
GoTo TheEnd

'' begin function something()
SomethingBody:
something_ret = something_x * something_y
GoTo CallSiteSomething1
'' end function something()

TheEnd:
```

* Will this `GoTo` craziness interact well with loops???  Hopefully,
  if `i` is unchanged?  Test!

* What if a proliferation of labels and `GoTo` make things slow?
  Think, extremely inefficient lookups?  It'd be weird, but hey.

## Findings

### Structs vs primitives

Appears to be nearly the same speed, perhaps primitives are a few
percent faster.  Not worth the convenience hit.

### Local loop vs TVec3_init function call

Calling out to the `TVec3_init` function is around 500% slower than
the local loop.  Even worse when compiled to Win32 exe - up to 20x ⚠⚠
slower!

### Can we jump in and out of loops with `GoTo`s?

Yes We Can™

### Does excessive use of `GoTo` make things slow?

We were worried we might be getting a similar performance hit to using
VB6 functions by reimplementing them ourselves.  Turns out it seems to
be.. pretty fast?  The same, or _faster_, than linear code.  That's
strange, but it doesn't look like a showstopper either way.

Time to build ourselves an optimising inliner!
