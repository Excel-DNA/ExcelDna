---
layout: page
title: FSharp Type Inference
---

When creating UDFs with F#, the flexible type inference might lead to function signatures that are not supported by Excel-DNA, or lead to unexpected results.

{% highlight fsharp %}
let MakeTwo x = 2 
{% endhighlight %}

This doesn't work (the UDF doesn't get registered) since the inferred type is _'a -> int_, so is generic over the argument. This is equivalent to the C# signature:

{% highlight fsharp %}
public int MakeTwo<T>(T input) = { return 2; }
{% endhighlight %}
However, the following, with explicit typing,  does work: 

{% highlight fsharp %}
let MakeTwo (x : float) = 2 
{% endhighlight %}

This would apply to any function that is generic over its input. Another example is:

{% highlight fsharp %}
let AddString x y = x.ToString() + y.ToString()
{% endhighlight %}

which is of the type a' -> b' -> string and doesn't get exposed as an UDF either. 

Adding explicit types removes the generic parameters:

{% highlight fsharp %}
let AddString (x:obj) (y:obj) = x.ToString() + y.ToString()
{% endhighlight %}

Even the simple example in the distribution can be a concern:

{% highlight fsharp %}
let Add x y = x + y 
{% endhighlight %}
F# infers this function to be of the type int -> int -> int, and if called in Excel as =Add(2.5,3.5) then this function will return 7 not 6.

{% highlight fsharp %}
let Add (x:float) (y:float) = x + y 
{% endhighlight %}