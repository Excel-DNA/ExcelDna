namespace Addin.Types.Managed;

public struct Variant
{
    public Variant(object? value)
    {
        Value = value;
    }

    public object? Value { get; set; }
    public IntPtr DispVal { get; set; }
}
