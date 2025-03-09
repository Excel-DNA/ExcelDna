namespace Addin.Types.Managed;

using Addin.Types.Unmanaged;
using System.Runtime.InteropServices;
using static Addin.Types.Unmanaged.ExcelConstants;

public class XlOper
{
    xloper12 _instance;

    public XlOper()
    {
        _instance = new();
    }

    public XlOper(string str)
    {
        // Calculate sizes
        var strLen = str.Length + 1;
        var charLen = Marshal.SizeOf(str[0]);
        var byteCount = charLen * strLen;

        // Create byte array, with the length stored in the first char and the string contents on the rest
        var bytes = new char[strLen];
        bytes[0] = (char)(strLen - 1);
        Buffer.BlockCopy(str.ToCharArray(), 0, bytes, charLen, byteCount - charLen);

        // Convert to unmanaged bytes
        var strPtr = Marshal.AllocHGlobal(byteCount);
        Marshal.Copy(bytes, 0, strPtr, strLen);

        // Add to XlOper structure
        xloper12 lpx = new() { xltype = xltypeStr };
        lpx.val.str = strPtr;
        _instance = lpx;
    }

    public XlOper(nint ptr)
    {
        _instance = Marshal.PtrToStructure<xloper12>(ptr);
    }

    public override string? ToString()
    {
        if (_instance.xltype != xltypeStr)
            throw new NotSupportedException();
        var str = Marshal.PtrToStringUni(_instance.val.str);
        var bytes = str?.ToCharArray().Skip(1).ToArray();
        return new string(bytes);
    }

    public nint ToPtr()
    {
        var ptr = Marshal.AllocHGlobal(Marshal.SizeOf(_instance));

        Marshal.StructureToPtr(_instance, ptr, false);

        return ptr;
    }
}
