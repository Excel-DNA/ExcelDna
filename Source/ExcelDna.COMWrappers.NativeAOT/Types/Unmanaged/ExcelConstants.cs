namespace Addin.Types.Unmanaged;

public static class ExcelConstants
{
    // Function number bits
    public static readonly int xlCommand = 0x8000;
    public static readonly int xlSpecial = 0x4000;
    public static readonly int xlIntl = 0x2000;
    public static readonly int xlPrompt = 0x1000;

    // Auxiliary function numbers
    // These functions are available only from the C API,
    // not from the Excel macro language.
    public static readonly int xlFree = 0 | xlSpecial;
    public static readonly int xlStack = 1 | xlSpecial;
    public static readonly int xlCoerce = 2 | xlSpecial;
    public static readonly int xlSet = 3 | xlSpecial;
    public static readonly int xlSheetId = 4 | xlSpecial;
    public static readonly int xlSheetNm = 5 | xlSpecial;
    public static readonly int xlAbort = 6 | xlSpecial;

    // Returns application's hinstance as an integer value, supported on 32-bit platform only
    public static readonly int xlGetInst = 7 | xlSpecial;
    public static readonly int xlGetHwnd = 8 | xlSpecial;
    public static readonly int xlGetName = 9 | xlSpecial;
    public static readonly int xlEnableXLMsgs = 10 | xlSpecial;
    public static readonly int xlDisableXLMsgs = 11 | xlSpecial;
    public static readonly int xlDefineBinaryName = 12 | xlSpecial;
    public static readonly int xlGetBinaryName = 13 | xlSpecial;

    // GetFooInfo are valid only for calls to LPenHelper
    public static readonly int xlGetFmlaInfo = 14 | xlSpecial;
    public static readonly int xlGetMouseInfo = 15 | xlSpecial;

    // Set return value from an asynchronous function call
    public static readonly int xlAsyncReturn = 16 | xlSpecial;

    // Register an XLL event
    public static readonly int xlEventRegister = 17 | xlSpecial;

    // Returns true if running on Compute Cluster
    public static readonly int xlRunningOnCluster = 18 | xlSpecial;

    // Returns application's hinstance as a handle, supported on both 32-bit and 64-bit platforms
    public static readonly int xlGetInstPtr = 19 | xlSpecial;

    public static readonly int xlfRegister = 149;

    public static readonly int xltypeNum = 0x0001;
    public static readonly int xltypeStr = 0x0002;
    public static readonly int xltypeBool = 0x0004;
    public static readonly int xltypeRef = 0x0008;
    public static readonly int xltypeErr = 0x0010;
    public static readonly int xltypeFlow = 0x0020;
    public static readonly int xltypeMulti = 0x0040;
    public static readonly int xltypeMissing = 0x0080;
    public static readonly int xltypeNil = 0x0100;
    public static readonly int xltypeSRef = 0x0400;
    public static readonly int xltypeInt = 0x0800;

    public static readonly int xlbitXLFree = 0x1000;
    public static readonly int xlbitDLLFree = 0x4000;

    public static readonly int xltypeBigData = xltypeStr | xltypeInt;

    // Error codes
    // Used for val.err field of XLOPER and XLOPER12 structures
    // when constructing error XLOPERs and XLOPER12s
    public static readonly int xlerrNull = 0;
    public static readonly int xlerrDiv0 = 7;
    public static readonly int xlerrValue = 15;
    public static readonly int xlerrRef = 23;
    public static readonly int xlerrName = 29;
    public static readonly int xlerrNum = 36;
    public static readonly int xlerrNA = 42;
    public static readonly int xlerrGettingData = 43;

    // Flow data types
    // Used for val.flow.xlflow field of XLOPER and XLOPER12 structures
    // when constructing flow-control XLOPERs and XLOPER12s
    public static readonly int xlflowHalt = 1;
    public static readonly int xlflowGoto = 2;
    public static readonly int xlflowRestart = 8;
    public static readonly int xlflowPause = 16;
    public static readonly int xlflowResume = 64;

    // Return codes
    // These values can be returned from Excel4(), Excel4v(), Excel12() or Excel12v().
    public static readonly int xlretSuccess = 0; // success
    public static readonly int xlretAbort = 1; // macro halted
    public static readonly int xlretInvXlfn = 2; // invalid function number
    public static readonly int xlretInvCount = 4; // invalid number of arguments
    public static readonly int xlretInvXloper = 8; // invalid OPER structure
    public static readonly int xlretStackOvfl = 16; // stack overflow
    public static readonly int xlretFailed = 32; // command failed
    public static readonly int xlretUncalced = 64; // uncalced cell
    public static readonly int xlretNotThreadSafe = 128; // not allowed during multi-threaded calc
    public static readonly int xlretInvAsynchronousContext = 256; // invalid asynchronous function handle
    public static readonly int xlretNotClusterSafe = 512; // not supported on cluster
}
