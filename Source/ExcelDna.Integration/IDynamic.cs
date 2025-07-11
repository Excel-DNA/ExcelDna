﻿#if COM_GENERATED

#nullable enable

namespace ExcelDna.Integration
{
    public interface IDynamic
    {
        object? Get(string propertyName);
        T Get<T>(string propertyName);

        object? Get(string propertyName, object[] args);
        T Get<T>(string propertyName, object[]? args);

        void Set(string propertyName, object value);

        object? this[int index] { get; }
        object? this[string index] { get; }

        object? Invoke(string functionName, object[]? args);
        T Invoke<T>(string functionName, object[]? args);
    }
}

#endif
