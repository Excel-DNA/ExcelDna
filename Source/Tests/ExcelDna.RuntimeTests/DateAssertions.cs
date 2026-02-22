using System.Globalization;

namespace ExcelDna.RuntimeTests
{
    internal static class DateAssertions
    {
        public static void Equal(object actualValue, DateTime expected)
        {
            Assert.NotNull(actualValue);
            string actualText = actualValue.ToString()!;

            Assert.True(
                DateTime.TryParse(actualText, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime actualDate),
                $"Could not parse DateTime value '{actualText}'.");

            Assert.Equal(expected.Ticks, actualDate.Ticks);
        }

        public static void EqualPrefixed(object actualValue, string prefix, DateTime expected)
        {
            Assert.NotNull(actualValue);
            string actualText = actualValue.ToString()!;

            Assert.StartsWith(prefix, actualText);
            string datePart = actualText.Substring(prefix.Length);

            Assert.True(
                DateTime.TryParse(datePart, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime actualDate),
                $"Could not parse DateTime portion '{datePart}' from '{actualText}'.");

            Assert.Equal(expected.Ticks, actualDate.Ticks);
        }
    }
}
