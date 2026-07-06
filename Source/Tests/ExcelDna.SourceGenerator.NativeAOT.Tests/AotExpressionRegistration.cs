namespace ExcelDna.SourceGenerator.NativeAOT.Tests
{
    public class AotExpressionRegistration
    {
        [Fact]
        public void CreateLambdaWithAotContextUsesExtendedActionForVoidWithMoreThan16Arguments()
        {
            var parameters = Enumerable.Range(1, 17)
                .Select(i => System.Linq.Expressions.Expression.Parameter(typeof(int), "i" + i))
                .ToArray();

            var lambda = ExcelDna.Registration.ExcelFunctionRegistration.CreateLambdaWithAotContext(
                System.Linq.Expressions.Expression.Empty(),
                "VoidArgs17",
                parameters,
                "Test");

            Assert.Equal(typeof(void), lambda.ReturnType);
            Assert.Equal(17, lambda.Parameters.Count);
            Assert.StartsWith("ExtendedAction17", lambda.Type.Name);
        }
    }
}
