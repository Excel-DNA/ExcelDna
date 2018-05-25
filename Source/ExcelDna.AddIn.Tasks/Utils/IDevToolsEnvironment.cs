namespace ExcelDna.AddIn.Tasks.Utils
{
    internal interface IDevToolsEnvironment
    {
        EnvDTE.Project GetProjectByName(string projectName);
    }
}
