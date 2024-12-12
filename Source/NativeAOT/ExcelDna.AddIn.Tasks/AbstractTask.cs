using Microsoft.Build.Framework;

namespace ExcelDna.AddIn.Tasks
{
    public abstract class AbstractTask : ITask
    {
        public abstract bool Execute();

        public IBuildEngine BuildEngine { get; set; }
        public ITaskHost HostObject { get; set; }
    }
}
