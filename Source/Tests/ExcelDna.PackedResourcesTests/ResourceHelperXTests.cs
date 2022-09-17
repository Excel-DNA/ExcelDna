using ExcelDna.PackedResources;
using NUnit.Framework;

namespace ExcelDna.PackedResourcesTests
{
    public class ResourceHelperXTests
    {
        [Test]
        public void AddResource()
        {
            ResourceHelperX.AddResource(null, null, null, null);
        }
    }
}
