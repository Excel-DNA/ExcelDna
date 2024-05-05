using System;
using System.Collections.Generic;

namespace ExcelDna.Integration.ObjectHandles
{
    public interface IHasRowVersion
    {
        ulong RowVersion { get; }
    }

    public interface IDataService
    {
        IHasRowVersion ProcessRequest(string objectType, object[] parameters);

        // Should return a list in the same order as the requestInfos, with nulls for objects not to be refreshed
        IList<IHasRowVersion> ProcessUpdateRequests(IList<Tuple<string, object[], ulong>> requestInfos);
    }

    public class DataObject1 : IHasRowVersion
    {
        public string Code { get; set; }
        public DateTime DateTime { get; set; }
        public double Value { get; set; }
        public ulong RowVersion { get; set; }   // Added for polling updates
    }

    public class DataObject2 : IHasRowVersion
    {
        public string[] Columns { get; set; }
        public object[,] Values { get; set; }
        public ulong RowVersion { get; set; }   // Added for polling updates
    }

    public class DataObjectGeneral : IHasRowVersion
    {
        public ulong Handle { get; set; }
        public ulong RowVersion { get; set; }   // Added for polling updates
    }

    public class DataService : IDataService // : IDisposable
    {
        Random rand;

        public DataService()
        {
            rand = new Random();
        }

        public IHasRowVersion ProcessRequest(string objectType, object[] parameters)
        {
            switch (objectType)
            {
                case "DataObject1":
                    // Do query against the back-end, using parameters
                    var result = new DataObject1
                    {
                        Code = (string)parameters[0],
                        DateTime = DateTime.Now,
                        Value = rand.Next(5, 50),
                        RowVersion = 123

                    };
                    return result;
                case "DataObject2":
                    // Do something
                    return null;
                case "DataObjectGeneral":
                    // Do query against the back-end, using parameters
                    return new DataObjectGeneral
                    {
                        Handle = (ulong)parameters[0],
                        RowVersion = 123

                    };
                default:
                    throw new ArgumentException("objectType");
            }

        }

        public IList<IHasRowVersion> ProcessUpdateRequests(IList<Tuple<string, object[], ulong>> requestInfos)
        {
            // Loop through the tuples - they represent the ObjectType, Parameters and RowVersion 
            // from which to build the batch stored proc call.
            // Maybe something like table-valued parameters: https://msdn.microsoft.com/en-us/library/bb675163%28v=vs.110%29.aspx

            // Should return a list in the same order as the requestInfos, with nulls for objects not to be refreshed
            var results = new List<IHasRowVersion>();

            // As a test, I just update every second object, ignoring the RowVersion
            for (int i = 0; i < requestInfos.Count; i++)
            {
                var requestInfo = requestInfos[i];
                if (i % 2 == 0)
                {
                    results.Add(ProcessRequest(requestInfo.Item1, requestInfo.Item2));
                }
                else
                {
                    results.Add(null);
                }
            }
            return results;
        }

        /* 
           Conversion from DB byte[] to ulong for SQL Server rowversion columns:

           byte[] byteArray = {0, 0, 0, 0, 0, 0, 0, 8};
           var value = BitConverter.ToUInt64(byteArray.Reverse().ToArray(), 0);
        */

        //public void Dispose()
        //{
        //    // This code will run when there are no more active requests
        //}
    }
}
