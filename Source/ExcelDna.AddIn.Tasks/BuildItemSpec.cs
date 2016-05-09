namespace ExcelDna.AddIn.Tasks
{
    internal class BuildItemSpec
    {
        public string InputDnaFileName { get; set; }

        public string InputDnaFileNameAs32Bit { get; set; }
        public string InputDnaFileNameAs64Bit { get; set; }

        public string InputConfigFileNameAs32Bit { get; set; }
        public string InputConfigFileNameFallbackAs32Bit { get; set; }

        public string InputConfigFileNameAs64Bit { get; set; }
        public string InputConfigFileNameFallbackAs64Bit { get; set; }

        public string OutputDnaFileNameAs32Bit { get; set; }
        public string OutputDnaFileNameAs64Bit { get; set; }

        public string OutputXllFileNameAs32Bit { get; set; }
        public string OutputXllFileNameAs64Bit { get; set; }

        public string OutputConfigFileNameAs32Bit { get; set; }
        public string OutputConfigFileNameAs64Bit { get; set; }
    }
}
