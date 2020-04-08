namespace PxlHandler
{
    internal class PxlPath
    {
        public PxlPath(string filename, string worksheetName, string rangeAddress)
        {
            IsValid = true;
            Filename = filename;
            WorksheetName = worksheetName;
            RangeAddress = rangeAddress;
        }

        private PxlPath()
        {
            IsValid = false;
        }

        public string Filename { get; }
        public bool IsValid { get; }
        public string RangeAddress { get; }
        public string WorksheetName { get; }

        public static PxlPath Invalid()
        {
            return new PxlPath();
        }

        public override string ToString()
        {
            return IsValid ? $"filename: {Filename}, worksheetName: {WorksheetName}, rangeAddress: {RangeAddress}" : "invalid";
        }
    }
}