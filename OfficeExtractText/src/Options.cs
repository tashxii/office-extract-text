namespace OfficeExtractText
{
    class OfficeTextExporterOption
    {
        public bool ExtractExcel { get; internal set; }
        public bool ExtractWord { get; internal set; }
        public bool ExtractPowerPoint { get; internal set; }
        public bool ExtractSubDir { get; internal set; }
        public string OutputDir { get; internal set; }
    }
}
