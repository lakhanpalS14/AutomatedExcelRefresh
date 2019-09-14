using Newtonsoft.Json;

namespace ExcelPivotRefresh
{
    class FileLocation
    {
        [JsonProperty("SourceLocation")]
        public string SourceLocation { get; set; }

        [JsonProperty("DestinationLocation")]
        public string DestinationLocation { get; set; }
    }
}