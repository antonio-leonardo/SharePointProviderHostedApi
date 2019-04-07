namespace SharePointProviderHostedApi.Json
{
    public class JsonMetadataDocument
    {
        public string serviceName { get; set; }
        public JsonEndpoint[] endpoints { get; set; }
        public JsonKey[] keys { get; set; }
    }
}