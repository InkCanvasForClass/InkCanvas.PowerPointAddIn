namespace PPTAgent.Contracts
{
    public static class PipeConstants
    {
        /// <summary>
        /// Default pipe name for PPTAgent communication.
        /// </summary>
        public const string DefaultPipeName = "PPTAgent_Pipe";

        /// <summary>
        /// Legacy pipe name used by InkCanvas for Class.
        /// </summary>
        public const string LegacyPipeName = "ICC_PPT_PIPE";

        /// <summary>
        /// Current pipe name. Can be changed at runtime before starting the host.
        /// </summary>
        public static string PipeName { get; set; } = DefaultPipeName;

        public const int ProtocolVersion = 1;
        public const int MaxFrameSize = 1024 * 1024; // 1 MB
        public const int ConnectTimeoutMilliseconds = 1000;
        public const int RequestTimeoutMilliseconds = 4000;
    }
}
