using Microsoft.VisualStudio.Shell;
using System.ComponentModel;

namespace SourcetrailDotnetIndexerExtension
{
    public class Settings : DialogPage
    {
        [Category("Sourcetrail")]
        [DisplayName("Executable Path")]
        public string Sourcetrail_ExecutablePath { get; set; } = @"C:\Program Files\Sourcetrail\Sourcetrail.exe";

        [Category("Sourcetrail")]
        [DisplayName("Open After Generating")]
        public bool Sourcetrail_OpenAfterGenerating { get; set; } = true;


        [Category("Indexer")]
        [DisplayName("Executable Path")]
        public string Indexer_ExecutablePath { get; set; }

        [Category("Indexer")]
        [DisplayName("Search Paths (-s)")]
        public string[] Indexer_SearchPaths { get; set; } = new[]
        {
            @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
        };

        [Category("Indexer")]
        [DisplayName("Name Filters (-f)")]
        public string[] Indexer_NameFilters { get; set; }

        [Category("Indexer")]
        [DisplayName("Output Path (-o)")]
        public string Indexer_OutputPath { get; set; }

        [Category("Indexer")]
        [DisplayName("Output Filename (-of)")]
        public string Indexer_OutputFilename { get; set; }
    }
}
