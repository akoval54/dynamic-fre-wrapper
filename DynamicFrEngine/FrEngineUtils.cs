using System.Collections.Generic;

namespace FrEngineLoader
{
    internal static class FrEngineUtils
    {
        public const string InprocComServerProgId = "FREngine.InprocLoader.11";
        public const string OutprocComServerProgId = "FREngine.OutprocLoader.11";
        public const string CloseMethodName = "Close";
        public const string FlushMethodName = "Flush";
        public const string ElementPropertyName = "Element";
        public const string CountPropertyName = "Count";

        public static readonly IEnumerable<string> FrEngineClosableInterfaces = new[]
        {"IDocumentInfo", "IFRDocument", "IExportFileWriter", "IFileWriter", "IReadStream"};

        public static readonly IEnumerable<string> FrEngineFlushableInterfaces = new[] {"IFRPage"};

        public static readonly IEnumerable<string> SupportedImageExtensions = new[]
        {
            ".bmp", ".dcx", ".djvu", ".djv", ".gif", ".jpg", ".jpeg", ".jfif", ".jp2", ".jpc", ".j2k", ".pcx",
            ".pdf", ".png", ".tif", ".tiff", ".jb2", ".wdp"
        };
    }

    public enum FrEngineLoadingMethod
    {
        InProcessComServer,
        OutOfProcessComServer
    }

    public enum FileExportFormat
    {
        Rtf = 0,
        HtmlVersion10Defaults,
        HtmlUnicodeDefaults,
        Xls,
        Pdf,
        TextVersion10Defaults,
        TextUnicodeDefaults,
        Xml,
        Docx,
        Xlsx,
        Pptx,
        Alto,
        Epub,
        Fb2,
        Odt,
        Xps
    }
}