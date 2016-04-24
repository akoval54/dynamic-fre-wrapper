using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using FrEngineLoader.Properties;

namespace FrEngineLoader
{
    internal static class FrEngineUtils
    {
        private const int SupportedFrEngineVersion = 11;

        public const uint LoadWithAlteredSearchPath = 0x00000008;
        public const string FrEngineDllFileName = "FREngine.dll";
        public const string FrEngineLoadingFunctionName = "GetEngineObjectEx";
        public const string FrEngineUnloadingFunctionName = "DeinitializeEngine";
        public const string InprocComServerProgId = "FREngine.InprocLoader.11";
        public const string OutprocComServerProgId = "FREngine.OutprocLoader.11";

        public const string CloseMethodName = "Close";
        public const string FlushMethodName = "Flush";
        public const string ElementPropertyName = "Element";
        public const string CountPropertyName = "Count";

        public static readonly IEnumerable<string> FrEngineClosableInterfaces = new[]
        {"IDocumentInfo", "IFRDocument", "IExportFileWriter", "IFileWriter", "IReadStream"};

        public static readonly IEnumerable<string> FrEngineFlushableInterfaces = new[] {"IFRPage"};

        [DllImport("kernel32.dll")]
        public static extern IntPtr LoadLibraryEx(string dllToLoad, IntPtr reserved, uint flags);

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetProcAddress(IntPtr hModule, string procedureName);

        [DllImport("kernel32.dll")]
        public static extern bool FreeLibrary(IntPtr hModule);

        public static void CheckFrEngineDllPath(string frEngineDllPath)
        {
            if (string.IsNullOrWhiteSpace(frEngineDllPath))
                throw new ApplicationException(string.Format(Resources.EXC_FRE_DLL_EMPTY_PATH, FrEngineDllFileName,
                    FrEngineLoadingMethod.Native));

            if (!string.Equals(FrEngineDllFileName, Path.GetFileName(frEngineDllPath),
                StringComparison.OrdinalIgnoreCase))
                throw new ApplicationException(string.Format(Resources.EXC_FRE_DLL_WRONG_PATH, FrEngineDllFileName,
                    frEngineDllPath, FrEngineLoadingMethod.Native));

            if (!File.Exists(frEngineDllPath))
                throw new ApplicationException(string.Format(Resources.EXC_FRE_DLL_FILE_NOT_EXIST, frEngineDllPath,
                    FrEngineDllFileName, FrEngineLoadingMethod.Native));

            // Check FREngine.dll version.
            var frEngineDllVErsion = FileVersionInfo.GetVersionInfo(frEngineDllPath).FileMajorPart;
            if (frEngineDllVErsion != SupportedFrEngineVersion)
                throw new ApplicationException(string.Format(Resources.EXC_FRE_DLL_FILE_VERSION, frEngineDllVErsion,
                    SupportedFrEngineVersion));
        }
    }

    // Delegate type for FREngine.dll GetEngineObjectEx() function.
    [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
    internal delegate int GetEngineObjectEx(
        string projectId,
        string frEngineDataFolder,
        string frEngineTempFolder,
        bool isSharedCpuCoresMode,
        string openLicensePath,
        string openLicensePassword,
        [MarshalAs(UnmanagedType.IUnknown)] ref object engine
        );

    public enum FrEngineLoadingMethod
    {
        Native,
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