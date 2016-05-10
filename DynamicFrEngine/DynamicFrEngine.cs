using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using FrEngineLoader.Properties;
using Microsoft.VisualBasic;

namespace FrEngineLoader
{
    /// <summary>
    ///     Loads and wraps ABBYY FineReader Engine 11 COM object so that it can be declared with "dynamic" keyword and
    ///     implements IDisposable pattern.
    /// </summary>
    public sealed class DynamicFrEngine : DynamicFrComObjectWrapper
    {
        private readonly string _password;
        private readonly string _projectId;
        private bool _disposed;
        private dynamic _engineLoader;

        /// <summary>
        ///     Initializes a new instance of wrapped ABBYY FineReader Engine 11 COM object using the specified loading
        ///     method.
        /// </summary>
        public DynamicFrEngine(FrEngineLoadingMethod loadingMethod, string projectId = null, string password = null)
            : base(null)
        {
            _projectId = projectId;
            _password = password;

            switch (loadingMethod)
            {
                case FrEngineLoadingMethod.InProcessComServer:
                    ComObject = LoadFrEngineAsComServer(FrEngineUtils.InprocComServerProgId);
                    break;
                case FrEngineLoadingMethod.OutOfProcessComServer:
                    ComObject = LoadFrEngineAsComServer(FrEngineUtils.OutprocComServerProgId);
                    break;
                default:
                    throw new ApplicationException(string.Format(Resources.EXC_LOADING_METHOD, loadingMethod));
            }

            ComObjectType = ComObject.GetType();
            NativeComObjectTypeName = Information.TypeName(ComObject);
        }

        public static IEnumerable<string> SupportedImageExtensions
        {
            get { return FrEngineUtils.SupportedImageExtensions; }
        }

        // Dispose pattern implementation for a derived class.
        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Cleanup managed resources.
                    base.Dispose(true); // Unload alive FrEngine COM object first.
                    UnloadFrEngineComServerIfLoaded();
                }
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        private object LoadFrEngineAsComServer(string progId)
        {
            var engineLoaderType = Type.GetTypeFromProgID(progId);
            if (engineLoaderType == null)
                throw new ApplicationException(string.Format(Resources.EXC_COM_SERVER_REG, progId));

            _engineLoader = Activator.CreateInstance(engineLoaderType);

            return _engineLoader.GetEngineObjectEx(_projectId, null, null, true, null, _password);
        }

        private void UnloadFrEngineComServerIfLoaded()
        {
            if (_engineLoader == null) return;

            _engineLoader.ExplicitlyUnload();
            Marshal.FinalReleaseComObject(_engineLoader);
            _engineLoader = null;
        }
    }
}