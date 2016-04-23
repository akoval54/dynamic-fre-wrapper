using System;
using System.Runtime.InteropServices;
using FrEngineLoader.Properties;

namespace FrEngineLoader
{
    public sealed partial class DynamicFrEngine
    {
        private readonly string _password;
        private readonly string _projectId;
        private Action _deinitializeEngine;
        private dynamic _engineLoader;
        private IntPtr _frEngineDllHandle;

        private object LoadFrEngineNatively(string frEngineDllPath)
        {
            FrEngineUtils.CheckFrEngineDllPath(frEngineDllPath);

            // Load FREngine.dll library.
            _frEngineDllHandle = FrEngineUtils.LoadLibraryEx(frEngineDllPath, IntPtr.Zero,
                FrEngineUtils.LoadWithAlteredSearchPath);
            if (_frEngineDllHandle == IntPtr.Zero)
                throw new ApplicationException(string.Format(Resources.EXC_FRE_DLL_CANNOT_LOAD, frEngineDllPath));

            // Prepare FineReader Engine unloading function delegate.
            var deinitializeEngineHandle = FrEngineUtils.GetProcAddress(_frEngineDllHandle,
                FrEngineUtils.FrEngineUnloadingFunctionName);
            if (deinitializeEngineHandle == IntPtr.Zero)
                throw new ApplicationException(string.Format(Resources.EXC_FRE_DLL_FUNC_SEARCH,
                    FrEngineUtils.FrEngineUnloadingFunctionName, frEngineDllPath));
            _deinitializeEngine =
                (Action) Marshal.GetDelegateForFunctionPointer(deinitializeEngineHandle, typeof(Action));

            // Prepare FineReader Engine loading function delegate and try to get Engine object.
            object engine = null;

            var getEngineObjectHandle = FrEngineUtils.GetProcAddress(_frEngineDllHandle,
                FrEngineUtils.FrEngineLoadingFunctionName);
            if (getEngineObjectHandle == IntPtr.Zero)
                throw new ApplicationException(string.Format(Resources.EXC_FRE_DLL_FUNC_SEARCH,
                    FrEngineUtils.FrEngineLoadingFunctionName, frEngineDllPath));
            var getEngineObjectEx =(GetEngineObjectEx)Marshal.GetDelegateForFunctionPointer(getEngineObjectHandle,
                    typeof(GetEngineObjectEx));
            Marshal.ThrowExceptionForHR(getEngineObjectEx(_projectId, null, null, true, null, _password, ref engine));

            return engine;
        }

        private object LoadFrEngineAsComServer(string progId)
        {
            var engineLoaderType = Type.GetTypeFromProgID(progId);
            if (engineLoaderType == null) throw new ApplicationException(string.Format(Resources.EXC_COM_SERVER_REG, progId,
                    FrEngineUtils.FrEngineDllFileName));

            _engineLoader = Activator.CreateInstance(engineLoaderType);

            return _engineLoader.GetEngineObjectEx(_projectId, null, null, true, null, _password);
        }
    }
}