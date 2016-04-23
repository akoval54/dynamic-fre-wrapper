using System;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;

namespace FrEngineLoader
{
    public partial class DynamicFrComObjectWrapper : DynamicObject, IDisposable
    {
        private static readonly UnknownWrapper NullObject = new UnknownWrapper(null);
        private bool _disposed;
        protected object ComObject;
        protected Type ComObjectType;
        protected string NativeComObjectTypeName;

        protected DynamicFrComObjectWrapper(object obj)
        {
            if (obj != null)
            {
                ComObject = obj;
                ComObjectType = ComObject.GetType();
                NativeComObjectTypeName = Information.TypeName(ComObject);
            }
        }

        // Provide original COM object for debugging purposes.
        public object WrappedObject
        {
            get { return ComObject; }
        }

        // Dispose pattern implementation for a base class.
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~DynamicFrComObjectWrapper()
        {
            Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;
            if (disposing) ReleaseFrComObject();
            _disposed = true;
        }

        private void ReleaseFrComObject()
        {
            if (ComObject == null) return;

            if (FrEngineUtils.FrEngineClosableInterfaces.Any(
                closableInterface => closableInterface == NativeComObjectTypeName))
                ComObjectType.InvokeMember(FrEngineUtils.CloseMethodName, BindingFlags.InvokeMethod, Type.DefaultBinder,
                    ComObject, null);

            if (FrEngineUtils.FrEngineFlushableInterfaces.Any(
                flushableInterface => flushableInterface == NativeComObjectTypeName))
                ComObjectType.InvokeMember(FrEngineUtils.FlushMethodName, BindingFlags.InvokeMethod, Type.DefaultBinder,
                    ComObject, new[] {(object) true});

            Marshal.FinalReleaseComObject(ComObject);
            ComObject = null;
        }
    }
}