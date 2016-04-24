using System;
using System.Collections;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using FrEngineLoader.Properties;
using Microsoft.VisualBasic;

namespace FrEngineLoader
{
    public partial class DynamicFrComObjectWrapper : DynamicObject, IDisposable, IEnumerable
    {
        private static readonly UnknownWrapper NullObject = new UnknownWrapper(null);
        private bool _disposed;
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

        public object ComObject { get; protected set; }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // IEnumerable implementation for FrEngine collections.
        public IEnumerator GetEnumerator()
        {
            int collectionItemsCount;
            try
            {
                collectionItemsCount =
                    (int) ComObjectType.InvokeMember(FrEngineUtils.CountPropertyName, BindingFlags.GetProperty,
                        Type.DefaultBinder, ComObject, null);
            }
            catch (Exception e)
            {
                throw new ApplicationException(
                    string.Format(Resources.EXC_GET_ENUMERATOR, FrEngineUtils.CountPropertyName,
                        NativeComObjectTypeName), e);
            }

            for (var collectionIndex = 0; collectionIndex < collectionItemsCount; collectionIndex++)
            {
                var result = ComObjectType.InvokeMember(FrEngineUtils.ElementPropertyName, BindingFlags.GetProperty,
                    Type.DefaultBinder, ComObject, new[] {(object) collectionIndex});
                WrapInvokeResult(ref result);
                yield return result;
            }
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