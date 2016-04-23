using System;
using System.Dynamic;
using System.Reflection;
using System.Runtime.InteropServices;
using FrEngineLoader.Properties;

namespace FrEngineLoader
{
    public partial class DynamicFrComObjectWrapper
    {
        // Serves for statements like "dynamic propValue = myComObject.ComProperty;".
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            try
            {
                result = ComObjectType.InvokeMember(binder.Name, BindingFlags.GetProperty, Type.DefaultBinder, ComObject,
                    null);
            }
            catch (COMException e)
            {
                throw new ApplicationException(string.Format(Resources.EXC_COM, NativeComObjectTypeName, binder.Name), e);
            }
            WrapInvokeResult(ref result);
            return true;
        }

        // Serves for statements like "dynamic propValue = myComObject.ComProperties[0];".
        public override bool TryGetIndex(GetIndexBinder binder, object[] indexes, out object result)
        {
            try
            {
                result = ComObjectType.InvokeMember(FrEngineUtils.ElementPropertyName, BindingFlags.GetProperty,
                    Type.DefaultBinder, ComObject, indexes);
            }
            catch (COMException e)
            {
                throw new ApplicationException(
                    string.Format(Resources.EXC_COM, NativeComObjectTypeName, FrEngineUtils.ElementPropertyName), e);
            }
            WrapInvokeResult(ref result);
            return true;
        }

        // Serves for statements like "myComObject.ComProperty = NewComProperty";
        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            value = UnwrapValue(value);
            try
            {
                ComObjectType.InvokeMember(binder.Name, BindingFlags.SetProperty, Type.DefaultBinder, ComObject,
                    new[] {value});
            }
            catch (COMException e)
            {
                throw new ApplicationException(string.Format(Resources.EXC_COM, NativeComObjectTypeName, binder.Name), e);
            }
            return true;
        }

        // Serves for statements like "myComObject.ComProperties[0] = NewComProperty";
        public override bool TrySetIndex(SetIndexBinder binder, object[] indexes, object value)
        {
            value = UnwrapValue(value);
            try
            {
                ComObjectType.InvokeMember(FrEngineUtils.ElementPropertyName, BindingFlags.SetProperty,
                    Type.DefaultBinder, ComObject, new[] {indexes[0], value});
            }
            catch (COMException e)
            {
                throw new ApplicationException(
                    string.Format(Resources.EXC_COM, NativeComObjectTypeName, FrEngineUtils.ElementPropertyName), e);
            }
            return true;
        }

        // Serves for statements like "myComObject.Do(arg1, arg2, null);";
        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            WrapNullArgs(args);
            try
            {
                result = ComObjectType.InvokeMember(binder.Name, BindingFlags.InvokeMethod, Type.DefaultBinder,
                    ComObject, args);
            }
            catch (COMException e)
            {
                throw new ApplicationException(string.Format(Resources.EXC_COM, NativeComObjectTypeName, binder.Name), e);
            }
            WrapInvokeResult(ref result);
            return true;
        }

        // Wrap only COM objects returned from System.Type.InvokeMember().
        public static void WrapInvokeResult(ref object result)
        {
            if (result != null && result.GetType().IsCOMObject)
                result = new DynamicFrComObjectWrapper(result);
        }

        // Replace null objects in args with UnknownWrapper objects.
        // Required for interaction with FrEngine COM methods.
        public static void WrapNullArgs(object[] args)
        {
            for (var i = 0; i < args.Length; i++)
                args[i] = args[i] ?? NullObject;
        }

        // Try to unwrap "value" if it's of type DynamicFreComObjectWrapper.
        private object UnwrapValue(object value)
        {
            return value is DynamicFrComObjectWrapper ? ComObject : value;
        }
    }
}