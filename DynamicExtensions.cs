using System;
using System.Reflection;

namespace PdfReducer
{
    // for some reason Acrobat crashes on IDispatch::GetTypeInfo, so C# dynamic keyword doesn't work
    // but "old" reflection way still works
    public static class DynamicExtensions
    {
        public static object GetPropertyValue(this object obj, string name)
        {
            if (obj == null)
                throw new ArgumentNullException(nameof(obj));

            if (name == null)
                throw new ArgumentNullException(nameof(name));

            return obj.GetType().InvokeMember(name, BindingFlags.Instance | BindingFlags.GetProperty | BindingFlags.Public, null, obj, null);
        }

        public static void SetPropertyValue(this object obj, string name, object value)
        {
            if (obj == null)
                throw new ArgumentNullException(nameof(obj));

            if (name == null)
                throw new ArgumentNullException(nameof(name));

            obj.GetType().InvokeMember(name, BindingFlags.Instance | BindingFlags.SetProperty | BindingFlags.Public, null, obj, new object[] { value });
        }

        // use Type.Missing for non set optional arguments
        public static object InvokeMember(this object obj, string name, params object[] args)
        {
            if (obj == null)
                throw new ArgumentNullException(nameof(obj));

            if (name == null)
                throw new ArgumentNullException(nameof(name));

            return obj.GetType().InvokeMember(name, BindingFlags.Instance | BindingFlags.InvokeMethod | BindingFlags.Public, null, obj, args);
        }
    }
}
