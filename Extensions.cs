using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace PdfReducer
{
    public static class Extensions
    {
        public static bool EqualsIgnoreCase(this string str, string text, bool trim = false)
        {
            if (trim)
            {
                str = str.Nullify();
                text = text.Nullify();
            }

            if (str == null)
                return text == null;

            if (text == null)
                return false;

            if (str.Length != text.Length)
                return false;

            return string.Compare(str, text, StringComparison.OrdinalIgnoreCase) == 0;
        }

        public static string Nullify(this string str)
        {
            if (str == null)
                return null;

            if (string.IsNullOrWhiteSpace(str))
                return null;

            var t = str.Trim();
            return t.Length == 0 ? null : t;
        }

        public static void EnsureFileDirectoryExists(string filePath)
        {
            if (filePath == null)
                throw new ArgumentNullException(nameof(filePath));

            if (!Path.IsPathRooted(filePath))
                throw new ArgumentException(null, nameof(filePath));

            var dir = Path.GetDirectoryName(filePath);
            if (string.IsNullOrWhiteSpace(dir))
                return;

            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        public static string GetAllMessagesWithDots(this Exception exception) => GetAllMessages(exception, s => s?.EndsWith(".") == true ? null : ". ");
        public static string GetAllMessages(this Exception exception) => GetAllMessages(exception, s => Environment.NewLine);
        public static string GetAllMessages(this Exception exception, Func<string, string> separator)
        {
            if (exception == null)
                return null;

            var sb = new StringBuilder();
            AppendMessages(sb, exception, separator);
            var msg = sb.ToString().Replace("..", ".").Nullify();
            return msg;
        }

        private static string GetExceptionTypeName(Exception exception)
        {
            if (exception == null)
                return null;

            var type = exception.GetType();
            if (type == null || string.IsNullOrWhiteSpace(type.FullName))
                return null;

            if (type.FullName.StartsWith("System.") ||
                type.FullName.StartsWith("Microsoft."))
                return null;

            return type.FullName;
        }

        private static void AppendMessages(StringBuilder sb, Exception e, Func<string, string> separator)
        {
            if (e == null)
                return;

            if (e is AggregateException agg)
            {
                foreach (var ex in agg.InnerExceptions)
                {
                    AppendMessages(sb, ex, separator);
                }
                return;
            }

            if (!(e is TargetInvocationException))
            {
                if (sb.Length > 0 && separator != null)
                {
                    var sep = separator(sb.ToString());
                    if (sep != null)
                    {
                        sb.Append(sep);
                    }
                }

                var typeName = GetExceptionTypeName(e);
                if (!string.IsNullOrWhiteSpace(typeName))
                {
                    sb.Append(typeName);
                    sb.Append(": ");
                    sb.Append(e.Message);
                }
                else
                {
                    sb.Append(e.Message);
                }
            }
            AppendMessages(sb, e.InnerException, separator);
        }

        public static IEnumerable<Exception> EnumerateAllExceptions(this Exception exception)
        {
            if (exception == null)
                yield break;

            yield return exception;
            if (exception is AggregateException agg)
            {
                foreach (var ae in agg.InnerExceptions)
                {
                    foreach (var child in EnumerateAllExceptions(ae))
                    {
                        yield return child;
                    }
                }
            }
            else
            {
                if (exception.InnerException != null)
                {
                    foreach (var child in EnumerateAllExceptions(exception.InnerException))
                    {
                        yield return child;
                    }
                }
            }
        }
    }
}
