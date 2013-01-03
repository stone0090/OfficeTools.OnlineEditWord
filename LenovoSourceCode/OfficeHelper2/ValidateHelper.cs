using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace CSN.DotNetLibrary.OfficeHelper
{

    public static class ValidateHelper
    {
        public static Validation Begin()
        {
            return null;
        }
    }

    public sealed class Validation
    {
        public bool IsValid { get; set; }
    }

    public static class ValidationExtensions
    {
        private static Validation Check<T>(this Validation validation, Func<bool> filterMethod, T exception) where T : Exception
        {
            if (filterMethod())
            {
                return validation ?? new Validation() { IsValid = true };
            }
            else
            {
                throw exception;
            }
        }

        public static Validation Check(this Validation validation, Func<bool> filterMethod)
        {
            return Check<Exception>(validation, filterMethod, new Exception("参数无效！"));
        }

        public static Validation NotNull(this Validation validation, object value)
        {
            return Check<ArgumentNullException>(
                validation,
                () => value != null,
                new ArgumentNullException("参数不能为null！")
            );
        }

        public static Validation NotNullAndEmpty(this Validation validation, string value)
        {
            return Check<ArgumentException>(
                validation,
                () => !string.IsNullOrEmpty(value),
                new ArgumentException("参数不能为空！")
            );
        }

        public static Validation InRange(this Validation validation, int value, int min, int max)
        {
            return Check<ArgumentOutOfRangeException>(
                validation,
                () => value >= min && value <= max,
                new ArgumentOutOfRangeException(string.Format("参数的范围必须在 {0} 与 {1} 之间！", min, max))
            );
        }

        public static Validation CheckFileType(this Validation validation, string value)
        {
            return Check<ArgumentException>(
                validation,
                () => value == ".pdf" || value == ".doc" || value == ".docx",
                new ArgumentException("无法识别的文件类型！")
            );
        }

        public static Validation FileExist(this Validation validation, string value)
        {
            return Check<ArgumentException>(
                validation,
                () => File.Exists(value),
                new ArgumentException("目标文件不存在！")
            );
        }
    }

}
