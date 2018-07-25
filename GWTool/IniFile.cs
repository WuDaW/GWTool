using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace 公文
{
    class IniFile
    {
        [DllImport("kernel32")]
        public static extern bool WritePrivateProfileString(
            string AppName,
            string KeyName,
            string ValueName,
            string FileName);
        [DllImport("kernel32")]
        public static extern int GetPrivateProfileString(
            string AppName,
            string KeyName,
            string lpDefault,
            StringBuilder lpReturnedString,
            int nSize,
            string FileName);
    }
}
