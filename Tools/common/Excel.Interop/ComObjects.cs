using System;
using System.Runtime.InteropServices;

namespace ExcelTools.Interop {
    public static class ComObjects {
        public static void Release(object value) {
            if (value != null && Marshal.IsComObject(value)) {
                Marshal.FinalReleaseComObject(value);
            }
        }

        public static void CollectReleasedObjects() {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}