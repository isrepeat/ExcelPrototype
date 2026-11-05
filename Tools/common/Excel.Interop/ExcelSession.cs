using System;
using System.Collections.Generic;

namespace ExcelTools.Interop {
    public sealed class ExcelSession : IDisposable {
        private object application;
        private object workbooks;
        private bool disposed;

        public dynamic Application {
            get {
                return application;
            }
        }
        public dynamic Workbooks {
            get {
                return workbooks;
            }
        }

        public ExcelSession() {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) {
                throw new InvalidOperationException("Desktop Microsoft Excel is not installed.");
            }
            try {
                application = Activator.CreateInstance(excelType);
                dynamic excel = application;
                excel.Visible = false;
                excel.ScreenUpdating = false;
                excel.DisplayAlerts = false;
                excel.EnableEvents = false;
                excel.AutomationSecurity = 3;
                workbooks = excel.Workbooks;
            }
            catch (Exception initializationError) {
                try {
                    Dispose();
                }
                catch (Exception cleanupError) {
                    throw new AggregateException(initializationError, cleanupError);
                }
                throw;
            }
        }

        public void Dispose() {
            if (disposed) {
                return;
            }
            disposed = true;
            var errors = new List<Exception>();
            try {
                if (application != null) {
                    ((dynamic)application).Quit();
                }
            }
            catch (Exception error) {
                errors.Add(error);
            }
            try {
                ComObjects.Release(workbooks);
            }
            catch (Exception error) {
                errors.Add(error);
            }
            try {
                ComObjects.Release(application);
            }
            catch (Exception error) {
                errors.Add(error);
            }
            workbooks = null;
            application = null;
            if (errors.Count > 0) {
                throw new AggregateException(errors);
            }
        }
    }
}