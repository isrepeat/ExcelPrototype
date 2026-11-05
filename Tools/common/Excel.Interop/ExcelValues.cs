using System.Runtime.InteropServices;

namespace ExcelTools.Interop {
    public static class ExcelValues {
        public static object PrepareForWrite(object values) {
            // Excel numbers arrive as Double; Int32 values represent VT_ERROR codes.
            // Rewrap them so managed COM does not write error codes as ordinary numbers.
            var cells = values as object[,];
            if (cells == null) {
                return values is int ? new ErrorWrapper((int)values) : values;
            }
            for (var row = cells.GetLowerBound(0); row <= cells.GetUpperBound(0); row++) {
                for (var column = cells.GetLowerBound(1); column <= cells.GetUpperBound(1); column++) {
                    if (cells[row, column] is int) {
                        cells[row, column] = new ErrorWrapper((int)cells[row, column]);
                    }
                }
            }
            return cells;
        }

    }
}