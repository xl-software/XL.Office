using System;
using System.IO;
using System.Runtime.InteropServices;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace XL.Office
{
    /// <summary>
    /// Envoltorio que hace la vida más fácil a la hora de tratar con documentos Excel.
    /// Utiliza las librerías oficiales de Microsoft "Microsoft.Office.Interop.Excel".
    /// SE NECESITA TENER INSTALADO EXCEL
    /// 
    /// Ejemplo de uso:
    /// 
    /// <code>
    /// using (Excel excel = new Excel(@"Document.xlsx"))
    /// {
    ///     for (int row = 0; row < excel.Rows; row++)
    ///     {
    ///         for (int col = 0; col < excel.Cols; col++)
    ///         {
    ///             // Leer valor
    ///             var value = excel[row, col]; 
    ///             
    ///             // Escribir valor
    ///             excel[row, col] = "New value";
    ///         }
    ///     }
    ///     
    ///     // Guardar cambios
    ///     excel.Save();
    /// }
    /// </code>
    /// </summary>
    public class Excel : IDisposable
    {
        private bool _inDrive;
        private string _filePath;

        private readonly MSExcel.Application _xlApp;
        private readonly MSExcel.Workbook _xlWorkbook;

        private MSExcel._Worksheet _xlWorksheet;
        private MSExcel.Range _xlRange;

        /// <summary>
        /// Nº de filas del documento
        /// </summary>
        public int Rows { get; private set; }

        /// <summary>
        /// Nº de columnas por fila del documento
        /// </summary>
        public int Cols { get; private set; }

        /// <summary>
        /// Crea un Excel a partir de un documento del sistema
        /// </summary>
        /// <param name="path">Ruta del archivo</param>
        /// <param name="sheet">Número de la hoja</param>
        public Excel(string path, int sheet = 1)
        {
            _xlApp = new MSExcel.Application();

            _filePath = path;

            if (File.Exists(path))
            {
                _inDrive = true;
                _xlWorkbook = _xlApp.Workbooks.Open(path);
            }
            else
            {
                _xlWorkbook = _xlApp.Workbooks.Add(true);
            }

            SetSheet(sheet);
        }

        /// <summary>
        /// Guarda los cambios del documento
        /// </summary>
        public void Save()
        {
            if (_inDrive)
            {
                _xlWorkbook.Save();
            }
            else
            {
                _xlWorkbook.SaveAs(_filePath);
                _inDrive = true;
            }
        }

        /// <summary>
        /// Guarda los cambios del documento en la ruta especificada
        /// </summary>
        /// <param name="path">Ruta del sistema</param>
        public void SaveAs(string path)
        {
            _xlWorkbook.SaveAs(path);
            _filePath = path;
            _inDrive = true;
        }

        /// <summary>
        /// Asigna el número de hoja del documento sobre la que operar
        /// </summary>
        /// <param name="n">Nº de hoja del archivo</param>
        public void SetSheet(int n)
        {
            _xlWorksheet = _xlWorkbook.Sheets[n];
            _xlRange = _xlWorksheet.UsedRange;

            Rows = _xlRange.Rows.Count;
            Cols = _xlRange.Columns.Count;
        }

        /// <summary>
        /// Sobrecarga del operador [,]
        /// Utilizado para acceder a los valores del archivo vía [fila, columna]
        /// EJEMPLO: excel[0, 0] = Valor de Fila 1, Columna A
        /// </summary>
        /// <param name="row">Nº de fila</param>
        /// <param name="col">Nº de columna</param>
        /// <returns></returns>
        public object this[int row, int col]
        {
            get
            {
                MSExcel.Range cell = _xlRange.Cells[row + 1, col + 1];

                return cell?.Value2;
            }
            set
            {
                _xlRange.Cells[row + 1, col + 1].Value2 = value;
            }
        }

        /// <summary>
        /// Libera y cierra un documento Excel de memoria.
        /// Es muy importante utilizar este método al acabar de realizar las operaciones.
        /// Si no, se quedará un proceso en memoria y el excel se quedará en modo "Solo lectura".
        /// </summary>
        public void Dispose()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(_xlRange);
            Marshal.ReleaseComObject(_xlWorksheet);

            _xlWorkbook.Close();
            Marshal.ReleaseComObject(_xlWorkbook);

            _xlApp.Quit();
            Marshal.ReleaseComObject(_xlApp);
        }
    }
}
