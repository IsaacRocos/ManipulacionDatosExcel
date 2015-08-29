using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace PruebaExcel
{
    class Program
    {
        private static Microsoft.Office.Interop.Excel.ApplicationClass appExcel;
        private static Workbook workBookExl = null;
        private static _Worksheet wSheet = null;


        static void Main(string[] args)
        {
            //excel_init("C:\\X.xlsx");
            excel_init("C:\\FORMATO_EJEMPLO.xls");
            int[] dimsTabla = excel_getTableSize();
            Console.WriteLine(dimsTabla[0] + " filas X " + dimsTabla[1] + " columnas");
            printRow(excel_getRow("A1","A"+ dimsTabla[1]));
            // lag
            Console.ReadKey();

            RunQuerys(dimsTabla[1], dimsTabla[0]);

            Console.ReadKey();
        }


        //Método para cargar un archivo de Excel
        static bool excel_init(String ruta)
        {
            appExcel = new ApplicationClass();
            if (System.IO.File.Exists(ruta)){
                workBookExl = appExcel.Workbooks.Open(ruta,0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",true, false, 0, true, false, false);
                wSheet = (_Worksheet)appExcel.ActiveWorkbook.ActiveSheet;
                return true;
            }else{
                Console.WriteLine("El documento " + ruta + " No puede abrirse");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
                appExcel = null;
                return false;
            }
        }


        //Método que Regresa las dimensiones de la tabla en excel
        static int[] excel_getTableSize() {
            int []filas_col = new int[2];
            wSheet.Columns.ClearFormats();
            wSheet.Rows.ClearFormats();
            filas_col[0] = wSheet.UsedRange.Rows.Count;
            filas_col[1] = wSheet.UsedRange.Columns.Count;
            return filas_col;
        }


        //Método para obtener un valor de una celda; El nombre de la celda puede ser A1,A2, B1,B2 etc... en excel.
        static string obtenerCelda(string cellname){
            string value = string.Empty;
            try{
                value = wSheet.get_Range(cellname).get_Value().ToString();
            }
            catch{
                value = "";
            }
            return value;
        }


        //Método para obtener un Rango de celdas en excel.
        static Microsoft.Office.Interop.Excel.Range excel_getRow(string celdainicio, string celdaFin){
            Console.WriteLine("Obteniendo rango: " + celdainicio + "," + celdaFin);
            Microsoft.Office.Interop.Excel.Range excelRow = (Microsoft.Office.Interop.Excel.Range)wSheet.get_Range(celdainicio, celdaFin);
            return excelRow;
        }

        // Método para imprimir una fila de excel
        static void printRow(Microsoft.Office.Interop.Excel.Range fila) {
         //algo va aquí   
        }

        // Método para crear una consulta 
        static String crearConsulta(int indiceFilaInicial, int numFilas, int numColumnas) {
            return "";
        }

        //Método principal para obtener las consultas de insersión en la base de datos
        static void RunQuerys(int numColumnas, int numFilas) {
            for (int i = 1; i <= numFilas; i++) {
                filaDeTabla(i, numColumnas);
                Console.WriteLine();
                Console.WriteLine();
            }
        }

        // Método para obtener todos los datos de una fila en un excel
        static void filaDeTabla(int numFila, int numColumnas) {
            char letraColumnaChar = 'A';
            ArrayList listaValores = new ArrayList();
            String celda = String.Empty;
            for (int letraColumna = letraColumnaChar; letraColumna<(letraColumnaChar+numColumnas);letraColumna++){
                celda = "" + (char)letraColumna + numFila;
                Console.Write(obtenerCelda(celda) + ",");
                listaValores.Add(celda);
            }
        }


        //Método para cerrar una conexión en excel
        static void excel_close()
        {
            if (appExcel != null){
                try{
                    workBookExl.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
                    appExcel = null;
                    wSheet = null;
                }catch (Exception ex){
                    appExcel = null;
                    Console.WriteLine("Ocurrieron problemas al intentar liberar los recursos: " + ex.ToString());
                }finally{
                    GC.Collect();
                }
            }
        }


    }
}
