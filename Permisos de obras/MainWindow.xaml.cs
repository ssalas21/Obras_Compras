using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Permisos_de_obras.Entity;
using Permisos_de_obras.BLL;

namespace Permisos_de_obras
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            List<Listado> lista2 = new List<Listado>();
            int anno = Convert.ToInt32(DateTime.Now.Year);
            for (int i = 2009; i <= anno; i++)
            {
                lista2.Add(new Listado(i, i));
            }
            CmbAnno.DisplayMemberPath = "AnnoD";
            CmbAnno.SelectedValuePath = "AnnoV";
            CmbAnno.ItemsSource = lista2;


        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\Pagina web\\Obras\\2021\\ley20898\\ley208982021muni.php"); // Abrir el txt
            // Cabeceras del html
            file.WriteLine("<? include (\"header.php\"); ?>");
            //Body del html
            file.WriteLine("<div class=\"row\">");
            file.WriteLine("<div class=\"col-sm-12\">");
            file.WriteLine("<center><b>LEY 20898 - 2021</b></center>");
            file.WriteLine("<P/>");
            file.WriteLine("<CENTER>");
            file.WriteLine("<style type=\"text/css\">");
            file.WriteLine(".tg  {border-collapse:collapse;border-spacing:0;border-color:#999;margin:0px auto;}");
            file.WriteLine(".tg td{font-family:Arial, sans-serif;font-size:10px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#444;background-color:#F7FDFA;}");
            file.WriteLine(".tg th{font-family:Arial, sans-serif;font-size:10px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#fff;background-color:#26ADE4;}");
            file.WriteLine(".tg .tg-yw4l{vertical-align:top}");
            file.WriteLine("th.tg-sort-header::-moz-selection { background:transparent; }th.tg-sort-header::selection      { background:transparent; }th.tg-sort-header { cursor:pointer; }table th.tg-sort-header:after {  content:'';  float:right;  margin-top:7px;  border-width:0 4px 4px;  border-style:solid;  border-color:#404040 transparent;  visibility:hidden;  }table th.tg-sort-header:hover:after {  visibility:visible;  }table th.tg-sort-desc:after,table th.tg-sort-asc:after,table th.tg-sort-asc:hover:after {  visibility:visible;  opacity:0.4;  }table th.tg-sort-desc:after {  border-bottom:none;  border-width:4px 4px 0;  }@media screen and (max-width: 767px) {.tg {width: auto !important;}.tg col {width: auto !important;}.tg-wrap {overflow-x: auto;-webkit-overflow-scrolling: touch;margin: auto 0px;}}</style>");
            //Crear tabla            
            file.WriteLine("<div class=\"tg-wrap\"><table id=\"tg-duF9v\" class=\"tg\">");
            //Cabecera de la tabla
            file.WriteLine("<tr>");
            file.WriteLine("<th class=\"tg-yw4l\">A&ntildeo</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Mes</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipologia del acto</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipo de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Denominaci&oacuten de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">N&uacutemero de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha de publicaci&oacuten en el DO (seg&uacuten Art. 45 y siguientes Ley 19.880)</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Indicaci&oacuten del medio y forma de publicidad</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tiene efectos generales</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha &uacuteltima actualizaci&oacuten</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Breve descripci&oacuten del objeto del acto</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace a la publicaci&oacuten o archivo correspondiente</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace a la modificaci&oacuten o archivo correspondiente</th>");
            file.WriteLine("</tr>");


            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\Pagina web\\Obras\\2021\\ley20898\\regley208982.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;

                //Llenar filas
                for (rCnt = 35; rCnt >= 8; rCnt--)
                {
                    string lines = "";
                    string tipo = "LEY 20.898";
                    int numero2 = Convert.ToInt32(((range.Cells[rCnt, 1] as Excel.Range).Value2));
                    int anno = Convert.ToInt32((range.Cells[rCnt, 2] as Excel.Range).Value2);
                    DateTime fecha = DateTime.FromOADate(((range.Cells[rCnt, 11] as Excel.Range).Value2));
                    string direccion = ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString();
                    string mes = "";
                    string numero = "";
                    if (numero2 < 10) numero = "00" + numero2;
                    if (numero2 > 9 && numero2 < 100) numero = "0" + numero2;
                    if (numero2 > 99) numero = "" + numero2;
                    switch (fecha.Month)
                    {
                        case 1:
                            mes = "ENERO";
                            break;
                        case 2:
                            mes = "FEBRERO";
                            break;
                        case 3:
                            mes = "MARZO";
                            break;
                        case 4:
                            mes = "ABRIL";
                            break;
                        case 5:
                            mes = "MAYO";
                            break;
                        case 6:
                            mes = "JUNIO";
                            break;
                        case 7:
                            mes = "JULIO";
                            break;
                        case 8:
                            mes = "AGOSTO";
                            break;
                        case 9:
                            mes = "SEPTIEMBRE";
                            break;
                        case 10:
                            mes = "OCTUBRE";
                            break;
                        case 11:
                            mes = "NOVIEMBRE";
                            break;
                        case 12:
                            mes = "DICIEMBRE";
                            break;
                    }
                    lines = "<tr>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + anno + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + mes + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">PERMISO DE OBRA</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">PERMISO</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + tipo + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + numero + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">PAGINA WEB INSTITUCIONAL</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + tipo + " EN " + direccion + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\" style=\"text-align:center;\"><a href=\"http://www.munipuchuncavi.cl/2.0/sitio10/pdf/transparencianew/permisosobras/leymono/2021/" + numero + ".pdf\"><img src=\"http://www.munipuchuncavi.cl/2.0/sitio10/images/pdf.jpg\" alt=\"Descargar\" style=\"width:25px;height:25px;\"></a></td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                    file.WriteLine(lines);
                    lines = "</tr>";
                    file.WriteLine(lines);
                }

                Marshal.ReleaseComObject(xlWorkSheet2);

            }

            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);


            //Cerrar tabla
            file.WriteLine("</table></div>");
            System.IO.StreamReader file2 = new System.IO.StreamReader("...\\...\\Script\\script.txt"); // Abrir el txt
            string linea;
            while ((linea = file2.ReadLine()) != null) file.WriteLine(linea);
            file.WriteLine();
            file.WriteLine("</CENTER>");
            file.WriteLine("</div>");
            file.WriteLine("</div>");

            //Footer
            file.WriteLine("<? include (\"footer.php\"); ?>");
            //Cerrar archivo
            file.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\Pagina web\\Obras\\2021\\permisos\\permisos2021muni.php"); // Abrir el txt
            // Cabeceras del html
            file.WriteLine("<? include (\"header.php\"); ?>");
            //Body del html
            file.WriteLine("<div class=\"row\">");
            file.WriteLine("<div class=\"col-sm-12\">");
            file.WriteLine("<center><b>PERMISOS EDIFICACI&OacuteN - 2021</b></center>");
            file.WriteLine("<P/>");
            file.WriteLine("<CENTER>");
            file.WriteLine("<style type=\"text/css\">");
            file.WriteLine(".tg  {border-collapse:collapse;border-spacing:0;border-color:#999;margin:0px auto;}");
            file.WriteLine(".tg td{font-family:Arial, sans-serif;font-size:10px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#444;background-color:#F7FDFA;}");
            file.WriteLine(".tg th{font-family:Arial, sans-serif;font-size:10px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#fff;background-color:#26ADE4;}");
            file.WriteLine(".tg .tg-yw4l{vertical-align:top}");
            file.WriteLine("th.tg-sort-header::-moz-selection { background:transparent; }th.tg-sort-header::selection      { background:transparent; }th.tg-sort-header { cursor:pointer; }table th.tg-sort-header:after {  content:'';  float:right;  margin-top:7px;  border-width:0 4px 4px;  border-style:solid;  border-color:#404040 transparent;  visibility:hidden;  }table th.tg-sort-header:hover:after {  visibility:visible;  }table th.tg-sort-desc:after,table th.tg-sort-asc:after,table th.tg-sort-asc:hover:after {  visibility:visible;  opacity:0.4;  }table th.tg-sort-desc:after {  border-bottom:none;  border-width:4px 4px 0;  }@media screen and (max-width: 767px) {.tg {width: auto !important;}.tg col {width: auto !important;}.tg-wrap {overflow-x: auto;-webkit-overflow-scrolling: touch;margin: auto 0px;}}</style>");
            //Crear tabla            
            file.WriteLine("<div class=\"tg-wrap\"><table id=\"tg-duF9v\" class=\"tg\">");
            //Cabecera de la tabla
            file.WriteLine("<tr>");
            file.WriteLine("<th class=\"tg-yw4l\">A&ntildeo</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Mes</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipologia del acto</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipo de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Denominaci&oacuten de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">N&uacutemero de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha de publicaci&oacuten en el DO (seg&uacuten Art. 45 y siguientes Ley 19.880)</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Indicaci&oacuten del medio y forma de publicidad</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tiene efectos generales</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha &uacuteltima actualizaci&oacuten</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Breve descripci&oacuten del objeto del acto</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace a la publicaci&oacuten o archivo correspondiente</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace a la modificaci&oacuten o archivo correspondiente</th>");
            file.WriteLine("</tr>");


            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\Pagina web\\Obras\\2021\\permisos\\permisos2020muni.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;

                //Llenar filas
                for (rCnt = 26; rCnt >= 6; rCnt--)
                {
                    string lines = "";
                    string tipo = ((range.Cells[rCnt, 1] as Excel.Range).Value2).ToString();
                    int numero2 = Convert.ToInt32(((range.Cells[rCnt, 2] as Excel.Range).Value2));
                    int anno = Convert.ToInt32((range.Cells[rCnt, 3] as Excel.Range).Value2);
                    DateTime fecha = DateTime.FromOADate(((range.Cells[rCnt, 14] as Excel.Range).Value2));
                    string direccion = ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                    string mes = "";
                    string numero = "";
                    if (numero2 < 10) numero = "00" + numero2;
                    if (numero2 > 9 && numero2 < 100) numero = "0" + numero2;
                    if (numero2 > 99) numero = numero2.ToString();
                    switch (fecha.Month)
                    {
                        case 1:
                            mes = "ENERO";
                            break;
                        case 2:
                            mes = "FEBRERO";
                            break;
                        case 3:
                            mes = "MARZO";
                            break;
                        case 4:
                            mes = "ABRIL";
                            break;
                        case 5:
                            mes = "MAYO";
                            break;
                        case 6:
                            mes = "JUNIO";
                            break;
                        case 7:
                            mes = "JULIO";
                            break;
                        case 8:
                            mes = "AGOSTO";
                            break;
                        case 9:
                            mes = "SEPTIEMBRE";
                            break;
                        case 10:
                            mes = "OCTUBRE";
                            break;
                        case 11:
                            mes = "NOVIEMBRE";
                            break;
                        case 12:
                            mes = "DICIEMBRE";
                            break;
                    }
                    lines = "<tr>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + anno + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + mes + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">PERMISO DE OBRA</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">PERMISO</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + tipo + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + numero + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">PAGINA WEB INSTITUCIONAL</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">AUTORIZA EJECUCION " + tipo + " EN " + direccion + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\" style=\"text-align:center;\"><a href=\"http://www.munipuchuncavi.cl/2.0/sitio10/pdf/transparencianew/permisosobras/permisos/2021/" + numero + ".pdf\"><img src=\"http://www.munipuchuncavi.cl/2.0/sitio10/images/pdf.jpg\" alt=\"Descargar\" style=\"width:25px;height:25px;\"></a></td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                    file.WriteLine(lines);
                    lines = "</tr>";
                    file.WriteLine(lines);
                }

                Marshal.ReleaseComObject(xlWorkSheet2);

            }

            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);


            //Cerrar tabla
            file.WriteLine("</table></div>");
            System.IO.StreamReader file2 = new System.IO.StreamReader("...\\...\\Script\\script.txt"); // Abrir el txt
            string linea;
            while ((linea = file2.ReadLine()) != null) file.WriteLine(linea);
            file.WriteLine();
            file.WriteLine("</CENTER>");
            file.WriteLine("</div>");
            file.WriteLine("</div>");
            //Footer
            file.WriteLine("<? include (\"footer.php\"); ?>");
            //Cerrar archivo
            file.Close();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\Pagina web\\Obras\\2021\\resoluciones\\resoluciones2021muni.php"); // Abrir el txt
            // Cabeceras del html
            file.WriteLine("<? include (\"header.php\"); ?>");
            //Body del html
            file.WriteLine("<div class=\"row\">");
            file.WriteLine("<div class=\"col-sm-12\">");
            file.WriteLine("<center><b>RESOLUCIONES - 2021</b></center>");
            file.WriteLine("<P/>");
            file.WriteLine("<CENTER>");
            file.WriteLine("<style type=\"text/css\">");
            file.WriteLine(".tg  {border-collapse:collapse;border-spacing:0;border-color:#999;margin:0px auto;}");
            file.WriteLine(".tg td{font-family:Arial, sans-serif;font-size:10px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#444;background-color:#F7FDFA;}");
            file.WriteLine(".tg th{font-family:Arial, sans-serif;font-size:10px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#fff;background-color:#26ADE4;}");
            file.WriteLine(".tg .tg-yw4l{vertical-align:top}");
            file.WriteLine("th.tg-sort-header::-moz-selection { background:transparent; }th.tg-sort-header::selection      { background:transparent; }th.tg-sort-header { cursor:pointer; }table th.tg-sort-header:after {  content:'';  float:right;  margin-top:7px;  border-width:0 4px 4px;  border-style:solid;  border-color:#404040 transparent;  visibility:hidden;  }table th.tg-sort-header:hover:after {  visibility:visible;  }table th.tg-sort-desc:after,table th.tg-sort-asc:after,table th.tg-sort-asc:hover:after {  visibility:visible;  opacity:0.4;  }table th.tg-sort-desc:after {  border-bottom:none;  border-width:4px 4px 0;  }@media screen and (max-width: 767px) {.tg {width: auto !important;}.tg col {width: auto !important;}.tg-wrap {overflow-x: auto;-webkit-overflow-scrolling: touch;margin: auto 0px;}}</style>");
            //Crear tabla            
            file.WriteLine("<div class=\"tg-wrap\"><table id=\"tg-duF9v\" class=\"tg\">");
            //Cabecera de la tabla
            file.WriteLine("<tr>");
            file.WriteLine("<th class=\"tg-yw4l\">A&ntildeo</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Mes</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipologia del acto</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipo de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Denominaci&oacuten de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">N&uacutemero de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha de publicaci&oacuten en el DO (seg&uacuten Art. 45 y siguientes Ley 19.880)</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Indicaci&oacuten del medio y forma de publicidad</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tiene efectos generales</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha &uacuteltima actualizaci&oacuten</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Breve descripci&oacuten del objeto del acto</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace a la publicaci&oacuten o archivo correspondiente</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace a la modificaci&oacuten o archivo correspondiente</th>");
            file.WriteLine("</tr>");


            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\Pagina web\\Obras\\2021\\resoluciones\\subdivisiones-2020.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;

                //Llenar filas
                for (rCnt = 76; rCnt >= 5; rCnt--)
                {
                    string lines = "";
                    string tipo = ((range.Cells[rCnt, 9] as Excel.Range).Value2).ToString();
                    int numero2 = Convert.ToInt32(((range.Cells[rCnt, 1] as Excel.Range).Value2));
                    int anno = Convert.ToInt32((range.Cells[rCnt, 2] as Excel.Range).Value2);
                    DateTime fecha = DateTime.FromOADate(((range.Cells[rCnt, 10] as Excel.Range).Value2));
                    string direccion1 = "";
                    string direccion2 = "";
                    try
                    {
                        direccion1 = ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString();
                    }
                    catch (Exception)
                    { }
                    try
                    {
                        direccion2 = ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                    }
                    catch (Exception)
                    {
                    }
                    string direccion = "";
                    if (direccion1.Equals("") && direccion2.Equals("")) direccion = "SIN DIRECCI&OacuteN REGISTRADA";
                    else direccion = "EN " + direccion1 + ", " + direccion2;
                    string mes = "";
                    string numero = "";
                    if (numero2 < 10) numero = "00" + numero2;
                    if (numero2 > 9 && numero2 < 100) numero = "0" + numero2;
                    if (numero2 > 99) numero = numero2.ToString();
                    switch (fecha.Month)
                    {
                        case 1:
                            mes = "ENERO";
                            break;
                        case 2:
                            mes = "FEBRERO";
                            break;
                        case 3:
                            mes = "MARZO";
                            break;
                        case 4:
                            mes = "ABRIL";
                            break;
                        case 5:
                            mes = "MAYO";
                            break;
                        case 6:
                            mes = "JUNIO";
                            break;
                        case 7:
                            mes = "JULIO";
                            break;
                        case 8:
                            mes = "AGOSTO";
                            break;
                        case 9:
                            mes = "SEPTIEMBRE";
                            break;
                        case 10:
                            mes = "OCTUBRE";
                            break;
                        case 11:
                            mes = "NOVIEMBRE";
                            break;
                        case 12:
                            mes = "DICIEMBRE";
                            break;
                    }
                    lines = "<tr>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + anno + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + mes + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">PERMISO DE OBRA</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">PERMISO</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + tipo + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + numero + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">PAGINA WEB INSTITUCIONAL</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + tipo + " " + direccion + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\" style=\"text-align:center;\"><a href=\"http://www.munipuchuncavi.cl/2.0/sitio10/pdf/transparencianew/permisosobras/resoluciones/2021/RE" + numero + ".pdf\"><img src=\"http://www.munipuchuncavi.cl/2.0/sitio10/images/pdf.jpg\" alt=\"Descargar\" style=\"width:25px;height:25px;\"></a></td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                    file.WriteLine(lines);
                    lines = "</tr>";
                    file.WriteLine(lines);
                }

                Marshal.ReleaseComObject(xlWorkSheet2);

            }

            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);


            //Cerrar tabla
            file.WriteLine("</table></div>");
            System.IO.StreamReader file2 = new System.IO.StreamReader("...\\...\\Script\\script.txt"); // Abrir el txt
            string linea;
            while ((linea = file2.ReadLine()) != null) file.WriteLine(linea);
            file.WriteLine();
            file.WriteLine("</CENTER>");
            file.WriteLine("</div>");
            file.WriteLine("</div>");
            //Footer
            file.WriteLine("<? include (\"footer.php\"); ?>");
            //Cerrar archivo
            file.Close();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\Pagina web\\Obras\\2021\\recepciones\\recepciones2021muni.php"); // Abrir el txt
            // Cabeceras del html
            file.WriteLine("<? include (\"header.php\"); ?>");
            //Body del html
            file.WriteLine("<div class=\"row\">");
            file.WriteLine("<div class=\"col-sm-12\">");
            file.WriteLine("<center><b>RECEPCIONES - 2021</b></center>");
            file.WriteLine("<P/>");
            file.WriteLine("<CENTER>");
            file.WriteLine("<style type=\"text/css\">");
            file.WriteLine(".tg  {border-collapse:collapse;border-spacing:0;border-color:#999;margin:0px auto;}");
            file.WriteLine(".tg td{font-family:Arial, sans-serif;font-size:10px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#444;background-color:#F7FDFA;}");
            file.WriteLine(".tg th{font-family:Arial, sans-serif;font-size:10px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#fff;background-color:#26ADE4;}");
            file.WriteLine(".tg .tg-yw4l{vertical-align:top}");
            file.WriteLine("th.tg-sort-header::-moz-selection { background:transparent; }th.tg-sort-header::selection      { background:transparent; }th.tg-sort-header { cursor:pointer; }table th.tg-sort-header:after {  content:'';  float:right;  margin-top:7px;  border-width:0 4px 4px;  border-style:solid;  border-color:#404040 transparent;  visibility:hidden;  }table th.tg-sort-header:hover:after {  visibility:visible;  }table th.tg-sort-desc:after,table th.tg-sort-asc:after,table th.tg-sort-asc:hover:after {  visibility:visible;  opacity:0.4;  }table th.tg-sort-desc:after {  border-bottom:none;  border-width:4px 4px 0;  }@media screen and (max-width: 767px) {.tg {width: auto !important;}.tg col {width: auto !important;}.tg-wrap {overflow-x: auto;-webkit-overflow-scrolling: touch;margin: auto 0px;}}</style>");
            //Crear tabla            
            file.WriteLine("<div class=\"tg-wrap\"><table id=\"tg-duF9v\" class=\"tg\">");
            //Cabecera de la tabla
            file.WriteLine("<tr>");
            file.WriteLine("<th class=\"tg-yw4l\">A&ntildeo</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Mes</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipologia del acto</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipo de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Denominaci&oacuten de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">N&uacutemero de norma</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha de publicaci&oacuten en el DO (seg&uacuten Art. 45 y siguientes Ley 19.880)</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Indicaci&oacuten del medio y forma de publicidad</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tiene efectos generales</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha &uacuteltima actualizaci&oacuten</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Breve descripci&oacuten del objeto del acto</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace a la publicaci&oacuten o archivo correspondiente</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace a la modificaci&oacuten o archivo correspondiente</th>");
            file.WriteLine("</tr>");


            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\Pagina web\\Obras\\2021\\recepciones\\recepciones2020muni.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;


                //Llenar filas
                for (rCnt = 18; rCnt >= 3; rCnt--)
                {
                    string lines = "";
                    int numero2 = Convert.ToInt32(((range.Cells[rCnt, 1] as Excel.Range).Value2));
                    int anno = Convert.ToInt32((range.Cells[rCnt, 2] as Excel.Range).Value2);
                    DateTime fecha = DateTime.FromOADate(((range.Cells[rCnt, 10] as Excel.Range).Value2));
                    string direccion = ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString();
                    string mes = "";
                    string numero = "";
                    if (numero2 < 10) numero = "00" + numero2;
                    if (numero2 > 9 && numero2 < 100) numero = "0" + numero2;
                    if (numero2 > 99) numero = numero2.ToString();
                    switch (fecha.Month)
                    {
                        case 1:
                            mes = "ENERO";
                            break;
                        case 2:
                            mes = "FEBRERO";
                            break;
                        case 3:
                            mes = "MARZO";
                            break;
                        case 4:
                            mes = "ABRIL";
                            break;
                        case 5:
                            mes = "MAYO";
                            break;
                        case 6:
                            mes = "JUNIO";
                            break;
                        case 7:
                            mes = "JULIO";
                            break;
                        case 8:
                            mes = "AGOSTO";
                            break;
                        case 9:
                            mes = "SEPTIEMBRE";
                            break;
                        case 10:
                            mes = "OCTUBRE";
                            break;
                        case 11:
                            mes = "NOVIEMBRE";
                            break;
                        case 12:
                            mes = "DICIEMBRE";
                            break;
                    }
                    lines = "<tr>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + anno + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + mes + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">PERMISO DE OBRA</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">CERTIFICADO</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">CERTIFICADO DE RECEPCI&OacuteN</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + numero + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">PAGINA WEB INSTITUCIONAL</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">CERTIFICADO DE RECEPCION EN " + direccion + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\" style=\"text-align:center;\"><a href=\"http://www.munipuchuncavi.cl/2.0/sitio10/pdf/transparencianew/permisosobras/recepciones/2021/" + numero + ".pdf\"><img src=\"http://www.munipuchuncavi.cl/2.0/sitio10/images/pdf.jpg\" alt=\"Descargar\" style=\"width:25px;height:25px;\"></a></td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                    file.WriteLine(lines);
                    lines = "</tr>";
                    file.WriteLine(lines);
                }

                Marshal.ReleaseComObject(xlWorkSheet2);

            }

            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);


            //Cerrar tabla
            file.WriteLine("</table></div>");
            System.IO.StreamReader file2 = new System.IO.StreamReader("...\\...\\Script\\script.txt"); // Abrir el txt
            string linea;
            while ((linea = file2.ReadLine()) != null) file.WriteLine(linea);
            file.WriteLine();
            file.WriteLine("</CENTER>");
            file.WriteLine("</div>");
            file.WriteLine("</div>");
            //Footer
            file.WriteLine("<? include (\"footer.php\"); ?>");
            //Cerrar archivo
            file.Close();
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\Pagina web\\Compras menores a 3 utm\\PDF\\Ordenesdecompra2017.php"); // Abrir el txt
            // Cabeceras del html
            file.WriteLine("<? include (\"header.php\"); ?>");
            //Body del html
            file.WriteLine("<div class=\"row\">");
            file.WriteLine("<div class=\"col-sm-12\">");
            file.WriteLine("<center><b>ORDENES DE COMPRA - 2017</b></center>");
            file.WriteLine("<P/>");
            file.WriteLine("<CENTER>");
            file.WriteLine("<style type=\"text/css\">");
            file.WriteLine(".tg  {border-collapse:collapse;border-spacing:0;border-color:#999;margin:0px auto;}");
            file.WriteLine(".tg td{font-family:Arial, sans-serif;font-size:10px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#444;background-color:#F7FDFA;}");
            file.WriteLine(".tg th{font-family:Arial, sans-serif;font-size:10px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#fff;background-color:#26ADE4;}");
            file.WriteLine(".tg .tg-yw4l{vertical-align:top}");
            file.WriteLine("th.tg-sort-header::-moz-selection { background:transparent; }th.tg-sort-header::selection      { background:transparent; }th.tg-sort-header { cursor:pointer; }table th.tg-sort-header:after {  content:'';  float:right;  margin-top:7px;  border-width:0 4px 4px;  border-style:solid;  border-color:#404040 transparent;  visibility:hidden;  }table th.tg-sort-header:hover:after {  visibility:visible;  }table th.tg-sort-desc:after,table th.tg-sort-asc:after,table th.tg-sort-asc:hover:after {  visibility:visible;  opacity:0.4;  }table th.tg-sort-desc:after {  border-bottom:none;  border-width:4px 4px 0;  }@media screen and (max-width: 767px) {.tg {width: auto !important;}.tg col {width: auto !important;}.tg-wrap {overflow-x: auto;-webkit-overflow-scrolling: touch;margin: auto 0px;}}</style>");
            //Crear tabla            
            file.WriteLine("<div class=\"tg-wrap\"><table id=\"tg-duF9v\" class=\"tg\">");
            //Cabecera de la tabla
            file.WriteLine("<tr>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipo de acto administrativo aprobatorio</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Denominaci&oacuten del acto administrativo</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha del acto administrativo aprobatorio del contrato</th>");
            file.WriteLine("<th class=\"tg-yw4l\">N&uacutemero del acto administrativo aprobatorio</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Nombre completo o raz&oacuten social de la persona contratada</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Rut de la persona contratada (Si aplica)</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Socios y accionistas principales (Si corresponde)</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Objeto de la contrataci&oacuten o adquisiciones</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Monto total de la operaci&oacuten</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha de inicio del contrato</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha de t&eacutermino del contrato</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace al texto &iacutentegrp del contrato</th>");
            file.WriteLine("</tr>");


            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\Pagina web\\Compras menores a 3 utm\\PDF\\2017\\Ordenesdecompra2017.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;


                //Llenar filas
                for (rCnt = 2; rCnt <= 349; rCnt++)
                {
                    string lines = "";
                    int numero = Convert.ToInt32(((range.Cells[rCnt, 1] as Excel.Range).Value2));
                    DateTime fecha = DateTime.FromOADate(((range.Cells[rCnt, 2] as Excel.Range).Value2));
                    string nombre = ((range.Cells[rCnt, 3] as Excel.Range).Value2).ToString();
                    int rut = Convert.ToInt32((range.Cells[rCnt, 4] as Excel.Range).Value2);
                    string digito = ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                    string detalle = ((range.Cells[rCnt, 6] as Excel.Range).Value2).ToString();
                    int precio = Convert.ToInt32((range.Cells[rCnt, 7] as Excel.Range).Value2);
                    DateTime fecha2 = fecha.AddMonths(1);
                    lines = "<tr>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">ORDEN DE COMPRA DAF</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">ORDEN N " + numero + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + numero + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + nombre + "</td>";
                    file.WriteLine(lines);
                    if (rut < 30000000)
                    {
                        lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                        file.WriteLine(lines);
                    }
                    else
                    {
                        lines = "<td class=\"tg-yw41\">" + rut + "-" + digito + "</td>";
                        file.WriteLine(lines);
                    }
                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + detalle + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">$" + precio + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha2.Day + "/" + fecha2.Month + "/" + fecha2.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\" style=\"text-align:center;\"><a href=\"http://www.munipuchuncavi.cl/2.0/sitio10/pdf/transparencianew/comprasmenores/2017/OC" + numero + "-17.pdf\"><img src=\"http://www.munipuchuncavi.cl/2.0/sitio10/images/pdf.jpg\" alt=\"Descargar\" style=\"width:25px;height:25px;\"></a></td>";
                    file.WriteLine(lines);
                    lines = "</tr>";
                    file.WriteLine(lines);
                }

                Marshal.ReleaseComObject(xlWorkSheet2);

            }

            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);


            //Cerrar tabla
            file.WriteLine("</table></div>");
            System.IO.StreamReader file2 = new System.IO.StreamReader("...\\...\\Script\\script.txt"); // Abrir el txt
            string linea;
            while ((linea = file2.ReadLine()) != null) file.WriteLine(linea);
            file.WriteLine();
            file.WriteLine("</CENTER>");
            file.WriteLine("</div>");
            file.WriteLine("</div>");
            //Footer
            file.WriteLine("<? include (\"footer.php\"); ?>");
            //Cerrar archivo
            file.Close();
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\Pagina web\\Compras menores a 3 utm\\PDF\\Ordenesdecompra2018.php"); // Abrir el txt
            // Cabeceras del html
            file.WriteLine("<? include (\"header.php\"); ?>");
            //Body del html
            file.WriteLine("<div id=\"contenido\">");
            file.WriteLine("<center><b>ORDENES DE COMPRA - 2018</b></center>");
            file.WriteLine("<P/>");
            file.WriteLine("<CENTER>");
            file.WriteLine("<style type=\"text/css\">");
            file.WriteLine(".tg  {border-collapse:collapse;border-spacing:0;border-color:#999;margin:0px auto;}");
            file.WriteLine(".tg td{font-family:Arial, sans-serif;font-size:10px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#444;background-color:#F7FDFA;}");
            file.WriteLine(".tg th{font-family:Arial, sans-serif;font-size:10px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#fff;background-color:#26ADE4;}");
            file.WriteLine(".tg .tg-yw4l{vertical-align:top}");
            file.WriteLine("th.tg-sort-header::-moz-selection { background:transparent; }th.tg-sort-header::selection      { background:transparent; }th.tg-sort-header { cursor:pointer; }table th.tg-sort-header:after {  content:'';  float:right;  margin-top:7px;  border-width:0 4px 4px;  border-style:solid;  border-color:#404040 transparent;  visibility:hidden;  }table th.tg-sort-header:hover:after {  visibility:visible;  }table th.tg-sort-desc:after,table th.tg-sort-asc:after,table th.tg-sort-asc:hover:after {  visibility:visible;  opacity:0.4;  }table th.tg-sort-desc:after {  border-bottom:none;  border-width:4px 4px 0;  }@media screen and (max-width: 767px) {.tg {width: auto !important;}.tg col {width: auto !important;}.tg-wrap {overflow-x: auto;-webkit-overflow-scrolling: touch;margin: auto 0px;}}</style>");
            //Crear tabla            
            file.WriteLine("<div class=\"tg-wrap\"><table id=\"tg-duF9v\" class=\"tg\">");
            //Cabecera de la tabla
            file.WriteLine("<tr>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipo de acto administrativo aprobatorio</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Denominaci&oacuten del acto administrativo</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha del acto administrativo aprobatorio del contrato</th>");
            file.WriteLine("<th class=\"tg-yw4l\">N&uacutemero del acto administrativo aprobatorio</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Nombre completo o raz&oacuten social de la persona contratada</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Rut de la persona contratada (Si aplica)</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Socios y accionistas principales (Si corresponde)</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Objeto de la contrataci&oacuten o adquisiciones</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Monto total de la operaci&oacuten</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha de inicio del contrato</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha de t&eacutermino del contrato</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace al texto &iacutentegrp del contrato</th>");
            file.WriteLine("</tr>");


            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\Pagina web\\Compras menores a 3 utm\\PDF\\2018\\Ordenesdecompra2018.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;


                //Llenar filas
                for (rCnt = 2; rCnt <= 219; rCnt++)
                {
                    string lines = "";
                    int numero = Convert.ToInt32(((range.Cells[rCnt, 1] as Excel.Range).Value2));
                    DateTime fecha = DateTime.FromOADate(((range.Cells[rCnt, 2] as Excel.Range).Value2));
                    string nombre = ((range.Cells[rCnt, 3] as Excel.Range).Value2).ToString();
                    int rut = Convert.ToInt32((range.Cells[rCnt, 4] as Excel.Range).Value2);
                    string digito = ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                    string detalle = ((range.Cells[rCnt, 6] as Excel.Range).Value2).ToString();
                    int precio = Convert.ToInt32((range.Cells[rCnt, 7] as Excel.Range).Value2);
                    DateTime fecha2 = fecha.AddMonths(1);
                    lines = "<tr>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">ORDEN DE COMPRA DAF</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">ORDEN N " + numero + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + numero + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + nombre + "</td>";
                    file.WriteLine(lines);
                    if (rut < 30000000)
                    {
                        lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                        file.WriteLine(lines);
                    }
                    else
                    {
                        lines = "<td class=\"tg-yw41\">" + rut + "-" + digito + "</td>";
                        file.WriteLine(lines);
                    }
                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + detalle + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">$" + precio + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\">" + fecha2.Day + "/" + fecha2.Month + "/" + fecha2.Year + "</td>";
                    file.WriteLine(lines);
                    lines = "<td class=\"tg-yw41\" style=\"text-align:center;\"><a href=\"http://www.munipuchuncavi.cl/2.0/sitio10/pdf/transparencianew/comprasmenores/2018/OC" + numero + "-18.pdf\"><img src=\"http://www.munipuchuncavi.cl/2.0/sitio10/images/pdf.jpg\" alt=\"Descargar\" style=\"width:25px;height:25px;\"></a></td>";
                    file.WriteLine(lines);
                    lines = "</tr>";
                    file.WriteLine(lines);
                }

                Marshal.ReleaseComObject(xlWorkSheet2);

            }

            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);


            //Cerrar tabla
            file.WriteLine("</table></div>");
            System.IO.StreamReader file2 = new System.IO.StreamReader("...\\...\\Script\\script.txt"); // Abrir el txt
            string linea;
            while ((linea = file2.ReadLine()) != null) file.WriteLine(linea);
            file.WriteLine();
            file.WriteLine("</CENTER>");
            file.WriteLine("</div>");
            //Footer
            file.WriteLine("<? include (\"footer.php\"); ?>");
            //Cerrar archivo
            file.Close();
        }

        private void BtnOrden_Click(object sender, RoutedEventArgs e)
        {
            int anno = Convert.ToInt32(CmbAnno.SelectedValue);
            string year = anno.ToString();
            string aux = "";
            for (int i = 2; i < year.Length; i++)
            {
                aux = aux + year[i];
            }
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\Pagina web\\Compras menores a 3 utm\\orden\\ordenes\\Ordenesdecompra" + anno + ".php"); // Abrir el txt
            System.IO.StreamWriter filemayor = new System.IO.StreamWriter("D:\\Pagina web\\Compras menores a 3 utm\\orden\\ordenes\\Ordenesdecompramayor" + anno + ".php"); // Abrir el txt
            // Cabeceras del html
            file.WriteLine("<? include (\"header.php\"); ?>");
            //Body del html
            file.WriteLine("<div class=\"row\">");
            file.WriteLine("<div class=\"col-sm-12\">");
            file.WriteLine("<center><b>ORDENES DE COMPRA - " + anno + "</b></center>");
            file.WriteLine("<P/>");
            file.WriteLine("<CENTER>");
            file.WriteLine("<style type=\"text/css\">");
            file.WriteLine(".tg  {border-collapse:collapse;border-spacing:0;border-color:#999;margin:0px auto;}");
            file.WriteLine(".tg td{font-family:Arial, sans-serif;font-size:10px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#444;background-color:#F7FDFA;}");
            file.WriteLine(".tg th{font-family:Arial, sans-serif;font-size:10px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#fff;background-color:#26ADE4;}");
            file.WriteLine(".tg .tg-yw4l{vertical-align:top}");
            file.WriteLine("th.tg-sort-header::-moz-selection { background:transparent; }th.tg-sort-header::selection      { background:transparent; }th.tg-sort-header { cursor:pointer; }table th.tg-sort-header:after {  content:'';  float:right;  margin-top:7px;  border-width:0 4px 4px;  border-style:solid;  border-color:#404040 transparent;  visibility:hidden;  }table th.tg-sort-header:hover:after {  visibility:visible;  }table th.tg-sort-desc:after,table th.tg-sort-asc:after,table th.tg-sort-asc:hover:after {  visibility:visible;  opacity:0.4;  }table th.tg-sort-desc:after {  border-bottom:none;  border-width:4px 4px 0;  }@media screen and (max-width: 767px) {.tg {width: auto !important;}.tg col {width: auto !important;}.tg-wrap {overflow-x: auto;-webkit-overflow-scrolling: touch;margin: auto 0px;}}</style>");
            //Crear tabla            
            file.WriteLine("<div class=\"tg-wrap\"><table id=\"tg-duF9v\" class=\"tg\">");
            //Cabecera de la tabla
            file.WriteLine("<tr>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipo de Compra</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Tipo de acto administrativo aprobatorio</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Denominaci&oacuten del acto administrativo</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha del acto administrativo aprobatorio del contrato</th>");
            file.WriteLine("<th class=\"tg-yw4l\">N&uacutemero del acto administrativo aprobatorio</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Nombre completo o raz&oacuten social de la persona contratada</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Rut de la persona contratada (Si aplica)</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Socios y accionistas principales (Si corresponde)</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Objeto de la contrataci&oacuten o adquisiciones</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Monto total de la operaci&oacuten</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha de inicio del contrato</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Fecha de t&eacutermino del contrato</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace al texto &iacutentegro del contrato</th>");
            file.WriteLine("<th class=\"tg-yw4l\">Enlace al texto &iacutentegro del acto administrativo aprobatorio</th>");
            //file.WriteLine("<th class=\"tg-yw4l\">Enlace al texto &iacutentegro del acto administrativo aprobatorio de la modificaci&oacuten</th>");
            file.WriteLine("</tr>");

            // Cabeceras del html
            filemayor.WriteLine("<? include (\"header.php\"); ?>");
            //Body del html
            filemayor.WriteLine("<div class=\"row\">");
            filemayor.WriteLine("<div class=\"col-sm-12\">");
            filemayor.WriteLine("<center><b>ORDENES DE COMPRA - " + anno + "</b></center>");
            filemayor.WriteLine("<P/>");
            filemayor.WriteLine("<CENTER>");
            filemayor.WriteLine("<style type=\"text/css\">");
            filemayor.WriteLine(".tg  {border-collapse:collapse;border-spacing:0;border-color:#999;margin:0px auto;}");
            filemayor.WriteLine(".tg td{font-family:Arial, sans-serif;font-size:10px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#444;background-color:#F7FDFA;}");
            filemayor.WriteLine(".tg th{font-family:Arial, sans-serif;font-size:10px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:#999;color:#fff;background-color:#26ADE4;}");
            filemayor.WriteLine(".tg .tg-yw4l{vertical-align:top}");
            filemayor.WriteLine("th.tg-sort-header::-moz-selection { background:transparent; }th.tg-sort-header::selection      { background:transparent; }th.tg-sort-header { cursor:pointer; }table th.tg-sort-header:after {  content:'';  float:right;  margin-top:7px;  border-width:0 4px 4px;  border-style:solid;  border-color:#404040 transparent;  visibility:hidden;  }table th.tg-sort-header:hover:after {  visibility:visible;  }table th.tg-sort-desc:after,table th.tg-sort-asc:after,table th.tg-sort-asc:hover:after {  visibility:visible;  opacity:0.4;  }table th.tg-sort-desc:after {  border-bottom:none;  border-width:4px 4px 0;  }@media screen and (max-width: 767px) {.tg {width: auto !important;}.tg col {width: auto !important;}.tg-wrap {overflow-x: auto;-webkit-overflow-scrolling: touch;margin: auto 0px;}}</style>");
            //Crear tabla            
            filemayor.WriteLine("<div class=\"tg-wrap\"><table id=\"tg-duF9v\" class=\"tg\">");
            //Cabecera de la tabla
            filemayor.WriteLine("<tr>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Tipo de Compra</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Tipo de acto administrativo aprobatorio</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Denominaci&oacuten del acto administrativo</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Fecha del acto administrativo aprobatorio del contrato</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">N&uacutemero del acto administrativo aprobatorio</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Nombre completo o raz&oacuten social de la persona contratada</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Rut de la persona contratada (Si aplica)</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Socios y accionistas principales (Si corresponde)</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Objeto de la contrataci&oacuten o adquisiciones</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Monto total de la operaci&oacuten</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Fecha de inicio del contrato</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Fecha de t&eacutermino del contrato</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Enlace al texto &iacutentegro del contrato</th>");
            filemayor.WriteLine("<th class=\"tg-yw4l\">Enlace al texto &iacutentegro del acto administrativo aprobatorio</th>");
            //filemayor.WriteLine("<th class=\"tg-yw4l\">Enlace al texto &iacutentegro del acto administrativo aprobatorio de la modificaci&oacuten</th>");
            filemayor.WriteLine("</tr>");


            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\Pagina web\\Compras menores a 3 utm\\orden\\ordenes\\ordenes.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;


                //Llenar filas
                for (rCnt = 2; rCnt <= 10152; rCnt++)
                {
                    DateTime fecha = DateTime.FromOADate(((range.Cells[rCnt, 2] as Excel.Range).Value2));
                    if (fecha.Year == anno)
                    {
                        string lines = "";
                        int numero = Convert.ToInt32(((range.Cells[rCnt, 1] as Excel.Range).Value2));
                        string nombre = ((range.Cells[rCnt, 3] as Excel.Range).Value2).ToString();
                        int rut = Convert.ToInt32((range.Cells[rCnt, 4] as Excel.Range).Value2);
                        string digito = ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                        string detalle = ((range.Cells[rCnt, 6] as Excel.Range).Value2).ToString();
                        int precio = Convert.ToInt32((range.Cells[rCnt, 7] as Excel.Range).Value2);
                        DateTime fecha2 = fecha.AddMonths(1);
                        Utm aux2 = (new UtmBLL()).ObtenerMesAnno(Convert.ToInt32(fecha.Year.ToString()), Convert.ToInt32(fecha.Month.ToString()));
                        int valor = 0;
                        valor = (int)aux2.Valor * 3;
                        if (System.IO.File.Exists("D:\\Pagina web\\Compras menores a 3 utm\\orden\\ordenes\\" + anno + "\\OC" + numero + "-" + aux + ".pdf"))
                        {
                            if (precio <= valor)
                            {
                                lines = "<tr>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">COMPRAS MENORES A 3 UTM</td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">ORDEN DE COMPRA DAF</td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">ORDEN N " + numero + "</td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + numero + "</td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + nombre + "</td>";
                                file.WriteLine(lines);
                                if (rut < 30000000)
                                {
                                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                                    file.WriteLine(lines);
                                }
                                else
                                {
                                    lines = "<td class=\"tg-yw41\">" + rut + "-" + digito + "</td>";
                                    file.WriteLine(lines);
                                }
                                lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + detalle + "</td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">$" + precio + "</td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + fecha2.Day + "/" + fecha2.Month + "/" + fecha2.Year + "</td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\" style=\"text-align:center;\"><a href=\"pdf/transparencianew/comprasmenores/" + anno + "/OC" + numero + "-" + aux + ".pdf\"><img src=\"images/pdf.jpg\" alt=\"Descargar\" style=\"width:25px;height:25px;\"></a></td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\" style=\"text-align:center;\"><a href=\"pdf/transparencianew/comprasmenores/" + anno + "/OC" + numero + "-" + aux + ".pdf\"><img src=\"images/pdf.jpg\" alt=\"Descargar\" style=\"width:25px;height:25px;\"></a></td>";
                                file.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                                file.WriteLine(lines);
                                lines = "</tr>";
                                file.WriteLine(lines);
                            }
                            else
                            {
                                lines = "<tr>";
                                filemayor.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">COMPRAS MAYORES A 3 UTM</td>";
                                filemayor.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">ORDEN DE COMPRA DAF</td>";
                                filemayor.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">ORDEN N " + numero + "</td>";
                                filemayor.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                                filemayor.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + numero + "</td>";
                                filemayor.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + nombre + "</td>";
                                filemayor.WriteLine(lines);
                                if (rut < 30000000)
                                {
                                    lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                                    filemayor.WriteLine(lines);
                                }
                                else
                                {
                                    lines = "<td class=\"tg-yw41\">" + rut + "-" + digito + "</td>";
                                    filemayor.WriteLine(lines);
                                }
                                lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                                filemayor.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + detalle + "</td>";
                                filemayor.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">$" + precio + "</td>";
                                filemayor.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + fecha.Day + "/" + fecha.Month + "/" + fecha.Year + "</td>";
                                filemayor.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\">" + fecha2.Day + "/" + fecha2.Month + "/" + fecha2.Year + "</td>";
                                filemayor.WriteLine(lines);
                                lines = "<td class=\"tg-yw41\" style=\"text-align:center;\"><a href=\"pdf/transparencianew/comprasmenores/" + anno + "/OC" + numero + "-" + aux + ".pdf\"><img src=\"images/pdf.jpg\" alt=\"Descargar\" style=\"width:25px;height:25px;\"></a></td>";
                                filemayor.WriteLine(lines);
                                int decreto;
                                try
                                {
                                    decreto = Convert.ToInt32((range.Cells[rCnt, 8] as Excel.Range).Value2);
                                }
                                catch (Exception)
                                {
                                    decreto = 0;
                                }
                                if (System.IO.File.Exists("D:\\Pagina web\\Decretos\\" + anno + "\\DA" + decreto + "-" + aux + ".pdf"))
                                {
                                    lines = "<td class=\"tg-yw41\" style=\"text-align:center;\"><a href=\"pdf/transparencianew/decretos/" + anno + "/DA" + decreto + "-" + aux + ".pdf\"><img src=\"images/pdf.jpg\" alt=\"Descargar\" style=\"width:25px;height:25px;\"></a></td>";
                                    filemayor.WriteLine(lines);
                                }
                                else {
                                    lines = "<td class=\"tg-yw41\" style=\"text-align:center;\"><a href=\"pdf/transparencianew/comprasmenores/" + anno + "/OC" + numero + "-" + aux + ".pdf\"><img src=\"images/pdf.jpg\" alt=\"Descargar\" style=\"width:25px;height:25px;\"></a></td>";
                                    filemayor.WriteLine(lines);
                                }
                                
                                lines = "<td class=\"tg-yw41\">NO APLICA</td>";
                                filemayor.WriteLine(lines);
                                lines = "</tr>";
                                filemayor.WriteLine(lines);
                            }
                        }
                    }
                }

                Marshal.ReleaseComObject(xlWorkSheet2);

            }

            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);


            //Cerrar tabla
            file.WriteLine("</table></div>");
            filemayor.WriteLine("</table></div>");
            System.IO.StreamReader file2 = new System.IO.StreamReader("...\\...\\Script\\script.txt"); // Abrir el txt
            string linea;
            while ((linea = file2.ReadLine()) != null)
            {
                file.WriteLine(linea);
                filemayor.WriteLine(linea);
            }
            file.WriteLine();
            file.WriteLine("</CENTER>");
            file.WriteLine("</div>");
            file.WriteLine("</div>");
            //Footer
            file.WriteLine("<? include (\"footer.php\"); ?>");
            //Cerrar archivo
            file.Close();
            filemayor.WriteLine();
            filemayor.WriteLine("</CENTER>");
            filemayor.WriteLine("</div>");
            filemayor.WriteLine("</div>");
            //Footer
            filemayor.WriteLine("<? include (\"footer.php\"); ?>");
            //Cerrar archivo
            filemayor.Close();
        }
    }
}
