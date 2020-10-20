using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Data;
using DataTable = System.Data.DataTable;

namespace Nöbetci_v2
{
    class Program
    {
        static Application m_XlApp;
        static Workbook m_XlWorkbook;
        static Worksheet m_XlWorksheet;
        static Range m_XlRange;

        static List<Analist> m_AnalistciList = new List<Analist>();
        static List<Nobet> m_NobetList = new List<Nobet>();
        static List<Izinli> m_Izınli = new List<Izinli>();
        static List<DateTime> m_Tatiller;


        static void Main(string[] args)
        {
            GetExcel(@"C:\Users\senev\Documents\VS\Nöbet_Analist.xlsx");

            Sheets excelSheets = m_XlWorkbook.Worksheets;

            DataTable dtAnalist = ExceltoDataTable("Analist", 1, 1);
            DataTable dtIzinli = ExceltoDataTable("İzinler", 1, 1);
            DataTable dtNobetci = ExceltoDataTable("Ocak", 1, 1);

            TatilGunleri();
            GetAnalist(dtAnalist);
            GetNobet(dtNobetci);
            GetIzin(dtIzinli);

            GonderSms();

            Console.ReadKey();
        }

        public static void GetExcel(string saveAsLocation)
        {
            try
            {
                m_XlApp = new Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                m_XlWorkbook = m_XlApp.Workbooks.Open(saveAsLocation);

                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(m_XlRange);
                Marshal.ReleaseComObject(m_XlWorksheet);

                m_XlWorkbook.Close();
                Marshal.ReleaseComObject(m_XlWorkbook);

                m_XlApp.Quit();
                Marshal.ReleaseComObject(m_XlApp);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);                
            }
        }
    
        public static DataTable ExceltoDataTable(string sheetName, int headerLine, int columnStart)
        {
            DataTable dtExcel = new DataTable();

            Worksheet son = (Worksheet)m_XlWorkbook.Worksheets.Item[m_XlWorkbook.Worksheets.Count];

            int columnCountSon = son.UsedRange.Columns.Count;
            int rowCountSon = son.UsedRange.Rows.Count;

            m_XlWorksheet = (Worksheet)m_XlWorkbook.Worksheets.Item[sheetName];
            m_XlRange = m_XlWorksheet.UsedRange;

            int columnCount = m_XlRange.Columns.Count;
            int rowCount = m_XlRange.Rows.Count;

            for (int j = columnStart; j <= columnCount; j++)
            {
                dtExcel.Columns.Add(Convert.ToString(m_XlRange.Cells[headerLine, j].value), typeof(string));
            }

            for (int i = headerLine + 1; i <= rowCount; i++)
            {
                DataRow dr = dtExcel.NewRow();

                for (int j = columnStart; j <= columnCount; j++)
                {
                    dr[j - columnStart] = Convert.ToString(m_XlRange.Cells[i, j].value);
                }

                dtExcel.Rows.InsertAt(dr, dtExcel.Rows.Count + 1);
            }

            return dtExcel;
        } 

        public static void GetAnalist(DataTable dataTable)
        {
            m_AnalistciList = new List<Analist>();

            foreach (DataRow dRow in dataTable.Rows)
            {
                m_AnalistciList.Add(new Analist()
                {
                    ID = Convert.ToInt32(dRow["ID"]),
                    Ad = dRow["ÇALIŞAN"].ToString(),
                    Telefon = dRow["TELEFON"].ToString(),
                    Mail = dRow["MAIL"].ToString()
                });
            }
        }

        public static void GetNobet(DataTable dataTable)
        {
            m_NobetList = new List<Nobet>();

            foreach (DataRow dRow in dataTable.Rows)
            {
                m_NobetList.Add(new Nobet()
                {
                    ID = Convert.ToInt32(dRow["ID"]),
                    Ad = dRow["ÇALIŞAN"].ToString(),
                    Tarih = Convert.ToDateTime(dRow["TARİH"])
                });
            }
        }

        public static void GetIzin(DataTable dataTable)
        {
            m_Izınli = new List<Izinli>();

            //foreach (DataRow dRow in dataTable.Rows)
            //{
            //    m_Izınli.Add(new Izinli()
            //    {
            //        ID = Convert.ToInt32(dRow["ID"]),
            //        Ad = dRow["ÇALIŞAN"].ToString(),
            //        StartDate = Convert.ToDateTime(dRow["BAŞLANGIÇ TARİHİ"]),
            //        EndDate = Convert.ToDateTime(dRow["BİTİŞ TARİHİ"])
            //    });
            //}
        }

        public static List<DateTime> TatilGunleri()
        {
            m_Tatiller = new List<DateTime>
            {
                new DateTime(2020, 10, 28),
                new DateTime(2020, 10, 29),
                new DateTime(2020, 12, 31)
            };

            return m_Tatiller;
        }
        
        private static void GonderSms()
        {

        }
    }
}

internal class Analist
{
    public int ID { get; set; }
    public string Ad { get; set; }
    public string Telefon { get; set; }
    public string Mail { get; set; }
}
internal class Nobet
{
    public int ID { get; set; }
    public string Ad { get; set; }
    public DateTime Tarih { get; set; }
    public string SmsGonderim { get; set; }
}
internal class Izinli
{
    public int ID { get; set; }
    public string Ad { get; set; }
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
}
