using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using Spire.Xls;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Net;

namespace WebApplication2
{
    public partial class WebForm1 : System.Web.UI.Page
    {


        string exceldosyaadı = "";
        DataTable dt3 = new DataTable();
     


        protected void Page_Load(object sender, EventArgs e)
        {
           
        }
        //YÜKLE BUTONU


        protected void btnUpload_Click(object sender, EventArgs e)
        {
          
            if (fluExcel.HasFile)
            {
                lblMessage.Text = "";
               
                string fileExtension = System.IO.Path.GetExtension(fluExcel.FileName).ToLower();
                if (fileExtension == ".xls" || fileExtension == ".xlsx")
                {


               
                   
                    exceldosyaadı = fluExcel.PostedFile.FileName;
                    Session.Add("path2_", exceldosyaadı);

                    string date = DateTime.Now.ToString("dd-MM-yyyy HH-mm");
                   // string newFileName = string.Concat(date + lblMessage);
                    string newFileName = date + exceldosyaadı;
                    string filePath = System.IO.Path.GetFullPath(Server.MapPath("~/App_Data/"));                  
                    fluExcel.PostedFile.SaveAs(filePath + newFileName);              
                    LoadExcel(newFileName);

                }
                else
                {
                    lblMessage.Text = "Excel Dosyası Seçiniz!";
                }
            }
            if (!fluExcel.HasFile)
            {
                lblMessage.Text = "Dosya Seçin!";
            }

           
        }

        private void LoadExcel(string newFileName)
        {
            try
            {    //OledbConnection Bağlantısı
                OleDbConnection oleDbConn = new OleDbConnection();
                Session.Add("path_", System.IO.Path.GetFullPath(Server.MapPath("~/App_Data/")) + newFileName);
              
                string path = System.IO.Path.GetFullPath(Server.MapPath("~/App_Data/")) + newFileName;//okumak istenilen dosyanın yolu alındı.

                if (System.IO.Path.GetExtension(path) == ".xls")
                {
                    oleDbConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties =\"Excel 8.0;HDR=Yes;IMEX=1\";");
                }
                else if (System.IO.Path.GetExtension(path) == ".xlsx")
                {
                    oleDbConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties =\"Excel 12.0 Xml;HDR=YES\";");
                }
                oleDbConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                DataTable dt = new DataTable();


                DataTable dbShema = oleDbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dbShema == null)
                {
                    lblMessage.Text = "Excel Dosyasında Sayfa Bulunmamaktadır!";
                    return;
                }

                string firstPageName = dbShema.Rows[0]["TABLE_NAME"].ToString(); //excel sayfalarının isimlerini alıyor.

                cmd.Connection = oleDbConn;
                cmd.CommandText = "SELECT * FROM [" + firstPageName + "]";
                cmd.CommandType = CommandType.Text;
                OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
                adapter.Fill(dt);
                oleDbConn.Close();

                grdExcel.DataSource = dt;
                grdExcel.DataBind();
                DataTable dt2 = new DataTable();


       
                


                if (dropdownlist.Items[0].Selected == true)
                {
                    lblMessage.Text = "Banka Seçin!";
                }



                for (int dropdown = 0; dropdown < dropdownlist.Items.Count; dropdown++)
                {
                    string BKolonu = "";
                    string[] aaa = { };
                    string ccc = "";
                    string bbb = "";
                    string control = "0";
                    string control2 = "5";
                    int k;
                    string numara = "";
                    string adsoyad = "";
                    string telefon = "";
                    string kategori = "";
                    string tc_no = "";
                    string tutar = "";

                   
                   


                    if (dropdownlist.Items[0].Selected == true)
                    {
                        lblMessage.Text = "Banka Seçin!";
                    }


                  
                        if (dropdownlist.Items[1].Selected == true && exceldosyaadı == "Akbank.xlsx") //Akbank
                        {


                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                                                     
                        if (i == 1)
                            {
                                //  gridColon2.Columns[0].HeaderText = "A";
                                dt2.Columns.Add("No");
                                dt2.Columns.Add("Ad Soyad");
                                dt2.Columns.Add("Kategoriler");
                                dt2.Columns.Add("Telefon");
                                dt2.Columns.Add("TC No");
                                dt2.Columns.Add("Bağış Miktarı");


                                for (int j = 0; j < dt.Rows.Count; j++)
                                {

                                    numara = (j + 2).ToString();
                                  
                                    BKolonu = " " + dt.Rows[j][i].ToString();
                                    //Ad Soyad
                                    if (BKolonu.IndexOf(".") != -1)
                                    {
                                        if (BKolonu.IndexOf("-") != -1)
                                        {
                                            aaa = BKolonu.Split('.');
                                            bbb = aaa[1].ToString();
                                            aaa = bbb.Split('-');



                                            adsoyad = aaa[0];
                                        }
                                        else
                                        {
                                            adsoyad = "NULL";

                                        }
                                    }
                                    else
                                    {
                                        adsoyad = "NULL";

                                    }

                                    //Kategoriler 
                                    if (BKolonu.IndexOf("/") != -1)
                                    {
                                        if (BKolonu.IndexOf(" ") != -1)
                                        {
                                            aaa = BKolonu.Split('/');
                                            bbb = aaa[1].ToString();
                                            aaa = bbb.Split(' ');
                                            kategori = aaa[0];

                                        }

                                    }
                                    else
                                    {
                                        kategori = "NULL";

                                    }

                                    //telefon numarası 

                                    aaa = BKolonu.Split(' ');
                                    for (k = 0; k < aaa.Length; k++)
                                    {
                                        telefon = "";
                                        ccc = aaa[k];
                                        if (ccc.IndexOf("5") != -1 && csSayisalDegerMi(ccc) == true)
                                        {
                                            if (ccc.StartsWith(control) && ccc.Length == 11 || ccc.StartsWith(control2) && ccc.Length == 10) //telefon numarası
                                            {
                                                bbb = aaa[k].ToString();
                                                aaa = bbb.Split(' ');


                                                telefon = aaa[0];
                                            }
                                            else
                                            {
                                                telefon = "NULL";

                                            }
                                        }
                                        if (ccc.IndexOf("-5") != -1)
                                        {
                                            bbb = aaa[k].ToString();
                                            aaa = bbb.Split('-', ' ');
                                            telefon = aaa[1];

                                        }
                                    }

                                    if (BKolonu.IndexOf("5") == -1)//telefon numarası değilse null
                                    {
                                        telefon = "NULL";

                                    }
                                    else if (csSayisalDegerMi(ccc) == false && ((ccc.Length != 11 && ccc.StartsWith(control)) || (ccc.Length != 10 && ccc.StartsWith(control2))))
                                    {
                                        telefon = "NULL";

                                    }

                                    //Tc Kimlik Numarası

                                    string[] tc = BKolonu.Split(' ');
                                    for (k = 0; k < tc.Length; k++)
                                    {
                                        ccc = tc[k];


                                        if (csSayisalDegerMi(ccc) == true)
                                        {
                                            if (ccc.Length == 11 && !ccc.StartsWith(control))
                                            {
                                                bbb = tc[k].ToString();
                                                tc = bbb.Split(' ');

                                                tc_no = tc[0];

                                            }
                                            else
                                            {
                                                tc_no = "NULL";

                                            }
                                        }
                                    }

                                    if (csSayisalDegerMi(ccc) == false && (ccc.Length != 11 || ccc.StartsWith(control)))
                                    {
                                        tc_no = "NULL";

                                    }


                                    dt2.Rows.Add(numara, adsoyad, kategori, telefon,tc_no);
                                   




                                }
                            }

                            if (i == 2)
                            {
                                for (int j = 0; j < dt.Rows.Count; j++)
                                {
                                    tutar = dt.Rows[j][i].ToString();
                                    dt2.Rows[j][5] = tutar;

                                }
                            }

                        }

                        gridColon2.DataSource = dt2;
                        gridColon2.DataBind();



                            //DATATABLE'I EXCELE AKTARMA//
                            try
                            {
                                if (dt2 == null || dt2.Columns.Count == 0)
                                    throw new Exception("Datatable boş.");


                                Excel.Application excelApp = new Excel.Application();
                                excelApp.Workbooks.Add();
                                Excel.Workbook xlKitap = excelApp.Workbooks.Open(Session["path_"].ToString(), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                                Excel.Sheets xlSayfa = xlKitap.Worksheets;
                                var xlYeniSayfa = (Excel.Worksheet)xlSayfa.Add(xlSayfa[1], Type.Missing, Type.Missing, Type.Missing);
                                xlYeniSayfa.Name = "Sayfa2";

                                xlYeniSayfa = (Excel.Worksheet)xlKitap.Worksheets.get_Item(1);
                                xlYeniSayfa.Select();



                                // column başlıkları
                                for (int i = 0; i < dt2.Columns.Count; i++)
                                {
                                    xlYeniSayfa.Cells[1, (i + 1)] = dt2.Columns[i].ColumnName;
                                }

                                // rows 
                                for (int i = 0; i < dt2.Rows.Count; i++)
                                {
                                    for (int j = 0; j < dt2.Columns.Count; j++)
                                    {
                                        xlYeniSayfa.Cells[(i + 2), (j + 1)] = dt2.Rows[i][j];
                                    }
                                }

                                //kaydedilecek excelin yolu
                                if (Session["path_"].ToString() != null && Session["path_"].ToString() != "")
                                {
                                    try
                                    {//excele kaydetme                               

                                        xlYeniSayfa.SaveAs(Session["path_"].ToString());
                                        excelApp.Quit();

                           

                                }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Excel dosyası kaydedilemedi.Dosya yolunu Kontrol et.\n"
                                        + ex.Message);
                                    }
                                }
                                else //dosya yolu yok
                                {
                                    excelApp.Visible = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Hata: \n" + ex.Message);
                            }



                        

                    }




                        if (dropdownlist.Items[2].Selected == true && exceldosyaadı == "Albaraka.xlsx") //Albaraka
                        {

                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if (i == 1)
                                {
                                    dt2.Columns.Add("No");
                                    dt2.Columns.Add("Ad Soyad");
                                    dt2.Columns.Add("Kategoriler");
                                    dt2.Columns.Add("Telefon");
                                    dt2.Columns.Add("TC No");
                                    dt2.Columns.Add("Bağış Miktarı");


                                    for (int j = 0; j < dt.Rows.Count; j++)
                                    {
                                        numara = (j + 2).ToString();
                                        BKolonu = " " + dt.Rows[j][i].ToString();
                                        //Ad Soyad
                                        if (BKolonu.IndexOf(" ") != -1)
                                        {
                                            if (BKolonu.IndexOf(" ") != -1)
                                            {
                                                aaa = BKolonu.Split(' ');
                                                //bbb = aaa[1].ToString();
                                                //aaa = bbb.Split(' ');



                                                adsoyad = aaa[2] + " " + aaa[3];

                                                if (BKolonu.IndexOf("SN") == -1)
                                                {
                                                    adsoyad = "NULL";
                                                }

                                            }
                                            else
                                            {
                                                adsoyad = "NULL";

                                            }
                                        }
                                        else
                                        {
                                            adsoyad = "NULL";

                                        }

                                        //Kategoriler 
                                        if (BKolonu.IndexOf(" ") != -1)
                                        {
                                            if (BKolonu.IndexOf(" ") != -1)
                                            {
                                                aaa = BKolonu.Split(' ');
                                                kategori = aaa[4];

                                            }
                                            if (BKolonu.IndexOf("SN") == -1)
                                            {
                                                kategori = "NULL";
                                            }
                                        }
                                        else
                                        {
                                            kategori = "NULL";

                                        }

                                        //telefon numarası 
                                        aaa = BKolonu.Split(' ');
                                        for (k = 0; k < aaa.Length; k++)
                                        {
                                            telefon = "";
                                            ccc = aaa[k];
                                            if (ccc.IndexOf("5") != -1 && csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.StartsWith(control) && ccc.Length == 11 || ccc.StartsWith(control2) && ccc.Length == 10) //telefon numarası
                                                {
                                                    bbb = aaa[k].ToString();
                                                    aaa = bbb.Split(' ');


                                                    telefon = aaa[0];
                                                }

                                                else
                                                {
                                                    telefon = "NULL";

                                                }

                                            }
                                            if (ccc.IndexOf("-5") != -1)
                                            {
                                                bbb = aaa[k].ToString();
                                                aaa = bbb.Split('-', ' ');
                                                telefon = aaa[1];

                                            }
                                        }

                                        if (BKolonu.IndexOf("5") == -1)//telefon numarası değilse null
                                        {
                                            telefon = "NULL";

                                        }
                                        else if (csSayisalDegerMi(ccc) == false && ((ccc.Length != 11 && ccc.StartsWith(control)) || (ccc.Length != 10 && ccc.StartsWith(control2))))
                                        {
                                            telefon = "NULL";

                                        }

                                        //Tc Kimlik Numarası

                                        string[] tc = BKolonu.Split(' ');
                                        for (k = 0; k < tc.Length; k++)
                                        {
                                            ccc = tc[k];


                                            if (csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.Length == 11 && !ccc.StartsWith(control))
                                                {
                                                    bbb = tc[k].ToString();
                                                    tc = bbb.Split(' ');

                                                    tc_no = tc[0];

                                                }
                                                else
                                                {
                                                    tc_no = "NULL";

                                                }
                                            }


                                        }
                                        if (csSayisalDegerMi(ccc) == false && (ccc.Length != 11 || ccc.StartsWith(control)))
                                        {
                                            tc_no = "NULL";

                                        }

                                        dt2.Rows.Add(numara, adsoyad, kategori, telefon, tc_no);
                                    }
                                }
                            if (i == 2)
                            {
                                for (int j = 0; j < dt.Rows.Count; j++)
                                {
                                    tutar = dt.Rows[j][i].ToString();
                                    dt2.Rows[j][5] = tutar;

                                }
                            }
                        }
                            gridColon2.DataSource = dt2;
                            gridColon2.DataBind();
                            //DATATABLE'I EXCELE AKTARMA//
                            try
                            {
                                if (dt2 == null || dt2.Columns.Count == 0)
                                    throw new Exception("Datatable boş.");


                                Excel.Application excelApp = new Excel.Application();
                                excelApp.Workbooks.Add();
                                Excel.Workbook xlKitap = excelApp.Workbooks.Open(Session["path_"].ToString(), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                                Excel.Sheets xlSayfa = xlKitap.Worksheets;
                                var xlYeniSayfa = (Excel.Worksheet)xlSayfa.Add(xlSayfa[1], Type.Missing, Type.Missing, Type.Missing);
                                xlYeniSayfa.Name = "Sayfa2";

                                xlYeniSayfa = (Excel.Worksheet)xlKitap.Worksheets.get_Item(1);
                                xlYeniSayfa.Select();



                                // column başlıkları
                                for (int i = 0; i < dt2.Columns.Count; i++)
                                {
                                    xlYeniSayfa.Cells[1, (i + 1)] = dt2.Columns[i].ColumnName;
                                }

                                // rows 
                                for (int i = 0; i < dt2.Rows.Count; i++)
                                {
                                    for (int j = 0; j < dt2.Columns.Count; j++)
                                    {
                                        xlYeniSayfa.Cells[(i + 2), (j + 1)] = dt2.Rows[i][j];
                                    }
                                }

                                //kaydedilecek excelin yolu
                                if (Session["path_"].ToString() != null && Session["path_"].ToString() != "")
                                {
                                    try
                                    {//excele kaydetme                               

                                        xlYeniSayfa.SaveAs(Session["path_"].ToString());
                                        excelApp.Quit();
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Excel dosyası kaydedilemedi.Dosya yolunu Kontrol et.\n"
                                        + ex.Message);
                                    }
                                }
                                else //dosya yolu yok
                                {
                                    excelApp.Visible = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Hata: \n" + ex.Message);
                            }

                       
                    }





                        if (dropdownlist.Items[3].Selected == true && exceldosyaadı == "Halkbank.xlsx") //Halkbank
                        {

                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if (i == 1)
                                {
                                    dt2.Columns.Add("No");
                                    dt2.Columns.Add("Ad Soyad");
                                    dt2.Columns.Add("Kategoriler");
                                    dt2.Columns.Add("Telefon");
                                    dt2.Columns.Add("TC No");
                                    dt2.Columns.Add("Bağış Miktarı");


                                    for (int j = 0; j < dt.Rows.Count; j++)
                                    {
                                        numara = (j + 2).ToString();
                                        BKolonu = " " + dt.Rows[j][i].ToString();
                                        //Ad Soyad
                                        if (BKolonu.IndexOf(" ") != -1)
                                        {
                                            if (BKolonu.IndexOf(" ") != -1)
                                            {
                                                aaa = BKolonu.Split(' ');
                                                adsoyad = aaa[1] + " " + aaa[2];
                                            }

                                            else
                                            {
                                                adsoyad = "NULL";

                                            }
                                        }

                                        else
                                        {
                                            adsoyad = "NULL";

                                        }

                                        //Kategoriler 
                                        if (BKolonu.IndexOf(" ") != -1)
                                        {
                                            if (BKolonu.IndexOf(" ") != -1)
                                            {
                                                aaa = BKolonu.Split(' ');
                                                kategori = aaa[3];
                                            }
                                        }
                                        else
                                        {
                                            kategori = "NULL";

                                        }

                                        //telefon numarası 
                                        aaa = BKolonu.Split(' ');
                                        for (k = 0; k < aaa.Length; k++)
                                        {
                                            telefon = "";
                                            ccc = aaa[k];
                                            if (ccc.IndexOf("5") != -1 && csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.StartsWith(control) && ccc.Length == 11 || ccc.StartsWith(control2) && ccc.Length == 10) //telefon numarası
                                                {
                                                    bbb = aaa[k].ToString();
                                                    aaa = bbb.Split(' ');


                                                    telefon = aaa[0];
                                                }

                                                else
                                                {
                                                    telefon = "NULL";

                                                }

                                            }
                                            if (ccc.IndexOf("-5") != -1)
                                            {
                                                bbb = aaa[k].ToString();
                                                aaa = bbb.Split('-', ' ');
                                                telefon = aaa[1];

                                            }
                                        }

                                        if (BKolonu.IndexOf("5") == -1)//telefon numarası değilse null
                                        {
                                            telefon = "NULL";

                                        }
                                        else if (csSayisalDegerMi(ccc) == false && ((ccc.Length != 11 && ccc.StartsWith(control)) || (ccc.Length != 10 && ccc.StartsWith(control2))))
                                        {
                                            telefon = "NULL";

                                        }

                                        //Tc Kimlik Numarası

                                        string[] tc = BKolonu.Split(' ');
                                        for (k = 0; k < tc.Length; k++)
                                        {
                                            ccc = tc[k];


                                            if (csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.Length == 11 && !ccc.StartsWith(control))
                                                {
                                                    bbb = tc[k].ToString();
                                                    tc = bbb.Split(' ');

                                                    tc_no = tc[0];

                                                }
                                                else
                                                {
                                                    tc_no = "NULL";

                                                }
                                            }


                                        }
                                        if (csSayisalDegerMi(ccc) == false && (ccc.Length != 11 || ccc.StartsWith(control)))
                                        {
                                            tc_no = "NULL";

                                        }

                                        dt2.Rows.Add(numara, adsoyad, kategori, telefon, tc_no);
                                    }
                                }
                            if (i == 2)
                            {
                                for (int j = 0; j < dt.Rows.Count; j++)
                                {
                                    tutar = dt.Rows[j][i].ToString();
                                    dt2.Rows[j][5] = tutar;

                                }
                            }
                        }
                            gridColon2.DataSource = dt2;
                            gridColon2.DataBind();



                            //DATATABLE'I EXCELE AKTARMA//
                            try
                            {
                                if (dt2 == null || dt2.Columns.Count == 0)
                                    throw new Exception("Datatable boş.");


                                Excel.Application excelApp = new Excel.Application();
                                excelApp.Workbooks.Add();
                                Excel.Workbook xlKitap = excelApp.Workbooks.Open(Session["path_"].ToString(), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                                Excel.Sheets xlSayfa = xlKitap.Worksheets;
                                var xlYeniSayfa = (Excel.Worksheet)xlSayfa.Add(xlSayfa[1], Type.Missing, Type.Missing, Type.Missing);
                                xlYeniSayfa.Name = "Sayfa2";

                                xlYeniSayfa = (Excel.Worksheet)xlKitap.Worksheets.get_Item(1);
                                xlYeniSayfa.Select();

                                // column başlıkları
                                for (int i = 0; i < dt2.Columns.Count; i++)
                                {
                                    xlYeniSayfa.Cells[1, (i + 1)] = dt2.Columns[i].ColumnName;
                                }

                                // rows 
                                for (int i = 0; i < dt2.Rows.Count; i++)
                                {
                                    for (int j = 0; j < dt2.Columns.Count; j++)
                                    {
                                        xlYeniSayfa.Cells[(i + 2), (j + 1)] = dt2.Rows[i][j];
                                    }
                                }

                                //kaydedilecek excelin yolu
                                if (Session["path_"].ToString() != null && Session["path_"].ToString() != "")
                                {
                                    try
                                    {//excele kaydetme                               

                                        xlYeniSayfa.SaveAs(Session["path_"].ToString());
                                        excelApp.Quit();
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Excel dosyası kaydedilemedi.Dosya yolunu Kontrol et.\n"
                                        + ex.Message);
                                    }
                                }
                                else //dosya yolu yok
                                {
                                    excelApp.Visible = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Hata: \n" + ex.Message);
                            }

                   
                    }

                        if (dropdownlist.Items[4].Selected == true && exceldosyaadı == "Kuveyt.xlsx") //Kuveyt
                        {

                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if (i == 1)
                                {
                                    dt2.Columns.Add("No");
                                    dt2.Columns.Add("Ad Soyad");
                                    dt2.Columns.Add("Kategoriler");
                                    dt2.Columns.Add("Telefon");
                                    dt2.Columns.Add("TC No");
                                    dt2.Columns.Add("Bağış Miktarı");

                                    for (int j = 0; j < dt.Rows.Count; j++)
                                    {
                                        numara = (j + 2).ToString();
                                        BKolonu = " " + dt.Rows[j][i].ToString();
                                        //Ad Soyad
                                        if (BKolonu.IndexOf("(") != -1)
                                        {
                                            if (BKolonu.IndexOf(")") != -1)
                                            {
                                                aaa = BKolonu.Split('(');
                                                bbb = aaa[1].ToString();
                                                aaa = bbb.Split(')');
                                                adsoyad = aaa[0];
                                            }
                                            else
                                            {
                                                adsoyad = "NULL";

                                            }
                                        }
                                        else
                                        {
                                            adsoyad = "NULL";

                                        }

                                        //Kategoriler 
                                        if (BKolonu.IndexOf("Aciklama=") != -1)
                                        {
                                            // aaa = BKolonu.Split(' ');
                                            //bbb = aaa[1].ToString();
                                            //aaa = bbb.Split(' ');
                                            //kategori = aaa[0];

                                          //  aaa = BKolonu.Split(' ');
                                           // kategori = aaa[6];
                                        }

                                        if (BKolonu.IndexOf("-Gönderen:") != -1)
                                        {
                                            aaa = BKolonu.Split(' ');
                                            kategori = aaa[1];
                                        }
                                        else
                                        {
                                            kategori = "NULL";

                                        }

                                        //telefon numarası 
                                        aaa = BKolonu.Split(' ');
                                        for (k = 0; k < aaa.Length; k++)
                                        {
                                            telefon = "";
                                            ccc = aaa[k];
                                            if (ccc.IndexOf("5") != -1 && csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.StartsWith(control) && ccc.Length == 11 || ccc.StartsWith(control2) && ccc.Length == 10) //telefon numarası
                                                {
                                                    bbb = aaa[k].ToString();
                                                    aaa = bbb.Split(' ');


                                                    telefon = aaa[0];
                                                }

                                                else
                                                {
                                                    telefon = "NULL";

                                                }

                                            }
                                            if (ccc.IndexOf("-5") != -1)
                                            {
                                                bbb = aaa[k].ToString();
                                                aaa = bbb.Split('-', ' ');
                                                telefon = aaa[1];

                                            }
                                        }

                                        if (BKolonu.IndexOf("5") == -1)//telefon numarası değilse null
                                        {
                                            telefon = "NULL";
                                           
                                        }
                                        else if (csSayisalDegerMi(ccc) == false && ((ccc.Length != 11 && ccc.StartsWith(control)) || (ccc.Length != 10 && ccc.StartsWith(control2))))
                                        {
                                            telefon = "NULL";

                                        }

                                        //Tc Kimlik Numarası

                                        string[] tc = BKolonu.Split(' ');
                                        for (k = 0; k < tc.Length; k++)
                                        {
                                            ccc = tc[k];


                                            if (csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.Length == 11 && !ccc.StartsWith(control))
                                                {
                                                    bbb = tc[k].ToString();
                                                    tc = bbb.Split(' ');

                                                    tc_no = tc[0];

                                                }
                                                else
                                                {
                                                    tc_no = "NULL";

                                                }
                                            }
                                        }
                                        if (csSayisalDegerMi(ccc) == false && (ccc.Length != 11 || ccc.StartsWith(control)))
                                        {
                                            tc_no = "NULL";

                                        }

                                        dt2.Rows.Add(numara, adsoyad, kategori, telefon, tc_no);
                                    }
                                }
                            if (i == 2)
                            {
                                for (int j = 0; j < dt.Rows.Count; j++)
                                {
                                    tutar = dt.Rows[j][i].ToString();
                                    dt2.Rows[j][5] = tutar;

                                }
                            }
                        }
                            gridColon2.DataSource = dt2;
                            gridColon2.DataBind();
                            //DATATABLE'I EXCELE AKTARMA//
                            try
                            {
                                if (dt2 == null || dt2.Columns.Count == 0)
                                    throw new Exception("Datatable boş.");


                                Excel.Application excelApp = new Excel.Application();
                                excelApp.Workbooks.Add();
                                Excel.Workbook xlKitap = excelApp.Workbooks.Open(Session["path_"].ToString(), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                                Excel.Sheets xlSayfa = xlKitap.Worksheets;
                                var xlYeniSayfa = (Excel.Worksheet)xlSayfa.Add(xlSayfa[1], Type.Missing, Type.Missing, Type.Missing);
                                xlYeniSayfa.Name = "Sayfa2";

                                xlYeniSayfa = (Excel.Worksheet)xlKitap.Worksheets.get_Item(1);
                                xlYeniSayfa.Select();

                                // column başlıkları
                                for (int i = 0; i < dt2.Columns.Count; i++)
                                {
                                    xlYeniSayfa.Cells[1, (i + 1)] = dt2.Columns[i].ColumnName;
                                }

                                // rows 
                                for (int i = 0; i < dt2.Rows.Count; i++)
                                {
                                    for (int j = 0; j < dt2.Columns.Count; j++)
                                    {
                                        xlYeniSayfa.Cells[(i + 2), (j + 1)] = dt2.Rows[i][j];
                                    }
                                }

                                //kaydedilecek excelin yolu
                                if (Session["path_"].ToString() != null && Session["path_"].ToString() != "")
                                {
                                    try
                                    {//excele kaydetme                               

                                        xlYeniSayfa.SaveAs(Session["path_"].ToString());
                                        excelApp.Quit();
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Excel dosyası kaydedilemedi.Dosya yolunu Kontrol et.\n"
                                        + ex.Message);
                                    }
                                }
                                else //dosya yolu yok
                                {
                                    excelApp.Visible = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Hata: \n" + ex.Message);
                            }

                    
                    }

                        if (dropdownlist.Items[5].Selected == true && exceldosyaadı == "Ptt.xlsx") //Ptt
                        {

                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            if (i == 1)
                            {
                                dt2.Columns.Add("No");
                                dt2.Columns.Add("Ad Soyad");
                                dt2.Columns.Add("Kategoriler");
                                dt2.Columns.Add("Telefon");
                                dt2.Columns.Add("TC No");
                                dt2.Columns.Add("Bağış Miktarı");

                                for (int j = 0; j < dt.Rows.Count; j++)
                                {
                                   
                                    numara = (j + 2).ToString();
                                    BKolonu = " " + dt.Rows[j][i].ToString();
                                    adsoyad = "";
                                    //Ad Soyad
                                    if (BKolonu.IndexOf(" ") != -1)
                                    {

                                        //  aaa = BKolonu.Split(' ');
                                        ////bbb = aaa[1].ToString();
                                        ////aaa = bbb.Split(' ');                                       
                                        // adsoyad = aaa[0];

                                        if (BKolonu.IndexOf("GON.") != -1)
                                        {
                                            aaa = BKolonu.Split(' ');
                                            adsoyad = aaa[3] + " " + aaa[4];
                                        }

                                        aaa = BKolonu.Split(' ');
                                        if (BKolonu.IndexOf("GON.") == -1 && aaa.Length == 4)
                                        {
                                            adsoyad = aaa[1] + " " + aaa[2];
                                        }
                                    }
                                    else
                                    {
                                        adsoyad = "NULL";

                                    }

                                    //Kategoriler 
                                    if (BKolonu.IndexOf(" ") != -1)
                                    {
                                        kategori = "";
                                        aaa = BKolonu.Split(' ');
                                        ////bbb = aaa[1].ToString();
                                        ////aaa = bbb.Split(' ');
                                        //kategori = aaa[0];
                                        if (BKolonu.IndexOf("GON.") != -1 && aaa.Length >= 6)
                                        {

                                            kategori = aaa[6];
                                        }
                                        aaa = BKolonu.Split(' ');
                                        if (BKolonu.IndexOf("GON.") == -1 && aaa.Length == 4)
                                        {

                                            kategori = aaa[3];
                                        }
                                    }
                                    else
                                    {
                                        kategori = "NULL";

                                    }
                                    //telefon numarası 
                                    aaa = BKolonu.Split(' ');
                                    for (k = 0; k < aaa.Length; k++)
                                    {
                                        telefon = "";
                                        ccc = aaa[k];
                                        if (ccc.IndexOf("5") != -1 && csSayisalDegerMi(ccc) == true)
                                        {
                                            if (ccc.StartsWith(control) && ccc.Length == 11 || ccc.StartsWith(control2) && ccc.Length == 10) //telefon numarası
                                            {
                                                bbb = aaa[k].ToString();
                                                aaa = bbb.Split(' ');


                                                telefon = aaa[0];
                                            }

                                            else
                                            {
                                                telefon = "NULL";

                                            }

                                        }
                                        if (ccc.IndexOf("-5") != -1)
                                        {
                                            bbb = aaa[k].ToString();
                                            aaa = bbb.Split('-', ' ');
                                            telefon = aaa[1];

                                        }
                                    }

                                    if (BKolonu.IndexOf("5") == -1)//telefon numarası değilse null
                                    {
                                        telefon = "NULL";

                                    }
                                    else if (csSayisalDegerMi(ccc) == false && ((ccc.Length != 11 && ccc.StartsWith(control)) || (ccc.Length != 10 && ccc.StartsWith(control2))))
                                    {
                                        telefon = "NULL";
                                    }

                                    //Tc Kimlik Numarası

                                    string[] tc = BKolonu.Split(' ');
                                    for (k = 0; k < tc.Length; k++)
                                    {
                                        ccc = tc[k];


                                        if (csSayisalDegerMi(ccc) == true)
                                        {
                                            if (ccc.Length == 11 && !ccc.StartsWith(control))
                                            {
                                                bbb = tc[k].ToString();
                                                tc = bbb.Split(' ');

                                                tc_no = tc[0];

                                            }
                                            else
                                            {
                                                tc_no = "NULL";

                                            }
                                        }


                                    }
                                    if (csSayisalDegerMi(ccc) == false && (ccc.Length != 11 || ccc.StartsWith(control)))
                                    {
                                        tc_no = "NULL";

                                    }

                                    dt2.Rows.Add(numara, adsoyad, kategori, telefon, tc_no);
                                }
                            }

                            if (i == 2)
                            {
                                for (int j = 0; j < dt.Rows.Count; j++)
                                {
                                    tutar = dt.Rows[j][i].ToString();
                                    dt2.Rows[j][5] = tutar;
                                   
                                }
                            }
                          


                        } //if 


                           
                        
                            gridColon2.DataSource = dt2;
                            gridColon2.DataBind();

                            //DATATABLE'I EXCELE AKTARMA//
                            try
                            {
                                if (dt2 == null || dt2.Columns.Count == 0)
                                    throw new Exception("Datatable boş.");


                                Excel.Application excelApp = new Excel.Application();
                                excelApp.Workbooks.Add();
                                Excel.Workbook xlKitap = excelApp.Workbooks.Open(Session["path_"].ToString(), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                                Excel.Sheets xlSayfa = xlKitap.Worksheets;
                                var xlYeniSayfa = (Excel.Worksheet)xlSayfa.Add(xlSayfa[1], Type.Missing, Type.Missing, Type.Missing);
                                xlYeniSayfa.Name = "Sayfa2";

                                xlYeniSayfa = (Excel.Worksheet)xlKitap.Worksheets.get_Item(1);
                                xlYeniSayfa.Select();

                                // column başlıkları
                                for (int i = 0; i < dt2.Columns.Count; i++)
                                {
                                    xlYeniSayfa.Cells[1, (i + 1)] = dt2.Columns[i].ColumnName;
                                }

                                // rows 
                                for (int i = 0; i < dt2.Rows.Count; i++)
                                {
                                    for (int j = 0; j < dt2.Columns.Count; j++)
                                    {
                                        xlYeniSayfa.Cells[(i + 2), (j + 1)] = dt2.Rows[i][j];
                                    }
                                }

                                //kaydedilecek excelin yolu
                                if (Session["path_"].ToString() != null && Session["path_"].ToString() != "")
                                {
                                    try
                                    {//excele kaydetme                               

                                        xlYeniSayfa.SaveAs(Session["path_"].ToString());
                                        excelApp.Quit();


                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Excel dosyası kaydedilemedi.Dosya yolunu Kontrol et.\n"
                                        + ex.Message);
                                    }
                                }
                                else //dosya yolu yok
                                {
                                    excelApp.Visible = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Hata: \n" + ex.Message);
                            }

                   
                    }
                        if (dropdownlist.Items[6].Selected == true && exceldosyaadı == "Vakıf Katılım.xlsx")  //Vakıf Katılım
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if (i == 1)
                                {
                                    dt2.Columns.Add("No");
                                    dt2.Columns.Add("Ad Soyad");
                                    dt2.Columns.Add("Kategoriler");
                                    dt2.Columns.Add("Telefon");
                                    dt2.Columns.Add("TC No");
                                dt2.Columns.Add("Bağış Miktarı");

                                    for (int j = 0; j < dt.Rows.Count; j++)
                                    {
                                        numara = (j + 2).ToString();
                                        BKolonu = " " + dt.Rows[j][i].ToString();
                                        //Ad Soyad
                                        if (BKolonu.IndexOf("=") != -1)
                                        {
                                            if (BKolonu.IndexOf("G") != -1)
                                            {
                                                aaa = BKolonu.Split('=');
                                                bbb = aaa[1].ToString();
                                                aaa = bbb.Split('G');
                                                adsoyad = aaa[0];
                                            }

                                            else
                                            {
                                                adsoyad = "NULL";

                                            }
                                        }
                                        else if (BKolonu.IndexOf("Amir=") == -1)
                                        {
                                            aaa = BKolonu.Split(' ');
                                            adsoyad = aaa[1] + " " + aaa[2] + " " + aaa[3];
                                        }
                                        else
                                        {
                                            adsoyad = "NULL";

                                        }

                                        //Kategoriler 
                                        if (BKolonu.IndexOf(" ") != -1)
                                        {
                                            //if (BKolonu.IndexOf(" ") != -1)
                                            //{
                                            //    aaa = BKolonu.Split(' ');
                                            //    //bbb = aaa[1].ToString();
                                            //    //aaa = bbb.Split(' ');
                                            //    kategori = aaa[0];
                                            //}
                                            string[] ddd = BKolonu.Split(' ');

                                            if (aaa.Length >= 6 && csSayisalDegerMi(ddd[4]) == false && csSayisalDegerMi(ddd[5]) == false && csSayisalDegerMi(ddd[6]) == false)
                                            {
                                                aaa = BKolonu.Split(' ');
                                                kategori = aaa[4] + " " + aaa[5] + " " + aaa[6];
                                            }

                                            else
                                            {
                                                kategori = "NULL";
                                            }
                                        }
                                        else
                                        {
                                            kategori = "NULL";

                                        }

                                        //telefon numarası 
                                        aaa = BKolonu.Split(' ');
                                        for (k = 0; k < aaa.Length; k++)
                                        {
                                            telefon = "";
                                            ccc = aaa[k];
                                            if (ccc.IndexOf("5") != -1 && csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.StartsWith(control) && ccc.Length == 11 || ccc.StartsWith(control2) && ccc.Length == 10) //telefon numarası
                                                {
                                                    bbb = aaa[k].ToString();
                                                    aaa = bbb.Split(' ');


                                                    telefon = aaa[0];
                                                }

                                                else
                                                {
                                                    telefon = "NULL";

                                                }

                                            }
                                            if (ccc.IndexOf("-5") != -1)
                                            {
                                                bbb = aaa[k].ToString();
                                                aaa = bbb.Split('-', ' ');
                                                telefon = aaa[1];

                                            }
                                        }

                                        if (BKolonu.IndexOf("5") == -1)//telefon numarası değilse null
                                        {
                                            telefon = "NULL";

                                        }
                                        else if (csSayisalDegerMi(ccc) == false && ((ccc.Length != 11 && ccc.StartsWith(control)) || (ccc.Length != 10 && ccc.StartsWith(control2))))
                                        {
                                            telefon = "NULL";

                                        }

                                        //Tc Kimlik Numarası

                                        string[] tc = BKolonu.Split(' ');
                                        for (k = 0; k < tc.Length; k++)
                                        {
                                            ccc = tc[k];


                                            if (csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.Length == 11 && !ccc.StartsWith(control))
                                                {
                                                    bbb = tc[k].ToString();
                                                    tc = bbb.Split(' ');

                                                    tc_no = tc[0];

                                                }
                                                else
                                                {
                                                    tc_no = "NULL";

                                                }
                                            }
                                        }
                                        if (csSayisalDegerMi(ccc) == false && (ccc.Length != 11 || ccc.StartsWith(control)))
                                        {
                                            tc_no = "NULL";

                                        }

                                        dt2.Rows.Add(numara, adsoyad, kategori, telefon, tc_no);
                                    }
                                }
                            if (i == 2)
                            {
                                for (int j = 0; j < dt.Rows.Count; j++)
                                {
                                    tutar = dt.Rows[j][i].ToString();
                                    dt2.Rows[j][5] = tutar;

                                }
                            }
                        }
                            gridColon2.DataSource = dt2;
                            gridColon2.DataBind();
                            //DATATABLE'I EXCELE AKTARMA//
                            try
                            {
                                if (dt2 == null || dt2.Columns.Count == 0)
                                    throw new Exception("Datatable boş.");


                                Excel.Application excelApp = new Excel.Application();
                                excelApp.Workbooks.Add();
                                Excel.Workbook xlKitap = excelApp.Workbooks.Open(Session["path_"].ToString(), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                                Excel.Sheets xlSayfa = xlKitap.Worksheets;
                                var xlYeniSayfa = (Excel.Worksheet)xlSayfa.Add(xlSayfa[1], Type.Missing, Type.Missing, Type.Missing);
                                xlYeniSayfa.Name = "Sayfa2";

                                xlYeniSayfa = (Excel.Worksheet)xlKitap.Worksheets.get_Item(1);
                                xlYeniSayfa.Select();



                                // column başlıkları
                                for (int i = 0; i < dt2.Columns.Count; i++)
                                {
                                    xlYeniSayfa.Cells[1, (i + 1)] = dt2.Columns[i].ColumnName;
                                }

                                // rows 
                                for (int i = 0; i < dt2.Rows.Count; i++)
                                {
                                    for (int j = 0; j < dt2.Columns.Count; j++)
                                    {
                                        xlYeniSayfa.Cells[(i + 2), (j + 1)] = dt2.Rows[i][j];
                                    }
                                }

                                //kaydedilecek excelin yolu
                                if (Session["path_"].ToString() != null && Session["path_"].ToString() != "")
                                {
                                    try
                                    {//excele kaydetme                               

                                        xlYeniSayfa.SaveAs(Session["path_"].ToString());
                                        excelApp.Quit();


                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Excel dosyası kaydedilemedi.Dosya yolunu Kontrol et.\n"
                                        + ex.Message);
                                    }
                                }
                                else //dosya yolu yok
                                {
                                    excelApp.Visible = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Hata: \n" + ex.Message);
                            }

                    }
                        if (dropdownlist.Items[7].Selected == true && exceldosyaadı == "Vakıf Bank.xlsx") //Vakıf Bank
                        {

                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if (i == 1)
                                {
                                    dt2.Columns.Add("No");
                                    dt2.Columns.Add("Ad Soyad");
                                    dt2.Columns.Add("Kategoriler");
                                    dt2.Columns.Add("Telefon");
                                    dt2.Columns.Add("TC No");
                                   dt2.Columns.Add("Bağış Miktarı");


                                    for (int j = 0; j < dt.Rows.Count; j++)
                                    {
                                        numara = (j + 2).ToString();
                                        BKolonu = " " + dt.Rows[j][i].ToString();
                                        //Ad Soyad
                                        if (BKolonu.IndexOf("nolu") != -1)
                                        {

                                            if (BKolonu.IndexOf("hesabından") != -1)
                                            {



                                                aaa = BKolonu.Split(' ');
                                                //bbb = aaa[1].ToString();
                                                //aaa = bbb.Split(' ');

                                                adsoyad = aaa[14] + " " + aaa[15];


                                                if (aaa[14] == "nolu")
                                                {
                                                    aaa = BKolonu.Split(' ');
                                                    adsoyad = aaa[15] + " " + aaa[16];
                                                }
                                                if (aaa[12] == "nolu")
                                                {
                                                    aaa = BKolonu.Split(' ');
                                                    adsoyad = aaa[13] + " " + aaa[14];
                                                }
                                                if (aaa[15] == "nolu")
                                                {
                                                    aaa = BKolonu.Split(' ');
                                                    adsoyad = aaa[16] + " " + aaa[17];
                                                }
                                                if (aaa[16] == "nolu")
                                                {
                                                    aaa = BKolonu.Split(' ');
                                                    adsoyad = aaa[17] + " " + aaa[18];
                                                }
                                                if (aaa[16] != "nolu" && aaa[12] != "nolu" && aaa[13] != "nolu" && aaa[14] != "nolu" && aaa[15] != "nolu")
                                                {
                                                    adsoyad = "NULL";
                                                }

                                            }
                                            else
                                            {
                                                adsoyad = "NULL";

                                            }
                                        }
                                        else
                                        {
                                            adsoyad = "NULL";

                                        }

                                        //Kategoriler 
                                        if (BKolonu.IndexOf("/") != -1)
                                        {
                                            if (BKolonu.IndexOf("/") != -1)
                                            {
                                                aaa = BKolonu.Split('/');
                                                bbb = aaa[1].ToString();
                                                aaa = bbb.Split('/');
                                                kategori = aaa[0];

                                                if (adsoyad == "NULL")
                                                {
                                                    kategori = "NULL";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            kategori = "NULL";

                                        }

                                        //telefon numarası 
                                        aaa = BKolonu.Split(' ');
                                        for (k = 0; k < aaa.Length; k++)
                                        {
                                            telefon = "";
                                            ccc = aaa[k];
                                            if (ccc.IndexOf("5") != -1 && csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.StartsWith(control) && ccc.Length == 11 || ccc.StartsWith(control2) && ccc.Length == 10) //telefon numarası
                                                {
                                                    bbb = aaa[k].ToString();
                                                    aaa = bbb.Split(' ');


                                                    telefon = aaa[0];
                                                }

                                                else
                                                {
                                                    telefon = "NULL";

                                                }

                                            }
                                            if (ccc.IndexOf("-5") != -1)
                                            {
                                                bbb = aaa[k].ToString();
                                                aaa = bbb.Split('-', ' ');
                                                telefon = aaa[1];

                                            }
                                        }

                                        if (BKolonu.IndexOf("5") == -1)//telefon numarası değilse null
                                        {
                                            telefon = "NULL";

                                        }
                                        else if (csSayisalDegerMi(ccc) == false && ((ccc.Length != 11 && ccc.StartsWith(control)) || (ccc.Length != 10 && ccc.StartsWith(control2))))
                                        {
                                            telefon = "NULL";

                                        }

                                        //Tc Kimlik Numarası

                                        string[] tc = BKolonu.Split(' ');
                                        for (k = 0; k < tc.Length; k++)
                                        {
                                            ccc = tc[k];


                                            if (csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.Length == 11 && !ccc.StartsWith(control))
                                                {
                                                    bbb = tc[k].ToString();
                                                    tc = bbb.Split(' ');

                                                    tc_no = tc[0];

                                                }
                                                else
                                                {
                                                    tc_no = "NULL";

                                                }
                                            }


                                        }
                                        if (csSayisalDegerMi(ccc) == false && (ccc.Length != 11 || ccc.StartsWith(control)))
                                        {
                                            tc_no = "NULL";

                                        }

                                        dt2.Rows.Add(numara, adsoyad, kategori, telefon, tc_no);
                                    }
                                }
                            if (i == 2)
                            {
                                for (int j = 0; j < dt.Rows.Count; j++)
                                {
                                    tutar = dt.Rows[j][i].ToString();
                                    dt2.Rows[j][5] = tutar;

                                }
                            }
                        }
                            gridColon2.DataSource = dt2;
                            gridColon2.DataBind();
                            //DATATABLE'I EXCELE AKTARMA//
                            try
                            {
                                if (dt2 == null || dt2.Columns.Count == 0)
                                    throw new Exception("Datatable boş.");


                                Excel.Application excelApp = new Excel.Application();
                                excelApp.Workbooks.Add();
                                Excel.Workbook xlKitap = excelApp.Workbooks.Open(Session["path_"].ToString(), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                                Excel.Sheets xlSayfa = xlKitap.Worksheets;
                                var xlYeniSayfa = (Excel.Worksheet)xlSayfa.Add(xlSayfa[1], Type.Missing, Type.Missing, Type.Missing);
                                xlYeniSayfa.Name = "Sayfa2";

                                xlYeniSayfa = (Excel.Worksheet)xlKitap.Worksheets.get_Item(1);
                                xlYeniSayfa.Select();



                                // column başlıkları
                                for (int i = 0; i < dt2.Columns.Count; i++)
                                {
                                    xlYeniSayfa.Cells[1, (i + 1)] = dt2.Columns[i].ColumnName;
                                }

                                // rows 
                                for (int i = 0; i < dt2.Rows.Count; i++)
                                {
                                    for (int j = 0; j < dt2.Columns.Count; j++)
                                    {
                                        xlYeniSayfa.Cells[(i + 2), (j + 1)] = dt2.Rows[i][j];
                                    }
                                }

                                //kaydedilecek excelin yolu
                                if (Session["path_"].ToString() != null && Session["path_"].ToString() != "")
                                {
                                    try
                                    {//excele kaydetme                               

                                        xlYeniSayfa.SaveAs(Session["path_"].ToString());
                                        excelApp.Quit();


                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Excel dosyası kaydedilemedi.Dosya yolunu Kontrol et.\n"
                                        + ex.Message);
                                    }
                                }
                                else //dosya yolu yok
                                {
                                    excelApp.Visible = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Hata: \n" + ex.Message);
                            }

                    }

                        if (dropdownlist.Items[8].Selected == true && exceldosyaadı == "Ziraat Bank.xlsx")  //Ziraat Bank
                        {


                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if (i == 1)
                                {
                                    dt2.Columns.Add("No");
                                    dt2.Columns.Add("Ad Soyad");
                                    dt2.Columns.Add("Kategoriler");
                                    dt2.Columns.Add("Telefon");
                                    dt2.Columns.Add("TC No");
                                    dt2.Columns.Add("Bağış Miktarı");
                                    

                                    for (int j = 0; j < dt.Rows.Count; j++)
                                    {
                                        numara = (j + 2).ToString();
                                        BKolonu = " " + dt.Rows[j][i].ToString();
                                        //Ad Soyad
                                        if (BKolonu.IndexOf(" ") != -1)
                                        {
                                            if (BKolonu.IndexOf(" ") != -1)
                                            {
                                                aaa = BKolonu.Split(' ');
                                                bbb = aaa[1].ToString();
                                                aaa = bbb.Split(' ');
                                                adsoyad = aaa[0];
                                            }
                                            else
                                            {
                                                adsoyad = "NULL";

                                            }
                                        }
                                        if (BKolonu.IndexOf("Gönd:") != -1)
                                        {
                                            aaa = BKolonu.Split(' ');
                                            adsoyad = aaa[2] + " " + aaa[3];
                                        }

                                        else
                                        {
                                            adsoyad = "NULL";

                                        }

                                        //Kategoriler 
                                        if (BKolonu.IndexOf(" ") != -1)
                                        {
                                            if (BKolonu.IndexOf(" ") != -1)
                                            {
                                                aaa = BKolonu.Split(' ');
                                                bbb = aaa[1].ToString();
                                                aaa = bbb.Split(' ');
                                                kategori = aaa[0];

                                            }

                                        }

                                        else
                                        {
                                            kategori = "NULL";

                                        }

                                        //telefon numarası 
                                        aaa = BKolonu.Split(' ');
                                        for (k = 0; k < aaa.Length; k++)
                                        {
                                            telefon = "";
                                            ccc = aaa[k];
                                            if (ccc.IndexOf("5") != -1 && csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.StartsWith(control) && ccc.Length == 11 || ccc.StartsWith(control2) && ccc.Length == 10) //telefon numarası
                                                {
                                                    bbb = aaa[k].ToString();
                                                    aaa = bbb.Split(' ');


                                                    telefon = aaa[0];
                                                }

                                                else
                                                {
                                                    telefon = "NULL";

                                                }

                                            }
                                            if (ccc.IndexOf("-5") != -1)
                                            {
                                                bbb = aaa[k].ToString();
                                                aaa = bbb.Split('-', ' ');
                                                telefon = aaa[1];

                                            }
                                        }

                                        if (BKolonu.IndexOf("5") == -1)//telefon numarası değilse null
                                        {
                                            telefon = "NULL";

                                        }
                                        else if (csSayisalDegerMi(ccc) == false && ((ccc.Length != 11 && ccc.StartsWith(control)) || (ccc.Length != 10 && ccc.StartsWith(control2))))
                                        {
                                            telefon = "NULL";

                                        }

                                        //Tc Kimlik Numarası

                                        string[] tc = BKolonu.Split(' ');
                                        for (k = 0; k < tc.Length; k++)
                                        {
                                            ccc = tc[k];


                                            if (csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.Length == 11 && !ccc.StartsWith(control))
                                                {
                                                    bbb = tc[k].ToString();
                                                    tc = bbb.Split(' ');

                                                    tc_no = tc[0];

                                                }
                                                else
                                                {
                                                    tc_no = "NULL";

                                                }
                                            }


                                        }
                                        if (csSayisalDegerMi(ccc) == false && (ccc.Length != 11 || ccc.StartsWith(control)))
                                        {
                                            tc_no = "NULL";

                                        }

                                        dt2.Rows.Add(numara, adsoyad, kategori, telefon, tc_no);
                                    }
                                }
                            if (i == 2)
                            {
                                for (int j = 0; j < dt.Rows.Count; j++)
                                {
                                    tutar = dt.Rows[j][i].ToString();
                                    dt2.Rows[j][5] = tutar;

                                }
                            }
                        }
                            gridColon2.DataSource = dt2;
                            gridColon2.DataBind();
                            //DATATABLE'I EXCELE AKTARMA//
                            try
                            {
                                if (dt2 == null || dt2.Columns.Count == 0)
                                    throw new Exception("Datatable boş.");


                                Excel.Application excelApp = new Excel.Application();
                                excelApp.Workbooks.Add();
                                Excel.Workbook xlKitap = excelApp.Workbooks.Open(Session["path_"].ToString(), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                                Excel.Sheets xlSayfa = xlKitap.Worksheets;
                                var xlYeniSayfa = (Excel.Worksheet)xlSayfa.Add(xlSayfa[1], Type.Missing, Type.Missing, Type.Missing);
                                xlYeniSayfa.Name = "Sayfa2";

                                xlYeniSayfa = (Excel.Worksheet)xlKitap.Worksheets.get_Item(1);
                                xlYeniSayfa.Select();



                                // column başlıkları
                                for (int i = 0; i < dt2.Columns.Count; i++)
                                {
                                    xlYeniSayfa.Cells[1, (i + 1)] = dt2.Columns[i].ColumnName;
                                }

                                // rows 
                                for (int i = 0; i < dt2.Rows.Count; i++)
                                {
                                    for (int j = 0; j < dt2.Columns.Count; j++)
                                    {
                                        xlYeniSayfa.Cells[(i + 2), (j + 1)] = dt2.Rows[i][j];
                                    }
                                }

                                //kaydedilecek excelin yolu
                                if (Session["path_"].ToString() != null && Session["path_"].ToString() != "")
                                {
                                    try
                                    {//excele kaydetme                               

                                        xlYeniSayfa.SaveAs(Session["path_"].ToString());
                                        excelApp.Quit();


                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Excel dosyası kaydedilemedi.Dosya yolunu Kontrol et.\n"
                                        + ex.Message);
                                    }
                                }
                                else //dosya yolu yok
                                {
                                    excelApp.Visible = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Hata: \n" + ex.Message);
                            }


                   

                    }
                        if (dropdownlist.Items[9].Selected == true && exceldosyaadı == "Ziraat Katılım.xlsx") //Ziraat Katılım
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if (i == 1)
                                {
                                    dt2.Columns.Add("No");
                                    dt2.Columns.Add("Ad Soyad");
                                    dt2.Columns.Add("Kategoriler");
                                    dt2.Columns.Add("Telefon");
                                    dt2.Columns.Add("TC No");
                                dt2.Columns.Add("Bağış Miktarı");

                                    for (int j = 0; j < dt.Rows.Count; j++)
                                    {
                                        numara = (j + 2).ToString();
                                        BKolonu = " " + dt.Rows[j][i].ToString();
                                        //Ad Soyad
                                        if (BKolonu.IndexOf(",") != -1)
                                        {
                                            if (BKolonu.IndexOf(",") != -1)
                                            {
                                                aaa = BKolonu.Split(',');
                                                bbb = aaa[1].ToString();
                                                aaa = bbb.Split(',');
                                                adsoyad = aaa[0];
                                            }
                                            else
                                            {
                                                adsoyad = "NULL";

                                            }
                                        }
                                        if (BKolonu.IndexOf("SN") != -1)
                                        {
                                            aaa = BKolonu.Split(' ');
                                            adsoyad = aaa[2] + " " + aaa[3];
                                        }
                                        else if (BKolonu.IndexOf("Gelen EFT Gönderen:") != -1)
                                        {
                                            aaa = BKolonu.Split(' ');
                                            adsoyad = aaa[4] + " " + aaa[5];
                                        }
                                        else
                                        {
                                            adsoyad = "NULL";

                                        }

                                        //Kategoriler 
                                        if (BKolonu.IndexOf("/") != -1)
                                        {
                                            if (BKolonu.IndexOf(" ") != -1)
                                            {
                                                aaa = BKolonu.Split('/');
                                                bbb = aaa[1].ToString();
                                                aaa = bbb.Split(' ');
                                                kategori = aaa[0];

                                            }
                                        }
                                        if (BKolonu.IndexOf("SN:") != -1)
                                        {
                                            aaa = BKolonu.Split(' ');
                                            kategori = aaa[4];
                                        }
                                        else
                                        {
                                            kategori = "NULL";

                                        }

                                        //telefon numarası 
                                        aaa = BKolonu.Split(' ');
                                        for (k = 0; k < aaa.Length; k++)
                                        {
                                            telefon = "";
                                            ccc = aaa[k];
                                            if (ccc.IndexOf("5") != -1 && csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.StartsWith(control) && ccc.Length == 11 || ccc.StartsWith(control2) && ccc.Length == 10) //telefon numarası
                                                {
                                                    bbb = aaa[k].ToString();
                                                    aaa = bbb.Split(' ');


                                                    telefon = aaa[0];
                                                }

                                                else
                                                {
                                                    telefon = "NULL";

                                                }

                                            }
                                            if (ccc.IndexOf("-5") != -1)
                                            {
                                                bbb = aaa[k].ToString();
                                                aaa = bbb.Split('-', ' ');
                                                telefon = aaa[1];

                                            }
                                        }

                                        if (BKolonu.IndexOf("5") == -1)//telefon numarası değilse null
                                        {
                                            telefon = "NULL";

                                        }
                                        else if (csSayisalDegerMi(ccc) == false && ((ccc.Length != 11 && ccc.StartsWith(control)) || (ccc.Length != 10 && ccc.StartsWith(control2))))
                                        {
                                            telefon = "NULL";

                                        }

                                        //Tc Kimlik Numarası

                                        string[] tc = BKolonu.Split(' ');
                                        for (k = 0; k < tc.Length; k++)
                                        {
                                            ccc = tc[k];


                                            if (csSayisalDegerMi(ccc) == true)
                                            {
                                                if (ccc.Length == 11 && !ccc.StartsWith(control))
                                                {
                                                    bbb = tc[k].ToString();
                                                    tc = bbb.Split(' ');

                                                    tc_no = tc[0];

                                                }
                                                else
                                                {
                                                    tc_no = "NULL";

                                                }
                                            }
                                        }
                                        if (csSayisalDegerMi(ccc) == false && (ccc.Length != 11 || ccc.StartsWith(control)))
                                        {
                                            tc_no = "NULL";

                                        }

                                        dt2.Rows.Add(numara, adsoyad, kategori, telefon, tc_no);
                                    }
                                }
                            if (i == 2)
                            {
                                for (int j = 0; j < dt.Rows.Count; j++)
                                {
                                    tutar = dt.Rows[j][i].ToString();
                                    dt2.Rows[j][5] = tutar;

                                }
                            }
                        }
                            gridColon2.DataSource = dt2;
                            gridColon2.DataBind();
                            //DATATABLE'I EXCELE AKTARMA//
                            try
                            {
                                if (dt2 == null || dt2.Columns.Count == 0)
                                    throw new Exception("Datatable boş.");


                                Excel.Application excelApp = new Excel.Application();
                                excelApp.Workbooks.Add();
                                Excel.Workbook xlKitap = excelApp.Workbooks.Open(Session["path_"].ToString(), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                                Excel.Sheets xlSayfa = xlKitap.Worksheets;
                                var xlYeniSayfa = (Excel.Worksheet)xlSayfa.Add(xlSayfa[1], Type.Missing, Type.Missing, Type.Missing);
                                xlYeniSayfa.Name = "Sayfa2";

                                xlYeniSayfa = (Excel.Worksheet)xlKitap.Worksheets.get_Item(1);
                                xlYeniSayfa.Select();

                                // column başlıkları
                                for (int i = 0; i < dt2.Columns.Count; i++)
                                {
                                    xlYeniSayfa.Cells[1, (i + 1)] = dt2.Columns[i].ColumnName;
                                }

                                // rows 
                                for (int i = 0; i < dt2.Rows.Count; i++)
                                {
                                    for (int j = 0; j < dt2.Columns.Count; j++)
                                    {
                                        xlYeniSayfa.Cells[(i + 2), (j + 1)] = dt2.Rows[i][j];
                                    }
                                }

                                //kaydedilecek excelin yolu
                                if (Session["path_"].ToString() != null && Session["path_"].ToString() != "")
                                {
                                    try
                                    {//excele kaydetme                               

                                        xlYeniSayfa.SaveAs(Session["path_"].ToString());
                                        excelApp.Quit();



                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Excel dosyası kaydedilemedi.Dosya yolunu Kontrol et.\n"
                                        + ex.Message);
                                    }
                                }
                                else //dosya yolu yok
                                {
                                    excelApp.Visible = true;
                                }                           

                        }                   
                            catch (Exception ex)
                            {
                                throw new Exception("Hata: \n" + ex.Message);
                            }
                    }                        
                    else if (fluExcel.PostedFile.FileName!=(dropdownlist.SelectedItem.ToString() + ".xlsx") )
                    {
                        lblMessage.Text = "Banka Seçimini Kontrol Edin!";
                    }
                   
                }               
            }
               

            catch (Exception ex)
            {
               // lblMessage.Text = "Hata:" + ex.Message;
            }

            //if (File.Exists(Session["path_"].ToString())) //App_Data klasörüne kopyalanan excel dosyalarını siliyor..
            //{
            //    File.Delete(Session["path_"].ToString());
            //}
        }

        //verinin sayısal değer olup olmadığını kontrol ediyor.
        public Boolean csSayisalDegerMi(String strVeri)
        {
            if (String.IsNullOrEmpty(strVeri) == true)
            {
                return false;
            }
            else
            {
                Regex desen = new Regex("^[0-9]*$");
                return desen.IsMatch(strVeri);
            }
        }
       
        protected void kaydet_Click(object sender, EventArgs e)
        {                    
            string date = DateTime.Now.ToString("dd-MM-yyyy HH-mm");
            string newFileName2 = date + Session["path2_"].ToString(); ;
            String FilePath = Session["path_"].ToString();
            System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
            response.ClearContent();
            response.Clear();
            response.ContentType = "text/plain";
            response.AddHeader("Content-Disposition", "attachment; filename=" + newFileName2);
         //   string dene = Session["path_"].ToString();
          //  response.AppendHeader("Content-Disposition", "attachment; filename=" + Session["path_"].ToString());
            response.TransmitFile(FilePath);
            response.Flush();
            response.End();
                 

        }
    }
}



