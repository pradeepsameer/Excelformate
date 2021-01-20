# Excelformate
Import image from Excel sheet into SQL Server table C#
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing.Imaging;
using Spire.Xls;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;


public partial class WorkbookDefault : System.Web.UI.Page
{
    int value = 0;
    string values = string.Empty;
    string valuess = string.Empty;
    string valuesss = string.Empty;
    byte[] imgdata;
    byte[] imgdataaddress;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\Umesh\Documents\Visual Studio 2012\WebSites\ExcelimageWebSite1\file\Ckyc.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            //CellRange range = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn]; 
            CellRange range = sheet.Range[1, 1, 1, 5];
            for (int j = 2; j <= 7; j++)
            {
                value = Convert.ToInt16(sheet.Range[j, 1].Value);
                values = sheet.Range[j, 2].Value;
                valuess = sheet.Range[j, 3].Value;
                valuesss = sheet.Range[j, 4].Value;
                if (sheet.Range[j, 5].HasPictures == true)
                {
                    ExcelPicture picture = sheet.Pictures[j - 2];
                    picture.Picture.Save(Server.MapPath("~/Image/" + picture.Name.Replace(" ", "" + value + "_") + ".jpg"), ImageFormat.Jpeg);
                    imgdata = System.IO.File.ReadAllBytes(HttpContext.Current.Server.MapPath("~/Image/" + picture.Name.Replace(" ", "" + value + "_") + ".jpg"));
                    string patha = Server.MapPath("~/Image/" + picture.Name.Replace(" ", "" + value + "_") + ".jpg");
                    if (System.IO.File.Exists(Server.MapPath("~/Image/" + picture.Name.Replace(" ", "" + value + "_") + ".jpg")))
                    {
                        System.IO.File.Delete(patha);
                    }
                    picture.Remove();
                }
                //if (sheet.Range[j, 6].HasPictures == true)
                //{

                //    ExcelPicture picture = sheet.Pictures[j - 2];
                //    picture.Picture.Save(Server.MapPath("~/Image/" + picture.Name.Replace(" ", "" + value + "_") + ".jpg"), ImageFormat.Jpeg);
                //    imgdataaddress = System.IO.File.ReadAllBytes(HttpContext.Current.Server.MapPath("~/Image/" + picture.Name.Replace(" ", "" + value + "_") + ".jpg"));



                //    string patha = Server.MapPath("~/Image/" + picture.Name.Replace(" ", "" + value + "_") + ".jpg");
                //    if (System.IO.File.Exists(Server.MapPath("~/Image/" + picture.Name.Replace(" ", "" + value + "_") + ".jpg")))
                //    {
                //        System.IO.File.Delete(patha);
                //    }
                //    picture.Remove();
                //}

                InsertImages(value, values, valuess, valuesss, imgdata, imgdataaddress);
            }
        }
    }

    public void InsertImages(int VALINT, string EMP_NAME, string dEPART, string designation, byte[] imgdata, byte[] imgdataaddress)
    {

        using (var con = new SqlConnection(ConfigurationManager.ConnectionStrings["constrd"].ConnectionString))
        {
            //Set up the command
            var cmd = new SqlCommand("SP_Import_Excel", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@employee_id", VALINT);
            cmd.Parameters.AddWithValue("@employee_name", EMP_NAME);
            cmd.Parameters.AddWithValue("@department", dEPART);
            cmd.Parameters.AddWithValue("@designation", designation);
            cmd.Parameters.AddWithValue("@employee_pic", imgdata);
            cmd.Parameters.AddWithValue("@Addressproff_pic", imgdataaddress);
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        using (var con = new SqlConnection(ConfigurationManager.ConnectionStrings["constrd"].ConnectionString))
        {
            con.Open();
            //Set up the command
            var cmd = new SqlCommand("SP_SearchExcel", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@employee_name", checkpic.Text.Trim());
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            sda.Fill(ds);

            byte[] arr = (byte[])ds.Tables[0].Rows[0]["employee_pic"];
            string base64String = Convert.ToBase64String(arr, 0, arr.Length);
            img.Src = "data:image/jpg;base64," + base64String;
            //byte[] arr1 = (byte[])ds.Tables[0].Rows[0]["Addressproff_pic"];
            //string base64Stringadd = Convert.ToBase64String(arr1, 0, arr1.Length);
            //img1.Src = "data:image/jpg;base64," + base64Stringadd;
            con.Close();
        }
    }
}
