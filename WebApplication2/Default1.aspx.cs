using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication2
{
    //public class Class1
    //{
    //}
    public partial class Default1 : System.Web.UI.Page
    {
        DataTable dt = new DataTable();
        DataTable CSVTable = new DataTable();
        protected void Page_Load(object sender, EventArgs e)
        {
        }
        /// <summary>  
        /// Exporting of the .CSV file.  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        protected void btnExport_OnClick(object sender, EventArgs e)
        {
            try
            {
                Response.Clear();
                Response.ClearContent();
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("ItemType,OrderPriority,OrderDate,UnitsSold,UnitPrice,TotalRevenue,TotalCost");
                Response.ContentType = "application/x-msexcel";

                Response.AddHeader("content-disposition", "attachment; filename=SalesRecords.csv");
                Response.Write(sb.ToString());
                if (ViewState["CSVTable"] != "")
                {
                    dt = ViewState["CSVTable"] as DataTable;
                    if ((dt != null) && (dt.Rows.Count > 0))
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            sb = new StringBuilder((string)row[0]);
                            for (int i = 1; i < dt.Columns.Count; i++)
                            {
                                if (row[i] is DBNull)
                                    sb.Append(",NULL");
                                else if (i == 2)
                                    sb.Append("," + new DateTime((long)row[i]).ToString("G"));
                                else
                                    sb.Append("," + row[i].ToString());

                            }
                            sb.AppendLine();
                            Response.Write(sb.ToString());
                        }
                    }
                }
                Response.Flush();
                Response.Close();
                Response.End();
            }
            catch (Exception ex) { }

        }
        /// <summary>  
        /// Importing of .CSV file.  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        protected void btn_import_Click(object sender, EventArgs e)
        {
            try
            {
                bool hasFile = FileUpload1.HasFile;
                if (hasFile)
                {
                    int flag = 0;
                    string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
                    string RandomName = DateTime.Now.ToFileTime().ToString();
                    string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
                    string FolderPath = "~/uploads/";

                    string FilePath = Server.MapPath(FolderPath + RandomName + FileName);

                    string[] filenames = Directory.GetFiles(Server.MapPath("~/uploads"));

                    if (filenames.Length > 0)
                    {
                        foreach (string filename in filenames)
                        {
                            if (FilePath == filename)
                            {

                                flag = 1;
                                break;
                            }
                        }

                        if (flag == 0)
                        {
                            FileUpload1.SaveAs(FilePath);
                            ReadCSVFile(FilePath);
                        }

                    }
                    else
                    {
                        FileUpload1.SaveAs(FilePath);
                        ReadCSVFile(FilePath);
                    }
                }
                else
                {
                    String msg = "Select a file then try to import";
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>  
        /// Reading of the .CSV file.  
        /// </summary>  
        /// <param name="fileName"></param>  
        public void ReadCSVFile(string fileName)
        {
            try
            {
                string connection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}\\;Extended Properties='Text;HDR=Yes;FMT=CSVDelimited'";

                connection = String.Format(connection, Path.GetDirectoryName(fileName));


                OleDbDataAdapter csvAdapter;
                csvAdapter = new OleDbDataAdapter("SELECT * FROM [" + Path.GetFileName(fileName) + "]", connection);

                if (File.Exists(fileName) && new FileInfo(fileName).Length > 0)
                {
                    try
                    {
                        csvAdapter.Fill(CSVTable);
                        if ((CSVTable != null) && (CSVTable.Rows.Count > 0))
                        {
                            ViewState["CSVTable"] = CSVTable;
                            wdgList.DataSource = CSVTable;
                            wdgList.DataBind();
                        }
                        else
                        {
                            String msg = "No records found";
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(String.Format("Error reading Table {0}.\n{1}", Path.GetFileName(fileName), ex.Message));
                    }
                }
            }
            catch (Exception ex) { }
        }
    }
}