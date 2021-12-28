using Dapper;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Uploadfiledata
{
    public partial class UploadIndex : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 上传数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Button1_Click(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            if (FileUpload1.HasFile == false)
            {
                uploadmessage.Text = "请上传文件！";
            }
            string IsXls = Path.GetExtension(FileUpload1.FileName).ToString().ToLower();//System.IO.Path.GetExtension获得文件的扩展名
            if (IsXls != ".xlsx")
            {
                Response.Write(FileUpload1.FileName);
                Response.Write("<script>alert('只可以选择Excel文件且后缀为.xlsx')</script>");
                return;//当选择的不是Excel文件时,返回
            }
            string filename = FileUpload1.FileName;
            string savePath = Server.MapPath("~/ExcelTemp/" + filename);//Server.MapPath 获得虚拟服务器相对路径
            FileUpload1.SaveAs(savePath);
            IWorkbook wk = null;
            string exion = Path.GetExtension(savePath);
            try
            {
                FileStream fs = File.OpenRead(savePath);
                //if (splitfilename == ".xls")
                //{
                //    wk = new HSSFWorkbook(fs);
                //}
                //if (splitfilename == ".xlsx")
                //{
                    wk = new XSSFWorkbook(fs);
                //}
                fs.Close();
                //读取当前表数据
                ISheet sheet = wk.GetSheetAt(0);
                //读取当前行数据
                IRow row = sheet.GetRow(0);
                int offset = sheet.LastRowNum+1;
                for (int i = 0; i < offset; i++)
                {
                    if (i != 0)//过滤表头
                    {
                        string addsql = string.Empty;
                        string key = string.Empty;
                        string clientid = string.Empty;
                        string product = string.Empty;
                        string iccid = string.Empty;
                        string imei = string.Empty;
                        string imsi = string.Empty;
                        string software_version = string.Empty;
                        string lac = string.Empty;
                        //string cid = string.Empty;
                        row = sheet.GetRow(i);
                        if (row != null)
                        {             
                            clientid = row.GetCell(0).ToString();
                            product = row.GetCell(1).ToString();
                            key = row.GetCell(2).ToString();
                            lac = row.GetCell(3).ToString();
                            //cid = row.GetCell(4).ToString();
                            imsi = row.GetCell(5).ToString();
                            imei = row.GetCell(6).ToString();
                            iccid = row.GetCell(7).ToString();
                            software_version = row.GetCell(8).ToString();
                            addsql = "insert into mqttpublish(clientid,product,keyss,lac,imsi,imei,iccid,software_version,addtime)" +
                            "values('"+clientid+"','"+product+"','"+key+"','"+lac+"','"+imsi+"','"+imei+"','"+iccid+"','"+software_version+"','"+time+"')";
                            using (IDbConnection conn = DapperService.MySqlConnection())
                            {
                                conn.Execute(addsql);
                            }
                        }
                    }
                }
                Response.Write("<script>alert('上传成功')</script>");
            }
            catch (Exception ex)
            {
                using (IDbConnection conn = DapperService.MySqlConnection())
                {
                    string sql = "insert into mqttpublishlog(logcontent,logtype,Addtime)values('" + ex + "','0','" + time + "')";
                    conn.Execute(sql);
                }
                Response.Write("<script>alert('上传失败，请联系管理员')</script>");
                uploadmessage.Text = ex.ToString();
            }

           
        }

        //protected void DownTemp_Click(object sender, EventArgs e)
        //{
        //    string fileName = "Temp.xlsx";//客户端保存的文件名
        //    string filePath = Server.MapPath("DownTemp/Temp.xlsx");//路径
        //    //以字符流的形式下载文件
        //    FileStream fs = new FileStream(filePath, FileMode.Open);
        //    byte[] bytes = new byte[(int)fs.Length];
        //    fs.Read(bytes, 0, bytes.Length);
        //    fs.Close();
        //    Response.ContentType = "application/octet-stream";
        //    //通知浏览器下载文件而不是打开
        //    Response.AddHeader("Content-Disposition", "attachment;  filename=" + HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8));
        //    Response.BinaryWrite(bytes);
        //    Response.Flush();
        //    Response.End();
        //}
    }
    
}