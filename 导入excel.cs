using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
namespace TempCms.comment
{
聽 聽 /// <summary>
聽 聽 /// Excl_In 鐨勬憳瑕佽鏄?
聽 聽 /// </summary>
聽 聽 public class Excl_In : IHttpHandler
聽 聽 {
聽 聽 聽 聽 public void ProcessRequest(HttpContext context)
聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 context.Response.ContentType = "text/plain";
聽 聽 聽 聽 聽 聽 HttpPostedFile file = context.Request.Files["daoruFile"];
聽 聽 聽 聽 聽 聽 string saname = context.Request["saname"];
聽 聽 聽 聽 聽 聽 string resule = "";
聽 聽 聽 聽 聽 聽 string fileExtenSion;
聽 聽 聽 聽 聽 聽 fileExtenSion = Path.GetExtension(file.FileName);
聽 聽 聽 聽 聽 聽 if (fileExtenSion.ToLower() != ".xls" && fileExtenSion.ToLower() != ".xlsx")
聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 resule="涓婁紶鐨勬枃浠舵牸寮忎笉姝ｇ‘";
聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 else
聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 DataTable dt = xsldata(file, saname, fileExtenSion);//excle杞崲鎴恉atatable
聽 聽 聽 聽 聽 聽 聽 聽 resule = DataInSql(dt);
聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 context.Response.Write(resule);
聽 聽 聽 聽 }
聽 聽 聽 聽 private string DataInSql(DataTable dt)
聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 try
聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 //dataGridView2.DataSource = ds.Tables[0]; 聽
聽 聽 聽 聽 聽 聽 聽 聽 int insertcount = 0;//璁板綍鎻掑叆鎴愬姛鏉℃暟 聽
聽 聽 聽 聽 聽 聽 聽 聽 int updatecount = 0;//璁板綍鏇存柊淇℃伅鏉℃暟 聽
聽 聽 聽 聽 聽 聽 聽 聽 string strcon = "server=(local);database=HXcms;uid=sa;pwd=sql@2008";
聽 聽 聽 聽 聽 聽 聽 聽 //string strcon = "server=192.168.0.190;database=HXcms;uid=sa;pwd=sql@2008";
聽 聽 聽 聽 聽 聽 聽 聽 SqlConnection conn = new SqlConnection(strcon);//閾炬帴鏁版嵁搴?聽
聽 聽 聽 聽 聽 聽 聽 聽 conn.Open();
聽 聽 聽 聽 聽 聽 聽 聽 for (int i = 0; i < dt.Rows.Count; i++)
聽 聽 聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 string cardNum = dt.Rows[i][0] == null ? "" : dt.Rows[i][0].ToString();//dt.Rows[i]["Name"].ToString(); "Name"鍗充负Excel涓璑ame鍒楃殑琛ㄥご 聽
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 string uname = dt.Rows[i][1] == null ? "" : dt.Rows[i][1].ToString();
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 int age = dt.Rows[i][2] == null ? 0 : int.Parse(dt.Rows[i][2].ToString());
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 string sex = dt.Rows[i][3] == null ? "" : dt.Rows[i][3].ToString();
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 switch (sex)
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 case "鐢?:
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 sex = "1";
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 break;
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 case "濂?:
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 sex = "2";
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 break;
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 default:
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 break;
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 string tel = dt.Rows[i][4] == null ? "" : dt.Rows[i][4].ToString(); ;
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 string address = dt.Rows[i][5] == null ? "" : dt.Rows[i][5].ToString();
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 string shenfenzheng = dt.Rows[i][6] == null ? "" : dt.Rows[i][6].ToString();
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 string pwd = dt.Rows[i][7] == null ? "" : dt.Rows[i][7].ToString();
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 decimal yu_e = dt.Rows[i][8] == null ? 0 : decimal.Parse(dt.Rows[i][8].ToString());
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 decimal zong_e_jia = dt.Rows[i][9] == null ? 0 : decimal.Parse(dt.Rows[i][9].ToString());
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 decimal zong_e_jian = dt.Rows[i][10] == null ? 0 : decimal.Parse(dt.Rows[i][10].ToString());
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 int jifen_zong = dt.Rows[i][11] == null ? 0 : int.Parse(dt.Rows[i][11].ToString());
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 int jifen = dt.Rows[i][12] == null ? 0 : int.Parse(dt.Rows[i][12].ToString());
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 string saname = "gzxxjs";
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 //if (Name != "" && Sex != "" && Age != 0 && Address != "")
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 //{
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 SqlCommand selectcmd = new SqlCommand("select count(*) from wxMemberCard where Wxmcstkh='" + cardNum + "' and Wxmcusername='" + "gzxxjs'", conn);
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 int count = Convert.ToInt32(selectcmd.ExecuteScalar());
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 if (count > 0)
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 updatecount++;
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 else
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 {
						string sql="insert into wxMemberCard(Wxmcnum,Wxmcstate,Wxmcshzt,Wxmcstkh,Wxmcname,wxmage,Wxmgender,Wxmctel,Wxmidress,Wxmidentity,wxmcstkhpwd,Wxmchykye,Wxmchykczze,Wxmchykxfze,Wxmclszjf,Wxmcjifen,Wxmcusername) values('',0,2,'" + cardNum + "','" + uname + "'," + age + ",'" + sex + "','" + tel + "','" + address + "','" + shenfenzheng + "','" + pwd + "'," + yu_e + "," + zong_e_jia + "," + zong_e_jian + "," + jifen_zong + "," + jifen + ",'" + saname + "')";
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 SqlCommand insertcmd = new SqlCommand(sql, conn);
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 insertcmd.ExecuteNonQuery();
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 insertcount++;
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 //}
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 //else
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 //{
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 // 聽 聽errorcount++;
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 //}
聽 聽 聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 聽 聽 return "鎴愬姛涓婁紶锛?" +insertcount + ")鏉℃暟鎹紒</br>閲嶅鏁版嵁锛?" + updatecount + ")鏉?;
聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 catch (Exception ex)
聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 return "鎿嶄綔澶辫触";
聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 }
聽 聽 聽 聽 /// <summary>
聽 聽 聽 聽 /// 璁瞖xcle杞崲鎴恉atatable
聽 聽 聽 聽 /// </summary>
聽 聽 聽 聽 /// <param name="file">excle鏂囦欢</param>
聽 聽 聽 聽 /// <param name="saname">鐧诲綍鍚?/param>
聽 聽 聽 聽 /// <param name="fileExtenSion">鎵╁睍鍚?/param>
聽 聽 聽 聽 /// <returns></returns>
聽 聽 聽 聽 private DataTable xsldata(HttpPostedFile file, string saname, string fileExtenSion)
聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 try
聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 string FileName = "App_Data/" + Path.GetFileName(file.FileName);
聽 聽 聽 聽 聽 聽 聽 聽 if (File.Exists(HttpRuntime.AppDomainAppPath + saname + "/" + file.FileName))
聽 聽 聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 File.Delete(HttpRuntime.AppDomainAppPath + saname + "/" + file.FileName);
聽 聽 聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 聽 聽 file.SaveAs(HttpRuntime.AppDomainAppPath + saname + "/" + file.FileName);
聽 聽 聽 聽 聽 聽 聽 聽 //HDR=Yes锛岃繖浠ｈ〃绗竴琛屾槸鏍囬锛屼笉鍋氫负鏁版嵁浣跨敤 锛屽鏋滅敤HDR=NO锛屽垯琛ㄧず绗竴琛屼笉鏄爣棰橈紝鍋氫负鏁版嵁鏉ヤ娇鐢ㄣ€傜郴缁熼粯璁ょ殑鏄痀ES 聽
聽 聽 聽 聽 聽 聽 聽 聽 string connstr2003 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + HttpRuntime.AppDomainAppPath + saname + "/" + file.FileName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
聽 聽 聽 聽 聽 聽 聽 聽 string connstr2007 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + HttpRuntime.AppDomainAppPath + saname + "/" + file.FileName + ";Extended Properties=\"Excel 12.0;HDR=YES\"";
聽 聽 聽 聽 聽 聽 聽 聽 OleDbConnection conn;
聽 聽 聽 聽 聽 聽 聽 聽 if (fileExtenSion.ToLower() == ".xls")
聽 聽 聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 conn = new OleDbConnection(connstr2003);
聽 聽 聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 聽 聽 else
聽 聽 聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 conn = new OleDbConnection(connstr2007);
聽 聽 聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 聽 聽 conn.Open();
聽 聽 聽 聽 聽 聽 聽 聽 string sql = "select * from [Sheet1$]";
聽 聽 聽 聽 聽 聽 聽 聽 OleDbCommand cmd = new OleDbCommand(sql, conn);
聽 聽 聽 聽 聽 聽 聽 聽 DataTable dt = new DataTable();
聽 聽 聽 聽 聽 聽 聽 聽 OleDbDataReader sdr = cmd.ExecuteReader();
聽 聽 聽 聽 聽 聽 聽 聽 dt.Load(sdr);
聽 聽 聽 聽 聽 聽 聽 聽 sdr.Close();
聽 聽 聽 聽 聽 聽 聽 聽 conn.Close();
聽 聽 聽 聽 聽 聽 聽 聽 //鍒犻櫎鏈嶅姟鍣ㄩ噷涓婁紶鐨勬枃浠?聽
聽 聽 聽 聽 聽 聽 聽 聽 if (File.Exists(HttpRuntime.AppDomainAppPath + saname + "/" + file.FileName))
聽 聽 聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 聽 聽 File.Delete(HttpRuntime.AppDomainAppPath + saname + "/" + file.FileName);
聽 聽 聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 聽 聽 return dt;
聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 聽 聽 catch (Exception e)
聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 return null;
聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 }
聽 聽 聽 聽 public bool IsReusable
聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 get
聽 聽 聽 聽 聽 聽 {
聽 聽 聽 聽 聽 聽 聽 聽 return false;
聽 聽 聽 聽 聽 聽 }
聽 聽 聽 聽 }
聽 聽 }
}