using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;

namespace PMWeb
{
    public partial class MissionCreat : System.Web.UI.Page
    {
        Systemset SS = new Systemset();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {                
                try
                {
                    Session["ProjectID"] = Request.QueryString["ProjectID"].ToString(); //取得專案編號
                    Session["UserID"] = Request.QueryString["UserID"].ToString(); //取得使用者ID
                    Session["Name"] = Request.QueryString["Name"].ToString(); //取得使用者姓名                    
                    TextBox_DateStart.Text = DateTime.Now.ToString("yyyy-MM-dd");
                    TextBox_DateEnd.Text = DateTime.Now.AddDays(3).ToString("yyyy-MM-dd");
                    Session["FileName"] = new string[0];
                    Session["FileSize"] = new int[0];
                    Session["FileExtension"] = new string[0];

                    #region 初始化listbox

                    string initTsql = "SELECT ID, TypeName FROM PMType WHERE (Alive = '1') ORDER BY ID;";
                    initTsql += "SELECT UserID, MemberName FROM PMPMember WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (Permission = '0') AND (Alive = '1') ORDER BY MemberName";

                    DataSet ListBoxInit = SS.GetSqlTable(initTsql);

                    #region 任務類型
                    if (ListBoxInit.Tables[0].Rows.Count > 0)
                    {
                        DropDownList_MType.DataSource = ListBoxInit;
                        DropDownList_MType.DataBind();
                    }
                    else //如果任務類型一項都沒有的話，就直接新增一項其他
                    {
                        string FirstTypeTsql = "INSERT INTO PMType (TypeName, SaveDate, Alive) VALUES ('其他','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','1')";
                        int FirstTypeWrite = SS.SetSql(FirstTypeTsql);

                        if (FirstTypeWrite > 0)
                        {
                            string GetTypeListTsql = "SELECT ID, TypeName FROM PMType WHERE (Alive = '1') ORDER BY ID";
                            DataSet GetTypeList = SS.GetSqlTable(GetTypeListTsql);

                            DropDownList_MType.DataSource = GetTypeList;
                            DropDownList_MType.DataBind();
                        }
                        else
                        {
                            Response.Write("<script> alert('任務類型列表讀取發生錯誤');</script>");
                        }
                    }
                    #endregion

                    #region 負責人

                    for (int i = 0; i < ListBoxInit.Tables[1].Rows.Count; i++)
                    {
                        ListItem PMember = new ListItem();
                        PMember.Text = ListBoxInit.Tables[1].Rows[i]["MemberName"].ToString();
                        PMember.Value = ListBoxInit.Tables[1].Rows[i]["UserID"].ToString();
                        DropDownList_Leader.Items.Add(PMember);
                    }

                    #endregion

                    GetHelperList();

                    #endregion
                }
                catch
                {
                    Response.Write("<script> alert('系統錯誤');</script>");
                    Server.Transfer("PMLobby.aspx");
                    Response.End();
                }               
            }
        }

        protected void Button_Edit_Click(object sender, EventArgs e)
        {
            if (TextBox_DateStart.Text == string.Empty || TextBox_DateEnd.Text == string.Empty)
            {
                Label_MissionStatus.Text = "日期欄不可為空";
            }
            else
            {
                DateTime StartDate = Convert.ToDateTime(TextBox_DateStart.Text);
                DateTime EndDate = Convert.ToDateTime(TextBox_DateEnd.Text);

                if (TextBox_Title.Text == string.Empty)
                {
                    Label_MissionStatus.Text = "標題欄不可為空";
                }

                else if (StartDate > EndDate)
                {
                    Label_MissionStatus.Text = "截止時間不可比開始時間早";
                }
                                    
                else if (StartDate < Convert.ToDateTime("2014-10-01") || EndDate < Convert.ToDateTime("2014-10-01"))
                {
                    Label_MissionStatus.Text = "時間設定錯誤";
                }

                else if (Session["UpdateTime"].ToString() != SS.GetUpdateTime(Session["ProjectID"].ToString()))
                {
                    Response.Write("<script> alert('閒置時間過久，將返回大廳');</script>");
                    Server.Transfer("PMLobby.aspx");
                    Response.End();
                }

                else
                {
                    try
                    {
                        string IdentityCode = Session["ProjectID"].ToString() + Session["UserID"].ToString() + DateTime.Now.ToString("yyyyMMddHHmmssfff"); //任務辨識碼

                        #region 寫入任務

                        string MissionIn = "INSERT INTO PMMission  (ProjectID, MType, Creator, Title, Description, StartDate, EndDate, CreatDate, ResponsibleName, Complete, Alive, IdentityCode) VALUES ";
                        MissionIn += " ('" + Session["ProjectID"].ToString() + "', '" + DropDownList_MType.SelectedValue + "', N'" + Session["Name"].ToString() + "', N'" + TextBox_Title.Text.Replace("'", "''") + "', N'" + TextBox_MText.Text.Replace("'", "''") + "', '" + TextBox_DateStart.Text + "', '" + TextBox_DateEnd.Text + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "', N'" + DropDownList_Leader.SelectedItem + "', '0', '1', '" + IdentityCode + "');";
                                             
                        int ResultA = SS.SetSql(MissionIn);

                        #endregion

                        if (ResultA > 0)
                        {
                            string MissionCheck = "SELECT MissionID FROM PMMission WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (Alive = '1') AND (IdentityCode = '" + IdentityCode + "')";
                            string ResultB = SS.GetSqlString(MissionCheck, "MissionID");

                            if (ResultB != string.Empty)
                            {
                                //#region 寫入成員

                                string MMember = "INSERT INTO PMMMember (ProjectID, MissionID, UserID, ExecutorName, JoinDate, Responsible, Complete, Alive) VALUES ";
                                MMember += " ('" + Session["ProjectID"].ToString() + "','" + ResultB.Substring(0, ResultB.Length - 1) + "','" + DropDownList_Leader.SelectedValue + "',N'" + DropDownList_Leader.SelectedItem + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','1','0','1');";

                                for (int i = 0; i < CheckBoxList_Helper.Items.Count; i++)
                                {
                                    if (CheckBoxList_Helper.Items[i].Selected == true)
                                    {
                                        MMember += "INSERT INTO PMMMember (ProjectID, MissionID, UserID, ExecutorName, JoinDate, Responsible, Complete, Alive) VALUES ";
                                        MMember += " ('" + Session["ProjectID"].ToString() + "','" + ResultB.Substring(0, ResultB.Length - 1) + "','" + CheckBoxList_Helper.Items[i].Value + "',N'" + CheckBoxList_Helper.Items[i].Text + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','0','0','1');";
                                    }
                                }

                                #region 寫入附件到sql
                                string[] FileName = Session["FileName"] as string[];
                                int[] FileSize = Session["FileSize"] as int[];
                                string[] FileExtension = Session["FileExtension"] as string[];

                                for (int i = 0; i < FileName.Length; i++)
                                {
                                    MMember += "INSERT INTO PMMissionFile (ProjectID, MissionID, FileName, FileSize, FileExtension, CreatDate, Alive) ";
                                    MMember += "VALUES ('" + Session["ProjectID"].ToString() + "','" + ResultB.Substring(0, ResultB.Length - 1) + "',N'" + FileName[i] + "','" + FileSize[i] + "','" + FileExtension[i] + "','"+ DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")+"','1');";
                                }
                                #endregion

                                int ResultC = SS.SetSql(MMember);

                                if (ResultC > 0)
                                {
                                    Response.Write("<script> alert('任務創建成功');</script>");
                                    Reset();
                                }
                                else
                                {
                                    Label_MissionStatus.Text = "任務成員寫入失敗(3)";
                                }
                            }
                            else
                            {
                                Label_MissionStatus.Text = "任務創建失敗(2)";
                            }
                        }
                        else
                        {
                            Label_MissionStatus.Text = "任務創建失敗(1)";
                        }
                    }
                    catch
                    {
                        Response.Write("<script> alert('系統錯誤');</script>");
                        Server.Transfer("PMLobby.aspx");
                        Response.End();
                    }
                }
            }
            
        }

        protected void DropDownList_Leader_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckBoxList_Helper.ClearSelection(); //取消所有勾選
            GetHelperList();
        }

        private void GetHelperList() //檢查誰能成為協力者
        {
            try
            {
                string Tsql = "SELECT UserID, MemberName FROM PMPMember WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (Permission = '0') AND (Alive = '1') AND (UserID <> '" + DropDownList_Leader.SelectedValue + "') ORDER BY MemberName";

                DataSet Helper = SS.GetSqlTable(Tsql);

                CheckBoxList_Helper.DataSource = Helper;
                CheckBoxList_Helper.DataBind();

                Session["UpdateTime"] = SS.GetUpdateTime(Session["ProjectID"].ToString());
            }
            catch
            {
                Response.Write("<script> alert('系統錯誤');</script>");
                Server.Transfer("PMLobby.aspx");
                Response.End();
            }
        }
               
        protected void Button_Delete_Click(object sender, EventArgs e)
        {
            Reset();
        }

        private void Reset()
        {
            CheckBoxList_Helper.ClearSelection(); //取消所有勾選
            TextBox_Title.Text = "";
            TextBox_MText.Text = "";
            DropDownList_MType.SelectedIndex = 0;
            TextBox_DateStart.Text = DateTime.Now.ToString("yyyy-MM-dd");
            TextBox_DateEnd.Text = DateTime.Now.AddDays(3).ToString("yyyy-MM-dd");
            Session["FileName"] = new string[0];
            Session["FileSize"] = new int[0];
            Session["FileExtension"] = new string[0];
            Session["FileIndex"] = 0;
            HiddenField_FileCout.Value = "0";
            CreatList();
        }

        protected void Button_Back_Click(object sender, EventArgs e)
        {
            Response.Redirect("PMLobby.aspx");
        }

        string ConvertObjectToString(object obj)
        {
            return (obj == null) ? string.Empty : obj.ToString(); // FIXME
        }

        private void CreatList()
        {
            int FileCout = Convert.ToInt32(HiddenField_FileCout.Value);

            //創建檔案烈表
            DataTable dt = new DataTable();
            DataRow drow;
            for (int i = 0; i < FileCout; i++)
            {
                drow = dt.NewRow();
                dt.Rows.Add(drow);
            }

            GridView_FileList.DataSource = dt;
            GridView_FileList.DataBind();
        }

        protected void Button_UploadFile_Click(object sender, EventArgs e)
        {
            string path = Server.MapPath("~/UpdateFile/"); //存放位置
            string Time = DateTime.Now.ToString("yyyyMMddHHmmss");
            string[] FileName = Session["FileName"] as string[];
            int[] FileSize = Session["FileSize"] as int[];
            string[] FileExtension = Session["FileExtension"] as string[];

            Session["FileIndex"] = 0;

            if (!FileUpload_Creat.HasFile)
            {
                Response.Write("<script> alert('請先選擇檔案');</script>");
            }
            else if (FileUpload_Creat.PostedFile.ContentLength > 4194304)
            {
                Response.Write("<script> alert('上傳檔案超過4MB');</script>");
            }
            else
            {
                try
                {
                    int FileCout = Convert.ToInt32(HiddenField_FileCout.Value);

                    //傳輸檔案
                    FileUpload_Creat.PostedFile.SaveAs(path + Time + FileUpload_Creat.FileName.Replace(" ", ""));

                    Array.Resize(ref FileName, FileName.Length + 1);
                    Array.Resize(ref FileSize, FileSize.Length + 1);
                    Array.Resize(ref FileExtension, FileExtension.Length + 1);

                    FileName[FileCout] = Time + FileUpload_Creat.FileName.Replace(" ","");
                    FileSize[FileCout] = FileUpload_Creat.PostedFile.ContentLength;
                    FileExtension[FileCout] = System.IO.Path.GetExtension(FileUpload_Creat.FileName).ToLower();

                    Session["FileName"] = FileName;
                    Session["FileSize"] = FileSize;
                    Session["FileExtension"] = FileExtension;

                    FileCout++;
                    HiddenField_FileCout.Value = FileCout.ToString();

                    CreatList();                    
                }
                catch
                {
                    Response.Write("<script> alert('上傳失敗');</script>");
                }
            }
        }

        protected void GridView_FileList_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                try
                {
                    string[] FileName = Session["FileName"] as string[];
                    int FileIndex = Convert.ToInt32(Session["FileIndex"]);

                    Label Label_ShowUpdate = (Label)e.Row.Cells[0].FindControl("Label_ShowUpdate");
                    Button Button_FileDelete = (Button)e.Row.Cells[1].FindControl("Button_FileDelete");

                    Label_ShowUpdate.Text = "<A href=UpdateFile/" + FileName[FileIndex] + ">" + FileName[FileIndex].Substring(14, FileName[FileIndex].Length - 14) + "</a>";

                    Button_FileDelete.CommandArgument = FileIndex.ToString();
                    Button_FileDelete.OnClientClick = @"if (confirm('確定刪除此附件嗎') == false) return false;";

                    FileIndex++;
                    Session["FileIndex"] = FileIndex;                    
                }
                catch
                {
                    Response.Write("<script> alert('列表顯示失敗');</script>");
                }
            }
        }

        protected void GridView_FileList_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            try
            {
                int FileCout = Convert.ToInt32(HiddenField_FileCout.Value);
                int TempCount = 0;

                Session["FileIndex"] = 0;
                string[] FileName = Session["FileName"] as string[];
                int[] FileSize = Session["FileSize"] as int[];
                string[] FileExtension = Session["FileExtension"] as string[];

                string[] TempFileName = new string[FileCout];
                int[] TempFileSize = new int[FileCout];
                string[] TempFileExtension = new string[FileCout];

                for (int i = 0; i < FileCout; i++)
                {
                    if (i.ToString() != e.CommandArgument.ToString())
                    {
                        TempFileName[TempCount] = FileName[i];
                        TempFileSize[TempCount] = FileSize[i];
                        TempFileExtension[TempCount] = FileExtension[i];
                        TempCount++;
                    }
                }

                Session["FileName"] = TempFileName;
                Session["FileSize"] = TempFileSize;
                Session["FileExtension"] = TempFileExtension;

                FileCout--;
                HiddenField_FileCout.Value = FileCout.ToString();
                CreatList();
            }
            catch
            {
                Response.Write("<script> alert('檔案刪除失敗');</script>");
            }

        }
    }
}