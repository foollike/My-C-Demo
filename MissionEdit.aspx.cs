using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using Insus.NET;

namespace PMWeb
{
    public partial class MissionEdit : System.Web.UI.Page
    {
        Systemset SS = new Systemset();
        InsusJsUtility IJU = new InsusJsUtility();

        protected void Page_Load(object sender, EventArgs e)
        {           
            if (!Page.IsPostBack)
            {
                try
                {
                    Session["ProjectID"] = Request.QueryString["ProjectID"].ToString(); //取得專案編號
                    Session["UserID"] = Request.QueryString["UserID"].ToString(); //取得使用者ID
                    Session["Name"] = Request.QueryString["Name"].ToString(); //取得使用者姓名
                    Session["MissionID"] = Request.QueryString["MissionID"].ToString(); //取得任務編號

                    TextBox_EDateStart.Text = DateTime.Now.ToString("yyyy-MM-dd");
                    TextBox_EDateEnd.Text = DateTime.Now.AddDays(3).ToString("yyyy-MM-dd");
                                       
                    #region 初始化listbox

                    string initTsql = "SELECT ID, TypeName FROM PMType WHERE (Alive = '1') ORDER BY ID;";
                    initTsql += "SELECT UserID, MemberName FROM PMPMember WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (Alive = '1') AND (Permission = '0') ORDER BY MemberName;";
                    initTsql += "SELECT ProjectID, MissionID, UserID, ExecutorName, JoinDate, Responsible, Complete, CompleteDate FROM PMMMember WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (Alive = '1') AND (Responsible = '1');";
                    

                    DataSet ListBoxInit = SS.GetSqlTable(initTsql);
                                       
                    #region 任務類型
                    if (ListBoxInit.Tables[0].Rows.Count > 0)
                    {
                        DropDownList_EType.DataSource = ListBoxInit.Tables[0];
                        DropDownList_EType.DataBind();
                    }
                    else 
                    {
                        Response.Write("<script> alert('系統錯誤(1)');</script>");
                        Response.Write("<script>window.opener=null;self.close()</script>");
                    }
                    #endregion

                    #region 負責人
                    if (ListBoxInit.Tables[1].Rows.Count > 0)
                    {
                        DropDownList_ELeader.DataSource = ListBoxInit.Tables[1];
                        DropDownList_ELeader.DataBind();
                        DropDownList_ELeader.SelectedIndex = DropDownList_ELeader.Items.IndexOf(DropDownList_ELeader.Items.FindByValue(ListBoxInit.Tables[2].Rows[0]["UserID"].ToString()));
                    }
                    else
                    {
                        Response.Write("<script> alert('系統錯誤(2)');</script>");
                        Response.Write("<script>window.opener=null;self.close()</script>");
                    }                   
                    #endregion                                        

                    Reload();

                    #endregion
                }
                catch
                {
                    Response.Write("<script> alert('系統錯誤(3)');</script>");
                    Response.Write("<script>window.opener=null;self.close()</script>");
                }             
            }
        }
                
        private void Reload()
        {
            try
            {
                // 取得mission表資料
                string DTsql = "SELECT MissionID, ProjectID, MType, Creator, Title, Description, StartDate, EndDate, CreatDate, ResponsibleName, Complete, CompleteDate FROM PMMission WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (Alive = '1');";
                
                DataSet ds = SS.GetSqlTable(DTsql);                         

                #region 資料顯示

                TextBox_ETitle.Text = ds.Tables[0].Rows[0]["Title"].ToString();
                DropDownList_EType.SelectedIndex = DropDownList_EType.Items.IndexOf(DropDownList_EType.Items.FindByValue(ds.Tables[0].Rows[0]["MType"].ToString()));                                               
                DateTime StartDate = Convert.ToDateTime(ds.Tables[0].Rows[0]["StartDate"]);
                DateTime EndDate = Convert.ToDateTime(ds.Tables[0].Rows[0]["EndDate"]);
                TextBox_EDateStart.Text = StartDate.ToString("yyyy-MM-dd");
                TextBox_EDateEnd.Text = EndDate.ToString("yyyy-MM-dd");
                TextBox_EText.Text = ds.Tables[0].Rows[0]["Description"].ToString(); //顯示內文

                ReloadHelper();
                ReloadFileList();

                #endregion
            }
            catch
            {
                Response.Write("<script> alert('系統錯誤(4)');</script>");
                Response.Write("<script>window.opener=null;self.close()</script>");
            }
        }

        private void ReloadHelper()
        {
            try
            {
                Session["UpdateTime"] = SS.GetUpdateTime(Session["ProjectID"].ToString());

                string DTsql = "SELECT UserID, MemberName FROM PMPMember WHERE (Alive = '1') AND (UserID <> '" + DropDownList_ELeader.SelectedValue + "') AND (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (Permission = '0') ORDER BY MemberName;";
                DTsql += "SELECT ProjectID, MissionID, UserID, ExecutorName, JoinDate, Responsible, Complete, CompleteDate FROM PMMMember WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (Alive = '1') AND (Responsible = '0')";

                DataSet ds = SS.GetSqlTable(DTsql);

                CheckBoxList_EHelper.Items.Clear();

                if (ds.Tables[0].Rows.Count > 0)
                {
                    CheckBoxList_EHelper.DataSource = ds.Tables[0];
                    CheckBoxList_EHelper.DataBind();

                    for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                    {
                        if (CheckBoxList_EHelper.Items.FindByValue(ds.Tables[1].Rows[i]["UserID"].ToString()) != null)
                        {
                            CheckBoxList_EHelper.Items.FindByValue(ds.Tables[1].Rows[i]["UserID"].ToString()).Selected = true;
                        }
                    }
                }
            }
            catch
            {
                Response.Write("<script> alert('協力者讀取錯誤');</script>");
            }
        }

        protected void DropDownList_ELeader_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReloadHelper();
        }

        protected void Button_EClean_Click(object sender, EventArgs e)
        {
            TextBox_ETitle.Text = "";
            TextBox_EText.Text = "";

            for (int i = 0; i < CheckBoxList_EHelper.Items.Count; i++)
            {
                CheckBoxList_EHelper.Items[i].Selected = false;
            }
        }

        protected void Button_ESend_Click(object sender, EventArgs e)
        {
            if (TextBox_EDateStart.Text == string.Empty || TextBox_EDateEnd.Text == string.Empty)
            {
                Label_MissionStatus.Text = "日期欄不可為空";
            }
            else
            {
                DateTime StartDate = Convert.ToDateTime(TextBox_EDateStart.Text);
                DateTime EndDate = Convert.ToDateTime(TextBox_EDateEnd.Text);

                if (TextBox_ETitle.Text == string.Empty)
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
                    Response.Write("<script>window.opener.location.href = 'PMLobby.aspx';</script>");
                    Response.Write("<script>window.opener=null;self.close()</script>");                    
                }

                else
                {
                    try
                    {
                        //更新任務內容
                        string Tsql = "UPDATE PMMission ";
                        Tsql += "SET MType ='" + DropDownList_EType.SelectedValue + "', Creator =N'" + Session["Name"].ToString() + "', Title =N'" + TextBox_ETitle.Text.Replace("'", "''") + "', Description =N'" + TextBox_EText.Text.Replace("'", "''") + "', StartDate ='" + TextBox_EDateStart.Text + "', EndDate ='" + TextBox_EDateEnd.Text + "', ResponsibleName =N'" + DropDownList_ELeader.SelectedItem + "', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                        Tsql += "WHERE (Alive = '1') AND (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "');";

                        //刪除舊的成員
                        Tsql += "UPDATE PMMMember SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "');";

                        //增加領導人
                        Tsql += "INSERT INTO PMMMember (ProjectID, MissionID, UserID, ExecutorName, JoinDate, Responsible, Complete, Alive) VALUES ";
                        Tsql += " ('" + Session["ProjectID"].ToString() + "','" + Session["MissionID"].ToString() + "','" + DropDownList_ELeader.SelectedValue + "',N'" + DropDownList_ELeader.SelectedItem + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','1','0','1');";

                        //增加協力者
                        string Helper = "";
                        for (int i = 0; i < CheckBoxList_EHelper.Items.Count; i++)
                        {
                            if (CheckBoxList_EHelper.Items[i].Selected == true)
                            {
                                Helper += "INSERT INTO PMMMember (ProjectID, MissionID, UserID, ExecutorName, JoinDate, Responsible, Complete, Alive) VALUES ";
                                Helper += " ('" + Session["ProjectID"].ToString() + "','" + Session["MissionID"].ToString() + "','" + CheckBoxList_EHelper.Items[i].Value + "',N'" + CheckBoxList_EHelper.Items[i].Text + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','0','0','1');";
                            }
                        }
                        Tsql += Helper;

                        int Result = SS.SetSql(Tsql);

                        if (Result > 0)
                        {                            
                            Response.Write("<script> alert('修改成功');</script>");
                            Response.Write("<script>window.opener.location.reload();</script>");
                            Response.Write("<script>window.opener=null;self.close()</script>");    
                        }
                        else
                        {
                            Label_MissionStatus.Text = "任務修改失敗";
                        }
                    }
                    catch
                    {
                        Label_MissionStatus.Text = "任務修改失敗";
                    }
                }
            }            
        }

        private void ReloadFileList()
        {
            try
            {
                string Tsql = "SELECT  ID, FileName FROM PMMissionFile WHERE (Alive = '1') AND (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "')";
                DataSet ds = SS.GetSqlTable(Tsql);

                GridView_FileList.DataSource = ds;
                GridView_FileList.DataBind();
            }
            catch
            {
                Response.Write("<script> alert('附件列表讀取失敗');</script>");
                Response.Write("<script>window.opener=null;self.close()</script>");
            }            
        }

        protected void Button_UploadFile_Click(object sender, EventArgs e)
        {
            string path = Server.MapPath("~/UpdateFile/"); //存放位置
            string Time = DateTime.Now.ToString("yyyyMMddHHmmss");

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
                    //傳輸檔案
                    FileUpload_Creat.PostedFile.SaveAs(path + Time + FileUpload_Creat.FileName.Replace(" ", ""));

                    string Tsql = "INSERT INTO PMMissionFile (ProjectID, MissionID, FileName, FileSize, FileExtension, CreatDate, Alive) ";
                    Tsql += "VALUES ('" + Session["ProjectID"].ToString() + "','" + Session["MissionID"].ToString() + "',N'" + Time + FileUpload_Creat.FileName.Replace(" ", "") + "','" + FileUpload_Creat.PostedFile.ContentLength + "','" + System.IO.Path.GetExtension(FileUpload_Creat.FileName).ToLower() + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','1');";

                    int Result = SS.SetSql(Tsql);

                    if (Result > 0)
                    {
                        ReloadFileList();
                    }
                    else
                    {
                        ReloadFileList();
                        Response.Write("<script> alert('上傳失敗');</script>");
                    }
                }
                catch
                {
                    ReloadFileList();
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
                    DataRowView board_Row = (DataRowView)e.Row.DataItem;

                    Label Label_ShowUpdate = (Label)e.Row.Cells[0].FindControl("Label_ShowUpdate");
                    Button Button_FileDelete = (Button)e.Row.Cells[1].FindControl("Button_FileDelete");

                    Label_ShowUpdate.Text = "<A href=UpdateFile/" + board_Row["FileName"].ToString() + ">" + board_Row["FileName"].ToString().Substring(14, board_Row["FileName"].ToString().Length - 14) + "</a>";

                    Button_FileDelete.CommandArgument = board_Row["ID"].ToString();
                    Button_FileDelete.OnClientClick = @"if (confirm('確定刪除此附件嗎') == false) return false;";
                }
                catch
                {
                    Response.Write("<script> alert('附件列表讀取失敗');</script>");
                }
            }
        }

        protected void GridView_FileList_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            try
            {
                string Tsql = "UPDATE PMMissionFile SET Alive ='0', Updatetime ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (ID = '" + e.CommandArgument.ToString() + "')";
                int Result = SS.SetSql(Tsql);

                if (Result > 0)
                {
                    ReloadFileList();
                }
                else
                {
                    ReloadFileList();
                    Response.Write("<script> alert('檔案刪除失敗');</script>");
                }
            }
            catch
            {
                ReloadFileList();
                Response.Write("<script> alert('檔案刪除失敗');</script>");
            }
        }        

    }
}