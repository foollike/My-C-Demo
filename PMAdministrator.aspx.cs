using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

namespace PMWeb
{
    public partial class PMAdministrator : System.Web.UI.Page
    {
        Systemset SS = new Systemset();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                try
                {
                    Session["SPermission"] = Request.QueryString["SPermission"].ToString(); //檢查是否為系統管理員
                    Session["UserID"] = Request.QueryString["UserID"].ToString(); //取得使用者ID
                    Session["Name"] = Request.QueryString["Name"].ToString(); //取得使用者姓名

                    #region 系統管理員按鈕開啟

                    if (Session["SPermission"].ToString() == "1")
                    {
                        Label_PM.Visible = true;
                        DropDownList_AProjectList.Visible = true;
                        RadioButtonList_AMember.Visible = true;
                        Button_PMSend.Visible = true;
                        Button_CreatProject.Visible = true;

                        GetPList();
                    }

                    #endregion

                    ReloadPList();
                    ReloadTypeList();
                }
                catch
                {
                    Response.Write("<script> alert('系統錯誤(1)');</script>");
                    Server.Transfer("PMLobby.aspx");
                    Response.End();
                }
            }
        }

        private void GetPList()
        {
            #region 顯示選擇專案管理員專案表
            try
            {
                string Tsql = "SELECT ProjectID, ProjectName FROM PMProject WHERE (Alive = '1')";

                DataSet ds = SS.GetSqlTable(Tsql);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    DropDownList_AProjectList.Enabled = true;
                    Button_PMSend.Visible = true;
                    DropDownList_AProjectList.DataSource = ds;
                    DropDownList_AProjectList.DataBind();
                    GetPMManagerList();
                }
                else
                {
                    ListItem items = new ListItem();
                    items.Text = "請先創建專案";
                    DropDownList_AProjectList.Items.Clear();
                    RadioButtonList_AMember.Items.Clear();
                    DropDownList_AProjectList.Items.Add(items);
                    DropDownList_AProjectList.Enabled = false;
                    Button_PMSend.Visible = false;
                }

            }
            catch
            {
                Response.Write("<script> alert('系統錯誤(2)');</script>");
            }
            #endregion
        }

        //檢查目前專案管理人是誰
        private void GetPMManagerList()
        {
            try
            {
                string Tsql = "SELECT UserID, MemberName FROM PMPMember WHERE (Alive = '1') AND (ProjectID = '" + DropDownList_AProjectList.SelectedValue + "') ORDER BY MemberName;";
                Tsql += "SELECT UserID FROM PMPMember WHERE (Alive = '1') AND (Permission = '1') AND (ProjectID = '" + DropDownList_AProjectList.SelectedValue + "')";

                DataSet ds = SS.GetSqlTable(Tsql);
                RadioButtonList_AMember.Items.Clear();

                #region 檢查專案是否有成員
                if (ds.Tables[0].Rows.Count > 0)
                {                    
                    RadioButtonList_AMember.DataSource = ds;
                    RadioButtonList_AMember.DataBind();
                    Button_PMSend.Enabled = true;
                }
                else
                {
                    Button_PMSend.Enabled = false;
                }
                #endregion

                if (ds.Tables[1].Rows.Count > 0)
                {
                    if (RadioButtonList_AMember.Items.FindByValue(ds.Tables[1].Rows[0]["UserID"].ToString()) != null)
                        RadioButtonList_AMember.Items.FindByValue(ds.Tables[1].Rows[0]["UserID"].ToString()).Selected = true;
                }
            }
            catch
            {
                Response.Write("<script> alert('系統錯誤(3)');</script>");
            }
        }

        protected void Button_Back_Click(object sender, EventArgs e)
        {
            Response.Redirect("PMLobby.aspx");
        }

        protected void Button_CreatProject_Click(object sender, EventArgs e)
        {
            if (HiddenField_PCreat.Value == "0")
            {
                HiddenField_PCreat.Value = "1";
                Label_P1.Visible = true;
                Label_P2.Visible = true;
                Label_P3.Visible = true;
                TextBox_PTitel.Visible = true;
                TextBox_PText.Visible = true;
                Button_PSend.Visible = true;
                Button_PClean.Visible = true;
                CheckBoxList_PMember.Visible = true;
                ShowMemberList("Creat", Session["UserID"].ToString(), "");
            }
            else
            {
                TextBox_PTitel.Text = "";
                TextBox_PText.Text = "";
                HiddenField_PCreat.Value = "0";
                Label_P1.Visible = false;
                Label_P2.Visible = false;
                Label_P3.Visible = false;
                TextBox_PTitel.Visible = false;
                TextBox_PText.Visible = false;
                Button_PSend.Visible = false;
                Button_PClean.Visible = false;
                CheckBoxList_PMember.Visible = false;
            }
        }

        protected void Button_PClean_Click(object sender, EventArgs e)
        {
            CheckBoxList_PMember.ClearSelection();
            TextBox_PTitel.Text = "";
            TextBox_PText.Text = "";
        }

        //創建專案按鈕
        protected void Button_PSend_Click(object sender, EventArgs e)
        {
            if (TextBox_PTitel.Text == string.Empty)
            {
                Label_PStatus.Text = "標題欄不可為空";
            }
            else if (TextBox_PText.Text == string.Empty)
            {
                Label_PStatus.Text = "描述欄不可為空";
            }
            else
            {
                try
                {
                    string IdentityCode = Session["UserID"].ToString() + DateTime.Now.ToString("yyyyMMddHHmmssfff"); //任務辨識碼

                    string Tsql = "INSERT INTO PMProject (ProjectName, Description, CreatDate, Alive,IdentityCode) VALUES ";
                    Tsql += "  (N'" + TextBox_PTitel.Text.Replace("'", "''") + "',N'" + TextBox_PText.Text.Replace("'", "''") + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','1','" + IdentityCode + "')";

                    int Result = SS.SetSql(Tsql);

                    if (Result > 0) //先寫入專案資訊到PMProject
                    {
                        string GetProjectID = "SELECT ProjectID FROM PMProject WHERE (Alive = '1') AND (IdentityCode = '" + IdentityCode + "')";

                        string ResultRead = SS.GetSqlString(GetProjectID, "ProjectID"); //讀取IdentityCode找出同一筆專案資料

                        if (ResultRead != string.Empty)
                        {
                            string InsertPMember = "";

                            for (int i = 0; i < CheckBoxList_PMember.Items.Count; i++) //確認有哪些使用者被加入專案
                            {
                                if (CheckBoxList_PMember.Items[i].Selected == true)
                                {
                                    InsertPMember += "INSERT INTO PMPMember (ProjectID, UserID, MemberName, CreatDate, Permission, Alive) VALUES ";
                                    InsertPMember += " ('" + ResultRead.Substring(0, ResultRead.Length - 1) + "','" + CheckBoxList_PMember.Items[i].Value + "',N'" + CheckBoxList_PMember.Items[i].Text + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','0','1');";
                                }
                            }

                            if (InsertPMember != string.Empty)
                            {
                                int MemberResult = SS.SetSql(InsertPMember); //寫入PMPMember表

                                if (MemberResult > 0)
                                {
                                    Response.Write("<script> alert('專案創建成功');</script>");
                                    Label_PStatus.Text = "";
                                    ReloadPList();
                                    GetPList();
                                    TextBox_PTitel.Text = "";
                                    TextBox_PText.Text = "";
                                    HiddenField_PCreat.Value = "0";
                                    Label_P1.Visible = false;
                                    Label_P2.Visible = false;
                                    Label_P3.Visible = false;
                                    TextBox_PTitel.Visible = false;
                                    TextBox_PText.Visible = false;
                                    Button_PSend.Visible = false;
                                    Button_PClean.Visible = false;
                                    CheckBoxList_PMember.Visible = false;
                                }
                                else
                                {
                                    Label_PStatus.Text = "專案成員加入失敗";
                                }
                            }
                            else //如沒有勾選任何人就直接略過加入使用者的部分
                            {
                                Response.Write("<script> alert('專案創建成功');</script>");
                                Label_PStatus.Text = "";
                                ReloadPList();
                                GetPList();
                                TextBox_PTitel.Text = "";
                                TextBox_PText.Text = "";
                                HiddenField_PCreat.Value = "0";
                                Label_P1.Visible = false;
                                Label_P2.Visible = false;
                                Label_P3.Visible = false;
                                TextBox_PTitel.Visible = false;
                                TextBox_PText.Visible = false;
                                Button_PSend.Visible = false;
                                Button_PClean.Visible = false;
                                CheckBoxList_PMember.Visible = false;
                            }
                        }
                        else
                        {
                            Label_PStatus.Text = "專案創建失敗";
                        }
                    }
                    else
                    {
                        Label_PStatus.Text = "專案創建失敗";
                    }

                }
                catch
                {
                    Label_PStatus.Text = "專案創建失敗";
                }
            }
        }

        //顯示專案列表
        private void ReloadPList()
        {
            string PList = "";

            try
            {
                if (Session["SPermission"].ToString() == "1")
                {
                    PList = "SELECT ProjectID, ProjectName, Description, CreatDate, UpdateDate FROM PMProject WHERE (Alive = '1') ORDER BY ProjectID";
                }
                else
                {
                    string MemberHave = "SELECT ProjectID FROM PMPMember WHERE (Alive = '1') AND (UserID = '" + Session["UserID"].ToString() + "') AND (Permission = '1')";
                    DataSet ds = SS.GetSqlTable(MemberHave);

                    string MemberHaveList = "AND (ProjectID = '" + ds.Tables[0].Rows[0]["ProjectID"].ToString() + "'";

                    for (int i = 1; i < ds.Tables[0].Rows.Count; i++)
                    {
                        MemberHaveList += " OR ProjectID = '" + ds.Tables[0].Rows[i]["ProjectID"].ToString() + "'";
                    }

                    MemberHaveList += " )";

                    PList = "SELECT ProjectID, ProjectName, Description, CreatDate, UpdateDate FROM PMProject WHERE (Alive = '1') " + MemberHaveList + " ORDER BY ProjectID";
                }

                DataSet PLds = SS.GetSqlTable(PList);
                GridView_PList.DataSource = PLds;
                GridView_PList.DataBind();

            }
            catch
            {
                Response.Write("<script> alert('系統錯誤(4)');</script>");
                Server.Transfer("PMLobby.aspx");
                Response.End();
            }
        }

        protected void GridView_PList_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                try
                {
                    DataRowView board_Row = (DataRowView)e.Row.DataItem;

                    Label Label_PTitle = (Label)e.Row.Cells[0].FindControl("Label_PTitle");
                    Label Label_PText = (Label)e.Row.Cells[1].FindControl("Label_PText");
                    Label Label_PDate = (Label)e.Row.Cells[2].FindControl("Label_PDate");
                    Label Label_PUpdate = (Label)e.Row.Cells[3].FindControl("Label_PUpdate");
                    Button Button_PEdit = (Button)e.Row.Cells[4].FindControl("Button_PEdit");
                    Button Button_PDelete = (Button)e.Row.Cells[5].FindControl("Button_PDelete");

                    Label_PTitle.Text = board_Row["ProjectName"].ToString();
                    Label_PText.Text = board_Row["Description"].ToString();
                    DateTime CreatDate = Convert.ToDateTime(board_Row["CreatDate"]);
                    Label_PDate.Text = CreatDate.ToString("yyyy-MM-dd HH:mm:ss");

                    #region 判斷是否開關刪除專案按鈕
                    if (Session["SPermission"].ToString() == "1")
                    {
                        Button_PDelete.Enabled = true;
                    }
                    #endregion

                    #region 判定是否已有UpdateDate
                    if (board_Row["UpdateDate"] is DBNull)
                    {
                        Label_PUpdate.Text = "";
                    }
                    else
                    {
                        DateTime PUpdate = Convert.ToDateTime(board_Row["UpdateDate"]);
                        Label_PUpdate.Text = PUpdate.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    #endregion

                    Button_PDelete.OnClientClick = @"if (confirm('確定刪除嗎') == false) return false;";

                    Button_PEdit.CommandArgument = board_Row["ProjectID"].ToString();
                    Button_PDelete.CommandArgument = board_Row["ProjectID"].ToString();
                }
                catch
                {
                    Response.Write("<script> alert('系統錯誤(5)');</script>");
                }
            }
        }

        protected void GridView_PList_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName == "Send_Edit")
            {
                try
                {
                    Session["EditIndex"] = e.CommandArgument.ToString();

                    if (CheckUpdate(0, Session["EditIndex"].ToString()))
                    {
                        string Tsql = "SELECT ProjectName, Description FROM PMProject WHERE (Alive = '1') AND (ProjectID = '" + Session["EditIndex"].ToString() + "')";
                        string[] Model = SS.GetSqlString(Tsql, "ProjectName,Description").Split(',');

                        Label_E1.Visible = true;
                        Label_E2.Visible = true;
                        Label_E3.Visible = true;
                        TextBox_ETitel.Text = Model[0];
                        TextBox_EText.Text = Model[1];

                        TextBox_ETitel.Visible = true;
                        TextBox_EText.Visible = true;
                        Button_ESend.Visible = true;
                        Button_EClean.Visible = true;
                        CheckBoxList_EMember.Visible = true;
                        ShowMemberList("Edit", Session["UserID"].ToString(), Session["EditIndex"].ToString());
                    }
                    else
                    {
                        Response.Write("<script> alert('系統錯誤(14)');</script>");
                        Server.Transfer("PMLobby.aspx");
                        Response.End();
                    }
                }
                catch
                {
                    Response.Write("<script> alert('系統錯誤(6)');</script>");
                }
            }
            else
            {
                try
                {
                    string Tsql = "UPDATE PMProject SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (ProjectID = '" + e.CommandArgument.ToString() + "') AND (Alive = '1');";
                    Tsql += "UPDATE PMPMember SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (ProjectID = '" + e.CommandArgument.ToString() + "');";
                    Tsql += "UPDATE PMMission SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (ProjectID = '" + e.CommandArgument.ToString() + "') AND (Alive = '1');";
                    Tsql += "UPDATE PMMMember SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (ProjectID = '" + e.CommandArgument.ToString() + "')";

                    int DeleteProject = SS.SetSql(Tsql);

                    if (DeleteProject > 0)
                    {
                        ReloadPList();
                        GetPList();
                        EditReset();
                        Response.Write("<script> alert('專案刪除成功');</script>");

                    }
                    else
                    {
                        ReloadPList();
                        GetPList();
                        EditReset();
                        Response.Write("<script> alert('專案刪除失敗');</script>");
                    }
                }
                catch
                {
                    Response.Write("<script> alert('系統錯誤(7)');</script>");
                }
            }
        }

        protected void Button_EClean_Click(object sender, EventArgs e)
        {
            TextBox_ETitel.Text = "";
            TextBox_EText.Text = "";
            CheckBoxList_EMember.ClearSelection();
        }
        private void EditReset()
        {
            Label_EStatus.Text = "";
            TextBox_ETitel.Text = "";
            TextBox_EText.Text = "";

            Label_E1.Visible = false;
            Label_E2.Visible = false;
            Label_E3.Visible = false;
            TextBox_ETitel.Visible = false;
            TextBox_EText.Visible = false;
            Button_ESend.Visible = false;
            Button_EClean.Visible = false;
            CheckBoxList_EMember.Visible = false;
        }

        protected void Button_ESend_Click(object sender, EventArgs e)
        {
            if (TextBox_ETitel.Text == string.Empty)
            {
                Label_EStatus.Text = "標題欄不可為空";
            }
            else if (TextBox_EText.Text == string.Empty)
            {
                Label_EStatus.Text = "描述欄不可為空";
            }
            else if (!CheckUpdate(1, Session["EditIndex"].ToString()))
            {
                Response.Write("<script> alert('閒置時間過久，請重新輸入');</script>");
                ReloadPList();
                EditReset();
                GetPList();
            }
            else
            {
                try
                {
                    string Tsql = "UPDATE PMProject SET ProjectName =N'" + TextBox_ETitel.Text.Replace("'", "''") + "', Description =N'" + TextBox_EText.Text.Replace("'", "''") + "', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (ProjectID = '" + Session["EditIndex"].ToString() + "');";

                    #region 檢查需要刪除任務

                    string GetDelTsql = ""; //刪除PMMMember列表人員
                    string GetProjectDelTsql = ""; //刪除PMPMember列表人員

                    bool NeedDel = false;
                    DataSet GetDelList = new DataSet();

                    for (int i = 0; i < CheckBoxList_EMember.Items.Count; i++)
                    {
                        if (i > 0)
                        {
                            GetProjectDelTsql += " OR UserID = '" + CheckBoxList_EMember.Items[i].Value + "'";
                        }
                        else
                        {
                            GetProjectDelTsql += "UPDATE PMPMember SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (Permission = '0') AND (ProjectID = '" + Session["EditIndex"].ToString() + "') AND (UserID = '" + CheckBoxList_EMember.Items[i].Value + "'";
                        }

                        if (CheckBoxList_EMember.Items[i].Selected == false)
                        {
                            if (NeedDel)
                            {
                                GetDelTsql += " OR UserID = '" + CheckBoxList_EMember.Items[i].Value + "'";
                            }
                            else
                            {
                                GetDelTsql += "SELECT UserID FROM PMMMember WHERE (ProjectID = '" + Session["EditIndex"].ToString() + "') AND (Alive = '1') AND (UserID = '" + CheckBoxList_EMember.Items[i].Value + "'";
                                NeedDel = true;
                            }
                        }
                    }

                    GetProjectDelTsql += ");";
                    Tsql += GetProjectDelTsql;

                    if (NeedDel)
                    {
                        GetDelTsql += " )";
                        GetDelList = SS.GetSqlTable(GetDelTsql);

                        for (int i = 0; i < GetDelList.Tables[0].Rows.Count; i++)
                        {
                            Tsql += "UPDATE PMMMember SET Alive = '0', UpdateDate = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (ProjectID = '" + Session["EditIndex"].ToString() + "') AND (UserID = '" + GetDelList.Tables[0].Rows[i]["UserID"].ToString() + "') AND (Alive = '1');";
                        }
                    }
                    #endregion

                    int Result = SS.SetSql(Tsql);

                    if (Result > 0)
                    {
                        string UpdatePMember = "";

                        for (int i = 0; i < CheckBoxList_EMember.Items.Count; i++)
                        {
                            if (CheckBoxList_EMember.Items[i].Selected == true)
                            {
                                UpdatePMember += "INSERT INTO PMPMember (ProjectID, UserID, MemberName, CreatDate, Permission, Alive) VALUES ";
                                UpdatePMember += " ('" + Session["EditIndex"].ToString() + "','" + CheckBoxList_EMember.Items[i].Value + "',N'" + CheckBoxList_EMember.Items[i].Text + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','0','1');";
                            }
                        }

                        if (UpdatePMember != string.Empty)
                        {
                            int UpdateResult = SS.SetSql(UpdatePMember);

                            if (UpdateResult > 0)
                            {
                                Response.Write("<script> alert('專案修改成功');</script>");
                                ReloadPList();
                                EditReset();
                                GetPList();
                            }
                            else
                            {
                                Label_EStatus.Text = "專案成員修改失敗";
                            }
                        }
                        else
                        {
                            Response.Write("<script> alert('專案修改成功');</script>");
                            ReloadPList();
                            EditReset();
                            GetPList();
                        }

                    }
                    else
                    {
                        Label_EStatus.Text = "專案修改失敗";
                    }

                }
                catch
                {
                    Label_EStatus.Text = "專案修改失敗";
                }
            }
        }

        private void ShowMemberList(string Type, string UserID, string ProjectID)
        {
            try
            {
                string Tsql = "";

                if (ProjectID != string.Empty)
                {
                    string GetMissionLeaderTsql = "SELECT UserID FROM PMMMember WHERE (Alive = '1') AND (ProjectID = '" + ProjectID + "') AND (Responsible = '1')";
                    GetMissionLeaderTsql += "SELECT UserID FROM PMPMember WHERE (ProjectID = '" + ProjectID + "') AND (Alive = '1') AND (Permission = '1')";

                    DataSet GetMissionLeader = SS.GetSqlTable(GetMissionLeaderTsql);

                    Tsql = "SELECT UserID, Name FROM PMAccount WHERE (Permission = '0') AND (UserID <> '" + UserID + "'";

                    for (int i = 0; i < GetMissionLeader.Tables[0].Rows.Count; i++)
                    {
                        if (UserID != GetMissionLeader.Tables[0].Rows[i]["UserID"].ToString())
                        {
                            Tsql += " AND UserID <> '" + GetMissionLeader.Tables[0].Rows[i]["UserID"].ToString() + "'";
                        }
                    }

                    for (int i = 0; i < GetMissionLeader.Tables[1].Rows.Count; i++)
                    {
                        if (UserID != GetMissionLeader.Tables[1].Rows[i]["UserID"].ToString())
                        {
                            Tsql += " AND UserID <> '" + GetMissionLeader.Tables[1].Rows[i]["UserID"].ToString() + "'";
                        }
                    }

                    Tsql += " ) ORDER BY Name;";
                }
                else
                {
                    Tsql = "SELECT UserID, Name FROM PMAccount WHERE (Permission = '0') AND (UserID <> '" + UserID + "') ORDER BY Name;";
                }

                DataSet ds = new DataSet();

                if (Type == "Creat") //創建專案
                {
                    ds = SS.GetSqlTable(Tsql);

                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        CheckBoxList_PMember.DataSource = ds;
                        CheckBoxList_PMember.DataBind();
                    }
                    else
                    {
                        Label_PStatus.Text = "目前無會員可加入專案";
                    }

                }
                else //編集專案
                {
                    Tsql += "SELECT UserID, MemberName FROM PMPMember WHERE (Alive = '1') AND (ProjectID = '" + Session["EditIndex"].ToString() + "') AND (Permission = '0') ORDER BY MemberName";
                    ds = SS.GetSqlTable(Tsql);

                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        Button_ESend.Enabled = true;
                        Button_EClean.Enabled = true;
                        CheckBoxList_EMember.DataSource = ds.Tables[0];
                        CheckBoxList_EMember.DataBind();

                        for (int i = 0; i < ds.Tables[1].Rows.Count; i++) //將已加入該專案的會員打勾
                        {
                            if (CheckBoxList_EMember.Items.FindByValue(ds.Tables[1].Rows[i]["UserID"].ToString()) != null)
                            {
                                CheckBoxList_EMember.Items.FindByValue(ds.Tables[1].Rows[i]["UserID"].ToString()).Selected = true;
                            }
                        }
                    }
                    else
                    {
                        Button_ESend.Enabled = false;
                        Button_EClean.Enabled = false;
                        Label_EStatus.Text = "目前無會員可加入專案";
                    }
                }

            }
            catch
            {
                Response.Write("<script> alert('系統錯誤(8)');</script>");
            }
        }

        protected void DropDownList_AProjectList_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetPMManagerList();
        }

        protected void Button_PMSend_Click(object sender, EventArgs e)
        {
            try
            {
                string TsqlA = "SELECT UserID FROM PMPMember WHERE (Alive = '1') AND (Permission = '1') AND (ProjectID = '" + DropDownList_AProjectList.SelectedValue + "')";
                string CheckLeader = SS.GetSqlString(TsqlA, "UserID");

                if (CheckLeader != string.Empty)
                {
                    if (CheckLeader.Substring(0, CheckLeader.Length - 1) != RadioButtonList_AMember.SelectedValue)
                    {
                        string TsqlB = "UPDATE PMPMember SET Permission ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (UserID = '" + CheckLeader.Substring(0, CheckLeader.Length - 1) + "') AND (ProjectID = '" + DropDownList_AProjectList.SelectedValue + "');";
                        TsqlB += "UPDATE PMPMember SET Permission ='1', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (UserID = '" + RadioButtonList_AMember.SelectedValue + "') AND (ProjectID = '" + DropDownList_AProjectList.SelectedValue + "')";

                        int EditLeader = SS.SetSql(TsqlB);

                        if (EditLeader > 0)
                        {
                            Response.Write("<script> alert('修改成功');</script>");
                            ReloadPList();
                            EditReset();
                        }
                        else
                        {
                            Response.Write("<script> alert('修改失敗');</script>");
                            GetPMManagerList();
                            ReloadPList();
                            EditReset();
                        }
                    }
                    else
                    {
                        Response.Write("<script> alert('選擇的成員已是專案領導人');</script>");
                    }
                }
                else
                {
                    string TsqlC = "UPDATE PMPMember SET Permission ='1', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (UserID = '" + RadioButtonList_AMember.SelectedValue + "') AND (ProjectID = '" + DropDownList_AProjectList.SelectedValue + "')";
                    int EditLeader = SS.SetSql(TsqlC);

                    if (EditLeader > 0)
                    {
                        Response.Write("<script> alert('修改成功');</script>");
                        ReloadPList();
                        EditReset();
                    }
                    else
                    {
                        Response.Write("<script> alert('修改失敗');</script>");
                        GetPMManagerList();
                        ReloadPList();
                        EditReset();
                    }
                }
            }
            catch
            {
                Response.Write("<script> alert('系統錯誤(9)');</script>");
            }

        }

        //專案類型烈表更新
        private void ReloadTypeList()
        {
            try
            {
                string Tsql = "SELECT ID, TypeName, SaveDate FROM PMType WHERE (Alive = '1')";
                DataSet ds = SS.GetSqlTable(Tsql);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    DropDownList_TypeList.Enabled = true;
                    DropDownList_TypeList.DataSource = ds;
                    DropDownList_TypeList.DataBind();
                    Button_TypeEdit.Enabled = true;
                    Button_TypeDelete.Enabled = true;
                }
                else
                {
                    DropDownList_TypeList.Items.Clear();
                    DropDownList_TypeList.Enabled = false;
                    ListItem DDL = new ListItem();
                    DDL.Text = "請開始新增任務類型";
                    DropDownList_TypeList.Items.Add(DDL);
                    Button_TypeEdit.Enabled = false;
                    Button_TypeDelete.Enabled = false;
                }
            }
            catch
            {
                Response.Write("<script> alert('系統錯誤(10)');</script>");
            }
        }

        protected void Button_TAddSend_Click(object sender, EventArgs e)
        {
            if (TextBox_TAdd.Text == string.Empty)
            {
                Response.Write("<script> alert('內容不可為空白');</script>");
            }
            else
            {
                try
                {
                    string Tsql = "INSERT INTO PMType (TypeName, SaveDate, Alive) VALUES (N'" + TextBox_TAdd.Text.Replace("'", "''") + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','1')";
                    int AddType = SS.SetSql(Tsql);

                    if (AddType > 0)
                    {
                        Response.Write("<script> alert('新增成功');</script>");
                        TextBox_TAdd.Text = "";
                        ReloadTypeList();
                    }
                    else
                    {
                        Response.Write("<script> alert('新增失敗');</script>");
                        ReloadTypeList();
                    }
                }
                catch
                {
                    Response.Write("<script> alert('系統錯誤(11)');</script>");
                }
            }
        }

        protected void Button_TypeEdit_Click(object sender, EventArgs e)
        {
            if (HiddenField_TEdit.Value == "0")
            {
                HiddenField_TEdit.Value = "1";
                Label_T2.Visible = true;
                TextBox_TEdit.Visible = true;
                Button_TESend.Visible = true;
                Button_TypeEdit.Text = "收回編輯欄";
            }
            else
            {
                HiddenField_TEdit.Value = "0";
                TextBox_TEdit.Text = "";
                Label_T2.Visible = false;
                TextBox_TEdit.Visible = false;
                Button_TESend.Visible = false;
                Button_TypeEdit.Text = "修改";
            }
        }

        protected void Button_TypeDelete_Click(object sender, EventArgs e)
        {
            try
            {
                //刪除type種類
                string DelTsql = "UPDATE PMType SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (ID = '" + DropDownList_TypeList.SelectedValue + "');";

                #region 檢查需刪除哪些任務
                string GetMissionTsql = "SELECT MissionID FROM PMMission WHERE (Alive = '1') AND (MType = '" + DropDownList_TypeList.SelectedValue + "')";
                DataSet GetMission = SS.GetSqlTable(GetMissionTsql);

                if (GetMission.Tables[0].Rows.Count > 0)
                {
                    //刪除使用此type的任務
                    DelTsql += "UPDATE PMMission SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (MType = '" + DropDownList_TypeList.SelectedValue + "');";

                    string DelMissionList = "UPDATE PMMMember SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (MissionID = '" + GetMission.Tables[0].Rows[0]["MissionID"].ToString() + "' ";

                    for (int i = 1; i < GetMission.Tables[0].Rows.Count; i++)
                    {
                        DelMissionList += " OR MissionID = '" + GetMission.Tables[0].Rows[i]["MissionID"].ToString() + "'";
                    }
                    DelMissionList += " )";

                    DelTsql += DelMissionList;
                }

                #endregion

                int DelType = SS.SetSql(DelTsql);

                if (DelType > 0)
                {
                    Response.Write("<script> alert('刪除類型成功');</script>");
                    ReloadTypeList();
                }
                else
                {
                    Response.Write("<script> alert('刪除類型失敗');</script>");
                    ReloadTypeList();
                }
            }
            catch
            {
                Response.Write("<script> alert('系統錯誤(13)');</script>");
            }
        }

        protected void Button_TESend_Click(object sender, EventArgs e)
        {
            if (TextBox_TEdit.Text == string.Empty)
            {
                Response.Write("<script> alert('內容不可為空白');</script>");
            }
            else
            {
                try
                {
                    string Tsql = "UPDATE PMType SET TypeName =N'" + TextBox_TEdit.Text.Replace("'", "''") + "', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (ID = '" + DropDownList_TypeList.SelectedValue + "')";

                    int TypeEdit = SS.SetSql(Tsql);

                    if (TypeEdit > 0)
                    {
                        TextBox_TEdit.Text = "";
                        Response.Write("<script> alert('修改成功');</script>");
                        ReloadTypeList();
                    }
                    else
                    {
                        Response.Write("<script> alert('修改失敗');</script>");
                        ReloadTypeList();
                    }
                }
                catch
                {
                    Response.Write("<script> alert('系統錯誤(12)');</script>");
                }
            }
        }

        private bool CheckUpdate(int dataType, string ProjectID) //檢查更新時間
        {
            if (dataType == 0)
            {
                try
                {
                    string Tsql = "SELECT MissionID, UpdateDate FROM PMMission WHERE (Alive = '1') AND (ProjectID = '" + ProjectID + "') ORDER BY MissionID DESC";
                    DataSet Result = SS.GetSqlTable(Tsql);
                    Session["CheckUpdate"] = Result;

                    return true;
                }
                catch
                {
                    return false;
                }
            }
            else
            {
                try
                {
                    string Tsql = "SELECT MissionID, UpdateDate FROM PMMission WHERE (Alive = '1') AND (ProjectID = '" + ProjectID + "') ORDER BY MissionID DESC";
                    DataSet NowResult = SS.GetSqlTable(Tsql);
                    DataSet PreResult = (DataSet)Session["CheckUpdate"];

                    if (NowResult.Tables[0].Rows.Count == PreResult.Tables[0].Rows.Count)
                    {
                        if (NowResult.Tables[0].Rows.Count != 0)
                        {
                            if (NowResult.Tables[0].Rows[0]["MissionID"].ToString() == PreResult.Tables[0].Rows[0]["MissionID"].ToString())
                            {
                                bool Result = true;

                                for (int i = 0; i < NowResult.Tables[0].Rows.Count; i++)
                                {
                                    if (NowResult.Tables[0].Rows[i]["UpdateDate"].ToString() != PreResult.Tables[0].Rows[i]["UpdateDate"].ToString())
                                    {
                                        Result = false;
                                    }
                                }

                                return Result;
                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            return true;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                catch
                {
                    return false;
                }
            }
        }       

    }
}