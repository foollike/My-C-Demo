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
    public partial class MissionDetail : System.Web.UI.Page
    {
        Systemset SS = new Systemset();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                try
                {
                    Session["ProjectID"] = Request.QueryString["ProjectID"].ToString(); //取得專案編號
                    Session["MissionID"] = Request.QueryString["MissionID"].ToString(); //取得任務編號
                    Session["UserID"] = Request.QueryString["UserID"].ToString(); //取得玩家ID
                    Session["Name"] = Request.QueryString["Name"].ToString(); //取得玩家姓名
                    Session["isAdministrator"] = Request.QueryString["isAdministrator"].ToString(); //檢查是否為管理員
                    Reload();

                }
                catch
                {
                    Response.Write("<script> alert('系統錯誤');</script>");
                    Response.Write("<script>window.opener.location.reload();</script>");
                    Response.Write("<script>window.opener=null;self.close()</script>");
                }
            }

            #region 顯示網頁名稱
            try
            {
                this.Title = Session["Title"].ToString();
            }
            catch
            {
                this.Title = "任務名稱 :";
            }
            #endregion

        }

        private void Reload()
        {
            try
            {
                // 取得mission表資料
                string DTsql = "SELECT MissionID, ProjectID, MType, Creator, Title, Description, StartDate, EndDate, CreatDate, ResponsibleName, Complete, CompleteDate FROM PMMission WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (Alive = '1');";

                // 取得mission協力者資料
                DTsql += "SELECT ProjectID, MissionID, UserID, ExecutorName, JoinDate, Responsible, Complete, CompleteDate FROM PMMMember WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (Alive = '1');";

                // 取得上傳檔案
                DTsql += "SELECT FileName FROM PMMissionFile WHERE (Alive = '1') AND (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "')";

                DataSet ds = SS.GetSqlTable(DTsql);

                Session["Title"] = "任務名稱 : " + ds.Tables[0].Rows[0]["Title"].ToString();
                string Helper = "";
                int HelperCount = Convert.ToInt32(HiddenField_HelperCount.Value);

                //檢查哪些人為協力者
                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    if (ds.Tables[1].Rows[i]["Responsible"].ToString() == "0")
                    {
                        if (HelperCount < 3)
                        {
                            Helper += ds.Tables[1].Rows[i]["ExecutorName"].ToString() + "&nbsp;&nbsp;&nbsp;&nbsp;";
                            HelperCount++;
                        }
                        else
                        {
                            Helper += ds.Tables[1].Rows[i]["ExecutorName"].ToString() + "<br />";
                            HelperCount = 0;
                        }
                        
                    }
                }

                #region 管理員按鈕開啟判定

                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (Convert.ToInt32(Session["isAdministrator"]) == 1)
                    {
                        Button_Edit.Visible = true;
                        Button_Delete.Visible = true;                        

                        if (ds.Tables[0].Rows[0]["Complete"].ToString() == "1")
                        {
                            Label_MissionStatus.Text = "任務已完成";
                            Label_MissionStatus.ForeColor = System.Drawing.Color.Green;
                            Button_StatusUpdate.Visible = false;
                        }
                        else
                        {
                            Label_MissionStatus.Text = "任務未完成";
                            Label_MissionStatus.ForeColor = System.Drawing.Color.Red;
                            Button_StatusUpdate.Visible = true;
                        }
                    }
                    else
                    {
                        if (ds.Tables[0].Rows[0]["Complete"].ToString() == "1")
                        {
                            Label_MissionStatus.Text = "任務已完成";
                            Label_MissionStatus.ForeColor = System.Drawing.Color.Green;
                            Button_Complete.Visible = false;
                        }
                        else
                        {
                            string MTsql = "SELECT Complete, CompleteDate FROM PMMMember WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (Alive = '1') AND (UserID = '" + Session["UserID"] + "')";
                            string[] IsComplete = SS.GetSqlString(MTsql, "Complete,CompleteDate").Split(',');

                            if (IsComplete[0] == "1")
                            {
                                Label_MissionStatus.Text = "你的部分完成囉 " + IsComplete[1];
                                Label_MissionStatus.ForeColor = System.Drawing.Color.Green;
                                Button_Complete.Visible = false;
                            }
                            else
                            {
                                Button_Complete.Visible = true;
                                Label_MissionStatus.Text = "任務未完成";
                                Label_MissionStatus.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                    }
                }
                else
                {
                    Response.Write("<script> alert('系統錯誤');</script>");
                    Response.Write("<script>window.opener.location.reload();</script>");
                    Response.Write("<script>window.opener=null;self.close()</script>");
                }


                #endregion

                #region 資料顯示

                string[] TypeName = SS.GetMissionType();
                Label_Title.Text = "主題  :  " + ds.Tables[0].Rows[0]["Title"].ToString();
                Label_MType.Text = "任務類型 : " + TypeName[Convert.ToInt32(ds.Tables[0].Rows[0]["MType"])];
                Label_Creater.Text = "任務創建者 : " + ds.Tables[0].Rows[0]["Creator"].ToString();
                Label_Responsible.Text = "負責人 : " + ds.Tables[0].Rows[0]["ResponsibleName"].ToString();

                if (Helper != string.Empty)
                {
                    Label_Helper.Text = " 協力者 : <br /><br />" + Helper;
                }
                else
                {
                    Label_Helper.Text = " 協力者 :  無";
                }                

                DateTime StartDate = Convert.ToDateTime(ds.Tables[0].Rows[0]["StartDate"]);
                DateTime EndDate = Convert.ToDateTime(ds.Tables[0].Rows[0]["EndDate"]);
                Label_Time.Text = "任務時間 : " + StartDate.ToString("yyyy-MM-dd") + " ~ " + EndDate.ToString("yyyy-MM-dd");

                TextBox_MText.Text = ds.Tables[0].Rows[0]["Description"].ToString(); //顯示內文

                #endregion

                #region 上傳檔案顯示
                GridView_FileList.DataSource = ds.Tables[2];
                GridView_FileList.DataBind();

                #endregion

                MessageReload();
            }
            catch
            {
                Response.Write("<script> alert('系統錯誤');</script>");
                Response.Write("<script>window.opener.location.reload();</script>");
                Response.Write("<script>window.opener=null;self.close()</script>");
            }
        }

        protected void Button_Complete_Click(object sender, EventArgs e)
        {
            try
            {
                //修改自己本身的完成狀態
                string SupComplete = "UPDATE PMMMember SET Complete = '1', CompleteDate = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                SupComplete += " WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (UserID = '" + Session["UserID"] + "') AND (Alive = '1')";

                int result = SS.SetSql(SupComplete);

                if (result > 0)
                {
                    StatusCheck();
                }
                else
                {
                    Response.Write("<script> alert('狀態提交失敗(S)');</script>");
                }

                Reload();

            }
            catch
            {
                Response.Write("<script> alert('系統錯誤');</script>");
            }
        }

        private void StatusCheck() //檢查是否所有人都完成，如是的話更改任務狀態
        {
            string IsMainComplete = "SELECT TOP (1) Complete FROM PMMMember WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (Complete = '0') AND (Alive = '1')";
            string isComplete = SS.GetSqlString(IsMainComplete, "Complete");

            if (isComplete == string.Empty)
            {
                string MainComplete = "UPDATE PMMission SET Complete ='1', CompleteDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (Alive = '1')";
                int MainResult = SS.SetSql(MainComplete);

                if (MainResult > 0)
                {
                    Response.Write("<script> alert('狀態提交成功(M)');</script>");
                    Response.Write("<script>window.opener.location.reload();</script>");
                }
                else
                {
                    Response.Write("<script> alert('狀態提交失敗(M)');</script>");
                }
            }
            else
            {
                Response.Write("<script> alert('狀態提交成功(S)');</script>");
                Response.Write("<script>window.opener.location.reload();</script>");
            }
        }

        protected void Button_StatusUpdate_Click(object sender, EventArgs e)
        {
            StatusCheck();
            Reload();
            Response.Write("<script>window.opener.location.reload();</script>");
        }

        protected void Button_Edit_Click(object sender, EventArgs e)
        {
            Response.Redirect("MissionEdit.aspx?ProjectID=" + Session["ProjectID"].ToString() + "&UserID=" + Session["UserID"].ToString() + "&Name=" + Session["Name"].ToString() + "&MissionID=" + Session["MissionID"].ToString());
        }

        protected void Button_Delete_Click(object sender, EventArgs e)
        {
            try
            {
                string GetMissionTsql = "SELECT MissionID, ProjectID, Alive FROM PMMission WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (Alive = '1')";
                string GetMission = SS.GetSqlString(GetMissionTsql, "MissionID,ProjectID,Alive");

                if (GetMission != string.Empty)
                {
                    string Tsql = "UPDATE PMMission SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (ProjectID = '" + Session["ProjectID"].ToString() + "');";
                    Tsql += "UPDATE PMMMember SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (Alive = '1');";
                    Tsql += "UPDATE PMMissionFile SET Alive ='0', Updatetime ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "')";

                    int Result = SS.SetSql(Tsql);

                    if (Result > 0)
                    {
                        Response.Write("<script> alert('任務刪除成功');</script>");
                        Response.Write("<script>window.opener.location.reload();</script>");
                        Response.Write("<script>window.opener=null;self.close()</script>");
                    }
                    else
                    {
                        Response.Write("<script> alert('任務刪除失敗');</script>");
                    }
                }
                else
                {
                    Response.Write("<script> alert('此任務已不存在');</script>");
                    Response.Write("<script>window.opener.location.reload();</script>");
                    Response.Write("<script>window.opener=null;self.close()</script>");
                }               
            }
            catch
            {
                Response.Write("<script> alert('系統錯誤');</script>");
                Response.Write("<script>window.opener.location.reload();</script>");
                Response.Write("<script>window.opener=null;self.close()</script>");
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

                    Label_ShowUpdate.Text = "<A href=UpdateFile/" + board_Row["FileName"].ToString() + ">" + board_Row["FileName"].ToString().Substring(14, board_Row["FileName"].ToString().Length - 14) + "</a>";                                        
                }
                catch
                {
                    Response.Write("<script> alert('列表顯示失敗');</script>");
                }
            }
        }

        private void MessageReload()
        {
            try
            {
                string Tsql = "SELECT ID, ProjectID, MissionID, Creater, Text, CreatDate FROM PMDiscussion WHERE (Alive = '1') AND (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') ORDER BY CreatDate DESC";
                DataSet ds = SS.GetSqlTable(Tsql);

                GridView_Message.DataSource = ds;
                GridView_Message.DataBind();
            }
            catch
            {
                Response.Write("<script> alert('留言讀取失敗');</script>");
            }
        }

        protected void Button_Send_Click(object sender, EventArgs e)
        {
            try
            {               
                if (TextBox_Message.Text == string.Empty)
                {
                    Label_Status.Text = "留言不可為空白";
                }                
                else
                {
                    string Tsql = "INSERT INTO PMDiscussion (ProjectID, MissionID, Creater, Text, CreatDate, Alive) ";
                    Tsql += "VALUES ('" + Session["ProjectID"].ToString() + "','" + Session["MissionID"].ToString() + "',N'" + Session["Name"].ToString() + "',N'" + TextBox_Message.Text.Replace("'", "''") + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','1')";

                    int Result = SS.SetSql(Tsql);

                    if (Result > 0)
                    {
                        Label_Status.Text = "";
                        TextBox_Message.Text = "";
                        MessageReload();
                    }
                    else
                    {
                        Label_Status.Text = "留言失敗";
                    }
                }
            }
            catch
            {
                Label_Status.Text = "留言失敗";
            }           

        }

        protected void GridView_Message_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            GridView_Message.PageIndex = e.NewPageIndex;
            MessageReload();
        }

        protected void GridView_Message_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                try
                {
                    DataRowView board_Row = (DataRowView)e.Row.DataItem;

                    Label Label_Chat = (Label)e.Row.Cells[0].FindControl("Label_Chat");
                    Button Button_ChatDelete = (Button)e.Row.Cells[1].FindControl("Button_ChatDelete");

                    //發言者
                    Label_Chat.Text = "<font color='#333399'>" + board_Row["Creater"].ToString() + "</font></a></b>" + "： <p> <font size='2'>";
                    //主發言
                    Label_Chat.Text += board_Row["Text"].ToString() + "<font size='2' color='#999999'>　" + Convert.ToDateTime(board_Row["CreatDate"].ToString()).ToString("yyyy-MM-dd HH:mm:ss") + "</font>";

                    Button_ChatDelete.CommandArgument = board_Row["ID"].ToString();
                    Button_ChatDelete.OnClientClick = @"if (confirm('確定刪除此留言嗎') == false) return false;";

                    if (Convert.ToInt32(Session["isAdministrator"]) == 1)
                    {
                        Button_ChatDelete.Visible = true;
                    }
                    else
                    {
                        Button_ChatDelete.Visible = false;
                    }
                }
                catch
                {
                    Response.Write("<script> alert('留言讀取失敗');</script>");
                }
            }
        }

        protected void GridView_Message_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            try
            {
                string Tsql = "UPDATE PMDiscussion SET Alive ='0', UpdateDate ='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE (Alive = '1') AND (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (MissionID = '" + Session["MissionID"].ToString() + "') AND (ID = '" + e.CommandArgument.ToString() + "')";
                int Result = SS.SetSql(Tsql);

                if (Result > 0)
                {
                    MessageReload();
                }
                else
                {
                    MessageReload();
                    Response.Write("<script> alert('留言刪除失敗');</script>");
                }
            }
            catch
            {
                MessageReload();
                Response.Write("<script> alert('留言刪除失敗');</script>");
            }
        }
    }
}