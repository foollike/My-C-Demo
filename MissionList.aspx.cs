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
    public partial class MissionList : System.Web.UI.Page
    {
        Systemset SS = new Systemset();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                DropDownList_LType.SelectedIndex = 0;
                DropDownList_SortItem.SelectedIndex = 0;
                DropDownList_SortType.SelectedIndex = 0;
                RadioButtonList_LRule.SelectedIndex = 0;                
                TextBox_StartDate.Text = DateTime.Now.AddDays(-30).ToString("yyyy-MM-dd");
                TextBox_EndDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

                try
                {
                    Session["ProjectID"] = Request.QueryString["ProjectID"].ToString(); //取得專案編號
                    Session["ProjectName"] = Request.QueryString["ProjectName"].ToString(); //取得專案名字
                    Session["UserID"] = Request.QueryString["UserID"].ToString(); //取得玩家ID
                    Session["isAdministrator"] = Request.QueryString["isAdministrator"].ToString(); //檢查是否為管理員

                    string Tsql = "SELECT ID, TypeName FROM PMType WHERE (Alive = '1') ORDER BY ID";
                    DataSet ds = SS.GetSqlTable(Tsql);

                    DropDownList_LTypeB.Items.Clear();
                    ListItem DDL = new ListItem();
                    DDL.Text = "全部類型";
                    DDL.Value = "0";
                    DropDownList_LTypeB.Items.Add(DDL);

                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            DDL = new ListItem();
                            DDL.Text = ds.Tables[0].Rows[i]["TypeName"].ToString();
                            DDL.Value = ds.Tables[0].Rows[i]["ID"].ToString();
                            DropDownList_LTypeB.Items.Add(DDL);
                        }
                    }
                }
                catch
                {
                    Response.Write("<script> alert('系統錯誤(1)');</script>");
                    Server.Transfer("PMLobby.aspx");
                    Response.End();
                }                
            }

            try
            {
                this.Title = Session["ProjectName"].ToString() + " 任務列表";
            }
            catch
            {
                Response.Write("<script> alert('系統錯誤(2)');</script>");
                Server.Transfer("PMLobby.aspx");
                Response.End();
            }
        }

        protected void Button_Back_Click(object sender, EventArgs e)
        {
            Response.Redirect("PMLobby.aspx");
        }

        protected void Button_LInquire_Click(object sender, EventArgs e)
        {
            ShowMissionList();
        }

        private void ShowMissionList()
        {
            if (TextBox_StartDate.Text == string.Empty && RadioButtonList_LRule.SelectedIndex == 1)
            {
                Response.Write("<script> alert('請選擇開始日期');</script>");
            }

            else if (TextBox_EndDate.Text == string.Empty && RadioButtonList_LRule.SelectedIndex == 1)
            {
                Response.Write("<script> alert('請選擇結束日期');</script>");
            }

            else
            {
                DateTime StartDate = Convert.ToDateTime(TextBox_StartDate.Text);
                DateTime EndDate = Convert.ToDateTime(TextBox_EndDate.Text);
                bool HaveProject = false;

                if (StartDate > EndDate)
                {
                    Response.Write("<script> alert('截止時間不可比開始時間早');</script>");
                }
                else if (StartDate < Convert.ToDateTime("2014-10-01") || EndDate < Convert.ToDateTime("2014-10-01"))
                {
                    Response.Write("<script> alert('時間設定錯誤');</script>");
                }
                else
                {
                    try
                    {
                        #region 查詢條件
                        string LType = ""; //查詢種類(已完成?)
                        switch (DropDownList_LType.SelectedIndex)
                        {
                            case 0:
                                LType = "";
                                break;
                            case 1:
                                LType = " AND (Complete = '0')";
                                break;
                            case 2:
                                LType = " AND (Complete = '1')";
                                break;
                        }

                        string LTitle = ""; //查詢關鍵字                        
                        string LRule = ""; //限定日期

                        if (RadioButtonList_LRule.SelectedIndex == 1)
                        {
                            LRule = " AND (StartDate >= '" + TextBox_StartDate.Text + "') AND (StartDate <= '" + TextBox_EndDate.Text + "')";

                            if (TextBox_KeyWord.Text != string.Empty)
                            {
                                LTitle = " AND (Title Like '%" + TextBox_KeyWord.Text.Replace("'", "''") + "%')";
                            }
                        }

                        string LType2 = ""; //選擇任務類型

                        if (DropDownList_LTypeB.SelectedValue != "0")
                        {
                            LType2 = " AND (MType = '" + DropDownList_LTypeB.SelectedValue + "')";
                        }

                        string LSort = "ORDER BY ";

                        switch (DropDownList_SortItem.SelectedValue)
                        {
                            case "0":
                                LSort += "MType";
                                break;

                            case "1":
                                LSort += "Title";
                                break;

                            case "2":
                                LSort += "StartDate";
                                break;

                            case "3":
                                LSort += "EndDate";
                                break;

                            case "4":
                                LSort += "ResponsibleName";
                                break;                               
                        }

                        if (DropDownList_SortType.SelectedIndex == 1)
                        {
                            LSort += " DESC";
                        }
                        
                        #endregion

                        #region 非管理員只能看自己的專案
                        string UserMissionID = "";
                        if (Session["isAdministrator"].ToString() == "0")
                        {
                            string LTsql = "SELECT MissionID FROM PMMMember WHERE (ProjectID = '" + Session["ProjectID"].ToString() + "') AND (UserID = '" + Session["UserID"].ToString() + "') AND (Alive = '1') ORDER BY MissionID";
                            DataSet MMerberList = SS.GetSqlTable(LTsql);

                            if (MMerberList.Tables[0].Rows.Count > 0)
                            {
                                UserMissionID = "AND (MissionID = '" + MMerberList.Tables[0].Rows[0]["MissionID"].ToString() + "'";

                                for (int i = 1; i < MMerberList.Tables[0].Rows.Count; i++)
                                {
                                    UserMissionID += " OR MissionID = '" + MMerberList.Tables[0].Rows[i]["MissionID"].ToString() + "'";
                                }

                                UserMissionID += ")";
                                HaveProject = true;
                            }                            
                        }
                        else
                        {
                            HaveProject = true;
                        }
                        #endregion

                        if (HaveProject)
                        {
                            string Tsql = "SELECT MissionID, MType, Creator, Title, Description, StartDate, EndDate, CreatDate, ResponsibleName, Complete, CompleteDate FROM PMMission WHERE (Alive = '1') AND (ProjectID = '" + Session["ProjectID"].ToString() + "')" + LType + LTitle + LRule + LType2 + UserMissionID + LSort;

                            DataSet ds = SS.GetSqlTable(Tsql);

                            GridView_Dis.DataSource = ds;
                            GridView_Dis.DataBind();

                            if (ds.Tables[0].Rows.Count < 1)
                            {
                                Response.Write("<script> alert('查詢不到任何任務');</script>");
                            }
                        }
                        else
                        {
                            Response.Write("<script> alert('查詢不到任何任務');</script>");
                        }
                    }
                    catch
                    {
                        Response.Write("<script> alert('查詢失敗');</script>");
                    }
                }
            }
        }

        protected void GridView_Dis_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                try
                {
                    DataRowView board_Row = (DataRowView)e.Row.DataItem;
                    
                    Label Label_DType = (Label)e.Row.Cells[0].FindControl("Label_DType");
                    Label Label_DTitle = (Label)e.Row.Cells[1].FindControl("Label_DTitle");
                    Label Label_DStartDate = (Label)e.Row.Cells[2].FindControl("Label_DStartDate");
                    Label Label_DEndDate = (Label)e.Row.Cells[3].FindControl("Label_DEndDate");
                    Label Label_DDuration = (Label)e.Row.Cells[4].FindControl("Label_DDuration");
                    Label Label_DMember = (Label)e.Row.Cells[5].FindControl("Label_DMember");
                    HyperLink HyperLink_DetailLink = (HyperLink)e.Row.Cells[6].FindControl("HyperLink_DetailLink");

                    DateTime StartDate = Convert.ToDateTime(board_Row["StartDate"]);
                    DateTime EndDate = Convert.ToDateTime(board_Row["EndDate"]);                    

                    string[] TypeName = SS.GetMissionType();
                    Label_DType.Text = TypeName[Convert.ToInt32(board_Row["MType"])]; //顯示類型                         
                    Label_DTitle.Text = board_Row["Title"].ToString(); //顯示標題
                    Label_DStartDate.Text = StartDate.ToString("yyyy-MM-dd"); //顯示工作開始時間
                    Label_DEndDate.Text = EndDate.ToString("yyyy-MM-dd"); //顯示工作結束時間

                    string Deadline = "";

                    if (board_Row["Complete"].ToString() == "1")
                    {
                        Deadline = "已完成";
                        Label_DEndDate.ForeColor = System.Drawing.Color.Green;
                        Label_DDuration.ForeColor = System.Drawing.Color.Green;
                    }
                    else
                    {
                        if ((EndDate - DateTime.Now).Days < 1)
                        {
                            Deadline = "已截止";
                            Label_DEndDate.ForeColor = System.Drawing.Color.Red;
                            Label_DDuration.ForeColor = System.Drawing.Color.Red;
                        }
                        else
                        {
                            Deadline = (EndDate - DateTime.Now).Days.ToString() + "天";
                        }
                    }

                    Label_DDuration.Text = Deadline; //顯示工作期限天數                    
                    Label_DMember.Text = board_Row["ResponsibleName"].ToString(); //顯示負責人

                    HyperLink_DetailLink.NavigateUrl = "javascript:void window.open('MissionDetail.aspx?ProjectID=" + Session["ProjectID"].ToString() + "&MissionID=" + board_Row["MissionID"].ToString() + "&isAdministrator=" + Session["isAdministrator"].ToString() + "&UserID=" + Session["UserID"].ToString() + "&Name=" + Session["Name"].ToString() + "','NewWindow',config='height=780,width=530,left=800,top=200','menubar=no','status=no','scrollbars=yes');"; //詳細資料連結
                }
                catch
                {
                    Response.Write("<script> alert('系統錯誤(3)');</script>");
                }
            }
        }

        protected void RadioButtonList_LRule_SelectedIndexChanged(object sender, EventArgs e)
        {
            TextBox_KeyWord.Text = "";
            TextBox_StartDate.Text = DateTime.Now.AddDays(-30).ToString("yyyy-MM-dd");
            TextBox_EndDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            if (RadioButtonList_LRule.SelectedIndex == 0)
            {
                TextBox_StartDate.Enabled = false;
                TextBox_EndDate.Enabled = false;
                TextBox_KeyWord.Enabled = false;
            }            
            else
            {
                TextBox_StartDate.Enabled = true;
                TextBox_EndDate.Enabled = true;
                TextBox_KeyWord.Enabled = true;
            }
        }

        protected void GridView_Dis_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            GridView_Dis.PageIndex = e.NewPageIndex;
            ShowMissionList();
        }



    }
}