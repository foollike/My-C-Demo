using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.International.Converters.TraditionalChineseToSimplifiedConverter;
using System.Data.OleDb;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;


namespace AdminTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region link settings

        PuzzleWeb_TW.SQLLink TWPuzzle = new PuzzleWeb_TW.SQLLink();
        PuzzleWeb_HK.SQLLink HKPuzzle = new PuzzleWeb_HK.SQLLink();
        PuzzleWeb_SG.SQLLink SGPuzzle = new PuzzleWeb_SG.SQLLink();
        PuzzleWeb_MA.SQLLink MAPuzzle = new PuzzleWeb_MA.SQLLink();
        PuzzleWeb_TH.SQLLink THPuzzle = new PuzzleWeb_TH.SQLLink();
        PuzzleWeb_KR.SQLLink KRPuzzle = new PuzzleWeb_KR.SQLLink();

        Server_TW.Charger SaveTW = new Server_TW.Charger();
        Server_HK.Charger SaveHK = new Server_HK.Charger();
        Server_CN.Charger SaveCN = new Server_CN.Charger();
        Server_SG.Charger SaveSG = new Server_SG.Charger();
        Server_MY.Charger SaveMY = new Server_MY.Charger();
        Server_TH.Charger SaveTH = new Server_TH.Charger();
        Server_Debug.Charger SaveDebug = new Server_Debug.Charger();
        Server_KR.Charger SaveKR = new Server_KR.Charger();
        Server_145.Charger Save145 = new Server_145.Charger();

        SqlLink_TW.SQLLink TWsql = new SqlLink_TW.SQLLink();
        SqlLink_HK.SQLLink HKsql = new SqlLink_HK.SQLLink();
        SqlLink_CN.SQLLink CNsql = new SqlLink_CN.SQLLink();

        #endregion

        string ServerPassWord = "Macintosh@1984#"; //儲值伺服器密碼    
        string SqlPassWord = "SQL@#1955"; //Sql密碼
        string SendPassWord = "Mac@1984";//查詢密碼

        private void Form1_Load(object sender, EventArgs e)
        {
            //單項儲值初始化
            comboBox_Country.SelectedIndex = 0;
            comboBox_ItemType.SelectedIndex = 0;
            comboBox_PointType.SelectedIndex = 0;
            comboBox_Spitem.SelectedIndex = 0;

            //全體發送初始化
            comboBox_ATime.SelectedIndex = 1;
            comboBox_ACountry.SelectedIndex = 0;
            comboBox_PickServer.SelectedIndex = 0;
            comboBox_AItemType.SelectedIndex = 0;
            textBox_ALevelLimit.Text = "1";
            textBox_ACount.Text = "1";

            dateTimePicker_Start.Value = DateTime.Now;
            dateTimePicker_End.Value = DateTime.Now.AddDays(1);

            ATimeStart = dateTimePicker_Start.Value.ToString("yyyy/MM/dd");
            ATimeEnd = dateTimePicker_End.Value.ToString("yyyy/MM/dd");

            comboBox_SendListItem.SelectedIndex = 0;

            //虛寶綁定初始化
            comboBox_TItemType.SelectedIndex = 0;
            comboBox_TSP.SelectedIndex = 0;
            comboBox_TCountry.SelectedIndex = 0;
            dateTimePicker_TStart.Value = DateTime.Now;
            dateTimePicker_TEnd.Value = DateTime.Now.AddDays(1);
            textBox6.Text = "1";
            comboBox_PomoType.SelectedIndex = 0;
            comboBox_PomoCountry.SelectedIndex = 0;
            textBox_PomoNum.Text = "0";

            //分析紀錄查詢
            comboBox_AnalysisCountry.SelectedIndex = 0;
            dateTimePicker_Analysis.Value = DateTime.Now;
            comboBox_MoneyList.SelectedIndex = 0;

            //新增推播訊息初始化
            comboBox_GiftMGNum.SelectedIndex = 0;
            comboBox_GiftMGCountry.SelectedIndex = 0;
            comboBox_FBMGCountry.SelectedIndex = 0;
            comboBox_FBMGNum.SelectedIndex = 0;
            comboBox_FBWorC.SelectedIndex = 0;

            //玩家分析
            comboBox_RCountry.SelectedIndex = 0;
            dateTimePicker_Date.Value = DateTime.Now;
            comboBox_TypePoint.SelectedIndex = 0;

            //密碼查詢
            comboBox_PSeachMethod.SelectedIndex = 0;
            comboBox_PCountry.SelectedIndex = 0;

            //取消全體發送
            comboBox_CECountry.SelectedIndex = 0;
            comboBox_CEServer.SelectedIndex = 0;

            //月登入查詢
            comboBox_LCCountry.SelectedIndex = 0;
            dateTimePicker_LCStart.Value = DateTime.Now.AddDays(-30);
            dateTimePicker_LCEnd.Value = DateTime.Now;
            dateTimePicker_RStart.Value = DateTime.Now.AddDays(-60); ;
            dateTimePicker_REnd.Value = DateTime.Now.AddDays(-30); ;

            //儲值人數查詢
            comboBox_CHCountry.SelectedIndex = 0;
            comboBox_CHServer.SelectedIndex = 0;
            dateTimePicker_CHStart.Value = DateTime.Now.AddDays(-30);
            dateTimePicker_CHEnd.Value = DateTime.Now;

            //FTP解綁定
            comboBox_FTPCountry.SelectedIndex = 0;

            //創角留存率分析
            comboBox_CRCountry.SelectedIndex = 0;
            comboBox_CRServer.SelectedIndex = 0;
            dateTimePicker_CRCStart.Value = DateTime.Now.AddDays(-11);
            dateTimePicker_CRCEnd.Value = DateTime.Now.AddDays(-11);
            dateTimePicker_CRLStart.Value = DateTime.Now.AddDays(-10);
            dateTimePicker_CRLEnd.Value = DateTime.Now;

            //角色花費
            comboBox_CMCountry.SelectedIndex = 0;
            comboBox_CMServer.SelectedIndex = 0;
            dateTimePicker_CMCStart.Value = DateTime.Now.AddDays(-11);
            dateTimePicker_CMCEnd.Value = DateTime.Now.AddDays(-11);
            dateTimePicker_CMLStart.Value = DateTime.Now.AddDays(-10);
            dateTimePicker_CMLEnd.Value = DateTime.Now;
            CdsDataTable.Columns.Add("UserID", typeof(string));
            dssDataTable.Columns.Add("UserID", typeof(string));
            dssDataTable.Columns.Add("_Money", typeof(string));
            SumdssDataTable.Columns.Add("UserID", typeof(string));
            SumdssDataTable.Columns.Add("_Money", typeof(string));
            //報表產生
            comboBox_PDCountry.SelectedIndex = 0;
            dateTimePicker_PDStart.Value = DateTime.Now.AddDays(-1);
            dateTimePicker_PDEnd.Value = DateTime.Now;
        }



        public int ID = 0; //流水號
        int Country = 0; // 國家代碼
        int ItemType = 0; //物品種類
        int UserID = 0; //使用者帳號
        int ItemID = 0; //物品代碼
        int ItemCount = 0; //數量
        int ItemLV = 0; //武將等級
        int PushMsg = 0; //推播號碼
        int Wop = 0; //武器編號
        int SHP = 0;
        int SATK = 0;
        int SRHP = 0;
        bool PointType = false; //是否為實點
        bool SpItem = false; //是否為閃卡
        string CardSituation = ""; // 卡片狀態
        string PointSituation = ""; // 點數狀態
        string WopType = ""; // 顯示武器種類的名稱
        string CoinType = ""; // 顯示貨幣種類的名稱
        string SaveType = ""; //儲值類型名稱
        bool Permission = false;

        #region 角色花費宣告非同步變數
        int CMCountry = 0;
        int CMServer = 0;
        int days = 0;
        int CDays = 0;
        DateTime CMLStartTime;
        DateTime CMCStartTime;
        DataTable CdsDataTable = new DataTable();
        DataTable dssDataTable = new DataTable();
        DataTable SumdssDataTable = new DataTable();
        DateTime DateTimePickerCMCStart;
        DateTime DateTimePickerCMLStart;
        #endregion

        private void comboBox_Country_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox_SLock.Checked = false;

            switch (comboBox_Country.SelectedIndex)
            {
                case 0:

                    Country = 1;
                    textBox_UserID.Text = "";
                    break;

                case 1:

                    Country = 2;
                    textBox_UserID.Text = "";
                    break;

                case 2:

                    Country = 3;
                    textBox_UserID.Text = "";
                    break;

                case 3:

                    Country = 4;
                    textBox_UserID.Text = "";
                    break;

                case 4:

                    Country = 5;
                    textBox_UserID.Text = "";
                    break;

                case 5:

                    Country = 6;
                    textBox_UserID.Text = "";
                    break;

                case 6:

                    Country = 7;
                    textBox_UserID.Text = "";
                    break;

                case 7:

                    Country = 8;
                    textBox_UserID.Text = "";
                    break;

                case 8:

                    Country = 9;
                    textBox_UserID.Text = "";
                    break;
            }
        }

        private void button_Send_Click(object sender, EventArgs e)
        {

            string[] aa = textBox_UserID.Text.Replace("\n","@").Split('@');

            if (textBox_Spassword.Text != SendPassWord)
            {
                MessageBox.Show("密碼輸入錯誤");
                textBox_Spassword.Text = "";
            }
            else if (textBox_UserID.Text == string.Empty)
            {
                MessageBox.Show("請輸入玩家ID");
                textBox_Spassword.Text = "";
            }
            else if (textBox_ItemID.Text == string.Empty)
            {
                if (comboBox_ItemType.SelectedIndex == 1)
                    MessageBox.Show("請輸入產品包ID");
                else if (comboBox_ItemType.SelectedIndex == 2)
                    MessageBox.Show("請輸入武將ID");
                else if (comboBox_ItemType.SelectedIndex == 3)
                    MessageBox.Show("請輸入武器或防具ID");
                textBox_Spassword.Text = "";
            }
            else if (textBox_Count.Text == string.Empty)
            {
                MessageBox.Show("數量至少為1");
                textBox_Count.Text = "1";
                textBox_Spassword.Text = "";
            }
            else if (comboBox_ItemType.SelectedIndex == 2 && (textBox_SHP.Text == string.Empty || textBox_SATK.Text == string.Empty || textBox_SRHP.Text == string.Empty))
            {
                MessageBox.Show("覺醒值欄不可為空，至少為0");
            }
            else if ((comboBox_CoinType.SelectedIndex == 1 || comboBox_CoinType.SelectedIndex >= 7) && comboBox_ItemType.SelectedIndex == 0) //貴重物品確認
            {
                if (checkBox_SLock.Checked)
                {
                    Permission = true;
                }
                else
                {
                    if (checkBox_SLock.Visible)
                    {
                        checkBox_SLock.BackColor = Color.Red;
                    }
                    else
                    {
                        if (comboBox_CoinType.SelectedIndex == 1)
                        {
                            MessageBox.Show("發送項目 : " + PointSituation + "\r\n" + "數量 : " + ItemCount + "\r\n" + "發送前請開啟安全鎖", "發送確認");
                            checkBox_SLock.Visible = true;
                        }
                        else
                        {
                            MessageBox.Show("發送項目 : " + CoinType + "\r\n" + "數量 : " + ItemCount + "\r\n" + "發送前請開啟安全鎖", "發送確認");
                            checkBox_SLock.Visible = true;
                        }
                    }
                }
            }
            else
            {
                Permission = true;
            }                                  

            if (Permission)
            {
                try
                {
                    textBox_Spassword.Text = "";
                    Permission = false;
                    checkBox_SLock.Checked = false;
                    checkBox_SLock.Visible = false;                    
                    WopType = "";
                    Wop = 0;

                    if (textBox_PushMsg.Text == string.Empty)
                    {
                        textBox_PushMsg.Text = "0";
                    }

                    #region 武器選取

                    string wopstring = "";

                    if (checkBox_Wop8.Checked == true)
                    {
                        wopstring += "1";
                        WopType += "捲軸,";
                    }
                    else
                    {
                        wopstring += "0";
                    }

                    if (checkBox_Wop7.Checked == true)
                    {
                        wopstring += "1";
                        WopType += "弓箭,";
                    }
                    else
                    {
                        wopstring += "0";
                    }

                    if (checkBox_Wop6.Checked == true)
                    {
                        wopstring += "1";
                        WopType += "法杖,";
                    }
                    else
                    {
                        wopstring += "0";
                    }

                    if (checkBox_Wop5.Checked == true)
                    {
                        wopstring += "1";
                        WopType += "羽扇,";
                    }
                    else
                    {
                        wopstring += "0";
                    }

                    if (checkBox_Wop4.Checked == true)
                    {
                        wopstring += "1";
                        WopType += "雙槌,";
                    }
                    else
                    {
                        wopstring += "0";
                    }

                    if (checkBox_Wop3.Checked == true)
                    {
                        wopstring += "1";
                        WopType += "刀劍,";
                    }
                    else
                    {
                        wopstring += "0";
                    }

                    if (checkBox_Wop2.Checked == true)
                    {
                        wopstring += "1";
                        WopType += "巨劍,";
                    }
                    else
                    {
                        wopstring += "0";
                    }

                    if (checkBox_Wop1.Checked == true)
                    {
                        wopstring += "1";
                        WopType += "長柄,";
                    }
                    else
                    {
                        wopstring += "0";
                    }

                    if (wopstring != "00000000")
                    {
                        Wop = Convert.ToInt32(wopstring, 2);
                        Convert.ToString(Wop, 2);

                        if (wopstring == "11111111")
                        {
                            WopType = "擁有全武器";
                        }
                        else
                        {
                            WopType = WopType.Substring(0, WopType.Length - 1);
                        }
                    }
                    else
                    {
                        WopType = "無武器";
                        Wop = 0;
                    }

                    #endregion

                    switch (Country)
                    {
                        case 1: //台灣

                            if (ItemType == 1) //貨幣
                            {
                                foreach (string bb in aa)
                                {

                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  台灣 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "數量:  " + ItemCount + "\r\n" + "點數類型:  " + PointSituation + "\r\n" + "儲值狀態:  " + SaveTW.Recharge(ItemID, ItemCount, UserID, PointType, PushMsg, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";

                                }
                            }

                            if (ItemType == 2) //武將
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  台灣 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "武器種類:  " + WopType + "\r\n" + "卡片類型:  " + CardSituation + "\r\n" + "卡片等級:  " + ItemLV + "\r\n" + "儲值狀態:  " + SaveTW.RechargeItem(ItemID, 1, UserID, PushMsg, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            break;

                        case 2: //香港

                            if (ItemType == 1) //貨幣
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  香港 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "數量:  " + ItemCount + "\r\n" + "點數類型:  " + PointSituation + "\r\n" + "儲值狀態:  " + SaveHK.Recharge(ItemID, ItemCount, UserID, PointType, PushMsg, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            if (ItemType == 2) //武將
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  香港 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "武器種類:  " + WopType + "\r\n" + "卡片類型:  " + CardSituation + "\r\n" + "卡片等級:  " + ItemLV + "\r\n" + "儲值狀態:  " + SaveHK.RechargeItem(ItemID, 1, UserID, PushMsg, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            break;

                        case 3: //中國

                            if (ItemType == 1) //貨幣
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  中國 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "數量:  " + ItemCount + "\r\n" + "點數類型:  " + PointSituation + "\r\n" + "儲值狀態:  " + SaveCN.Recharge(ItemID, ItemCount, UserID, PointType, PushMsg, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            if (ItemType == 2) //武將
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  中國 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "武器種類:  " + WopType + "\r\n" + "卡片類型:  " + CardSituation + "\r\n" + "卡片等級:  " + ItemLV + "\r\n" + "儲值狀態:  " + SaveCN.RechargeItem(ItemID, 1, UserID, PushMsg, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            break;

                        case 4: //新加坡

                            if (ItemType == 1) //貨幣
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  新加坡 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "數量:  " + ItemCount + "\r\n" + "點數類型:  " + PointSituation + "\r\n" + "儲值狀態:  " + SaveSG.Recharge(ItemID, ItemCount, UserID, PointType, PushMsg, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            if (ItemType == 2) //武將
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  新加坡 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "武器種類:  " + WopType + "\r\n" + "卡片類型:  " + CardSituation + "\r\n" + "卡片等級:  " + ItemLV + "\r\n" + "儲值狀態:  " + SaveSG.RechargeItem(ItemID, 1, UserID, PushMsg, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            break;

                        case 5: //馬來西亞

                            if (ItemType == 1) //貨幣
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  馬來西亞 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "數量:  " + ItemCount + "\r\n" + "點數類型:  " + PointSituation + "\r\n" + "儲值狀態:  " + SaveMY.Recharge(ItemID, ItemCount, UserID, PointType, PushMsg, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            if (ItemType == 2) //武將
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  馬來西亞 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "武器種類:  " + WopType + "\r\n" + "卡片類型:  " + CardSituation + "\r\n" + "卡片等級:  " + ItemLV + "\r\n" + "儲值狀態:  " + SaveMY.RechargeItem(ItemID, 1, UserID, PushMsg, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            break;

                        case 6: //Debug

                            if (ItemType == 1) //貨幣
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  Debug " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "數量:  " + ItemCount + "\r\n" + "點數類型:  " + PointSituation + "\r\n" + "儲值狀態:  " + SaveDebug.Recharge(ItemID, ItemCount, UserID, PointType, PushMsg, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            if (ItemType == 2) //武將
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  Debug " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "武器種類:  " + WopType + "\r\n" + "卡片類型:  " + CardSituation + "\r\n" + "卡片等級:  " + ItemLV + "\r\n" + "儲值狀態:  " + SaveDebug.RechargeItem(ItemID, 1, UserID, PushMsg, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            break;

                        case 7: //泰國

                            if (ItemType == 1) //貨幣
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  泰國 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "數量:  " + ItemCount + "\r\n" + "點數類型:  " + PointSituation + "\r\n" + "儲值狀態:  " + SaveTH.Recharge(ItemID, ItemCount, UserID, PointType, PushMsg, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            if (ItemType == 2) //武將
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  泰國 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "武器種類:  " + WopType + "\r\n" + "卡片類型:  " + CardSituation + "\r\n" + "卡片等級:  " + ItemLV + "\r\n" + "儲值狀態:  " + SaveTH.RechargeItem(ItemID, 1, UserID, PushMsg, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            break;

                        case 8: //韓國

                            if (ItemType == 1) //貨幣
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  韓國 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "數量:  " + ItemCount + "\r\n" + "點數類型:  " + PointSituation + "\r\n" + "儲值狀態:  " + SaveKR.Recharge(ItemID, ItemCount, UserID, PointType, PushMsg, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            if (ItemType == 2) //武將
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  韓國 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "武器種類:  " + WopType + "\r\n" + "卡片類型:  " + CardSituation + "\r\n" + "卡片等級:  " + ItemLV + "\r\n" + "儲值狀態:  " + SaveKR.RechargeItem(ItemID, 1, UserID, PushMsg, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            break;

                        case 9: //145

                            if (ItemType == 1) //貨幣
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  38.145 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "數量:  " + ItemCount + "\r\n" + "點數類型:  " + PointSituation + "\r\n" + "儲值狀態:  " + Save145.Recharge(ItemID, ItemCount, UserID, PointType, PushMsg, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            if (ItemType == 2) //武將
                            {
                                foreach (string bb in aa)
                                {
                                    ID++;
                                    UserID = Convert.ToInt32(bb);
                                    textBox_Display.Text += "儲值單號:  " + ID + "\r\n" + "國別:  38.145 " + "\r\n" + "玩家ID:  " + UserID + "\r\n" + SaveType + ItemID + "\r\n" + "推播訊息號碼:  " + PushMsg + "\r\n" + "武器種類:  " + WopType + "\r\n" + "卡片類型:  " + CardSituation + "\r\n" + "卡片等級:  " + ItemLV + "\r\n" + "儲值狀態:  " + Save145.RechargeItem(ItemID, 1, UserID, PushMsg, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord) + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n" + "\r\n";
                                }
                            }

                            break;

                    }

                }
                catch
                {
                    MessageBox.Show("玩家ID錯誤");
                }
            }
        }

        private void textBox_UserID_TextChanged(object sender, EventArgs e)
        {
            checkBox_SLock.Checked = false;

            string[] aa = textBox_UserID.Text.Replace("\n", "@").Split('@');

            textBox_UserIdDis.Text = string.Empty;

            foreach (string bb in aa)
            {

                textBox_UserIdDis.Text += "玩家ID:  " + bb + "\r\n" + "\r\n";

            }
        }

        private void comboBox_ItemType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_ItemType.SelectedIndex)
            {
                case 0:
                    textBox_Count.Enabled = true;
                    comboBox_Spitem.Enabled = false;
                    comboBox_CoinType.Enabled = true;
                    comboBox_Spitem.Text = "";
                    textBox_ItemID.Enabled = false;
                    textBox_ItemID.Text = "0";
                    comboBox_CoinType.SelectedIndex = 0;
                    ItemType = 1;
                    ChangeLable.Text = "活動包ID:";
                    SaveType = "貨幣ID:  ";
                    textBox_SHP.Text = "0";
                    textBox_SATK.Text = "0";
                    textBox_SRHP.Text = "0";
                    textBox_SHP.Enabled = false;
                    textBox_SATK.Enabled = false;
                    textBox_SRHP.Enabled = false;
                    SHP = 0;
                    SATK = 0;
                    SRHP = 0;
                    WopInitialize();
                    textBox_Level.Text = "0";
                    textBox_Level.Enabled = false;

                    break;

                case 1:
                    textBox_Count.Enabled = false;
                    comboBox_PointType.Enabled = false;;
                    comboBox_Spitem.Enabled = false;
                    comboBox_CoinType.Enabled = false;
                    comboBox_Spitem.Text = "";
                    textBox_ItemID.Enabled = true;
                    textBox_Count.Text = "1";
                    ItemType = 1;
                    ChangeLable.Text = "活動包ID:";
                    SaveType = "產品包ID:  ";
                    textBox_SHP.Enabled = false;
                    textBox_SATK.Enabled = false;
                    textBox_SRHP.Enabled = false;
                    SHP = 0;
                    SATK = 0;
                    SRHP = 0;
                    WopInitialize();
                    textBox_Level.Text = "0";
                    textBox_Level.Enabled = false;
                    checkBox_SLock.Checked = false;
                    checkBox_SLock.Visible = false;

                    break;


                case 2:

                    textBox_Count.Enabled = false;
                    comboBox_PointType.Enabled = false;
                    comboBox_Spitem.Enabled = true;
                    comboBox_CoinType.Enabled = false;
                    textBox_ItemID.Enabled = true;
                    comboBox_PointType.Text = "";
                    textBox_Count.Text = "1";
                    ItemType = 2;
                    ChangeLable.Text = "武將ID";
                    SaveType = "武將ID:  ";
                    textBox_SHP.Enabled = true;
                    textBox_SATK.Enabled = true;
                    textBox_SRHP.Enabled = true;
                    comboBox_Spitem.SelectedIndex = 0;

                    checkBox_Wop1.Enabled = true;
                    checkBox_Wop2.Enabled = true;
                    checkBox_Wop3.Enabled = true;
                    checkBox_Wop4.Enabled = true;
                    checkBox_Wop5.Enabled = true;
                    checkBox_Wop6.Enabled = true;
                    checkBox_Wop7.Enabled = true;
                    checkBox_Wop8.Enabled = true;

                    textBox_Level.Text = "0";
                    textBox_Level.Enabled = true;
                    checkBox_SLock.Checked = false;
                    checkBox_SLock.Visible = false;

                    break;

                case 3:

                    textBox_Count.Enabled = false;
                    comboBox_PointType.Enabled = false;
                    comboBox_Spitem.Enabled = true;
                    comboBox_CoinType.Enabled = false;
                    textBox_ItemID.Enabled = true;
                    comboBox_PointType.Text = "";
                    textBox_Count.Text = "1";
                    ItemType = 2;
                    ChangeLable.Text = "武器防具ID:";
                    SaveType = "武器防具ID:  ";
                    textBox_SHP.Enabled = false;
                    textBox_SATK.Enabled = false;
                    textBox_SRHP.Enabled = false;
                    SHP = 0;
                    SATK = 0;
                    SRHP = 0;
                    WopInitialize();
                    comboBox_Spitem.SelectedIndex = 0;
                    textBox_Level.Text = "0";
                    textBox_Level.Enabled = true;
                    checkBox_SLock.Checked = false;
                    checkBox_SLock.Visible = false;

                    break;

            }
        }

        private void WopInitialize()
        {
            checkBox_Wop1.Enabled = false;
            checkBox_Wop2.Enabled = false;
            checkBox_Wop3.Enabled = false;
            checkBox_Wop4.Enabled = false;
            checkBox_Wop5.Enabled = false;
            checkBox_Wop6.Enabled = false;
            checkBox_Wop7.Enabled = false;
            checkBox_Wop8.Enabled = false;

            checkBox_Wop1.Checked = false;
            checkBox_Wop2.Checked = false;
            checkBox_Wop3.Checked = false;
            checkBox_Wop4.Checked = false;
            checkBox_Wop5.Checked = false;
            checkBox_Wop6.Checked = false;
            checkBox_Wop7.Checked = false;
            checkBox_Wop8.Checked = false;
        }

        private void comboBox_CoinType_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox_SLock.Checked = false;

            switch (comboBox_CoinType.SelectedIndex)
            {

                case 0:

                    ItemID = 2;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "金幣";

                    break;

                case 1:                                        

                    ItemID = 1;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = true;
                    CoinType = "鑽石";

                    break;

                case 2:

                    ItemID = 3;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "銀幣";

                    break;

                case 3:

                    ItemID = 4;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "金牌";

                    break;

                case 4:

                    ItemID = 5;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "聖旨";

                    break;

                case 5:

                    ItemID = 6;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "小體力卷";

                    break;

                case 6:

                    ItemID = 7;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "中體力卷";

                    break;

                case 7:

                    ItemID = 8;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "大體力卷";

                    break;

                case 8:

                    ItemID = 9;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "白寶石";

                    break;

                case 9:

                    ItemID = 10;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "紅寶石";

                    break;

                case 10:

                    ItemID = 11;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "藍寶石";

                    break;

                case 11:

                    ItemID = 12;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "黃寶石";

                    break;

                case 12:

                    ItemID = 13;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "紫寶石";

                    break;

                case 13:

                    ItemID = 14;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "統御值";

                    break;

                case 14:

                    ItemID = 17;
                    comboBox_PointType.SelectedIndex = 0;
                    comboBox_PointType.Enabled = false;
                    CoinType = "友情值";

                    break;

            }
        }

        private void textBox_ItemID_TextChanged(object sender, EventArgs e)
        {
            if (textBox_ItemID.Text != string.Empty)
            {
                try
                {
                    ItemID = Convert.ToInt32(textBox_ItemID.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字");
                    textBox_ItemID.Text = textBox_ItemID.Text.Substring(0, textBox_ItemID.Text.Length - 1);
                }
            }
        }

        private void textBox_PushMsg_TextChanged(object sender, EventArgs e)
        {
            checkBox_SLock.Checked = false;

            if (textBox_PushMsg.Text != string.Empty)
            {
                try
                {
                    PushMsg = Convert.ToInt32(textBox_PushMsg.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字");
                    textBox_PushMsg.Text = textBox_PushMsg.Text.Substring(0, textBox_PushMsg.Text.Length - 1);
                }
            }
        }

        private void textBox_Count_TextChanged(object sender, EventArgs e)
        {
            checkBox_SLock.Checked = false;

            if (textBox_Count.Text != string.Empty)
            {
                try
                {
                    ItemCount = Convert.ToInt32(textBox_Count.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_Count.Text = textBox_Count.Text.Substring(0, textBox_Count.Text.Length - 1);
                }
            }
        }

        private void comboBox_PointType_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox_SLock.Checked = false;

            if (comboBox_PointType.SelectedIndex == 0)
            {
                PointType = false;
                PointSituation = "虛點";
            }

            if (comboBox_PointType.SelectedIndex == 1)
            {
                PointType = true;
                PointSituation = "實點";
            }
        }

        private void comboBox_Spitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (comboBox_Spitem.SelectedIndex == 0)
                {
                    SpItem = false;
                    CardSituation = "普卡";
                }

                if (comboBox_Spitem.SelectedIndex == 1)
                {
                    SpItem = true;
                    CardSituation = "閃卡";
                }
            }
            catch
            {

            }
        }

        private void button_Clean_Click(object sender, EventArgs e)
        {
            checkBox_SLock.Checked = false;
            checkBox_SLock.Visible = false;

            textBox_Spassword.Text = "";

            if (textBox_UserID.Enabled == true)
                textBox_UserID.Text = "";

            if (textBox_ItemID.Enabled == true)
                textBox_ItemID.Text = "";

            textBox_PushMsg.Text = "0";

            if (textBox_Count.Enabled == true)
                textBox_Count.Text = "";

            checkBox_Wop1.Checked = false;
            checkBox_Wop2.Checked = false;
            checkBox_Wop3.Checked = false;
            checkBox_Wop4.Checked = false;
            checkBox_Wop5.Checked = false;
            checkBox_Wop6.Checked = false;
            checkBox_Wop7.Checked = false;
            checkBox_Wop8.Checked = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox_Display.Text = "";
        }


        int ACountry = 0; //全體發送的國家
        int PickServer = 0; //伺服器選擇 
        string APushMsg = ""; //全體推播訊息
        int ALevelLimit = 1; //全體等級限制
        int AType = 0; //貨幣或裝備
        int AItemType = 0; //貨幣種類
        int ACount = 0;//全體發送數量
        int AItemNum = 0; //武將或武器編號
        int AItemLevel = 0; //武將或武器等級
        bool ASpItem = false; // 是否閃卡
        int AWop = 0; //武器種類
        int AHP = 0;
        int AATK = 0;
        int ARHP = 0;
        bool ATime = true; //是否開啟時間限制
        string AUser = "";//使用者
        string ATimeStart = "yyyy/MM/dd"; //開始時間
        string ATimeEnd = "yyyy/MM/dd"; //結束時間
        string ItemTypeString = "";// 物品種類名稱
        string CoinTypeString = ""; //貨幣種類名稱
        string CardTypeString = ""; //卡片種類名稱
        string WopTypeString = ""; //武器種類名稱
        string TimeLimitString = ""; //是否限制時間
        string ServerName = ""; //伺服器名稱
        int AID = 0; // 流水單號
        bool APermission = false; //驗證


        private void comboBox_ACountry_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox_ALock.Checked = false;

            switch (comboBox_ACountry.SelectedIndex)
            {
                case 0:

                    ACountry = 1;

                    break;

                case 1:

                    ACountry = 2;

                    break;

                case 2:

                    ACountry = 3;
                    
                    break;

                case 3:
                    
                    ACountry = 4;
                    
                    break;

                case 4:
                    
                    ACountry = 5;
                    
                    break;

                case 5:
                    
                    ACountry = 6;
                                        
                    break;

                case 6:

                    ACountry = 7;

                    break;

                case 7:

                    ACountry = 8;

                    break;

                case 8:

                    ACountry = 9;

                    break;
            }
        }

        private void textBox_APushMeg_TextChanged(object sender, EventArgs e)
        {
            APushMsg = textBox_APushMeg.Text;
        }

        private void textBox_ALevelLimit_TextChanged(object sender, EventArgs e)
        {
            checkBox_ALock.Checked = false;

            if (textBox_ALevelLimit.Text != string.Empty)
            {
                try
                {
                    ALevelLimit = Convert.ToInt32(textBox_ALevelLimit.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_ALevelLimit.Text = textBox_ALevelLimit.Text.Substring(0, textBox_ALevelLimit.Text.Length - 1);
                }
            }
        }

        private void comboBox_AItemType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_AItemType.SelectedIndex)
            {
                case 0:

                    AType = 1;
                    AItemNum = 0;

                    label_AChange.Text = "不需填寫";
                    comboBox_ACoinType.Enabled = true;
                    comboBox_ACoinType.SelectedIndex = 0;
                    textBox_ACount.Enabled = true;
                    textBox_AItemNum.Enabled = false;
                    textBox_AItemLevel.Enabled = false;
                    comboBox_ACardType.Enabled = false;
                    comboBox_AWop.Enabled = false;
                    textBox_AItemNum.Text = "0";
                    textBox_AItemLevel.Text = "0";
                    comboBox_ACardType.Text = "1. 普卡";
                    comboBox_AWop.Text = "0. 無裝備武器";
                    ItemTypeString = "貨幣類";
                    textBox_Count.Text = "1";
                    textBox_AHP.Text = "0";
                    textBox_AATK.Text = "0";
                    textBox_ARHP.Text = "0";
                    textBox_AHP.Enabled = false;
                    textBox_AATK.Enabled = false;
                    textBox_ARHP.Enabled = false;
                    AHP = 0;
                    AATK = 0;
                    ARHP = 0;

                    break;

                case 1:

                    AType = 2;
                    AItemType = 9;
                    ACount = 1;

                    label_AChange.Text = "武將編號:";
                    comboBox_ACoinType.Enabled = false;
                    textBox_ACount.Enabled = false;
                    textBox_ACount.Text = "1";
                    textBox_AItemNum.Enabled = true;
                    textBox_AItemLevel.Enabled = true;
                    comboBox_ACardType.Enabled = true;
                    comboBox_AWop.Enabled = true;
                    ItemTypeString = "武將";
                    textBox_AHP.Enabled = true;
                    textBox_AATK.Enabled = true;
                    textBox_ARHP.Enabled = true;
                    checkBox_ALock.Checked = false;
                    checkBox_ALock.Visible = false;

                    break;


                case 2:

                    AType = 2;
                    AItemType = 10;
                    ACount = 1;

                    label_AChange.Text = "武器編號:";
                    comboBox_ACoinType.Enabled = false;
                    textBox_ACount.Enabled = false;
                    textBox_ACount.Text = "1";
                    textBox_AItemNum.Enabled = true;
                    textBox_AItemLevel.Enabled = true;
                    comboBox_ACardType.Enabled = true;
                    comboBox_AWop.Enabled = true;
                    ItemTypeString = "武器防具";
                    textBox_AHP.Text = "0";
                    textBox_AATK.Text = "0";
                    textBox_ARHP.Text = "0";
                    textBox_AHP.Enabled = false;
                    textBox_AATK.Enabled = false;
                    textBox_ARHP.Enabled = false;
                    AHP = 0;
                    AATK = 0;
                    ARHP = 0;
                    checkBox_ALock.Checked = false;
                    checkBox_ALock.Visible = false;

                    break;


            }
        }

        private void comboBox_ACoinType_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox_ALock.Checked = false;

            switch (comboBox_ACoinType.SelectedIndex)
            {

                case 0:

                    AItemType = 2;
                    CoinTypeString = "金幣";
                                        
                    break;

                case 1:

                    AItemType = 1;
                    CoinTypeString = "鑽石";

                    break;

                case 2:

                    AItemType = 3;
                    CoinTypeString = "銀幣";

                    break;

                case 3:

                    AItemType = 4;
                    CoinTypeString = "金牌";

                    break;


                case 4:

                    AItemType = 5;
                    CoinTypeString = "聖旨";

                    break;

                case 5:

                    AItemType = 6;
                    CoinTypeString = "小體力卷";

                    break;

                case 6:

                    AItemType = 7;
                    CoinTypeString = "中體力卷";

                    break;

                case 7:

                    AItemType = 8;
                    CoinTypeString = "大體力卷";

                    break;

                case 8:

                    AItemType = 17;
                    CoinTypeString = "名聲";

                    break;

                case 9:

                    AItemType = 19;
                    CoinTypeString = "個人碎片";

                    break;

                case 10:

                    AItemType = 18;
                    CoinTypeString = "統御力";

                    break;

                case 11:

                    AItemType = 11;
                    CoinTypeString = "白寶石";

                    break;

                case 12:

                    AItemType = 12;
                    CoinTypeString = "紅寶石";

                    break;

                case 13:

                    AItemType = 13;
                    CoinTypeString = "藍寶石";

                    break;

                case 14:

                    AItemType = 14;
                    CoinTypeString = "黃寶石";

                    break;

                case 15:

                    AItemType = 15;
                    CoinTypeString = "紫寶石";

                    break;

            }
        }

        private void textBox_ACount_TextChanged(object sender, EventArgs e)
        {
            checkBox_ALock.Checked = false;

            if (textBox_ACount.Text != string.Empty)
            {
                try
                {
                    ACount = Convert.ToInt32(textBox_ACount.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_ACount.Text = textBox_ACount.Text.Substring(0, textBox_ACount.Text.Length - 1);
                }
            }
        }

        private void textBox_AItemNum_TextChanged(object sender, EventArgs e)
        {
            if (textBox_AItemNum.Text != string.Empty)
            {
                try
                {
                    AItemNum = Convert.ToInt32(textBox_AItemNum.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_AItemNum.Text = textBox_AItemNum.Text.Substring(0, textBox_AItemNum.Text.Length - 1);
                }
            }
        }

        private void textBox_AItemLevel_TextChanged(object sender, EventArgs e)
        {

            if (textBox_AItemLevel.Text != string.Empty)
            {
                try
                {
                    AItemLevel = Convert.ToInt32(textBox_AItemLevel.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_AItemLevel.Text = textBox_AItemLevel.Text.Substring(0, textBox_AItemLevel.Text.Length - 1);
                }
            }
        }

        private void comboBox_ACardType_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (comboBox_ACardType.SelectedIndex == 0)
                {
                    ASpItem = false;
                    CardTypeString = "普卡";
                }

                if (comboBox_ACardType.SelectedIndex == 1)
                {
                    ASpItem = true;
                    CardTypeString = "閃卡";
                }
            }
            catch
            {

            }
        }

        private void comboBox_AWop_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_AWop.SelectedIndex)
            {
                case 0:

                    AWop = 0;
                    WopTypeString = "無裝備武器";

                    break;

                case 1:

                    AWop = 1;
                    WopTypeString = "長柄:(長槍、關刀、戟)";

                    break;

                case 2:

                    AWop = 2;
                    WopTypeString = "巨劍";

                    break;

                case 3:

                    AWop = 3;
                    WopTypeString = "刀劍";

                    break;

                case 4:

                    AWop = 4;
                    WopTypeString = "雙錘:(錘類、狼牙棒)";

                    break;

                case 5:

                    AWop = 5;
                    WopTypeString = "羽扇";

                    break;

                case 6:

                    AWop = 6;
                    WopTypeString = "法杖";

                    break;

                case 7:

                    AWop = 7;
                    WopTypeString = "弓箭";

                    break;

                case 8:

                    AWop = 8;
                    WopTypeString = "捲軸";

                    break;

            }
        }

        private void comboBox_ATime_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox_ALock.Checked = false;

            try
            {
                if (comboBox_ATime.SelectedIndex == 0)
                {
                    ATime = false;
                    TimeLimitString = "否";
                }

                if (comboBox_ATime.SelectedIndex == 1)
                {
                    ATime = true;
                    TimeLimitString = "是";
                }

            }
            catch
            {

            }
        }

        private void dateTimePicker_Start_ValueChanged(object sender, EventArgs e)
        {
            checkBox_ALock.Checked = false;
            ATimeStart = dateTimePicker_Start.Value.ToString("yyyy/MM/dd");
        }

        private void dateTimePicker_End_ValueChanged(object sender, EventArgs e)
        {
            checkBox_ALock.Checked = false;
            ATimeEnd = dateTimePicker_End.Value.ToString("yyyy/MM/dd");
        }

        private void button_ASend_Click(object sender, EventArgs e)
        {

            if (textBox_APassword.Text != SendPassWord)
            {
                MessageBox.Show("密碼輸入錯誤");
                textBox_APassword.Text = "";
            }
            else if (textBox_APushMeg.Text == string.Empty)
            {
                MessageBox.Show("請輸入推播文字");
                textBox_APassword.Text = "";
            }
            else if (textBox_ALevelLimit.Text == string.Empty)
            {
                MessageBox.Show("等級限制至少為1");
                textBox_ALevelLimit.Text = "1";
                textBox_APassword.Text = "";
            }
            else if (textBox_AUser.Text == string.Empty)
            {
                MessageBox.Show("請輸入儲值人員英文名以供紀錄");
                textBox_APassword.Text = "";
            }
            else if (textBox_ACount.Text == string.Empty)
            {
                MessageBox.Show("數量至少為1");
                textBox_ACount.Text = "1";
                textBox_APassword.Text = "";
            }
            else if (textBox_AItemNum.Text == string.Empty)
            {
                if (comboBox_AItemType.SelectedIndex == 1)
                    MessageBox.Show("請輸入武將編號");
                if (comboBox_AItemType.SelectedIndex == 2)
                    MessageBox.Show("請輸入武器編號");
                textBox_APassword.Text = "";
            }
            else if (textBox_AItemLevel.Text == string.Empty)
            {
                if (comboBox_AItemType.SelectedIndex == 1)
                    MessageBox.Show("武將等級至少為0");
                if (comboBox_AItemType.SelectedIndex == 2)
                    MessageBox.Show("武器等級至少為0");
                textBox_AItemLevel.Text = "0";
                textBox_APassword.Text = "";               
            }
            else if ((comboBox_ACoinType.SelectedIndex == 1 || comboBox_ACoinType.SelectedIndex >= 7) && comboBox_AItemType.SelectedIndex == 0) //貴重物品確認
            {
                if (checkBox_ALock.Checked)
                {
                    APermission = true;
                }
                else
                {
                    if (checkBox_ALock.Visible)
                    {
                        checkBox_ALock.BackColor = Color.Red;
                    }
                    else
                    {
                        MessageBox.Show("發送項目 : " + CoinTypeString + "\r\n" + "數量 : " + ACount + "\r\n" + "發送前請開啟安全鎖", "發送確認");
                        checkBox_ALock.Visible = true;                           
                    }
                }
            }
            else
            {
                APermission = true;
            }

            if (APermission)
            {

                try
                {
                    AID++;
                    APermission = false;
                    checkBox_ALock.Checked = false;
                    checkBox_ALock.Visible = false;                    

                    if (textBox_ALevelLimit.Text == string.Empty)
                    {
                        textBox_ALevelLimit.Text = "1";
                    }

                    try
                    {
                        if (ACountry == 3 || ACountry == 5)
                            APushMsg = ChineseConverter.Convert(textBox_APushMeg.Text, ChineseConversionDirection.TraditionalToSimplified);

                        else
                        {
                            APushMsg = textBox_APushMeg.Text;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("推播文字轉換錯誤");
                        APushMsg = textBox_APushMeg.Text;
                    }

                    switch (ACountry)
                    {
                        case 1: //台灣

                            if (AType == 1 && textBox_ACount.Text != string.Empty) //貨幣
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  台灣 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "貨幣種類:  " + CoinTypeString + "\r\n" + "數量:  " + ACount + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveTW.RechargeItemToAll(PickServer, APushMsg, AItemType, 0, 0, ACount, 0, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", false, 0, 0, 0, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            if (AType == 2 && textBox_ACount.Text != string.Empty) //武將or武器
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  台灣 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "武將編號:  " + AItemNum + "\r\n" + "武將等級:  " + AItemLevel + "\r\n" + "卡片種類:  " + CardTypeString + "\r\n" + "武器種類:  " + WopTypeString + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveTW.RechargeItemToAll(PickServer, APushMsg, AItemType, AItemNum, AWop, 1, AItemLevel, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", ASpItem, AHP, AATK, ARHP, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            break;

                        case 2: //香港

                            if (AType == 1 && textBox_ACount.Text != string.Empty) //貨幣
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  香港 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "貨幣種類:  " + CoinTypeString + "\r\n" + "數量:  " + ACount + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveHK.RechargeItemToAll(PickServer, APushMsg, AItemType, 0, 0, ACount, 0, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", false, 0, 0, 0, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            if (AType == 2 && textBox_ACount.Text != string.Empty) //武將or武器
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  香港 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "武將編號:  " + AItemNum + "\r\n" + "武將等級:  " + AItemLevel + "\r\n" + "卡片種類:  " + CardTypeString + "\r\n" + "武器種類:  " + WopTypeString + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveHK.RechargeItemToAll(PickServer, APushMsg, AItemType, AItemNum, AWop, 1, AItemLevel, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", ASpItem, AHP, AATK, ARHP, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            break;

                        case 3: //中國

                            if (AType == 1 && textBox_ACount.Text != string.Empty) //貨幣
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  中國 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "貨幣種類:  " + CoinTypeString + "\r\n" + "數量:  " + ACount + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveCN.RechargeItemToAll(PickServer, APushMsg, AItemType, 0, 0, ACount, 0, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", false, 0, 0, 0, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            if (AType == 2 && textBox_ACount.Text != string.Empty) //武將or武器
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  中國 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "武將編號:  " + AItemNum + "\r\n" + "武將等級:  " + AItemLevel + "\r\n" + "卡片種類:  " + CardTypeString + "\r\n" + "武器種類:  " + WopTypeString + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveCN.RechargeItemToAll(PickServer, APushMsg, AItemType, AItemNum, AWop, 1, AItemLevel, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", ASpItem, AHP, AATK, ARHP, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            break;

                        case 4: //新加坡

                            if (AType == 1 && textBox_ACount.Text != string.Empty) //貨幣
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  新加坡 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "貨幣種類:  " + CoinTypeString + "\r\n" + "數量:  " + ACount + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveSG.RechargeItemToAll(PickServer, APushMsg, AItemType, 0, 0, ACount, 0, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", false, 0, 0, 0, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            if (AType == 2 && textBox_ACount.Text != string.Empty) //武將or武器
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  新加坡 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "武將編號:  " + AItemNum + "\r\n" + "武將等級:  " + AItemLevel + "\r\n" + "卡片種類:  " + CardTypeString + "\r\n" + "武器種類:  " + WopTypeString + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveSG.RechargeItemToAll(PickServer, APushMsg, AItemType, AItemNum, AWop, 1, AItemLevel, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", ASpItem, AHP, AATK, ARHP, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            break;

                        case 5: //馬來西亞

                            if (AType == 1 && textBox_ACount.Text != string.Empty) //貨幣
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  馬來西亞 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "貨幣種類:  " + CoinTypeString + "\r\n" + "數量:  " + ACount + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveMY.RechargeItemToAll(PickServer, APushMsg, AItemType, 0, 0, ACount, 0, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", false, 0, 0, 0, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            if (AType == 2 && textBox_ACount.Text != string.Empty) //武將or武器
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  馬來西亞 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "武將編號:  " + AItemNum + "\r\n" + "武將等級:  " + AItemLevel + "\r\n" + "卡片種類:  " + CardTypeString + "\r\n" + "武器種類:  " + WopTypeString + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveMY.RechargeItemToAll(PickServer, APushMsg, AItemType, AItemNum, AWop, 1, AItemLevel, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", ASpItem, AHP, AATK, ARHP, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            break;

                        case 6: //Debug

                            if (AType == 1 && textBox_ACount.Text != string.Empty) //貨幣
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  Debug " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "貨幣種類:  " + CoinTypeString + "\r\n" + "數量:  " + ACount + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveDebug.RechargeItemToAll(PickServer, APushMsg, AItemType, 0, 0, ACount, 0, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", false, 0, 0, 0, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            if (AType == 2 && textBox_ACount.Text != string.Empty) //武將or武器
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  Debug " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "武將編號:  " + AItemNum + "\r\n" + "武將等級:  " + AItemLevel + "\r\n" + "卡片種類:  " + CardTypeString + "\r\n" + "武器種類:  " + WopTypeString + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveDebug.RechargeItemToAll(PickServer, APushMsg, AItemType, AItemNum, AWop, 1, AItemLevel, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", ASpItem, AHP, AATK, ARHP, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            break;

                        case 7: //泰國

                            if (AType == 1 && textBox_ACount.Text != string.Empty) //貨幣
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  泰國 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "貨幣種類:  " + CoinTypeString + "\r\n" + "數量:  " + ACount + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveTH.RechargeItemToAll(PickServer, APushMsg, AItemType, 0, 0, ACount, 0, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", false, 0, 0, 0, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            if (AType == 2 && textBox_ACount.Text != string.Empty) //武將or武器
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  泰國 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "武將編號:  " + AItemNum + "\r\n" + "武將等級:  " + AItemLevel + "\r\n" + "卡片種類:  " + CardTypeString + "\r\n" + "武器種類:  " + WopTypeString + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveTH.RechargeItemToAll(PickServer, APushMsg, AItemType, AItemNum, AWop, 1, AItemLevel, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", ASpItem, AHP, AATK, ARHP, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            break;

                        case 8: //韓國

                            if (AType == 1 && textBox_ACount.Text != string.Empty) //貨幣
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  韓國 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "貨幣種類:  " + CoinTypeString + "\r\n" + "數量:  " + ACount + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveKR.RechargeItemToAll(PickServer, APushMsg, AItemType, 0, 0, ACount, 0, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", false, 0, 0, 0, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            if (AType == 2 && textBox_ACount.Text != string.Empty) //武將or武器
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  韓國 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "武將編號:  " + AItemNum + "\r\n" + "武將等級:  " + AItemLevel + "\r\n" + "卡片種類:  " + CardTypeString + "\r\n" + "武器種類:  " + WopTypeString + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + SaveKR.RechargeItemToAll(PickServer, APushMsg, AItemType, AItemNum, AWop, 1, AItemLevel, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", ASpItem, AHP, AATK, ARHP, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            break;

                        case 9: //38.145

                            if (AType == 1 && textBox_ACount.Text != string.Empty) //貨幣
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  38.145 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "貨幣種類:  " + CoinTypeString + "\r\n" + "數量:  " + ACount + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + Save145.RechargeItemToAll(PickServer, APushMsg, AItemType, 0, 0, ACount, 0, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", false, 0, 0, 0, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            if (AType == 2 && textBox_ACount.Text != string.Empty) //武將or武器
                            {
                                textBox_ADisplay.Text += "儲值單號:  " + AID + "\r\n" + "國別:  38.145 " + "\r\n" + "伺服器:  " + ServerName + "\r\n" + "推播文字:  " + APushMsg + "\r\n" + "等級限制:  " + ALevelLimit + "\r\n" + "儲值類型:  " + ItemTypeString + "\r\n" + "武將編號:  " + AItemNum + "\r\n" + "武將等級:  " + AItemLevel + "\r\n" + "卡片種類:  " + CardTypeString + "\r\n" + "武器種類:  " + WopTypeString + "\r\n" + "是否限制創角時間:  " + TimeLimitString + "\r\n" + "開始時間:  " + ATimeStart + "\r\n" + "結束時間:  " + ATimeEnd + "\r\n" + "儲值狀態:  " + Save145.RechargeItemToAll(PickServer, APushMsg, AItemType, AItemNum, AWop, 1, AItemLevel, ATime, Convert.ToDateTime(ATimeStart), Convert.ToDateTime(ATimeEnd), ALevelLimit, AUser, "chart", ASpItem, AHP, AATK, ARHP, DateTime.Now.ToString("yyyyMMddHHmmss"), ServerPassWord) + "\r\n" + "儲值人員:  " + AUser + "\r\n" + "儲值時間:  " + DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss") + "\r\n" + "\r\n";
                                textBox_APassword.Text = "";
                            }

                            break;

                    }

                }
                catch
                {
                    MessageBox.Show("儲值失敗");

                }
            }

        }

        private void button_AOptionCancel_Click(object sender, EventArgs e)
        {
            checkBox_ALock.Checked = false;
            checkBox_ALock.Visible = false;

            textBox_APushMeg.Text = "";
            textBox_ALevelLimit.Text = "";

            if (textBox_ACount.Enabled == true)
                textBox_ACount.Text = "";

            if (textBox_AItemNum.Enabled == true)
                textBox_AItemNum.Text = "";

            if (textBox_AItemLevel.Enabled == true)
                textBox_AItemLevel.Text = "";

            textBox_AUser.Text = "";
            textBox_APassword.Text = "";

        }

        private void textBox_AUser_TextChanged(object sender, EventArgs e)
        {
            AUser = textBox_AUser.Text;
        }

        private void button_ADisCancel_Click(object sender, EventArgs e)
        {
            textBox_ADisplay.Text = "";
        }

        private void comboBox_PickServer_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox_ALock.Checked = false;

            switch (comboBox_PickServer.SelectedIndex)
            {

                case 0:

                    PickServer = 0;
                    ServerName = "全伺服器";

                    break;

                case 1:

                    PickServer = 1;
                    ServerName = "第一伺服器";

                    break;

                case 2:

                    PickServer = 2;
                    ServerName = "第二伺服器";

                    break;

                case 3:

                    PickServer = 3;
                    ServerName = "第三伺服器";

                    break;

                case 4:

                    PickServer = 4;
                    ServerName = "第四伺服器";

                    break;

                case 5:

                    PickServer = 5;
                    ServerName = "第五伺服器";

                    break;

                case 6:

                    PickServer = 6;
                    ServerName = "第六伺服器";

                    break;

                case 7:

                    PickServer = 7;
                    ServerName = "第七伺服器";

                    break;

                case 8:

                    PickServer = 8;
                    ServerName = "第八伺服器";

                    break;

                case 9:

                    PickServer = 9;
                    ServerName = "第九伺服器";

                    break;
            }
        }

        private void uTextBox1_TextChanged(object sender, EventArgs e)
        {
            checkBox_ALock.Checked = false;
        }

        private void button_SendListItem_SelectFile_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog file = new OpenFileDialog();

                file.ShowDialog();

                label_SendListItem_FilePath.Text = file.FileName;


                string[] lines = System.IO.File.ReadAllLines(file.FileName);
                DataTable DT = new DataTable("SendList");
                DT.Columns.Add("MAINID", typeof(string));
                //DT.Columns.Add("ITEMNUMBER", typeof(string));
                //DT.Columns.Add("ISSPITEM", typeof(string));
                DataRow workRow;

                foreach (string line in lines)
                {
                    string[] list = line.Split(',');

                    workRow = DT.NewRow();
                    workRow[0] = list[0];
                    //workRow[1] = list[1];
                    //workRow[2] = list[2];
                    DT.Rows.Add(workRow);
                }

                dataGridView1.DataSource = DT;
            }
            catch
            {
                MessageBox.Show("選取失敗");
            }
        }

        private void button_SendListItem_Send_Click(object sender, EventArgs e)
        {
            string[] lines = System.IO.File.ReadAllLines(label_SendListItem_FilePath.Text);
            DataTable DT = new DataTable("SendList");
            DT.Columns.Add("MAINID", typeof(string));
            //DT.Columns.Add("ITEMNUMBER", typeof(string));
            //DT.Columns.Add("ISSPITEM", typeof(string));
            DT.Columns.Add("ServerResp", typeof(string));
            DataRow workRow;

            foreach (string line in lines)
            {

                string[] list = line.Split(',');
                workRow = DT.NewRow();
                workRow[0] = list[0];
                //workRow[1] = list[1];
                //workRow[2] = list[2];

                if (comboBox_SendListItem.SelectedIndex == 0)
                {
                    Server_TW.Charger cp = new Server_TW.Charger();
                    //workRow[1] = cp.RechargeItem(565, 1, Convert.ToInt32(list[0]), 284, 0, true, ServerPassWord);
                }
                else if (comboBox_SendListItem.SelectedIndex == 1)
                {
                    Server_HK.Charger cp = new Server_HK.Charger();
                    //workRow[1] = cp.RechargeItem(565, 1, Convert.ToInt32(list[0]), 210, 0, true, ServerPassWord);
                }
                else if (comboBox_SendListItem.SelectedIndex == 2)
                {
                    Server_CN.Charger cp = new Server_CN.Charger();
                    //workRow[1] = cp.RechargeItem(565, 1, Convert.ToInt32(list[0]), 81, 0, true, ServerPassWord);
                }
                DT.Rows.Add(workRow);
            }

            dataGridView1.DataSource = DT;
            MessageBox.Show("發送完成");
        }

        private void TWopInitialize()
        {
            checkBox_TWop1.Enabled = false;
            checkBox_TWop2.Enabled = false;
            checkBox_TWop3.Enabled = false;
            checkBox_TWop4.Enabled = false;
            checkBox_TWop5.Enabled = false;
            checkBox_TWop6.Enabled = false;
            checkBox_TWop7.Enabled = false;
            checkBox_TWop8.Enabled = false;

            checkBox_TWop1.Checked = false;
            checkBox_TWop2.Checked = false;
            checkBox_TWop3.Checked = false;
            checkBox_TWop4.Checked = false;
            checkBox_TWop5.Checked = false;
            checkBox_TWop6.Checked = false;
            checkBox_TWop7.Checked = false;
            checkBox_TWop8.Checked = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox_TPassword.Text != SendPassWord)
            {
                MessageBox.Show("密碼欄錯誤");
            }

            else if (richTextBox1.Text == string.Empty)
            {
                MessageBox.Show("序號欄不可為空");
            }

            else if (textBox1.Text == string.Empty)
            {
                MessageBox.Show("推播訊息欄不可為空");
            }

            else if (textBox4.Text == "0" && (comboBox_TItemType.SelectedIndex == 2 || comboBox_TItemType.SelectedIndex == 5))
            {
                MessageBox.Show("武將ID不合法");
            }

            else if (textBox4.Text == "0" && (comboBox_TItemType.SelectedIndex == 3 || comboBox_TItemType.SelectedIndex == 6))
            {
                MessageBox.Show("武器ID不合法");
            }

            else if ((textBox_PHP.Text == string.Empty || textBox_PATK.Text == string.Empty || textBox_PRHP.Text == string.Empty) && (comboBox_TItemType.SelectedIndex == 2 || comboBox_TItemType.SelectedIndex == 5))
            {
                MessageBox.Show("覺醒欄不可為空，至少為0");
            }
            
            else if (textBox_TLevel.Text == string.Empty)
            {
                MessageBox.Show("武將等級不可為空，至少為0");
            }

            else
            {
                textBox_TPassword.Text = "";
                TWopType = "";
                TWOP = 0;
                                
                #region 武器選取

                string Twopstring = "";

                if (checkBox_TWop8.Checked == true)
                {
                    Twopstring += "1";
                    TWopType += "捲軸,";
                }
                else
                {
                    Twopstring += "0";
                }

                if (checkBox_TWop7.Checked == true)
                {
                    Twopstring += "1";
                    TWopType += "弓箭,";
                }
                else
                {
                    Twopstring += "0";
                }

                if (checkBox_TWop6.Checked == true)
                {
                    Twopstring += "1";
                    TWopType += "法杖,";
                }
                else
                {
                    Twopstring += "0";
                }

                if (checkBox_TWop5.Checked == true)
                {
                    Twopstring += "1";
                    TWopType += "羽扇,";
                }
                else
                {
                    Twopstring += "0";
                }

                if (checkBox_TWop4.Checked == true)
                {
                    Twopstring += "1";
                    TWopType += "雙槌,";
                }
                else
                {
                    Twopstring += "0";
                }

                if (checkBox_TWop3.Checked == true)
                {
                    Twopstring += "1";
                    TWopType += "刀劍,";
                }
                else
                {
                    Twopstring += "0";
                }

                if (checkBox_TWop2.Checked == true)
                {
                    Twopstring += "1";
                    TWopType += "巨劍,";
                }
                else
                {
                    Twopstring += "0";
                }

                if (checkBox_TWop1.Checked == true)
                {
                    Twopstring += "1";
                    TWopType += "長柄,";
                }
                else
                {
                    Twopstring += "0";
                }

                if (Twopstring != "00000000")
                {
                    TWOP = Convert.ToInt32(Twopstring, 2);
                    Convert.ToString(TWOP, 2);

                    if (Twopstring == "11111111")
                    {
                        TWopType = "擁有全武器";
                    }
                    else
                    {
                        TWopType = TWopType.Substring(0, TWopType.Length - 1);
                    }
                }
                else
                {
                    TWopType = "無武器";
                    TWOP = 0;
                }

                #endregion

                if (TCountry == 1)
                {
                    string[] Get_EventID = TWPomocard.Create_Event("Jway", textBox1.Text, TItemType, TItemID, Convert.ToInt32(textBox6.Text), dateTimePicker_TStart.Value, dateTimePicker_TEnd.Value).Split('|');

                    if (Get_EventID[0] == "Y")
                    {
                        string[] Codes = (richTextBox1.Text).Replace("\n", "@").Split('@');
                        string AA = "";
                        int error = 0;
                        foreach (string ac in Codes)
                        {
                            if (ac != string.Empty)
                            {
                                string getReq = TWPomocard.Create_Pomocards(ac, Get_EventID[1], TItemType, TSP, TWOP + "|" + textBox_PHP.Text + "|" + textBox_PATK.Text + "|" + textBox_PRHP.Text,TLevel, TItemID, Convert.ToInt32(textBox6.Text), dateTimePicker_TEnd.Value);
                                if (getReq != "Y|Y")
                                {
                                    AA += ac + ",";
                                }
                            }
                            else
                            {
                                error++;
                            }
                        }
                        textBox5.Text = AA;
                        MessageBox.Show("OK" + (Codes.Length - error).ToString());
                    }
                    else
                    {
                        MessageBox.Show("加入event錯誤");
                    }
                }

                else if (TCountry == 2)
                {
                    string[] Get_EventID = HKPomocard.Create_Event("Jway", textBox1.Text, TItemType, TItemID, Convert.ToInt32(textBox6.Text), dateTimePicker_TStart.Value, dateTimePicker_TEnd.Value).Split('|');

                    try
                    {

                        if (Get_EventID[0] == "Y")
                        {
                            string[] Codes = (richTextBox1.Text).Replace("\n", "@").Split('@');
                            string AA = "";
                            int error = 0;
                            foreach (string ac in Codes)
                            {
                                if (ac != string.Empty)
                                {
                                    string getReq = HKPomocard.Create_Pomocards(ac, Get_EventID[1], TItemType, TSP, TWOP + "|" + textBox_PHP.Text + "|" + textBox_PATK.Text + "|" + textBox_PRHP.Text, TLevel, TItemID, Convert.ToInt32(textBox6.Text), dateTimePicker_TEnd.Value);
                                    if (getReq != "Y|Y")
                                    {
                                        AA += ac + ",";
                                    }
                                }
                                else
                                {
                                    error++;
                                }
                            }
                            textBox5.Text = AA;
                            MessageBox.Show("OK" + (Codes.Length - error).ToString());
                        }
                        else
                        {
                            MessageBox.Show("加入event錯誤");
                        }

                    }

                    catch
                    {
                        MessageBox.Show("綁定失敗");
                    }
                }

                else if (TCountry == 3)
                {
                    string[] Get_EventID = CNPomocard.Create_Event("Jway", textBox1.Text, TItemType, TItemID, Convert.ToInt32(textBox6.Text), dateTimePicker_TStart.Value, dateTimePicker_TEnd.Value).Split('|');

                    if (Get_EventID[0] == "Y")
                    {
                        string[] Codes = (richTextBox1.Text).Replace("\n", "@").Split('@');
                        string AA = "";
                        int error = 0;
                        foreach (string ac in Codes)
                        {
                            if (ac != string.Empty)
                            {
                                string getReq = CNPomocard.Create_Pomocards(ac, Get_EventID[1], TItemType, TSP, TWOP + "|" + textBox_PHP.Text + "|" + textBox_PATK.Text + "|" + textBox_PRHP.Text, TItemID, Convert.ToInt32(textBox6.Text), dateTimePicker_TEnd.Value);
                                if (getReq != "Y|Y")
                                {
                                    AA += ac + ",";
                                }
                            }
                            else
                            {
                                error++;
                            }

                        }
                        textBox5.Text = AA;
                        MessageBox.Show("OK" + (Codes.Length - error).ToString());
                    }
                    else
                    {
                        MessageBox.Show("加入event錯誤");
                    }
                }

                else
                {
                    string[] Get_EventID = DebugPomocard.Create_Event("Jway", textBox1.Text, TItemType, TItemID, Convert.ToInt32(textBox6.Text), dateTimePicker_TStart.Value, dateTimePicker_TEnd.Value).Split('|');

                    try
                    {

                        if (Get_EventID[0] == "Y")
                        {
                            string[] Codes = (richTextBox1.Text).Replace("\n", "@").Split('@');
                            string AA = "";
                            int error = 0;
                            foreach (string ac in Codes)
                            {
                                if (ac != string.Empty)
                                {

                                    string getReq = DebugPomocard.Create_Pomocards(ac, Get_EventID[1], TItemType, TSP, TWOP + "|" + textBox_PHP.Text + "|" + textBox_PATK.Text + "|" + textBox_PRHP.Text, TLevel, TItemID, Convert.ToInt32(textBox6.Text), dateTimePicker_TEnd.Value);
                                    if (getReq != "Y|Y")
                                    {
                                        AA += ac + ",";
                                    }
                                }
                                else
                                {
                                    error++;
                                }
                            }
                            textBox5.Text = AA;
                            MessageBox.Show("OK" + (Codes.Length - error).ToString());
                        }
                        else
                        {
                            MessageBox.Show("加入event錯誤");
                        }

                    }

                    catch
                    {
                        MessageBox.Show("綁定失敗");
                    }
                }
            }
        }

        DebugPomoCard.PomoCard DebugPomocard = new DebugPomoCard.PomoCard();
        TWPomoCard.PomoCard TWPomocard = new TWPomoCard.PomoCard();
        HKPomoCard.PomoCard HKPomocard = new HKPomoCard.PomoCard();
        CNPomoCard.PomoCard CNPomocard = new CNPomoCard.PomoCard();

        string TItemType = ""; //物品種類
        string TSP = ""; //是否為閃卡
        int TWOP = 0; //武器種類
        string TWopType = ""; // 顯示武器種類的名稱
        int TCountry = 0; //國家編號
        string TItemID = ""; //物品編號
        int TLevel = 0; //武將等級

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_TItemType.SelectedIndex)
            {
                case 0:

                    TItemType = "1";
                    comboBox_TSP.SelectedIndex = 0;
                    comboBox_TSP.Enabled = false;
                    textBox4.Text = "0";
                    textBox4.Enabled = false;
                    comboBox_TCoinType.Enabled = true;
                    comboBox_TCoinType.SelectedIndex = 0;
                    TItemID = "1";
                    label_TChange.Text = "不需填寫 :";
                    textBox6.Enabled = true;
                    textBox_PHP.Text = "0";
                    textBox_PATK.Text = "0";
                    textBox_PRHP.Text = "0";
                    textBox_PHP.Enabled = false;
                    textBox_PATK.Enabled = false;
                    textBox_PRHP.Enabled = false;
                    textBox_TLevel.Text = "0";
                    textBox_TLevel.Enabled = false;
                    TWopInitialize();
                    
                    break;

                case 1:

                    TItemType = "1";
                    comboBox_TSP.SelectedIndex = 0;
                    comboBox_TSP.Enabled = false;
                    comboBox_TCoinType.Enabled = false;
                    comboBox_TCoinType.SelectedIndex = 0;
                    textBox4.Text = "0";
                    textBox4.Enabled = true;
                    label_TChange.Text = "產品包ID :";
                    textBox6.Enabled = false;
                    textBox6.Text = "1";
                    textBox_PHP.Text = "0";
                    textBox_PATK.Text = "0";
                    textBox_PRHP.Text = "0";
                    textBox_PHP.Enabled = false;
                    textBox_PATK.Enabled = false;
                    textBox_PRHP.Enabled = false;
                    textBox_TLevel.Text = "0";
                    textBox_TLevel.Enabled = false;
                    TWopInitialize();

                    break;

                case 2:

                    TItemType = "2";
                    comboBox_TSP.Enabled = true;
                    comboBox_TCoinType.Enabled = false;
                    comboBox_TCoinType.SelectedIndex = 0;
                    textBox4.Text = "0";
                    textBox4.Enabled = true;
                    label_TChange.Text = "武將ID :";
                    textBox6.Enabled = false;
                    textBox6.Text = "1";
                    textBox_PHP.Enabled = true;
                    textBox_PATK.Enabled = true;
                    textBox_PRHP.Enabled = true;
                    textBox_TLevel.Text = "0";
                    textBox_TLevel.Enabled = true;

                    checkBox_TWop1.Enabled = true;
                    checkBox_TWop2.Enabled = true;
                    checkBox_TWop3.Enabled = true;
                    checkBox_TWop4.Enabled = true;
                    checkBox_TWop5.Enabled = true;
                    checkBox_TWop6.Enabled = true;
                    checkBox_TWop7.Enabled = true;
                    checkBox_TWop8.Enabled = true;

                    break;

                case 3:

                    TItemType = "2";
                    comboBox_TSP.Enabled = true;
                    comboBox_TCoinType.Enabled = false;
                    comboBox_TCoinType.SelectedIndex = 0;
                    textBox4.Text = "0";
                    textBox4.Enabled = true;
                    label_TChange.Text = "武器ID :";
                    textBox6.Enabled = false;
                    textBox6.Text = "1";
                    textBox_PHP.Text = "0";
                    textBox_PATK.Text = "0";
                    textBox_PRHP.Text = "0";
                    textBox_PHP.Enabled = false;
                    textBox_PATK.Enabled = false;
                    textBox_PRHP.Enabled = false;
                    textBox_TLevel.Text = "0";
                    textBox_TLevel.Enabled = true;

                    checkBox_TWop1.Enabled = true;
                    checkBox_TWop2.Enabled = true;
                    checkBox_TWop3.Enabled = true;
                    checkBox_TWop4.Enabled = true;
                    checkBox_TWop5.Enabled = true;
                    checkBox_TWop6.Enabled = true;
                    checkBox_TWop7.Enabled = true;
                    checkBox_TWop8.Enabled = true;

                    break;

                case 4:

                    TItemType = "5,1";
                    comboBox_TSP.SelectedIndex = 0;
                    comboBox_TSP.Enabled = false;
                    textBox4.Text = "0";
                    textBox4.Enabled = false;
                    comboBox_TCoinType.Enabled = true;
                    comboBox_TCoinType.SelectedIndex = 0;
                    TItemID = "1";
                    label_TChange.Text = "不需填寫 :";
                    textBox6.Enabled = true;
                    textBox_PHP.Text = "0";
                    textBox_PATK.Text = "0";
                    textBox_PRHP.Text = "0";
                    textBox_PHP.Enabled = false;
                    textBox_PATK.Enabled = false;
                    textBox_PRHP.Enabled = false;
                    textBox_TLevel.Text = "0";
                    textBox_TLevel.Enabled = false;
                    TWopInitialize();

                    break;

                case 5:

                    TItemType = "5,1";
                    comboBox_TSP.Enabled = false;
                    comboBox_TCoinType.Enabled = false;
                    textBox4.Text = "0";
                    textBox4.Enabled = true;
                    label_TChange.Text = "產品包ID :";
                    comboBox_TCoinType.SelectedIndex = 0;
                    textBox6.Enabled = false;
                    textBox6.Text = "1";
                    textBox_PHP.Enabled = false;
                    textBox_PATK.Enabled = false;
                    textBox_PRHP.Enabled = false;
                    textBox_PHP.Text = "0";
                    textBox_PATK.Text = "0";
                    textBox_PRHP.Text = "0";
                    textBox_TLevel.Text = "0";
                    textBox_TLevel.Enabled = false;
                    TWopInitialize();

                    break;

                case 6:

                    TItemType = "5,2";
                    comboBox_TSP.Enabled = true;
                    comboBox_TCoinType.Enabled = false;
                    comboBox_TCoinType.SelectedIndex = 0;
                    textBox4.Text = "0";
                    textBox4.Enabled = true;
                    label_TChange.Text = "武將ID :";
                    textBox6.Enabled = false;
                    textBox6.Text = "1";
                    textBox_PHP.Enabled = true;
                    textBox_PATK.Enabled = true;
                    textBox_PRHP.Enabled = true;
                    textBox_TLevel.Text = "0";
                    textBox_TLevel.Enabled = true;

                    checkBox_TWop1.Enabled = true;
                    checkBox_TWop2.Enabled = true;
                    checkBox_TWop3.Enabled = true;
                    checkBox_TWop4.Enabled = true;
                    checkBox_TWop5.Enabled = true;
                    checkBox_TWop6.Enabled = true;
                    checkBox_TWop7.Enabled = true;
                    checkBox_TWop8.Enabled = true;

                    break;

                case 7:

                    TItemType = "5,2";
                    comboBox_TSP.Enabled = true;
                    comboBox_TCoinType.Enabled = false;
                    comboBox_TCoinType.SelectedIndex = 0;
                    textBox4.Text = "0";
                    textBox4.Enabled = true;
                    label_TChange.Text = "武器ID :";
                    textBox6.Enabled = false;
                    textBox6.Text = "1";
                    textBox_PHP.Text = "0";
                    textBox_PATK.Text = "0";
                    textBox_PRHP.Text = "0";
                    textBox_PHP.Enabled = false;
                    textBox_PATK.Enabled = false;
                    textBox_PRHP.Enabled = false;
                    textBox_TLevel.Text = "0";
                    textBox_TLevel.Enabled = true;

                    checkBox_TWop1.Enabled = true;
                    checkBox_TWop2.Enabled = true;
                    checkBox_TWop3.Enabled = true;
                    checkBox_TWop4.Enabled = true;
                    checkBox_TWop5.Enabled = true;
                    checkBox_TWop6.Enabled = true;
                    checkBox_TWop7.Enabled = true;
                    checkBox_TWop8.Enabled = true;

                    break;
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_TSP.SelectedIndex)
            {
                case 0:

                    TSP = "0";

                    break;

                case 1:

                    TSP = "1";

                    break;
            }
        }
                
        private void dateTimePicker_TEnd_ValueChanged(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged_2(object sender, EventArgs e)
        {
            switch (comboBox_TCountry.SelectedIndex)
            {
                case 0:

                    TCountry = 1;

                    break;

                case 1:

                    TCountry = 2;

                    break;

                case 2:
                    
                    TCountry = 3;

                    break;

                case 3:

                    TCountry = 4;

                    break;
            }
        }

        private void comboBox_TCoinType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_TCoinType.SelectedIndex)
            {

                case 0:

                    TItemID = "1";
                    //CoinType = "鑽石";

                    break;

                case 1:

                    TItemID = "2";
                    //CoinType = "金幣";

                    break;

                case 2:

                    TItemID = "3";
                    //CoinType = "銀幣";

                    break;

                case 3:

                    TItemID = "4";
                    //CoinType = "金牌";

                    break;

                case 4:

                    TItemID = "5";
                    //CoinType = "聖旨";

                    break;

                case 5:

                    TItemID = "6";
                    //CoinType = "小體力卷";

                    break;

                case 6:

                    TItemID = "7";
                    //CoinType = "中體力卷";

                    break;

                case 7:

                    TItemID = "8";
                    //CoinType = "大體力卷";

                    break;

                case 8:

                    TItemID = "9";
                    // CoinType = "白寶石";

                    break;

                case 9:

                    TItemID = "10";
                    //CoinType = "紅寶石";

                    break;

                case 10:

                    TItemID = "11";
                    //CoinType = "藍寶石";

                    break;

                case 11:

                    TItemID = "12";
                    //CoinType = "黃寶石";

                    break;

                case 12:

                    TItemID = "13";
                    //CoinType = "紫寶石";

                    break;
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            TItemID = textBox4.Text;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        #region 分析紀錄查詢

        int AnalysisCountry = 0;
        string AnalysisPickTime = "yyyyMMdd"; //選擇的日期
        string MoneySql = "";

        private void comboBox_AnalysisCountry_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_AnalysisCountry.SelectedIndex)
            {
                case 0:

                    AnalysisCountry = 1;

                    break;

                case 1:

                    AnalysisCountry = 2;

                    break;

                case 2:

                    AnalysisCountry = 3;

                    break;

            }
        }

        private void dateTimePicker_Analysis_ValueChanged(object sender, EventArgs e)
        {
            AnalysisPickTime = dateTimePicker_Analysis.Value.ToString("yyyyMMdd");
        }

        private void button_AnalysisClean_Click(object sender, EventArgs e)
        {
            textBox_AnalysisPassword.Text = "";
            textBox_AnalysisDis.Text = "";
            textBox_LoginDis.Text = "";
            textBox_SavePointDis.Text = "";
            textBox_MoneyList.Text = "";
        }

        private void button_AnalysisSearch_Click(object sender, EventArgs e)
        {
            if (textBox_AnalysisPassword.Text == string.Empty)
            {
                MessageBox.Show("請填入密碼");
            }

            else
            {
                //初始化
                textBox_AnalysisDis.Text = "";
                textBox_LoginDis.Text = "";
                textBox_SavePointDis.Text = "";
                textBox_MoneyList.Text = "";

                //耗點
                int CostPoint = 29;
                int[] Counts = new int[CostPoint + 1];
                int[] Role = new int[CostPoint + 1];
                int AccountNum = 0;

                for (int i = 0; i <= CostPoint; i++)
                {
                    Counts[i] = 0;
                    Role[i] = 0;
                }
                int Total_CON = 0;

                switch (AnalysisCountry)
                {
                    case 1:
                        if (textBox_AnalysisPassword.Text == SendPassWord)
                        {
                            string Tsql = "SELECT Kind, Cause, V2 FROM Money_" + AnalysisPickTime + " WHERE (Kind = 2);";
                            Tsql += "SELECT UserID, COUNT(PuzzleWeb) AS Expr1 FROM Money_" + AnalysisPickTime + " WHERE (Kind = 2) GROUP BY UserID;";
                            Tsql += "SELECT Kind, ServerID, Cause, V2 ,V4  FROM Money_" + AnalysisPickTime + " WHERE (Kind = 1);";
                            Tsql += "SELECT Kind, Cause, UserID, V2, V4, SaveDate FROM Money_" + AnalysisPickTime + " WHERE " + MoneySql + " ORDER BY SaveDate";

                            try
                            {
                                DataSet ds = new DataSet();
                                ds = TWsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);

                                #region 耗點分析

                                for (int k = 0; k < ds.Tables[0].Rows.Count; k++)
                                {
                                    switch (ds.Tables[0].Rows[k]["Cause"].ToString())
                                    {
                                        case "1":
                                            // = "購買能量　　　　　";
                                            Role[1]++;
                                            Counts[1] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            break;
                                        case "2":
                                            // = "抽武將　　　　　　";
                                            Counts[2] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[2]++;
                                            break;
                                        case "3":
                                            // = "開寶箱　　　　　　";
                                            Counts[3] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[3]++;
                                            break;
                                        case "4":
                                            // = "消除競技場ＣＤ　　";
                                            Counts[4] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[4]++;
                                            break;
                                        case "5":
                                            // = "重置攻城戰　　　　";
                                            Counts[5] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[5]++;
                                            break;
                                        case "6":
                                            // = "重置攻城寶箱　　　";
                                            Counts[6] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[6]++;
                                            break;
                                        case "7":
                                            // = "購買懸賞格子　　　";
                                            Counts[7] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[7]++;
                                            break;
                                        case "8":
                                            // = "轉蛋扣除　　　　　";
                                            Counts[8] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[8]++;
                                            break;
                                        case "9":
                                            // = "購買武將欄位　　　";
                                            Counts[9] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[9]++;
                                            break;
                                        case "10":
                                            // = "推圖復活　　　　　";
                                            Counts[10] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[10]++;
                                            break;
                                        case "11":
                                            // = "戰場使用精英援軍　";
                                            Counts[11] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[11]++;
                                            break;
                                        case "12":
                                            // = "快速通關　　　　　";
                                            Counts[12] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[12]++;
                                            break;
                                        case "13":
                                            Counts[13] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[13]++;
                                            // = "購買ＰＶＰ挑戰次數";
                                            break;
                                        case "16":
                                            Counts[14] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[14]++;
                                            // = "轉盤重置　　　　　";
                                            break;
                                        case "17":
                                            Counts[15] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[15]++;
                                            // = "購買裝備欄位　　 　";
                                            break;
                                        case "18":
                                            if (Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString()) == 30)
                                            {
                                                Counts[16] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                                Role[16]++;
                                                // = "ＶＩＰ強化－３０點 ";
                                            }

                                            if (Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString()) == 100)
                                            {
                                                Counts[17] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                                Role[17]++;
                                                //Cause = "ＶＩＰ強化－１００點";
                                            }
                                            break;
                                        case "14":
                                            Counts[18] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[18]++;
                                            // = "消世界ＢｏｓｓＣＤ　";
                                            break;
                                        case "15":
                                            Counts[19] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[19]++;
                                            // = "消掃蕩ＣＤ　　　　　";
                                            break;
                                        case "19":
                                            Counts[20] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[20]++;
                                            // = "消跨服ＰＶＰ　ＣＤ　"
                                            break;
                                        case "20":
                                            Counts[21] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[21]++;
                                            // = "買跨服ＰＶＰ挑戰次數";
                                            break;
                                        case "21":
                                            Counts[22] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[22]++;
                                            // = "建立軍團　　　　　　";
                                            break;
                                        case "22":
                                            Counts[23] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[23]++;
                                            // = "軍團捐獻　　　　　　";
                                            break;
                                        case "23":
                                            Counts[24] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[25]++;
                                            // = "接關扣鑽　　　　　　";
                                            break;
                                        case "24":
                                            Counts[25] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[25]++;
                                            // = "購買累積儲值福袋　　";
                                            break;
                                        case "25":
                                            Counts[26] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[26]++;
                                            // = "購買勢力禮盒　　　　";
                                            break;
                                        case "26":
                                            Counts[27] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[27]++;
                                            // = "買全球Ｂｏｓｓ挑戰次數";
                                            break;
                                        case "27":
                                            Counts[28] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[28]++;
                                            // = "全球Ｂｏｓｓ升級　　";
                                            break;
                                        case "28":
                                            Counts[29] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[29]++;
                                            // = "ＷＥＢ後台　　　　　";
                                            break;
                                    }

                                    Total_CON += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());

                                }

                                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                                {
                                    AccountNum++;

                                }
                                textBox_AnalysisDis.Text += "耗點總額度　　　　　 : " + Total_CON.ToString() + "\r\n" + "\r\n" + "購買能量　　　　　　 : " + Counts[1].ToString("00000000") + "  /  " + Role[1].ToString("0000") + "\r\n" + "抽武將　　　　　　　 : " + Counts[2].ToString("00000000") + "  /  " + Role[2].ToString("0000") + "\r\n" + "開寶箱　　　　　　　 : " + Counts[3].ToString("00000000") + "  /  " + Role[3].ToString("0000") + "\r\n" + "消除ＰＶＰ　ＣＤ　　 : " + Counts[4].ToString("00000000") + "  /  " + Role[4].ToString("0000") + "\r\n" + "重置攻城戰　　　　　 : " + Counts[5].ToString("00000000") + "  /  " + Role[5].ToString("0000") + "\r\n" + "重置攻城寶箱　　　　 : " + Counts[6].ToString("00000000") + "  /  " + Role[6].ToString("0000") + "\r\n" + "購買懸賞格子　　　　 : " + Counts[7].ToString("00000000") + "  /  " + Role[7].ToString("0000") + "\r\n" + "轉蛋　　　　　　　　 : " + Counts[8].ToString("00000000") + "  /  " + Role[8].ToString("0000") + "\r\n" + "購買武將欄位　　　　 : " + Counts[9].ToString("00000000") + "  /  " + Role[9].ToString("0000") + "\r\n" + "推圖復活　　　　　　 : " + Counts[10].ToString("00000000") + "  /  " + Role[10].ToString("0000") + "\r\n" + "戰場使用精英援軍　　 : " + Counts[11].ToString("00000000") + "  /  " + Role[11].ToString("0000") + "\r\n" + "快速通關　　　　　　 : " + Counts[12].ToString("00000000") + "  /  " + Role[12].ToString("0000") + "\r\n" + "購買ＰＶＰ挑戰次數　 : " + Counts[13].ToString("00000000") + "  /  " + Role[13].ToString("0000") + "\r\n" + "拉ＢＡＲ重置　　　　 : " + Counts[14].ToString("00000000") + "  /  " + Role[14].ToString("0000") + "\r\n" + "購買裝備欄位　　　　 : " + Counts[15].ToString("00000000") + "  /  " + Role[15].ToString("0000") + "\r\n" + "ＶＩＰ強化－３０點　 : " + Counts[16].ToString("00000000") + "  /  " + Role[16].ToString("0000") + "\r\n" + "ＶＩＰ強化－１００點 : " + Counts[17].ToString("00000000") + "  /  " + Role[17].ToString("0000") + "\r\n" + "消世界ＢｏｓｓＣＤ　 : " + Counts[18].ToString("00000000") + "  /  " + Role[18].ToString("0000") + "\r\n" + "消掃蕩ＣＤ　　　　     : " + Counts[19].ToString("00000000") + "  /  " + Role[19].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "消跨服ＰＶＰ　ＣＤ　 : " + Counts[20].ToString("00000000") + "  /  " + Role[20].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "買跨服ＰＶＰ挑戰次數 : " + Counts[21].ToString("00000000") + "  /  " + Role[21].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "建立軍團　　　　 　　: " + Counts[22].ToString("00000000") + "  /  " + Role[22].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "軍團捐獻　　　　 　　: " + Counts[23].ToString("00000000") + "  /  " + Role[23].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "接關扣鑽　　　　 　　: " + Counts[24].ToString("00000000") + "  /  " + Role[24].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "購買累積儲值福袋　　 : " + Counts[25].ToString("00000000") + "  /  " + Role[25].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "購買勢力禮盒　　　　 : " + Counts[26].ToString("00000000") + "  /  " + Role[26].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "買全球Ｂｏｓｓ挑戰次數: " + Counts[27].ToString("00000000") + "  /  " + Role[27].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "全球Ｂｏｓｓ升級　　 : " + Counts[28].ToString("00000000") + "  /  " + Role[28].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "ＷＥＢ後台　　　　　 : " + Counts[29].ToString("00000000") + "  /  " + Role[29].ToString("0000");
                                textBox_AnalysisDis.Text += "\r\n" + "\r\n" + "耗點帳號數量 : " + AccountNum;
                                #endregion

                                #region 儲點分析

                                int ServerNum = 10; //伺服器總數
                                int AllServerTotal = 0;
                                int AllServerPoint_Apple = 0;
                                int AllServerPoint_Google = 0;
                                int AllServerPoint_GoogleF = 0;
                                int AllServerPoint_MyCard = 0;
                                int AllServerPoint_MyCardGoogle = 0;
                                int AllServerPoint_980x = 0;
                                int AllServerPoint_Bug = 0;
                                int AllServerPoint_東方 = 0;
                                int AllServerPoint_松崗 = 0;

                                textBox_SavePointDis.Text += "台幣資訊 : " + "\r\n";

                                for (int j = 1; j <= ServerNum; j++)
                                {
                                    int Point_Apple13 = 0;
                                    int Point_Google13 = 0;
                                    int Point_Google13F = 0;
                                    int Point_MyCard13 = 0;
                                    int Point_MyCardGoogle13 = 0;
                                    int Point_980x13 = 0;
                                    int Point_Bug13 = 0;
                                    int Point_東方 = 0;
                                    int Point_松崗 = 0;
                                    int Point_Total13 = 0;

                                    for (int i = 0; i < ds.Tables[2].Rows.Count; i++)
                                    {
                                        if (Convert.ToInt32(ds.Tables[2].Rows[i]["ServerID"].ToString()) == j)
                                        {
                                            switch (ds.Tables[2].Rows[i]["Cause"].ToString())
                                            {
                                                #region 流向分類

                                                case "9":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "20":

                                                            Point_Bug13 += 20;

                                                            break;

                                                        case "30":

                                                            Point_Bug13 += 30;

                                                            break;

                                                        case "53":

                                                            Point_Bug13 += 50;

                                                            break;

                                                        case "168":

                                                            Point_Bug13 += 150;

                                                            break;

                                                        case "480":

                                                            Point_Bug13 += 400;

                                                            break;

                                                        case "1000":

                                                            Point_Bug13 += 800;

                                                            break;

                                                        case "1560":

                                                            Point_Bug13 += 1200;

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());

                                                            Point_Bug13 += Convert.ToInt32(Math.Floor(x / 1.3));

                                                            break;
                                                    }

                                                    break;

                                                case "10":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "10":
                                                            Point_MyCardGoogle13 += 10;
                                                            break;

                                                        case "53":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 50;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 50;
                                                            }

                                                            break;

                                                        case "168":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 150;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 150;
                                                            }

                                                            break;

                                                        case "336":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 300;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 300;
                                                            }

                                                            break;

                                                        case "403":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 350;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 350;
                                                            }

                                                            break;

                                                        case "480":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 400;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 400;
                                                            }

                                                            break;

                                                        case "540":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 450;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 450;
                                                            }

                                                            break;

                                                        case "600":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 500;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 500;
                                                            }

                                                            break;

                                                        case "1270":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 1000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 1000;
                                                            }

                                                            break;

                                                        case "1460":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 1150;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 1150;
                                                            }

                                                            break;

                                                        case "2600":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 2000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 2000;
                                                            }

                                                            break;

                                                        case "4000":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 3000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 3000;
                                                            }

                                                            break;

                                                        case "6750":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 5000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 5000;
                                                            }

                                                            break;

                                                    }
                                                    break;

                                                case "13":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "64":

                                                            Point_Apple13 += 60;

                                                            break;

                                                        case "336":

                                                            Point_Apple13 += 300;

                                                            break;

                                                        case "540":

                                                            Point_Apple13 += 450;

                                                            break;

                                                        case "958":

                                                            Point_Apple13 += 750;

                                                            break;

                                                        case "4000":

                                                            Point_Apple13 += 2990;

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());

                                                            Point_Apple13 += Convert.ToInt32(Math.Floor(x / 1.3));

                                                            break;

                                                    }
                                                    break;
                                                case "14":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "10":
                                                            Point_Google13F += 0;
                                                            break;

                                                        case "64":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 60;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 60;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 60;
                                                            }

                                                            break;

                                                        case "336":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 300;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 300;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 300;
                                                            }

                                                            break;

                                                        case "540":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 450;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 450;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 450;
                                                            }

                                                            break;

                                                        case "958":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 750;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 750;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 750;
                                                            }

                                                            break;

                                                        case "4000":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 2990;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 2990;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 2990;
                                                            }

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());


                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += Convert.ToInt32(Math.Floor(x / 1.3));
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += Convert.ToInt32(Math.Floor(x / 1.3));
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += Convert.ToInt32(Math.Floor(x / 1.3));
                                                            }

                                                            break;
                                                    }
                                                    break;
                                                case "16":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "168":

                                                            Point_980x13 += 150;

                                                            break;

                                                        case "480":

                                                            Point_980x13 += 400;

                                                            break;

                                                        case "1560":

                                                            Point_980x13 += 1200;

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());

                                                            Point_980x13 += Convert.ToInt32(Math.Floor(x / 1.3));

                                                            break;
                                                    }

                                                    break;

                                                case "24":
                                                    Point_東方 += Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());
                                                    break;

                                                case "15":
                                                    Point_松崗 += Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());
                                                    break;


                                                #endregion
                                            }
                                        }
                                    }

                                    Point_Total13 = Convert.ToInt32(Math.Floor(Point_Apple13 + Point_Google13 + Point_Google13F + Point_MyCard13 + Point_MyCardGoogle13 + Point_980x13 + Point_Bug13 + (Point_東方 / 1.3)));

                                    AllServerTotal += Point_Total13;
                                    AllServerPoint_Apple += Point_Apple13;
                                    AllServerPoint_Google += Point_Google13;
                                    AllServerPoint_GoogleF += Point_Google13F;
                                    AllServerPoint_MyCard += Point_MyCard13;
                                    AllServerPoint_MyCardGoogle += Point_MyCardGoogle13;
                                    AllServerPoint_980x += Point_980x13;
                                    AllServerPoint_Bug += Point_Bug13;
                                    AllServerPoint_東方 += Point_東方;
                                    AllServerPoint_松崗 += Point_松崗;

                                    textBox_SavePointDis.Text += "\r\n" + "Server : " + j + "\r\n" + "Total : " + Point_Total13.ToString("0") + "\r\n" + " Apple : " + Point_Apple13.ToString("0") + "\r\n" + " Google : " + Point_Google13.ToString("0") + "\r\n" + " GoogleFull : " + Point_Google13F.ToString("0") + "\r\n" + " MyCard : " + Point_MyCard13.ToString("0") + "\r\n" + " MCG : " + Point_MyCardGoogle13.ToString("0") + "\r\n" + " 980xCard : " + Point_980x13.ToString("0") + "\r\n" + " BugCard : " + Point_Bug13.ToString("0") + "\r\n" + " 老李卡 : " + (Point_東方 / 1.3).ToString("0") + "\r\n" + " 松崗 : " + (Point_松崗 / 1.3).ToString("0") + "\r\n";
                                }

                                textBox_SavePointDis.Text += "\r\n" + "全伺服器總合 : " + "\r\n" + "Total : " + AllServerTotal + "\r\n" + " Apple : " + AllServerPoint_Apple + "\r\n" + " Google : " + AllServerPoint_Google + "\r\n" + " GoogleFull : " + AllServerPoint_GoogleF + "\r\n" + " MyCard : " + AllServerPoint_MyCard + "\r\n" + " MCG : " + AllServerPoint_MyCardGoogle + "\r\n" + " 980xCard : " + AllServerPoint_980x + "\r\n" + " BugCard : " + AllServerPoint_Bug + "\r\n" + " 老李卡 : " + AllServerPoint_東方 + "\r\n" + " 松崗 : " + AllServerPoint_松崗 + "\r\n";


                                #endregion

                                #region 金錢流向



                                switch (comboBox_MoneyList.SelectedIndex)
                                {
                                    case 0:
                                        textBox_MoneyList.Text += "松崗金流 : " + "\r\n" + "\r\n";

                                        for (int x = 0; x < ds.Tables[3].Rows.Count; x++)
                                        {
                                            textBox_MoneyList.Text += "ID : " + ds.Tables[3].Rows[x]["UserID"].ToString() + "   儲值鑽石 : " + ds.Tables[3].Rows[x]["V2"].ToString() + "   時間 : " + Convert.ToDateTime(ds.Tables[3].Rows[x]["SaveDate"]).ToString("yyyy/MM/dd HH:mm:ss") + "\r\n";
                                        }

                                        break;

                                    case 1:
                                        textBox_MoneyList.Text += "吞食Q鑽金流 :  " + "\r\n" + "\r\n"; ;

                                        for (int x = 0; x < ds.Tables[3].Rows.Count; x++)
                                        {
                                            int MCG = 0;
                                            string MCGCome = "";

                                            switch (ds.Tables[3].Rows[x]["Cause"].ToString())
                                            {
                                                case "10":
                                                case "22":
                                                    MCGCome = "ＭｙＣａｒｄ　";
                                                    break;

                                                case "14":
                                                    MCGCome = "Ｇｏｏｇｌｅ　";
                                                    break;
                                            }

                                            switch (Convert.ToInt32(ds.Tables[3].Rows[x]["V2"]))
                                            {
                                                case 10:
                                                    MCG = 10;
                                                    break;

                                                case 53:
                                                    MCG = 50;
                                                    break;

                                                case 64:
                                                    MCG = 60;
                                                    break;

                                                case 168:
                                                    MCG = 150;
                                                    break;

                                                case 336:
                                                    MCG = 300;
                                                    break;

                                                case 403:
                                                    MCG = 350;
                                                    break;

                                                case 480:
                                                    MCG = 400;
                                                    break;

                                                case 540:
                                                    MCG = 450;
                                                    break;

                                                case 600:
                                                    MCG = 500;
                                                    break;

                                                case 958:
                                                    MCG = 750;
                                                    break;

                                                case 1270:
                                                    MCG = 1000;
                                                    break;

                                                case 1460:
                                                    MCG = 1150;
                                                    break;

                                                case 2600:
                                                    MCG = 2000;
                                                    break;

                                                case 4000:
                                                    MCG = 3000;
                                                    break;

                                                case 6750:
                                                    MCG = 5000;
                                                    break;

                                                default:
                                                    MCG = Convert.ToInt32(ds.Tables[3].Rows[x]["V2"]);

                                                    break;

                                            }

                                            textBox_MoneyList.Text += MCGCome +"ID : " + ds.Tables[3].Rows[x]["UserID"].ToString() + "   儲值鑽石 : " + ds.Tables[3].Rows[x]["V2"].ToString() + "   儲值台幣 : " + MCG + "   時間 : " + Convert.ToDateTime(ds.Tables[3].Rows[x]["SaveDate"]).ToString("yyyy/MM/dd HH:mm:ss") + "\r\n";
                                        }
                                        break;
                                }

                                #endregion

                            }

                            catch
                            {
                                MessageBox.Show("資料查詢錯誤");
                            }

                            finally
                            {
                            }
                        }


                        else
                        {
                            MessageBox.Show("密碼錯誤");
                        }

                        if (textBox_AnalysisPassword.Text == SendPassWord)
                        {
                            string Tsql = "SELECT TOP (1) Area, Account, DistinctAccount, GS1, GS1Max, GS2, GS2Max, GS3, GS3Max, GS4, GS4Max, GS5, GS5Max, GS6, GS6Max, GS7, GS7Max, GS8, GS8Max, GS9, GS9Max, SaveDAte FROM OnlineRole WHERE (Area = 'TW') AND (SaveDAte > CONVERT(DATETIME, '" + AnalysisPickTime + " 00:00:00', 000)) AND (SaveDAte < CONVERT(DATETIME, '" + AnalysisPickTime + " 23:59:59', 000)) ORDER BY SaveDAte DESC";

                            try
                            {
                                DataSet ds = new DataSet();
                                ds = TWsql.Get_SQL_Analysis_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                                textBox_LoginDis.Text += "國家 : 台灣" + "\r\n" + "\r\n" + "總帳號數 : " + ds.Tables[0].Rows[0]["Account"].ToString() + "\r\n" + "\r\n" + "不重複登入數 : " + ds.Tables[0].Rows[0]["DistinctAccount"].ToString() + "\r\n" + "\r\n" + "\r\n";
                                textBox_LoginDis.Text += "即時人數 : " + "\r\n" + "\r\n" + "Server 1 :  " + ds.Tables[0].Rows[0]["GS1"].ToString() + "\r\n" + "Server 2 :  " + ds.Tables[0].Rows[0]["GS2"].ToString() + "\r\n" + "Server 3 :  " + ds.Tables[0].Rows[0]["GS3"].ToString() + "\r\n" + "Server 4 :  " + ds.Tables[0].Rows[0]["GS4"].ToString() + "\r\n" + "Server 5 :  " + ds.Tables[0].Rows[0]["GS5"].ToString() + "\r\n" + "Server 6 :  " + ds.Tables[0].Rows[0]["GS6"].ToString() + "\r\n" + "Server 7 :  " + ds.Tables[0].Rows[0]["GS7"].ToString() + "\r\n" + "Server 8 :  " + ds.Tables[0].Rows[0]["GS8"].ToString() + "\r\n" + "Server 9 :  " + ds.Tables[0].Rows[0]["GS9"].ToString() + "\r\n" + "Server 10 :  " + ds.Tables[0].Rows[0]["GS10"].ToString() + "\r\n" + "\r\n" + "\r\n";
                                textBox_LoginDis.Text += "最高人數 : " + "\r\n" + "\r\n" + "Server 1 :  " + ds.Tables[0].Rows[0]["GS1Max"].ToString() + "\r\n" + "Server 2 :  " + ds.Tables[0].Rows[0]["GS2Max"].ToString() + "\r\n" + "Server 3  : " + ds.Tables[0].Rows[0]["GS3Max"].ToString() + "\r\n" + "Server 4 :  " + ds.Tables[0].Rows[0]["GS4Max"].ToString() + "\r\n" + "Server 5 :  " + ds.Tables[0].Rows[0]["GS5Max"].ToString() + "\r\n" + "Server 6 :  " + ds.Tables[0].Rows[0]["GS6Max"].ToString() + "\r\n" + "Server 7 :  " + ds.Tables[0].Rows[0]["GS7Max"].ToString() + "\r\n" + "Server 8 :  " + ds.Tables[0].Rows[0]["GS8Max"].ToString() + "\r\n" + "Server 9 :  " + ds.Tables[0].Rows[0]["GS9Max"].ToString() + "\r\n" + "Server 10 :  " + ds.Tables[0].Rows[0]["GS10Max"].ToString() + "\r\n" + "\r\n";
                                MessageBox.Show("資料來了");
                            }

                            catch
                            {
                                MessageBox.Show("資料查詢錯誤");
                            }
                        }
                        else
                        {
                        }

                        break;

                    case 2:

                        if (textBox_AnalysisPassword.Text == SendPassWord)
                        {
                            string Tsql = "SELECT Kind, Cause, V2 FROM Money_" + AnalysisPickTime + " WHERE (Kind = 2);";
                            Tsql += "SELECT UserID, COUNT(PuzzleWeb) AS Expr1 FROM Money_" + AnalysisPickTime + " WHERE (Kind = 2) GROUP BY UserID;";
                            Tsql += "SELECT Kind, ServerID, Cause, V2 ,V4  FROM Money_" + AnalysisPickTime + " WHERE (Kind = 1);";
                            Tsql += "SELECT Kind, Cause, UserID, V2, V4, SaveDate FROM Money_" + AnalysisPickTime + " WHERE " + MoneySql + " ORDER BY SaveDate";

                            try
                            {
                                DataSet ds = new DataSet();
                                ds = HKsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);

                                #region 耗點分析

                                for (int k = 0; k < ds.Tables[0].Rows.Count; k++)
                                {
                                    switch (ds.Tables[0].Rows[k]["Cause"].ToString())
                                    {
                                        case "1":
                                            // = "購買能量　　　　　";
                                            Role[1]++;
                                            Counts[1] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            break;
                                        case "2":
                                            // = "抽武將　　　　　　";
                                            Counts[2] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[2]++;
                                            break;
                                        case "3":
                                            // = "開寶箱　　　　　　";
                                            Counts[3] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[3]++;
                                            break;
                                        case "4":
                                            // = "消除競技場ＣＤ　　";
                                            Counts[4] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[4]++;
                                            break;
                                        case "5":
                                            // = "重置攻城戰　　　　";
                                            Counts[5] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[5]++;
                                            break;
                                        case "6":
                                            // = "重置攻城寶箱　　　";
                                            Counts[6] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[6]++;
                                            break;
                                        case "7":
                                            // = "購買懸賞格子　　　";
                                            Counts[7] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[7]++;
                                            break;
                                        case "8":
                                            // = "轉蛋扣除　　　　　";
                                            Counts[8] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[8]++;
                                            break;
                                        case "9":
                                            // = "購買武將欄位　　　";
                                            Counts[9] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[9]++;
                                            break;
                                        case "10":
                                            // = "推圖復活　　　　　";
                                            Counts[10] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[10]++;
                                            break;
                                        case "11":
                                            // = "戰場使用精英援軍　";
                                            Counts[11] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[11]++;
                                            break;
                                        case "12":
                                            // = "快速通關　　　　　";
                                            Counts[12] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[12]++;
                                            break;
                                        case "13":
                                            Counts[13] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[13]++;
                                            // = "購買ＰＶＰ挑戰次數";
                                            break;
                                        case "16":
                                            Counts[14] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[14]++;
                                            // = "轉盤重置　　　　　";
                                            break;
                                        case "17":
                                            Counts[15] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[15]++;
                                            // = "購買裝備欄位　　 　";
                                            break;
                                        case "18":
                                            if (Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString()) == 30)
                                            {
                                                Counts[16] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                                Role[16]++;
                                                // = "ＶＩＰ強化－３０點 ";
                                            }

                                            if (Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString()) == 100)
                                            {
                                                Counts[17] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                                Role[17]++;
                                                //Cause = "ＶＩＰ強化－１００點";
                                            }
                                            break;
                                        case "14":
                                            Counts[18] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[18]++;
                                            // = "消世界ＢｏｓｓＣＤ　";
                                            break;
                                        case "15":
                                            Counts[19] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[19]++;
                                            // = "消掃蕩ＣＤ　　　　　";
                                            break;
                                        case "19":
                                            Counts[20] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[20]++;
                                            // = "消跨服ＰＶＰ　ＣＤ　"
                                            break;
                                        case "20":
                                            Counts[21] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[21]++;
                                            // = "買跨服ＰＶＰ挑戰次數";
                                            break;
                                        case "21":
                                            Counts[22] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[22]++;
                                            // = "建立軍團　　　　　　";
                                            break;
                                        case "22":
                                            Counts[23] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[23]++;
                                            // = "軍團捐獻　　　　　　";
                                            break;
                                        case "23":
                                            Counts[24] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[25]++;
                                            // = "接關扣鑽　　　　　　";
                                            break;
                                        case "24":
                                            Counts[25] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[25]++;
                                            // = "購買累積儲值福袋　　";
                                            break;
                                        case "25":
                                            Counts[26] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[26]++;
                                            // = "購買勢力禮盒　　　　";
                                            break;
                                        case "26":
                                            Counts[27] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[27]++;
                                            // = "買全球Ｂｏｓｓ挑戰次數";
                                            break;
                                        case "27":
                                            Counts[28] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[28]++;
                                            // = "全球Ｂｏｓｓ升級　　";
                                            break;
                                        case "28":
                                            Counts[29] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[29]++;
                                            // = "ＷＥＢ後台　　　　　";
                                            break;
                                    }

                                    Total_CON += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());

                                }

                                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                                {
                                    AccountNum++;

                                }
                                textBox_AnalysisDis.Text += "耗點總額度　　　　　 : " + Total_CON.ToString() + "\r\n" + "\r\n" + "購買能量　　　　　　 : " + Counts[1].ToString("00000000") + "  /  " + Role[1].ToString("0000") + "\r\n" + "抽武將　　　　　　　 : " + Counts[2].ToString("00000000") + "  /  " + Role[2].ToString("0000") + "\r\n" + "開寶箱　　　　　　　 : " + Counts[3].ToString("00000000") + "  /  " + Role[3].ToString("0000") + "\r\n" + "消除ＰＶＰ　ＣＤ　　 : " + Counts[4].ToString("00000000") + "  /  " + Role[4].ToString("0000") + "\r\n" + "重置攻城戰　　　　　 : " + Counts[5].ToString("00000000") + "  /  " + Role[5].ToString("0000") + "\r\n" + "重置攻城寶箱　　　　 : " + Counts[6].ToString("00000000") + "  /  " + Role[6].ToString("0000") + "\r\n" + "購買懸賞格子　　　　 : " + Counts[7].ToString("00000000") + "  /  " + Role[7].ToString("0000") + "\r\n" + "轉蛋　　　　　　　　 : " + Counts[8].ToString("00000000") + "  /  " + Role[8].ToString("0000") + "\r\n" + "購買武將欄位　　　　 : " + Counts[9].ToString("00000000") + "  /  " + Role[9].ToString("0000") + "\r\n" + "推圖復活　　　　　　 : " + Counts[10].ToString("00000000") + "  /  " + Role[10].ToString("0000") + "\r\n" + "戰場使用精英援軍　　 : " + Counts[11].ToString("00000000") + "  /  " + Role[11].ToString("0000") + "\r\n" + "快速通關　　　　　　 : " + Counts[12].ToString("00000000") + "  /  " + Role[12].ToString("0000") + "\r\n" + "購買ＰＶＰ挑戰次數　 : " + Counts[13].ToString("00000000") + "  /  " + Role[13].ToString("0000") + "\r\n" + "拉ＢＡＲ重置　　　　 : " + Counts[14].ToString("00000000") + "  /  " + Role[14].ToString("0000") + "\r\n" + "購買裝備欄位　　　　 : " + Counts[15].ToString("00000000") + "  /  " + Role[15].ToString("0000") + "\r\n" + "ＶＩＰ強化－３０點　 : " + Counts[16].ToString("00000000") + "  /  " + Role[16].ToString("0000") + "\r\n" + "ＶＩＰ強化－１００點 : " + Counts[17].ToString("00000000") + "  /  " + Role[17].ToString("0000") + "\r\n" + "消世界ＢｏｓｓＣＤ　 : " + Counts[18].ToString("00000000") + "  /  " + Role[18].ToString("0000") + "\r\n" + "消掃蕩ＣＤ　　　　     : " + Counts[19].ToString("00000000") + "  /  " + Role[19].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "消跨服ＰＶＰ　ＣＤ　 : " + Counts[20].ToString("00000000") + "  /  " + Role[20].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "買跨服ＰＶＰ挑戰次數 : " + Counts[21].ToString("00000000") + "  /  " + Role[21].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "建立軍團　　　　 　　: " + Counts[22].ToString("00000000") + "  /  " + Role[22].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "軍團捐獻　　　　 　　: " + Counts[23].ToString("00000000") + "  /  " + Role[23].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "接關扣鑽　　　　 　　: " + Counts[24].ToString("00000000") + "  /  " + Role[24].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "購買累積儲值福袋　　 : " + Counts[25].ToString("00000000") + "  /  " + Role[25].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "購買勢力禮盒　　　　 : " + Counts[26].ToString("00000000") + "  /  " + Role[26].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "買全球Ｂｏｓｓ挑戰次數: " + Counts[27].ToString("00000000") + "  /  " + Role[27].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "全球Ｂｏｓｓ升級　　 : " + Counts[28].ToString("00000000") + "  /  " + Role[28].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "ＷＥＢ後台　　　　　 : " + Counts[29].ToString("00000000") + "  /  " + Role[29].ToString("0000");
                                textBox_AnalysisDis.Text += "\r\n" + "\r\n" + "耗點帳號數量 : " + AccountNum;
                                #endregion

                                #region 儲點分析

                                int ServerNum = 8; //伺服器總數
                                int AllServerTotal = 0;
                                int AllServerPoint_Apple = 0;
                                int AllServerPoint_Google = 0;
                                int AllServerPoint_GoogleF = 0;
                                int AllServerPoint_MyCard = 0;
                                int AllServerPoint_MyCardGoogle = 0;
                                int AllServerPoint_980x = 0;
                                int AllServerPoint_Bug = 0;
                                int AllServerPoint_東方 = 0;
                                int AllServerPoint_松崗 = 0;

                                textBox_SavePointDis.Text += "台幣資訊 : " + "\r\n";

                                for (int j = 1; j <= ServerNum; j++)
                                {
                                    int Point_Apple13 = 0;
                                    int Point_Google13 = 0;
                                    int Point_Google13F = 0;
                                    int Point_MyCard13 = 0;
                                    int Point_MyCardGoogle13 = 0;
                                    int Point_980x13 = 0;
                                    int Point_Bug13 = 0;
                                    int Point_東方 = 0;
                                    int Point_松崗 = 0;
                                    int Point_Total13 = 0;

                                    for (int i = 0; i < ds.Tables[2].Rows.Count; i++)
                                    {
                                        if (Convert.ToInt32(ds.Tables[2].Rows[i]["ServerID"].ToString()) == j)
                                        {
                                            switch (ds.Tables[2].Rows[i]["Cause"].ToString())
                                            {
                                                #region 流向分類

                                                case "9":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "20":

                                                            Point_Bug13 += 20;

                                                            break;

                                                        case "30":

                                                            Point_Bug13 += 30;

                                                            break;

                                                        case "53":

                                                            Point_Bug13 += 50;

                                                            break;

                                                        case "168":

                                                            Point_Bug13 += 150;

                                                            break;

                                                        case "480":

                                                            Point_Bug13 += 400;

                                                            break;

                                                        case "1000":

                                                            Point_Bug13 += 800;

                                                            break;

                                                        case "1560":

                                                            Point_Bug13 += 1200;

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());

                                                            Point_Bug13 += Convert.ToInt32(Math.Floor(x / 1.3));

                                                            break;
                                                    }

                                                    break;

                                                case "10":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "10":
                                                            Point_MyCardGoogle13 += 10;
                                                            break;

                                                        case "53":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 50;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 50;
                                                            }

                                                            break;

                                                        case "168":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 150;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 150;
                                                            }

                                                            break;

                                                        case "336":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 300;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 300;
                                                            }

                                                            break;

                                                        case "403":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 350;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 350;
                                                            }

                                                            break;

                                                        case "480":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 400;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 400;
                                                            }

                                                            break;

                                                        case "540":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 450;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 450;
                                                            }

                                                            break;

                                                        case "600":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 500;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 500;
                                                            }

                                                            break;

                                                        case "1270":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 1000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 1000;
                                                            }

                                                            break;

                                                        case "1460":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 1150;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 1150;
                                                            }

                                                            break;

                                                        case "2600":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 2000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 2000;
                                                            }

                                                            break;

                                                        case "4000":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 3000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 3000;
                                                            }

                                                            break;

                                                        case "6750":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 5000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 5000;
                                                            }

                                                            break;

                                                    }
                                                    break;

                                                case "22":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "10":
                                                            Point_MyCardGoogle13 += 10;
                                                            break;

                                                        case "53":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 50;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 50;
                                                            }

                                                            break;

                                                        case "168":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 150;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 150;
                                                            }

                                                            break;

                                                        case "336":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 300;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 300;
                                                            }

                                                            break;

                                                        case "403":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 350;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 350;
                                                            }

                                                            break;

                                                        case "480":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 400;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 400;
                                                            }

                                                            break;

                                                        case "540":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 450;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 450;
                                                            }

                                                            break;

                                                        case "600":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 500;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 500;
                                                            }

                                                            break;

                                                        case "1270":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 1000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 1000;
                                                            }

                                                            break;

                                                        case "1460":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 1150;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 1150;
                                                            }

                                                            break;

                                                        case "2600":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 2000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 2000;
                                                            }

                                                            break;

                                                        case "4000":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 3000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 3000;
                                                            }

                                                            break;

                                                        case "6750":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 5000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 5000;
                                                            }

                                                            break;

                                                    }
                                                    break;

                                                case "13":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "64":

                                                            Point_Apple13 += 60;

                                                            break;

                                                        case "336":

                                                            Point_Apple13 += 300;

                                                            break;

                                                        case "540":

                                                            Point_Apple13 += 450;

                                                            break;

                                                        case "958":

                                                            Point_Apple13 += 750;

                                                            break;

                                                        case "4000":

                                                            Point_Apple13 += 2990;

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());

                                                            Point_Apple13 += Convert.ToInt32(Math.Floor(x / 1.3));

                                                            break;

                                                    }
                                                    break;
                                                case "14":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "10":
                                                            Point_Google13F += 0;
                                                            break;

                                                        case "64":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 60;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 60;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 60;
                                                            }

                                                            break;

                                                        case "336":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 300;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 300;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 300;
                                                            }

                                                            break;

                                                        case "540":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 450;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 450;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 450;
                                                            }

                                                            break;

                                                        case "958":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 750;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 750;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 750;
                                                            }

                                                            break;

                                                        case "4000":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 2990;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 2990;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 2990;
                                                            }

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());


                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += Convert.ToInt32(Math.Floor(x / 1.3));
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += Convert.ToInt32(Math.Floor(x / 1.3));
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += Convert.ToInt32(Math.Floor(x / 1.3));
                                                            }

                                                            break;
                                                    }
                                                    break;
                                                case "16":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "168":

                                                            Point_980x13 += 150;

                                                            break;

                                                        case "480":

                                                            Point_980x13 += 400;

                                                            break;

                                                        case "1560":

                                                            Point_980x13 += 1200;

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());

                                                            Point_980x13 += Convert.ToInt32(Math.Floor(x / 1.3));

                                                            break;
                                                    }

                                                    break;

                                                case "24":
                                                    Point_東方 += Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());
                                                    break;

                                                case "15":
                                                    Point_松崗 += Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());
                                                    break;


                                                #endregion
                                            }
                                        }
                                    }

                                    Point_Total13 = Convert.ToInt32(Math.Floor(Point_Apple13 + Point_Google13 + Point_Google13F + Point_MyCard13 + Point_MyCardGoogle13 + Point_980x13 + Point_Bug13 + (Point_東方 / 1.3)));

                                    AllServerTotal += Point_Total13;
                                    AllServerPoint_Apple += Point_Apple13;
                                    AllServerPoint_Google += Point_Google13;
                                    AllServerPoint_GoogleF += Point_Google13F;
                                    AllServerPoint_MyCard += Point_MyCard13;
                                    AllServerPoint_MyCardGoogle += Point_MyCardGoogle13;
                                    AllServerPoint_980x += Point_980x13;
                                    AllServerPoint_Bug += Point_Bug13;
                                    AllServerPoint_東方 += Point_東方;
                                    AllServerPoint_松崗 += Point_松崗;

                                    textBox_SavePointDis.Text += "\r\n" + "Server : " + j + "\r\n" + "Total : " + Point_Total13.ToString("0") + "\r\n" + " Apple : " + Point_Apple13.ToString("0") + "\r\n" + " Google : " + Point_Google13.ToString("0") + "\r\n" + " GoogleFull : " + Point_Google13F.ToString("0") + "\r\n" + " MyCard : " + Point_MyCard13.ToString("0") + "\r\n" + " MCG : " + Point_MyCardGoogle13.ToString("0") + "\r\n" + " 980xCard : " + Point_980x13.ToString("0") + "\r\n" + " BugCard : " + Point_Bug13.ToString("0") + "\r\n" + " 老李卡 : " + (Point_東方 / 1.3).ToString("0") + "\r\n" + " 松崗 : " + (Point_松崗 / 1.3).ToString("0") + "\r\n";
                                }

                                textBox_SavePointDis.Text += "\r\n" + "全伺服器總合 : " + "\r\n" + "Total : " + AllServerTotal + "\r\n" + " Apple : " + AllServerPoint_Apple + "\r\n" + " Google : " + AllServerPoint_Google + "\r\n" + " GoogleFull : " + AllServerPoint_GoogleF + "\r\n" + " MyCard : " + AllServerPoint_MyCard + "\r\n" + " MCG : " + AllServerPoint_MyCardGoogle + "\r\n" + " 980xCard : " + AllServerPoint_980x + "\r\n" + " BugCard : " + AllServerPoint_Bug + "\r\n" + " 老李卡 : " + AllServerPoint_東方 + "\r\n" + " 松崗 : " + AllServerPoint_松崗 + "\r\n";


                                #endregion

                                #region 金錢流向



                                switch (comboBox_MoneyList.SelectedIndex)
                                {
                                    case 0:
                                        textBox_MoneyList.Text += "松崗金流 : " + "\r\n" + "\r\n";

                                        for (int x = 0; x < ds.Tables[3].Rows.Count; x++)
                                        {
                                            textBox_MoneyList.Text += "ID : " + ds.Tables[3].Rows[x]["UserID"].ToString() + "   儲值鑽石 : " + ds.Tables[3].Rows[x]["V2"].ToString() + "   時間 : " + Convert.ToDateTime(ds.Tables[3].Rows[x]["SaveDate"]).ToString("yyyy/MM/dd HH:mm:ss") + "\r\n";
                                        }

                                        break;

                                    case 1:
                                        textBox_MoneyList.Text += "吞食Q鑽金流 :  " + "\r\n" + "\r\n"; ;

                                        for (int x = 0; x < ds.Tables[3].Rows.Count; x++)
                                        {
                                            int MCG = 0;
                                            string MCGCome = "";

                                            switch (ds.Tables[3].Rows[x]["Cause"].ToString())
                                            {
                                                case "10":
                                                    MCGCome = "ＭｙＣａｒｄ　";
                                                    break;

                                                case "14":
                                                    MCGCome = "Ｇｏｏｇｌｅ　";
                                                    break;
                                            }

                                            switch (Convert.ToInt32(ds.Tables[3].Rows[x]["V2"]))
                                            {
                                                case 10:
                                                    MCG = 10;
                                                    break;

                                                case 53:
                                                    MCG = 50;
                                                    break;

                                                case 64:
                                                    MCG = 60;
                                                    break;

                                                case 168:
                                                    MCG = 150;
                                                    break;

                                                case 336:
                                                    MCG = 300;
                                                    break;

                                                case 403:
                                                    MCG = 350;
                                                    break;

                                                case 480:
                                                    MCG = 400;
                                                    break;

                                                case 540:
                                                    MCG = 450;
                                                    break;

                                                case 600:
                                                    MCG = 500;
                                                    break;

                                                case 958:
                                                    MCG = 750;
                                                    break;

                                                case 1270:
                                                    MCG = 1000;
                                                    break;

                                                case 1460:
                                                    MCG = 1150;
                                                    break;

                                                case 2600:
                                                    MCG = 2000;
                                                    break;

                                                case 4000:
                                                    MCG = 3000;
                                                    break;

                                                case 6750:
                                                    MCG = 5000;
                                                    break;

                                                default:
                                                    MCG = Convert.ToInt32(ds.Tables[3].Rows[x]["V2"]);

                                                    break;

                                            }

                                            textBox_MoneyList.Text += MCGCome + "ID : " + ds.Tables[3].Rows[x]["UserID"].ToString() + "   儲值鑽石 : " + ds.Tables[3].Rows[x]["V2"].ToString() + "   儲值台幣 : " + MCG + "   時間 : " + Convert.ToDateTime(ds.Tables[3].Rows[x]["SaveDate"]).ToString("yyyy/MM/dd HH:mm:ss") + "\r\n";
                                        }
                                        break;
                                }

                                #endregion

                            }

                            catch
                            {
                                MessageBox.Show("資料查詢錯誤");
                            }

                            finally
                            {
                            }
                        }


                        else
                        {
                            MessageBox.Show("密碼錯誤");
                        }


                        if (textBox_AnalysisPassword.Text == SendPassWord)
                        {
                            string Tsql = "SELECT TOP (1) Area, Account, DistinctAccount, GS1, GS1Max, GS2, GS2Max, GS3, GS3Max, GS4, GS4Max, GS5, GS5Max, GS6, GS6Max, SaveDAte FROM OnlineRole WHERE (Area = 'HK') AND (SaveDAte > CONVERT(DATETIME, '" + AnalysisPickTime + " 00:00:00', 000)) AND (SaveDAte < CONVERT(DATETIME, '" + AnalysisPickTime + " 23:59:59', 000)) ORDER BY SaveDAte DESC";

                            try
                            {
                                DataSet ds = new DataSet();
                                ds = TWsql.Get_SQL_Analysis_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                                textBox_LoginDis.Text += "國家 : 香港" + "\r\n" + "\r\n" + "總帳號數 : " + ds.Tables[0].Rows[0]["Account"].ToString() + "\r\n" + "\r\n" + "不重複登入數 : " + ds.Tables[0].Rows[0]["DistinctAccount"].ToString() + "\r\n" + "\r\n" + "\r\n";
                                textBox_LoginDis.Text += "即時人數 : " + "\r\n" + "\r\n" + "Server 1 :  " + ds.Tables[0].Rows[0]["GS1"].ToString() + "\r\n" + "Server 2 :  " + ds.Tables[0].Rows[0]["GS2"].ToString() + "\r\n" + "Server 3 :  " + ds.Tables[0].Rows[0]["GS3"].ToString() + "\r\n" + "Server 4 :  " + ds.Tables[0].Rows[0]["GS4"].ToString() + "\r\n" + "Server 5 :  " + ds.Tables[0].Rows[0]["GS5"].ToString() + "\r\n" + "Server 6 :  " + ds.Tables[0].Rows[0]["GS6"].ToString() + "\r\n" + "Server 7 :  " + ds.Tables[0].Rows[0]["GS7"].ToString() + "\r\n" + "Server 8 :  " + ds.Tables[0].Rows[0]["GS8"].ToString() + "\r\n" + "\r\n" + "\r\n";
                                textBox_LoginDis.Text += "最高人數 : " + "\r\n" + "\r\n" + "Server 1 :  " + ds.Tables[0].Rows[0]["GS1Max"].ToString() + "\r\n" + "Server 2 :  " + ds.Tables[0].Rows[0]["GS2Max"].ToString() + "\r\n" + "Server 3  : " + ds.Tables[0].Rows[0]["GS3Max"].ToString() + "\r\n" + "Server 4 :  " + ds.Tables[0].Rows[0]["GS4Max"].ToString() + "\r\n" + "Server 5 :  " + ds.Tables[0].Rows[0]["GS5Max"].ToString() + "\r\n" + "Server 6 :  " + ds.Tables[0].Rows[0]["GS6ax"].ToString() + "\r\n" + "Server 7 :  " + ds.Tables[0].Rows[0]["GS7ax"].ToString() + "\r\n" + "Server 8 :  " + ds.Tables[0].Rows[0]["GS8ax"].ToString() + "\r\n" + "\r\n";
                                MessageBox.Show("資料來了");
                            }

                            catch
                            {
                                MessageBox.Show("資料查詢錯誤");
                            }
                        }
                        else
                        {
                        }

                        break;

                    case 3:

                        if (textBox_AnalysisPassword.Text == SendPassWord)
                        {
                            string Tsql = "SELECT Kind, Cause, V2 FROM Money_" + AnalysisPickTime + " WHERE (Kind = 2);";
                            Tsql += "SELECT UserID, COUNT(PuzzleWeb) AS Expr1 FROM Money_" + AnalysisPickTime + " WHERE (Kind = 2) GROUP BY UserID;";
                            Tsql += "SELECT Kind, ServerID, Cause, V2 ,V4  FROM Money_" + AnalysisPickTime + " WHERE (Kind = 1);";
                            Tsql += "SELECT Kind, Cause, UserID, V2, V4, SaveDate FROM Money_" + AnalysisPickTime + " WHERE " + MoneySql + " ORDER BY SaveDate";

                            try
                            {
                                DataSet ds = new DataSet();
                                ds = CNsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);

                                #region 耗點分析

                                for (int k = 0; k < ds.Tables[0].Rows.Count; k++)
                                {
                                    switch (ds.Tables[0].Rows[k]["Cause"].ToString())
                                    {
                                        case "1":
                                            // = "購買能量　　　　　";
                                            Role[1]++;
                                            Counts[1] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            break;
                                        case "2":
                                            // = "抽武將　　　　　　";
                                            Counts[2] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[2]++;
                                            break;
                                        case "3":
                                            // = "開寶箱　　　　　　";
                                            Counts[3] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[3]++;
                                            break;
                                        case "4":
                                            // = "消除競技場ＣＤ　　";
                                            Counts[4] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[4]++;
                                            break;
                                        case "5":
                                            // = "重置攻城戰　　　　";
                                            Counts[5] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[5]++;
                                            break;
                                        case "6":
                                            // = "重置攻城寶箱　　　";
                                            Counts[6] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[6]++;
                                            break;
                                        case "7":
                                            // = "購買懸賞格子　　　";
                                            Counts[7] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[7]++;
                                            break;
                                        case "8":
                                            // = "轉蛋扣除　　　　　";
                                            Counts[8] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[8]++;
                                            break;
                                        case "9":
                                            // = "購買武將欄位　　　";
                                            Counts[9] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[9]++;
                                            break;
                                        case "10":
                                            // = "推圖復活　　　　　";
                                            Counts[10] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[10]++;
                                            break;
                                        case "11":
                                            // = "戰場使用精英援軍　";
                                            Counts[11] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[11]++;
                                            break;
                                        case "12":
                                            // = "快速通關　　　　　";
                                            Counts[12] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[12]++;
                                            break;
                                        case "13":
                                            Counts[13] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[13]++;
                                            // = "購買ＰＶＰ挑戰次數";
                                            break;
                                        case "16":
                                            Counts[14] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[14]++;
                                            // = "轉盤重置　　　　　";
                                            break;
                                        case "17":
                                            Counts[15] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[15]++;
                                            // = "購買裝備欄位　　 　";
                                            break;
                                        case "18":
                                            if (Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString()) == 30)
                                            {
                                                Counts[16] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                                Role[16]++;
                                                // = "ＶＩＰ強化－３０點 ";
                                            }

                                            if (Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString()) == 100)
                                            {
                                                Counts[17] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                                Role[17]++;
                                                //Cause = "ＶＩＰ強化－１００點";
                                            }
                                            break;
                                        case "14":
                                            Counts[18] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[18]++;
                                            // = "消世界ＢｏｓｓＣＤ　";
                                            break;
                                        case "15":
                                            Counts[19] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[19]++;
                                            // = "消掃蕩ＣＤ　　　　　";
                                            break;
                                        case "19":
                                            Counts[20] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[20]++;
                                            // = "消跨服ＰＶＰ　ＣＤ　"
                                            break;
                                        case "20":
                                            Counts[21] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[21]++;
                                            // = "買跨服ＰＶＰ挑戰次數";
                                            break;
                                        case "21":
                                            Counts[22] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[22]++;
                                            // = "建立軍團　　　　　　";
                                            break;
                                        case "22":
                                            Counts[23] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[23]++;
                                            // = "軍團捐獻　　　　　　";
                                            break;
                                        case "23":
                                            Counts[24] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[25]++;
                                            // = "接關扣鑽　　　　　　";
                                            break;
                                        case "24":
                                            Counts[25] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[25]++;
                                            // = "購買累積儲值福袋　　";
                                            break;
                                        case "25":
                                            Counts[26] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[26]++;
                                            // = "購買勢力禮盒　　　　";
                                            break;
                                        case "26":
                                            Counts[27] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[27]++;
                                            // = "買全球Ｂｏｓｓ挑戰次數";
                                            break;
                                        case "27":
                                            Counts[28] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[28]++;
                                            // = "全球Ｂｏｓｓ升級　　";
                                            break;
                                        case "28":
                                            Counts[29] += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());
                                            Role[29]++;
                                            // = "ＷＥＢ後台　　　　　";
                                            break;
                                    }

                                    Total_CON += Convert.ToInt32(ds.Tables[0].Rows[k]["V2"].ToString());

                                }

                                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                                {
                                    AccountNum++;

                                }
                                textBox_AnalysisDis.Text += "耗點總額度　　　　　 : " + Total_CON.ToString() + "\r\n" + "\r\n" + "購買能量　　　　　　 : " + Counts[1].ToString("00000000") + "  /  " + Role[1].ToString("0000") + "\r\n" + "抽武將　　　　　　　 : " + Counts[2].ToString("00000000") + "  /  " + Role[2].ToString("0000") + "\r\n" + "開寶箱　　　　　　　 : " + Counts[3].ToString("00000000") + "  /  " + Role[3].ToString("0000") + "\r\n" + "消除ＰＶＰ　ＣＤ　　 : " + Counts[4].ToString("00000000") + "  /  " + Role[4].ToString("0000") + "\r\n" + "重置攻城戰　　　　　 : " + Counts[5].ToString("00000000") + "  /  " + Role[5].ToString("0000") + "\r\n" + "重置攻城寶箱　　　　 : " + Counts[6].ToString("00000000") + "  /  " + Role[6].ToString("0000") + "\r\n" + "購買懸賞格子　　　　 : " + Counts[7].ToString("00000000") + "  /  " + Role[7].ToString("0000") + "\r\n" + "轉蛋　　　　　　　　 : " + Counts[8].ToString("00000000") + "  /  " + Role[8].ToString("0000") + "\r\n" + "購買武將欄位　　　　 : " + Counts[9].ToString("00000000") + "  /  " + Role[9].ToString("0000") + "\r\n" + "推圖復活　　　　　　 : " + Counts[10].ToString("00000000") + "  /  " + Role[10].ToString("0000") + "\r\n" + "戰場使用精英援軍　　 : " + Counts[11].ToString("00000000") + "  /  " + Role[11].ToString("0000") + "\r\n" + "快速通關　　　　　　 : " + Counts[12].ToString("00000000") + "  /  " + Role[12].ToString("0000") + "\r\n" + "購買ＰＶＰ挑戰次數　 : " + Counts[13].ToString("00000000") + "  /  " + Role[13].ToString("0000") + "\r\n" + "拉ＢＡＲ重置　　　　 : " + Counts[14].ToString("00000000") + "  /  " + Role[14].ToString("0000") + "\r\n" + "購買裝備欄位　　　　 : " + Counts[15].ToString("00000000") + "  /  " + Role[15].ToString("0000") + "\r\n" + "ＶＩＰ強化－３０點　 : " + Counts[16].ToString("00000000") + "  /  " + Role[16].ToString("0000") + "\r\n" + "ＶＩＰ強化－１００點 : " + Counts[17].ToString("00000000") + "  /  " + Role[17].ToString("0000") + "\r\n" + "消世界ＢｏｓｓＣＤ　 : " + Counts[18].ToString("00000000") + "  /  " + Role[18].ToString("0000") + "\r\n" + "消掃蕩ＣＤ　　　　     : " + Counts[19].ToString("00000000") + "  /  " + Role[19].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "消跨服ＰＶＰ　ＣＤ　 : " + Counts[20].ToString("00000000") + "  /  " + Role[20].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "買跨服ＰＶＰ挑戰次數 : " + Counts[21].ToString("00000000") + "  /  " + Role[21].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "建立軍團　　　　 　　: " + Counts[22].ToString("00000000") + "  /  " + Role[22].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "軍團捐獻　　　　 　　: " + Counts[23].ToString("00000000") + "  /  " + Role[23].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "接關扣鑽　　　　 　　: " + Counts[24].ToString("00000000") + "  /  " + Role[24].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "購買累積儲值福袋　　 : " + Counts[25].ToString("00000000") + "  /  " + Role[25].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "購買勢力禮盒　　　　 : " + Counts[26].ToString("00000000") + "  /  " + Role[26].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "買全球Ｂｏｓｓ挑戰次數: " + Counts[27].ToString("00000000") + "  /  " + Role[27].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "全球Ｂｏｓｓ升級　　 : " + Counts[28].ToString("00000000") + "  /  " + Role[28].ToString("0000") + "\r\n";
                                textBox_AnalysisDis.Text += "ＷＥＢ後台　　　　　 : " + Counts[29].ToString("00000000") + "  /  " + Role[29].ToString("0000");
                                textBox_AnalysisDis.Text += "\r\n" + "\r\n" + "耗點帳號數量 : " + AccountNum;

                                #endregion

                                #region 儲點分析

                                int ServerNum = 5; //伺服器總數
                                int AllServerTotal = 0;
                                int AllServerPoint_Apple = 0;
                                int AllServerPoint_Google = 0;
                                int AllServerPoint_GoogleF = 0;
                                int AllServerPoint_MyCard = 0;
                                int AllServerPoint_MyCardGoogle = 0;
                                int AllServerPoint_980x = 0;
                                int AllServerPoint_Bug = 0;
                                int AllServerPoint_東方 = 0;

                                textBox_SavePointDis.Text += "台幣資訊 : " + "\r\n";

                                for (int j = 1; j <= ServerNum; j++)
                                {
                                    int Point_Apple13 = 0;
                                    int Point_Google13 = 0;
                                    int Point_Google13F = 0;
                                    int Point_MyCard13 = 0;
                                    int Point_MyCardGoogle13 = 0;
                                    int Point_980x13 = 0;
                                    int Point_Bug13 = 0;
                                    int Point_東方 = 0;
                                    int Point_Total13 = 0;

                                    for (int i = 0; i < ds.Tables[2].Rows.Count; i++)
                                    {
                                        if (Convert.ToInt32(ds.Tables[2].Rows[i]["ServerID"].ToString()) == j)
                                        {
                                            switch (ds.Tables[2].Rows[i]["Cause"].ToString())
                                            {
                                                #region 流向分類

                                                case "9":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "20":

                                                            Point_Bug13 += 20;

                                                            break;

                                                        case "30":

                                                            Point_Bug13 += 30;

                                                            break;

                                                        case "53":

                                                            Point_Bug13 += 50;

                                                            break;

                                                        case "168":

                                                            Point_Bug13 += 150;

                                                            break;

                                                        case "480":

                                                            Point_Bug13 += 400;

                                                            break;

                                                        case "1000":

                                                            Point_Bug13 += 800;

                                                            break;

                                                        case "1560":

                                                            Point_Bug13 += 1200;

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());

                                                            Point_Bug13 += Convert.ToInt32(Math.Floor(x / 1.3));

                                                            break;
                                                    }

                                                    break;

                                                case "10":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "53":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 50;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 50;
                                                            }

                                                            break;

                                                        case "168":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 150;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 150;
                                                            }

                                                            break;

                                                        case "336":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 300;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 300;
                                                            }

                                                            break;

                                                        case "403":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 350;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 350;
                                                            }

                                                            break;

                                                        case "480":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 400;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 400;
                                                            }

                                                            break;

                                                        case "540":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 450;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 450;
                                                            }

                                                            break;

                                                        case "600":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 500;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 500;
                                                            }

                                                            break;

                                                        case "1270":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 1000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 1000;
                                                            }

                                                            break;

                                                        case "1460":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 1150;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 1150;
                                                            }

                                                            break;

                                                        case "2600":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 2000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 2000;
                                                            }

                                                            break;

                                                        case "4000":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 3000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 3000;
                                                            }

                                                            break;

                                                        case "6750":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_MyCardGoogle13 += 5000;
                                                            }

                                                            else
                                                            {
                                                                Point_MyCard13 += 5000;
                                                            }

                                                            break;

                                                    }
                                                    break;

                                                case "13":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "64":

                                                            Point_Apple13 += 60;

                                                            break;

                                                        case "336":

                                                            Point_Apple13 += 300;

                                                            break;

                                                        case "540":

                                                            Point_Apple13 += 450;

                                                            break;

                                                        case "958":

                                                            Point_Apple13 += 750;

                                                            break;

                                                        case "4000":

                                                            Point_Apple13 += 2990;

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());

                                                            Point_Apple13 += Convert.ToInt32(Math.Floor(x / 1.3));

                                                            break;

                                                    }
                                                    break;
                                                case "14":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "64":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 60;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 60;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 60;
                                                            }

                                                            break;

                                                        case "336":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 300;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 300;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 300;
                                                            }

                                                            break;

                                                        case "540":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 450;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 450;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 450;
                                                            }

                                                            break;

                                                        case "958":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 750;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 750;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 750;
                                                            }

                                                            break;

                                                        case "4000":

                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += 2990;
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += 2990;
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += 2990;
                                                            }

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());


                                                            if (ds.Tables[2].Rows[i]["V4"].ToString() == "11")
                                                            {
                                                                Point_Google13 += Convert.ToInt32(Math.Floor(x / 1.3));
                                                            }

                                                            else if (ds.Tables[2].Rows[i]["V4"].ToString() == "17")
                                                            {
                                                                Point_Google13F += Convert.ToInt32(Math.Floor(x / 1.3));
                                                            }

                                                            else
                                                            {
                                                                Point_Google13 += Convert.ToInt32(Math.Floor(x / 1.3));
                                                            }

                                                            break;
                                                    }
                                                    break;
                                                case "16":

                                                    switch (ds.Tables[2].Rows[i]["V2"].ToString())
                                                    {
                                                        case "168":

                                                            Point_980x13 += 150;

                                                            break;

                                                        case "480":

                                                            Point_980x13 += 400;

                                                            break;

                                                        case "1560":

                                                            Point_980x13 += 1200;

                                                            break;

                                                        default:

                                                            int x = Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());

                                                            Point_980x13 += Convert.ToInt32(Math.Floor(x / 1.3));

                                                            break;
                                                    }

                                                    break;

                                                case "24":
                                                    Point_東方 += Convert.ToInt32(ds.Tables[2].Rows[i]["V2"].ToString());
                                                    break;


                                                #endregion
                                            }
                                        }
                                    }

                                    Point_Total13 = Convert.ToInt32(Math.Floor(Point_Apple13 + Point_Google13 + Point_Google13F + Point_MyCard13 + Point_MyCardGoogle13 + Point_980x13 + Point_Bug13 + (Point_東方 / 1.3)));

                                    AllServerTotal += Point_Total13;
                                    AllServerPoint_Apple += Point_Apple13;
                                    AllServerPoint_Google += Point_Google13;
                                    AllServerPoint_GoogleF += Point_Google13F;
                                    AllServerPoint_MyCard += Point_MyCard13;
                                    AllServerPoint_MyCardGoogle += Point_MyCardGoogle13;
                                    AllServerPoint_980x += Point_980x13;
                                    AllServerPoint_Bug += Point_Bug13;
                                    AllServerPoint_東方 += Point_東方;

                                    textBox_SavePointDis.Text += "\r\n" + "Server : " + j + "\r\n" + "Total : " + Point_Total13.ToString("0") + "\r\n" + " Apple : " + Point_Apple13.ToString("0") + "\r\n" + " Google : " + Point_Google13.ToString("0") + "\r\n" + " GoogleFull : " + Point_Google13F.ToString("0") + "\r\n" + " MyCard : " + Point_MyCard13.ToString("0") + "\r\n" + " MCG : " + Point_MyCardGoogle13.ToString("0") + "\r\n" + " 980xCard : " + Point_980x13.ToString("0") + "\r\n" + " BugCard : " + Point_Bug13.ToString("0") + "\r\n" + " 老李卡 : " + (Point_東方 / 1.3).ToString("0") + "\r\n";
                                }

                                textBox_SavePointDis.Text += "\r\n" + "全伺服器總合 : " + "\r\n" + "Total : " + AllServerTotal + "\r\n" + " Apple : " + AllServerPoint_Apple + "\r\n" + " Google : " + AllServerPoint_Google + "\r\n" + " GoogleFull : " + AllServerPoint_GoogleF + "\r\n" + " MyCard : " + AllServerPoint_MyCard + "\r\n" + " MCG : " + AllServerPoint_MyCardGoogle + "\r\n" + " 980xCard : " + AllServerPoint_980x + "\r\n" + " BugCard : " + AllServerPoint_Bug + "\r\n" + " 老李卡 : " + AllServerPoint_東方 + "\r\n";


                                #endregion

                                #region 金錢流向



                                switch (comboBox_MoneyList.SelectedIndex)
                                {
                                    case 0:
                                        textBox_MoneyList.Text += "松崗金流 : " + "\r\n" + "\r\n";

                                        for (int x = 0; x < ds.Tables[3].Rows.Count; x++)
                                        {
                                            textBox_MoneyList.Text += "ID : " + ds.Tables[3].Rows[x]["UserID"].ToString() + "   儲值鑽石 : " + ds.Tables[3].Rows[x]["V2"].ToString() + "   時間 : " + Convert.ToDateTime(ds.Tables[3].Rows[x]["SaveDate"]).ToString("yyyy/MM/dd HH:mm:ss") + "\r\n";
                                        }

                                        break;

                                    case 1:
                                        textBox_MoneyList.Text += "吞食Q鑽金流 :  " + "\r\n" + "\r\n"; ;

                                        for (int x = 0; x < ds.Tables[3].Rows.Count; x++)
                                        {
                                            int MCG = 0;
                                            string MCGCome = "";

                                            switch (ds.Tables[3].Rows[x]["Cause"].ToString())
                                            {
                                                case "10":
                                                    MCGCome = "ＭｙＣａｒｄ　";
                                                    break;

                                                case "14":
                                                    MCGCome = "Ｇｏｏｇｌｅ　";
                                                    break;
                                            }

                                            switch (Convert.ToInt32(ds.Tables[3].Rows[x]["V2"]))
                                            {
                                                case 10:
                                                    MCG = 10;
                                                    break;

                                                case 53:
                                                    MCG = 50;
                                                    break;

                                                case 64:
                                                    MCG = 60;
                                                    break;

                                                case 168:
                                                    MCG = 150;
                                                    break;

                                                case 336:
                                                    MCG = 300;
                                                    break;

                                                case 403:
                                                    MCG = 350;
                                                    break;

                                                case 480:
                                                    MCG = 400;
                                                    break;

                                                case 540:
                                                    MCG = 450;
                                                    break;

                                                case 600:
                                                    MCG = 500;
                                                    break;

                                                case 958:
                                                    MCG = 750;
                                                    break;

                                                case 1270:
                                                    MCG = 1000;
                                                    break;

                                                case 1460:
                                                    MCG = 1150;
                                                    break;

                                                case 2600:
                                                    MCG = 2000;
                                                    break;

                                                case 4000:
                                                    MCG = 3000;
                                                    break;

                                                case 6750:
                                                    MCG = 5000;
                                                    break;

                                                default:
                                                    MCG = Convert.ToInt32(ds.Tables[3].Rows[x]["V2"]);

                                                    break;

                                            }

                                            textBox_MoneyList.Text += MCGCome + "ID : " + ds.Tables[3].Rows[x]["UserID"].ToString() + "   儲值鑽石 : " + ds.Tables[3].Rows[x]["V2"].ToString() + "   儲值台幣 : " + MCG + "   時間 : " + Convert.ToDateTime(ds.Tables[3].Rows[x]["SaveDate"]).ToString("yyyy/MM/dd HH:mm:ss") + "\r\n";
                                        }
                                        break;
                                }

                                #endregion

                            }

                            catch
                            {
                                MessageBox.Show("資料查詢錯誤");
                            }

                            finally
                            {
                            }
                        }


                        else
                        {
                            MessageBox.Show("密碼錯誤");
                        }



                        if (textBox_AnalysisPassword.Text == SendPassWord)
                        {
                            string Tsql = "SELECT TOP (1) Area, Account, DistinctAccount, GS1, GS1Max, GS2, GS2Max, GS3, GS3Max, GS4, GS4Max, GS5, GS5Max, SaveDAte FROM OnlineRole WHERE (Area = 'CN') AND (SaveDAte > CONVERT(DATETIME, '" + AnalysisPickTime + "  00:00:00', 000)) AND (SaveDAte < CONVERT(DATETIME, '" + AnalysisPickTime + " 23:59:59', 000)) ORDER BY SaveDAte DESC";

                            try
                            {
                                DataSet ds = new DataSet();
                                ds = TWsql.Get_SQL_Analysis_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);

                                textBox_LoginDis.Text += "國家 : 中國" + "\r\n" + "\r\n" + "總帳號數 : " + ds.Tables[0].Rows[0]["Account"].ToString() + "\r\n" + "\r\n" + "不重複登入數 : " + ds.Tables[0].Rows[0]["DistinctAccount"].ToString() + "\r\n" + "\r\n" + "\r\n";
                                textBox_LoginDis.Text += "即時人數 : " + "\r\n" + "\r\n" + "Server 1 :  " + ds.Tables[0].Rows[0]["GS1"].ToString() + "\r\n" + "Server 2 :  " + ds.Tables[0].Rows[0]["GS2"].ToString() + "\r\n" + "Server 3 :  " + ds.Tables[0].Rows[0]["GS3"].ToString() + "\r\n" + "Server 4 :  " + ds.Tables[0].Rows[0]["GS4"].ToString() + "\r\n" + "Server 5 :  " + ds.Tables[0].Rows[0]["GS5"].ToString() + "\r\n" + "\r\n" + "\r\n";
                                textBox_LoginDis.Text += "最高人數 : " + "\r\n" + "\r\n" + "Server 1 :  " + ds.Tables[0].Rows[0]["GS1Max"].ToString() + "\r\n" + "Server 2 :  " + ds.Tables[0].Rows[0]["GS2Max"].ToString() + "\r\n" + "Server 3  : " + ds.Tables[0].Rows[0]["GS3Max"].ToString() + "\r\n" + "Server 4 :  " + ds.Tables[0].Rows[0]["GS4Max"].ToString() + "\r\n" + "Server 5 :  " + ds.Tables[0].Rows[0]["GS5Max"].ToString() + "\r\n" + "\r\n";
                                MessageBox.Show("資料來了");
                            }

                            catch
                            {
                                MessageBox.Show("資料查詢錯誤");
                            }

                            finally
                            {
                            }
                        }

                        break;

                }


            }

            if (textBox_AnalysisDis.Text == string.Empty)
            {
                textBox_AnalysisDis.Text = "沒有資料";
            }

            if (textBox_LoginDis.Text == string.Empty)
            {
                textBox_LoginDis.Text = "沒有資料";
            }

            if (textBox_SavePointDis.Text == string.Empty)
            {
                textBox_SavePointDis.Text = "沒有資料";
            }

            if (textBox_MoneyList.Text == string.Empty)
            {
                textBox_MoneyList.Text = "沒有資料";
            }

        #endregion

        }

        #region 新增推播訊息

        int GiftMGCountry = 0; //選擇國家

        private void button_MessageInquir_Click(object sender, EventArgs e)
        {
            string GiftMGNum = "";

            switch (comboBox_GiftMGNum.SelectedIndex)
            {
                case 0:
                    GiftMGNum = "TOP (10)";
                    break;

                case 1:
                    GiftMGNum = "TOP (20)";
                    break;

                case 2:
                    GiftMGNum = "TOP (30)";
                    break;

                case 3:
                    GiftMGNum = "";
                    break;
            }

            string Tsql = "SELECT " + GiftMGNum + " EventID, EventName, ItemNumber, ItemType, Count, Creater, CreateDate, StartDate, EndDate FROM Event ORDER BY EventID DESC";

            try
            {
                DataSet ds = new DataSet();

                switch (GiftMGCountry)
                {
                    case 0:
                        ds = TWPuzzle.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        label50.Text = "資料顯示國家 : 台灣";
                        break;

                    case 1:
                        ds = HKPuzzle.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        label50.Text = "資料顯示國家 : 香港";
                        break;

                    case 2:
                        ds = CNsql.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        label50.Text = "資料顯示國家 : 中國";
                        break;

                    case 3:
                        ds = SGPuzzle.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        label50.Text = "資料顯示國家 : 新加坡";
                        break;

                    case 4:
                        ds = MAPuzzle.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        label50.Text = "資料顯示國家 : 馬來西亞";
                        break;

                    case 5:
                        ds = THPuzzle.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        label50.Text = "資料顯示國家 : 泰國";
                        break;

                    case 6:
                        ds = KRPuzzle.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        label50.Text = "資料顯示國家 : 韓國";
                        break;

                }

                dataGridView_MessageList.DataSource = ds.Tables[0];
            }
            catch
            {
                MessageBox.Show("查詢失敗");
            }
            finally
            {
                MessageBox.Show("資料來了");
            }

        }

        private void comboBox1_SelectedIndexChanged_3(object sender, EventArgs e)
        {

        }

        private void comboBox_GiftMGCountry_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_GiftMGCountry.SelectedIndex)
            {
                case 0:
                    GiftMGCountry = 0;
                    break;

                case 1:
                    GiftMGCountry = 1;
                    break;

                case 2:
                    GiftMGCountry = 2;
                    break;

                case 3:
                    GiftMGCountry = 3;
                    break;

                case 4:
                    GiftMGCountry = 4;
                    break;

                case 5:
                    GiftMGCountry = 5;
                    break;

                case 6:
                    GiftMGCountry = 6;
                    break;

            }
        }

        private void button_giftMGSend_Click(object sender, EventArgs e)
        {
            if (textBox_giftMGPassword.Text != SendPassWord )
            {
                MessageBox.Show("密碼錯誤");
            }

            else if (textBox_giftMGCreater.Text == string.Empty)
            {
                MessageBox.Show("Creater不可為空");
            }

            else if (TextBox_giftMGEventName.Text == string.Empty)
            {
                MessageBox.Show("EventName不可為空");
            }

            else
            {
                if (checkBox_Continu.Checked)
                {

                }
                else
                {
                    textBox_giftMGPassword.Text = "";
                }

                try
                {
                    int xr = 0;

                    string Tsql = "INSERT INTO Event( EventName, ItemNumber, ItemType, Count, Creater, CreateDate, StartDate, EndDate) ";
                    Tsql += "VALUES ( N'" + TextBox_giftMGEventName.Text + "','0','0','0','" + textBox_giftMGCreater.Text + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                   
                    switch (GiftMGCountry)
                    {
                        case 0:
                            xr = TWPuzzle.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;

                        case 1:
                            xr = HKPuzzle.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;

                        case 2:
                            xr = CNsql.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;

                        case 3:
                            xr = SGPuzzle.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;

                        case 4:
                            xr = MAPuzzle.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;

                        case 5:
                            xr = THPuzzle.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;

                        case 6:
                            xr = KRPuzzle.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;
                    }

                }
                catch
                {
                    MessageBox.Show("寫入失敗");
                }
                finally
                {
                    button_MessageInquir.PerformClick();
                }

            }

        }

        #endregion

        int FBMGCountry = 0;//選擇國家

        private void comboBox_FBMGCountry_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_FBMGCountry.SelectedIndex)
            {
                case 0:
                    FBMGCountry = 0;
                    break;

                case 1:
                    FBMGCountry = 1;
                    break;

                case 2:
                    FBMGCountry = 2;
                    break;

            }
        }

        private void button_MessageInquirFB_Click(object sender, EventArgs e)
        {
            string FBMGNum = "";

            switch (comboBox_FBMGNum.SelectedIndex)
            {
                case 0:
                    FBMGNum = "TOP (10)";
                    break;

                case 1:
                    FBMGNum = "TOP (20)";
                    break;

                case 2:
                    FBMGNum = "TOP (30)";
                    break;

                case 3:
                    FBMGNum = "";
                    break;
            }

            string Tsql = "SELECT " + FBMGNum + " MessageID, Title, TitleLink, Body, PhotoLink, Message FROM Facebook_Message ORDER BY MessageID DESC";

            try
            {
                DataSet ds = new DataSet();

                switch (FBMGCountry)
                {
                    case 0:
                        ds = TWPuzzle.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        label51.Text = "資料顯示國家 : 台灣";
                        break;

                    case 1:
                        ds = HKPuzzle.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        label51.Text = "資料顯示國家 : 香港";
                        break;

                    case 2:
                        ds = SGPuzzle.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        label51.Text = "資料顯示國家 : 新加坡";
                        break;

                }

                dataGridView_MessageListFB.DataSource = ds.Tables[0];
            }
            catch
            {
                MessageBox.Show("查詢失敗");
            }
            finally
            {
                MessageBox.Show("資料來了");
            }
        }

        private void button_FBMGSend_Click(object sender, EventArgs e)
        {
            if (textBox_FBPassWord.Text != SendPassWord)
            {
                MessageBox.Show("密碼錯誤");
            }

            else if (textBox_FBTitle.Text == string.Empty)
            {
                MessageBox.Show("Title不可為空");
            }

            else if (textBox_FBTitleLink.Text == string.Empty)
            {
                MessageBox.Show("TitleLink不可為空");
            }

            else if (textBox_FBBody.Text == string.Empty)
            {
                MessageBox.Show("Body不可為空");
            }

            else if (textBox_FBPhotoLink.Text == string.Empty)
            {
                MessageBox.Show("PhotoLink不可為空");
            }

            else if (textBox_FBMessage.Text == string.Empty)
            {
                MessageBox.Show("Message不可為空");
            }

            else
            {
                if (checkBox_ContinuFB.Checked)
                {

                }
                else
                {
                    textBox_FBPassWord.Text = "";
                }
                
                try
                {
                    int xr = 0;
                    string Tsql = "";
                    switch (comboBox_FBWorC.SelectedIndex)
                    {
                        case 0:
                            Tsql = "INSERT INTO Facebook_Message ( Title, TitleLink, Body, PhotoLink, Message) ";
                            Tsql += "VALUES ('" + textBox_FBTitle.Text + "','" + textBox_FBTitleLink.Text + "','" + textBox_FBBody.Text + "','" + textBox_FBPhotoLink.Text + "','" + textBox_FBMessage.Text + "')";
                            break;

                        case 1:
                            Tsql = "UPDATE    Facebook_Message ";
                            Tsql += "SET Title ='" + textBox_FBTitle.Text + "', TitleLink ='" + textBox_FBTitleLink.Text + "', Body ='" + textBox_FBBody.Text + "', PhotoLink ='" + textBox_FBPhotoLink.Text + "', Message ='" + textBox_FBMessage.Text + "' WHERE (MessageID = '" + textBox_FBWorC.Text + "')";
                            break;

                    }

                    switch (FBMGCountry)
                    {
                        case 0:
                            xr = TWPuzzle.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;

                        case 1:
                            xr = HKPuzzle.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;

                        case 2:
                            xr = SGPuzzle.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;
                    }

                }
                catch
                {
                    MessageBox.Show("寫入失敗");
                }
                finally
                {
                    button_MessageInquirFB.PerformClick();
                }

            }

        }

        private void comboBox_WorC_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_FBWorC.SelectedIndex)
            {
                case 0:
                    button_FBMGSend.Text = "創建訊息";
                    textBox_FBWorC.Enabled = false;
                    textBox_FBWorC.Text = "0";
                    break;

                case 1:
                    button_FBMGSend.Text = "修改訊息";
                    textBox_FBWorC.Enabled = true;
                    break;

            }
        }

        private void tabPage7_Click(object sender, EventArgs e)
        {

        }

        private void comboBox_GiftWorC_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox_MoneyList_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_MoneyList.SelectedIndex)
            {
                case 0:
                    MoneySql = "(Kind = '1') AND (Cause = '15')";
                    break;

                case 1:
                    MoneySql = " (Kind = '1') AND (V4 = '17') AND (Cause = '10' OR Cause = '14' OR Cause = '22')";
                    break;

            }
        }

        
        void NewForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            button_OpenBTG.Enabled = true;
        }

        string PickTime = "yyyyMMdd"; //選擇的日期
        int RCountry = 0; //選擇的國家
        string Description = "";
        int SportPoint = 0; //積分

        private void comboBox_RCountry_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_RCountry.SelectedIndex)
            {
                case 0:

                    RCountry = 1;

                    break;

                case 1:

                    RCountry = 2;

                    break;

                case 2:

                    RCountry = 3;

                    break;

            }
        }

        private void dateTimePicker_Date_ValueChanged(object sender, EventArgs e)
        {
            PickTime = dateTimePicker_Date.Value.ToString("yyyyMMdd");
        }

        private void button_Analysis_Click(object sender, EventArgs e)
        {
            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > Convert.ToInt32(DateTime.Now.ToString("yyyyMMdd")))
            {
                MessageBox.Show("此日期還沒有分析資料");
            }

            else if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130828)
            {
                label_inquire.Text = "目前查詢日期:  " + dateTimePicker_Date.Value.ToString("yyyy/MM/dd"); //日期顯示

                switch (RCountry)
                {
                    case 1:

                        label_Country.Text = "目前查詢國家:  台灣";

                        try
                        {
                            //玩家等級分析
                            string Tsql = "SELECT Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 1) ORDER BY SaveDate DESC, CAST(V1 AS int), CAST(V2 AS int);";

                            //玩家名聲分析
                            Tsql += "SELECT  Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 2) AND (V1 = 1 OR V1 = 2 OR V1 = 3 OR V1 = 4 OR V1 = 5 OR V1 = 6 OR V1 = 7 OR V1 = 8 OR V1 = 9 OR V1 = 10) ORDER BY CAST(V1 AS int), SaveDate DESC, CAST(V2 AS int);";

                            //Vip等級
                            Tsql += "SELECT  Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE  (Kind = 75) AND (Cause = 3)  ORDER BY CAST(V1 AS int), SaveDate DESC,CAST(V2 AS int);";

                            //玩家現有鑽石數
                            Tsql += "SELECT Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 4) AND (V1 = 1 OR V1 = 2 OR V1 = 3 OR V1 = 4 OR V1 = 5 OR V1 = 6 OR V1 = 7 OR V1 = 8 OR V1 = 9 OR V1 = 10) ORDER BY CAST(V1 AS int), SaveDate DESC, CAST(V2 AS int);";

                            //登入方式
                            Tsql += "SELECT Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 5) ORDER BY SaveDate DESC,CAST(V2 AS int);";

                            //卡包分析
                            Tsql += "SELECT Cause, COUNT(UserID) AS Expr1 FROM Egg_" + PickTime + " WHERE (Kind = 78) AND (Cause = 1 OR Cause = 2 OR Cause = 3 OR Cause = 4 OR Cause = 6 OR Cause = 7 OR Cause = 9) GROUP BY Cause ORDER BY CAST(Cause AS int);";
                            Tsql += "SELECT V1, COUNT(UserID) AS Expr1 FROM Egg_" + PickTime + " WHERE (Kind = 78) AND (Cause = 5) OR (Kind = 78) AND (Cause = 8) OR (Kind = 78) AND (Cause = 10) GROUP BY V1 ORDER BY CAST(V1 AS int);";

                            //活動積分分析
                            Tsql += "SELECT UserID, MAX(DISTINCT CAST(V1 AS int)) AS Expr1 FROM Role_" + PickTime + " WHERE (Kind = 61) GROUP BY UserID HAVING (MAX(DISTINCT CAST(V1 AS int)) > " + SportPoint + ") ORDER BY MAX(DISTINCT CAST(V1 AS int)) DESC;";

                            //創角人物分析
                            Tsql += "SELECT Kind, Cause, V1, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 6 OR Cause = 7 OR Cause = 8) ORDER BY SaveDate , CAST(V1 AS int), Cause;";

                            //儲值玩家清單
                            Tsql += "SELECT V2, Cause, V4, UserID, SaveDate FROM Money_" + PickTime + " WHERE (Kind = '1') AND (Cause = '22' OR Cause = '23' OR Cause = '24' OR Cause = '15' OR Cause = '14' OR Cause = '13' OR Cause = '10' OR Cause = '9' OR Cause = '16') ORDER BY SaveDate;"; 

                            //當日儲值人數
                            Tsql += "SELECT UserID FROM Money_" + PickTime + " WHERE (Kind = '1') AND (Cause = '22' OR Cause = '23' OR Cause = '24' OR Cause = '15' OR Cause = '14' OR Cause = '13' OR Cause = '10' OR Cause = '9' OR Cause = '16') GROUP BY UserID";

                            textBox_DisLv.Text = "";
                            textBox_DisDimon.Text = "";
                            textBox_DisLog.Text = "";
                            textBox_DisName.Text = "";
                            textBox_DisVip.Text = "";
                            textBox_CardBox.Text = "";
                            textBox_Point25W.Text = "";
                            textBox_Creat.Text = "";
                            richTextBox_ChargeList.Text = "";

                            SqlLink_TW.SQLLink TWsql = new SqlLink_TW.SQLLink();
                            DataSet ds = new DataSet();
                            ds = TWsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);

                            //玩家等級分析
                            int LvCount = 0;
                            bool LVSwitch = true;

                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20130913) //舊協定
                            {
                                for (int x = 0; x < ds.Tables[0].Rows.Count; x++)
                                {
                                    if ((x % 13) == 0)
                                    {
                                        textBox_DisLv.Text += "\r\n";
                                    }

                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V2"]))
                                    {
                                        case 1:
                                            Description = "3 級以下: ";
                                            break;

                                        case 2:
                                            Description = "5 級以下: ";
                                            break;

                                        case 3:
                                            Description = "10 級: ";
                                            break;

                                        case 4:
                                            Description = "15 級: ";
                                            break;

                                        case 5:
                                            Description = "20 級: ";
                                            break;

                                        case 6:
                                            Description = "30 級: ";
                                            break;

                                        case 7:
                                            Description = "40 級: ";
                                            break;

                                        case 8:
                                            Description = "50 級: ";
                                            break;

                                        case 9:
                                            Description = "60 級: ";
                                            break;

                                        case 10:
                                            Description = "70 級: ";
                                            break;

                                        case 11:
                                            Description = "80 級: ";
                                            break;

                                        case 12:
                                            Description = "90 級: ";
                                            break;

                                        case 13:
                                            Description = "超過90 級: ";
                                            break;
                                    }

                                    textBox_DisLv.Text += "Server (" + ds.Tables[0].Rows[x]["V1"].ToString() + ") : " + Description + "  " + ds.Tables[0].Rows[x]["V3"].ToString() + "\r\n";
                                }
                            }

                            else if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130912 && Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20131128) //新協定.v1
                            {
                                string RLv = "";

                                for (int x = 0; x < ds.Tables[0].Rows.Count; x++) //玩家等級分析
                                {
                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]))
                                    {
                                        case 1:
                                            RLv = "3 級以下: ";
                                            break;

                                        case 2:
                                            RLv = "5 級以下: ";
                                            break;

                                        case 3:
                                            RLv = "10 級: ";
                                            break;

                                        case 4:
                                            RLv = "15 級: ";
                                            break;

                                        case 5:
                                            RLv = "20 級: ";
                                            break;

                                        case 6:
                                            RLv = "30 級: ";
                                            break;

                                        case 7:
                                            RLv = "40 級: ";
                                            break;

                                        case 8:
                                            RLv = "50 級: ";
                                            break;

                                        case 9:
                                            RLv = "60 級: ";
                                            break;

                                        case 10:
                                            RLv = "70 級: ";
                                            break;

                                        case 11:
                                            RLv = "80 級: ";
                                            break;

                                        case 12:
                                            RLv = "90 級: ";
                                            break;

                                        case 13:
                                            RLv = "100 級: ";
                                            break;

                                        case 14:
                                            RLv = "115 級: ";
                                            break;

                                        case 15:
                                            RLv = "130 級: ";
                                            break;

                                        case 16:
                                            RLv = "150以下: ";
                                            break;

                                        case 17:
                                            RLv = "超過150 級: ";
                                            break;
                                    }

                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V2"]))
                                    {
                                        case 0:
                                            Description = "其它: ";
                                            break;

                                        case 1:
                                            Description = "一般 Apple: ";
                                            break;

                                        case 2:
                                            Description = "完整 Apple: ";
                                            break;

                                        case 3:
                                            Description = "一般 Goole: ";
                                            break;

                                        case 4:
                                            Description = "完整 Goole: ";
                                            break;

                                        case 5:
                                            Description = "一般 MyCard: ";
                                            break;

                                        case 6:
                                            Description = "完整 MyCard: ";
                                            break;

                                        case 7:
                                            Description = "一般台灣老李 MyCard: ";
                                            break;

                                        case 8:
                                            Description = "完整台灣老李 MyCard: ";
                                            break;

                                        case 9:
                                            Description = "一般 Android: ";
                                            break;

                                        case 10:
                                            Description = "完整 Android: ";
                                            break;

                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) - LvCount == 1)
                                    {
                                        textBox_DisLv.Text += "\r\n" + "\r\n" + " 玩家等級: " + RLv + "\r\n" + "\r\n";
                                        LvCount++;
                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) >= LvCount && LVSwitch == true)
                                    {
                                        textBox_DisLv.Text += Description + "  " + ds.Tables[0].Rows[x]["V3"].ToString() + "\r\n";
                                    }
                                    else
                                    {
                                        LVSwitch = false;
                                    }

                                }
                            }

                            else  //新協定.v2
                            {
                                string RLv = "";
                                for (int x = 0; x < ds.Tables[0].Rows.Count; x++) //玩家等級分析
                                {
                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]))
                                    {
                                        case 1:
                                            RLv = "3 級以下: ";
                                            break;

                                        case 2:
                                            RLv = "5 級以下: ";
                                            break;

                                        case 3:
                                            RLv = "10 級: ";
                                            break;

                                        case 4:
                                            RLv = "15 級: ";
                                            break;

                                        case 5:
                                            RLv = "20 級: ";
                                            break;

                                        case 6:
                                            RLv = "30 級: ";
                                            break;

                                        case 7:
                                            RLv = "40 級: ";
                                            break;

                                        case 8:
                                            RLv = "50 級: ";
                                            break;

                                        case 9:
                                            RLv = "60 級: ";
                                            break;

                                        case 10:
                                            RLv = "70 級: ";
                                            break;

                                        case 11:
                                            RLv = "80 級: ";
                                            break;

                                        case 12:
                                            RLv = "90 級: ";
                                            break;

                                        case 13:
                                            RLv = "100 級: ";
                                            break;

                                        case 14:
                                            RLv = "115 級: ";
                                            break;

                                        case 15:
                                            RLv = "130 級: ";
                                            break;

                                        case 16:
                                            RLv = "150以下: ";
                                            break;

                                        case 17:
                                            RLv = "超過150 級: ";
                                            break;
                                    }

                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V2"]))
                                    {
                                        case 0:
                                            Description = "Ios 其他: ";
                                            break;

                                        case 1:
                                            Description = "香港 IOS: ";
                                            break;

                                        case 2:
                                            Description = "香港 IOS 完整版: ";
                                            break;

                                        case 3:
                                            Description = "台灣 IOS 完整版: ";
                                            break;

                                        case 4:
                                            Description = "台灣 IOS: ";
                                            break;

                                        case 5:
                                            Description = "大陸 IOS 完整版: ";
                                            break;

                                        case 6:
                                            Description = "Android 其他: ";
                                            break;

                                        case 7:
                                            Description = "香港 MyCard: ";
                                            break;

                                        case 8:
                                            Description = "香港 MyCard 完整版: ";
                                            break;

                                        case 9:
                                            Description = "香港 Google: ";
                                            break;

                                        case 10:
                                            Description = "台灣 MyCard 完整版: ";
                                            break;

                                        case 11:
                                            Description = "台灣 Google: ";
                                            break;

                                        case 12:
                                            Description = "大陸 Android 完整版 (未): ";
                                            break;

                                        case 13:
                                            Description = "IOS Debug: ";
                                            break;

                                        case 14:
                                            Description = "Android Debug: ";
                                            break;

                                        case 15:
                                            Description = "台灣老李 MyCard: ";
                                            break;

                                        case 16:
                                            Description = "台灣 MyCard: ";
                                            break;

                                        case 17:
                                            Description = "台灣 Google 完整版: ";
                                            break;

                                        case 18:
                                            Description = "台灣老李 MyCard 完整版: ";
                                            break;

                                        case 19:
                                            Description = "香港 Google 完整版: ";
                                            break;

                                        case 20:
                                            Description = "松崗: ";
                                            break;

                                        case 21:
                                            Description = "新加坡 IOS: ";
                                            break;

                                        case 22:
                                            Description = "新加坡 Android 完整版: ";
                                            break;

                                        case 23:
                                            Description = "新加坡 Android 當地金流: ";
                                            break;

                                        case 24:
                                            Description = "大陸91 IOS: ";
                                            break;

                                        case 25:
                                            Description = "大陸UC Android: ";
                                            break;

                                        case 26:
                                            Description = "大陸360 Android: ";
                                            break;

                                        case 27:
                                            Description = "大陸App助手 IOS: ";
                                            break;

                                        case 28:
                                            Description = "大陸官方apk Android: ";
                                            break;

                                        case 29:
                                            Description = "大陸91 Android: ";
                                            break;

                                        case 30:
                                            Description = "馬幹線 IOS: ";
                                            break;

                                        case 31:
                                            Description = "馬幹線 Android: ";
                                            break;

                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) - LvCount == 1)
                                    {
                                        textBox_DisLv.Text += "\r\n" + "\r\n" + " 玩家等級: " + RLv + "\r\n" + "\r\n";
                                        LvCount++;
                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) >= LvCount && LVSwitch == true)
                                    {
                                        textBox_DisLv.Text += Description + "  " + ds.Tables[0].Rows[x]["V3"].ToString() + "\r\n";
                                    }
                                    else
                                    {
                                        LVSwitch = false;
                                    }

                                }
                            }


                            //玩家名聲分析

                            int y = 0;

                            for (int x = 0; x < ds.Tables[1].Rows.Count; x++)
                            {

                                if (Convert.ToInt32(ds.Tables[1].Rows[x]["V1"]) - y == 1)
                                {
                                    textBox_DisName.Text += "\r\n" + "伺服器: " + ds.Tables[1].Rows[x]["V1"].ToString() + "\r\n" + "\r\n";
                                    y++;
                                }

                                switch (Convert.ToInt32(ds.Tables[1].Rows[x]["V2"]))
                                {

                                    case 1:
                                        Description = "1000 以下: ";
                                        break;

                                    case 2:
                                        Description = "2000: ";
                                        break;

                                    case 3:
                                        Description = "5000: ";
                                        break;

                                    case 4:
                                        Description = "10000: ";
                                        break;

                                    case 5:
                                        Description = "20000: ";
                                        break;

                                    case 6:
                                        Description = "50000: ";
                                        break;

                                    case 7:
                                        Description = "100000: ";
                                        break;

                                    case 8:
                                        Description = "200000: ";
                                        break;

                                    case 9:
                                        Description = "250000: ";
                                        break;

                                    case 10:
                                        Description = "300000: ";
                                        break;

                                    case 11:
                                        Description = "350000: ";
                                        break;

                                    case 12:
                                        Description = "400000: ";
                                        break;

                                    case 13:
                                        Description = "450000: ";
                                        break;

                                    case 14:
                                        Description = "500000: ";
                                        break;

                                    case 15:
                                        Description = "超過 500000: ";
                                        break;

                                    case 16:
                                        Description = "1000000: ";
                                        break;

                                    case 17:
                                        Description = "1500000: ";
                                        break;
                                }

                                textBox_DisName.Text += Description + "  " + ds.Tables[1].Rows[x]["V3"].ToString() + "\r\n";

                            }




                            //Vip分析

                            int VipCount = -1;

                            for (int x = 0; x < ds.Tables[2].Rows.Count; x++)
                            {
                                if (Convert.ToInt32(ds.Tables[2].Rows[x]["V1"]) - VipCount == 1)
                                {
                                    textBox_DisVip.Text += "\r\n" + "Vip等級: " + ds.Tables[2].Rows[x]["V1"].ToString() + "\r\n" + "\r\n";
                                    VipCount++;
                                }

                                if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20130912) //舊協定
                                {
                                    switch (Convert.ToInt32(ds.Tables[2].Rows[x]["V2"].ToString()))
                                    {

                                        case 0:
                                            Description = "其他: ";
                                            break;

                                        case 1:
                                            Description = "Apple: ";
                                            break;

                                        case 2:
                                            Description = "一般 My Card: ";
                                            break;

                                        case 3:
                                            Description = "完整 My Card: ";
                                            break;

                                        case 4:
                                            Description = "Google: ";
                                            break;

                                        case 5:
                                            Description = "東方阿李: ";
                                            break;
                                    }
                                }

                                else if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130911 && Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20131128) //舊協定
                                {
                                    switch (Convert.ToInt32(ds.Tables[2].Rows[x]["V2"].ToString()))
                                    {

                                        case 0:
                                            Description = "其他: ";
                                            break;

                                        case 1:
                                            Description = "一般 Apple: ";
                                            break;

                                        case 2:
                                            Description = "完整 Apple: ";
                                            break;

                                        case 3:
                                            Description = "一般 Google: ";
                                            break;

                                        case 4:
                                            Description = "完整 Google: ";
                                            break;

                                        case 5:
                                            Description = "一般 MyCard: ";
                                            break;

                                        case 6:
                                            Description = "完整 MyCard: ";
                                            break;

                                        case 7:
                                            Description = "一般台灣老李MyCard: ";
                                            break;

                                        case 8:
                                            Description = "完整台灣老李 MyCard: ";
                                            break;

                                        case 9:
                                            Description = "一般 Android: ";
                                            break;

                                        case 10:
                                            Description = "完整 Android: ";
                                            break;
                                    }
                                }

                                else
                                {
                                    switch (Convert.ToInt32(ds.Tables[2].Rows[x]["V2"].ToString()))
                                    {

                                        case 0:
                                            Description = "Ios 其他: ";
                                            break;

                                        case 1:
                                            Description = "香港 IOS: ";
                                            break;

                                        case 2:
                                            Description = "香港 IOS 完整版: ";
                                            break;

                                        case 3:
                                            Description = "台灣 IOS 完整版: ";
                                            break;

                                        case 4:
                                            Description = "台灣 IOS: ";
                                            break;

                                        case 5:
                                            Description = "大陸 IOS 完整版: ";
                                            break;

                                        case 6:
                                            Description = "Android 其他: ";
                                            break;

                                        case 7:
                                            Description = "香港 MyCard: ";
                                            break;

                                        case 8:
                                            Description = "香港 MyCard 完整版: ";
                                            break;

                                        case 9:
                                            Description = "香港 Google: ";
                                            break;

                                        case 10:
                                            Description = "台灣 MyCard 完整版: ";
                                            break;

                                        case 11:
                                            Description = "台灣 Google: ";
                                            break;

                                        case 12:
                                            Description = "大陸 Android 完整版 (未): ";
                                            break;

                                        case 13:
                                            Description = "IOS Debug: ";
                                            break;

                                        case 14:
                                            Description = "Android Debug: ";
                                            break;

                                        case 15:
                                            Description = "台灣老李 MyCard: ";
                                            break;

                                        case 16:
                                            Description = "台灣 MyCard: ";
                                            break;

                                        case 17:
                                            Description = "台灣 Google 完整版: ";
                                            break;

                                        case 18:
                                            Description = "台灣老李 MyCard 完整版: ";
                                            break;

                                        case 19:
                                            Description = "香港 Google 完整版: ";
                                            break;

                                        case 20:
                                            Description = "松崗: ";
                                            break;

                                        case 21:
                                            Description = "新加坡 IOS: ";
                                            break;

                                        case 22:
                                            Description = "新加坡 Android 完整版: ";
                                            break;

                                        case 23:
                                            Description = "新加坡 Android 當地金流: ";
                                            break;

                                        case 24:
                                            Description = "大陸91 IOS: ";
                                            break;

                                        case 25:
                                            Description = "大陸UC Android: ";
                                            break;

                                        case 26:
                                            Description = "大陸360 Android: ";
                                            break;

                                        case 27:
                                            Description = "大陸App助手 IOS: ";
                                            break;

                                        case 28:
                                            Description = "大陸官方apk Android: ";
                                            break;

                                        case 29:
                                            Description = "大陸91 Android: ";
                                            break;

                                        case 30:
                                            Description = "馬幹線 IOS: ";
                                            break;

                                        case 31:
                                            Description = "馬幹線 Android: ";
                                            break;
                                    }
                                }

                                textBox_DisVip.Text += Description + "  " + ds.Tables[2].Rows[x]["V3"].ToString() + "\r\n";
                            }



                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130905)
                            {
                                //玩家現有鑽石分析

                                int DimonCount = 0;

                                for (int x = 0; x < ds.Tables[3].Rows.Count; x++)
                                {
                                    if (Convert.ToInt32(ds.Tables[3].Rows[x]["V1"]) - DimonCount == 1)
                                    {
                                        textBox_DisDimon.Text += "\r\n" + "伺服器: " + ds.Tables[3].Rows[x]["V1"].ToString() + "\r\n" + "\r\n";
                                        DimonCount++;
                                    }

                                    switch (Convert.ToInt32(ds.Tables[3].Rows[x]["V2"]))
                                    {
                                        case 0:
                                            Description = "60 以下: ";
                                            break;

                                        case 1:
                                            Description = "500: ";
                                            break;

                                        case 2:
                                            Description = "1000: ";
                                            break;

                                        case 3:
                                            Description = "2000: ";
                                            break;

                                        case 4:
                                            Description = "5000: ";
                                            break;

                                        case 5:
                                            Description = "10000: ";
                                            break;

                                        case 6:
                                            Description = "20000: ";
                                            break;

                                        case 7:
                                            Description = "50000: ";
                                            break;

                                        case 8:
                                            Description = "100000: ";
                                            break;

                                        case 9:
                                            Description = "10萬到50萬: ";
                                            break;

                                        case 10:
                                            Description = "500000以上: ";
                                            break;
                                    }

                                    textBox_DisDimon.Text += Description + "  " + ds.Tables[3].Rows[x]["V3"].ToString() + "\r\n";
                                }

                            }
                            else
                            {
                                textBox_DisDimon.Text = "資料庫無資料可分析";
                            }


                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130904)
                            {
                                for (int j = 0; j < 1; j++) //登入方式
                                {
                                    textBox_DisLog.Text += "\r\n" + "登入方式: " + "\r\n" + "\r\n";

                                    for (int x = 0; x < ds.Tables[j + 4].Rows.Count; x++)
                                    {
                                        switch (Convert.ToInt32(ds.Tables[j + 4].Rows[x]["V2"]))
                                        {

                                            case 0:
                                                Description = "其它: ";
                                                break;

                                            case 1:
                                                Description = "香港一般版: ";
                                                break;

                                            case 2:
                                                Description = "香港 IOS 完整版: ";
                                                break;

                                            case 3:
                                                Description = "台灣 IOS 完整版: ";
                                                break;

                                            case 4:
                                                Description = "台灣一般版: ";
                                                break;

                                            case 5:
                                                Description = "阿六版: ";
                                                break;

                                            case 6:
                                                Description = "未登錄: ";
                                                break;

                                            case 7:
                                                Description = "香港一般MYCARD: ";
                                                break;

                                            case 8:
                                                Description = "香港完整MYCARD: ";
                                                break;

                                            case 9:
                                                Description = "香港一般Google: ";
                                                break;

                                            case 10:
                                                Description = "台灣完整MYCARD: ";
                                                break;

                                            case 11:
                                                Description = "台灣一般google: ";
                                                break;

                                            case 12:
                                                Description = "大陸 Android 完整版 (未): ";
                                                break;

                                            case 13:
                                                Description = "IOS Debug: ";
                                                break;

                                            case 14:
                                                Description = "Android Debug: ";
                                                break;

                                            case 15:
                                                Description = "東方阿李: ";
                                                break;

                                            case 16:
                                                Description = "台灣 MyCard: ";
                                                break;

                                            case 17:
                                                Description = "台灣 Google 完整版: ";
                                                break;

                                            case 18:
                                                Description = "台灣老李 MyCard 完整版: ";
                                                break;

                                            case 19:
                                                Description = "香港 Google 完整版: ";
                                                break;

                                            case 20:
                                                Description = "松崗: ";
                                                break;

                                            case 21:
                                                Description = "新加坡 IOS: ";
                                                break;

                                            case 22:
                                                Description = "新加坡 Android 完整版: ";
                                                break;

                                            case 23:
                                                Description = "新加坡 Android 當地金流: ";
                                                break;

                                            case 24:
                                                Description = "大陸91 IOS: ";
                                                break;

                                            case 25:
                                                Description = "大陸UC Android: ";
                                                break;

                                            case 26:
                                                Description = "大陸360 Android: ";
                                                break;

                                            case 27:
                                                Description = "大陸App助手 IOS: ";
                                                break;

                                            case 28:
                                                Description = "大陸官方apk Android: ";
                                                break;

                                            case 29:
                                                Description = "大陸91 Android: ";
                                                break;

                                            case 30:
                                                Description = "馬幹線 IOS: ";
                                                break;

                                            case 31:
                                                Description = "馬幹線 Android: ";
                                                break;
                                        }

                                        textBox_DisLog.Text += Description + "  " + ds.Tables[j + 4].Rows[x]["V3"].ToString() + "\r\n" + "\r\n";
                                    }
                                }
                            }
                            else
                            {
                                textBox_DisLog.Text = "資料庫無資料可分析";
                            }

                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130905)
                            {
                                //卡包分析

                                textBox_CardBox.Text += "\r\n";

                                for (int x = 0; x < ds.Tables[5].Rows.Count; x++)
                                {
                                    switch (Convert.ToInt32(ds.Tables[5].Rows[x]["Cause"]))
                                    {
                                        case 0:
                                            Description = "聖旨單抽: ";
                                            break;

                                        case 1:
                                            Description = "友情單抽: ";
                                            break;

                                        case 2:
                                            Description = "銀幣單抽: ";
                                            break;

                                        case 3:
                                            Description = "鑽石單抽: ";
                                            break;

                                        case 6:
                                            Description = "友情10連抽: ";
                                            break;

                                        case 7:
                                            Description = "銀幣10連抽: ";
                                            break;

                                        case 9:
                                            Description = "９鑽抽:     ";
                                            break;
                                    }

                                    textBox_CardBox.Text += Description + "  " + ds.Tables[5].Rows[x]["Expr1"].ToString() + "\r\n" + "\r\n";
                                }


                                textBox_CardBox.Text += "\r\n";

                                for (int x = 0; x < ds.Tables[6].Rows.Count; x++) //卡包分析
                                {
                                    switch (Convert.ToInt32(ds.Tables[6].Rows[x]["V1"]))
                                    {
                                        case 4:
                                            Description = "必中卡包: ";
                                            break;

                                        case 7:
                                            Description = "一般卡包: ";
                                            break;

                                        case 11:
                                            Description = "卡包1(吳國太): ";
                                            break;

                                        case 12:
                                            Description = "卡包2(甘夫人): ";
                                            break;

                                        case 13:
                                            Description = "卡包3(蔡文姬): ";
                                            break;

                                        case 14:
                                            Description = "卡包4(夏日美女包): ";
                                            break;

                                        case 15:
                                            Description = "卡包5(闇屬趙雲): ";
                                            break;

                                        case 16:
                                            Description = "卡包6(光屬張飛): ";
                                            break;

                                        case 17:
                                            Description = "卡包7(蜀國勢力包:龐統必中): ";
                                            break;

                                        case 18:
                                            Description = "卡包8(吳國勢力包:龐統必中): ";
                                            break;

                                        case 19:
                                            Description = "卡包9(魏國勢力包:司馬懿): ";
                                            break;

                                        case 20:
                                            Description = "卡包10(他國勢力包:水呂布): ";
                                            break;

                                        case 21:
                                            Description = "卡包11: ";
                                            break;

                                        case 22:
                                            Description = "卡包12: ";
                                            break;

                                        case 23:
                                            Description = "卡包13: ";
                                            break;

                                        case 24:
                                            Description = "卡包14: ";
                                            break;

                                        case 25:
                                            Description = "卡包15: ";
                                            break;

                                        case 26:
                                            Description = "卡包16: ";
                                            break;

                                        case 27:
                                            Description = "卡包17: ";
                                            break;

                                        case 28:
                                            Description = "卡包18: ";
                                            break;

                                        case 29:
                                            Description = "卡包19: ";
                                            break;

                                        case 30:
                                            Description = "卡包20: ";
                                            break;


                                    }

                                    textBox_CardBox.Text += Description + "  " + ds.Tables[6].Rows[x]["Expr1"].ToString() + "\r\n" + "\r\n";
                                }
                            }
                            else
                            {
                                textBox_CardBox.Text = "資料庫無資料可分析";
                            }

                            //活動積分分析
                            textBox_Point25W.Text += "\r\n";
                            int ID = 0;
                            for (int x = 0; x < ds.Tables[7].Rows.Count; x++)
                            {
                                ID++;
                                textBox_Point25W.Text += " (" + ID + "). " + "玩家ID: " + ds.Tables[7].Rows[x]["UserID"].ToString() + "     積分: " + ds.Tables[7].Rows[x]["Expr1"].ToString() + "\r\n" + "\r\n";
                            }

                            if (textBox_Point25W.Text == "\r\n")
                            {
                                textBox_Point25W.Text = "無對應此積分資料";
                            }

                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130917)
                            {
                                if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20131129)
                                {
                                    #region CreatCharacter

                                    //當日創角
                                    textBox_Creat.Text += "當日創角人數 /" + "\r\n" + "當日創角且超過5等 /" + "\r\n" + "前天創角且連續兩天有上線" + "\r\n" + "\r\n";

                                    #region 數字空間
                                    string Num6_0 = "";
                                    string Num6_1 = "";
                                    string Num6_2 = "";
                                    string Num6_3 = "";
                                    string Num6_4 = "";
                                    string Num6_5 = "";
                                    string Num6_6 = "";
                                    string Num6_7 = "";
                                    string Num6_8 = "";
                                    string Num6_9 = "";
                                    string Num6_10 = "";

                                    string Num7_0 = "";
                                    string Num7_1 = "";
                                    string Num7_2 = "";
                                    string Num7_3 = "";
                                    string Num7_4 = "";
                                    string Num7_5 = "";
                                    string Num7_6 = "";
                                    string Num7_7 = "";
                                    string Num7_8 = "";
                                    string Num7_9 = "";
                                    string Num7_10 = "";

                                    string Num8_0 = "";
                                    string Num8_1 = "";
                                    string Num8_2 = "";
                                    string Num8_3 = "";
                                    string Num8_4 = "";
                                    string Num8_5 = "";
                                    string Num8_6 = "";
                                    string Num8_7 = "";
                                    string Num8_8 = "";
                                    string Num8_9 = "";
                                    string Num8_10 = "";
                                    #endregion

                                    for (int x = 0; x < ds.Tables[8].Rows.Count; x++)
                                    {
                                        switch (Convert.ToInt32(ds.Tables[8].Rows[x]["Cause"].ToString()))
                                        {
                                            case 6:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"其他:　　　　　　　　　　　";
                                                        Num6_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"一般Ａｐｐｌｅ:　　　　　　";
                                                        Num6_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"完整Ａｐｐｌｅ:　　　　　　";
                                                        Num6_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"一般Ｇｏｏｇｌｅ:　　　　　";
                                                        Num6_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"完整Ｇｏｏｇｌｅ:　　　　　";
                                                        Num6_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"一般ＭｙＣａｒｄ:　　　　　";
                                                        Num6_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"完整ＭｙＣａｒｄ:　　　　　";
                                                        Num6_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"一般台灣老李ＭｙＣａｒｄ:　";
                                                        Num6_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"完整台灣老李ＭｙＣａｒｄ:　";
                                                        Num6_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"一般Ａｎｄｒｏｉｄ:　　　　";
                                                        Num6_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"完整Ａｎｄｒｏｉｄ:　　　　";
                                                        Num6_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                            case 7:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"其他:　　　　　　　　　　　";
                                                        Num7_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"一般Ａｐｐｌｅ:　　　　　　";
                                                        Num7_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"完整Ａｐｐｌｅ:　　　　　　";
                                                        Num7_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"一般Ｇｏｏｇｌｅ:　　　　　";
                                                        Num7_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"完整Ｇｏｏｇｌｅ:　　　　　";
                                                        Num7_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"一般ＭｙＣａｒｄ:　　　　　";
                                                        Num7_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"完整ＭｙＣａｒｄ:　　　　　";
                                                        Num7_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"一般台灣老李ＭｙＣａｒｄ:　";
                                                        Num7_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"完整台灣老李ＭｙＣａｒｄ:　";
                                                        Num7_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"一般Ａｎｄｒｏｉｄ:　　　　";
                                                        Num7_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"完整Ａｎｄｒｏｉｄ:　　　　";
                                                        Num7_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }


                                                break;

                                            case 8:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"其他:　　　　　　　　　　　";
                                                        Num8_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"一般Ａｐｐｌｅ:　　　　　　";
                                                        Num8_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"完整Ａｐｐｌｅ:　　　　　　";
                                                        Num8_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"一般Ｇｏｏｇｌｅ:　　　　　";
                                                        Num8_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"完整Ｇｏｏｇｌｅ:　　　　　";
                                                        Num8_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"一般ＭｙＣａｒｄ:　　　　　";
                                                        Num8_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"完整ＭｙＣａｒｄ:　　　　　";
                                                        Num8_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"一般台灣老李ＭｙＣａｒｄ:　";
                                                        Num8_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"完整台灣老李ＭｙＣａｒｄ:　";
                                                        Num8_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"一般Ａｎｄｒｏｉｄ:　　　　";
                                                        Num8_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"完整Ａｎｄｒｏｉｄ:　　　　";
                                                        Num8_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                        }

                                    }

                                    textBox_Creat.Text += "其他:　　　　　　　　　　　" + Num6_0 + " / " + Num7_0 + " / " + Num8_0 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般Ａｐｐｌｅ:　　　　　　" + Num6_1 + " / " + Num7_1 + " / " + Num8_1 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整Ａｐｐｌｅ:　　　　　　" + Num6_2 + " / " + Num7_2 + " / " + Num8_2 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般Ｇｏｏｇｌｅ:　　　　　" + Num6_3 + " / " + Num7_3 + " / " + Num8_3 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整Ｇｏｏｇｌｅ:　　　　　" + Num6_4 + " / " + Num7_4 + " / " + Num8_4 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般ＭｙＣａｒｄ:　　　　　" + Num6_5 + " / " + Num7_5 + " / " + Num8_5 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整ＭｙＣａｒｄ:　　　　　" + Num6_6 + " / " + Num7_6 + " / " + Num8_6 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般台灣老李ＭｙＣａｒｄ:　" + Num6_7 + " / " + Num7_7 + " / " + Num8_7 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整台灣老李ＭｙＣａｒｄ:　" + Num6_8 + " / " + Num7_8 + " / " + Num8_8 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般Ａｎｄｒｏｉｄ:　　　　" + Num6_9 + " / " + Num7_9 + " / " + Num8_9 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整Ａｎｄｒｏｉｄ:　　　　" + Num6_10 + " / " + Num7_10 + " / " + Num8_10 + "\r\n" + "\r\n";

                                    #endregion
                                }

                                else
                                {
                                    #region CreatCharacter

                                    //當日創角
                                    textBox_Creat.Text += "當日創角人數 /" + "\r\n" + "當日創角且超過5等 /" + "\r\n" + "前天創角且連續兩天有上線" + "\r\n" + "\r\n";

                                    #region 數字空間
                                    string Num6_0 = "";
                                    string Num6_1 = "";
                                    string Num6_2 = "";
                                    string Num6_3 = "";
                                    string Num6_4 = "";
                                    string Num6_5 = "";
                                    string Num6_6 = "";
                                    string Num6_7 = "";
                                    string Num6_8 = "";
                                    string Num6_9 = "";
                                    string Num6_10 = "";
                                    string Num6_11 = "";
                                    string Num6_12 = "";
                                    string Num6_13 = "";
                                    string Num6_14 = "";
                                    string Num6_15 = "";
                                    string Num6_16 = "";
                                    string Num6_17 = "";
                                    string Num6_18 = "";
                                    string Num6_19 = "";
                                    string Num6_20 = "";
                                    string Num6_21 = "";
                                    string Num6_22 = "";

                                    string Num7_0 = "";
                                    string Num7_1 = "";
                                    string Num7_2 = "";
                                    string Num7_3 = "";
                                    string Num7_4 = "";
                                    string Num7_5 = "";
                                    string Num7_6 = "";
                                    string Num7_7 = "";
                                    string Num7_8 = "";
                                    string Num7_9 = "";
                                    string Num7_10 = "";
                                    string Num7_11 = "";
                                    string Num7_12 = "";
                                    string Num7_13 = "";
                                    string Num7_14 = "";
                                    string Num7_15 = "";
                                    string Num7_16 = "";
                                    string Num7_17 = "";
                                    string Num7_18 = "";
                                    string Num7_19 = "";
                                    string Num7_20 = "";
                                    string Num7_21 = "";
                                    string Num7_22 = "";

                                    string Num8_0 = "";
                                    string Num8_1 = "";
                                    string Num8_2 = "";
                                    string Num8_3 = "";
                                    string Num8_4 = "";
                                    string Num8_5 = "";
                                    string Num8_6 = "";
                                    string Num8_7 = "";
                                    string Num8_8 = "";
                                    string Num8_9 = "";
                                    string Num8_10 = "";
                                    string Num8_11 = "";
                                    string Num8_12 = "";
                                    string Num8_13 = "";
                                    string Num8_14 = "";
                                    string Num8_15 = "";
                                    string Num8_16 = "";
                                    string Num8_17 = "";
                                    string Num8_18 = "";
                                    string Num8_19 = "";
                                    string Num8_20 = "";
                                    string Num8_21 = "";
                                    string Num8_22 = "";

                                    #endregion

                                    for (int x = 0; x < ds.Tables[8].Rows.Count; x++)
                                    {
                                        switch (Convert.ToInt32(ds.Tables[8].Rows[x]["Cause"].ToString()))
                                        {
                                            case 6:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"Ｉｏｓ其他:　　　　　　　　　　　　　";
                                                        Num6_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"香港　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num6_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"香港　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num6_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"台灣　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num6_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"台灣　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num6_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"大陸　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num6_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"Ａｎｄｒｏｉｄ其他:　　　　　　　　　";
                                                        Num6_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"香港　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num6_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"香港　ＭｙＣａｒｄ完整版:　　　　　　";
                                                        Num6_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"香港　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num6_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"台灣　Ｍｙｃａｒｄ完整版:　　　　　　";
                                                        Num6_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 11:
                                                        //"台灣　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num6_11 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 12:
                                                        //"大陸　Ａｎｄｒｏｉｄ 完整版 (未):　　";
                                                        Num6_12 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 13:
                                                        //"ＩｏｓＤｅｂｕｇ:　　　　　　　　　　";
                                                        Num6_13 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 14:
                                                        //"ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　";
                                                        Num6_14 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 15:
                                                        //"台灣老李　ＭｙＣａｒｄ:　　　　　　　";
                                                        Num6_15 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 16:
                                                        //"台灣　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num6_16 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 17:
                                                        //"台灣　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num6_17 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 18:
                                                        //"台灣老李　ＭｙＣａｒｄ完整版:　　　　";
                                                        Num6_18 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 19:
                                                        //"香港　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num6_19 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 20:
                                                        //"松崗:　　　　　　　　　　　　　　　　";
                                                        Num6_20 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 21:
                                                        //"新加坡　Ｉｏｓ:　　　　　　　　　　　";
                                                        Num6_21 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 22:
                                                        //"新加坡　Ａｎｄｒｏｉｄ完整版:　　　　";
                                                        Num6_22 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                            case 7:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"Ｉｏｓ其他:　　　　　　　　　　　　　";
                                                        Num7_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"香港　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num7_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"香港　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num7_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"台灣　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num7_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"台灣　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num7_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"大陸　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num7_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"Ａｎｄｒｏｉｄ其他:　　　　　　　　　";
                                                        Num7_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"香港　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num7_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"香港　ＭｙＣａｒｄ完整版:　　　　　　";
                                                        Num7_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"香港　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num7_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"台灣　Ｍｙｃａｒｄ完整版:　　　　　　";
                                                        Num7_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 11:
                                                        //"台灣　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num7_11 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 12:
                                                        //"大陸　Ａｎｄｒｏｉｄ 完整版 (未):　　";
                                                        Num7_12 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 13:
                                                        //"ＩｏｓＤｅｂｕｇ:　　　　　　　　　　";
                                                        Num7_13 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 14:
                                                        //"ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　";
                                                        Num7_14 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 15:
                                                        //"台灣老李　ＭｙＣａｒｄ:　　　　　　　";
                                                        Num7_15 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 16:
                                                        //"台灣　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num7_16 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 17:
                                                        //"台灣　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num7_17 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 18:
                                                        //"台灣老李　ＭｙＣａｒｄ完整版:　　　　";
                                                        Num7_18 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 19:
                                                        //"香港　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num7_19 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 20:
                                                        //"松崗:　　　　　　　　　　　　　　　　";
                                                        Num7_20 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 21:
                                                        //"新加坡　Ｉｏｓ:　　　　　　　　　　　";
                                                        Num7_21 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 22:
                                                        //"新加坡　Ａｎｄｒｏｉｄ完整版:　　　　";
                                                        Num7_22 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }


                                                break;

                                            case 8:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"Ｉｏｓ其他:　　　　　　　　　　　　　";
                                                        Num8_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"香港　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num8_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"香港　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num8_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"台灣　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num8_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"台灣　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num8_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"大陸　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num8_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"Ａｎｄｒｏｉｄ其他:　　　　　　　　　";
                                                        Num8_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"香港　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num8_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"香港　ＭｙＣａｒｄ完整版:　　　　　　";
                                                        Num8_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"香港　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num8_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"台灣　Ｍｙｃａｒｄ完整版:　　　　　　";
                                                        Num8_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 11:
                                                        //"台灣　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num8_11 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 12:
                                                        //"大陸　Ａｎｄｒｏｉｄ 完整版 (未):　　";
                                                        Num8_12 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 13:
                                                        //"ＩｏｓＤｅｂｕｇ:　　　　　　　　　　";
                                                        Num8_13 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 14:
                                                        //"ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　";
                                                        Num8_14 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 15:
                                                        //"台灣老李　ＭｙＣａｒｄ:　　　　　　　";
                                                        Num8_15 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 16:
                                                        //"台灣　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num8_16 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 17:
                                                        //"台灣　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num8_17 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 18:
                                                        //"台灣老李　ＭｙＣａｒｄ完整版:　　　　";
                                                        Num8_18 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 19:
                                                        //"香港　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num8_19 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 20:
                                                        //"松崗:　　　　　　　　　　　　　　　　";
                                                        Num8_20 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 21:
                                                        //"新加坡　Ｉｏｓ:　　　　　　　　　　　";
                                                        Num8_21 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 22:
                                                        //"新加坡　Ａｎｄｒｏｉｄ完整版:　　　　";
                                                        Num8_22 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                        }

                                    }

                                    textBox_Creat.Text += "Ｉｏｓ其他:　　　　　　　　　　　　　" + Num6_0 + " / " + Num7_0 + " / " + Num8_0 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｉｏｓ:　　　　　　　　　　　　" + Num6_1 + " / " + Num7_1 + " / " + Num8_1 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｉｏｓ完整版:　　　　　　　　　" + Num6_2 + " / " + Num7_2 + " / " + Num8_2 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｉｏｓ完整版:　　　　　　　　　" + Num6_3 + " / " + Num7_3 + " / " + Num8_3 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｉｏｓ:　　　　　　　　　　　　" + Num6_4 + " / " + Num7_4 + " / " + Num8_4 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "大陸　Ｉｏｓ完整版:　　　　　　　　　" + Num6_5 + " / " + Num7_5 + " / " + Num8_5 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "Ａｎｄｒｏｉｄ其他:　　　　　　　　　" + Num6_6 + " / " + Num7_6 + " / " + Num8_6 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　ＭｙＣａｒｄ:　　　　　　　　　" + Num6_7 + " / " + Num7_7 + " / " + Num8_7 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　ＭｙＣａｒｄ完整版:　　　　　　" + Num6_8 + " / " + Num7_8 + " / " + Num8_8 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｇｏｏｇｌｅ:　　　　　　　　　" + Num6_9 + " / " + Num7_9 + " / " + Num8_9 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｍｙｃａｒｄ完整版:　　　　　　" + Num6_10 + " / " + Num7_10 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｇｏｏｇｌｅ:　　　　　　　　　" + Num6_11 + " / " + Num7_11 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "大陸　Ａｎｄｒｏｉｄ完整版（未）:　　" + Num6_12 + " / " + Num7_12 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "ＩｏｓＤｅｂｕｇ:　　　　　　　　　　" + Num6_13 + " / " + Num7_13 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　" + Num6_14 + " / " + Num7_14 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣老李　ＭｙＣａｒｄ:　　　　　　　" + Num6_15 + " / " + Num7_15 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　ＭｙＣａｒｄ:　　　　　　　　　" + Num6_16 + " / " + Num7_16 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｇｏｏｇｌｅ完整版:　　　　　　" + Num6_17 + " / " + Num7_17 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣老李　ＭｙＣａｒｄ完整版:　　　　" + Num6_18 + " / " + Num7_18 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｇｏｏｇｌｅ完整版:　　　　　　" + Num6_19 + " / " + Num7_19 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "松崗:　　　　　　　　　　　　　　　　" + Num6_20 + " / " + Num7_20 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "新加坡　Ｉｏｓ:　　　　　　　　　　　" + Num6_21 + " / " + Num7_21 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "新加坡　Ａｎｄｒｏｉｄ完整版:　　　　" + Num6_22 + " / " + Num7_22 + " / " + Num8_10 + "\r\n" + "\r\n";

                                    #endregion
                                }
                            }

                            else
                            {
                                textBox_Creat.Text = "資料庫無資料可分析";
                            }

                            //儲值人數
                            richTextBox_ChargeList.Text += "當日總儲值人數 : " + ds.Tables[10].Rows.Count + "\r\n" + "\r\n";

                            //儲值玩家清單
                            for (int x = 0; x < ds.Tables[9].Rows.Count; x++)
                            {
                                string CauseName = "";
                                float TWD = 0;

                                switch (ds.Tables[9].Rows[x]["Cause"].ToString())
                                {
                                    #region 流向分類

                                    case "10":

                                        if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                        {
                                            CauseName = "Ｍ吞食Ｑ鑽　　　　";
                                        }
                                        else
                                        {
                                            CauseName = "ＭｙＣａｒｄ　　　";
                                        }

                                        switch (ds.Tables[9].Rows[x]["V2"].ToString())
                                        {
                                            case "10":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 10;
                                                }

                                                else
                                                {
                                                    TWD = 10;
                                                }

                                                break;

                                            case "53":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 50;
                                                }

                                                else
                                                {
                                                    TWD = 50;
                                                }

                                                break;

                                            case "168":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 150;
                                                }

                                                else
                                                {
                                                    TWD = 150;
                                                }

                                                break;

                                            case "336":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 300;
                                                }

                                                else
                                                {
                                                    TWD = 300;
                                                }

                                                break;

                                            case "403":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 350;
                                                }

                                                else
                                                {
                                                    TWD = 350;
                                                }

                                                break;

                                            case "480":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 400;
                                                }

                                                else
                                                {
                                                    TWD = 400;
                                                }

                                                break;

                                            case "540":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 450;
                                                }

                                                else
                                                {
                                                    TWD = 450;
                                                }

                                                break;

                                            case "600":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 500;
                                                }

                                                else
                                                {
                                                    TWD = 500;
                                                }

                                                break;

                                            case "1270":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 1000;
                                                }

                                                else
                                                {
                                                    TWD = 1000;
                                                }

                                                break;

                                            case "1460":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 1150;
                                                }

                                                else
                                                {
                                                    TWD = 1150;
                                                }

                                                break;

                                            case "2600":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 2000;
                                                }

                                                else
                                                {
                                                    TWD = 2000;
                                                }

                                                break;

                                            case "4000":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 3000;
                                                }

                                                else
                                                {
                                                    TWD = 3000;
                                                }

                                                break;

                                            case "6750":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 5000;
                                                }

                                                else
                                                {
                                                    TWD = 5000;
                                                }

                                                break;

                                        }
                                        break;

                                    case "13":

                                        CauseName = "ＡＰＰＬＥ　　　　";

                                        switch (ds.Tables[9].Rows[x]["V2"].ToString())
                                        {
                                            case "64":

                                                TWD = 60;

                                                break;

                                            case "168":

                                                TWD = 150;

                                                break;

                                            case "336":

                                                TWD = 300;

                                                break;

                                            case "540":

                                                TWD = 450;

                                                break;

                                            case "958":

                                                TWD = 750;

                                                break;

                                            case "4000":

                                                TWD = 2990;

                                                break;

                                            default:

                                                int z = Convert.ToInt32(ds.Tables[9].Rows[x]["V2"].ToString());

                                                TWD = Convert.ToInt32(Math.Floor(z / 1.3));

                                                break;

                                        }
                                        break;
                                    case "14":

                                        if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                        {
                                            CauseName = "Ｇ吞食Ｑ鑽　　　　";
                                        }
                                        else
                                        {
                                            CauseName = "Ｇｏｏｇｌｅ　　　";
                                        }

                                        switch (ds.Tables[9].Rows[x]["V2"].ToString())
                                        {
                                            case "64":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "11")
                                                {
                                                    TWD = 60;
                                                }

                                                else if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 60;
                                                }

                                                else
                                                {
                                                    TWD = 60;
                                                }

                                                break;

                                            case "168":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "11")
                                                {
                                                    TWD = 150;
                                                }

                                                else if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 150;
                                                }

                                                else
                                                {
                                                    TWD = 150;
                                                }

                                                break;

                                            case "336":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "11")
                                                {
                                                    TWD = 300;
                                                }

                                                else if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 300;
                                                }

                                                else
                                                {
                                                    TWD = 300;
                                                }

                                                break;

                                            case "540":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "11")
                                                {
                                                    TWD = 450;
                                                }

                                                else if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 450;
                                                }

                                                else
                                                {
                                                    TWD = 450;
                                                }

                                                break;

                                            case "958":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "11")
                                                {
                                                    TWD = 750;
                                                }

                                                else if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 750;
                                                }

                                                else
                                                {
                                                    TWD = 750;
                                                }

                                                break;

                                            case "4000":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "11")
                                                {
                                                    TWD = 2990;
                                                }

                                                else if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 2990;
                                                }

                                                else
                                                {
                                                    TWD = 2990;
                                                }

                                                break;

                                            default:

                                                int z = Convert.ToInt32(ds.Tables[9].Rows[x]["V2"].ToString());


                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "11")
                                                {
                                                    TWD = Convert.ToInt32(Math.Floor(z / 1.3));
                                                }

                                                else if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = Convert.ToInt32(Math.Floor(z / 1.3));
                                                }

                                                else
                                                {
                                                    TWD = Convert.ToInt32(Math.Floor(z / 1.3));
                                                }

                                                break;
                                        }
                                        break;
                                    
                                    case "22":

                                        if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                        {
                                            CauseName = "Ｍ吞食Ｑ鑽　　　　";
                                        }
                                        else
                                        {
                                            CauseName = "ＭｙＣａｒｄ　　　";
                                        }

                                        switch (ds.Tables[9].Rows[x]["V2"].ToString())
                                        {
                                            case "53":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 50;
                                                }

                                                else
                                                {
                                                    TWD = 50;
                                                }

                                                break;

                                            case "168":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 150;
                                                }

                                                else
                                                {
                                                    TWD = 150;
                                                }

                                                break;

                                            case "336":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 300;
                                                }

                                                else
                                                {
                                                    TWD = 300;
                                                }

                                                break;

                                            case "403":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 350;
                                                }

                                                else
                                                {
                                                    TWD = 350;
                                                }

                                                break;

                                            case "480":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 400;
                                                }

                                                else
                                                {
                                                    TWD = 400;
                                                }

                                                break;

                                            case "540":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 450;
                                                }

                                                else
                                                {
                                                    TWD = 450;
                                                }

                                                break;

                                            case "600":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 500;
                                                }

                                                else
                                                {
                                                    TWD = 500;
                                                }

                                                break;

                                            case "1270":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 1000;
                                                }

                                                else
                                                {
                                                    TWD = 1000;
                                                }

                                                break;

                                            case "1460":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 1150;
                                                }

                                                else
                                                {
                                                    TWD = 1150;
                                                }

                                                break;

                                            case "2600":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 2000;
                                                }

                                                else
                                                {
                                                    TWD = 2000;
                                                }

                                                break;

                                            case "4000":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 3000;
                                                }

                                                else
                                                {
                                                    TWD = 3000;
                                                }

                                                break;

                                            case "6750":

                                                if (ds.Tables[9].Rows[x]["V4"].ToString() == "17")
                                                {
                                                    TWD = 5000;
                                                }

                                                else
                                                {
                                                    TWD = 5000;
                                                }

                                                break;

                                        }

                                        break;

                                    default:

                                        CauseName = "Cause編號 : " + ds.Tables[9].Rows[x]["Cause"].ToString();

                                        TWD = Convert.ToInt32(ds.Tables[9].Rows[x]["V2"].ToString());

                                        break;


                                    #endregion
                                }

                                richTextBox_ChargeList.Text += CauseName + "-玩家ID-" + ds.Tables[9].Rows[x]["UserID"].ToString() + "-時間-" + (Convert.ToDateTime(ds.Tables[9].Rows[x]["SaveDate"])).ToString("yyyy/MM/dd HH:mm:ss") + "-儲鑽數-" + Convert.ToInt32(ds.Tables[9].Rows[x]["V2"]).ToString("0000") + "-台幣-" + TWD + "\r\n";

                            }


                        }
                        catch
                        {
                            MessageBox.Show("資料查詢錯誤");

                        }

                        finally
                        {



                        }

                        break;


                    case 2:

                        label_Country.Text = "目前查詢國家:  香港";

                        try
                        {
                            //玩家等級分析
                            string Tsql = "SELECT Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 1) ORDER BY SaveDate DESC, CAST(V1 AS int), CAST(V2 AS int);";

                            //玩家名聲分析
                            Tsql += "SELECT  Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE  (Kind = 75) AND (Cause = 2) AND (V1 = 1 OR V1 = 2 OR V1 = 3 OR V1 = 4 OR V1 = 5 OR V1 = 6) ORDER BY V1, SaveDate DESC, CAST(V2 AS int);";

                            //Vip等級
                            Tsql += "SELECT  Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE  (Kind = 75) AND (Cause = 3)  ORDER BY CAST(V1 AS int), SaveDate DESC,CAST(V2 AS int);";

                            //玩家現有鑽石數
                            Tsql += "SELECT Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 4) AND (V1 = 1 OR V1 = 2 OR V1 = 3 OR V1 = 4 OR V1 = 5 OR V1 = 6) ORDER BY V1, SaveDate DESC, CAST(V2 AS int);";

                            //登入方式
                            Tsql += "SELECT Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 5) ORDER BY SaveDate DESC,CAST(V2 AS int);";

                            //卡包分析
                            Tsql += "SELECT Cause, COUNT(UserID) AS Expr1 FROM Egg_" + PickTime + " WHERE (Kind = 78) AND (Cause = 1 OR Cause = 2 OR Cause = 3 OR Cause = 4 OR Cause = 6 OR Cause = 7 OR Cause = 9) GROUP BY Cause ORDER BY CAST(Cause AS int);";
                            Tsql += "SELECT V1, COUNT(UserID) AS Expr1 FROM Egg_" + PickTime + " WHERE (Kind = 78) AND (Cause = 5) OR (Kind = 78) AND (Cause = 8) OR (Kind = 78) AND (Cause = 10) GROUP BY V1 ORDER BY CAST(V1 AS int);";

                            //活動積分分析
                            Tsql += "SELECT UserID, MAX(DISTINCT CAST(V1 AS int)) AS Expr1 FROM Role_" + PickTime + " WHERE (Kind = 61) GROUP BY UserID HAVING (MAX(DISTINCT CAST(V1 AS int)) > " + SportPoint + ") ORDER BY MAX(DISTINCT CAST(V1 AS int)) DESC;";

                            //創角人物分析
                            Tsql += "SELECT Kind, Cause, V1, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 6 OR Cause = 7 OR Cause = 8) ORDER BY SaveDate , CAST(V1 AS int), Cause";

                            textBox_DisLv.Text = "";
                            textBox_DisDimon.Text = "";
                            textBox_DisLog.Text = "";
                            textBox_DisName.Text = "";
                            textBox_DisVip.Text = "";
                            textBox_CardBox.Text = "";
                            textBox_Point25W.Text = "";
                            textBox_Creat.Text = "";

                            SqlLink_HK.SQLLink HKsql = new SqlLink_HK.SQLLink();
                            DataSet ds = new DataSet();
                            ds = HKsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);

                            //玩家等級分析
                            int LvCount = 0;
                            bool LVSwitch = true;

                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20130913) //舊協定
                            {
                                for (int x = 0; x < ds.Tables[0].Rows.Count; x++)
                                {
                                    if ((x % 13) == 0)
                                    {
                                        textBox_DisLv.Text += "\r\n";
                                    }

                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V2"]))
                                    {
                                        case 1:
                                            Description = "3 級以下: ";
                                            break;

                                        case 2:
                                            Description = "5 級以下: ";
                                            break;

                                        case 3:
                                            Description = "10 級: ";
                                            break;

                                        case 4:
                                            Description = "15 級: ";
                                            break;

                                        case 5:
                                            Description = "20 級: ";
                                            break;

                                        case 6:
                                            Description = "30 級: ";
                                            break;

                                        case 7:
                                            Description = "40 級: ";
                                            break;

                                        case 8:
                                            Description = "50 級: ";
                                            break;

                                        case 9:
                                            Description = "60 級: ";
                                            break;

                                        case 10:
                                            Description = "70 級: ";
                                            break;

                                        case 11:
                                            Description = "80 級: ";
                                            break;

                                        case 12:
                                            Description = "90 級: ";
                                            break;

                                        case 13:
                                            Description = "超過90 級: ";
                                            break;
                                    }

                                    textBox_DisLv.Text += "Server (" + ds.Tables[0].Rows[x]["V1"].ToString() + ") : " + Description + "  " + ds.Tables[0].Rows[x]["V3"].ToString() + "\r\n";
                                }
                            }

                            else if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130912 && Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20131128) //新協定.v1
                            {
                                string RLv = "";

                                for (int x = 0; x < ds.Tables[0].Rows.Count; x++) //玩家等級分析
                                {
                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]))
                                    {
                                        case 1:
                                            RLv = "3 級以下: ";
                                            break;

                                        case 2:
                                            RLv = "5 級以下: ";
                                            break;

                                        case 3:
                                            RLv = "10 級: ";
                                            break;

                                        case 4:
                                            RLv = "15 級: ";
                                            break;

                                        case 5:
                                            RLv = "20 級: ";
                                            break;

                                        case 6:
                                            RLv = "30 級: ";
                                            break;

                                        case 7:
                                            RLv = "40 級: ";
                                            break;

                                        case 8:
                                            RLv = "50 級: ";
                                            break;

                                        case 9:
                                            RLv = "60 級: ";
                                            break;

                                        case 10:
                                            RLv = "70 級: ";
                                            break;

                                        case 11:
                                            RLv = "80 級: ";
                                            break;

                                        case 12:
                                            RLv = "90 級: ";
                                            break;

                                        case 13:
                                            RLv = "100 級: ";
                                            break;

                                        case 14:
                                            RLv = "115 級: ";
                                            break;

                                        case 15:
                                            RLv = "130 級: ";
                                            break;

                                        case 16:
                                            RLv = "150以下: ";
                                            break;

                                        case 17:
                                            RLv = "超過150 級: ";
                                            break;
                                    }

                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V2"]))
                                    {
                                        case 0:
                                            Description = "其它: ";
                                            break;

                                        case 1:
                                            Description = "一般 Apple: ";
                                            break;

                                        case 2:
                                            Description = "完整 Apple: ";
                                            break;

                                        case 3:
                                            Description = "一般 Goole: ";
                                            break;

                                        case 4:
                                            Description = "完整 Goole: ";
                                            break;

                                        case 5:
                                            Description = "一般 MyCard: ";
                                            break;

                                        case 6:
                                            Description = "完整 MyCard: ";
                                            break;

                                        case 7:
                                            Description = "一般台灣老李 MyCard: ";
                                            break;

                                        case 8:
                                            Description = "完整台灣老李 MyCard: ";
                                            break;

                                        case 9:
                                            Description = "一般 Android: ";
                                            break;

                                        case 10:
                                            Description = "完整 Android: ";
                                            break;

                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) - LvCount == 1)
                                    {
                                        textBox_DisLv.Text += "\r\n" + "\r\n" + " 玩家等級: " + RLv + "\r\n" + "\r\n";
                                        LvCount++;
                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) >= LvCount && LVSwitch == true)
                                    {
                                        textBox_DisLv.Text += Description + "  " + ds.Tables[0].Rows[x]["V3"].ToString() + "\r\n";
                                    }
                                    else
                                    {
                                        LVSwitch = false;
                                    }

                                }
                            }

                            else  //新協定.v2
                            {
                                string RLv = "";
                                for (int x = 0; x < ds.Tables[0].Rows.Count; x++) //玩家等級分析
                                {
                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]))
                                    {
                                        case 1:
                                            RLv = "3 級以下: ";
                                            break;

                                        case 2:
                                            RLv = "5 級以下: ";
                                            break;

                                        case 3:
                                            RLv = "10 級: ";
                                            break;

                                        case 4:
                                            RLv = "15 級: ";
                                            break;

                                        case 5:
                                            RLv = "20 級: ";
                                            break;

                                        case 6:
                                            RLv = "30 級: ";
                                            break;

                                        case 7:
                                            RLv = "40 級: ";
                                            break;

                                        case 8:
                                            RLv = "50 級: ";
                                            break;

                                        case 9:
                                            RLv = "60 級: ";
                                            break;

                                        case 10:
                                            RLv = "70 級: ";
                                            break;

                                        case 11:
                                            RLv = "80 級: ";
                                            break;

                                        case 12:
                                            RLv = "90 級: ";
                                            break;

                                        case 13:
                                            RLv = "100 級: ";
                                            break;

                                        case 14:
                                            RLv = "115 級: ";
                                            break;

                                        case 15:
                                            RLv = "130 級: ";
                                            break;

                                        case 16:
                                            RLv = "150以下: ";
                                            break;

                                        case 17:
                                            RLv = "超過150 級: ";
                                            break;
                                    }

                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V2"]))
                                    {
                                        case 0:
                                            Description = "Ios 其他: ";
                                            break;

                                        case 1:
                                            Description = "香港 IOS: ";
                                            break;

                                        case 2:
                                            Description = "香港 IOS 完整版: ";
                                            break;

                                        case 3:
                                            Description = "台灣 IOS 完整版: ";
                                            break;

                                        case 4:
                                            Description = "台灣 IOS: ";
                                            break;

                                        case 5:
                                            Description = "大陸 IOS 完整版: ";
                                            break;

                                        case 6:
                                            Description = "Android 其他: ";
                                            break;

                                        case 7:
                                            Description = "香港 MyCard: ";
                                            break;

                                        case 8:
                                            Description = "香港 MyCard 完整版: ";
                                            break;

                                        case 9:
                                            Description = "香港 Google: ";
                                            break;

                                        case 10:
                                            Description = "台灣 MyCard 完整版: ";
                                            break;

                                        case 11:
                                            Description = "台灣 Google: ";
                                            break;

                                        case 12:
                                            Description = "大陸 Android 完整版 (未): ";
                                            break;

                                        case 13:
                                            Description = "IOS Debug: ";
                                            break;

                                        case 14:
                                            Description = "Android Debug: ";
                                            break;

                                        case 15:
                                            Description = "台灣老李 MyCard: ";
                                            break;

                                        case 16:
                                            Description = "台灣 MyCard: ";
                                            break;

                                        case 17:
                                            Description = "台灣 Google 完整版: ";
                                            break;

                                        case 18:
                                            Description = "台灣老李 MyCard 完整版: ";
                                            break;

                                        case 19:
                                            Description = "香港 Google 完整版: ";
                                            break;

                                        case 20:
                                            Description = "松崗: ";
                                            break;

                                        case 21:
                                            Description = "新加坡 IOS: ";
                                            break;

                                        case 22:
                                            Description = "新加坡 Android 完整版: ";
                                            break;

                                        case 23:
                                            Description = "新加坡 Android 當地金流: ";
                                            break;

                                        case 24:
                                            Description = "大陸91 IOS: ";
                                            break;

                                        case 25:
                                            Description = "大陸UC Android: ";
                                            break;

                                        case 26:
                                            Description = "大陸360 Android: ";
                                            break;

                                        case 27:
                                            Description = "大陸App助手 IOS: ";
                                            break;

                                        case 28:
                                            Description = "大陸官方apk Android: ";
                                            break;

                                        case 29:
                                            Description = "大陸91 Android: ";
                                            break;

                                        case 30:
                                            Description = "馬幹線 IOS: ";
                                            break;

                                        case 31:
                                            Description = "馬幹線 Android: ";
                                            break;

                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) - LvCount == 1)
                                    {
                                        textBox_DisLv.Text += "\r\n" + "\r\n" + " 玩家等級: " + RLv + "\r\n" + "\r\n";
                                        LvCount++;
                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) >= LvCount && LVSwitch == true)
                                    {
                                        textBox_DisLv.Text += Description + "  " + ds.Tables[0].Rows[x]["V3"].ToString() + "\r\n";
                                    }
                                    else
                                    {
                                        LVSwitch = false;
                                    }

                                }
                            }


                            //玩家名聲分析

                            int y = 0;

                            for (int x = 0; x < ds.Tables[1].Rows.Count; x++)
                            {

                                if (Convert.ToInt32(ds.Tables[1].Rows[x]["V1"]) - y == 1)
                                {
                                    textBox_DisName.Text += "\r\n" + "伺服器: " + ds.Tables[1].Rows[x]["V1"].ToString() + "\r\n" + "\r\n";
                                    y++;
                                }

                                switch (Convert.ToInt32(ds.Tables[1].Rows[x]["V2"]))
                                {

                                    case 1:
                                        Description = "1000 以下: ";
                                        break;

                                    case 2:
                                        Description = "2000: ";
                                        break;

                                    case 3:
                                        Description = "5000: ";
                                        break;

                                    case 4:
                                        Description = "10000: ";
                                        break;

                                    case 5:
                                        Description = "20000: ";
                                        break;

                                    case 6:
                                        Description = "50000: ";
                                        break;

                                    case 7:
                                        Description = "100000: ";
                                        break;

                                    case 8:
                                        Description = "200000: ";
                                        break;

                                    case 9:
                                        Description = "250000: ";
                                        break;

                                    case 10:
                                        Description = "300000: ";
                                        break;

                                    case 11:
                                        Description = "350000: ";
                                        break;

                                    case 12:
                                        Description = "400000: ";
                                        break;

                                    case 13:
                                        Description = "450000: ";
                                        break;

                                    case 14:
                                        Description = "500000: ";
                                        break;

                                    case 15:
                                        Description = "超過 500000: ";
                                        break;

                                    case 16:
                                        Description = "1000000: ";
                                        break;

                                    case 17:
                                        Description = "1500000: ";
                                        break;
                                }

                                textBox_DisName.Text += Description + "  " + ds.Tables[1].Rows[x]["V3"].ToString() + "\r\n";

                            }




                            //Vip分析

                            int VipCount = -1;

                            for (int x = 0; x < ds.Tables[2].Rows.Count; x++)
                            {
                                if (Convert.ToInt32(ds.Tables[2].Rows[x]["V1"]) - VipCount == 1)
                                {
                                    textBox_DisVip.Text += "\r\n" + "Vip等級: " + ds.Tables[2].Rows[x]["V1"].ToString() + "\r\n" + "\r\n";
                                    VipCount++;
                                }

                                if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20130912) //舊協定
                                {
                                    switch (Convert.ToInt32(ds.Tables[2].Rows[x]["V2"].ToString()))
                                    {

                                        case 0:
                                            Description = "其他: ";
                                            break;

                                        case 1:
                                            Description = "Apple: ";
                                            break;

                                        case 2:
                                            Description = "一般 My Card: ";
                                            break;

                                        case 3:
                                            Description = "完整 My Card: ";
                                            break;

                                        case 4:
                                            Description = "Google: ";
                                            break;

                                        case 5:
                                            Description = "東方阿李: ";
                                            break;
                                    }
                                }

                                else if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130911 && Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20131128) //舊協定
                                {
                                    switch (Convert.ToInt32(ds.Tables[2].Rows[x]["V2"].ToString()))
                                    {

                                        case 0:
                                            Description = "其他: ";
                                            break;

                                        case 1:
                                            Description = "一般 Apple: ";
                                            break;

                                        case 2:
                                            Description = "完整 Apple: ";
                                            break;

                                        case 3:
                                            Description = "一般 Google: ";
                                            break;

                                        case 4:
                                            Description = "完整 Google: ";
                                            break;

                                        case 5:
                                            Description = "一般 MyCard: ";
                                            break;

                                        case 6:
                                            Description = "完整 MyCard: ";
                                            break;

                                        case 7:
                                            Description = "一般台灣老李MyCard: ";
                                            break;

                                        case 8:
                                            Description = "完整台灣老李 MyCard: ";
                                            break;

                                        case 9:
                                            Description = "一般 Android: ";
                                            break;

                                        case 10:
                                            Description = "完整 Android: ";
                                            break;
                                    }
                                }

                                else
                                {
                                    switch (Convert.ToInt32(ds.Tables[2].Rows[x]["V2"].ToString()))
                                    {

                                        case 0:
                                            Description = "Ios 其他: ";
                                            break;

                                        case 1:
                                            Description = "香港 IOS: ";
                                            break;

                                        case 2:
                                            Description = "香港 IOS 完整版: ";
                                            break;

                                        case 3:
                                            Description = "台灣 IOS 完整版: ";
                                            break;

                                        case 4:
                                            Description = "台灣 IOS: ";
                                            break;

                                        case 5:
                                            Description = "大陸 IOS 完整版: ";
                                            break;

                                        case 6:
                                            Description = "Android 其他: ";
                                            break;

                                        case 7:
                                            Description = "香港 MyCard: ";
                                            break;

                                        case 8:
                                            Description = "香港 MyCard 完整版: ";
                                            break;

                                        case 9:
                                            Description = "香港 Google: ";
                                            break;

                                        case 10:
                                            Description = "台灣 MyCard 完整版: ";
                                            break;

                                        case 11:
                                            Description = "台灣 Google: ";
                                            break;

                                        case 12:
                                            Description = "大陸 Android 完整版 (未): ";
                                            break;

                                        case 13:
                                            Description = "IOS Debug: ";
                                            break;

                                        case 14:
                                            Description = "Android Debug: ";
                                            break;

                                        case 15:
                                            Description = "台灣老李 MyCard: ";
                                            break;

                                        case 16:
                                            Description = "台灣 MyCard: ";
                                            break;

                                        case 17:
                                            Description = "台灣 Google 完整版: ";
                                            break;

                                        case 18:
                                            Description = "台灣老李 MyCard 完整版: ";
                                            break;

                                        case 19:
                                            Description = "香港 Google 完整版: ";
                                            break;

                                        case 20:
                                            Description = "松崗: ";
                                            break;

                                        case 21:
                                            Description = "新加坡 IOS: ";
                                            break;

                                        case 22:
                                            Description = "新加坡 Android 完整版: ";
                                            break;

                                        case 23:
                                            Description = "新加坡 Android 當地金流: ";
                                            break;

                                        case 24:
                                            Description = "大陸91 IOS: ";
                                            break;

                                        case 25:
                                            Description = "大陸UC Android: ";
                                            break;

                                        case 26:
                                            Description = "大陸360 Android: ";
                                            break;

                                        case 27:
                                            Description = "大陸App助手 IOS: ";
                                            break;

                                        case 28:
                                            Description = "大陸官方apk Android: ";
                                            break;

                                        case 29:
                                            Description = "大陸91 Android: ";
                                            break;

                                        case 30:
                                            Description = "馬幹線 IOS: ";
                                            break;

                                        case 31:
                                            Description = "馬幹線 Android: ";
                                            break;
                                    }
                                }

                                textBox_DisVip.Text += Description + "  " + ds.Tables[2].Rows[x]["V3"].ToString() + "\r\n";
                            }



                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130905)
                            {
                                //玩家現有鑽石分析

                                int DimonCount = 0;

                                for (int x = 0; x < ds.Tables[3].Rows.Count; x++)
                                {
                                    if (Convert.ToInt32(ds.Tables[3].Rows[x]["V1"]) - DimonCount == 1)
                                    {
                                        textBox_DisDimon.Text += "\r\n" + "伺服器: " + ds.Tables[3].Rows[x]["V1"].ToString() + "\r\n" + "\r\n";
                                        DimonCount++;
                                    }

                                    switch (Convert.ToInt32(ds.Tables[3].Rows[x]["V2"]))
                                    {
                                        case 0:
                                            Description = "60 以下: ";
                                            break;

                                        case 1:
                                            Description = "500: ";
                                            break;

                                        case 2:
                                            Description = "1000: ";
                                            break;

                                        case 3:
                                            Description = "2000: ";
                                            break;

                                        case 4:
                                            Description = "5000: ";
                                            break;

                                        case 5:
                                            Description = "10000: ";
                                            break;

                                        case 6:
                                            Description = "20000: ";
                                            break;

                                        case 7:
                                            Description = "50000: ";
                                            break;

                                        case 8:
                                            Description = "100000: ";
                                            break;

                                        case 9:
                                            Description = "10萬到50萬: ";
                                            break;

                                        case 10:
                                            Description = "500000以上: ";
                                            break;
                                    }

                                    textBox_DisDimon.Text += Description + "  " + ds.Tables[3].Rows[x]["V3"].ToString() + "\r\n";
                                }

                            }
                            else
                            {
                                textBox_DisDimon.Text = "資料庫無資料可分析";
                            }


                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130904)
                            {
                                for (int j = 0; j < 1; j++) //登入方式
                                {
                                    textBox_DisLog.Text += "\r\n" + "登入方式: " + "\r\n" + "\r\n";

                                    for (int x = 0; x < ds.Tables[j + 4].Rows.Count; x++)
                                    {
                                        switch (Convert.ToInt32(ds.Tables[j + 4].Rows[x]["V2"]))
                                        {

                                            case 0:
                                                Description = "其它: ";
                                                break;

                                            case 1:
                                                Description = "香港一般版: ";
                                                break;

                                            case 2:
                                                Description = "香港 IOS 完整版: ";
                                                break;

                                            case 3:
                                                Description = "台灣 IOS 完整版: ";
                                                break;

                                            case 4:
                                                Description = "台灣一般版: ";
                                                break;

                                            case 5:
                                                Description = "阿六版: ";
                                                break;

                                            case 6:
                                                Description = "未登錄: ";
                                                break;

                                            case 7:
                                                Description = "香港一般MYCARD: ";
                                                break;

                                            case 8:
                                                Description = "香港完整MYCARD: ";
                                                break;

                                            case 9:
                                                Description = "香港一般Google: ";
                                                break;

                                            case 10:
                                                Description = "台灣完整MYCARD: ";
                                                break;

                                            case 11:
                                                Description = "台灣一般google: ";
                                                break;

                                            case 12:
                                                Description = "大陸 Android 完整版 (未): ";
                                                break;

                                            case 13:
                                                Description = "IOS Debug: ";
                                                break;

                                            case 14:
                                                Description = "Android Debug: ";
                                                break;

                                            case 15:
                                                Description = "東方阿李: ";
                                                break;

                                            case 16:
                                                Description = "台灣 MyCard: ";
                                                break;

                                            case 17:
                                                Description = "台灣 Google 完整版: ";
                                                break;

                                            case 18:
                                                Description = "台灣老李 MyCard 完整版: ";
                                                break;

                                            case 19:
                                                Description = "香港 Google 完整版: ";
                                                break;

                                            case 20:
                                                Description = "松崗: ";
                                                break;

                                            case 21:
                                                Description = "新加坡 IOS: ";
                                                break;

                                            case 22:
                                                Description = "新加坡 Android 完整版: ";
                                                break;

                                            case 23:
                                                Description = "新加坡 Android 當地金流: ";
                                                break;

                                            case 24:
                                                Description = "大陸91 IOS: ";
                                                break;

                                            case 25:
                                                Description = "大陸UC Android: ";
                                                break;

                                            case 26:
                                                Description = "大陸360 Android: ";
                                                break;

                                            case 27:
                                                Description = "大陸App助手 IOS: ";
                                                break;

                                            case 28:
                                                Description = "大陸官方apk Android: ";
                                                break;

                                            case 29:
                                                Description = "大陸91 Android: ";
                                                break;

                                            case 30:
                                                Description = "馬幹線 IOS: ";
                                                break;

                                            case 31:
                                                Description = "馬幹線 Android: ";
                                                break;
                                        }

                                        textBox_DisLog.Text += Description + "  " + ds.Tables[j + 4].Rows[x]["V3"].ToString() + "\r\n" + "\r\n";
                                    }
                                }
                            }
                            else
                            {
                                textBox_DisLog.Text = "資料庫無資料可分析";
                            }

                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130905)
                            {
                                //卡包分析

                                textBox_CardBox.Text += "\r\n";

                                for (int x = 0; x < ds.Tables[5].Rows.Count; x++)
                                {
                                    switch (Convert.ToInt32(ds.Tables[5].Rows[x]["Cause"]))
                                    {
                                        case 0:
                                            Description = "聖旨單抽: ";
                                            break;

                                        case 1:
                                            Description = "友情單抽: ";
                                            break;

                                        case 2:
                                            Description = "銀幣單抽: ";
                                            break;

                                        case 3:
                                            Description = "鑽石單抽: ";
                                            break;

                                        case 6:
                                            Description = "友情10連抽: ";
                                            break;

                                        case 7:
                                            Description = "銀幣10連抽: ";
                                            break;

                                        case 9:
                                            Description = "９鑽抽:     ";
                                            break;
                                    }

                                    textBox_CardBox.Text += Description + "  " + ds.Tables[5].Rows[x]["Expr1"].ToString() + "\r\n" + "\r\n";
                                }


                                textBox_CardBox.Text += "\r\n";

                                for (int x = 0; x < ds.Tables[6].Rows.Count; x++) //卡包分析
                                {
                                    switch (Convert.ToInt32(ds.Tables[6].Rows[x]["V1"]))
                                    {
                                        case 4:
                                            Description = "必中卡包: ";
                                            break;

                                        case 7:
                                            Description = "一般卡包: ";
                                            break;

                                        case 11:
                                            Description = "卡包1(吳國太): ";
                                            break;

                                        case 12:
                                            Description = "卡包2(甘夫人): ";
                                            break;

                                        case 13:
                                            Description = "卡包3(蔡文姬): ";
                                            break;

                                        case 14:
                                            Description = "卡包4(夏日美女包): ";
                                            break;

                                        case 15:
                                            Description = "卡包5(闇屬趙雲): ";
                                            break;

                                        case 16:
                                            Description = "卡包6(光屬張飛): ";
                                            break;

                                        case 17:
                                            Description = "卡包7(蜀國勢力包:龐統必中): ";
                                            break;

                                        case 18:
                                            Description = "卡包8(吳國勢力包:龐統必中): ";
                                            break;

                                        case 19:
                                            Description = "卡包9(魏國勢力包:司馬懿): ";
                                            break;

                                        case 20:
                                            Description = "卡包10(他國勢力包:水呂布): ";
                                            break;

                                        case 21:
                                            Description = "卡包11: ";
                                            break;

                                        case 22:
                                            Description = "卡包12: ";
                                            break;

                                        case 23:
                                            Description = "卡包13: ";
                                            break;

                                        case 24:
                                            Description = "卡包14: ";
                                            break;

                                        case 25:
                                            Description = "卡包15: ";
                                            break;

                                        case 26:
                                            Description = "卡包16: ";
                                            break;

                                        case 27:
                                            Description = "卡包17: ";
                                            break;

                                        case 28:
                                            Description = "卡包18: ";
                                            break;

                                        case 29:
                                            Description = "卡包19: ";
                                            break;

                                        case 30:
                                            Description = "卡包20: ";
                                            break;


                                    }

                                    textBox_CardBox.Text += Description + "  " + ds.Tables[6].Rows[x]["Expr1"].ToString() + "\r\n" + "\r\n";
                                }
                            }
                            else
                            {
                                textBox_CardBox.Text = "資料庫無資料可分析";
                            }

                            //活動積分分析
                            textBox_Point25W.Text += "\r\n";
                            int ID = 0;
                            for (int x = 0; x < ds.Tables[7].Rows.Count; x++)
                            {
                                ID++;
                                textBox_Point25W.Text += " (" + ID + "). " + "玩家ID: " + ds.Tables[7].Rows[x]["UserID"].ToString() + "     積分: " + ds.Tables[7].Rows[x]["Expr1"].ToString() + "\r\n" + "\r\n";
                            }

                            if (textBox_Point25W.Text == "\r\n")
                            {
                                textBox_Point25W.Text = "無對應此積分資料";
                            }

                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130917)
                            {
                                if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20131129)
                                {
                                    #region CreatCharacter

                                    //當日創角
                                    textBox_Creat.Text += "當日創角人數 /" + "\r\n" + "當日創角且超過5等 /" + "\r\n" + "前天創角且連續兩天有上線" + "\r\n" + "\r\n";

                                    #region 數字空間
                                    string Num6_0 = "";
                                    string Num6_1 = "";
                                    string Num6_2 = "";
                                    string Num6_3 = "";
                                    string Num6_4 = "";
                                    string Num6_5 = "";
                                    string Num6_6 = "";
                                    string Num6_7 = "";
                                    string Num6_8 = "";
                                    string Num6_9 = "";
                                    string Num6_10 = "";

                                    string Num7_0 = "";
                                    string Num7_1 = "";
                                    string Num7_2 = "";
                                    string Num7_3 = "";
                                    string Num7_4 = "";
                                    string Num7_5 = "";
                                    string Num7_6 = "";
                                    string Num7_7 = "";
                                    string Num7_8 = "";
                                    string Num7_9 = "";
                                    string Num7_10 = "";

                                    string Num8_0 = "";
                                    string Num8_1 = "";
                                    string Num8_2 = "";
                                    string Num8_3 = "";
                                    string Num8_4 = "";
                                    string Num8_5 = "";
                                    string Num8_6 = "";
                                    string Num8_7 = "";
                                    string Num8_8 = "";
                                    string Num8_9 = "";
                                    string Num8_10 = "";
                                    #endregion

                                    for (int x = 0; x < ds.Tables[8].Rows.Count; x++)
                                    {
                                        switch (Convert.ToInt32(ds.Tables[8].Rows[x]["Cause"].ToString()))
                                        {
                                            case 6:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"其他:　　　　　　　　　　　";
                                                        Num6_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"一般Ａｐｐｌｅ:　　　　　　";
                                                        Num6_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"完整Ａｐｐｌｅ:　　　　　　";
                                                        Num6_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"一般Ｇｏｏｇｌｅ:　　　　　";
                                                        Num6_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"完整Ｇｏｏｇｌｅ:　　　　　";
                                                        Num6_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"一般ＭｙＣａｒｄ:　　　　　";
                                                        Num6_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"完整ＭｙＣａｒｄ:　　　　　";
                                                        Num6_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"一般台灣老李ＭｙＣａｒｄ:　";
                                                        Num6_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"完整台灣老李ＭｙＣａｒｄ:　";
                                                        Num6_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"一般Ａｎｄｒｏｉｄ:　　　　";
                                                        Num6_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"完整Ａｎｄｒｏｉｄ:　　　　";
                                                        Num6_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                            case 7:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"其他:　　　　　　　　　　　";
                                                        Num7_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"一般Ａｐｐｌｅ:　　　　　　";
                                                        Num7_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"完整Ａｐｐｌｅ:　　　　　　";
                                                        Num7_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"一般Ｇｏｏｇｌｅ:　　　　　";
                                                        Num7_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"完整Ｇｏｏｇｌｅ:　　　　　";
                                                        Num7_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"一般ＭｙＣａｒｄ:　　　　　";
                                                        Num7_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"完整ＭｙＣａｒｄ:　　　　　";
                                                        Num7_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"一般台灣老李ＭｙＣａｒｄ:　";
                                                        Num7_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"完整台灣老李ＭｙＣａｒｄ:　";
                                                        Num7_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"一般Ａｎｄｒｏｉｄ:　　　　";
                                                        Num7_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"完整Ａｎｄｒｏｉｄ:　　　　";
                                                        Num7_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }


                                                break;

                                            case 8:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"其他:　　　　　　　　　　　";
                                                        Num8_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"一般Ａｐｐｌｅ:　　　　　　";
                                                        Num8_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"完整Ａｐｐｌｅ:　　　　　　";
                                                        Num8_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"一般Ｇｏｏｇｌｅ:　　　　　";
                                                        Num8_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"完整Ｇｏｏｇｌｅ:　　　　　";
                                                        Num8_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"一般ＭｙＣａｒｄ:　　　　　";
                                                        Num8_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"完整ＭｙＣａｒｄ:　　　　　";
                                                        Num8_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"一般台灣老李ＭｙＣａｒｄ:　";
                                                        Num8_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"完整台灣老李ＭｙＣａｒｄ:　";
                                                        Num8_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"一般Ａｎｄｒｏｉｄ:　　　　";
                                                        Num8_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"完整Ａｎｄｒｏｉｄ:　　　　";
                                                        Num8_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                        }

                                    }

                                    textBox_Creat.Text += "其他:　　　　　　　　　　　" + Num6_0 + " / " + Num7_0 + " / " + Num8_0 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般Ａｐｐｌｅ:　　　　　　" + Num6_1 + " / " + Num7_1 + " / " + Num8_1 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整Ａｐｐｌｅ:　　　　　　" + Num6_2 + " / " + Num7_2 + " / " + Num8_2 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般Ｇｏｏｇｌｅ:　　　　　" + Num6_3 + " / " + Num7_3 + " / " + Num8_3 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整Ｇｏｏｇｌｅ:　　　　　" + Num6_4 + " / " + Num7_4 + " / " + Num8_4 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般ＭｙＣａｒｄ:　　　　　" + Num6_5 + " / " + Num7_5 + " / " + Num8_5 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整ＭｙＣａｒｄ:　　　　　" + Num6_6 + " / " + Num7_6 + " / " + Num8_6 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般台灣老李ＭｙＣａｒｄ:　" + Num6_7 + " / " + Num7_7 + " / " + Num8_7 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整台灣老李ＭｙＣａｒｄ:　" + Num6_8 + " / " + Num7_8 + " / " + Num8_8 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般Ａｎｄｒｏｉｄ:　　　　" + Num6_9 + " / " + Num7_9 + " / " + Num8_9 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整Ａｎｄｒｏｉｄ:　　　　" + Num6_10 + " / " + Num7_10 + " / " + Num8_10 + "\r\n" + "\r\n";

                                    #endregion
                                }

                                else
                                {
                                    #region CreatCharacter

                                    //當日創角
                                    textBox_Creat.Text += "當日創角人數 /" + "\r\n" + "當日創角且超過5等 /" + "\r\n" + "前天創角且連續兩天有上線" + "\r\n" + "\r\n";

                                    #region 數字空間
                                    string Num6_0 = "";
                                    string Num6_1 = "";
                                    string Num6_2 = "";
                                    string Num6_3 = "";
                                    string Num6_4 = "";
                                    string Num6_5 = "";
                                    string Num6_6 = "";
                                    string Num6_7 = "";
                                    string Num6_8 = "";
                                    string Num6_9 = "";
                                    string Num6_10 = "";
                                    string Num6_11 = "";
                                    string Num6_12 = "";
                                    string Num6_13 = "";
                                    string Num6_14 = "";
                                    string Num6_15 = "";
                                    string Num6_16 = "";
                                    string Num6_17 = "";
                                    string Num6_18 = "";
                                    string Num6_19 = "";
                                    string Num6_20 = "";
                                    string Num6_21 = "";
                                    string Num6_22 = "";

                                    string Num7_0 = "";
                                    string Num7_1 = "";
                                    string Num7_2 = "";
                                    string Num7_3 = "";
                                    string Num7_4 = "";
                                    string Num7_5 = "";
                                    string Num7_6 = "";
                                    string Num7_7 = "";
                                    string Num7_8 = "";
                                    string Num7_9 = "";
                                    string Num7_10 = "";
                                    string Num7_11 = "";
                                    string Num7_12 = "";
                                    string Num7_13 = "";
                                    string Num7_14 = "";
                                    string Num7_15 = "";
                                    string Num7_16 = "";
                                    string Num7_17 = "";
                                    string Num7_18 = "";
                                    string Num7_19 = "";
                                    string Num7_20 = "";
                                    string Num7_21 = "";
                                    string Num7_22 = "";

                                    string Num8_0 = "";
                                    string Num8_1 = "";
                                    string Num8_2 = "";
                                    string Num8_3 = "";
                                    string Num8_4 = "";
                                    string Num8_5 = "";
                                    string Num8_6 = "";
                                    string Num8_7 = "";
                                    string Num8_8 = "";
                                    string Num8_9 = "";
                                    string Num8_10 = "";
                                    string Num8_11 = "";
                                    string Num8_12 = "";
                                    string Num8_13 = "";
                                    string Num8_14 = "";
                                    string Num8_15 = "";
                                    string Num8_16 = "";
                                    string Num8_17 = "";
                                    string Num8_18 = "";
                                    string Num8_19 = "";
                                    string Num8_20 = "";
                                    string Num8_21 = "";
                                    string Num8_22 = "";

                                    #endregion

                                    for (int x = 0; x < ds.Tables[8].Rows.Count; x++)
                                    {
                                        switch (Convert.ToInt32(ds.Tables[8].Rows[x]["Cause"].ToString()))
                                        {
                                            case 6:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"Ｉｏｓ其他:　　　　　　　　　　　　　";
                                                        Num6_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"香港　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num6_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"香港　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num6_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"台灣　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num6_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"台灣　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num6_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"大陸　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num6_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"Ａｎｄｒｏｉｄ其他:　　　　　　　　　";
                                                        Num6_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"香港　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num6_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"香港　ＭｙＣａｒｄ完整版:　　　　　　";
                                                        Num6_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"香港　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num6_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"台灣　Ｍｙｃａｒｄ完整版:　　　　　　";
                                                        Num6_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 11:
                                                        //"台灣　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num6_11 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 12:
                                                        //"大陸　Ａｎｄｒｏｉｄ 完整版 (未):　　";
                                                        Num6_12 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 13:
                                                        //"ＩｏｓＤｅｂｕｇ:　　　　　　　　　　";
                                                        Num6_13 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 14:
                                                        //"ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　";
                                                        Num6_14 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 15:
                                                        //"台灣老李　ＭｙＣａｒｄ:　　　　　　　";
                                                        Num6_15 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 16:
                                                        //"台灣　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num6_16 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 17:
                                                        //"台灣　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num6_17 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 18:
                                                        //"台灣老李　ＭｙＣａｒｄ完整版:　　　　";
                                                        Num6_18 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 19:
                                                        //"香港　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num6_19 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 20:
                                                        //"松崗:　　　　　　　　　　　　　　　　";
                                                        Num6_20 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 21:
                                                        //"新加坡　Ｉｏｓ:　　　　　　　　　　　";
                                                        Num6_21 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 22:
                                                        //"新加坡　Ａｎｄｒｏｉｄ完整版:　　　　";
                                                        Num6_22 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                            case 7:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"Ｉｏｓ其他:　　　　　　　　　　　　　";
                                                        Num7_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"香港　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num7_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"香港　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num7_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"台灣　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num7_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"台灣　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num7_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"大陸　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num7_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"Ａｎｄｒｏｉｄ其他:　　　　　　　　　";
                                                        Num7_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"香港　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num7_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"香港　ＭｙＣａｒｄ完整版:　　　　　　";
                                                        Num7_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"香港　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num7_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"台灣　Ｍｙｃａｒｄ完整版:　　　　　　";
                                                        Num7_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 11:
                                                        //"台灣　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num7_11 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 12:
                                                        //"大陸　Ａｎｄｒｏｉｄ 完整版 (未):　　";
                                                        Num7_12 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 13:
                                                        //"ＩｏｓＤｅｂｕｇ:　　　　　　　　　　";
                                                        Num7_13 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 14:
                                                        //"ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　";
                                                        Num7_14 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 15:
                                                        //"台灣老李　ＭｙＣａｒｄ:　　　　　　　";
                                                        Num7_15 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 16:
                                                        //"台灣　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num7_16 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 17:
                                                        //"台灣　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num7_17 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 18:
                                                        //"台灣老李　ＭｙＣａｒｄ完整版:　　　　";
                                                        Num7_18 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 19:
                                                        //"香港　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num7_19 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 20:
                                                        //"松崗:　　　　　　　　　　　　　　　　";
                                                        Num7_20 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 21:
                                                        //"新加坡　Ｉｏｓ:　　　　　　　　　　　";
                                                        Num7_21 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 22:
                                                        //"新加坡　Ａｎｄｒｏｉｄ完整版:　　　　";
                                                        Num7_22 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }


                                                break;

                                            case 8:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"Ｉｏｓ其他:　　　　　　　　　　　　　";
                                                        Num8_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"香港　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num8_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"香港　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num8_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"台灣　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num8_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"台灣　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num8_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"大陸　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num8_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"Ａｎｄｒｏｉｄ其他:　　　　　　　　　";
                                                        Num8_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"香港　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num8_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"香港　ＭｙＣａｒｄ完整版:　　　　　　";
                                                        Num8_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"香港　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num8_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"台灣　Ｍｙｃａｒｄ完整版:　　　　　　";
                                                        Num8_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 11:
                                                        //"台灣　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num8_11 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 12:
                                                        //"大陸　Ａｎｄｒｏｉｄ 完整版 (未):　　";
                                                        Num8_12 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 13:
                                                        //"ＩｏｓＤｅｂｕｇ:　　　　　　　　　　";
                                                        Num8_13 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 14:
                                                        //"ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　";
                                                        Num8_14 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 15:
                                                        //"台灣老李　ＭｙＣａｒｄ:　　　　　　　";
                                                        Num8_15 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 16:
                                                        //"台灣　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num8_16 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 17:
                                                        //"台灣　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num8_17 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 18:
                                                        //"台灣老李　ＭｙＣａｒｄ完整版:　　　　";
                                                        Num8_18 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 19:
                                                        //"香港　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num8_19 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 20:
                                                        //"松崗:　　　　　　　　　　　　　　　　";
                                                        Num8_20 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 21:
                                                        //"新加坡　Ｉｏｓ:　　　　　　　　　　　";
                                                        Num8_21 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 22:
                                                        //"新加坡　Ａｎｄｒｏｉｄ完整版:　　　　";
                                                        Num8_22 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                        }

                                    }

                                    textBox_Creat.Text += "Ｉｏｓ其他:　　　　　　　　　　　　　" + Num6_0 + " / " + Num7_0 + " / " + Num8_0 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｉｏｓ:　　　　　　　　　　　　" + Num6_1 + " / " + Num7_1 + " / " + Num8_1 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｉｏｓ完整版:　　　　　　　　　" + Num6_2 + " / " + Num7_2 + " / " + Num8_2 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｉｏｓ完整版:　　　　　　　　　" + Num6_3 + " / " + Num7_3 + " / " + Num8_3 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｉｏｓ:　　　　　　　　　　　　" + Num6_4 + " / " + Num7_4 + " / " + Num8_4 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "大陸　Ｉｏｓ完整版:　　　　　　　　　" + Num6_5 + " / " + Num7_5 + " / " + Num8_5 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "Ａｎｄｒｏｉｄ其他:　　　　　　　　　" + Num6_6 + " / " + Num7_6 + " / " + Num8_6 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　ＭｙＣａｒｄ:　　　　　　　　　" + Num6_7 + " / " + Num7_7 + " / " + Num8_7 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　ＭｙＣａｒｄ完整版:　　　　　　" + Num6_8 + " / " + Num7_8 + " / " + Num8_8 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｇｏｏｇｌｅ:　　　　　　　　　" + Num6_9 + " / " + Num7_9 + " / " + Num8_9 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｍｙｃａｒｄ完整版:　　　　　　" + Num6_10 + " / " + Num7_10 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｇｏｏｇｌｅ:　　　　　　　　　" + Num6_11 + " / " + Num7_11 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "大陸　Ａｎｄｒｏｉｄ完整版（未）:　　" + Num6_12 + " / " + Num7_12 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "ＩｏｓＤｅｂｕｇ:　　　　　　　　　　" + Num6_13 + " / " + Num7_13 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　" + Num6_14 + " / " + Num7_14 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣老李　ＭｙＣａｒｄ:　　　　　　　" + Num6_15 + " / " + Num7_15 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　ＭｙＣａｒｄ:　　　　　　　　　" + Num6_16 + " / " + Num7_16 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｇｏｏｇｌｅ完整版:　　　　　　" + Num6_17 + " / " + Num7_17 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣老李　ＭｙＣａｒｄ完整版:　　　　" + Num6_18 + " / " + Num7_18 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｇｏｏｇｌｅ完整版:　　　　　　" + Num6_19 + " / " + Num7_19 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "松崗:　　　　　　　　　　　　　　　　" + Num6_20 + " / " + Num7_20 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "新加坡　Ｉｏｓ:　　　　　　　　　　　" + Num6_21 + " / " + Num7_21 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "新加坡　Ａｎｄｒｏｉｄ完整版:　　　　" + Num6_22 + " / " + Num7_22 + " / " + Num8_10 + "\r\n" + "\r\n";

                                    #endregion
                                }
                            }

                            else
                            {
                                textBox_Creat.Text = "資料庫無資料可分析";
                            }


                        }
                        catch
                        {
                            MessageBox.Show("資料查詢錯誤");

                        }

                        finally
                        {
                        }

                        break;


                    case 3:

                        label_Country.Text = "目前查詢國家:  中國";

                        try
                        {

                            //玩家等級分析
                            string Tsql = "SELECT Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 1) ORDER BY SaveDate DESC, CAST(V1 AS int), CAST(V2 AS int);";

                            //玩家名聲分析
                            Tsql += "SELECT  Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE  (Kind = 75) AND (Cause = 2) AND (V1 = 1 OR V1 = 2 OR V1 = 3 OR V1 = 4 OR V1 = 5) ORDER BY V1, SaveDate DESC, CAST(V2 AS int);";

                            //Vip等級
                            Tsql += "SELECT  Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE  (Kind = 75) AND (Cause = 3)  ORDER BY CAST(V1 AS int), SaveDate DESC,CAST(V2 AS int);";

                            //玩家現有鑽石數
                            Tsql += "SELECT Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 4) AND (V1 = 1 OR V1 = 2 OR V1 = 3 OR V1 = 4 OR V1 = 5) ORDER BY V1, SaveDate DESC, CAST(V2 AS int);";

                            //登入方式
                            Tsql += "SELECT Kind, Cause, V1, V2, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 5) ORDER BY SaveDate DESC,CAST(V2 AS int);";

                            //卡包分析
                            Tsql += "SELECT Cause, COUNT(UserID) AS Expr1 FROM Egg_" + PickTime + " WHERE (Kind = 78) AND (Cause = 1 OR Cause = 2 OR Cause = 3 OR Cause = 4 OR Cause = 6 OR Cause = 7 OR Cause = 9) GROUP BY Cause ORDER BY CAST(Cause AS int);";
                            Tsql += "SELECT V1, COUNT(UserID) AS Expr1 FROM Egg_" + PickTime + " WHERE (Kind = 78) AND (Cause = 5) OR (Kind = 78) AND (Cause = 8) OR (Kind = 78) AND (Cause = 10) GROUP BY V1 ORDER BY CAST(V1 AS int);";

                            //活動積分分析
                            Tsql += "SELECT UserID, MAX(DISTINCT CAST(V1 AS int)) AS Expr1 FROM Role_" + PickTime + " WHERE (Kind = 61) GROUP BY UserID HAVING (MAX(DISTINCT CAST(V1 AS int)) > " + SportPoint + ") ORDER BY MAX(DISTINCT CAST(V1 AS int)) DESC;";

                            //創角人物分析
                            Tsql += "SELECT Kind, Cause, V1, V3, SaveDate FROM Role_" + PickTime + " WHERE (Kind = 75) AND (Cause = 6 OR Cause = 7 OR Cause = 8) ORDER BY SaveDate , CAST(V1 AS int), Cause";

                            textBox_DisLv.Text = "";
                            textBox_DisDimon.Text = "";
                            textBox_DisLog.Text = "";
                            textBox_DisName.Text = "";
                            textBox_DisVip.Text = "";
                            textBox_CardBox.Text = "";
                            textBox_Point25W.Text = "";
                            textBox_Creat.Text = "";

                            SqlLink_CN.SQLLink CNsql = new SqlLink_CN.SQLLink();
                            DataSet ds = new DataSet();
                            ds = CNsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);

                            //玩家等級分析
                            int LvCount = 0;
                            bool LVSwitch = true;

                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20130913) //舊協定
                            {
                                for (int x = 0; x < ds.Tables[0].Rows.Count; x++)
                                {
                                    if ((x % 13) == 0)
                                    {
                                        textBox_DisLv.Text += "\r\n";
                                    }

                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V2"]))
                                    {
                                        case 1:
                                            Description = "3 級以下: ";
                                            break;

                                        case 2:
                                            Description = "5 級以下: ";
                                            break;

                                        case 3:
                                            Description = "10 級: ";
                                            break;

                                        case 4:
                                            Description = "15 級: ";
                                            break;

                                        case 5:
                                            Description = "20 級: ";
                                            break;

                                        case 6:
                                            Description = "30 級: ";
                                            break;

                                        case 7:
                                            Description = "40 級: ";
                                            break;

                                        case 8:
                                            Description = "50 級: ";
                                            break;

                                        case 9:
                                            Description = "60 級: ";
                                            break;

                                        case 10:
                                            Description = "70 級: ";
                                            break;

                                        case 11:
                                            Description = "80 級: ";
                                            break;

                                        case 12:
                                            Description = "90 級: ";
                                            break;

                                        case 13:
                                            Description = "超過90 級: ";
                                            break;
                                    }

                                    textBox_DisLv.Text += "Server (" + ds.Tables[0].Rows[x]["V1"].ToString() + ") : " + Description + "  " + ds.Tables[0].Rows[x]["V3"].ToString() + "\r\n";
                                }
                            }

                            else if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130912 && Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20131128) //新協定.v1
                            {
                                string RLv = "";

                                for (int x = 0; x < ds.Tables[0].Rows.Count; x++) //玩家等級分析
                                {
                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]))
                                    {
                                        case 1:
                                            RLv = "3 級以下: ";
                                            break;

                                        case 2:
                                            RLv = "5 級以下: ";
                                            break;

                                        case 3:
                                            RLv = "10 級: ";
                                            break;

                                        case 4:
                                            RLv = "15 級: ";
                                            break;

                                        case 5:
                                            RLv = "20 級: ";
                                            break;

                                        case 6:
                                            RLv = "30 級: ";
                                            break;

                                        case 7:
                                            RLv = "40 級: ";
                                            break;

                                        case 8:
                                            RLv = "50 級: ";
                                            break;

                                        case 9:
                                            RLv = "60 級: ";
                                            break;

                                        case 10:
                                            RLv = "70 級: ";
                                            break;

                                        case 11:
                                            RLv = "80 級: ";
                                            break;

                                        case 12:
                                            RLv = "90 級: ";
                                            break;

                                        case 13:
                                            RLv = "100 級: ";
                                            break;

                                        case 14:
                                            RLv = "115 級: ";
                                            break;

                                        case 15:
                                            RLv = "130 級: ";
                                            break;

                                        case 16:
                                            RLv = "150以下: ";
                                            break;

                                        case 17:
                                            RLv = "超過150 級: ";
                                            break;
                                    }

                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V2"]))
                                    {
                                        case 0:
                                            Description = "其它: ";
                                            break;

                                        case 1:
                                            Description = "一般 Apple: ";
                                            break;

                                        case 2:
                                            Description = "完整 Apple: ";
                                            break;

                                        case 3:
                                            Description = "一般 Goole: ";
                                            break;

                                        case 4:
                                            Description = "完整 Goole: ";
                                            break;

                                        case 5:
                                            Description = "一般 MyCard: ";
                                            break;

                                        case 6:
                                            Description = "完整 MyCard: ";
                                            break;

                                        case 7:
                                            Description = "一般台灣老李 MyCard: ";
                                            break;

                                        case 8:
                                            Description = "完整台灣老李 MyCard: ";
                                            break;

                                        case 9:
                                            Description = "一般 Android: ";
                                            break;

                                        case 10:
                                            Description = "完整 Android: ";
                                            break;

                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) - LvCount == 1)
                                    {
                                        textBox_DisLv.Text += "\r\n" + "\r\n" + " 玩家等級: " + RLv + "\r\n" + "\r\n";
                                        LvCount++;
                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) >= LvCount && LVSwitch == true)
                                    {
                                        textBox_DisLv.Text += Description + "  " + ds.Tables[0].Rows[x]["V3"].ToString() + "\r\n";
                                    }
                                    else
                                    {
                                        LVSwitch = false;
                                    }

                                }
                            }

                            else  //新協定.v2
                            {
                                string RLv = "";
                                for (int x = 0; x < ds.Tables[0].Rows.Count; x++) //玩家等級分析
                                {
                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]))
                                    {
                                        case 1:
                                            RLv = "3 級以下: ";
                                            break;

                                        case 2:
                                            RLv = "5 級以下: ";
                                            break;

                                        case 3:
                                            RLv = "10 級: ";
                                            break;

                                        case 4:
                                            RLv = "15 級: ";
                                            break;

                                        case 5:
                                            RLv = "20 級: ";
                                            break;

                                        case 6:
                                            RLv = "30 級: ";
                                            break;

                                        case 7:
                                            RLv = "40 級: ";
                                            break;

                                        case 8:
                                            RLv = "50 級: ";
                                            break;

                                        case 9:
                                            RLv = "60 級: ";
                                            break;

                                        case 10:
                                            RLv = "70 級: ";
                                            break;

                                        case 11:
                                            RLv = "80 級: ";
                                            break;

                                        case 12:
                                            RLv = "90 級: ";
                                            break;

                                        case 13:
                                            RLv = "100 級: ";
                                            break;

                                        case 14:
                                            RLv = "115 級: ";
                                            break;

                                        case 15:
                                            RLv = "130 級: ";
                                            break;

                                        case 16:
                                            RLv = "150以下: ";
                                            break;

                                        case 17:
                                            RLv = "超過150 級: ";
                                            break;
                                    }

                                    switch (Convert.ToInt32(ds.Tables[0].Rows[x]["V2"]))
                                    {
                                        case 0:
                                            Description = "Ios 其他: ";
                                            break;

                                        case 1:
                                            Description = "香港 IOS: ";
                                            break;

                                        case 2:
                                            Description = "香港 IOS 完整版: ";
                                            break;

                                        case 3:
                                            Description = "台灣 IOS 完整版: ";
                                            break;

                                        case 4:
                                            Description = "台灣 IOS: ";
                                            break;

                                        case 5:
                                            Description = "大陸 IOS 完整版: ";
                                            break;

                                        case 6:
                                            Description = "Android 其他: ";
                                            break;

                                        case 7:
                                            Description = "香港 MyCard: ";
                                            break;

                                        case 8:
                                            Description = "香港 MyCard 完整版: ";
                                            break;

                                        case 9:
                                            Description = "香港 Google: ";
                                            break;

                                        case 10:
                                            Description = "台灣 MyCard 完整版: ";
                                            break;

                                        case 11:
                                            Description = "台灣 Google: ";
                                            break;

                                        case 12:
                                            Description = "大陸 Android 完整版 (未): ";
                                            break;

                                        case 13:
                                            Description = "IOS Debug: ";
                                            break;

                                        case 14:
                                            Description = "Android Debug: ";
                                            break;

                                        case 15:
                                            Description = "台灣老李 MyCard: ";
                                            break;

                                        case 16:
                                            Description = "台灣 MyCard: ";
                                            break;

                                        case 17:
                                            Description = "台灣 Google 完整版: ";
                                            break;

                                        case 18:
                                            Description = "台灣老李 MyCard 完整版: ";
                                            break;

                                        case 19:
                                            Description = "香港 Google 完整版: ";
                                            break;

                                        case 20:
                                            Description = "松崗: ";
                                            break;

                                        case 21:
                                            Description = "新加坡 IOS: ";
                                            break;

                                        case 22:
                                            Description = "新加坡 Android 完整版: ";
                                            break;

                                        case 23:
                                            Description = "新加坡 Android 當地金流: ";
                                            break;

                                        case 24:
                                            Description = "大陸91 IOS: ";
                                            break;

                                        case 25:
                                            Description = "大陸UC Android: ";
                                            break;

                                        case 26:
                                            Description = "大陸360 Android: ";
                                            break;

                                        case 27:
                                            Description = "大陸App助手 IOS: ";
                                            break;

                                        case 28:
                                            Description = "大陸官方apk Android: ";
                                            break;

                                        case 29:
                                            Description = "大陸91 Android: ";
                                            break;

                                        case 30:
                                            Description = "馬幹線 IOS: ";
                                            break;

                                        case 31:
                                            Description = "馬幹線 Android: ";
                                            break;

                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) - LvCount == 1)
                                    {
                                        textBox_DisLv.Text += "\r\n" + "\r\n" + " 玩家等級: " + RLv + "\r\n" + "\r\n";
                                        LvCount++;
                                    }

                                    if (Convert.ToInt32(ds.Tables[0].Rows[x]["V1"]) >= LvCount && LVSwitch == true)
                                    {
                                        textBox_DisLv.Text += Description + "  " + ds.Tables[0].Rows[x]["V3"].ToString() + "\r\n";
                                    }
                                    else
                                    {
                                        LVSwitch = false;
                                    }

                                }
                            }


                            //玩家名聲分析

                            int y = 0;

                            for (int x = 0; x < ds.Tables[1].Rows.Count; x++)
                            {

                                if (Convert.ToInt32(ds.Tables[1].Rows[x]["V1"]) - y == 1)
                                {
                                    textBox_DisName.Text += "\r\n" + "伺服器: " + ds.Tables[1].Rows[x]["V1"].ToString() + "\r\n" + "\r\n";
                                    y++;
                                }

                                switch (Convert.ToInt32(ds.Tables[1].Rows[x]["V2"]))
                                {

                                    case 1:
                                        Description = "1000 以下: ";
                                        break;

                                    case 2:
                                        Description = "2000: ";
                                        break;

                                    case 3:
                                        Description = "5000: ";
                                        break;

                                    case 4:
                                        Description = "10000: ";
                                        break;

                                    case 5:
                                        Description = "20000: ";
                                        break;

                                    case 6:
                                        Description = "50000: ";
                                        break;

                                    case 7:
                                        Description = "100000: ";
                                        break;

                                    case 8:
                                        Description = "200000: ";
                                        break;

                                    case 9:
                                        Description = "250000: ";
                                        break;

                                    case 10:
                                        Description = "300000: ";
                                        break;

                                    case 11:
                                        Description = "350000: ";
                                        break;

                                    case 12:
                                        Description = "400000: ";
                                        break;

                                    case 13:
                                        Description = "450000: ";
                                        break;

                                    case 14:
                                        Description = "500000: ";
                                        break;

                                    case 15:
                                        Description = "超過 500000: ";
                                        break;

                                    case 16:
                                        Description = "1000000: ";
                                        break;

                                    case 17:
                                        Description = "1500000: ";
                                        break;
                                }

                                textBox_DisName.Text += Description + "  " + ds.Tables[1].Rows[x]["V3"].ToString() + "\r\n";

                            }




                            //Vip分析

                            int VipCount = -1;

                            for (int x = 0; x < ds.Tables[2].Rows.Count; x++)
                            {
                                if (Convert.ToInt32(ds.Tables[2].Rows[x]["V1"]) - VipCount == 1)
                                {
                                    textBox_DisVip.Text += "\r\n" + "Vip等級: " + ds.Tables[2].Rows[x]["V1"].ToString() + "\r\n" + "\r\n";
                                    VipCount++;
                                }

                                if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20130912) //舊協定
                                {
                                    switch (Convert.ToInt32(ds.Tables[2].Rows[x]["V2"].ToString()))
                                    {

                                        case 0:
                                            Description = "其他: ";
                                            break;

                                        case 1:
                                            Description = "Apple: ";
                                            break;

                                        case 2:
                                            Description = "一般 My Card: ";
                                            break;

                                        case 3:
                                            Description = "完整 My Card: ";
                                            break;

                                        case 4:
                                            Description = "Google: ";
                                            break;

                                        case 5:
                                            Description = "東方阿李: ";
                                            break;
                                    }
                                }

                                else if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130911 && Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20131128) //舊協定
                                {
                                    switch (Convert.ToInt32(ds.Tables[2].Rows[x]["V2"].ToString()))
                                    {

                                        case 0:
                                            Description = "其他: ";
                                            break;

                                        case 1:
                                            Description = "一般 Apple: ";
                                            break;

                                        case 2:
                                            Description = "完整 Apple: ";
                                            break;

                                        case 3:
                                            Description = "一般 Google: ";
                                            break;

                                        case 4:
                                            Description = "完整 Google: ";
                                            break;

                                        case 5:
                                            Description = "一般 MyCard: ";
                                            break;

                                        case 6:
                                            Description = "完整 MyCard: ";
                                            break;

                                        case 7:
                                            Description = "一般台灣老李MyCard: ";
                                            break;

                                        case 8:
                                            Description = "完整台灣老李 MyCard: ";
                                            break;

                                        case 9:
                                            Description = "一般 Android: ";
                                            break;

                                        case 10:
                                            Description = "完整 Android: ";
                                            break;
                                    }
                                }

                                else
                                {
                                    switch (Convert.ToInt32(ds.Tables[2].Rows[x]["V2"].ToString()))
                                    {

                                        case 0:
                                            Description = "Ios 其他: ";
                                            break;

                                        case 1:
                                            Description = "香港 IOS: ";
                                            break;

                                        case 2:
                                            Description = "香港 IOS 完整版: ";
                                            break;

                                        case 3:
                                            Description = "台灣 IOS 完整版: ";
                                            break;

                                        case 4:
                                            Description = "台灣 IOS: ";
                                            break;

                                        case 5:
                                            Description = "大陸 IOS 完整版: ";
                                            break;

                                        case 6:
                                            Description = "Android 其他: ";
                                            break;

                                        case 7:
                                            Description = "香港 MyCard: ";
                                            break;

                                        case 8:
                                            Description = "香港 MyCard 完整版: ";
                                            break;

                                        case 9:
                                            Description = "香港 Google: ";
                                            break;

                                        case 10:
                                            Description = "台灣 MyCard 完整版: ";
                                            break;

                                        case 11:
                                            Description = "台灣 Google: ";
                                            break;

                                        case 12:
                                            Description = "大陸 Android 完整版 (未): ";
                                            break;

                                        case 13:
                                            Description = "IOS Debug: ";
                                            break;

                                        case 14:
                                            Description = "Android Debug: ";
                                            break;

                                        case 15:
                                            Description = "台灣老李 MyCard: ";
                                            break;

                                        case 16:
                                            Description = "台灣 MyCard: ";
                                            break;

                                        case 17:
                                            Description = "台灣 Google 完整版: ";
                                            break;

                                        case 18:
                                            Description = "台灣老李 MyCard 完整版: ";
                                            break;

                                        case 19:
                                            Description = "香港 Google 完整版: ";
                                            break;

                                        case 20:
                                            Description = "松崗: ";
                                            break;

                                        case 21:
                                            Description = "新加坡 IOS: ";
                                            break;

                                        case 22:
                                            Description = "新加坡 Android 完整版: ";
                                            break;

                                        case 23:
                                            Description = "新加坡 Android 當地金流: ";
                                            break;

                                        case 24:
                                            Description = "大陸91 IOS: ";
                                            break;

                                        case 25:
                                            Description = "大陸UC Android: ";
                                            break;

                                        case 26:
                                            Description = "大陸360 Android: ";
                                            break;

                                        case 27:
                                            Description = "大陸App助手 IOS: ";
                                            break;

                                        case 28:
                                            Description = "大陸官方apk Android: ";
                                            break;

                                        case 29:
                                            Description = "大陸91 Android: ";
                                            break;

                                        case 30:
                                            Description = "馬幹線 IOS: ";
                                            break;

                                        case 31:
                                            Description = "馬幹線 Android: ";
                                            break;
                                    }
                                }

                                textBox_DisVip.Text += Description + "  " + ds.Tables[2].Rows[x]["V3"].ToString() + "\r\n";
                            }



                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130905)
                            {
                                //玩家現有鑽石分析

                                int DimonCount = 0;

                                for (int x = 0; x < ds.Tables[3].Rows.Count; x++)
                                {
                                    if (Convert.ToInt32(ds.Tables[3].Rows[x]["V1"]) - DimonCount == 1)
                                    {
                                        textBox_DisDimon.Text += "\r\n" + "伺服器: " + ds.Tables[3].Rows[x]["V1"].ToString() + "\r\n" + "\r\n";
                                        DimonCount++;
                                    }

                                    switch (Convert.ToInt32(ds.Tables[3].Rows[x]["V2"]))
                                    {
                                        case 0:
                                            Description = "60 以下: ";
                                            break;

                                        case 1:
                                            Description = "500: ";
                                            break;

                                        case 2:
                                            Description = "1000: ";
                                            break;

                                        case 3:
                                            Description = "2000: ";
                                            break;

                                        case 4:
                                            Description = "5000: ";
                                            break;

                                        case 5:
                                            Description = "10000: ";
                                            break;

                                        case 6:
                                            Description = "20000: ";
                                            break;

                                        case 7:
                                            Description = "50000: ";
                                            break;

                                        case 8:
                                            Description = "100000: ";
                                            break;

                                        case 9:
                                            Description = "10萬到50萬: ";
                                            break;

                                        case 10:
                                            Description = "500000以上: ";
                                            break;
                                    }

                                    textBox_DisDimon.Text += Description + "  " + ds.Tables[3].Rows[x]["V3"].ToString() + "\r\n";
                                }

                            }
                            else
                            {
                                textBox_DisDimon.Text = "資料庫無資料可分析";
                            }


                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130904)
                            {
                                for (int j = 0; j < 1; j++) //登入方式
                                {
                                    textBox_DisLog.Text += "\r\n" + "登入方式: " + "\r\n" + "\r\n";

                                    for (int x = 0; x < ds.Tables[j + 4].Rows.Count; x++)
                                    {
                                        switch (Convert.ToInt32(ds.Tables[j + 4].Rows[x]["V2"]))
                                        {

                                            case 0:
                                                Description = "其它: ";
                                                break;

                                            case 1:
                                                Description = "香港一般版: ";
                                                break;

                                            case 2:
                                                Description = "香港 IOS 完整版: ";
                                                break;

                                            case 3:
                                                Description = "台灣 IOS 完整版: ";
                                                break;

                                            case 4:
                                                Description = "台灣一般版: ";
                                                break;

                                            case 5:
                                                Description = "阿六版: ";
                                                break;

                                            case 6:
                                                Description = "未登錄: ";
                                                break;

                                            case 7:
                                                Description = "香港一般MYCARD: ";
                                                break;

                                            case 8:
                                                Description = "香港完整MYCARD: ";
                                                break;

                                            case 9:
                                                Description = "香港一般Google: ";
                                                break;

                                            case 10:
                                                Description = "台灣完整MYCARD: ";
                                                break;

                                            case 11:
                                                Description = "台灣一般google: ";
                                                break;

                                            case 12:
                                                Description = "大陸 Android 完整版 (未): ";
                                                break;

                                            case 13:
                                                Description = "IOS Debug: ";
                                                break;

                                            case 14:
                                                Description = "Android Debug: ";
                                                break;

                                            case 15:
                                                Description = "東方阿李: ";
                                                break;

                                            case 16:
                                                Description = "台灣 MyCard: ";
                                                break;

                                            case 17:
                                                Description = "台灣 Google 完整版: ";
                                                break;

                                            case 18:
                                                Description = "台灣老李 MyCard 完整版: ";
                                                break;

                                            case 19:
                                                Description = "香港 Google 完整版: ";
                                                break;

                                            case 20:
                                                Description = "松崗: ";
                                                break;

                                            case 21:
                                                Description = "新加坡 IOS: ";
                                                break;

                                            case 22:
                                                Description = "新加坡 Android 完整版: ";
                                                break;

                                            case 23:
                                                Description = "新加坡 Android 當地金流: ";
                                                break;

                                            case 24:
                                                Description = "大陸91 IOS: ";
                                                break;

                                            case 25:
                                                Description = "大陸UC Android: ";
                                                break;

                                            case 26:
                                                Description = "大陸360 Android: ";
                                                break;

                                            case 27:
                                                Description = "大陸App助手 IOS: ";
                                                break;

                                            case 28:
                                                Description = "大陸官方apk Android: ";
                                                break;

                                            case 29:
                                                Description = "大陸91 Android: ";
                                                break;

                                            case 30:
                                                Description = "馬幹線 IOS: ";
                                                break;

                                            case 31:
                                                Description = "馬幹線 Android: ";
                                                break;
                                        }

                                        textBox_DisLog.Text += Description + "  " + ds.Tables[j + 4].Rows[x]["V3"].ToString() + "\r\n" + "\r\n";
                                    }
                                }
                            }
                            else
                            {
                                textBox_DisLog.Text = "資料庫無資料可分析";
                            }

                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130905)
                            {
                                //卡包分析

                                textBox_CardBox.Text += "\r\n";

                                for (int x = 0; x < ds.Tables[5].Rows.Count; x++)
                                {
                                    switch (Convert.ToInt32(ds.Tables[5].Rows[x]["Cause"]))
                                    {
                                        case 0:
                                            Description = "聖旨單抽: ";
                                            break;

                                        case 1:
                                            Description = "友情單抽: ";
                                            break;

                                        case 2:
                                            Description = "銀幣單抽: ";
                                            break;

                                        case 3:
                                            Description = "鑽石單抽: ";
                                            break;

                                        case 6:
                                            Description = "友情10連抽: ";
                                            break;

                                        case 7:
                                            Description = "銀幣10連抽: ";
                                            break;

                                        case 9:
                                            Description = "９鑽抽:     ";
                                            break;
                                    }

                                    textBox_CardBox.Text += Description + "  " + ds.Tables[5].Rows[x]["Expr1"].ToString() + "\r\n" + "\r\n";
                                }


                                textBox_CardBox.Text += "\r\n";

                                for (int x = 0; x < ds.Tables[6].Rows.Count; x++) //卡包分析
                                {
                                    switch (Convert.ToInt32(ds.Tables[6].Rows[x]["V1"]))
                                    {
                                        case 4:
                                            Description = "必中卡包: ";
                                            break;

                                        case 7:
                                            Description = "一般卡包: ";
                                            break;

                                        case 11:
                                            Description = "卡包1(吳國太): ";
                                            break;

                                        case 12:
                                            Description = "卡包2(甘夫人): ";
                                            break;

                                        case 13:
                                            Description = "卡包3(蔡文姬): ";
                                            break;

                                        case 14:
                                            Description = "卡包4(夏日美女包): ";
                                            break;

                                        case 15:
                                            Description = "卡包5(闇屬趙雲): ";
                                            break;

                                        case 16:
                                            Description = "卡包6(光屬張飛): ";
                                            break;

                                        case 17:
                                            Description = "卡包7(蜀國勢力包:龐統必中): ";
                                            break;

                                        case 18:
                                            Description = "卡包8(吳國勢力包:龐統必中): ";
                                            break;

                                        case 19:
                                            Description = "卡包9(魏國勢力包:司馬懿): ";
                                            break;

                                        case 20:
                                            Description = "卡包10(他國勢力包:水呂布): ";
                                            break;

                                        case 21:
                                            Description = "卡包11: ";
                                            break;

                                        case 22:
                                            Description = "卡包12: ";
                                            break;

                                        case 23:
                                            Description = "卡包13: ";
                                            break;

                                        case 24:
                                            Description = "卡包14: ";
                                            break;

                                        case 25:
                                            Description = "卡包15: ";
                                            break;

                                        case 26:
                                            Description = "卡包16: ";
                                            break;

                                        case 27:
                                            Description = "卡包17: ";
                                            break;

                                        case 28:
                                            Description = "卡包18: ";
                                            break;

                                        case 29:
                                            Description = "卡包19: ";
                                            break;

                                        case 30:
                                            Description = "卡包20: ";
                                            break;


                                    }

                                    textBox_CardBox.Text += Description + "  " + ds.Tables[6].Rows[x]["Expr1"].ToString() + "\r\n" + "\r\n";
                                }
                            }
                            else
                            {
                                textBox_CardBox.Text = "資料庫無資料可分析";
                            }

                            //活動積分分析
                            textBox_Point25W.Text += "\r\n";
                            int ID = 0;
                            for (int x = 0; x < ds.Tables[7].Rows.Count; x++)
                            {
                                ID++;
                                textBox_Point25W.Text += " (" + ID + "). " + "玩家ID: " + ds.Tables[7].Rows[x]["UserID"].ToString() + "     積分: " + ds.Tables[7].Rows[x]["Expr1"].ToString() + "\r\n" + "\r\n";
                            }

                            if (textBox_Point25W.Text == "\r\n")
                            {
                                textBox_Point25W.Text = "無對應此積分資料";
                            }

                            if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) > 20130917)
                            {
                                if (Convert.ToInt32(dateTimePicker_Date.Value.ToString("yyyyMMdd")) < 20131129)
                                {
                                    #region CreatCharacter

                                    //當日創角
                                    textBox_Creat.Text += "當日創角人數 /" + "\r\n" + "當日創角且超過5等 /" + "\r\n" + "前天創角且連續兩天有上線" + "\r\n" + "\r\n";

                                    #region 數字空間
                                    string Num6_0 = "";
                                    string Num6_1 = "";
                                    string Num6_2 = "";
                                    string Num6_3 = "";
                                    string Num6_4 = "";
                                    string Num6_5 = "";
                                    string Num6_6 = "";
                                    string Num6_7 = "";
                                    string Num6_8 = "";
                                    string Num6_9 = "";
                                    string Num6_10 = "";

                                    string Num7_0 = "";
                                    string Num7_1 = "";
                                    string Num7_2 = "";
                                    string Num7_3 = "";
                                    string Num7_4 = "";
                                    string Num7_5 = "";
                                    string Num7_6 = "";
                                    string Num7_7 = "";
                                    string Num7_8 = "";
                                    string Num7_9 = "";
                                    string Num7_10 = "";

                                    string Num8_0 = "";
                                    string Num8_1 = "";
                                    string Num8_2 = "";
                                    string Num8_3 = "";
                                    string Num8_4 = "";
                                    string Num8_5 = "";
                                    string Num8_6 = "";
                                    string Num8_7 = "";
                                    string Num8_8 = "";
                                    string Num8_9 = "";
                                    string Num8_10 = "";
                                    #endregion

                                    for (int x = 0; x < ds.Tables[8].Rows.Count; x++)
                                    {
                                        switch (Convert.ToInt32(ds.Tables[8].Rows[x]["Cause"].ToString()))
                                        {
                                            case 6:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"其他:　　　　　　　　　　　";
                                                        Num6_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"一般Ａｐｐｌｅ:　　　　　　";
                                                        Num6_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"完整Ａｐｐｌｅ:　　　　　　";
                                                        Num6_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"一般Ｇｏｏｇｌｅ:　　　　　";
                                                        Num6_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"完整Ｇｏｏｇｌｅ:　　　　　";
                                                        Num6_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"一般ＭｙＣａｒｄ:　　　　　";
                                                        Num6_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"完整ＭｙＣａｒｄ:　　　　　";
                                                        Num6_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"一般台灣老李ＭｙＣａｒｄ:　";
                                                        Num6_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"完整台灣老李ＭｙＣａｒｄ:　";
                                                        Num6_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"一般Ａｎｄｒｏｉｄ:　　　　";
                                                        Num6_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"完整Ａｎｄｒｏｉｄ:　　　　";
                                                        Num6_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                            case 7:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"其他:　　　　　　　　　　　";
                                                        Num7_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"一般Ａｐｐｌｅ:　　　　　　";
                                                        Num7_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"完整Ａｐｐｌｅ:　　　　　　";
                                                        Num7_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"一般Ｇｏｏｇｌｅ:　　　　　";
                                                        Num7_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"完整Ｇｏｏｇｌｅ:　　　　　";
                                                        Num7_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"一般ＭｙＣａｒｄ:　　　　　";
                                                        Num7_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"完整ＭｙＣａｒｄ:　　　　　";
                                                        Num7_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"一般台灣老李ＭｙＣａｒｄ:　";
                                                        Num7_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"完整台灣老李ＭｙＣａｒｄ:　";
                                                        Num7_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"一般Ａｎｄｒｏｉｄ:　　　　";
                                                        Num7_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"完整Ａｎｄｒｏｉｄ:　　　　";
                                                        Num7_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }


                                                break;

                                            case 8:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"其他:　　　　　　　　　　　";
                                                        Num8_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"一般Ａｐｐｌｅ:　　　　　　";
                                                        Num8_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"完整Ａｐｐｌｅ:　　　　　　";
                                                        Num8_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"一般Ｇｏｏｇｌｅ:　　　　　";
                                                        Num8_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"完整Ｇｏｏｇｌｅ:　　　　　";
                                                        Num8_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"一般ＭｙＣａｒｄ:　　　　　";
                                                        Num8_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"完整ＭｙＣａｒｄ:　　　　　";
                                                        Num8_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"一般台灣老李ＭｙＣａｒｄ:　";
                                                        Num8_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"完整台灣老李ＭｙＣａｒｄ:　";
                                                        Num8_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"一般Ａｎｄｒｏｉｄ:　　　　";
                                                        Num8_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"完整Ａｎｄｒｏｉｄ:　　　　";
                                                        Num8_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                        }

                                    }

                                    textBox_Creat.Text += "其他:　　　　　　　　　　　" + Num6_0 + " / " + Num7_0 + " / " + Num8_0 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般Ａｐｐｌｅ:　　　　　　" + Num6_1 + " / " + Num7_1 + " / " + Num8_1 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整Ａｐｐｌｅ:　　　　　　" + Num6_2 + " / " + Num7_2 + " / " + Num8_2 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般Ｇｏｏｇｌｅ:　　　　　" + Num6_3 + " / " + Num7_3 + " / " + Num8_3 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整Ｇｏｏｇｌｅ:　　　　　" + Num6_4 + " / " + Num7_4 + " / " + Num8_4 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般ＭｙＣａｒｄ:　　　　　" + Num6_5 + " / " + Num7_5 + " / " + Num8_5 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整ＭｙＣａｒｄ:　　　　　" + Num6_6 + " / " + Num7_6 + " / " + Num8_6 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般台灣老李ＭｙＣａｒｄ:　" + Num6_7 + " / " + Num7_7 + " / " + Num8_7 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整台灣老李ＭｙＣａｒｄ:　" + Num6_8 + " / " + Num7_8 + " / " + Num8_8 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "一般Ａｎｄｒｏｉｄ:　　　　" + Num6_9 + " / " + Num7_9 + " / " + Num8_9 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "完整Ａｎｄｒｏｉｄ:　　　　" + Num6_10 + " / " + Num7_10 + " / " + Num8_10 + "\r\n" + "\r\n";

                                    #endregion
                                }

                                else
                                {
                                    #region CreatCharacter

                                    //當日創角
                                    textBox_Creat.Text += "當日創角人數 /" + "\r\n" + "當日創角且超過5等 /" + "\r\n" + "前天創角且連續兩天有上線" + "\r\n" + "\r\n";

                                    #region 數字空間
                                    string Num6_0 = "";
                                    string Num6_1 = "";
                                    string Num6_2 = "";
                                    string Num6_3 = "";
                                    string Num6_4 = "";
                                    string Num6_5 = "";
                                    string Num6_6 = "";
                                    string Num6_7 = "";
                                    string Num6_8 = "";
                                    string Num6_9 = "";
                                    string Num6_10 = "";
                                    string Num6_11 = "";
                                    string Num6_12 = "";
                                    string Num6_13 = "";
                                    string Num6_14 = "";
                                    string Num6_15 = "";
                                    string Num6_16 = "";
                                    string Num6_17 = "";
                                    string Num6_18 = "";
                                    string Num6_19 = "";
                                    string Num6_20 = "";
                                    string Num6_21 = "";
                                    string Num6_22 = "";

                                    string Num7_0 = "";
                                    string Num7_1 = "";
                                    string Num7_2 = "";
                                    string Num7_3 = "";
                                    string Num7_4 = "";
                                    string Num7_5 = "";
                                    string Num7_6 = "";
                                    string Num7_7 = "";
                                    string Num7_8 = "";
                                    string Num7_9 = "";
                                    string Num7_10 = "";
                                    string Num7_11 = "";
                                    string Num7_12 = "";
                                    string Num7_13 = "";
                                    string Num7_14 = "";
                                    string Num7_15 = "";
                                    string Num7_16 = "";
                                    string Num7_17 = "";
                                    string Num7_18 = "";
                                    string Num7_19 = "";
                                    string Num7_20 = "";
                                    string Num7_21 = "";
                                    string Num7_22 = "";

                                    string Num8_0 = "";
                                    string Num8_1 = "";
                                    string Num8_2 = "";
                                    string Num8_3 = "";
                                    string Num8_4 = "";
                                    string Num8_5 = "";
                                    string Num8_6 = "";
                                    string Num8_7 = "";
                                    string Num8_8 = "";
                                    string Num8_9 = "";
                                    string Num8_10 = "";
                                    string Num8_11 = "";
                                    string Num8_12 = "";
                                    string Num8_13 = "";
                                    string Num8_14 = "";
                                    string Num8_15 = "";
                                    string Num8_16 = "";
                                    string Num8_17 = "";
                                    string Num8_18 = "";
                                    string Num8_19 = "";
                                    string Num8_20 = "";
                                    string Num8_21 = "";
                                    string Num8_22 = "";

                                    #endregion

                                    for (int x = 0; x < ds.Tables[8].Rows.Count; x++)
                                    {
                                        switch (Convert.ToInt32(ds.Tables[8].Rows[x]["Cause"].ToString()))
                                        {
                                            case 6:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"Ｉｏｓ其他:　　　　　　　　　　　　　";
                                                        Num6_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"香港　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num6_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"香港　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num6_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"台灣　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num6_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"台灣　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num6_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"大陸　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num6_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"Ａｎｄｒｏｉｄ其他:　　　　　　　　　";
                                                        Num6_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"香港　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num6_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"香港　ＭｙＣａｒｄ完整版:　　　　　　";
                                                        Num6_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"香港　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num6_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"台灣　Ｍｙｃａｒｄ完整版:　　　　　　";
                                                        Num6_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 11:
                                                        //"台灣　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num6_11 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 12:
                                                        //"大陸　Ａｎｄｒｏｉｄ 完整版 (未):　　";
                                                        Num6_12 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 13:
                                                        //"ＩｏｓＤｅｂｕｇ:　　　　　　　　　　";
                                                        Num6_13 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 14:
                                                        //"ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　";
                                                        Num6_14 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 15:
                                                        //"台灣老李　ＭｙＣａｒｄ:　　　　　　　";
                                                        Num6_15 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 16:
                                                        //"台灣　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num6_16 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 17:
                                                        //"台灣　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num6_17 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 18:
                                                        //"台灣老李　ＭｙＣａｒｄ完整版:　　　　";
                                                        Num6_18 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 19:
                                                        //"香港　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num6_19 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 20:
                                                        //"松崗:　　　　　　　　　　　　　　　　";
                                                        Num6_20 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 21:
                                                        //"新加坡　Ｉｏｓ:　　　　　　　　　　　";
                                                        Num6_21 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 22:
                                                        //"新加坡　Ａｎｄｒｏｉｄ完整版:　　　　";
                                                        Num6_22 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                            case 7:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"Ｉｏｓ其他:　　　　　　　　　　　　　";
                                                        Num7_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"香港　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num7_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"香港　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num7_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"台灣　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num7_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"台灣　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num7_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"大陸　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num7_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"Ａｎｄｒｏｉｄ其他:　　　　　　　　　";
                                                        Num7_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"香港　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num7_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"香港　ＭｙＣａｒｄ完整版:　　　　　　";
                                                        Num7_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"香港　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num7_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"台灣　Ｍｙｃａｒｄ完整版:　　　　　　";
                                                        Num7_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 11:
                                                        //"台灣　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num7_11 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 12:
                                                        //"大陸　Ａｎｄｒｏｉｄ 完整版 (未):　　";
                                                        Num7_12 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 13:
                                                        //"ＩｏｓＤｅｂｕｇ:　　　　　　　　　　";
                                                        Num7_13 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 14:
                                                        //"ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　";
                                                        Num7_14 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 15:
                                                        //"台灣老李　ＭｙＣａｒｄ:　　　　　　　";
                                                        Num7_15 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 16:
                                                        //"台灣　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num7_16 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 17:
                                                        //"台灣　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num7_17 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 18:
                                                        //"台灣老李　ＭｙＣａｒｄ完整版:　　　　";
                                                        Num7_18 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 19:
                                                        //"香港　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num7_19 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 20:
                                                        //"松崗:　　　　　　　　　　　　　　　　";
                                                        Num7_20 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 21:
                                                        //"新加坡　Ｉｏｓ:　　　　　　　　　　　";
                                                        Num7_21 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 22:
                                                        //"新加坡　Ａｎｄｒｏｉｄ完整版:　　　　";
                                                        Num7_22 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }


                                                break;

                                            case 8:

                                                switch (Convert.ToInt32(ds.Tables[8].Rows[x]["V1"].ToString()))
                                                {
                                                    case 0:
                                                        //"Ｉｏｓ其他:　　　　　　　　　　　　　";
                                                        Num8_0 = ds.Tables[8].Rows[x]["V3"].ToString();

                                                        break;

                                                    case 1:
                                                        //"香港　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num8_1 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 2:
                                                        //"香港　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num8_2 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 3:
                                                        //"台灣　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num8_3 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 4:
                                                        //"台灣　Ｉｏｓ:　　　　　　　　　　　　";
                                                        Num8_4 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 5:
                                                        //"大陸　Ｉｏｓ完整版:　　　　　　　　　";
                                                        Num8_5 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 6:
                                                        //"Ａｎｄｒｏｉｄ其他:　　　　　　　　　";
                                                        Num8_6 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 7:
                                                        //"香港　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num8_7 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 8:
                                                        //"香港　ＭｙＣａｒｄ完整版:　　　　　　";
                                                        Num8_8 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 9:
                                                        //"香港　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num8_9 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 10:
                                                        //"台灣　Ｍｙｃａｒｄ完整版:　　　　　　";
                                                        Num8_10 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 11:
                                                        //"台灣　Ｇｏｏｇｌｅ:　　　　　　　　　";
                                                        Num8_11 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 12:
                                                        //"大陸　Ａｎｄｒｏｉｄ 完整版 (未):　　";
                                                        Num8_12 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 13:
                                                        //"ＩｏｓＤｅｂｕｇ:　　　　　　　　　　";
                                                        Num8_13 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 14:
                                                        //"ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　";
                                                        Num8_14 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 15:
                                                        //"台灣老李　ＭｙＣａｒｄ:　　　　　　　";
                                                        Num8_15 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 16:
                                                        //"台灣　ＭｙＣａｒｄ:　　　　　　　　　";
                                                        Num8_16 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 17:
                                                        //"台灣　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num8_17 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 18:
                                                        //"台灣老李　ＭｙＣａｒｄ完整版:　　　　";
                                                        Num8_18 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 19:
                                                        //"香港　Ｇｏｏｇｌｅ完整版:　　　　　　";
                                                        Num8_19 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 20:
                                                        //"松崗:　　　　　　　　　　　　　　　　";
                                                        Num8_20 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 21:
                                                        //"新加坡　Ｉｏｓ:　　　　　　　　　　　";
                                                        Num8_21 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                    case 22:
                                                        //"新加坡　Ａｎｄｒｏｉｄ完整版:　　　　";
                                                        Num8_22 = ds.Tables[8].Rows[x]["V3"].ToString();
                                                        break;

                                                }

                                                break;

                                        }

                                    }

                                    textBox_Creat.Text += "Ｉｏｓ其他:　　　　　　　　　　　　　" + Num6_0 + " / " + Num7_0 + " / " + Num8_0 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｉｏｓ:　　　　　　　　　　　　" + Num6_1 + " / " + Num7_1 + " / " + Num8_1 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｉｏｓ完整版:　　　　　　　　　" + Num6_2 + " / " + Num7_2 + " / " + Num8_2 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｉｏｓ完整版:　　　　　　　　　" + Num6_3 + " / " + Num7_3 + " / " + Num8_3 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｉｏｓ:　　　　　　　　　　　　" + Num6_4 + " / " + Num7_4 + " / " + Num8_4 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "大陸　Ｉｏｓ完整版:　　　　　　　　　" + Num6_5 + " / " + Num7_5 + " / " + Num8_5 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "Ａｎｄｒｏｉｄ其他:　　　　　　　　　" + Num6_6 + " / " + Num7_6 + " / " + Num8_6 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　ＭｙＣａｒｄ:　　　　　　　　　" + Num6_7 + " / " + Num7_7 + " / " + Num8_7 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　ＭｙＣａｒｄ完整版:　　　　　　" + Num6_8 + " / " + Num7_8 + " / " + Num8_8 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｇｏｏｇｌｅ:　　　　　　　　　" + Num6_9 + " / " + Num7_9 + " / " + Num8_9 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｍｙｃａｒｄ完整版:　　　　　　" + Num6_10 + " / " + Num7_10 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｇｏｏｇｌｅ:　　　　　　　　　" + Num6_11 + " / " + Num7_11 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "大陸　Ａｎｄｒｏｉｄ完整版（未）:　　" + Num6_12 + " / " + Num7_12 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "ＩｏｓＤｅｂｕｇ:　　　　　　　　　　" + Num6_13 + " / " + Num7_13 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "ＡｎｄｒｏｉｄＤｅｂｕｇ:　　　　　　" + Num6_14 + " / " + Num7_14 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣老李　ＭｙＣａｒｄ:　　　　　　　" + Num6_15 + " / " + Num7_15 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　ＭｙＣａｒｄ:　　　　　　　　　" + Num6_16 + " / " + Num7_16 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣　Ｇｏｏｇｌｅ完整版:　　　　　　" + Num6_17 + " / " + Num7_17 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "台灣老李　ＭｙＣａｒｄ完整版:　　　　" + Num6_18 + " / " + Num7_18 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "香港　Ｇｏｏｇｌｅ完整版:　　　　　　" + Num6_19 + " / " + Num7_19 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "松崗:　　　　　　　　　　　　　　　　" + Num6_20 + " / " + Num7_20 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "新加坡　Ｉｏｓ:　　　　　　　　　　　" + Num6_21 + " / " + Num7_21 + " / " + Num8_10 + "\r\n" + "\r\n";
                                    textBox_Creat.Text += "新加坡　Ａｎｄｒｏｉｄ完整版:　　　　" + Num6_22 + " / " + Num7_22 + " / " + Num8_10 + "\r\n" + "\r\n";

                                    #endregion
                                }
                            }

                            else
                            {
                                textBox_Creat.Text = "資料庫無資料可分析";
                            }


                        }
                        catch
                        {
                            MessageBox.Show("資料查詢錯誤");

                        }

                        finally
                        {



                        }

                        break;
                }

            }
            else
            {
                MessageBox.Show("8/29前沒有資料");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox_DisLv.Text = "";
            textBox_DisDimon.Text = "";
            textBox_DisLog.Text = "";
            textBox_DisName.Text = "";
            textBox_DisVip.Text = "";
            textBox_CardBox.Text = "";
            textBox_Point25W.Text = "";
            label_inquire.Text = "目前查詢日期:";
            label_Country.Text = "目前查詢國家:";
            textBox_Creat.Text = "";
            richTextBox_ChargeList.Text = "";
        }

        private void comboBox_TypePoint_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_TypePoint.SelectedIndex)
            {
                case 0:
                    SportPoint = 250000;
                    break;

                case 1:
                    SportPoint = 300000;
                    break;

                case 2:
                    SportPoint = 400000;
                    break;

                case 3:
                    SportPoint = 500000;
                    break;
            }
        }

        string PassWordCode = ""; //從sql抓到的密碼原始碼

        private void button_Psearch_Click(object sender, EventArgs e)
        {
            if (textBox_SendPassword.Text != SendPassWord)
            {
                MessageBox.Show("查詢密碼錯誤");
            }

            else
            {
                textBox_PassWordDis.Text = "";
                label_PType.Text = "";

                try
                {
                    string Type = "";
                    string Tsql = "";

                    switch (comboBox_PSeachMethod.SelectedIndex)
                    {
                        case 0:

                            Tsql = "SELECT MainID, Type, UserAccount, UserPassword, Account FROM Account WHERE (UserAccount = '" + textBox_AccountInput.Text + "')";

                            if (textBox_AccountInput.Text != string.Empty)
                            {
                                string[] dr = ",".Split(',');

                                switch (comboBox_PCountry.SelectedIndex)
                                {
                                    case 0:

                                        dr = TWPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_MainIDDis.Text = dr[0];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;

                                    case 1:

                                        dr = HKPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_MainIDDis.Text = dr[0];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;

                                    case 2:

                                        dr = CNsql.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_MainIDDis.Text = dr[0];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;

                                    case 3:

                                        dr = SGPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_MainIDDis.Text = dr[0];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;

                                    case 4:

                                        dr = MAPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_MainIDDis.Text = dr[0];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;

                                    case 5:

                                        dr = THPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_MainIDDis.Text = dr[0];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;


                                }
                            }
                            else
                            {
                                MessageBox.Show("請輸入正確帳號");
                            }

                            break;

                        case 1:

                            Tsql = "SELECT MainID, Type, UserAccount, UserPassword, Account FROM Account WHERE  (MainID = '" + textBox_MainIDDis.Text + "')";

                            if (textBox_MainIDDis.Text != string.Empty)
                            {
                                string[] dr = ",".Split(',');

                                switch (comboBox_PCountry.SelectedIndex)
                                {
                                    case 0:

                                        dr = TWPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_AccountInput.Text = dr[2];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;

                                    case 1:

                                        dr = HKPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_AccountInput.Text = dr[2];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;

                                    case 2:

                                        dr = CNsql.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_AccountInput.Text = dr[2];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;

                                    case 3:

                                        dr = SGPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_AccountInput.Text = dr[2];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;

                                    case 4:

                                        dr = MAPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_AccountInput.Text = dr[2];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;

                                    case 5:

                                        dr = THPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "MainID,Type,UserAccount,UserPassword,Account").Split(',');

                                        textBox_AccountInput.Text = dr[2];
                                        label_PType.Text = dr[1];
                                        Type = dr[1];
                                        PassWordCode = dr[3];
                                        textBox_OnlyCodDis.Text = dr[4];

                                        break;

                                }
                            }
                            else
                            {
                                MessageBox.Show("請輸入正確帳號");
                            }

                            break;

                    }

                    if (Type == "SN" || Type == "16")
                    {
                        textBox_PassWordDis.Text = GSecurity.Decrypt(PassWordCode);
                    }
                    else if (Type == "FB")
                    {
                        textBox_PassWordDis.Text = "FB綁定無法查詢密碼";
                    }
                    else
                    {
                        textBox_PassWordDis.Text = "無法查詢密碼";
                    }

                }
                catch
                {
                    MessageBox.Show("查詢錯誤");
                }

            }

        }

        private void button_PClean_Click(object sender, EventArgs e)
        {
            textBox_AccountInput.Text = "";
            textBox_MainIDDis.Text = "";
            textBox_PassWordDis.Text = "";
            textBox_SendPassword.Text = "";
            label_PType.Text = "";
            textBox_OnlyCodDis.Text = "";
        }

        private void comboBox_PSeachMethod_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_PSeachMethod.SelectedIndex)
            {
                case 0:
                    textBox_MainIDDis.Text = "";
                    textBox_MainIDDis.ReadOnly = true;
                    textBox_AccountInput.ReadOnly = false;
                    break;

                case 1:
                    textBox_AccountInput.Text = "";
                    textBox_AccountInput.ReadOnly = true;
                    textBox_MainIDDis.ReadOnly = false;
                    break;

            }
        }

        private void checkBox_Continu_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_Continu.Checked)
            {
                MessageBox.Show("連續發送安全機制已關閉，請小心使用");
            }
            else
            {
                textBox_giftMGPassword.Text = "";
            }
        }

        private void button_OpenBTG_Click(object sender, EventArgs e)
        {
            Form2 NewForm = new Form2();
            NewForm.FormClosed += new FormClosedEventHandler(NewForm_FormClosed);

            NewForm.Show();
            button_OpenBTG.Enabled = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form3 NewForm = new Form3();
            NewForm.FormClosed += new FormClosedEventHandler(NewForm_Form3Closed);

            NewForm.Show();
            button_StopAccount.Enabled = false;
        }

        void NewForm_Form3Closed(object sender, FormClosedEventArgs e)
        {
            button_StopAccount.Enabled = true;
        }

        private void checkBox_ContinuFB_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_ContinuFB.Checked)
            {
                MessageBox.Show("連續發送安全機制已關閉，請小心使用");
            }
            else
            {
                textBox_FBPassWord.Text = "";
            }
        }

        private void textBox_SHP_TextChanged(object sender, EventArgs e)
        {
            if (textBox_SHP.Text != string.Empty)
            {
                try
                {
                    SHP = Convert.ToInt32(textBox_SHP.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_SHP.Text = textBox_SHP.Text.Substring(0, textBox_SHP.Text.Length - 1);
                }
            }
        }

        private void textBox_SATK_TextChanged(object sender, EventArgs e)
        {
            if (textBox_SATK.Text != string.Empty)
            {
                try
                {
                    SATK = Convert.ToInt32(textBox_SATK.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_SATK.Text = textBox_SATK.Text.Substring(0, textBox_SATK.Text.Length - 1);
                }
            }
        }

        private void textBox_SRHP_TextChanged(object sender, EventArgs e)
        {
            if (textBox_SRHP.Text != string.Empty)
            {
                try
                {
                    SRHP = Convert.ToInt32(textBox_SRHP.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_SRHP.Text = textBox_SRHP.Text.Substring(0, textBox_SRHP.Text.Length - 1);
                }
            }
        }

        private void textBox_PHP_TextChanged(object sender, EventArgs e)
        {
       
        }

        int PomoNum = 0;

        private void button_PomoCreat_Click(object sender, EventArgs e)
        {
            if (textBox_PomoPassWord.Text != SendPassWord)
            {
                MessageBox.Show("密碼錯誤");
            }

            else
            {

                string outputstr = "";
                string[] Test = new string[PomoNum];

                for (int i = 0; i < PomoNum; i++)
                {

                    string snr = GetRandomPassword(Convert.ToInt32(comboBox_PomoType.SelectedItem));

                    string Tsql = "SELECT CardCode FROM Charger_PomoCards WHERE (CardCode = '" + snr + "')";
                    bool PomoAlive = false;

                    try
                    {
                        string re = "";

                        switch (comboBox_PomoCountry.SelectedIndex)
                        {
                            case 0:
                                 re = TWPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "CardCode");
                                break;

                            case 1:
                                 re = HKPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "CardCode");
                                break;

                            case 2:
                                re = CNsql.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "CardCode");
                                break;

                            case 3:
                                 re = SGPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "CardCode");
                                break;

                            case 4:
                                 re = MAPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "CardCode");
                                break;

                        }
                        
                        if (re == string.Empty)
                        {
                            for (int x = 0; x < i; x++)
                            {
                                if (snr == Test[x])
                                {
                                    PomoAlive = true;
                                }
                            }

                            if (PomoAlive == false)
                            {
                                Test[i] = snr;
                                outputstr += snr + "\r\n";
                            }
                            else
                            {
                                i--;
                            }

                        }
                        else
                        {
                            i--;
                        }
                    }
                    catch
                    {
                        i--;
                    }                    
                }
                                
                richTextBox_PomoDis.Text = outputstr;
                textBox_PomoPassWord.Text = "";
                MessageBox.Show("產生完畢");
            }
        }

        private void textBox_PomoNum_TextChanged(object sender, EventArgs e)
        {
            if (textBox_PomoNum.Text != string.Empty)
            {
                try
                {
                    PomoNum = Convert.ToInt32(textBox_PomoNum.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_PomoNum.Text = textBox_PomoNum.Text.Substring(0, textBox_PomoNum.Text.Length - 1);
                }
            }
        }

        private System.Security.Cryptography.RNGCryptoServiceProvider rngp = new System.Security.Cryptography.RNGCryptoServiceProvider();
        private byte[] rb = new byte[4];

        /// <summary>
        /// 產生一個非負數的亂數
        /// </summary>
        private int Next()
        {
            rngp.GetBytes(rb);
            int value = BitConverter.ToInt32(rb, 0);
            if (value < 0) value = -value;
            return value;
        }
        /// <summary>
        /// 產生一個非負數且最大值 max 以下的亂數
        /// </summary>
        /// <param name="max">最大值</param>
        private int Next(int max)
        {
            rngp.GetBytes(rb);
            int value = BitConverter.ToInt32(rb, 0);
            value = value % (max + 1);
            if (value < 0) value = -value;
            return value;
        }
        /// <summary>
        /// 產生一個非負數且最小值在 min 以上最大值在 max 以下的亂數
        /// </summary>
        /// <param name="min">最小值</param>
        /// <param name="max">最大值</param>
        private int Next(int min, int max)
        {
            int value = Next(max - min) + min;
            return value;
        }

        public string GetRandomPassword(int length)
        {
            StringBuilder sb = new StringBuilder();
            char[] chars = "0123456789abcdefghijklmnopqrstuvwxyz".ToCharArray();

            for (int i = 0; i < length; i++)
            {
                sb.Append(chars[this.Next(chars.Length - 1)]);
            }
            string Password = sb.ToString();
            return Password;
        }

        private void textBox_AHP_TextChanged(object sender, EventArgs e)
        {
            if (textBox_AHP.Text != string.Empty)
            {
                try
                {
                    AHP = Convert.ToInt32(textBox_AHP.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_AHP.Text = textBox_AHP.Text.Substring(0, textBox_AHP.Text.Length - 1);
                }
            }
        }

        private void textBox_AATK_TextChanged(object sender, EventArgs e)
        {
            if (textBox_AATK.Text != string.Empty)
            {
                try
                {
                    AATK = Convert.ToInt32(textBox_AATK.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_AATK.Text = textBox_AATK.Text.Substring(0, textBox_AATK.Text.Length - 1);
                }
            }
        }

        private void textBox_ARHP_TextChanged(object sender, EventArgs e)
        {
            if (textBox_ARHP.Text != string.Empty)
            {
                try
                {
                    ARHP = Convert.ToInt32(textBox_ARHP.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_ARHP.Text = textBox_ARHP.Text.Substring(0, textBox_ARHP.Text.Length - 1);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
        }

        bool Start = false;

        private void comboBox_PomoType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox_PomoType.SelectedIndex)
            {
                case 0:
                    if (Start)
                    {
                        MessageBox.Show("一組帳號只能儲一組序號，序號使用過即銷毀");
                    }                    
                    Start = true;

                    break;

                case 1:
                    MessageBox.Show("一組帳號可儲值多組序號，序號使用過即銷毀");
                    break;

                case 2:
                    MessageBox.Show("一組帳號只能儲一組序號，序號可重複使用");
                    break;
                    
            }
        }

        private void button_6Pomo_Click_1(object sender, EventArgs e)
        {
             string Tsql = "SELECT CardCode FROM Charger_PomoCards WHERE (CardCode = '" + textBox_6Pomo.Text + "')";

             try
             {
                 string re = "";

                 switch (comboBox_TCountry.SelectedIndex)
                 {
                     case 0:
                         re = TWPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "CardCode");
                         break;

                     case 1:
                         re = HKPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "CardCode");
                         break;

                     case 2:
                         re = CNsql.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "CardCode");
                         break;
                   
                     default:
                         re = TWPuzzle.Get_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "CardCode");
                         break;

                 }

                 if (re == string.Empty)
                 {
                     MessageBox.Show("可使用");
                 }
                 else
                 {
                     MessageBox.Show("已重複，不可使用");
                 }

             }
             catch
             {
                 MessageBox.Show("查詢錯誤，請重新查詢");
             }    
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        int EventNum = 0; //要取消的Event編號

        private void button_CEUse_Click(object sender, EventArgs e)
        {            
            string sr = "";

            if (textBox_CEPassWord.Text != SendPassWord)
            {
                MessageBox.Show("密碼錯誤");
            }

            else if (textBox_CENumber.Text == string.Empty)
            {
                MessageBox.Show("Event編號不可為空");
            }
            else
            {
                try
                {
                    switch (comboBox_CECountry.SelectedIndex)
                    {
                        case 0:
                            sr = SaveTW.CancelEvent(EventNum, Convert.ToInt32(comboBox_CEServer.SelectedIndex), ServerPassWord);
                            break;

                        case 1:
                            sr = SaveHK.CancelEvent(EventNum, Convert.ToInt32(comboBox_CEServer.SelectedIndex), ServerPassWord);
                            break;

                        case 2:
                            sr = SaveCN.CancelEvent(EventNum, Convert.ToInt32(comboBox_CEServer.SelectedIndex), ServerPassWord);
                            break;

                        case 3:
                            sr = SaveSG.CancelEvent(EventNum, Convert.ToInt32(comboBox_CEServer.SelectedIndex), ServerPassWord);
                            break;

                        case 4:
                            sr = SaveMY.CancelEvent(EventNum, Convert.ToInt32(comboBox_CEServer.SelectedIndex), ServerPassWord);
                            break;

                        case 5:
                            sr = SaveDebug.CancelEvent(EventNum, Convert.ToInt32(comboBox_CEServer.SelectedIndex), ServerPassWord);
                            break;

                        case 6:
                            sr = SaveTH.CancelEvent(EventNum, Convert.ToInt32(comboBox_CEServer.SelectedIndex), ServerPassWord);
                            break;
                    }

                    MessageBox.Show(sr);
                }
                catch
                {
                    MessageBox.Show("取消失敗");
                }
                finally
                {
                    //textBox_CEPassWord.Text = "";
                }
            }           
            
        }

        private void textBox_CENumber_TextChanged(object sender, EventArgs e)
        {
            if (textBox_CENumber.Text != string.Empty)
            {
                try
                {
                    EventNum = Convert.ToInt32(textBox_CENumber.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_CENumber.Text = textBox_CENumber.Text.Substring(0, textBox_CENumber.Text.Length - 1);
                }
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void textBox_Level_TextChanged(object sender, EventArgs e)
        {
            if (textBox_Level.Text != string.Empty)
            {
                try
                {
                    ItemLV = Convert.ToInt32(textBox_Level.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_Level.Text = textBox_Level.Text.Substring(0, textBox_Level.Text.Length - 1);
                }
            }
        }
        
        private DataSet LoginCount(int Country,int Type,DateTime StartDate,int days,int ChoiseServer)
        {
            string Tsql = "";
            string Server = "";

            #region 選擇server
            if (ChoiseServer == 0)
            {
                Server = "";
            }
            else
            {
                switch (comboBox_CRServer.SelectedIndex)
                {
                    case 0:
                        Server = "";
                        break;
                    case 1:
                        Server = "AND (ServerID = '1')";
                        break;
                    case 2:
                        Server = "AND (ServerID = '2')";
                        break;
                    case 3:
                        Server = "AND (ServerID = '3')";
                        break;
                    case 4:
                        Server = "AND (ServerID = '4')";
                        break;
                    case 5:
                        Server = "AND (ServerID = '5')";
                        break;
                    case 6:
                        Server = "AND (ServerID = '6')";
                        break;
                    case 7:
                        Server = "AND (ServerID = '7')";
                        break;
                    case 8:
                        Server = "AND (ServerID = '8')";
                        break;
                    case 9:
                        Server = "AND (ServerID = '9')";
                        break;
                    case 10:
                        Server = "AND (ServerID = '10')";
                        break;
                }
            }
            #endregion

            if (Type == 0)
            {
                Tsql = "SELECT UserID FROM Login_" + StartDate.ToString("yyyyMMdd") + " WHERE (Kind = '15') " + Server;

                for (int x = 0; x < days - 1; x++)
                {
                    Tsql += "UNION SELECT UserID FROM Login_" + StartDate.AddDays(x + 1).ToString("yyyyMMdd") + " WHERE (Kind = '15') " + Server;
                }

                Tsql += "GROUP BY UserID ORDER BY UserID";
            }
            else
            {
                Tsql = "SELECT UserID FROM Role_" + StartDate.ToString("yyyyMMdd") + " WHERE (Kind = '19') " + Server;

                for (int x = 0; x < days - 1; x++)
                {
                    Tsql += " UNION SELECT UserID FROM Role_" + StartDate.AddDays(x + 1).ToString("yyyyMMdd") + " WHERE (Kind = '19') " + Server;
                }

                Tsql += " GROUP BY UserID ORDER BY UserID";
            }

            try
            {
                DataSet ds = new DataSet();

                switch (Country)
                {
                    case 0:
                        ds = TWsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        break;

                    case 1:
                        ds = HKsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        break;

                    case 2:
                        ds = CNsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        break;
                }
                              
                return ds;
            }
            catch
            {
                return null;
            }            

        }


        private DataSet LoginCountMony(int Country, int Type, DateTime StartDate, int days, int ChoiseServer)
        {
            string Tsql = "";
            string Server = "";

            #region 選擇server
            if (ChoiseServer == 0)
            {
                Server = "";
            }
            else
            {
                switch (ChoiseServer)
                {
                    case 0:
                        Server = "";
                        break;
                    case 1:
                        Server = "AND (ServerID = '1')";
                        break;
                    case 2:
                        Server = "AND (ServerID = '2')";
                        break;
                    case 3:
                        Server = "AND (ServerID = '3')";
                        break;
                    case 4:
                        Server = "AND (ServerID = '4')";
                        break;
                    case 5:
                        Server = "AND (ServerID = '5')";
                        break;
                    case 6:
                        Server = "AND (ServerID = '6')";
                        break;
                    case 7:
                        Server = "AND (ServerID = '7')";
                        break;
                    case 8:
                        Server = "AND (ServerID = '8')";
                        break;
                    case 9:
                        Server = "AND (ServerID = '9')";
                        break;
                    case 10:
                        Server = "AND (ServerID = '10')";
                        break;
                }
            }
            #endregion

            if (Type == 0)
            {
                Tsql = "SELECT Kind,UserID, SUM(CAST(V2 AS int)) AS _Money FROM   (";

                Tsql += "SELECT  ID, Kind, ServerID, Cause, UserID, PuzzleWeb, V1, V2 FROM   Money_" + StartDate.ToString("yyyyMMdd") + " WHERE (Kind = '1') " + Server;

                for (int x = 0; x < days - 1; x++)
                {
                    Tsql += "UNION ALL SELECT     ID, Kind, ServerID, Cause, UserID, PuzzleWeb, V1, V2  FROM Money_" + StartDate.AddDays(x + 1).ToString("yyyyMMdd") + " WHERE (Kind = '1') " + Server;
        }

                Tsql += " ) AS derivedtbl_1 WHERE  (Cause = '10') OR (Cause = '13') OR (Cause = '14') OR (Cause = '15') OR (Cause = '5') OR (Cause = '16') OR (Cause = '17') OR (Cause = '18') AND (ServerID = '10') GROUP BY Kind,UserID,ServerID HAVING (Kind = '1')" + Server;
            }
            else
            {
                Tsql = "SELECT UserID FROM Role_" + StartDate.ToString("yyyyMMdd") + " WHERE (Kind = '19') " + Server;

                for (int x = 0; x < days - 1; x++)
                {
                    Tsql += " UNION SELECT UserID FROM Role_" + StartDate.AddDays(x + 1).ToString("yyyyMMdd") + " WHERE (Kind = '19') " + Server;
                }

                //請問玩家ID不需
                //Tsql += " GROUP BY UserID ORDER BY UserID";
            }

            try
            {
                DataSet ds = new DataSet();

                switch (Country)
                {
                    case 0:
                        ds = TWsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        break;

                    case 1:
                        ds = HKsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        break;

                    case 2:
                        ds = CNsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        break;
                }

                return ds;
            }
            catch
            {
                return null;
            }

        }


        private void button_LCStart_Click(object sender, EventArgs e)
        {
            if (dateTimePicker_LCStart.Value > dateTimePicker_LCEnd.Value)
            {
                MessageBox.Show("日期設定錯誤");
            }

            else if (dateTimePicker_LCEnd.Value > DateTime.Now)
            {
                MessageBox.Show("無法查詢未來的記錄");
            }

            else
            {
                            
                int days = (dateTimePicker_LCEnd.Value - dateTimePicker_LCStart.Value).Days +1;

                //string Tsql = "SELECT UserID FROM Login_" + dateTimePicker_LCStart.Value.ToString("yyyyMMdd") + " " + Server + " GROUP BY UserID ORDER BY UserID;";

                //for (int x = 0; x < days-1; x++)
                //{
                //    Tsql += "SELECT UserID FROM Login_" + dateTimePicker_LCStart.Value.AddDays(x + 1).ToString("yyyyMMdd") + " " + Server + " GROUP BY UserID ORDER BY UserID;";
                //}
                                
                try
                {
                    DataSet ds = LoginCount(comboBox_LCCountry.SelectedIndex, 0, dateTimePicker_LCStart.Value, days, 0);

                    //string[] MainID = new string[0];
                    string MainID = "";
                    int AccountNum = 0;
                    int[] ServerCount = new int[11];

                    #region 建立初始數據

                    //for (int x = 0; x < ds.Tables[0].Rows.Count; x++)
                    //{
                    //    Array.Resize(ref MainID, MainID.Length + 1);
                    //    MainID[x] = ds.Tables[0].Rows[x]["UserID"].ToString();
                    //    AccountNum++;
                    //}

                    #endregion

                    #region 處理後續天數

                    //if (days > 0)
                    //{
                    //    for (int Total = 1; Total < days; Total++) //迴圈看有幾天
                    //    {                            
                    //        for (int x = 0; x < ds.Tables[Total].Rows.Count; x++) //迴圈該天的每一筆
                    //        {
                    //            bool Alive = false;

                    //            for (int AllGroup = 0; AllGroup < MainID.Length; AllGroup++) //驗證該MainID有沒有出現過
                    //            {
                    //                if (ds.Tables[Total].Rows[x]["UserID"].ToString() == MainID[AllGroup])
                    //                {
                    //                    Alive = true;
                    //                }
                    //            }

                    //            if (Alive == false)
                    //            {
                    //                Array.Resize(ref MainID, MainID.Length + 1);
                    //                MainID[AccountNum] = ds.Tables[Total].Rows[x]["UserID"].ToString();  
                    //                AccountNum++;                                                                      
                    //            }
                    //        }
                    //    }
                    //}

                    #endregion

                    #region 合併表處理
                    
                    for (int x = 0; x < ds.Tables[0].Rows.Count; x++)
                    {
                        MainID += ds.Tables[0].Rows[x]["UserID"].ToString() + "\n";
                        AccountNum++;

                        switch (ds.Tables[0].Rows[x]["UserID"].ToString().Substring(6, 2))
                        {
                            case "01":
                                ServerCount[1]++;
                                break;

                            case "02":
                                ServerCount[2]++;
                                break;

                            case "03":
                                ServerCount[3]++;
                                break;

                            case "04":
                                ServerCount[4]++;
                                break;

                            case "05":
                                ServerCount[5]++;
                                break;

                            case "06":
                                ServerCount[6]++;
                                break;

                            case "07":
                                ServerCount[7]++;
                                break;

                            case "08":
                                ServerCount[8]++;
                                break;

                            case "09":
                                ServerCount[9]++;
                                break;

                            case "10":
                                ServerCount[10]++;
                                break;

                            default:
                                ServerCount[0]++;
                                break;
                        }
                    }

                    #endregion

                    #region 留存率處理

                    string RemainDis = "0";

                    if (checkBox_Remain.Checked)
                    {
                        int RDays = (dateTimePicker_REnd.Value - dateTimePicker_RStart.Value).Days + 1;
                        string[] Radd = new string[0];
                        string[] TM = MainID.Substring(0, MainID.Length - 1).Replace("\n", "@").Split('@');
                        int Total = 0;

                        DataSet Rds = LoginCount(comboBox_LCCountry.SelectedIndex, 0, dateTimePicker_RStart.Value, RDays, 0);

                        for (int i = 0; i < Rds.Tables[0].Rows.Count; i++)
                        {
                            bool alive = false;

                            for (int y = 0; y < TM.Length; y++)
                            {
                                if (TM[y] == Rds.Tables[0].Rows[i]["UserID"].ToString())
                                {
                                    alive = true;
                                }
                            }

                            if (alive == true)
                            {
                                Array.Resize(ref Radd, Radd.Length + 1);
                                Radd[Total] += Rds.Tables[0].Rows[i]["UserID"].ToString();
                                Total++;
                            }

                        }

                        double ReMain = Radd.Length;
                        double TotalReMain = TM.Length;

                        RemainDis = (ReMain / TotalReMain).ToString("P");

                    }

                    #endregion

                    #region 資料顯示

                    richTextBox_LCDis.Text = "查詢日期 : " + dateTimePicker_LCStart.Value.ToString("yyyy-MM-dd") + " 到 " + dateTimePicker_LCEnd.Value.ToString("yyyy-MM-dd") + "\n";
                    richTextBox_LCDis.Text += "共 " + days + " 天" + "\n";
                    richTextBox_LCDis.Text += "不重複登入人數共 : " + AccountNum + " 人" + "\n";
                    richTextBox_LCDis.Text += "Server1 : " + ServerCount[1] + " 人" + "\n";
                    richTextBox_LCDis.Text += "Server2 : " + ServerCount[2] + " 人" + "\n";
                    richTextBox_LCDis.Text += "Server3 : " + ServerCount[3] + " 人" + "\n";
                    richTextBox_LCDis.Text += "Server4 : " + ServerCount[4] + " 人" + "\n";
                    richTextBox_LCDis.Text += "Server5 : " + ServerCount[5] + " 人" + "\n";
                    richTextBox_LCDis.Text += "Server6 : " + ServerCount[6] + " 人" + "\n";
                    richTextBox_LCDis.Text += "Server7 : " + ServerCount[7] + " 人" + "\n";
                    richTextBox_LCDis.Text += "Server8 : " + ServerCount[8] + " 人" + "\n";
                    richTextBox_LCDis.Text += "Server9 : " + ServerCount[9] + " 人" + "\n";
                    richTextBox_LCDis.Text += "Server10 : " + ServerCount[10] + " 人" + "\n";
                    richTextBox_LCDis.Text += "其他 : " + ServerCount[0] + " 人" + "\n";

                    if (checkBox_Remain.Checked)
                    {
                        richTextBox_LCDis.Text += "\n" + "留存率 : " + RemainDis + "\n";
                    }
                                        
                    if (checkBox_LCShowID.Checked)
                    {
                        richTextBox_LCDis.Text += "\n" + "登入玩家ID : " + "\n";
                        richTextBox_LCDis.Text += MainID;
                    }

                    #endregion

                    MessageBox.Show("資料來了");
                }
                catch
                {
                    MessageBox.Show("查詢失敗");
                }
            }
           
        }

        private void button_LCClean_Click(object sender, EventArgs e)
        {
            richTextBox_LCDis.Text = "";
        }

        private void button_CHClean_Click(object sender, EventArgs e)
        {
            richTextBox_CHDis.Text = "";
        }

        private void button_CHStart_Click(object sender, EventArgs e)
        {
            if (dateTimePicker_CHStart.Value > dateTimePicker_LCEnd.Value)
            {
                MessageBox.Show("日期設定錯誤");
            }

            else if (dateTimePicker_CHEnd.Value > DateTime.Now)
            {
                MessageBox.Show("無法查詢未來的記錄");
            }

            else
            {
                string Server = "";

                switch (comboBox_CHServer.SelectedIndex)
                {
                    case 0:
                        Server = "";
                        break;
                    case 1:
                        Server = "AND (ServerID = '1')";
                        break;
                    case 2:
                        Server = "AND (ServerID = '2')";
                        break;
                    case 3:
                        Server = "AND (ServerID = '3')";
                        break;
                    case 4:
                        Server = "AND (ServerID = '4')";
                        break;
                    case 5:
                        Server = "AND (ServerID = '5')";
                        break;
                    case 6:
                        Server = "AND (ServerID = '6')";
                        break;
                    case 7:
                        Server = "AND (ServerID = '7')";
                        break;
                    case 8:
                        Server = "AND (ServerID = '8')";
                        break;
                    case 9:
                        Server = "AND (ServerID = '9')";
                        break;
                    case 10:
                        Server = "AND (ServerID = '10')";
                        break;
                }


                int days = (dateTimePicker_CHEnd.Value - dateTimePicker_CHStart.Value).Days + 1;

                string Tsql = "SELECT UserID FROM Money_" + dateTimePicker_CHStart.Value.ToString("yyyyMMdd") + " WHERE (Kind = 1) AND (Cause = 10 OR Cause = 13 OR Cause = 14 OR Cause = 15 OR Cause = 22) " + Server + " GROUP BY UserID ORDER BY UserID;";

                for (int x = 0; x < days - 1; x++)
                {
                    Tsql += "SELECT UserID FROM Money_" + dateTimePicker_CHStart.Value.AddDays(x + 1).ToString("yyyyMMdd") + " WHERE (Kind = 1) AND (Cause = 10 OR Cause = 13 OR Cause = 14 OR Cause = 15 OR Cause = 22) " + Server + " GROUP BY UserID ORDER BY UserID;";
                }

                try
                {
                    DataSet ds = new DataSet();

                    switch (comboBox_CHCountry.SelectedIndex)
                    {
                        case 0:
                            ds = TWsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;

                        case 1:
                            ds = HKsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;

                        case 2:
                            ds = CNsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                            break;
                    }

                    string[] MainID = new string[0];
                    int AccountNum = 0;

                    #region 建立初始數據

                    for (int x = 0; x < ds.Tables[0].Rows.Count; x++)
                    {
                        Array.Resize(ref MainID, MainID.Length + 1);
                        MainID[x] = ds.Tables[0].Rows[x]["UserID"].ToString();
                        AccountNum++;
                    }

                    #endregion

                    #region 處理後續天數

                    if (days > 0)
                    {
                        for (int Total = 1; Total < days; Total++) //迴圈看有幾天
                        {
                            for (int x = 0; x < ds.Tables[Total].Rows.Count; x++) //迴圈該天的每一筆
                            {
                                bool Alive = false;

                                for (int AllGroup = 0; AllGroup < MainID.Length; AllGroup++) //驗證該MainID有沒有出現過
                                {
                                    if (ds.Tables[Total].Rows[x]["UserID"].ToString() == MainID[AllGroup])
                                    {
                                        Alive = true;
                                    }
                                }

                                if (Alive == false)
                                {
                                    Array.Resize(ref MainID, MainID.Length + 1);
                                    MainID[AccountNum] = ds.Tables[Total].Rows[x]["UserID"].ToString();
                                    AccountNum++;
                                }
                            }
                        }
                    }

                    #endregion

                    richTextBox_CHDis.Text = "查詢日期 : " + dateTimePicker_CHStart.Value.ToString("yyyy-MM-dd") + " 到 " + dateTimePicker_CHEnd.Value.ToString("yyyy-MM-dd") + "\n";
                    richTextBox_CHDis.Text += "共 " + days + " 天" + "\n";
                    richTextBox_CHDis.Text += "不重複儲值人數共 : " + AccountNum + " 人" + "\n" + "\n";

                    if (checkBox_CHShowID.Checked)
                    {
                        richTextBox_CHDis.Text += "儲值玩家ID : " + "\n";

                        for (int x = 0; x < MainID.Length; x++)
                        {
                            richTextBox_CHDis.Text += MainID[x] + "\n";
                        }
                    }

                    MessageBox.Show("資料來了");
                }
                catch
                {
                    MessageBox.Show("查詢失敗");
                }
            }

        }

        bool FtpStatus = false; //FTP有沒有人使用
        string ControNum = ""; //獨立單號

        private void button_FTPInquire_Click(object sender, EventArgs e)
        {
            string Tsql = "SELECT UserID, StartDate, ControlNumber FROM FTP_UserControl WHERE (Status = '1')";

            string[] Status = new string[3];
            try
            {
                switch (comboBox_FTPCountry.SelectedIndex)
                {
                        #region Debug
                    case 0:

                        string SqlLink_Recorder_Debug = "Data Source=192.168.38.15; Initial Catalog=Analysis; User ID='sa'; Password='Steve*Jobs@1955#2011'";                        
                        
                        using (SqlConnection SCN = new SqlConnection(SqlLink_Recorder_Debug))
                        {
                            SqlCommand sda = new SqlCommand(Tsql, SCN);

                            SCN.Open();

                            SqlDataReader reader = sda.ExecuteReader();

                            if (reader.Read())
                            {
                                Status[0] = reader[0].ToString();
                                Status[1] = reader[1].ToString();
                                Status[2] = reader[2].ToString();
                                FtpStatus = true;
                            }
                            else
                            {
                                FtpStatus = false;
                            }

                            SCN.Close();
                        }

                        break;

                        #endregion

                        #region TW
                    case 1:
                        Status = TWsql.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                        if (Status[0] != string.Empty)
                        {
                            FtpStatus = true;
                        }
                        else
                        {
                            FtpStatus = false;
                        }
                        break;
                        #endregion

                        #region HK
                    case 2:
                        Status = HKsql.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                        if (Status[0] != string.Empty)
                        {
                            FtpStatus = true;
                        }
                        else
                        {
                            FtpStatus = false;
                        }
                        break;
                        #endregion

                        #region CN
                    case 3:
                        Status = CNsql.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                        if (Status[0] != string.Empty)
                        {
                            FtpStatus = true;
                        }
                        else
                        {
                            FtpStatus = false;
                        }
                        break;
                        #endregion

                        #region SG
                    case 4:
                        Status = SGPuzzle.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                        if (Status[0] != string.Empty)
                        {
                            FtpStatus = true;
                        }
                        else
                        {
                            FtpStatus = false;
                        }
                        break;
                        #endregion

                        #region MY
                    case 5:
                        Status = MAPuzzle.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                        if (Status[0] != string.Empty)
                        {
                            FtpStatus = true;
                        }
                        else
                        {
                            FtpStatus = false;
                        }
                        break;
                        #endregion

                        #region TH
                    case 6:
                        Status = THPuzzle.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                        if (Status[0] != string.Empty)
                        {
                            FtpStatus = true;
                        }
                        else
                        {
                            FtpStatus = false;
                        }
                        break;
                        #endregion

                        #region KR
                    case 7:
                        Status = KRPuzzle.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                        if (Status[0] != string.Empty)
                        {
                            FtpStatus = true;
                        }
                        else
                        {
                            FtpStatus = false;
                        }
                        break;
                        #endregion
                }

                if (FtpStatus)
                {
                    label_FTPDis.Text = "使用狀態 :         使用者 :  " + Status[0] + "         使用時間 : " + Status[1];
                    button_FTPUntight.Enabled = true;
                    ControNum = Status[2].ToString();
                }
                else
                {
                    label_FTPDis.Text = "使用狀態 :         " + comboBox_FTPCountry.SelectedItem.ToString() + " 無人使用";
                    button_FTPUntight.Enabled = false;
                    ControNum = "";
                }

            }
            catch
            {
                FtpStatus = false;
                label_FTPDis.Text = "使用狀態 :         " + comboBox_FTPCountry.SelectedItem.ToString() + " 無人使用";
                ControNum = "";
            }
                     
        }

        private void button_FTPUntight_Click(object sender, EventArgs e)
        {
            if (textBox_FTPPassword.Text != SendPassWord)
            {
                MessageBox.Show("解綁定密碼錯誤");
            }
            else if (ControNum == string.Empty)
            {
                label_FTPDis.Text = "使用狀態 :         資料錯誤，請重新查詢";
                FTPReset();
            }
            else
            {
                string[] Status = new string[3];

                string Tsql = "SELECT UserID, StartDate, ControlNumber FROM FTP_UserControl WHERE (Status = '1') AND (ControlNumber = '" + ControNum + "')";

                try
                {
                    switch (comboBox_FTPCountry.SelectedIndex)
                    {
                        #region Debug
                        case 0:

                            string SqlLink_Recorder_Debug = "Data Source=192.168.38.15; Initial Catalog=Analysis; User ID='sa'; Password='Steve*Jobs@1955#2011'";

                            using (SqlConnection SCN = new SqlConnection(SqlLink_Recorder_Debug))
                            {
                                SqlCommand sda = new SqlCommand(Tsql, SCN);

                                SCN.Open();

                                SqlDataReader reader = sda.ExecuteReader();

                                if (reader.Read())
                                {
                                    FtpStatus = true;
                                }
                                else
                                {
                                    FtpStatus = false;
                                }

                                SCN.Close();
                            }

                            break;

                        #endregion

                        #region TW
                        case 1:
                            Status = TWsql.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                            if (Status[0] != string.Empty)
                            {
                                FtpStatus = true;
                            }
                            else
                            {
                                FtpStatus = false;
                            }
                            break;
                        #endregion

                        #region HK
                        case 2:
                            Status = HKsql.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                            if (Status[0] != string.Empty)
                            {
                                FtpStatus = true;
                            }
                            else
                            {
                                FtpStatus = false;
                            }
                            break;
                        #endregion

                        #region CN
                        case 3:
                            Status = CNsql.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                            if (Status[0] != string.Empty)
                            {
                                FtpStatus = true;
                            }
                            else
                            {
                                FtpStatus = false;
                            }
                            break;
                        #endregion

                        #region SG
                        case 4:
                            Status = SGPuzzle.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                            if (Status[0] != string.Empty)
                            {
                                FtpStatus = true;
                            }
                            else
                            {
                                FtpStatus = false;
                            }
                            break;
                        #endregion

                        #region MY
                        case 5:
                            Status = MAPuzzle.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                            if (Status[0] != string.Empty)
                            {
                                FtpStatus = true;
                            }
                            else
                            {
                                FtpStatus = false;
                            }
                            break;
                        #endregion

                        #region TH
                        case 6:
                            Status = THPuzzle.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                            if (Status[0] != string.Empty)
                            {
                                FtpStatus = true;
                            }
                            else
                            {
                                FtpStatus = false;
                            }
                            break;
                        #endregion

                        #region KR
                        case 7:
                            Status = KRPuzzle.Get_SQL_Analysis_Command(GSecurity.Encrypt(Tsql), SqlPassWord, "UserID,StartDate,ControlNumber").Split(',');
                            if (Status[0] != string.Empty)
                            {
                                FtpStatus = true;
                            }
                            else
                            {
                                FtpStatus = false;
                            }
                            break;
                        #endregion
                    }
                }
                catch
                {
                    label_FTPDis.Text = "使用狀態 :         資料錯誤，請重新查詢";
                    FTPReset();
                }


                if (FtpStatus)
                {
                    try
                    {
                        string WTsql = "UPDATE FTP_UserControl SET Status = '2' WHERE (Status = '1') AND (ControlNumber = '" + ControNum + "')";
                        int Resault = 0;

                        switch (comboBox_FTPCountry.SelectedIndex)
                        {
                            #region Debug
                            case 0:

                                string SqlLink_Recorder_Debug = "Data Source=192.168.38.15; Initial Catalog=Analysis; User ID='sa'; Password='Steve*Jobs@1955#2011'";

                                using (SqlConnection SCN = new SqlConnection(SqlLink_Recorder_Debug))
                                {
                                    SqlCommand sda = new SqlCommand(WTsql, SCN);

                                    SCN.Open();

                                    Resault = sda.ExecuteNonQuery();

                                    SCN.Close();
                                }

                                break;

                            #endregion

                            #region TW
                            case 1:

                                Resault = TWsql.Set_SQL_Analysis_Command(GSecurity.Encrypt(WTsql), SqlPassWord);
                                
                                break;
                            #endregion

                            #region HK
                            case 2:

                                Resault = HKsql.Set_SQL_Analysis_Command(GSecurity.Encrypt(WTsql), SqlPassWord);

                                break;
                            #endregion

                            #region CN
                            case 3:

                                Resault = CNsql.Set_SQL_Analysis_Command(GSecurity.Encrypt(WTsql), SqlPassWord);

                                break;
                            #endregion

                            #region SG
                            case 4:

                                Resault = SGPuzzle.Set_SQL_Analysis_Command(GSecurity.Encrypt(WTsql), SqlPassWord);

                                break;
                            #endregion

                            #region MY
                            case 5:

                                Resault = MAPuzzle.Set_SQL_Analysis_Command(GSecurity.Encrypt(WTsql), SqlPassWord);

                                break;
                            #endregion

                            #region TH
                            case 6:

                                Resault = THPuzzle.Set_SQL_Analysis_Command(GSecurity.Encrypt(WTsql), SqlPassWord);

                                break;
                            #endregion

                            #region KR
                            case 7:

                                Resault = KRPuzzle.Set_SQL_Analysis_Command(GSecurity.Encrypt(WTsql), SqlPassWord);

                                break;
                            #endregion
                        }

                        if (Resault == 1)
                        {
                            label_FTPDis.Text = "使用狀態 :         解綁定成功";
                        }
                        else
                        {
                            label_FTPDis.Text = "使用狀態 :         解綁定失敗";
                        }
                    }
                    catch
                    {
                        label_FTPDis.Text = "使用狀態 :         解綁定失敗";
                    }

                    FTPReset();

                }
                else
                {
                    label_FTPDis.Text = "使用狀態 :         資料錯誤，請重新查詢";
                    FTPReset();
                }

            }

        }

        private void FTPReset()
        {
            FtpStatus = false;
            ControNum = "";
            button_FTPUntight.Enabled = false;
            textBox_FTPPassword.Text = "";
        }
                
        private void textBox_TLevel_TextChanged(object sender, EventArgs e)
        {
            if (textBox_TLevel.Text != string.Empty)
            {
                try
                {
                    TLevel = Convert.ToInt32(textBox_TLevel.Text);
                }
                catch
                {
                    MessageBox.Show("請填寫數字喔");
                    textBox_TLevel.Text = textBox_TLevel.Text.Substring(0, textBox_TLevel.Text.Length - 1);
                }
            }
        }

        private void checkBox_Remain_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_Remain.Checked)
            {
                dateTimePicker_RStart.Enabled = true;
                dateTimePicker_REnd.Enabled = true;
            }
            else
            {
                dateTimePicker_RStart.Enabled = false;
                dateTimePicker_REnd.Enabled = false;
            }
        }

        private void button_CRClean_Click(object sender, EventArgs e)
        {
            richTextBox_CRDis.Text = "";
        }


        public class Data
        {
            public string Id { set; get; }
            public int SUM { set; get; }
        }

        private void button_CRInquire_Click(object sender, EventArgs e)
        {
            if (dateTimePicker_CRCStart.Value > dateTimePicker_CRCEnd.Value || dateTimePicker_CRLStart.Value > dateTimePicker_CRLEnd.Value)
            {
                MessageBox.Show("日期設定錯誤");
            }                           

            else if (dateTimePicker_CRCEnd.Value > DateTime.Now || dateTimePicker_CRLEnd.Value > DateTime.Now)
            {
                MessageBox.Show("無法查詢未來的記錄");
            }            

            else
            {
                if (dateTimePicker_CRCStart.Value >= dateTimePicker_CRLStart.Value || dateTimePicker_CRCEnd.Value >= dateTimePicker_CRLStart.Value)
                {
                    MessageBox.Show("創角日期和登入日期有重疊，將導致留存率不準確");
                }

                int days = (dateTimePicker_CRLEnd.Value - dateTimePicker_CRLStart.Value).Days + 1;
                int CDays = (dateTimePicker_CRCEnd.Value - dateTimePicker_CRCStart.Value).Days + 1;

                try
                {
                    List<Data> myData = new List<Data>();

                    DataSet ds = LoginCount(comboBox_CRCountry.SelectedIndex, 0, dateTimePicker_CRLStart.Value, days, 1);
                    DataSet Cds = LoginCount(comboBox_CRCountry.SelectedIndex, 1, dateTimePicker_CRCStart.Value, CDays, 1);
                    string RemainDis = "0";
                    string[] RemainID = new string[0];
                    int Total = 0;

                    for (int i = 0; i < Cds.Tables[0].Rows.Count; i++)
                    {
                        bool alive = false;

                        for (int y = 0; y < ds.Tables[0].Rows.Count; y++)
                        {
                            if (ds.Tables[0].Rows[y]["UserID"].ToString() == Cds.Tables[0].Rows[i]["UserID"].ToString())
                            {
                                alive = true;
                            }
                        }                        

                        if (alive)
                        {
                            Array.Resize(ref RemainID, RemainID.Length + 1);

                           

                            RemainID[Total] += Cds.Tables[0].Rows[i]["UserID"].ToString();
                            Total++;
                        }
                    }

                    double ReMain = Total;
                    double TotalReMain = Cds.Tables[0].Rows.Count;

                    if (TotalReMain > 0)
                    {
                        RemainDis = (ReMain / TotalReMain).ToString("P");
                    }
                    else
                    {
                        RemainDis = "0.00%";
                    }
                    

                    #region 資料顯示

                    richTextBox_CRDis.Text = "伺服器 : " + comboBox_CRServer.SelectedItem.ToString() + "\n";
                    richTextBox_CRDis.Text += "創角日期 : " + dateTimePicker_CRCStart.Value.ToString("yyyy-MM-dd") + " 到 " + dateTimePicker_CRCEnd.Value.ToString("yyyy-MM-dd") + "\n";
                    richTextBox_CRDis.Text += "上線日期 : " + dateTimePicker_CRLStart.Value.ToString("yyyy-MM-dd") + " 到 " + dateTimePicker_CRLEnd.Value.ToString("yyyy-MM-dd") + "\n";

                    richTextBox_CRDis.Text += "\n" + "留存人數共 : " + Total + " 人" + "\n";
                    richTextBox_CRDis.Text += "創角人數共 : " + Cds.Tables[0].Rows.Count + " 人" + "\n";
                    richTextBox_CRDis.Text += "\n" + "留存率 : " + RemainDis + "\n";             

                    if (checkBox_CRShowID.Checked)
                    {
                        richTextBox_CRDis.Text += "\n" + "創角後登入玩家ID : " + "\n";

                        for (int i = 0; i < RemainID.Length; i++)
                        {
                            richTextBox_CRDis.Text += RemainID[i] + "\n";
                        }
                    }

                    if (checkBox_CRNewID.Checked)
                    {
                        richTextBox_CRDis.Text += "\n" + "新創角玩家ID : " + "\n";

                        for (int i = 0; i < Cds.Tables[0].Rows.Count; i++)
                        {
                            richTextBox_CRDis.Text += Cds.Tables[0].Rows[i]["UserID"].ToString() + "\n";
                        }
                    }

                    #endregion

                    MessageBox.Show("資料來了");
                }
                catch
                {
                    MessageBox.Show("查詢失敗");
                }
            }
        }

        private void CreateData(string Dialog)
        {
            int days = (dateTimePicker_PDEnd.Value - dateTimePicker_PDStart.Value).Days + 1;
            DateTime StartDate = dateTimePicker_PDStart.Value;

            #region 表格準備

            // 設定儲存檔名，不用設定副檔名，系統自動判斷 excel 版本，產生 .xls 或 .xlsx 副檔名
            string pathFile = Dialog;            

            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            //Excel.Range wRange;

            // 開啟一個新的應用程式
            excelApp = new Excel.Application();
                       
            // 停用警告訊息
            excelApp.DisplayAlerts = false;

            // 加入新的活頁簿

            excelApp.Workbooks.Add(Type.Missing);

            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];

            // 設定活頁簿焦點
            wBook.Activate();

            wBook.Sheets[3].Delete();
            wBook.Sheets[2].Delete();

            #endregion            

            try
            {                
                // 引用第一個工作表
                wSheet = (Excel._Worksheet)wBook.Worksheets[1];
                // 命名工作表的名稱
                wSheet.Name = StartDate.ToString("yyyy");
                // 設定工作表焦點
                wSheet.Activate();

                #region 取得每日不重複登入人數

                string TsqlA = "SELECT ServerID, UserID FROM Login_" + dateTimePicker_PDStart.Value.ToString("yyyyMMdd") + " WHERE (Kind = '15') GROUP BY ServerID, UserID ORDER BY ServerID;";                

                for (int x = 0; x < days - 1; x++)
                {
                    TsqlA += "SELECT ServerID, UserID FROM Login_" + dateTimePicker_PDStart.Value.AddDays(x + 1).ToString("yyyyMMdd") + " WHERE (Kind = '15') GROUP BY ServerID, UserID ORDER BY ServerID;";
                }

                DataSet ds = GetDate(TsqlA);                               

                for (int t = 0; t < ds.Tables.Count; t++)
                {
                    int S1 = 0, S2 = 0, S3 = 0, S4 = 0, S5 = 0, S6 = 0, S7 = 0, S8 = 0, S9 = 0, S10 = 0;

                    for (int i = 0; i < ds.Tables[t].Rows.Count; i++)
                    {
                        switch (ds.Tables[t].Rows[i]["ServerID"].ToString())
                        {
                            case "1":
                                S1++;
                                break;
                            case "2":
                                S2++;
                                break;
                            case "3":
                                S3++;
                                break;
                            case "4":
                                S4++;
                                break;
                            case "5":
                                S5++;
                                break;
                            case "6":
                                S6++;
                                break;
                            case "7":
                                S7++;
                                break;
                            case "8":
                                S8++;
                                break;
                            case "9":
                                S9++;
                                break;
                            case "10":
                                S10++;
                                break;
                        }
                    }                   

                    excelApp.Cells[1, t * 3 + 2] = StartDate.AddDays(t).ToString("MM 月 dd 日");
                    excelApp.Cells[2, t * 3 + 2] = "不重複登入人數";

                    excelApp.Cells[3, t * 3 + 2] = "Server";
                    excelApp.Cells[3, t * 3 + 3] = "人數";

                    excelApp.Cells[4, t * 3 + 2] = "Server 1";
                    excelApp.Cells[4, t * 3 + 3] = S1.ToString();
                    excelApp.Cells[5, t * 3 + 2] = "Server 2";
                    excelApp.Cells[5, t * 3 + 3] = S2.ToString();
                    excelApp.Cells[6, t * 3 + 2] = "Server 3";
                    excelApp.Cells[6, t * 3 + 3] = S3.ToString();
                    excelApp.Cells[7, t * 3 + 2] = "Server 4";
                    excelApp.Cells[7, t * 3 + 3] = S4.ToString();
                    excelApp.Cells[8, t * 3 + 2] = "Server 5";
                    excelApp.Cells[8, t * 3 + 3] = S5.ToString();
                    excelApp.Cells[9, t * 3 + 2] = "Server 6";
                    excelApp.Cells[9, t * 3 + 3] = S6.ToString();
                    excelApp.Cells[10, t * 3 + 2] = "Server 7";
                    excelApp.Cells[10, t * 3 + 3] = S7.ToString();
                    excelApp.Cells[11, t * 3 + 2] = "Server 8";
                    excelApp.Cells[11, t * 3 + 3] = S8.ToString();
                    excelApp.Cells[12, t * 3 + 2] = "Server 9";
                    excelApp.Cells[12, t * 3 + 3] = S9.ToString();
                    excelApp.Cells[13, t * 3 + 2] = "Server 10";
                    excelApp.Cells[13, t * 3 + 3] = S10.ToString();

                    // 設定總和公式 =SUM(B2:B4)
                    excelApp.Cells[14, t * 3 + 2] = "總共 : ";
                    excelApp.Cells[14, t * 3 + 3] = (S1 + S2 + S3 + S4 + S5 + S6 + S7 + S8 + S9 + S10).ToString();
                    
                    #region 表格美化設定
                    // 設定第1列顏色
                    //wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 2]];
                    //wRange.Select();
                    //wRange.Font.Color = ColorTranslator.ToOle(Color.White);
                    //wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);

                    //// 設定第5列顏色
                    //wRange = wSheet.Range[wSheet.Cells[5, 1], wSheet.Cells[5, 2]];
                    //wRange.Select();
                    //wRange.Font.Color = ColorTranslator.ToOle(Color.Red);
                    //wRange.Interior.Color = ColorTranslator.ToOle(Color.Yellow);

                    //// 自動調整欄寬
                    //wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[5, 2]];
                    //wRange.Select();
                    //wRange.Columns.AutoFit();                    
                    #endregion                    
                }

                #endregion

                #region 取得每日不重複儲值人數

                string TsqlB = "SELECT UserID, SUM(CAST(V2 AS int)) AS Money FROM Money_" + dateTimePicker_PDStart.Value.ToString("yyyyMMdd") + " WHERE (Kind = 1) AND (Cause = 10 OR Cause = 13 OR Cause = 14 OR Cause = 15 OR Cause = 22) GROUP BY UserID ORDER BY Money DESC;";
                string TsqlD = "SELECT TotalTWD FROM MonthMoney WHERE (Area = 'TW') AND (SaveDate = '2014-12-23');";


                for (int x = 0; x < days - 1; x++)
                {
                    TsqlB += "SELECT UserID, SUM(CAST(V2 AS int)) AS Money FROM Money_" + dateTimePicker_PDStart.Value.AddDays(x + 1).ToString("yyyyMMdd") + " WHERE (Kind = 1) AND (Cause = 10 OR Cause = 13 OR Cause = 14 OR Cause = 15 OR Cause = 22) GROUP BY UserID ORDER BY Money DESC;";
                    
                }

                ds = GetDate(TsqlB);

                for (int t = 0; t < ds.Tables.Count; t++)
                {
                    float DiamonTotal = 0;
                    string Charger = "";

                    excelApp.Cells[16, t * 3 + 2] = "不重複儲值人數";                    

                    for (int i = 0; i < ds.Tables[t].Rows.Count; i++)
                    {
                        DiamonTotal += Convert.ToInt32(ds.Tables[t].Rows[i]["Money"]);
                        Charger += ds.Tables[t].Rows[i]["UserID"].ToString() + "  儲值 : " + ds.Tables[t].Rows[i]["Money"].ToString() + "\r\n";
                    }

                    excelApp.Cells[17, t * 3 + 2] = "儲鑽數 : ";
                    excelApp.Cells[17, t * 3 + 3] = DiamonTotal.ToString() + "鑽 / " + Math.Round(DiamonTotal / 1.3) + " 元";
                    excelApp.Cells[18, t * 3 + 2] = "儲值人數 : ";
                    excelApp.Cells[18, t * 3 + 3] = ds.Tables[t].Rows.Count.ToString();

                    if (checkBox_PDShowCharger.Checked)
                    {
                        excelApp.Cells[19, t * 3 + 2] = "儲值細項 : ";
                        excelApp.Cells[19, t * 3 + 3] = Charger; 
                    }
                    
                }

                #endregion                

                #region 每日耗鑽項目

                string TsqlC = "SELECT Cause, SUM(CAST(V2 AS int)) AS Money FROM Money_" + dateTimePicker_PDStart.Value.ToString("yyyyMMdd") + " WHERE (Kind = 2) GROUP BY Cause ORDER BY Cause;";

                for (int x = 0; x < days - 1; x++)
                {
                    TsqlC += "SELECT Cause, SUM(CAST(V2 AS int)) AS Money FROM Money_" + dateTimePicker_PDStart.Value.AddDays(x + 1).ToString("yyyyMMdd") + " WHERE (Kind = 2) GROUP BY Cause ORDER BY Cause;";
                }

                ds = GetDate(TsqlC);

                for (int t = 0; t < ds.Tables.Count; t++)
                {
                    excelApp.Cells[21, t * 3 + 2] = "每日耗鑽項目";

                    for (int i = 0; i < ds.Tables[t].Rows.Count; i++)
                    {
                        string Cause = "其他";

                        switch (ds.Tables[t].Rows[i]["Cause"].ToString())
                        {
                            case "1":
                                Cause = "購買能量";
                                break;
                            case "2":
                                Cause = "抽武將";
                                break;
                            case "3":
                                Cause = "開寶箱";
                                break;
                            case "4":
                                Cause = "消除競技場ＣＤ";
                                break;
                            case "5":
                                Cause = "重置攻城戰";
                                break;
                            case "6":
                                Cause = "重置攻城寶箱";
                                break;
                            case "7":
                                Cause = "購買懸賞格子";
                                break;
                            case "8":
                                Cause = "轉蛋扣除";
                                break;
                            case "9":
                                Cause = "購買武將欄位";
                                break;
                            case "10":
                                Cause = "攻城戰復活";
                                break;
                            case "11":
                                Cause = "戰場使用精英援軍";
                                break;
                            case "12":
                                Cause = "快速通關";
                                break;
                            case "13":
                                Cause = "購買ＰＶＰ挑戰次數";
                                break;
                            case "14":
                                Cause = "消世界ＢｏｓｓＣＤ";
                                break;
                            case "15":
                                Cause = "消連續掃蕩ＣＤ";
                                break;
                            case "16":
                                Cause = "拉ＢＡＲ轉盤重置";
                                break;
                            case "17":
                                Cause = "購買武器欄位";
                                break;
                            case "18":
                                Cause = "Ｖｉｐ強化武將耗鑽";
                                break;
                            case "19":
                                Cause = "消跨服ＰＶＰ　ＣＤ";
                                break;
                            case "20":
                                Cause = "買跨服ＰＶＰ挑戰次數";
                                break;
                            case "21":
                                Cause = "建立軍團";
                                break;
                            case "22":
                                Cause = "軍團捐獻";
                                break;
                            case "23":
                                Cause = "接關扣鑽";
                                break;
                            case "24":
                                Cause = "購買累積儲值福袋";
                                break;
                            case "25":
                                Cause = "購買勢力禮盒";
                                break;
                            case "26":
                                Cause = "買全球Ｂｏｓｓ挑戰次數";
                                break;
                            case "27":
                                Cause = "全球Ｂｏｓｓ升級";
                                break;
                            case "28":
                                Cause = "ＷＥＢ後台";
                                break;
                            case "29":
                                Cause = "軍團建設";
                                break;
                            case "30":
                                Cause = "軍團商店購買";
                                break;
                            case "31":
                                Cause = "軍團商店購買失敗回補";
                                break;
                            case "32":
                                Cause = "武將解放";
                                break;
                            case "33":
                                Cause = "50萬金幣轉90鑽10連抽";
                                break;
                            case "34":
                                Cause = "每日簽到補簽";
                                break;
                        }

                        excelApp.Cells[22 + i, t * 3 + 2] = Cause;
                        excelApp.Cells[22 + i, t * 3 + 3] = ds.Tables[t].Rows[i]["Money"].ToString();
                    }
                }

                #endregion

                excelApp.Columns.AutoFit();
                excelApp.Columns.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;

                try
                {
                    //另存活頁簿
                    wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MessageBox.Show("產生成功! 儲存文件於 " + Environment.NewLine + pathFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("產生報表時出錯！" + Environment.NewLine + ex.Message);
            }

            #region 其他
            // 讓Excel文件可見
            excelApp.Visible = true;

            //關閉活頁簿
            //wBook.Close(false, Type.Missing, Type.Missing);

            //關閉Excel
            //excelApp.Quit();

            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            //wRange = null;
            excelApp = null;
            GC.Collect();
            Console.Read();
            #endregion

        }

        private void button_PDPrint_Click(object sender, EventArgs e)
        {
            DateTime StartDate = dateTimePicker_PDStart.Value;
            DateTime EndDate = dateTimePicker_PDEnd.Value;

            if (StartDate > EndDate)
            {
                MessageBox.Show("日期設定錯誤");
            }
            else if ((dateTimePicker_PDEnd.Value - dateTimePicker_PDStart.Value).Days > 40)
            {
                MessageBox.Show("查詢範圍過大");
            }
            else
            {
                SaveFileDialog dialog = new SaveFileDialog();

                dialog.CreatePrompt = false;
                dialog.OverwritePrompt = true;

                dialog.Filter = "All Files (*.*)|*.*";
                dialog.Title = "報表存檔位置";
                dialog.AddExtension = true;
                dialog.FileName = DateTime.Now.ToString("yyyyMMdd") + "分析資料報表";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string savepos = dialog.FileName;
                    CreateData(savepos);
                }
            }
        }

        private DataSet GetDate(string Tsql)
        {
            DataSet ds = new DataSet();

            try
            {
                switch (comboBox_PDCountry.SelectedIndex)
                {
                    case 0:
                        ds = TWsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        break;
                    case 1:
                        ds = HKsql.Get_SQL_Recoarder_DataSet(GSecurity.Encrypt(Tsql), SqlPassWord);
                        break;
                }

                return ds;

            }
            catch
            {
                return null;
            }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }


        private void button_CMInquire_Click(object sender, EventArgs e)
        {
            #region 禁止修改
            comboBox_CMCountry.Enabled = false;
            comboBox_CMServer.Enabled = false;
            dateTimePicker_CMCStart.Enabled = false;
            dateTimePicker_CMCEnd.Enabled = false;
            dateTimePicker_CMLStart.Enabled = false;
            dateTimePicker_CMLEnd.Enabled = false;
            button_CMInquire.Enabled = false;
            button_CMClean.Enabled = false;
            #endregion

            #region 初始化
            progressBar1.Maximum = 100;
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            CMCountry = comboBox_CMCountry.SelectedIndex;
            CMServer = comboBox_CMServer.SelectedIndex;
            DateTimePickerCMLStart = dateTimePicker_CMLStart.Value;
            DateTimePickerCMCStart = dateTimePicker_CMCStart.Value;
            CdsDataTable.Clear();
            dssDataTable.Clear();
            SumdssDataTable.Clear();
            richTextBox_CMDis.Text = "";
            #endregion

            #region 變數

            days = (dateTimePicker_CMLEnd.Value - dateTimePicker_CMLStart.Value).Days + 1;
            CDays = (dateTimePicker_CMCEnd.Value - dateTimePicker_CMCStart.Value).Days + 1;

            CMLStartTime = dateTimePicker_CMLStart.Value;
            CMCStartTime = dateTimePicker_CMCStart.Value;

            #endregion

            #region 運算

            if (dateTimePicker_CMCStart.Value > dateTimePicker_CMCEnd.Value || dateTimePicker_CMLStart.Value > dateTimePicker_CMLEnd.Value)
            {
                MessageBox.Show("日期設定錯誤");
            }
            else if (dateTimePicker_CMCEnd.Value > DateTime.Now || dateTimePicker_CMLEnd.Value > DateTime.Now)
            {
                MessageBox.Show("無法查詢未來的記錄");
            }
            else if (dateTimePicker_CMCStart.Value >= dateTimePicker_CMLStart.Value || dateTimePicker_CMCEnd.Value >= dateTimePicker_CMLStart.Value)
            {
                MessageBox.Show("創角日期和儲值日期有重疊，將導致留存率不準確");
            }
            else if (comboBox_CMCountry.SelectedIndex == 2)
            {
                MessageBox.Show("中國目前沒開放查詢");
            }
            else
            {
                try
                {
                    backgroundWorker2.RunWorkerAsync();
                }
                catch
                {
                    MessageBox.Show("查詢失敗");
                }
            }
            #endregion
          
        }
        
        private void button_CMClean_Click(object sender, EventArgs e)
        {
            richTextBox_CMDis.Text = "";
        }

        #region 角色花錢

        private void startJob(BackgroundWorker myWork)
        {
            for (int i = 0; i <= 80; i++)
            {
                myWork.ReportProgress(i);
            }
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            this.backgroundWorker2.WorkerReportsProgress = true;
            #region 傳送花費與創建

            while (CMLStartTime < dateTimePicker_CMLStart.Value.AddDays(days))
            {
                DataSet ds = new DataSet();
                ds = LoginCountMony(CMCountry, 0, CMLStartTime, 1, CMServer);
                for (int y = 0; y < ds.Tables[0].Rows.Count; y++)
                {
                    DataRow myDataRow = dssDataTable.NewRow();
                    myDataRow["UserID"] = ds.Tables[0].Rows[y]["UserID"].ToString().Trim();
                    myDataRow["_Money"] = ds.Tables[0].Rows[y]["_Money"].ToString().Trim();
                    dssDataTable.Rows.Add(myDataRow);
                    startJob(sender as BackgroundWorker);
                }
                CMLStartTime = CMLStartTime.AddDays(1);
        }        

            while (CMCStartTime < dateTimePicker_CMCStart.Value.AddDays(CDays))
            {
                DataSet Cds = new DataSet();
                Cds = LoginCountMony(CMCountry, 1, CMCStartTime, 1, CMServer);
                for (int i = 0; i < Cds.Tables[0].Rows.Count; i++)
                {
                    DataRow myDataRow = CdsDataTable.NewRow();
                    myDataRow["UserID"] = Cds.Tables[0].Rows[i]["UserID"].ToString().Trim();
                    CdsDataTable.Rows.Add(myDataRow);
                    startJob(sender as BackgroundWorker);
                }
                CMCStartTime = CMCStartTime.AddDays(1);
            }
            #endregion   
        }


        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            #region 變數
            int UserID = 0;
            int _Money = 0;
            List<string> UserId = new List<string>();
            List<string> UserIdCheck = new List<string>();
            #endregion

            #region 不重複
            for (int y = 0; y < dssDataTable.Rows.Count; y++)
            {
                UserId.Add(dssDataTable.Rows[y]["UserID"].ToString().Trim());
            }

            for (int y = 0; y < UserId.Count; y++)
            {
                bool UserTalk = false;

                if (y == 0)
                {
                    UserIdCheck.Add(UserId[y].ToString().Trim());
                }

                for (int i = 0; i < UserIdCheck.Count; i++)
                {
                    if (UserIdCheck[i].ToString().Trim() == UserId[y].ToString().Trim())
                    {
                        UserTalk = true;
                    }
                }

                if (UserTalk == false)
                {
                    UserIdCheck.Add(UserId[y].ToString().Trim());
                }

            }
            #endregion

            #region 資料加總

            for (int y = 0; y < UserIdCheck.Count; y++)
            {
                string UserIDALL = UserIdCheck[y].ToString().Trim();
                int _MoneyALL = 0;

                for (int i = 0; i < dssDataTable.Rows.Count; i++)
                {
                    string UserIDALLCheck = dssDataTable.Rows[i]["UserID"].ToString().Trim();

                    if (UserIDALL == UserIDALLCheck)
                    {
                        UserIDALL = dssDataTable.Rows[i]["UserID"].ToString().Trim();
                        _MoneyALL = _MoneyALL + int.Parse(dssDataTable.Rows[i]["_Money"].ToString().Trim());
                    }
                }

                DataRow myDataRow = SumdssDataTable.NewRow();
                myDataRow["UserID"] = UserIDALL;
                myDataRow["_Money"] = _MoneyALL;
                SumdssDataTable.Rows.Add(myDataRow);
            }

            #endregion

            #region 顯示資料
            this.progressBar1.Value = 96;
            for (int y = 0; y < SumdssDataTable.Rows.Count; y++)
            {
                for (int i = 0; i < CdsDataTable.Rows.Count; i++)
                {
                    if (SumdssDataTable.Rows[y]["UserID"].ToString().Trim() == CdsDataTable.Rows[i]["UserID"].ToString())
                    {
                        richTextBox_CMDis.Text += " 玩家 : " + SumdssDataTable.Rows[y]["UserID"].ToString() + " 花費 : " + SumdssDataTable.Rows[y]["_Money"].ToString() + "\n";
                        UserID++;
                        _Money = _Money + int.Parse(SumdssDataTable.Rows[y]["_Money"].ToString().Trim());
                    }
                }
            }
            richTextBox_CMDis.Text += " 總玩家人數 : " + UserID.ToString() + " 總玩家花費 : " + _Money.ToString() + "\n";

            this.backgroundWorker2.WorkerReportsProgress = false;
            this.progressBar1.Value = 100;
            MessageBox.Show("資料來了");
            #endregion

            #region 開放修改
            comboBox_CMCountry.Enabled = true;
            comboBox_CMServer.Enabled = true;
            dateTimePicker_CMCStart.Enabled = true;
            dateTimePicker_CMCEnd.Enabled = true;
            dateTimePicker_CMLStart.Enabled = true;
            dateTimePicker_CMLEnd.Enabled = true;
            button_CMInquire.Enabled = true;
            button_CMClean.Enabled = true;
            #endregion
        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar1.Value = e.ProgressPercentage;
        }

        #endregion
    }
       

    public partial class uTextBox : TextBox
    {
        uint _maxByteLength = 0;
        public uint MaxByteLength
        {
            get { return _maxByteLength; }
            set { _maxByteLength = value; }
        }


        protected override void OnKeyPress(KeyPressEventArgs e)
        {
            base.OnKeyPress(e);

            if (ReadOnly) return; //唯讀不處理
            if (_maxByteLength == 0) return; //沒設定MaxByteLength不處理
            if (char.IsControl(e.KeyChar)) return; //Backspace, Enter...等控制鍵不處理

            int textByteLength = Encoding.GetEncoding(950).GetByteCount(Text + e.KeyChar.ToString()); //取得原本字串和新字串相加後的Byte長度
            int selectTextByteLength = Encoding.GetEncoding(950).GetByteCount(SelectedText); //取得選取字串的Byte長度, 選取字串將會被取代
            if (textByteLength - selectTextByteLength > _maxByteLength) e.Handled = true; //相減後長度若大於設定值, 則不送出該字元
        }

        int WM_PASTEDATA = 0x0302; //貼上資料的訊息

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_PASTEDATA) //如果收到貼上資料的訊息, 包括Ctrl+V, Shift+Ins和滑鼠右鍵選單中的貼上
                SendCharFromClipboard(); //就把剪貼簿中的字串一個字元一個字元的拆開, 再傳給自己
            else
                base.WndProc(ref m);
        }

        int WM_CHAR = 0x0102;

        private void SendCharFromClipboard()
        {
            foreach (char c in Clipboard.GetText())
            {
                Message msg = new Message();
                msg.HWnd = Handle;
                msg.Msg = WM_CHAR;
                msg.WParam = (IntPtr)c;
                msg.LParam = IntPtr.Zero;
                base.WndProc(ref msg);
            }
        }
    }



}

