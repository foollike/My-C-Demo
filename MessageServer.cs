using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.IO;

namespace MessageServer
{
    public partial class Form1 : Form
    {        
        //版本資料區

        int Ver = 1; //活動版號

        DateTime LimitDate = Convert.ToDateTime("2014-12-25 8:00"); //截止日期(格式別亂改)，要比真正結束日期多一天

        int CountryNum = 3; //目前運作活動的有幾個國家

        int Times = 0; //資訊欄保留顯示筆數

        /*-----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
        /*-----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/

        #region WebServer

        SynchronizationContext _syncContext;

        Server_Debug.Charger CPDebug = new Server_Debug.Charger();
        Server_TW.Charger CPTW = new Server_TW.Charger();
        Server_HK.Charger CPHK = new Server_HK.Charger();

        SqlLink_TW.SQLLink SQLTW = new SqlLink_TW.SQLLink();
        SqlLink_HK.SQLLink SQLHK = new SqlLink_HK.SQLLink();

        string ServerPassWord = "Macintosh@1984#"; //儲值伺服器密碼
        string SqlPassWord = "SQL@#1955"; //Sql密碼

        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label_Ver.Text = "版本 : " + Ver;
            label_Country.Text = "運作國家數目 : " + CountryNum;

            _syncContext = SynchronizationContext.Current;

            Thread Thread_A = new Thread(new ThreadStart(WorkA));
            Thread_A.IsBackground = true;
            Thread_A.Start();
        }

        public void WorkA()
        {
            //1小時檢查一次
            while (true)
            {
                DateTime NowDate = DateTime.Now;

                if (NowDate <= LimitDate)
                {
                    if (NowDate.Hour == 1)
                    {
                        CheckVip(NowDate);
                    }
                }

                Thread.Sleep(3600000);
            }
        }

        private void CheckVip(DateTime NowDate) //確認3個幸運兒並送出禮物
        {
            string TsqlA = "SELECT Ver, UserID, NickName, SaveDate FROM XmasMessage WHERE (Ver = '" + Ver.ToString() + "') AND (DATEPART(yyyy, SaveDate) = '" + DateTime.Now.AddDays(-1).ToString("yyyy") + "') AND (DATEPART(MM, SaveDate) = '" + DateTime.Now.AddDays(-1).ToString("MM") + "') AND (DATEPART(dd, SaveDate) = '" + DateTime.Now.AddDays(-1).ToString("dd") + "');";
            TsqlA += "SELECT Ver, UserID, NickName, CheckDate, SaveDate FROM XmasVip WHERE (Ver = '" + Ver.ToString() + "');";
            TsqlA += "SELECT Ver, UserID, NickName, CheckDate, SaveDate FROM XmasVip WHERE (Ver = '" + Ver.ToString() + "') AND (DATEPART(yyyy, SaveDate) = '" + DateTime.Now.ToString("yyyy") + "') AND (DATEPART(MM, SaveDate) = '" + DateTime.Now.ToString("MM") + "') AND (DATEPART(dd, SaveDate) = '" + DateTime.Now.ToString("dd") + "')";

            DataSet ds = new DataSet();

            for (int i = 0; i < CountryNum; i++)
            {
                try
                {
                    #region 讀表

                    switch (i)
                    {                       
                        case 1:
                            ds = SQLTW.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(TsqlA), SqlPassWord);
                            break;
                        case 2:
                            ds = SQLHK.Get_SQL_PuzzleWeb_DataSet(GSecurity.Encrypt(TsqlA), SqlPassWord);
                            break;
                        default:
                            ds = CPDebug.Get_SQL_DataSet(GSecurity.Encrypt(TsqlA), ServerPassWord);
                            break;
                    }

                    #endregion

                    if (ds.Tables[2].Rows.Count == 0)
                    {
                        if (ds.Tables[0].Rows.Count >= 3)
                        {
                            if (ds.Tables[1].Rows.Count > 0)
                            {
                                //可使用名單
                                int TotalU = 0;
                                string[] NIckName = new string[0];
                                string[] MainID = new string[0];

                                //已重複名單
                                int TotalNU = 0;
                                string[] BNIckName = new string[0];
                                string[] BMainID = new string[0];

                                #region 篩選不重複玩家名單

                                for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                                {
                                    bool Alive = false;

                                    for (int k = 0; k < ds.Tables[1].Rows.Count; k++)
                                    {
                                        if (ds.Tables[0].Rows[j]["UserID"].ToString() == ds.Tables[1].Rows[k]["UserID"].ToString())
                                        {
                                            Alive = true;
                                        }
                                    }

                                    if (!Alive)
                                    {
                                        Array.Resize(ref NIckName, NIckName.Length + 1);
                                        Array.Resize(ref MainID, MainID.Length + 1);

                                        NIckName[TotalU] = ds.Tables[0].Rows[j]["NickName"].ToString();
                                        MainID[TotalU] = ds.Tables[0].Rows[j]["UserID"].ToString();

                                        TotalU++;
                                    }
                                    else
                                    {
                                        Array.Resize(ref BNIckName, BNIckName.Length + 1);
                                        Array.Resize(ref BMainID, BMainID.Length + 1);

                                        BNIckName[TotalNU] = ds.Tables[0].Rows[j]["NickName"].ToString();
                                        BMainID[TotalNU] = ds.Tables[0].Rows[j]["UserID"].ToString();

                                        TotalNU++;
                                    }
                                }
                                #endregion

                                if (TotalU >= 3)
                                {
                                    int[] RandomPick = GetRandomNum(TotalU, 3);

                                    for (int j = 0; j < RandomPick.Length; j++)
                                    {
                                        FinishVipWork(i, Convert.ToInt32(MainID[RandomPick[j]]), NIckName[RandomPick[j]].ToString(), NowDate);
                                    }
                                }
                                else
                                {
                                    if (TotalU > 0)
                                    {
                                        if (TotalU > 1)
                                        {
                                            int[] RandomPick = GetRandomNum(TotalU, TotalU);

                                            for (int j = 0; j < RandomPick.Length; j++)
                                            {
                                                FinishVipWork(i, Convert.ToInt32(MainID[RandomPick[j]]), NIckName[RandomPick[j]].ToString(), NowDate);
                                            }
                                        }
                                        else
                                        {
                                            FinishVipWork(i, Convert.ToInt32(MainID[0]), NIckName[0].ToString(), NowDate);
                                        }
                                    }

                                    int[] NotUse = GetRandomNum(TotalNU, 3 - TotalU);

                                    for (int j = 0; j < NotUse.Length; j++)
                                    {
                                        FinishVipWork(i, Convert.ToInt32(BMainID[NotUse[j]]), BNIckName[NotUse[j]].ToString(), NowDate);
                                    }
                                }
                            }
                            else
                            {
                                if (ds.Tables[0].Rows.Count > 3)
                                {
                                    int[] RandomPick = GetRandomNum(ds.Tables[0].Rows.Count, 3);

                                    for (int j = 0; j < RandomPick.Length; j++)
                                    {
                                        FinishVipWork(i, Convert.ToInt32(ds.Tables[0].Rows[RandomPick[j]]["UserID"]), ds.Tables[0].Rows[RandomPick[j]]["NickName"].ToString(), NowDate);
                                    }
                                }
                                else
                                {
                                    for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                                    {
                                        FinishVipWork(i, Convert.ToInt32(ds.Tables[0].Rows[j]["UserID"]), ds.Tables[0].Rows[j]["NickName"].ToString(), NowDate);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                                {
                                    FinishVipWork(i, Convert.ToInt32(ds.Tables[0].Rows[j]["UserID"]), ds.Tables[0].Rows[j]["NickName"].ToString(), NowDate);
                                }
                            }
                            else
                            {
                                ShowMessage(i, DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 無任何人留言");
                            }
                        }
                    }
                    else
                    {
                        ShowMessage(i, "已選出幸運兒");   
                    }
                }
                catch
                {
                    ShowMessage(i, "資料寫入錯誤");                    
                }
            }

        }

        private void FinishVipWork(int CountryIndex, int TargetUser, string NickName, DateTime NowDate) //確認VIP後發送禮物
        {
            int DataCreatResult = RecordVip(0, CountryIndex, TargetUser, NickName, NowDate);

            if (DataCreatResult > 0)
            {
                string SendGiftResult = SendGift(CountryIndex, DateTime.Now.AddDays(-1).ToString("MMdd"), TargetUser);

                if (SendGiftResult == "OK,發送成功")
                {
                    int DataFinishResult = RecordVip(1, CountryIndex, TargetUser, NickName, NowDate);

                    if (DataFinishResult > 0)
                    {
                        ShowMessage(CountryIndex, "發送給 " + TargetUser + " 的 " + DateTime.Now.ToString("MMdd") + " 禮物成功");
                    }
                    else
                    {
                        ShowMessage(CountryIndex, "發送給 " + TargetUser + " 的 " + DateTime.Now.ToString("MMdd") + " 禮物成功，但SQL記錄失敗");
                    }
                }
                else
                {
                    if (SendGiftResult == "")
                    {
                        ShowMessage(CountryIndex, "發送給 " + TargetUser + " 的 " + DateTime.Now.ToString("MMdd") + " 禮物失敗");   
                    }
                    else
                    {
                        ShowMessage(CountryIndex,SendGiftResult);                        
                    }
                }
            }
            else
            {
                ShowMessage(CountryIndex, "資料寫入錯誤");                
            }
        }

        private int RecordVip(int Type, int CountryIndex, int TargetUser, string NickName, DateTime NowDate) //寫入當日幸運兒 Type: 0=寫入資料 1=將資料改成已送出禮物
        {
            try
            {
                #region SQL語法

                string Tsql = "";
                if (Type == 0)
                {
                    Tsql = "INSERT INTO XmasVip (Ver, UserID, NickName, CheckDate, SaveDate, IsCharge, IdentityCode) ";
                    Tsql += "VALUES ('" + Ver + "','" + TargetUser + "',N'" + NickName.Replace("'", "''") + "','" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','0','" + TargetUser.ToString() + NowDate.ToString("yyyyMMddHHmmss") + "')";
                }
                else
                {
                    Tsql = "UPDATE XmasVip SET IsCharge = '1' WHERE (Ver = '" + Ver + "') AND (UserID = '" + TargetUser + "') AND (IdentityCode = '" + TargetUser.ToString() + NowDate.ToString("yyyyMMddHHmmss") + "')";
                }

                #endregion

                int Result = 0;

                switch (CountryIndex)
                {                    
                    case 1:
                        Result = SQLTW.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                        break;
                    case 2:
                        Result = SQLHK.Set_SQL_PuzzleWeb_Command(GSecurity.Encrypt(Tsql), SqlPassWord);
                        break;
                    default:
                        Result = CPDebug.Set_SQL_Command(GSecurity.Encrypt(Tsql), ServerPassWord);
                        break;
                }

                return Result;
            }
            catch
            {
                return 0;
            }
        }

        private string SendGift(int CountryIndex, string GiftNum, int TargetUser) //單發送禮物Function
        {           
            try
            {
                int GiftType = 0; // 0=物品 1=武將

                #region 送物品
                //物品ID= 1.鑽石 2.金幣 3.銀幣 4.金牌 5.聖旨 6.小體力卷 7.中體力卷 8.大體力卷 9.白寶石 10.紅寶石 11.藍寶石 12.黃寶石 13.紫寶石 14.統御值 17.友情值
                int ItemID = 0;

                //物品數量
                int ItemCount = 0;
                #endregion

                #region 送武將
                int HeroID = 0; //武將ID
                int Wop = 0; //武將ID
                bool SpItem = false; //是否為閃卡
                int SHP = 0; //覺醒值HP
                int SATK = 0; //覺醒值ATK
                int SRHP = 0; //覺醒值RHP
                int ItemLV = 0; //武將等級
                #endregion

                //禮物單(單人)
                switch (GiftNum)
                {
                    case "1218":                        
                    case "1219":
                    case "1220":
                    case "1221":
                    case "1222":
                    case "1223":
                    case "1224":
                    default:
                        GiftType = 1;
                        HeroID = 259;
                        Wop = 0;
                        SpItem = false;
                        SHP = 0;
                        SATK = 0;
                        SRHP = 0;
                        ItemLV = 0;
                        break;
                }

                #region 發送

                string GiftResult = "";
                switch (CountryIndex)
                {
                    case 1:
                        int PushMsgTW = 982; //訊息編號

                        if (GiftType == 0)
                        {
                            GiftResult = CPTW.Recharge(ItemID, ItemCount, TargetUser, false, PushMsgTW, ServerPassWord);
                        }
                        else if (GiftType == 1)
                        {
                            GiftResult = CPTW.RechargeItem(HeroID, 1, TargetUser, PushMsgTW, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord);
                        }
                        break;

                    case 2:
                        int PushMsgHK = 730;

                        if (GiftType == 0)
                        {
                            GiftResult = CPHK.Recharge(ItemID, ItemCount, TargetUser, false, PushMsgHK, ServerPassWord);
                        }
                        else if (GiftType == 1)
                        {
                            GiftResult = CPHK.RechargeItem(HeroID, 1, TargetUser, PushMsgHK, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord);
                        }
                        break;

                    default:
                        int PushMsgDebug = 982;

                        if (GiftType == 0)
                        {
                            GiftResult = CPDebug.Recharge(ItemID, ItemCount, TargetUser, false, PushMsgDebug, ServerPassWord);
                        }
                        else if (GiftType == 1)
                        {
                            GiftResult = CPDebug.RechargeItem(HeroID, 1, TargetUser, PushMsgDebug, Wop, SpItem, SHP, SATK, SRHP, ItemLV, ServerPassWord);
                        }
                        break;
                }

                #endregion

                return GiftResult;
            }
            catch
            {
                return "";                
            }

        }

        private int[] GetRandomNum(int Max,int Count) //取3個不重複亂數
        {
            int Target = 0; //發送名單
            int[] RandomNo = new int[Count];

            while (Target < Count)
            {
                Random rnd = new Random();
                bool SameNum = false;

                RandomNo[Target] = rnd.Next(0, Max);

                switch (Target)
                {
                    case 0:
                        SameNum = true;
                        break;
                    case 1:
                        if (RandomNo[0] != RandomNo[1])
                        {
                            SameNum = true;
                        }
                        break;
                    case 2:
                        if (RandomNo[0] != RandomNo[1])
                        {
                            if (RandomNo[0] != RandomNo[2])
                            {
                                if (RandomNo[1] != RandomNo[2])
                                {
                                    SameNum = true;
                                }
                            }
                        }
                        break;
                }

                if (SameNum)
                {
                    Target++;
                }
            }

            return RandomNo;
                            
        }

        private void ShowMessage(int CountryIndex, string Text)
        {
            string CountryName = "Debug";

            switch (CountryIndex)
            {
                case 1:
                    CountryName = "TW ";
                    _syncContext.Post(Result, CountryName + Text);
                    break;
                case 2:
                    CountryName = "HK ";
                    _syncContext.Post(Result, CountryName + Text);
                    break;
                default:
                    CountryName = "Debug ";
                    _syncContext.Post(Result, CountryName + Text);
                    break;
            }
        } //顯示訊息

        private void Result(object result) //資訊顯示
        {
            DateTime NowDate = DateTime.Now; //現在時間

            try
            {
                if (Times <= 100)
                {
                    richTextBox_Display.Text += result.ToString() + "---" + NowDate.ToString("yyyy-MM-dd HH:mm:ss") + "\n";
                    Times++;
                }
                else
                {
                    richTextBox_Display.Text = "";
                    Times = 0;
                    richTextBox_Display.Text += result.ToString() + "---" + NowDate.ToString("yyyy-MM-dd HH:mm:ss") + "\n";
                    Times++;
                }
            }
            catch
            {
                richTextBox_Display.Text += "顯示記錄失敗" + "\n";
                Times++;
            }

        }

    }
}
