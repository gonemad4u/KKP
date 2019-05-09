using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using DAL;
using BLL;
using System.IO;
using System.Diagnostics;


namespace Mitsunori
{
    public partial class FmMNRKYDHS01 : Form
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger
        (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        DialogResult myResult;
        private static bool init = true;
        public FmMNRKYDHS01()
        {
            InitializeComponent();
            GR_LIST.AutoGenerateColumns = false;
            E_UNCHIN1.Value = DateTime.Now;
            E_UNCHIN1.CustomFormat = " ";
            E_UNCHIN2.Value = DateTime.Now;
            E_UNCHIN2.CustomFormat = " ";
            E_EXLSRD1.Value = DateTime.Now;
            E_EXLSRD1.CustomFormat = " ";
            E_EXLSRD2.Value = DateTime.Now;
            E_EXLSRD2.CustomFormat = " ";
            E_SYUKABI1.Value = DateTime.Now;
            E_SYUKABI1.CustomFormat = " ";
            E_SYUKABI2.Value = DateTime.Now;
            E_SYUKABI2.CustomFormat = " ";
            var Areanm = Tools.GetAreanm();
            CB_AREA.DataSource = Areanm;
            CB_AREA.ValueMember = "areanm";
            CB_AREA.DisplayMember = "areanm";            
        }

        //Clears all selections
        private void B_CLEAR_Click(object sender, EventArgs e)
        {
            Control grpValue = this.GR_MNRKYDHS01;
            for (int index = 0; index < grpValue.Controls.Count; index++)
            {
                switch (grpValue.Controls[index].GetType().Name)
                {
                    case "TextBox":
                        ((TextBox)grpValue.Controls[index]).Text = "";
                        break;
                    //case "RadioButton":
                    //    ((RadioButton)(grpValue.Controls[index])).Checked = false;
                    //    break;
                    case "CheckBox":
                        ((CheckBox)(grpValue.Controls[index])).Checked = false;
                        break;
                    case "ComboBox":
                        ((ComboBox)(grpValue.Controls[index])).Text = "";
                        break;
                    case "DateTimePicker":
                        ((DateTimePicker)(grpValue.Controls[index])).CustomFormat = " ";
                        ((DateTimePicker)(grpValue.Controls[index])).Checked = true;
                        break;
                }
            }
        }

        //Check the number of checked 
        private int GrdCheck()
        {
            int intCount = 0;
            string message = "一覧で該当データを選択してください。";

            for (int i = 0; i < GR_LIST.RowCount; i++)
            {
                if ((bool)GR_LIST.Rows[i].Cells[0].EditedFormattedValue == true)
                {
                    intCount++;
                }
            }
            if (intCount == 0)
            {
                MessageBox.Show(message, " Waring", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            return intCount;
        }

        //check if there are multiple status selected
        private bool StatusCheck(string[] statusList)
        {
            HashSet<string> ts = new HashSet<string>();
            List<string> status = statusList.ToList();
            string message = "ステータスをチェックしてください。";

            for (int i = 0; i < GR_LIST.RowCount; i++)
            {
                if ((bool)GR_LIST.Rows[i].Cells[0].EditedFormattedValue == true)
                {
                    ts.Add(GR_LIST.Rows[i].Cells["CODE1"].Value.ToString());
                }
            }
            foreach(string a in ts)
            {
                if (!status.Contains(a))
                {
                    MessageBox.Show(message, " Waring", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
            }
            return true;
        }

        //Attemp to delete the selected data shown in the gridview, physically in Oracle
        private void B_DELETE_Click(object sender, EventArgs e)
        {
            if (GrdCheck() != 0)
            {
                DialogResult myResult;
                string mess = GrdCheck() + "件のデータを削除します。よろしいですか？";
                myResult = MessageBox.Show(mess, "Delete Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (myResult == DialogResult.Yes)
                {
                    Delete delete = new Delete();
                    delete.DeleteFromGL(GR_LIST);
                    Search();
                }
                else
                {
                    MessageBox.Show("処理を中止しました。");
                }
            }
        }

        /*
         * Read a folder that is selected by user, which is filled with csv files, and upload the data onto
         * Oracle
        */
        string folder = @"C:\";
        private void B_READ_Click(object sender, EventArgs e)
        {

            Load L = new Load();

            using (var fbd = new FolderBrowserDialog())
            {
                
                fbd.SelectedPath = folder;
                DialogResult result = fbd.ShowDialog();
                
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    folder = fbd.SelectedPath;
                    DirectoryInfo d = new DirectoryInfo(fbd.SelectedPath);
                    FileInfo[] Files = d.GetFiles("*.csv");
                    log.Info("EXEC LOAD BEGIN");
                    try
                    {
                        if (Files.Length == 0) MessageBox.Show("指定されたフォルダーにCSVファイルが見つからない");
                        foreach (FileInfo f in Files)
                        {
                            L.csvToDT(f);
                        }

                        L.insertData();

                        GR_LIST.DataSource = L.getData();
                    }
                    catch(Exception err)
                    {
                        log.Error("Unexpected Error: " + err.Message);
                        myResult = MessageBox.Show("Unexpected Error, please check csv file! See log?", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                        if (myResult == DialogResult.OK)
                        {
                            Process.Start("explorer.exe", @"logs");
                        }
                    }
                    finally
                    {
                        log.Info("EXEC END");
                    }
                }
                else
                {
                    MessageBox.Show("処理を中止しました。");
                }
            }
        }

        //Read Syukahyo info from oracle and output an excel file
        private void B_SYUKAHYO_Click(object sender, EventArgs e)
        {
            if (GrdCheck() != 0 && StatusCheck(new string[] { "未処理", "集荷表", "集約" }) == true)
            {
                DataGridView gr = GR_LIST;
                bool ZanKa = CH_ZANKA.Checked;
                if (ZanKa)
                {
                    string mess = "選択したレコードは全て残貨対象のため処理できません。";
                    MessageBox.Show(mess, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                {
                    SyuKaHyo syukahyo = new SyuKaHyo();
                    syukahyo.SearchParameter(gr);
                    Search();
                }
            }
        }

        //Read Okurijyo info from oracle and output an excel file
        private void B_OKURIJYO_Click(object sender, EventArgs e)
        {
            if (GrdCheck() != 0 && StatusCheck(new string[] { "集荷表", "集約", "配送済", "運賃計算", "実績送信済" }) == true)
            {
                DataGridView gr = GR_LIST;

                bool ZanKa = CH_ZANKA.Checked;
                if (ZanKa)
                {
                    string mess = "選択したレコードは全て残貨対象のため処理できません。";
                    MessageBox.Show(mess, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    OkuriJyo okuriJyo = new OkuriJyo();
                    okuriJyo.SearchParameter(gr);
                }
            }
        }

        //Read Haiso info from oracle and output an excel file
        private void B_HAISODOWNLOAD_Click(object sender, EventArgs e)
        {
            if (GrdCheck() != 0 && StatusCheck(new string[] { "集約", "配送済", "運賃計算", "実績送信済" }) == true)
            {

                bool ZanKa = CH_ZANKA.Checked;
                if (ZanKa)
                {
                    string mess = "選択したレコードは全て残貨対象のため処理できません。";
                    MessageBox.Show(mess, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    HaisoDl haisoDl = new HaisoDl();
                    haisoDl.CreateExcel(GR_LIST);
                    Search();
                }
            }
        }

        //Compute Unchin info and update the database with unchin
        private void B_UNCHINCAL_Click(object sender, EventArgs e)
        {
            if (GrdCheck() != 0 && StatusCheck(new string[] { "配送済", "運賃計算" }) == true)
            {
                DialogResult myResult;

                string mess = GrdCheck() + "件の運賃計算を行います。よろしいですか？";
                myResult = MessageBox.Show(mess, "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (myResult == DialogResult.Yes)
                {

                    Search();
                }
                else
                {
                    MessageBox.Show("処理を中止しました。");
                }
            }
        }

        //Read Unchin info from oracle and output an excel file
        private void B_UNCHINDOWNLOAD_Click(object sender, EventArgs e)
        {
            if (GrdCheck() != 0 && StatusCheck(new string[] { "運賃計算", "実績送信済" }) == true)
            {
                bool ZanKa = CH_ZANKA.Checked;
                if (ZanKa)
                {
                    string mess = "選択したレコードは全て残貨対象のため処理できません。";
                    MessageBox.Show(mess, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    DataGridView gr = GR_LIST;
                    UnchinDl unchinDl = new UnchinDl();
                    unchinDl.SearchParameter(gr);
                    Search();
                }

            }
        }
        
        private void E_SYUKABI1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete)
            {
                E_SYUKABI1.CustomFormat = " ";
                E_SYUKABI1.Format = DateTimePickerFormat.Custom;
                E_SYUKABI1.Checked = true;
            }

        }

        private void E_SYUKABI1_ValueChanged(object sender, EventArgs e)
        {
            if (init == false)
            {
                E_SYUKABI1.CustomFormat = "yyyy/MM/dd";
                E_SYUKABI1.Checked = false;
            }
        }

        private void E_SYUKABI2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete)
            {
                E_SYUKABI2.CustomFormat = " ";
                E_SYUKABI2.Format = DateTimePickerFormat.Custom;
                E_SYUKABI2.Checked = true;
            }
        }
        private void E_SYUKABI2_ValueChanged(object sender, EventArgs e)
        {
            if (init == false)
            {
                E_SYUKABI2.CustomFormat = "yyyy/MM/dd";
                E_SYUKABI2.Checked = false;
            }
            else { init = false; }
        }

        private void E_EXLSRD1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete)
            {
                E_EXLSRD1.CustomFormat = " ";
                E_EXLSRD1.Format = DateTimePickerFormat.Custom;
                E_EXLSRD1.Checked = true;
            }
        }
        private void E_EXLSRD1_ValueChanged(object sender, EventArgs e)
        {
            if (init == false)
            {
                E_EXLSRD1.CustomFormat = "yyyy/MM/dd";
                E_EXLSRD1.Checked = false;
            }
        }

        private void E_EXLSRD2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete)
            {
                E_EXLSRD2.CustomFormat = " ";
                E_EXLSRD2.Format = DateTimePickerFormat.Custom;
                E_EXLSRD2.Checked = true;
            }
        }
        private void E_EXLSRD2_ValueChanged(object sender, EventArgs e)
        {
            if (init == false)
            { 
                E_EXLSRD2.CustomFormat = "yyyy/MM/dd";
                E_EXLSRD2.Checked = false;
            }
        }

        private void E_UNCHIN1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete)
            {
                E_UNCHIN1.CustomFormat = " ";
                E_UNCHIN1.Format = DateTimePickerFormat.Custom;
                E_UNCHIN1.Checked = true;
            }
        }
        private void E_UNCHIN1_ValueChanged(object sender, EventArgs e)
        {
            if (init == false)
            {
                E_UNCHIN1.CustomFormat = "yyyy/MM/dd";
                E_UNCHIN1.Checked = false;
            }
        }

        private void E_UNCHIN2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete)
            {
                E_UNCHIN2.CustomFormat = " ";
                E_UNCHIN2.Format = DateTimePickerFormat.Custom;
                E_UNCHIN2.Checked = true;
            }
        }
        private void E_UNCHIN2_ValueChanged(object sender, EventArgs e)
        {
            if(init == false)
            { 
                E_UNCHIN2.CustomFormat = "yyyy/MM/dd";
                E_UNCHIN2.Checked = false;
            }
        }
        

        //Search Oracle with parameters provided from interface
        private void B_SEARCH_Click(object sender, EventArgs e)
        {
            try
            {
                log.Info("SEARCH BEGIN");
                Search();
            }
            catch(Exception err)
            {
                log.Fatal("UNEXPECTED ERROR: " + err.Message);
                log.Info("SEARCH END");
            }
            finally
            {
                log.Info("SEARCH END");
            }
        }
        //Search function 
        public void Search()
        {
            var aaa = E_UNCHIN2.Checked;
            GR_LIST.DataSource = null;
            object Area = CB_AREA.SelectedValue;

            bool MNR = CH_MNR.Checked;
            bool SNK = CH_SNK.Checked;
            bool TYU = CH_TYU.Checked;
            bool CYU = CH_CYU.Checked;
            bool APU = CH_APU.Checked;

            bool Status1 = radioButton1.Checked;
            bool Status2 = radioButton3.Checked;
            bool Status3 = radioButton2.Checked;
            bool Status4 = radioButton4.Checked;
            bool Status5 = radioButton5.Checked;
            bool Status6 = radioButton6.Checked;
            bool Status7 = radioButton7.Checked;

            string SyukaBi1 = E_SYUKABI1.Value.ToString();
            string SyukaBi2 = E_SYUKABI2.Value.ToString();
            string Exlsrd1 = E_EXLSRD1.Value.ToString();
            string Exlsrd2 = E_EXLSRD2.Value.ToString();
            string Unchin1 = E_UNCHIN1.Value.ToString();
            string Unchin2 = E_UNCHIN2.Value.ToString();

            string OrderNO = E_ORDERNO.Text;
            string HaisoBnNO = E_HAISOBNNO.Text;

            bool ZanKa = CH_ZANKA.Checked;

            if (OrderNO.Length >= 0 || SyukaBi1.Length >= 0)
            {

                Search search = new Search();
                DataTable ds = new DataTable();
                ds = search.SearchByParameter(MNR, SNK, TYU, CYU, APU, Status1, Status2, Status3, Status4, Status5, Status6, Status7, SyukaBi1, SyukaBi2, Exlsrd1,
                    Exlsrd2, Unchin1, Unchin2, OrderNO, HaisoBnNO, ZanKa, Area, !E_SYUKABI1.Checked, !E_SYUKABI2.Checked, !E_EXLSRD1.Checked, !E_EXLSRD2.Checked, !E_UNCHIN1.Checked, !E_UNCHIN2.Checked);
                if (ds.Rows.Count <= 0)
                {
                    MessageBox.Show("対象データがありません。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    label6.Text = label7.Text = "0";
                }
                else if (ds.Rows.Count > int.Parse(Tools.GetKenSu()))
                {
                    string mess = "最大件数" + Tools.GetKenSu() + "を越えています。検索条件を変更してください。";
                    MessageBox.Show(mess, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    label6.Text = label7.Text = "0";
                }
                else
                {
                    GR_LIST.DataSource = ds;
                    GR_LIST.Sort(ZNKFLG, 0);
                    int totalCount = 0;
                    double totalWeight = 0;
                    for (int i = 0; i < GR_LIST.RowCount; i++)
                    {
                        var nmsl = GR_LIST.Rows[i].Cells["ZNKFLG"].Value;
                        if (nmsl is DBNull)
                        {
                            totalCount += Convert.ToInt32(GR_LIST.Rows[i].Cells["KOSU"].Value);
                            var bb = GR_LIST.Rows[i].Cells["WT"].Value;
                            totalWeight += Convert.ToDouble(bb);
                        }
                    }
                    label6.Text = totalCount.ToString();
                    label7.Text = totalWeight.ToString();
                }
            }
            else
            {
                MessageBox.Show("Error");

            }
        }
        //Read Syuyaku info from oracle and output an excel file
        private void B_SYKDOWNLOAD_Click(object sender, EventArgs e)
        {
            if (GrdCheck() != 0 && StatusCheck(new string[] { "未処理", "集荷表", "集約" }) == true)
            {
                object Area = CB_AREA.SelectedValue;

                bool MNR = CH_MNR.Checked;
                bool SNK = CH_SNK.Checked;
                bool TYU = CH_TYU.Checked;
                bool CYU = CH_CYU.Checked;
                bool APU = CH_APU.Checked;

                bool Status1 = radioButton1.Checked;
                bool Status2 = radioButton3.Checked;
                bool Status3 = radioButton2.Checked;
                bool Status4 = radioButton4.Checked;
                bool Status5 = radioButton5.Checked;
                bool Status6 = radioButton6.Checked;
                bool Status7 = radioButton7.Checked;

                string SyukaBi1 = E_SYUKABI1.Value.ToString();
                string SyukaBi2 = E_SYUKABI2.Value.ToString();
                string Exlsrd1 = E_EXLSRD1.Value.ToString();
                string Exlsrd2 = E_EXLSRD2.Value.ToString();
                string Unchin1 = E_UNCHIN1.Value.ToString();
                string Unchin2 = E_UNCHIN2.Value.ToString();

                string OrderNO = E_ORDERNO.Text;
                string HaisoBnNO = E_HAISOBNNO.Text;

                bool ZanKa = CH_ZANKA.Checked;


                DataGridView gr = GR_LIST;
                DownLoad dl = new DownLoad();
                DataTable ds = new DataTable();
                //検索
                ds = dl.SearchParameter(MNR, SNK, TYU, CYU, APU, Status1, Status2, Status3, Status4, Status5, Status6, Status7, SyukaBi1, SyukaBi2, Exlsrd1,
                    Exlsrd2, Unchin1, Unchin2, OrderNO, HaisoBnNO, ZanKa, Area, !E_SYUKABI1.Checked, !E_SYUKABI2.Checked, !E_EXLSRD1.Checked, !E_EXLSRD2.Checked, !E_UNCHIN1.Checked, !E_UNCHIN2.Checked);
                //excel出力
                dl.SyuYakuDL(gr, ds);
            }
        }

        //Read the Syuyaku file info and update Oracle accordingly
        private void B_SYKUPLOAD_Click(object sender, EventArgs e)
        {
            UploadSyuyaku up = new UploadSyuyaku();

            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        log.Info("EXEC BEGIN");
                        //Get the path of specified file
                        filePath = openFileDialog.FileName;
                        up.UpdateByDatable(filePath);
                        Search();
                    }
                    catch (Exception err)
                    {
                        log.Fatal("Unexpected Error: " + err.Message);
                    }
                    finally
                    {
                        log.Info("EXEC END");
                    }
                }
                else
                {
                    MessageBox.Show("処理を中止しました。");
                }
            }
        }

        

        //Select all / Deselect all
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            bool selectAllFlg = true;

            if (checkBox1.Checked)
            {
                selectAllFlg = true;
            }
            else
            {
                selectAllFlg = false;
            }
            for (int i = 0; i < this.GR_LIST.RowCount; i++)
            {
                //this.GR_LIST.EndEdit();

                //string re_value = this.GR_LIST.Rows[i].Cells[0].EditedFormattedValue.ToString();

                this.GR_LIST.Rows[i].Cells[0].Value = selectAllFlg;
                this.GR_LIST.Rows[i].Selected = selectAllFlg;
            }

        }

        //Enable / Disable the Unchin date selector
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkUnchin())
            {
                LabE_UNCHIN.Enabled = true;
                E_UNCHIN1.Enabled = true;
                E_UNCHIN2.Enabled = true;
            }
            else
            {
                LabE_UNCHIN.Enabled = false;
                E_UNCHIN1.Enabled = false;
                E_UNCHIN2.Enabled = false;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkUnchin())
            {
                LabE_UNCHIN.Enabled = true;
                E_UNCHIN1.Enabled = true;
                E_UNCHIN2.Enabled = true;
            }
            else
            {
                LabE_UNCHIN.Enabled = false;
                E_UNCHIN1.Enabled = false;
                E_UNCHIN2.Enabled = false;
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkUnchin())
            {
                LabE_UNCHIN.Enabled = true;
                E_UNCHIN1.Enabled = true;
                E_UNCHIN2.Enabled = true;
            }
            else
            {
                LabE_UNCHIN.Enabled = false;
                E_UNCHIN1.Enabled = false;
                E_UNCHIN2.Enabled = false;
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkUnchin())
            {
                LabE_UNCHIN.Enabled = true;
                E_UNCHIN1.Enabled = true;
                E_UNCHIN2.Enabled = true;
            }
            else
            {
                LabE_UNCHIN.Enabled = false;
                E_UNCHIN1.Enabled = false;
                E_UNCHIN2.Enabled = false;
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkUnchin())
            {
                LabE_UNCHIN.Enabled = true;
                E_UNCHIN1.Enabled = true;
                E_UNCHIN2.Enabled = true;
            }
            else
            {
                LabE_UNCHIN.Enabled = false;
                E_UNCHIN1.Enabled = false;
                E_UNCHIN2.Enabled = false;
            }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkUnchin())
            {
                LabE_UNCHIN.Enabled = true;
                E_UNCHIN1.Enabled = true;
                E_UNCHIN2.Enabled = true;
            }
            else
            {
                LabE_UNCHIN.Enabled = false;
                E_UNCHIN1.Enabled = false;
                E_UNCHIN2.Enabled = false;
            }
        }
        public bool checkUnchin()
        {
            if (radioButton5.Checked || radioButton6.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void B_CLOSE_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //If row header selected then select this row
        private void myDataGrid_OnCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == myCheckBoxColumn.Index && e.RowIndex != -1)
            {
                for (int i = 0; i < this.GR_LIST.RowCount; i++)
                {
                    this.GR_LIST.Rows[i].Selected = (bool)GR_LIST.Rows[i].Cells[0].EditedFormattedValue;
                }
            }
        }
        private void myDataGrid_OnCellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            // End of edition on each click on column of checkbox
            if (e.ColumnIndex == myCheckBoxColumn.Index && e.RowIndex != -1)
            {
                GR_LIST.EndEdit();
            }
        }

     
    }
}