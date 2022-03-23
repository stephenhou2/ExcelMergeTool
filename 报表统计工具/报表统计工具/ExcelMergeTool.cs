using System;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace 报表统计工具
{
    public partial class ExcelMergeTool : Form
    {
        public static ExcelMergeTool Ins;

        public ExcelMergeTool()
        {
            Ins = this;
            InitializeComponent();
            InitializeOther();
        }

        private void InitializeOther()
        {
            Utility.LoadExcelMergeToolSetting();
        }

        #region LOG
        private StringBuilder _excute_result = new StringBuilder();
        private StringBuilder _excute_errors = new StringBuilder();
        public void UpdateOutputFormat(string s,params object[] args)
        {
            _excute_result.Append("\r\n");
            _excute_result.AppendFormat(s, args);
            Label_Excute_Result.Text = _excute_result.ToString();
        }

        int errorCnt = 0;
        public void UpdateErrorOutputFormat(string s, params object[] args)
        {
            _excute_result.AppendFormat("\r\n!!!!!!!!!!!!!!  ");
            _excute_result.AppendFormat(s, args);
            Label_Excute_Result.Text = _excute_result.ToString();

            _excute_errors.AppendFormat("\r\n{0}.",errorCnt++);
            _excute_errors.AppendFormat(s, args);
        }
        #endregion

        private string _selected_dir_path = string.Empty;
        private void Btn_Select_Dir_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "选择文件夹";
            fbd.ShowNewFolderButton = false;
            
            if(fbd.ShowDialog() == DialogResult.OK)
            {
                _selected_dir_path = fbd.SelectedPath;
                Label_Merge_Dir.Text = _selected_dir_path;
                UpdateOutputFormat("合并报表--选择文件夹  {0}", _selected_dir_path);
            }
        }

        private void Btn_Excute_Merge_Click(object sender, EventArgs e)
        {
            errorCnt = 0;
            if (string.IsNullOrEmpty(_selected_dir_path) || !Directory.Exists(_selected_dir_path))
            {
                UpdateErrorOutputFormat("非法的文件夹路径，请检查!  {0}",_selected_dir_path);
                return;
            }

            Utility.MergeExcelInDir(_selected_dir_path);

            if(_excute_errors.Length == 0)
            {
                UpdateOutputFormat(@"
-----------------------------------------------------------------------------------------------------------------------
                                                                  
                                               报表合并成功                     
                                                                  
-----------------------------------------------------------------------------------------------------------------------
");
            }
            else
            {
                Label_Excute_Result.Text = _excute_result.ToString() + string.Format(@"


xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                                                                  
                                             报表合并中有错误!                       
{0}
                                                                 
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
", _excute_errors.ToString());
            }
        }

        private void Btn_SetExcelModule_Click(object sender, EventArgs e)
        {

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "选择Excel模板文件";
            ofd.InitialDirectory = Application.StartupPath;
            ofd.CheckFileExists = true;
            ofd.Multiselect = false;
            ofd.CheckPathExists = true;
            ofd.RestoreDirectory = true;
            ofd.Filter = "excel文件（03）|*.xls|excel文件（07）|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                DataCenter.MoudlePath = ofd.FileName;
                Label_ExcelDataModule.Text = ofd.FileName;
                Utility.SaveExcelModulePath(ofd.FileName);
            }
        }
    }
}
