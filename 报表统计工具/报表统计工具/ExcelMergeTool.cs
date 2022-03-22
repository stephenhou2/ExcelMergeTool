using System;
using System.Windows.Forms;
using System.IO;
using System.Text;

namespace 报表统计工具
{
    public partial class ExcelMergeTool : Form
    {
        public ExcelMergeTool()
        {
            InitializeComponent();
            InitializeOther();
        }

        private void InitializeOther()
        {
            LoadExcelMergeToolSetting();
        }

        private StringBuilder _excute_result = new StringBuilder();
        private StringBuilder _excute_errors = new StringBuilder();
        private void UpdateOutputFormat(string s,params object[] args)
        {
            _excute_result.Append("\r\n");
            _excute_result.AppendFormat(s, args);
            Label_Excute_Result.Text = _excute_result.ToString();
        }

        int errorCnt = 0;
        private void UpdateErrorOutputFormat(string s, params object[] args)
        {
            _excute_result.AppendFormat("\r\n!!!!!!!!!!!!!!  ");
            _excute_result.AppendFormat(s, args);
            Label_Excute_Result.Text = _excute_result.ToString();

            _excute_errors.AppendFormat("\r\n{0}.",errorCnt++);
            _excute_errors.AppendFormat(s, args);
        }

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

            MergeExcelInDir(_selected_dir_path);

            if(_excute_errors.Length == 0)
            {
                UpdateOutputFormat(@"
-------------------------------------------------------------------------------------------------------------------------------------------------
                                                                  
                                               报表合并成功                     
                                                                  
--------------------------------------------------------------------------------------------------------------------------------------------------
");
            }
            else
            {
                Label_Excute_Result.Text = _excute_result.ToString() + string.Format(@"


xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                                                                  
                                             报表合并中有错误!                       
{0}
                                                                 
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
", _excute_errors.ToString());
            }
        }
    }
}
