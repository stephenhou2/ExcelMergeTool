using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace 报表统计工具
{
    partial class ExcelMergeTool : Form
    {
        private void LoadExcelMergeToolSetting()
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(ExcelMergeToolDef.ToolSettingsPath);

            XmlNode root = xml.SelectSingleNode("root");
            if(root != null)
            {
                XmlElement modelPathNode = root.SelectSingleNode("模板路径") as XmlElement;

                if(modelPathNode != null)
                {
                    if(File.Exists(modelPathNode.InnerText))
                    {
                        _excel_module_path = modelPathNode.InnerText;
                        Label_ExcelDataModule.Text = _excel_module_path;
                    }
                    else
                    {
                        UpdateErrorOutputFormat("模板路径不存在，请检查!  {0}", modelPathNode.InnerText);
                    }
                }
                else
                {
                    UpdateErrorOutputFormat("配置文件中模板数据格式错误，请检查!");
                }
            }
        }

        private string _excel_module_path = string.Empty;
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
                _excel_module_path = ofd.FileName;
                Label_ExcelDataModule.Text = _excel_module_path;
                SaveExcelModulePath(_excel_module_path);
            }
        }

        private void SaveExcelModulePath(string modulePath)
        {
            if(string.IsNullOrEmpty(modulePath) || !File.Exists(modulePath))
            {
                UpdateErrorOutputFormat("配置模板文件失败,请检查!  {0}", modulePath);
                return;
            }

            XmlDocument xml = new XmlDocument();
            XmlElement root_node = xml.CreateElement("root");
            xml.AppendChild(root_node);

            XmlElement mp_node = xml.CreateElement("模板路径");
            mp_node.InnerText = modulePath;
            root_node.AppendChild(mp_node);

            xml.Save(ExcelMergeToolDef.ToolSettingsPath);
            UpdateOutputFormat("模板文件更改成功 {0}",modulePath);
        }


        private Dictionary<string, IWorkbook> ReadAllExcels(string dirPath)
        {
            Dictionary<string, IWorkbook> allWorkbooks = new Dictionary<string, IWorkbook>();
            if (string.IsNullOrEmpty(dirPath) || !Directory.Exists(dirPath))
            {
                UpdateErrorOutputFormat("合并报表失败，无效的文件夹路径--{0}", dirPath);
                return allWorkbooks;
            }

            DirectoryInfo di = new DirectoryInfo(dirPath);
            FileInfo[] files = di.GetFiles();
            if (files == null || files.Length == 0)
            {
                UpdateErrorOutputFormat("合并报表失败，空文件夹--{0}", dirPath);
                return allWorkbooks;
            }

            FileStream fs = null;
            for (int i = 0; i < files.Length; i++)
            {
                string filePath = files[i].FullName;
                if (!filePath.EndsWith(".xls") && !filePath.EndsWith(".xlsl") || filePath.Contains("统计结果"))
                {
                    continue;
                }

                fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                if (fs != null)
                {
                    IWorkbook workbook = null;
                    if (filePath.EndsWith(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    else if (filePath.EndsWith(".xls"))
                    {
                        workbook = new HSSFWorkbook(fs);
                    }

                    if (workbook != null)
                    {
                        if(!allWorkbooks.ContainsKey(filePath))
                        {
                            allWorkbooks.Add(filePath,workbook);
                            UpdateOutputFormat("文件{0}解析成功", filePath);
                        }
                        else
                        {
                            UpdateErrorOutputFormat("文件名重复--{0}", filePath);
                        }
                    }
                    else
                    {
                        UpdateErrorOutputFormat("文件{0}解析失败", filePath);
                    }
                }
                else
                {
                    UpdateErrorOutputFormat("文件{0}读取失败", filePath);
                }

                if (fs != null)
                {
                    fs.Dispose();
                    fs.Close();
                }
            }


            return allWorkbooks;
        }

        private IWorkbook ReadModuleWorkbook()
        {
            IWorkbook workbook = null;
            FileStream fs = File.Open(_excel_module_path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            if (fs != null)
            {
                if (_excel_module_path.EndsWith(".xlsx"))
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (_excel_module_path.EndsWith(".xls"))
                {
                    workbook = new HSSFWorkbook(fs);
                }

                if (workbook != null)
                {
                    UpdateOutputFormat("模板文件{0}解析成功", _excel_module_path);
                }
                else
                {
                    UpdateErrorOutputFormat("模板文件{0}解析失败", _excel_module_path);
                }
            }
            else
            {
                UpdateErrorOutputFormat("文件{0}读取失败", _excel_module_path);
            }

            if(workbook != null)
            {
                return workbook.Copy();
            }

            return null;
        }

        private double MergeSumAt(Dictionary<string,IWorkbook> toMerges,string sheetName,int row,int col)
        {
            double result = 0;
            foreach(KeyValuePair<string,IWorkbook> kv in toMerges)
            {
                string filename = kv.Key;
                IWorkbook workbook = kv.Value;
                ISheet sheet = workbook.GetSheet(sheetName);

                if(sheet == null)
                {
                    UpdateErrorOutputFormat("表单{0}中没有{1}切页", filename, sheetName);
                    continue;
                }

                IRow rowData = sheet.GetRow(row);
                if(rowData == null)
                {
                    UpdateErrorOutputFormat("表单  【{0}】  切页  【{1}】  中缺少数据行row={2}", filename, sheetName,row);
                    continue;
                }

                ICell cellData = rowData.GetCell(col);
                if(cellData == null)
                {
                    UpdateErrorOutputFormat("表单  【{0}】   切页 【{1}】 中数据行row={2}缺少数据列col={3}", filename, sheetName, row);
                    continue;
                }

                if(cellData.CellType == CellType.Numeric)
                {
                    result += cellData.NumericCellValue;
                }
                else
                {
                    UpdateErrorOutputFormat("表单  【{0}】  切页  【{1}】  中数据行row={2}数据列col={3} 数据格式错误，应该是一个数字", filename, sheetName, row, col);
                    continue;
                }
            }

            return result;
        }

        private void Merge(IWorkbook module, Dictionary<string, IWorkbook> toMerges)
        {
            int sheetCnt = module.NumberOfSheets;
            for(int i=0;i<sheetCnt;i++)
            {
                // 待合并的页
                ISheet sheet = module.GetSheetAt(i);
                int rowNum = sheet.LastRowNum;
                for(int row=0;row<rowNum;row++)
                {
                    IRow rowData = sheet.GetRow(row);
                    if (rowData == null)
                        continue;

                    int colNum = rowData.LastCellNum;
                    for (int col = 0; col < colNum; col++)
                    {
                        ICell cellData = rowData.GetCell(col);
                        if (cellData == null)
                            continue;

                        if(cellData.CellComment != null)
                        {
                            if(cellData.CellComment.String.String == "SUM")
                            {
                                cellData.SetCellValue(MergeSumAt(toMerges, sheet.SheetName, row, col));
                            }
                        }
                    }
                }
            }
        }

        private void MergeExcelInDir(string dirPath)
        {
            Dictionary<string, IWorkbook> allWorkbooks = ReadAllExcels(dirPath);
            IWorkbook moduleWorkbook = ReadModuleWorkbook();

            if (allWorkbooks.Count == 0 || moduleWorkbook == null)
                return;

            Merge(moduleWorkbook, allWorkbooks);

            FileStream fs = new FileStream(dirPath + "/统计结果.xls",FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.ReadWrite);
            moduleWorkbook.Write(fs);
            fs.Dispose();
            fs.Close();
        }
    }
}
