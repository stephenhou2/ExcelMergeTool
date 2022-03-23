using System;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace 报表统计工具
{
    partial class Utility
    {
        /// <summary>
        /// 加载配置文件，读取配置信息
        /// </summary>
        /// <returns></returns>
        public static (int, string) LoadExcelMergeToolSetting()
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(ExcelMergeToolDef.ToolSettingsPath);

            XmlNode root = xml.SelectSingleNode("root");
            if (root != null)
            {
                XmlElement modelPathNode = root.SelectSingleNode("模板路径") as XmlElement;

                if (modelPathNode != null)
                {
                    if (File.Exists(modelPathNode.InnerText))
                    {
                        DataCenter.MoudlePath = modelPathNode.InnerText;
                        return (Define.ReadSettingResult_Success, modelPathNode.InnerText);
                    }
                    else
                    {
                        ExcelMergeTool.Ins.UpdateErrorOutputFormat("模板路径不存在，请检查!  {0}", modelPathNode.InnerText);
                        return (Define.ReadSettingResult_Failed, null);

                    }
                }
            }

            ExcelMergeTool.Ins.UpdateErrorOutputFormat("配置文件中模板数据格式错误，请检查!");
            return (Define.ReadSettingResult_Failed, null);
        }

        /// <summary>
        /// 保存配置文件
        /// 目前存储模板文件路径
        /// </summary>
        /// <param name="modulePath"></param>
        /// <returns></returns>
        public static int SaveExcelModulePath(string modulePath)
        {
            if (string.IsNullOrEmpty(modulePath) || !File.Exists(modulePath))
            {
                ExcelMergeTool.Ins.UpdateErrorOutputFormat("配置模板文件失败,请检查!  {0}", modulePath);
                return Define.SaveSettingResult_Failed;
            }

            XmlDocument xml = new XmlDocument();
            XmlElement root_node = xml.CreateElement("root");
            xml.AppendChild(root_node);

            XmlElement mp_node = xml.CreateElement("模板路径");
            mp_node.InnerText = modulePath;
            root_node.AppendChild(mp_node);

            xml.Save(ExcelMergeToolDef.ToolSettingsPath);

            ExcelMergeTool.Ins.UpdateOutputFormat("模板文件更改成功 {0}", modulePath);
            return Define.SaveSettingResult_Success;
        }


        private static Dictionary<string, IWorkbook> ReadAllExcels(string dirPath)
        {
            Dictionary<string, IWorkbook> allWorkbooks = new Dictionary<string, IWorkbook>();
            if (string.IsNullOrEmpty(dirPath) || !Directory.Exists(dirPath))
            {
                ExcelMergeTool.Ins.UpdateErrorOutputFormat("合并报表失败，无效的文件夹路径--{0}", dirPath);
                return allWorkbooks;
            }

            DirectoryInfo di = new DirectoryInfo(dirPath);
            FileInfo[] files = di.GetFiles();
            if (files == null || files.Length == 0)
            {
                ExcelMergeTool.Ins.UpdateErrorOutputFormat("合并报表失败，空文件夹--{0}", dirPath);
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
                        if (!allWorkbooks.ContainsKey(filePath))
                        {
                            allWorkbooks.Add(filePath, workbook);
                            ExcelMergeTool.Ins.UpdateOutputFormat("文件{0}解析成功", filePath);
                        }
                        else
                        {
                            ExcelMergeTool.Ins.UpdateErrorOutputFormat("文件名重复--{0}", filePath);
                        }
                    }
                    else
                    {
                        ExcelMergeTool.Ins.UpdateErrorOutputFormat("文件{0}解析失败", filePath);
                    }
                }
                else
                {
                    ExcelMergeTool.Ins.UpdateErrorOutputFormat("文件{0}读取失败", filePath);
                }

                if (fs != null)
                {
                    fs.Dispose();
                    fs.Close();
                }
            }


            return allWorkbooks;
        }

        private static IWorkbook ReadModuleWorkbook()
        {
            IWorkbook workbook = null;
            FileStream fs = File.Open(DataCenter.MoudlePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            if (fs != null)
            {
                if (DataCenter.MoudlePath.EndsWith(".xlsx"))
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (DataCenter.MoudlePath.EndsWith(".xls"))
                {
                    workbook = new HSSFWorkbook(fs);
                }

                if (workbook != null)
                {
                    ExcelMergeTool.Ins.UpdateOutputFormat("模板文件{0}解析成功", DataCenter.MoudlePath);
                }
                else
                {
                    ExcelMergeTool.Ins.UpdateErrorOutputFormat("模板文件{0}解析失败", DataCenter.MoudlePath);
                }
                fs.Dispose();
                fs.Close();
            }
            else
            {
                ExcelMergeTool.Ins.UpdateErrorOutputFormat("文件{0}读取失败", DataCenter.MoudlePath);
            }


            if (workbook != null)
            {
                return workbook.Copy();
            }


            return null;
        }

        private static double MergeSumAt(Dictionary<string, IWorkbook> toMerges, string sheetName, int row, int col)
        {
            double result = 0;
            foreach (KeyValuePair<string, IWorkbook> kv in toMerges)
            {
                string filename = kv.Key;
                IWorkbook workbook = kv.Value;
                ISheet sheet = workbook.GetSheet(sheetName);

                if (sheet == null)
                {
                    ExcelMergeTool.Ins.UpdateErrorOutputFormat("表单{0}中没有{1}切页", filename, sheetName);
                    continue;
                }

                IRow rowData = sheet.GetRow(row);
                if (rowData == null)
                {
                    ExcelMergeTool.Ins.UpdateErrorOutputFormat("表单  【{0}】  切页  【{1}】  中缺少数据行row={2}", filename, sheetName, row);
                    continue;
                }

                ICell cellData = rowData.GetCell(col);
                if (cellData == null)
                {
                    ExcelMergeTool.Ins.UpdateErrorOutputFormat("表单  【{0}】   切页 【{1}】 中数据行row={2}缺少数据列col={3}", filename, sheetName, row);
                    continue;
                }

                if (cellData.CellType == CellType.Numeric)
                {
                    result += cellData.NumericCellValue;
                }
                else
                {
                    ExcelMergeTool.Ins.UpdateErrorOutputFormat("表单  【{0}】  切页  【{1}】  中数据行row={2}数据列col={3} 数据格式错误，应该是一个数字", filename, sheetName, row, col);
                    continue;
                }
            }

            return result;
        }

        private static string GetCellString(ICell cellData)
        {
            switch (cellData.CellType)
            {
                case CellType.Numeric:
                    return cellData.NumericCellValue.ToString();
                case CellType.String:
                    return cellData.StringCellValue;
                default:
                    return null;
            }
        }

        private static void Merge(IWorkbook module, Dictionary<string, IWorkbook> toMerges)
        {
            int sheetCnt = module.NumberOfSheets;

            // 第一轮遍历，根据模板表的配置合并数据
            for (int i = 0; i < sheetCnt; i++)
            {
                // 待合并的页
                ISheet sheet = module.GetSheetAt(i);
                int rowNum = sheet.LastRowNum;

                for (int row = 0; row < rowNum; row++)
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

                        string cellStr = GetCellString(cellData);
                        if (!string.IsNullOrEmpty(cellStr))
                        {
                            if (cellStr == "SUM")
                            {
                                cellData.SetCellValue(MergeSumAt(toMerges, sheet.SheetName, row, col));
                            }
                        }
                    }
                }
            }

            // 第二轮遍历，根据公式刷新表格内容
            for (int i = 0; i < sheetCnt; i++)
            {
                ISheet sheet = module.GetSheetAt(i);
                sheet.ForceFormulaRecalculation = true;
            }
        }

        public static void MergeExcelInDir(string dirPath)
        {
            Dictionary<string, IWorkbook> allWorkbooks = ReadAllExcels(dirPath);
            IWorkbook moduleWorkbook = ReadModuleWorkbook();

            if (allWorkbooks.Count == 0 || moduleWorkbook == null)
                return;


            Merge(moduleWorkbook, allWorkbooks);

            string filePath = dirPath + "/统计结果.xls";
            try
            {
                FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
                moduleWorkbook.Write(fs);
                fs.Dispose();
                fs.Close();
            }
            catch (System.IO.IOException e)
            {
                ExcelMergeTool.Ins.UpdateErrorOutputFormat("写入异常，请确保 文件【{0}】已经关闭，否则无法正常写入", filePath);
            }
        }
    }
}
