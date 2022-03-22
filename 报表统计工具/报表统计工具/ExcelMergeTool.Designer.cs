
namespace 报表统计工具
{
    partial class ExcelMergeTool
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.Label_ExcelDataModule = new System.Windows.Forms.Label();
            this.Btn_SetExcelModule = new System.Windows.Forms.Button();
            this.Btn_Select_Dir = new System.Windows.Forms.Button();
            this.Label_Merge_Dir = new System.Windows.Forms.Label();
            this.ListView_ExcelsInDir = new System.Windows.Forms.ListView();
            this.label1 = new System.Windows.Forms.Label();
            this.Btn_Excute_Merge = new System.Windows.Forms.Button();
            this.Label_Excute_Result = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // Label_ExcelDataModule
            // 
            this.Label_ExcelDataModule.AutoSize = true;
            this.Label_ExcelDataModule.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Label_ExcelDataModule.Location = new System.Drawing.Point(104, 13);
            this.Label_ExcelDataModule.Name = "Label_ExcelDataModule";
            this.Label_ExcelDataModule.Size = new System.Drawing.Size(83, 12);
            this.Label_ExcelDataModule.TabIndex = 0;
            this.Label_ExcelDataModule.Text = "Excel模板路径";
            // 
            // Btn_SetExcelModule
            // 
            this.Btn_SetExcelModule.Location = new System.Drawing.Point(12, 8);
            this.Btn_SetExcelModule.Name = "Btn_SetExcelModule";
            this.Btn_SetExcelModule.Size = new System.Drawing.Size(75, 23);
            this.Btn_SetExcelModule.TabIndex = 1;
            this.Btn_SetExcelModule.Text = "设置模板";
            this.Btn_SetExcelModule.UseVisualStyleBackColor = true;
            this.Btn_SetExcelModule.Click += new System.EventHandler(this.Btn_SetExcelModule_Click);
            // 
            // Btn_Select_Dir
            // 
            this.Btn_Select_Dir.Location = new System.Drawing.Point(13, 38);
            this.Btn_Select_Dir.Name = "Btn_Select_Dir";
            this.Btn_Select_Dir.Size = new System.Drawing.Size(75, 23);
            this.Btn_Select_Dir.TabIndex = 2;
            this.Btn_Select_Dir.Text = "选择目录";
            this.Btn_Select_Dir.UseVisualStyleBackColor = true;
            this.Btn_Select_Dir.Click += new System.EventHandler(this.Btn_Select_Dir_Click);
            // 
            // Label_Merge_Dir
            // 
            this.Label_Merge_Dir.AutoSize = true;
            this.Label_Merge_Dir.Location = new System.Drawing.Point(104, 43);
            this.Label_Merge_Dir.Name = "Label_Merge_Dir";
            this.Label_Merge_Dir.Size = new System.Drawing.Size(53, 12);
            this.Label_Merge_Dir.TabIndex = 3;
            this.Label_Merge_Dir.Text = "文件目录";
            // 
            // ListView_ExcelsInDir
            // 
            this.ListView_ExcelsInDir.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.ListView_ExcelsInDir.HideSelection = false;
            this.ListView_ExcelsInDir.Location = new System.Drawing.Point(13, 68);
            this.ListView_ExcelsInDir.Name = "ListView_ExcelsInDir";
            this.ListView_ExcelsInDir.Size = new System.Drawing.Size(763, 139);
            this.ListView_ExcelsInDir.TabIndex = 4;
            this.ListView_ExcelsInDir.UseCompatibleStateImageBehavior = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 214);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "输出：";
            // 
            // Btn_Excute_Merge
            // 
            this.Btn_Excute_Merge.Location = new System.Drawing.Point(701, 418);
            this.Btn_Excute_Merge.Name = "Btn_Excute_Merge";
            this.Btn_Excute_Merge.Size = new System.Drawing.Size(75, 23);
            this.Btn_Excute_Merge.TabIndex = 7;
            this.Btn_Excute_Merge.Text = "运行";
            this.Btn_Excute_Merge.UseVisualStyleBackColor = true;
            this.Btn_Excute_Merge.Click += new System.EventHandler(this.Btn_Excute_Merge_Click);
            // 
            // Label_Excute_Result
            // 
            this.Label_Excute_Result.BackColor = System.Drawing.SystemColors.Window;
            this.Label_Excute_Result.Location = new System.Drawing.Point(12, 230);
            this.Label_Excute_Result.Multiline = true;
            this.Label_Excute_Result.Name = "Label_Excute_Result";
            this.Label_Excute_Result.ReadOnly = true;
            this.Label_Excute_Result.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label_Excute_Result.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.Label_Excute_Result.Size = new System.Drawing.Size(764, 182);
            this.Label_Excute_Result.TabIndex = 8;
            // 
            // ExcelMergeTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.Label_Excute_Result);
            this.Controls.Add(this.Btn_Excute_Merge);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ListView_ExcelsInDir);
            this.Controls.Add(this.Label_Merge_Dir);
            this.Controls.Add(this.Btn_Select_Dir);
            this.Controls.Add(this.Btn_SetExcelModule);
            this.Controls.Add(this.Label_ExcelDataModule);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "ExcelMergeTool";
            this.Text = "报表统计工具";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label Label_ExcelDataModule;
        private System.Windows.Forms.Button Btn_SetExcelModule;
        private System.Windows.Forms.Button Btn_Select_Dir;
        private System.Windows.Forms.Label Label_Merge_Dir;
        private System.Windows.Forms.ListView ListView_ExcelsInDir;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button Btn_Excute_Merge;
        private System.Windows.Forms.TextBox Label_Excute_Result;
    }
}

