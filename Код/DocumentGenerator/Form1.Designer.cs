
namespace DocumentGenerator
{
    partial class Program1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Program1));
            this.Generate_Button = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.СhooseData = new System.Windows.Forms.Button();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ChooseFolderTemplates = new System.Windows.Forms.Button();
            this.SaveButton = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // Generate_Button
            // 
            this.Generate_Button.BackColor = System.Drawing.SystemColors.HotTrack;
            this.Generate_Button.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.Generate_Button.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.Generate_Button.Location = new System.Drawing.Point(226, 162);
            this.Generate_Button.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Generate_Button.Name = "Generate_Button";
            this.Generate_Button.Size = new System.Drawing.Size(137, 49);
            this.Generate_Button.TabIndex = 1;
            this.Generate_Button.Text = "Выгрузить";
            this.Generate_Button.UseVisualStyleBackColor = false;
            this.Generate_Button.Click += new System.EventHandler(this.Generate_Button_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 219);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(559, 37);
            this.progressBar1.TabIndex = 2;
            this.progressBar1.Visible = false;
            // 
            // СhooseData
            // 
            this.СhooseData.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.СhooseData.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.СhooseData.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.СhooseData.Location = new System.Drawing.Point(12, 13);
            this.СhooseData.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.СhooseData.Name = "СhooseData";
            this.СhooseData.Size = new System.Drawing.Size(136, 37);
            this.СhooseData.TabIndex = 3;
            this.СhooseData.Text = "Выбрать Бланк";
            this.СhooseData.UseVisualStyleBackColor = false;
            this.СhooseData.Click += new System.EventHandler(this.СhooseData_Click);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // ChooseFolderTemplates
            // 
            this.ChooseFolderTemplates.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.ChooseFolderTemplates.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.ChooseFolderTemplates.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ChooseFolderTemplates.Location = new System.Drawing.Point(12, 58);
            this.ChooseFolderTemplates.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.ChooseFolderTemplates.Name = "ChooseFolderTemplates";
            this.ChooseFolderTemplates.Size = new System.Drawing.Size(136, 53);
            this.ChooseFolderTemplates.TabIndex = 6;
            this.ChooseFolderTemplates.Text = "Выбрать Папку с Шаблонами";
            this.ChooseFolderTemplates.UseVisualStyleBackColor = false;
            this.ChooseFolderTemplates.Click += new System.EventHandler(this.ChooseFolderTemplates_Click);
            // 
            // SaveButton
            // 
            this.SaveButton.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.SaveButton.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.SaveButton.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.SaveButton.Location = new System.Drawing.Point(12, 119);
            this.SaveButton.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(136, 37);
            this.SaveButton.TabIndex = 7;
            this.SaveButton.Text = "Сохранить в ";
            this.SaveButton.UseVisualStyleBackColor = false;
            this.SaveButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // fff
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.Disable;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(583, 269);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.ChooseFolderTemplates);
            this.Controls.Add(this.СhooseData);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.Generate_Button);
            this.Font = new System.Drawing.Font("Microsoft YaHei", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "fff";
            this.Text = "Генератор Актов BetaV1.0";
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button Generate_Button;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button СhooseData;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.Button ChooseFolderTemplates;
        private System.Windows.Forms.Button SaveButton;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}

