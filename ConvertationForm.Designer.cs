namespace Convertation
{
    partial class ConvertationForm
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
            this.Process = new System.Windows.Forms.Button();
            this.LogBox = new System.Windows.Forms.TextBox();
            this.SetInput = new System.Windows.Forms.Button();
            this.SetOutput = new System.Windows.Forms.Button();
            this.InputDirectory = new System.Windows.Forms.TextBox();
            this.OutputDirectory = new System.Windows.Forms.TextBox();
            this.FolderDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // Process
            // 
            this.Process.Location = new System.Drawing.Point(487, 12);
            this.Process.Name = "Process";
            this.Process.Size = new System.Drawing.Size(75, 23);
            this.Process.TabIndex = 0;
            this.Process.Text = "Process";
            this.Process.UseVisualStyleBackColor = true;
            this.Process.Click += new System.EventHandler(this.Button1_Click);
            // 
            // LogBox
            // 
            this.LogBox.Location = new System.Drawing.Point(12, 74);
            this.LogBox.Multiline = true;
            this.LogBox.Name = "LogBox";
            this.LogBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.LogBox.Size = new System.Drawing.Size(550, 327);
            this.LogBox.TabIndex = 1;
            // 
            // SetInput
            // 
            this.SetInput.Location = new System.Drawing.Point(13, 12);
            this.SetInput.Name = "SetInput";
            this.SetInput.Size = new System.Drawing.Size(89, 23);
            this.SetInput.TabIndex = 2;
            this.SetInput.Text = "Input folder";
            this.SetInput.UseVisualStyleBackColor = true;
            this.SetInput.Click += new System.EventHandler(this.SetInput_Click);
            // 
            // SetOutput
            // 
            this.SetOutput.Location = new System.Drawing.Point(12, 42);
            this.SetOutput.Name = "SetOutput";
            this.SetOutput.Size = new System.Drawing.Size(90, 23);
            this.SetOutput.TabIndex = 3;
            this.SetOutput.Text = "Output folder";
            this.SetOutput.UseVisualStyleBackColor = true;
            this.SetOutput.Click += new System.EventHandler(this.SetOutput_Click);
            // 
            // InputDirectory
            // 
            this.InputDirectory.Location = new System.Drawing.Point(108, 13);
            this.InputDirectory.Name = "InputDirectory";
            this.InputDirectory.Size = new System.Drawing.Size(370, 20);
            this.InputDirectory.TabIndex = 4;
            // 
            // OutputDirectory
            // 
            this.OutputDirectory.Location = new System.Drawing.Point(108, 42);
            this.OutputDirectory.Name = "OutputDirectory";
            this.OutputDirectory.Size = new System.Drawing.Size(370, 20);
            this.OutputDirectory.TabIndex = 5;
            // 
            // ConvertationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(574, 413);
            this.Controls.Add(this.OutputDirectory);
            this.Controls.Add(this.InputDirectory);
            this.Controls.Add(this.SetOutput);
            this.Controls.Add(this.SetInput);
            this.Controls.Add(this.LogBox);
            this.Controls.Add(this.Process);
            this.MinimumSize = new System.Drawing.Size(350, 350);
            this.Name = "ConvertationForm";
            this.RightToLeftLayout = true;
            this.Text = "Convertation to html";
            this.Load += new System.EventHandler(this.ConvertationForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Process;
        private System.Windows.Forms.TextBox LogBox;
        private System.Windows.Forms.Button SetInput;
        private System.Windows.Forms.Button SetOutput;
        private System.Windows.Forms.TextBox InputDirectory;
        private System.Windows.Forms.TextBox OutputDirectory;
        private System.Windows.Forms.FolderBrowserDialog FolderDialog;
    }
}

