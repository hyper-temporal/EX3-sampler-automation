
namespace ExPlus
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.SelectedPresetNum = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.SelectedPreset = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(178, 57);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(57, 35);
            this.button1.TabIndex = 22;
            this.button1.Text = "COPY";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // SelectedPresetNum
            // 
            this.SelectedPresetNum.AutoSize = true;
            this.SelectedPresetNum.Location = new System.Drawing.Point(15, 32);
            this.SelectedPresetNum.Name = "SelectedPresetNum";
            this.SelectedPresetNum.Size = new System.Drawing.Size(116, 15);
            this.SelectedPresetNum.TabIndex = 37;
            this.SelectedPresetNum.Text = "Selected Preset Num";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(178, 96);
            this.button4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(57, 35);
            this.button4.TabIndex = 38;
            this.button4.Text = "PASTE";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(15, 57);
            this.button6.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(79, 74);
            this.button6.TabIndex = 40;
            this.button6.Text = "MERGE";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click_1);
            // 
            // SelectedPreset
            // 
            this.SelectedPreset.AutoSize = true;
            this.SelectedPreset.Location = new System.Drawing.Point(13, 12);
            this.SelectedPreset.Name = "SelectedPreset";
            this.SelectedPreset.Size = new System.Drawing.Size(121, 15);
            this.SelectedPreset.TabIndex = 36;
            this.SelectedPreset.Text = "Selected Preset Name";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(280, 149);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.SelectedPresetNum);
            this.Controls.Add(this.SelectedPreset);
            this.Controls.Add(this.button1);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label SelectedPresetNum;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Label SelectedPreset;
    }
}

