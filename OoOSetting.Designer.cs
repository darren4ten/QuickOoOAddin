
namespace OoOAddin
{
    partial class OoOSetting
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblSubjectOoO = new System.Windows.Forms.Label();
            this.txtSubject_OoO = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.txtSubject_WFH = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtReceivers_OoO = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtBody_OoO = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtBody_WFH = new System.Windows.Forms.TextBox();
            this.txtReceivers_WFH = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.lblSaveResult = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblSubjectOoO
            // 
            this.lblSubjectOoO.AutoSize = true;
            this.lblSubjectOoO.Location = new System.Drawing.Point(11, 25);
            this.lblSubjectOoO.Name = "lblSubjectOoO";
            this.lblSubjectOoO.Size = new System.Drawing.Size(68, 13);
            this.lblSubjectOoO.TabIndex = 0;
            this.lblSubjectOoO.Text = "OoO Subject";
            // 
            // txtSubject_OoO
            // 
            this.txtSubject_OoO.Location = new System.Drawing.Point(92, 21);
            this.txtSubject_OoO.Name = "txtSubject_OoO";
            this.txtSubject_OoO.Size = new System.Drawing.Size(261, 20);
            this.txtSubject_OoO.TabIndex = 1;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(278, 340);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // txtSubject_WFH
            // 
            this.txtSubject_WFH.Location = new System.Drawing.Point(92, 157);
            this.txtSubject_WFH.Name = "txtSubject_WFH";
            this.txtSubject_WFH.Size = new System.Drawing.Size(261, 20);
            this.txtSubject_WFH.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 160);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "WFH Subject";
            // 
            // txtReceivers_OoO
            // 
            this.txtReceivers_OoO.Location = new System.Drawing.Point(92, 47);
            this.txtReceivers_OoO.Name = "txtReceivers_OoO";
            this.txtReceivers_OoO.Size = new System.Drawing.Size(261, 20);
            this.txtReceivers_OoO.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 51);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "OoO Receivers";
            // 
            // txtBody_OoO
            // 
            this.txtBody_OoO.Location = new System.Drawing.Point(92, 74);
            this.txtBody_OoO.Multiline = true;
            this.txtBody_OoO.Name = "txtBody_OoO";
            this.txtBody_OoO.Size = new System.Drawing.Size(261, 63);
            this.txtBody_OoO.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 77);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "OoO Body";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(11, 212);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "WFH Body";
            // 
            // txtBody_WFH
            // 
            this.txtBody_WFH.Location = new System.Drawing.Point(92, 209);
            this.txtBody_WFH.Multiline = true;
            this.txtBody_WFH.Name = "txtBody_WFH";
            this.txtBody_WFH.Size = new System.Drawing.Size(261, 63);
            this.txtBody_WFH.TabIndex = 11;
            // 
            // txtReceivers_WFH
            // 
            this.txtReceivers_WFH.AccessibleDescription = "use \";\" to split multiple uses.";
            this.txtReceivers_WFH.Location = new System.Drawing.Point(92, 183);
            this.txtReceivers_WFH.Name = "txtReceivers_WFH";
            this.txtReceivers_WFH.Size = new System.Drawing.Size(261, 20);
            this.txtReceivers_WFH.TabIndex = 10;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(7, 186);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "WFH Receivers";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(175, 340);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 13;
            this.button1.Text = "Cancel";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblSaveResult
            // 
            this.lblSaveResult.AutoSize = true;
            this.lblSaveResult.Location = new System.Drawing.Point(56, 298);
            this.lblSaveResult.Name = "lblSaveResult";
            this.lblSaveResult.Size = new System.Drawing.Size(0, 13);
            this.lblSaveResult.TabIndex = 14;
            // 
            // OoOSetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblSaveResult);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtBody_WFH);
            this.Controls.Add(this.txtReceivers_WFH);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtBody_OoO);
            this.Controls.Add(this.txtReceivers_OoO);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtSubject_WFH);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.txtSubject_OoO);
            this.Controls.Add(this.lblSubjectOoO);
            this.Name = "OoOSetting";
            this.Size = new System.Drawing.Size(379, 384);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblSubjectOoO;
        private System.Windows.Forms.TextBox txtSubject_OoO;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox txtSubject_WFH;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtReceivers_OoO;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtBody_OoO;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtBody_WFH;
        private System.Windows.Forms.TextBox txtReceivers_WFH;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label lblSaveResult;
    }
}
