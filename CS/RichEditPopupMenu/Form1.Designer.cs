namespace RichEditPopupMenuExample {
    partial class Form1 {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.components = new System.ComponentModel.Container();
            DevExpress.XtraSpellChecker.OptionsSpelling optionsSpelling1 = new DevExpress.XtraSpellChecker.OptionsSpelling();
            this.richEditControl1 = new DevExpress.XtraRichEdit.RichEditControl();
            this.spellChecker1 = new DevExpress.XtraSpellChecker.SpellChecker();
            this.sharedDictionaryStorage1 = new DevExpress.XtraSpellChecker.SharedDictionaryStorage(this.components);
            this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
            this.panelControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // richEditControl1
            // 
            this.richEditControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richEditControl1.Location = new System.Drawing.Point(0, 26);
            this.richEditControl1.Name = "richEditControl1";
            this.spellChecker1.SetShowSpellCheckMenu(this.richEditControl1, false);
            this.richEditControl1.Size = new System.Drawing.Size(668, 376);
            this.richEditControl1.SpellChecker = this.spellChecker1;
            this.spellChecker1.SetSpellCheckerOptions(this.richEditControl1, optionsSpelling1);
            this.richEditControl1.TabIndex = 0;
            this.richEditControl1.Text = "richEditControl1";
            // 
            // spellChecker1
            // 
            this.spellChecker1.Culture = new System.Globalization.CultureInfo("en-US");
            this.spellChecker1.ParentContainer = null;
            // 
            // panelControl1
            // 
            this.panelControl1.Controls.Add(this.labelControl1);
            this.panelControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelControl1.Location = new System.Drawing.Point(0, 0);
            this.panelControl1.Name = "panelControl1";
            this.panelControl1.Size = new System.Drawing.Size(668, 26);
            this.panelControl1.TabIndex = 1;
            // 
            // labelControl1
            // 
            this.labelControl1.Location = new System.Drawing.Point(12, 5);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(63, 13);
            this.labelControl1.TabIndex = 0;
            this.labelControl1.Text = "labelControl1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(668, 402);
            this.Controls.Add(this.richEditControl1);
            this.Controls.Add(this.panelControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
            this.panelControl1.ResumeLayout(false);
            this.panelControl1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraRichEdit.RichEditControl richEditControl1;
        private DevExpress.XtraSpellChecker.SpellChecker spellChecker1;
        private DevExpress.XtraSpellChecker.SharedDictionaryStorage sharedDictionaryStorage1;
        private DevExpress.XtraEditors.PanelControl panelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl1;
    }
}

