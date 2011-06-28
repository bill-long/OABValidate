namespace OABValidate
{
	partial class Form1
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components;

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

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBoxGCandOAB = new System.Windows.Forms.GroupBox();
            this.listViewChooseOAB = new System.Windows.Forms.ListView();
            this.labelChooseOAB = new System.Windows.Forms.Label();
            this.buttonGetOABs = new System.Windows.Forms.Button();
            this.textBoxGC = new System.Windows.Forms.TextBox();
            this.labelGC = new System.Windows.Forms.Label();
            this.buttonGo = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.textBoxObjectsProcessed = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxObjectsFound = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxProblemObjects = new System.Windows.Forms.TextBox();
            this.groupBoxGCandOAB.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBoxGCandOAB
            // 
            this.groupBoxGCandOAB.Controls.Add(this.listViewChooseOAB);
            this.groupBoxGCandOAB.Controls.Add(this.labelChooseOAB);
            this.groupBoxGCandOAB.Controls.Add(this.buttonGetOABs);
            this.groupBoxGCandOAB.Controls.Add(this.textBoxGC);
            this.groupBoxGCandOAB.Controls.Add(this.labelGC);
            this.groupBoxGCandOAB.Location = new System.Drawing.Point(12, 12);
            this.groupBoxGCandOAB.Name = "groupBoxGCandOAB";
            this.groupBoxGCandOAB.Size = new System.Drawing.Size(415, 157);
            this.groupBoxGCandOAB.TabIndex = 0;
            this.groupBoxGCandOAB.TabStop = false;
            this.groupBoxGCandOAB.Text = "GC and OAB selection";
            // 
            // listViewChooseOAB
            // 
            this.listViewChooseOAB.CheckBoxes = true;
            this.listViewChooseOAB.Location = new System.Drawing.Point(9, 63);
            this.listViewChooseOAB.Name = "listViewChooseOAB";
            this.listViewChooseOAB.Size = new System.Drawing.Size(394, 81);
            this.listViewChooseOAB.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listViewChooseOAB.TabIndex = 4;
            this.listViewChooseOAB.UseCompatibleStateImageBehavior = false;
            this.listViewChooseOAB.View = System.Windows.Forms.View.List;
            // 
            // labelChooseOAB
            // 
            this.labelChooseOAB.AutoSize = true;
            this.labelChooseOAB.Location = new System.Drawing.Point(6, 47);
            this.labelChooseOAB.Name = "labelChooseOAB";
            this.labelChooseOAB.Size = new System.Drawing.Size(128, 13);
            this.labelChooseOAB.TabIndex = 3;
            this.labelChooseOAB.Text = "Choose OABs to validate:";
            // 
            // buttonGetOABs
            // 
            this.buttonGetOABs.Location = new System.Drawing.Point(328, 22);
            this.buttonGetOABs.Name = "buttonGetOABs";
            this.buttonGetOABs.Size = new System.Drawing.Size(75, 23);
            this.buttonGetOABs.TabIndex = 2;
            this.buttonGetOABs.Text = "Get OABs";
            this.buttonGetOABs.UseVisualStyleBackColor = true;
            this.buttonGetOABs.Click += new System.EventHandler(this.buttonGetOABs_Click);
            // 
            // textBoxGC
            // 
            this.textBoxGC.Location = new System.Drawing.Point(37, 24);
            this.textBoxGC.Name = "textBoxGC";
            this.textBoxGC.Size = new System.Drawing.Size(285, 20);
            this.textBoxGC.TabIndex = 1;
            // 
            // labelGC
            // 
            this.labelGC.AutoSize = true;
            this.labelGC.Location = new System.Drawing.Point(6, 27);
            this.labelGC.Name = "labelGC";
            this.labelGC.Size = new System.Drawing.Size(25, 13);
            this.labelGC.TabIndex = 0;
            this.labelGC.Text = "GC:";
            // 
            // buttonGo
            // 
            this.buttonGo.Location = new System.Drawing.Point(277, 172);
            this.buttonGo.Name = "buttonGo";
            this.buttonGo.Size = new System.Drawing.Size(150, 66);
            this.buttonGo.TabIndex = 2;
            this.buttonGo.Text = "Go!";
            this.buttonGo.UseVisualStyleBackColor = true;
            this.buttonGo.Click += new System.EventHandler(this.buttonGo_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.richTextBox1.Location = new System.Drawing.Point(12, 247);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(414, 125);
            this.richTextBox1.TabIndex = 4;
            this.richTextBox1.Text = "";
            this.richTextBox1.WordWrap = false;
            // 
            // textBoxObjectsProcessed
            // 
            this.textBoxObjectsProcessed.Location = new System.Drawing.Point(113, 195);
            this.textBoxObjectsProcessed.Name = "textBoxObjectsProcessed";
            this.textBoxObjectsProcessed.ReadOnly = true;
            this.textBoxObjectsProcessed.Size = new System.Drawing.Size(158, 20);
            this.textBoxObjectsProcessed.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 198);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Objects processed:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(34, 175);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Objects found:";
            // 
            // textBoxObjectsFound
            // 
            this.textBoxObjectsFound.Location = new System.Drawing.Point(113, 172);
            this.textBoxObjectsFound.Name = "textBoxObjectsFound";
            this.textBoxObjectsFound.ReadOnly = true;
            this.textBoxObjectsFound.Size = new System.Drawing.Size(158, 20);
            this.textBoxObjectsFound.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(25, 221);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(85, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Problem objects:";
            // 
            // textBoxProblemObjects
            // 
            this.textBoxProblemObjects.Location = new System.Drawing.Point(113, 218);
            this.textBoxProblemObjects.Name = "textBoxProblemObjects";
            this.textBoxProblemObjects.ReadOnly = true;
            this.textBoxProblemObjects.Size = new System.Drawing.Size(158, 20);
            this.textBoxProblemObjects.TabIndex = 10;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(441, 384);
            this.Controls.Add(this.textBoxProblemObjects);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxObjectsFound);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxObjectsProcessed);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.buttonGo);
            this.Controls.Add(this.groupBoxGCandOAB);
            this.Name = "Form1";
            this.Text = "OABValidate 2";
            this.groupBoxGCandOAB.ResumeLayout(false);
            this.groupBoxGCandOAB.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

        private System.Windows.Forms.GroupBox groupBoxGCandOAB;
		private System.Windows.Forms.Label labelChooseOAB;
		private System.Windows.Forms.Button buttonGetOABs;
		private System.Windows.Forms.TextBox textBoxGC;
        private System.Windows.Forms.Label labelGC;
        private System.Windows.Forms.Button buttonGo;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.ListView listViewChooseOAB;
        private System.Windows.Forms.TextBox textBoxObjectsProcessed;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxObjectsFound;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxProblemObjects;
	}
}

