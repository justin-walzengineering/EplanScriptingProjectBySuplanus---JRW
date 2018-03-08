// InsertComment, Version 1.2.0, vom 16.04.2013
//
// Erweitert Eplan Electric P8 um die Möglichkeit Commente einzufügen,
// diese können dann mit dem Commente-Navigator verwaltet werden.
//
// Copyright by Frank Schöneck, 2013
// Last change: Frank Schöneck, 28.02.2013 V1.0.0, Projektbeginn
//					Frank Schöneck, 01.03.2013 V1.1.0, Level, Linetype and Musterlänge als Variable eingesetzt
//                  Frank Schöneck, 16.04.2013 V1.2.0, New tab "Settings" with the possibility to group,
//                      Name changed from "InsertPDFComment" to "InsertComment"
//
// For Eplan Electric P8, ab V2.2
//

// InsertComment, Version 1.2.0, dated 16.04.2013
//
// Extends Eplan Electric P8 to allow comments to be added
// these can then be managed using the Comments Navigator.
//
// Copyright by Frank Schöneck, 2013
// last change: Frank Schöneck, 28.02.2013 V1.0.0, start of project
// Frank Schöneck, 01.03.2013 V1.1.0, level, line type and pattern length used as variables
// Frank Schöneck, 16.04.2013 V1.2.0, New tab "Settings" with the possibility to group,
// Name changed from "InsertPDFComment" to "InsertComment"
//
// for Eplan Electric P8, from V2.2
//

// Menu created at location: Page > Insert comment

using System.Windows.Forms;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using System;
using System.Xml;

public partial class frmInsertComment : System.Windows.Forms.Form
{
	private Button btnOK;
	private Button btnAbort;
	private TabControl tabControl1;
	private Label label1;
	private Label label2;
	private TextBox txtWriter;
	private TextBox txtCommentext;
	private Label label4;
	private TabPage tabComment;
	private DateTimePicker dCreationDate;
	private ComboBox cBStatus;
	private Label label3;
	private TabPage tabSettings;
	private CheckBox chkCommentTextGroup;

	#region Vom Windows Form-Designer generierter Code

	// <summary>
	// Required designer variable.
	// </summary>
	private System.ComponentModel.IContainer components = null;

	// <summary>
	// Clean used resources.
	// </summary>
	// <param name="disposing">True, when managed resources
	// to be deleted; otherwise False. </ param>
	protected override void Dispose(bool disposing)
	{
		if (disposing && (components != null))
		{
			components.Dispose();
		}
		base.Dispose(disposing);
	}

	// <summary>
	// Required method for designer support.
	// The content of the method must not be with the code editor
	// be changed.
	// </summary>
	private void InitializeComponent()
	{
		this.btnOK = new System.Windows.Forms.Button();
		this.btnAbort = new System.Windows.Forms.Button();
		this.tabControl1 = new System.Windows.Forms.TabControl();
		this.tabComment = new System.Windows.Forms.TabPage();
		this.cBStatus = new System.Windows.Forms.ComboBox();
		this.dCreationDate = new System.Windows.Forms.DateTimePicker();
		this.txtCommentext = new System.Windows.Forms.TextBox();
		this.label4 = new System.Windows.Forms.Label();
		this.label3 = new System.Windows.Forms.Label();
		this.label2 = new System.Windows.Forms.Label();
		this.txtWriter = new System.Windows.Forms.TextBox();
		this.label1 = new System.Windows.Forms.Label();
		this.tabSettings = new System.Windows.Forms.TabPage();
		this.chkCommentTextGroup = new System.Windows.Forms.CheckBox();
		this.tabControl1.SuspendLayout();
		this.tabComment.SuspendLayout();
		this.tabSettings.SuspendLayout();
		this.SuspendLayout();
		// 
		// btnOK
		// 
		this.btnOK.Location = new System.Drawing.Point(199, 382);
		this.btnOK.Name = "btnOK";
		this.btnOK.Size = new System.Drawing.Size(110, 25);
		this.btnOK.TabIndex = 0;
		this.btnOK.Text = "OK";
		this.btnOK.UseVisualStyleBackColor = true;
		this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
		// 
		// btnAbort
		// 
		this.btnAbort.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		this.btnAbort.Location = new System.Drawing.Point(328, 382);
		this.btnAbort.Name = "btnAbort";
		this.btnAbort.Size = new System.Drawing.Size(110, 25);
		this.btnAbort.TabIndex = 1;
		this.btnAbort.Text = "Abort";
		this.btnAbort.UseVisualStyleBackColor = true;
		this.btnAbort.Click += new System.EventHandler(this.btnAbort_Click);
		// 
		// tabControl1
		// 
		this.tabControl1.Controls.Add(this.tabComment);
		this.tabControl1.Controls.Add(this.tabSettings);
		this.tabControl1.Location = new System.Drawing.Point(12, 12);
		this.tabControl1.Name = "tabControl1";
		this.tabControl1.SelectedIndex = 0;
		this.tabControl1.Size = new System.Drawing.Size(426, 357);
		this.tabControl1.TabIndex = 2;
		this.tabControl1.TabStop = false;
		// 
		// tabComment
		// 
		this.tabComment.BackColor = System.Drawing.Color.Transparent;
		this.tabComment.Controls.Add(this.cBStatus);
		this.tabComment.Controls.Add(this.dCreationDate);
		this.tabComment.Controls.Add(this.txtCommentext);
		this.tabComment.Controls.Add(this.label4);
		this.tabComment.Controls.Add(this.label3);
		this.tabComment.Controls.Add(this.label2);
		this.tabComment.Controls.Add(this.txtWriter);
		this.tabComment.Controls.Add(this.label1);
		this.tabComment.Location = new System.Drawing.Point(4, 22);
		this.tabComment.Name = "tabComment";
		this.tabComment.Padding = new System.Windows.Forms.Padding(3);
		this.tabComment.Size = new System.Drawing.Size(418, 331);
		this.tabComment.TabIndex = 0;
		this.tabComment.Text = "Comment";
		// 
		// cBStatus
		// 
		this.cBStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.cBStatus.Items.AddRange(new object[] {
            "No status",
            "Accepted",
            "Declined",
            "Canceled",
            "Completed"});
		this.cBStatus.Location = new System.Drawing.Point(201, 94);
		this.cBStatus.MaxDropDownItems = 5;
		this.cBStatus.Name = "cBStatus";
		this.cBStatus.Size = new System.Drawing.Size(211, 21);
		this.cBStatus.TabIndex = 2;
		// 
		// dCreationDate
		// 
		this.dCreationDate.CustomFormat = "MM.dd.yyyy HH:mm:ss";
		this.dCreationDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
		this.dCreationDate.Location = new System.Drawing.Point(9, 94);
		this.dCreationDate.Name = "dCreationDate";
		this.dCreationDate.ShowUpDown = true;
		this.dCreationDate.Size = new System.Drawing.Size(155, 20);
		this.dCreationDate.TabIndex = 1;
		// 
		// txtCommentext
		// 
		this.txtCommentext.Location = new System.Drawing.Point(6, 145);
		this.txtCommentext.Multiline = true;
		this.txtCommentext.Name = "txtCommentext";
		this.txtCommentext.Size = new System.Drawing.Size(406, 175);
		this.txtCommentext.TabIndex = 3;
		// 
		// label4
		// 
		this.label4.AutoSize = true;
		this.label4.Location = new System.Drawing.Point(6, 129);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(80, 13);
		this.label4.TabIndex = 6;
		this.label4.Text = "Comment text:";
		// 
		// label3
		// 
		this.label3.AutoSize = true;
		this.label3.Location = new System.Drawing.Point(198, 78);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(40, 13);
		this.label3.TabIndex = 2;
		this.label3.Text = "Status:";
		// 
		// label2
		// 
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(6, 78);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(90, 13);
		this.label2.TabIndex = 2;
		this.label2.Text = "Creation Date:";
		// 
		// txtWriter
		// 
		this.txtWriter.Location = new System.Drawing.Point(6, 44);
		this.txtWriter.Name = "txtWriter";
		this.txtWriter.Size = new System.Drawing.Size(406, 20);
		this.txtWriter.TabIndex = 0;
		// 
		// label1
		// 
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(6, 28);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(54, 13);
		this.label1.TabIndex = 0;
		this.label1.Text = "Author:";
		// 
		// tabSettings
		// 
		this.tabSettings.BackColor = System.Drawing.Color.Transparent;
		this.tabSettings.Controls.Add(this.chkCommentTextGroup);
		this.tabSettings.Location = new System.Drawing.Point(4, 22);
		this.tabSettings.Name = "tabSettings";
		this.tabSettings.Padding = new System.Windows.Forms.Padding(3);
		this.tabSettings.Size = new System.Drawing.Size(418, 331);
		this.tabSettings.TabIndex = 1;
		this.tabSettings.Text = "Settings";
		// 
		// chkCommentTextGroup
		// 
		this.chkCommentTextGroup.AutoSize = true;
		this.chkCommentTextGroup.Checked = true;
		this.chkCommentTextGroup.CheckState = System.Windows.Forms.CheckState.Checked;
		this.chkCommentTextGroup.Location = new System.Drawing.Point(23, 35);
		this.chkCommentTextGroup.Name = "chkCommentTextGroup";
		this.chkCommentTextGroup.Size = new System.Drawing.Size(246, 17);
		this.chkCommentTextGroup.TabIndex = 0;
		this.chkCommentTextGroup.Text = "Symbol and Comment group text placed";
		this.chkCommentTextGroup.UseVisualStyleBackColor = true;
		// 
		// frmInsertComment
		// 
		this.AcceptButton = this.btnOK;
		this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
		this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.CancelButton = this.btnAbort;
		this.ClientSize = new System.Drawing.Size(450, 419);
		this.Controls.Add(this.tabControl1);
		this.Controls.Add(this.btnAbort);
		this.Controls.Add(this.btnOK);
		this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		this.MaximizeBox = false;
		this.MinimizeBox = false;
		this.Name = "frmInsertComment";
		this.ShowInTaskbar = false;
		this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
		this.Text = "Insert Comment";
		this.Load += new System.EventHandler(this.frmInsertComment_Load);
		this.tabControl1.ResumeLayout(false);
		this.tabComment.ResumeLayout(false);
		this.tabComment.PerformLayout();
		this.tabSettings.ResumeLayout(false);
		this.tabSettings.PerformLayout();
		this.ResumeLayout(false);

	}

	public frmInsertComment()
	{
		InitializeComponent();
	}

    #endregion

    //Create menu item
    [DeclareMenu()]
	public void Comment_Menu()
	{
		Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
		oMenu.AddMenuItem("Comment insert", "DialogInsertComment", "Comment insert is running", 35381, 0, false, false);
	}

    //Action to call the form
    [DeclareAction("DialogInsertComment")]
	public void DialogInsertComment_action()
	{
		frmInsertComment frm = new frmInsertComment();
		frm.ShowDialog();
		return;
	}

	//Form_Load
	private void frmInsertComment_Load(object sender, System.EventArgs e)
	{
		Text = "insert";

        //Preset the author
#if DEBUG
        txtWriter.Text = "Justin R. Walz";
#else

		UserRights oUserRights = new UserRights();
		txtWriter.Text = oUserRights.GetUser();
#endif
		//Status preset
		cBStatus.Text = "no Status";

        //Input to comment text
        txtCommentext.Select();
	}

	//Button Abort Click
	private void btnAbort_Click(object sender, System.EventArgs e)
	{
		Close();
	}

	//Button OK Click
	private void btnOK_Click(object sender, System.EventArgs e)
	{
        //Level of Comment(A411 = 519(EPLAN519, Graphics.Comment))

        string sLevel = "519";

        //Linetype of Comment(A412 = L(Layer) / 0(solid) / 41(~~~~~))

        string sLinetype = "41";

        //Pattern length of the comment(A415 = L(layer) / -1.5(1.50 mm) / -32(32.00 mm))
        string sMusterlänge = "-1.5";

        //Comment color (A413 = 0(black) / 1(red) / 2(yellow) / 3(light green) / 4(light blue) / 5(dark blue) / 6(violet) / 8(white) / 40(orange))
        string sColor = "40";

        //Author (A2521)
        string sAuthor = txtWriter.Text;

		//Creation Date (A2524)
		string sCreationDate = DateTimeToUnixTimestamp(dCreationDate.Value).ToString(); // Convert DateTime value to Unix Timestamp format

        //Comment Text (A511)
        string sCommenttext = txtCommentext.Text;
		if (sCommenttext.EndsWith(Environment.NewLine)) //Comment can not end with a line break
 
        { sCommenttext = sCommenttext.Substring(0, sCommenttext.Length - 2); }
		sCommenttext = sCommenttext.Replace(Environment.NewLine, "&#10;"); //Comment Convert line break
        sCommenttext = "??_??@" + sCommenttext; //Comment MultiLanguage String
		if (!sCommenttext.EndsWith(";")) //Comment must be with ";" end up
        { sCommenttext += ";"; }

        //Status, (A2527 = 0 (no status) / 1 (Accepted) / 2 (Declined) / 3 (Canceled) / 4 (Finished))
        string sStatus = cBStatus.SelectedIndex.ToString();

        //Path and filename of the Temp.file
        string sTempFile;
#if DEBUG
		sTempFile = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\tmpInsertComment.ema";
#else
		sTempFile = PathMap.SubstitutePath(@"$(TMP)") + @"\tmpInsertComment.ema";
#endif
		XmlWriterSettings settings = new XmlWriterSettings();
		settings.Indent = true;
		XmlWriter writer = XmlWriter.Create(sTempFile, settings);

        //Macro file content
        writer.WriteRaw("\n<EplanPxfRoot Name=\"#Comment\" Type=\"WindowMacro\" Version=\"2.2.6360\" PxfVersion=\"1.23\" SchemaVersion=\"1.7.6360\" Source=\"\" SourceProject=\"\" Description=\"\" ConfigurationFlags=\"0\" NumMainObjects=\"0\" NumProjectSteps=\"0\" NumMDSteps=\"0\" Custompagescaleused=\"true\" StreamSchema=\"EBitmap,BaseTypes,2,1,2;EPosition3D,BaseTypes,0,3,2;ERay3D,BaseTypes,0,4,2;EStreamableVector,BaseTypes,0,5,2;DMNCDataSet,TrDMProject,1,20,2;DMNCDataSetVector,TrDMProject,1,21,2;DMPlaceHolderRuleData,TrDMProject,0,22,2;Arc3d@W3D,W3dBaseGeometry,0,36,2;Box3d@W3D,W3dBaseGeometry,0,37,2;Circle3d@W3D,W3dBaseGeometry,0,38,2;Color@W3D,W3dBaseGeometry,0,39,2;ContourPlane3d@W3D,W3dBaseGeometry,0,40,2;CTexture@W3D,W3dBaseGeometry,0,41,2;CTextureMap@W3D,W3dBaseGeometry,0,42,2;Line3d@W3D,W3dBaseGeometry,0,43,2;Linetype@W3D,W3dBaseGeometry,0,44,2;Material@W3D,W3dBaseGeometry,3,45,2;Path3d@W3D,W3dBaseGeometry,0,46,2;Mesh3dX@W3D,W3dMeshModeller,2,47,2;MeshBox@W3D,W3dMeshModeller,5,48,2;MeshMate@W3D,W3dMeshModeller,7,49,2;MeshMateFace@W3D,W3dMeshModeller,1,50,2;MeshMateGrid@W3D,W3dMeshModeller,8,51,2;MeshMateGridLine@W3D,W3dMeshModeller,1,52,2;MeshMateLine@W3D,W3dMeshModeller,1,53,2;MeshText3dX@W3D,W3dMeshModeller,0,55,2;BaseTextLine@W3D,W3dMeshModeller,2,56,2;Mesh3d@W3D,W3dMeshModeller,8,57,2;MeshEdge3d@W3D,W3dMeshModeller,0,58,2;MeshFace3d@W3D,W3dMeshModeller,2,59,2;MeshPoint3d@W3D,W3dMeshModeller,1,60,2;MeshPolygon3d@W3D,W3dMeshModeller,1,61,2;MeshSimpleTextureTriangle3d@W3D,W3dMeshModeller,2,62,2;MeshSimpleTriangle3d@W3D,W3dMeshModeller,1,63,2;MeshTriangle3d@W3D,W3dMeshModeller,2,64,2;MeshTriangleFace3d@W3D,W3dMeshModeller,0,65,2;MeshTriangleFaceEdge3D@W3D,W3dMeshModeller,0,66,2\">");
		writer.WriteRaw("\n <MacroVariant MacroFuncType=\"1\" VariantId=\"0\" ReferencePoint=\"64/248/0\" Version=\"2.2.6360\" PxfVersion=\"1.23\" SchemaVersion=\"1.7.6360\" Source=\"\" SourceProject=\"\" Description=\"\" ConfigurationFlags=\"0\" DocumentType=\"1\" Customgost=\"0\">");
		writer.WriteRaw("\n  <O4 Build=\"6360\" A1=\"4/18\" A3=\"0\" A13=\"0\" A14=\"0\" A47=\"1\" A48=\"1362057551\" A50=\"1\" A59=\"1\" A404=\"1\" A405=\"64\" A406=\"0\" A407=\"0\" A431=\"1\" A1101=\"17\" A1102=\"\" A1103=\"\">");

        //if you want to group
        if (chkCommentTextGroup.Checked == true)
		{
			writer.WriteRaw("\n  <O26 Build=\"6360\" A1=\"26/128740\" A3=\"0\" A13=\"0\" A14=\"0\" A404=\"1\" A405=\"64\" A406=\"0\" A407=\"0\" A431=\"1\">");
		}
        //Characteristics Comment
        writer.WriteRaw("\n  <O165 Build=\"6360\" A1=\"165/128741\" A3=\"0\" A13=\"0\" A14=\"0\" A404=\"1\" A405=\"64\" A406=\"0\" A407=\"0\" A411=\"" + sLevel + "\" A412=\"" + sLinetype + "\" A413=\"" + sColor + "\" A414=\"0.352777238812552\" A415=\"" + sMusterlänge + "\" A416=\"0\" A501=\"64/248\" A503=\"0\" A504=\"0\" A506=\"22\" A511=\"" + sCommenttext + "\" A2521=\"" + sAuthor + "\" A2522=\"\" A2523=\"\" A2524=\"" + sCreationDate + "\" A2525=\"" + sCreationDate + "\" A2526=\"2\" A2527=\"" + sStatus + "\" A2528=\"0\" A2529=\"0\" A2531=\"0\" A2532=\"0\" A2533=\"64/248;70.349110320284/254.349110320285\" A2534=\"2\" A2539=\"0\" A2540=\"0\">");
		writer.WriteRaw("\n  <S54x505 A961=\"L\" A962=\"L\" A963=\"0\" A964=\"L\" A965=\"0\" A966=\"0\" A967=\"0\" A968=\"0\" A969=\"0\" A4000=\"L\" A4001=\"L\" A4013=\"0\"/>");
		writer.WriteRaw("\n  </O165>");

        //Characteristics Text
        writer.WriteRaw("\n  <O30 Build=\"6360\" A1=\"30/128742\" A3=\"0\" A13=\"0\" A14=\"0\" A404=\"1\" A405=\"64\" A406=\"0\" A407=\"0\" A411=\"" + sLevel + "\" A412=\"L\" A413=\"L\" A414=\"L\" A415=\"L\" A416=\"0\" A501=\"72/248\" A503=\"0\" A504=\"0\" A506=\"0\" A511=\"" + sCommenttext + "\">");
		writer.WriteRaw("\n  <S54x505 A961=\"L\" A962=\"L\" A963=\"0\" A964=\"L\" A965=\"0\" A966=\"0\" A967=\"0\" A968=\"0\" A969=\"0\" A4000=\"L\" A4001=\"L\" A4013=\"0\"/>");
		writer.WriteRaw("\n  </O30>");

        //Ff you want to group
        if (chkCommentTextGroup.Checked == true)
		{
			writer.WriteRaw("\n  </O26>");
		}
		writer.WriteRaw("\n  <O37 Build=\"6360\" A1=\"37/128743\" A3=\"1\" A13=\"0\" A14=\"0\" A404=\"1\" A405=\"64\" A406=\"0\" A407=\"0\" A682=\"1\" A683=\"26/128740\" A684=\"0\" A687=\"8\" A688=\"2\" A689=\"-1\" A690=\"-1\" A691=\"0\" A693=\"1\" A792=\"0\" A793=\"0\" A794=\"0\" A1261=\"0\" A1262=\"44\" A1263=\"0\" A1631=\"0/-6.34911032028501\" A1632=\"8/0\">");
		writer.WriteRaw("\n  <S109x692 Build=\"6360\" A3=\"0\" A13=\"0\" A14=\"0\" R1906=\"165/128741\"/>");
		writer.WriteRaw("\n  <S109x692 Build=\"6360\" A3=\"0\" A13=\"0\" A14=\"0\" R1906=\"30/128742\"/>");
		writer.WriteRaw("\n  <S40x1201 A762=\"64/254.349110320285\">");
		writer.WriteRaw("\n  <S39x761 A751=\"1\" A752=\"0\" A753=\"0\" A754=\"1\"/>");
		writer.WriteRaw("\n  </S40x1201>");
		writer.WriteRaw("\n  <S89x5 Build=\"6360\" A3=\"0\" A4=\"1\" R7=\"37/128743\" A13=\"0\" A14=\"0\" A404=\"9\" A405=\"64\" A406=\"0\" A407=\"0\" A411=\"308\" A412=\"L\" A413=\"L\" A414=\"L\" A415=\"L\" A416=\"0\" A1651=\"0/-6.34911032028501\" A1652=\"8/0\" A1653=\"0\" A1654=\"0\" A1655=\"0\" A1656=\"0\" A1657=\"0\"/>");
		writer.WriteRaw("\n  </O37>");
		writer.WriteRaw("\n  </O4>");
		writer.WriteRaw("\n </MacroVariant>");
		writer.WriteRaw("\n</EplanPxfRoot>");

		// Write the XML to file and close the writer.
		writer.Flush();
		writer.Close();

        //Insert macro
#if DEBUG
        MessageBox.Show(sTempFile);
#else
		CommandLineInterpreter oCli = new CommandLineInterpreter();
		ActionCallingContext oAcc = new ActionCallingContext();
		oAcc.AddParameter("Name", "XMIaInsertMacro");
		oAcc.AddParameter("filename", sTempFile);
		oAcc.AddParameter("variant", "0");
		oCli.Execute("XGedStartInteractionAction", oAcc);
#endif

        //Break up
        Close();
		return;
	}

    // <summary>
    // Methods to convert Unix time stamp to DateTime
    // </ summary>
    // <param name = "_ UnixTimeStamp"> Unix time stamp to convert </ param>
    // <returns> Return DateTime </ returns>
    //
    // Example:
    // convert given Unix time stamp to DateTime and update dateTimePicker value
    // dateTimePicker1.Value = UnixTimestampToDateTime (181913235);
    public DateTime UnixTimestampToDateTime(long _UnixTimeStamp)
	{
		return (new DateTime(1970, 1, 1, 1, 0, 0)).AddSeconds(_UnixTimeStamp);

	}

    // <summary>
    // Methods to convert DateTime to Unix time stamp
    // </ summary>
    // <param name = "_ UnixTimeStamp"> Unix time stamp to convert </ param>
    // <returns> Return Unix time stamp as long type </ returns>
    //
    // Example:
    // Convert current DateTime to Unix time stamp
    // Console.WriteLine ("Current unix time stamp is:" + DateTimeToUnixTimestamp (DateTime.Now) .ToString ());
    public long DateTimeToUnixTimestamp(DateTime _DateTime)
	{
		TimeSpan _UnixTimeSpan = (_DateTime - new DateTime(1970, 1, 1, 1, 0, 0));
		return (long)_UnixTimeSpan.TotalSeconds;
	}

}