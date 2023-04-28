using SACO.SolidWorks.ComptProps.Models;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SACO.SolidWorks.ComptProps
{
    [ProgId(PropsAddinMain.SWTASKPANE_PROGID)]
    public partial class TaskpaneHost : UserControl
    {

        #region General Properties

        private List<string> DesignersInitals = new List<string>();
        private List<string> CheckedByInitials = new List<string>();
        private List<string> SignedBy = new List<string>();
        private List<string> MaterialInfo = new List<string>();
        private List<string> FinishIngo = new List<string>();
        private List<string> MaterialThickness = new List<string>();
        private List<string> ProductGroup = new List<string>();
        private List<string> Revision = new List<string>();

        private SacoComponentPropertiesModel PropertiesModel = new SacoComponentPropertiesModel();

        #endregion

        #region SW Properties

        private SldWorks swApp;
        private ModelDoc2 swModel;
        private ModelDocExtension modelExt;
        private CustomPropertyManager customProps;


        private int mSwCookie;

        //static DSldWorksEvents_FileOpenNotify2EventHandler fileOpenNotify;

        static FileOpenNotifyHandler fileOpenNotify;



        #endregion







        #region Control Properties

        private Button btnRefreshProps;
        private GroupBox grpGeneral;
        private TextBox textTitle;
        private Label label5;
        private Label label4;
        private Label label3;
        private Label label2;
        private Label Title;
        private TextBox textProject;
        private ComboBox dropThickness;
        private ComboBox dropMaterial;
        private ComboBox dropFinish;
        private GroupBox groupUncontrolled;
        private Label label10;
        private Label label9;
        private Label label8;
        private Label label7;
        private Label label6;
        private Label label1;
        private ComboBox dropStatusUncontrolled;
        private ComboBox dropCheckedThreeUncontrolled;
        private ComboBox dropCheckedTwoUncontrolled;
        private ComboBox dropUncontrolledDrawnBy;
        private ComboBox dropCheckedOneUncontrolled;
        private ComboBox dropUncontrolledRevision;
        private DateTimePicker dateUncontrolledDateCheckedThree;
        private DateTimePicker dateUncontrolledDateCheckedTwo;
        private DateTimePicker dateUncontolledCheckedOne;
        private DateTimePicker dateUncontrolledDrawn;
        private GroupBox groupControlled;
        private DateTimePicker dateControlledDateCheckedThree;
        private DateTimePicker dateControlledDateCheckedTwo;
        private DateTimePicker dateContolledCheckedOne;
        private DateTimePicker dateControlledDrawn;
        private ComboBox dropStatusControlled;
        private ComboBox dropCheckedThreeControlled;
        private ComboBox dropCheckedTwoControlled;
        private ComboBox dropDrawnByControlled;
        private ComboBox dropCheckedOneControlled;
        private ComboBox dropControlledRevision;
        private Label label11;
        private Label label12;
        private Label label13;
        private Label label14;
        private Label label15;
        private Label label16;
        private DateTimePicker dateSignedBy;
        private ComboBox dropSignedBy;
        private Label label17;
        private Label label18;
        private ComboBox dropProductGroup;
        private Label label19;
        private Label lblSacoComptHeader;
        private Button btnSave;
        private Button btnUncontrolledCheckedDateRemove;
        private TextBox textNotesComments;

        #endregion

        public TaskpaneHost()
        {
            InitializeComponent();

            this.dateUncontrolledDrawn.ValueChanged += delegate (object sender, EventArgs e) { DateTimePicker_ValueChanged(sender, e, "Drawn_Date_Uncontrolled"); }; ;

            GetLists();

            //swModel = swApp.ActiveDoc;
            fileOpenNotify = new FileOpenNotifyHandler();
            fileOpenNotify.Init(swApp);

        }


        private void InitializeComponent()
        {
            this.btnRefreshProps = new System.Windows.Forms.Button();
            this.grpGeneral = new System.Windows.Forms.GroupBox();
            this.dropProductGroup = new System.Windows.Forms.ComboBox();
            this.label19 = new System.Windows.Forms.Label();
            this.textProject = new System.Windows.Forms.TextBox();
            this.dropThickness = new System.Windows.Forms.ComboBox();
            this.dropMaterial = new System.Windows.Forms.ComboBox();
            this.dropFinish = new System.Windows.Forms.ComboBox();
            this.textTitle = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Title = new System.Windows.Forms.Label();
            this.groupUncontrolled = new System.Windows.Forms.GroupBox();
            this.btnUncontrolledCheckedDateRemove = new System.Windows.Forms.Button();
            this.dateUncontrolledDateCheckedThree = new System.Windows.Forms.DateTimePicker();
            this.dateUncontrolledDateCheckedTwo = new System.Windows.Forms.DateTimePicker();
            this.dateUncontolledCheckedOne = new System.Windows.Forms.DateTimePicker();
            this.dateUncontrolledDrawn = new System.Windows.Forms.DateTimePicker();
            this.dropStatusUncontrolled = new System.Windows.Forms.ComboBox();
            this.dropCheckedThreeUncontrolled = new System.Windows.Forms.ComboBox();
            this.dropCheckedTwoUncontrolled = new System.Windows.Forms.ComboBox();
            this.dropUncontrolledDrawnBy = new System.Windows.Forms.ComboBox();
            this.dropCheckedOneUncontrolled = new System.Windows.Forms.ComboBox();
            this.dropUncontrolledRevision = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupControlled = new System.Windows.Forms.GroupBox();
            this.dateSignedBy = new System.Windows.Forms.DateTimePicker();
            this.dropSignedBy = new System.Windows.Forms.ComboBox();
            this.label17 = new System.Windows.Forms.Label();
            this.dateControlledDateCheckedThree = new System.Windows.Forms.DateTimePicker();
            this.dateControlledDateCheckedTwo = new System.Windows.Forms.DateTimePicker();
            this.dateContolledCheckedOne = new System.Windows.Forms.DateTimePicker();
            this.dateControlledDrawn = new System.Windows.Forms.DateTimePicker();
            this.dropStatusControlled = new System.Windows.Forms.ComboBox();
            this.dropCheckedThreeControlled = new System.Windows.Forms.ComboBox();
            this.dropCheckedTwoControlled = new System.Windows.Forms.ComboBox();
            this.dropDrawnByControlled = new System.Windows.Forms.ComboBox();
            this.dropCheckedOneControlled = new System.Windows.Forms.ComboBox();
            this.dropControlledRevision = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.textNotesComments = new System.Windows.Forms.TextBox();
            this.lblSacoComptHeader = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.grpGeneral.SuspendLayout();
            this.groupUncontrolled.SuspendLayout();
            this.groupControlled.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnRefreshProps
            // 
            this.btnRefreshProps.BackColor = System.Drawing.Color.DarkRed;
            this.btnRefreshProps.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnRefreshProps.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnRefreshProps.Location = new System.Drawing.Point(230, 3);
            this.btnRefreshProps.Name = "btnRefreshProps";
            this.btnRefreshProps.Size = new System.Drawing.Size(95, 44);
            this.btnRefreshProps.TabIndex = 0;
            this.btnRefreshProps.Text = "Get Properties";
            this.btnRefreshProps.UseVisualStyleBackColor = false;
            this.btnRefreshProps.Click += new System.EventHandler(this.button1_Click);
            // 
            // grpGeneral
            // 
            this.grpGeneral.BackColor = System.Drawing.Color.White;
            this.grpGeneral.Controls.Add(this.dropProductGroup);
            this.grpGeneral.Controls.Add(this.label19);
            this.grpGeneral.Controls.Add(this.textProject);
            this.grpGeneral.Controls.Add(this.dropThickness);
            this.grpGeneral.Controls.Add(this.dropMaterial);
            this.grpGeneral.Controls.Add(this.dropFinish);
            this.grpGeneral.Controls.Add(this.textTitle);
            this.grpGeneral.Controls.Add(this.label5);
            this.grpGeneral.Controls.Add(this.label4);
            this.grpGeneral.Controls.Add(this.label3);
            this.grpGeneral.Controls.Add(this.label2);
            this.grpGeneral.Controls.Add(this.Title);
            this.grpGeneral.Location = new System.Drawing.Point(21, 53);
            this.grpGeneral.Name = "grpGeneral";
            this.grpGeneral.Size = new System.Drawing.Size(380, 199);
            this.grpGeneral.TabIndex = 4;
            this.grpGeneral.TabStop = false;
            this.grpGeneral.Text = "General";
            // 
            // dropProductGroup
            // 
            this.dropProductGroup.FormattingEnabled = true;
            this.dropProductGroup.Location = new System.Drawing.Point(116, 164);
            this.dropProductGroup.Name = "dropProductGroup";
            this.dropProductGroup.Size = new System.Drawing.Size(245, 21);
            this.dropProductGroup.TabIndex = 11;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(24, 167);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(76, 13);
            this.label19.TabIndex = 10;
            this.label19.Text = "Product Group";
            // 
            // textProject
            // 
            this.textProject.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textProject.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.textProject.Location = new System.Drawing.Point(116, 138);
            this.textProject.Name = "textProject";
            this.textProject.Size = new System.Drawing.Size(245, 20);
            this.textProject.TabIndex = 9;
            // 
            // dropThickness
            // 
            this.dropThickness.FormattingEnabled = true;
            this.dropThickness.Location = new System.Drawing.Point(116, 109);
            this.dropThickness.Name = "dropThickness";
            this.dropThickness.Size = new System.Drawing.Size(245, 21);
            this.dropThickness.TabIndex = 8;
            // 
            // dropMaterial
            // 
            this.dropMaterial.FormattingEnabled = true;
            this.dropMaterial.Location = new System.Drawing.Point(116, 83);
            this.dropMaterial.Name = "dropMaterial";
            this.dropMaterial.Size = new System.Drawing.Size(245, 21);
            this.dropMaterial.TabIndex = 7;
            // 
            // dropFinish
            // 
            this.dropFinish.FormattingEnabled = true;
            this.dropFinish.Location = new System.Drawing.Point(116, 56);
            this.dropFinish.Name = "dropFinish";
            this.dropFinish.Size = new System.Drawing.Size(245, 21);
            this.dropFinish.TabIndex = 6;
            // 
            // textTitle
            // 
            this.textTitle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textTitle.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.textTitle.Location = new System.Drawing.Point(116, 30);
            this.textTitle.Name = "textTitle";
            this.textTitle.Size = new System.Drawing.Size(245, 20);
            this.textTitle.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(24, 138);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(40, 13);
            this.label5.TabIndex = 4;
            this.label5.Text = "Project";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(24, 112);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Thickness";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 86);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Material Info";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(24, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Finish Info";
            // 
            // Title
            // 
            this.Title.AutoSize = true;
            this.Title.Location = new System.Drawing.Point(24, 33);
            this.Title.Name = "Title";
            this.Title.Size = new System.Drawing.Size(27, 13);
            this.Title.TabIndex = 0;
            this.Title.Text = "Title";
            // 
            // groupUncontrolled
            // 
            this.groupUncontrolled.BackColor = System.Drawing.Color.White;
            this.groupUncontrolled.Controls.Add(this.btnUncontrolledCheckedDateRemove);
            this.groupUncontrolled.Controls.Add(this.dateUncontrolledDateCheckedThree);
            this.groupUncontrolled.Controls.Add(this.dateUncontrolledDateCheckedTwo);
            this.groupUncontrolled.Controls.Add(this.dateUncontolledCheckedOne);
            this.groupUncontrolled.Controls.Add(this.dateUncontrolledDrawn);
            this.groupUncontrolled.Controls.Add(this.dropStatusUncontrolled);
            this.groupUncontrolled.Controls.Add(this.dropCheckedThreeUncontrolled);
            this.groupUncontrolled.Controls.Add(this.dropCheckedTwoUncontrolled);
            this.groupUncontrolled.Controls.Add(this.dropUncontrolledDrawnBy);
            this.groupUncontrolled.Controls.Add(this.dropCheckedOneUncontrolled);
            this.groupUncontrolled.Controls.Add(this.dropUncontrolledRevision);
            this.groupUncontrolled.Controls.Add(this.label10);
            this.groupUncontrolled.Controls.Add(this.label9);
            this.groupUncontrolled.Controls.Add(this.label8);
            this.groupUncontrolled.Controls.Add(this.label7);
            this.groupUncontrolled.Controls.Add(this.label6);
            this.groupUncontrolled.Controls.Add(this.label1);
            this.groupUncontrolled.Location = new System.Drawing.Point(21, 274);
            this.groupUncontrolled.Name = "groupUncontrolled";
            this.groupUncontrolled.Size = new System.Drawing.Size(402, 191);
            this.groupUncontrolled.TabIndex = 5;
            this.groupUncontrolled.TabStop = false;
            this.groupUncontrolled.Text = "UnControlled";
            // 
            // btnUncontrolledCheckedDateRemove
            // 
            this.btnUncontrolledCheckedDateRemove.Location = new System.Drawing.Point(373, 56);
            this.btnUncontrolledCheckedDateRemove.Name = "btnUncontrolledCheckedDateRemove";
            this.btnUncontrolledCheckedDateRemove.Size = new System.Drawing.Size(18, 18);
            this.btnUncontrolledCheckedDateRemove.TabIndex = 17;
            this.btnUncontrolledCheckedDateRemove.Text = "x";
            this.btnUncontrolledCheckedDateRemove.UseVisualStyleBackColor = true;
            this.btnUncontrolledCheckedDateRemove.Click += new System.EventHandler(this.btnUncontrolledCheckedDateRemove_Click);
            // 
            // dateUncontrolledDateCheckedThree
            // 
            this.dateUncontrolledDateCheckedThree.Location = new System.Drawing.Point(241, 129);
            this.dateUncontrolledDateCheckedThree.Name = "dateUncontrolledDateCheckedThree";
            this.dateUncontrolledDateCheckedThree.Size = new System.Drawing.Size(126, 20);
            this.dateUncontrolledDateCheckedThree.TabIndex = 16;
            
            // 
            // dateUncontrolledDateCheckedTwo
            // 
            this.dateUncontrolledDateCheckedTwo.Location = new System.Drawing.Point(241, 104);
            this.dateUncontrolledDateCheckedTwo.Name = "dateUncontrolledDateCheckedTwo";
            this.dateUncontrolledDateCheckedTwo.Size = new System.Drawing.Size(126, 20);
            this.dateUncontrolledDateCheckedTwo.TabIndex = 15;
            //this.dateUncontrolledDrawn.ValueChanged += delegate (object sender, EventArgs e) { DateTimePicker_ValueChanged(this, e, ""); }; ;
            // 
            // dateUncontolledCheckedOne
            // 
            this.dateUncontolledCheckedOne.Location = new System.Drawing.Point(241, 77);
            this.dateUncontolledCheckedOne.Name = "dateUncontolledCheckedOne";
            this.dateUncontolledCheckedOne.Size = new System.Drawing.Size(126, 20);
            //this.dateUncontrolledDrawn.ValueChanged += delegate (object sender, EventArgs e) { DateTimePicker_ValueChanged(this, e, ""); }; ;
            this.dateUncontolledCheckedOne.TabIndex = 14;
            // 
            // dateUncontrolledDrawn
            // 
            this.dateUncontrolledDrawn.Location = new System.Drawing.Point(241, 53);
            this.dateUncontrolledDrawn.Name = "dateUncontrolledDrawn";
            this.dateUncontrolledDrawn.Size = new System.Drawing.Size(126, 20);
            this.dateUncontrolledDrawn.TabIndex = 13;
            //this.dateUncontrolledDrawn.ValueChanged += delegate (object sender, EventArgs e) { DateTimePicker_ValueChanged(sender, e, "Drawn_Date_Uncontrolled"); }; ;
            // 
            // dropStatusUncontrolled
            // 
            this.dropStatusUncontrolled.FormattingEnabled = true;
            this.dropStatusUncontrolled.Location = new System.Drawing.Point(127, 156);
            this.dropStatusUncontrolled.Name = "dropStatusUncontrolled";
            this.dropStatusUncontrolled.Size = new System.Drawing.Size(105, 21);
            this.dropStatusUncontrolled.TabIndex = 12;
            // 
            // dropCheckedThreeUncontrolled
            // 
            this.dropCheckedThreeUncontrolled.FormattingEnabled = true;
            this.dropCheckedThreeUncontrolled.Location = new System.Drawing.Point(127, 129);
            this.dropCheckedThreeUncontrolled.Name = "dropCheckedThreeUncontrolled";
            this.dropCheckedThreeUncontrolled.Size = new System.Drawing.Size(105, 21);
            this.dropCheckedThreeUncontrolled.TabIndex = 11;
            // 
            // dropCheckedTwoUncontrolled
            // 
            this.dropCheckedTwoUncontrolled.FormattingEnabled = true;
            this.dropCheckedTwoUncontrolled.Location = new System.Drawing.Point(127, 103);
            this.dropCheckedTwoUncontrolled.Name = "dropCheckedTwoUncontrolled";
            this.dropCheckedTwoUncontrolled.Size = new System.Drawing.Size(105, 21);
            this.dropCheckedTwoUncontrolled.TabIndex = 10;
            // 
            // dropUncontrolledDrawnBy
            // 
            this.dropUncontrolledDrawnBy.FormattingEnabled = true;
            this.dropUncontrolledDrawnBy.Location = new System.Drawing.Point(127, 53);
            this.dropUncontrolledDrawnBy.Name = "dropUncontrolledDrawnBy";
            this.dropUncontrolledDrawnBy.Size = new System.Drawing.Size(105, 21);
            this.dropUncontrolledDrawnBy.TabIndex = 9;
            // 
            // dropCheckedOneUncontrolled
            // 
            this.dropCheckedOneUncontrolled.FormattingEnabled = true;
            this.dropCheckedOneUncontrolled.Location = new System.Drawing.Point(127, 77);
            this.dropCheckedOneUncontrolled.Name = "dropCheckedOneUncontrolled";
            this.dropCheckedOneUncontrolled.Size = new System.Drawing.Size(105, 21);
            this.dropCheckedOneUncontrolled.TabIndex = 8;
            // 
            // dropUncontrolledRevision
            // 
            this.dropUncontrolledRevision.FormattingEnabled = true;
            this.dropUncontrolledRevision.Location = new System.Drawing.Point(127, 27);
            this.dropUncontrolledRevision.Name = "dropUncontrolledRevision";
            this.dropUncontrolledRevision.Size = new System.Drawing.Size(105, 21);
            this.dropUncontrolledRevision.TabIndex = 7;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(24, 159);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(37, 13);
            this.label10.TabIndex = 5;
            this.label10.Text = "Status";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(24, 132);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(96, 13);
            this.label9.TabIndex = 4;
            this.label9.Text = "Checked By Three";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(24, 106);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(89, 13);
            this.label8.TabIndex = 3;
            this.label8.Text = "Checked By Two";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(24, 80);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(88, 13);
            this.label7.TabIndex = 2;
            this.label7.Text = "Checked By One";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(24, 56);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 13);
            this.label6.TabIndex = 1;
            this.label6.Text = "Drawn By";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Revision";
            // 
            // groupControlled
            // 
            this.groupControlled.BackColor = System.Drawing.Color.White;
            this.groupControlled.Controls.Add(this.dateSignedBy);
            this.groupControlled.Controls.Add(this.dropSignedBy);
            this.groupControlled.Controls.Add(this.label17);
            this.groupControlled.Controls.Add(this.dateControlledDateCheckedThree);
            this.groupControlled.Controls.Add(this.dateControlledDateCheckedTwo);
            this.groupControlled.Controls.Add(this.dateContolledCheckedOne);
            this.groupControlled.Controls.Add(this.dateControlledDrawn);
            this.groupControlled.Controls.Add(this.dropStatusControlled);
            this.groupControlled.Controls.Add(this.dropCheckedThreeControlled);
            this.groupControlled.Controls.Add(this.dropCheckedTwoControlled);
            this.groupControlled.Controls.Add(this.dropDrawnByControlled);
            this.groupControlled.Controls.Add(this.dropCheckedOneControlled);
            this.groupControlled.Controls.Add(this.dropControlledRevision);
            this.groupControlled.Controls.Add(this.label11);
            this.groupControlled.Controls.Add(this.label12);
            this.groupControlled.Controls.Add(this.label13);
            this.groupControlled.Controls.Add(this.label14);
            this.groupControlled.Controls.Add(this.label15);
            this.groupControlled.Controls.Add(this.label16);
            this.groupControlled.Location = new System.Drawing.Point(21, 481);
            this.groupControlled.Name = "groupControlled";
            this.groupControlled.Size = new System.Drawing.Size(380, 221);
            this.groupControlled.TabIndex = 17;
            this.groupControlled.TabStop = false;
            this.groupControlled.Text = "Controlled";
            // 
            // dateSignedBy
            // 
            this.dateSignedBy.Location = new System.Drawing.Point(241, 155);
            this.dateSignedBy.Name = "dateSignedBy";
            this.dateSignedBy.Size = new System.Drawing.Size(126, 20);
            this.dateSignedBy.TabIndex = 19;
            // 
            // dropSignedBy
            // 
            this.dropSignedBy.FormattingEnabled = true;
            this.dropSignedBy.Location = new System.Drawing.Point(127, 155);
            this.dropSignedBy.Name = "dropSignedBy";
            this.dropSignedBy.Size = new System.Drawing.Size(105, 21);
            this.dropSignedBy.TabIndex = 18;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(24, 158);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(55, 13);
            this.label17.TabIndex = 17;
            this.label17.Text = "Signed By";
            // 
            // dateControlledDateCheckedThree
            // 
            this.dateControlledDateCheckedThree.Location = new System.Drawing.Point(241, 129);
            this.dateControlledDateCheckedThree.Name = "dateControlledDateCheckedThree";
            this.dateControlledDateCheckedThree.Size = new System.Drawing.Size(126, 20);
            this.dateControlledDateCheckedThree.TabIndex = 16;
            // 
            // dateControlledDateCheckedTwo
            // 
            this.dateControlledDateCheckedTwo.Location = new System.Drawing.Point(241, 104);
            this.dateControlledDateCheckedTwo.Name = "dateControlledDateCheckedTwo";
            this.dateControlledDateCheckedTwo.Size = new System.Drawing.Size(126, 20);
            this.dateControlledDateCheckedTwo.TabIndex = 15;
            // 
            // dateContolledCheckedOne
            // 
            this.dateContolledCheckedOne.Location = new System.Drawing.Point(241, 77);
            this.dateContolledCheckedOne.Name = "dateContolledCheckedOne";
            this.dateContolledCheckedOne.Size = new System.Drawing.Size(126, 20);
            this.dateContolledCheckedOne.TabIndex = 14;
            // 
            // dateControlledDrawn
            // 
            this.dateControlledDrawn.Location = new System.Drawing.Point(241, 53);
            this.dateControlledDrawn.Name = "dateControlledDrawn";
            this.dateControlledDrawn.Size = new System.Drawing.Size(126, 20);
            this.dateControlledDrawn.TabIndex = 13;
            //this.dateUncontrolledDrawn.ValueChanged += delegate(object sender, EventArgs e) { DateTimePicker_ValueChanged(this, e, ""); };
            // 
            // dropStatusControlled
            // 
            this.dropStatusControlled.FormattingEnabled = true;
            this.dropStatusControlled.Location = new System.Drawing.Point(127, 181);
            this.dropStatusControlled.Name = "dropStatusControlled";
            this.dropStatusControlled.Size = new System.Drawing.Size(105, 21);
            this.dropStatusControlled.TabIndex = 12;
            // 
            // dropCheckedThreeControlled
            // 
            this.dropCheckedThreeControlled.FormattingEnabled = true;
            this.dropCheckedThreeControlled.Location = new System.Drawing.Point(127, 129);
            this.dropCheckedThreeControlled.Name = "dropCheckedThreeControlled";
            this.dropCheckedThreeControlled.Size = new System.Drawing.Size(105, 21);
            this.dropCheckedThreeControlled.TabIndex = 11;
            // 
            // dropCheckedTwoControlled
            // 
            this.dropCheckedTwoControlled.FormattingEnabled = true;
            this.dropCheckedTwoControlled.Location = new System.Drawing.Point(127, 103);
            this.dropCheckedTwoControlled.Name = "dropCheckedTwoControlled";
            this.dropCheckedTwoControlled.Size = new System.Drawing.Size(105, 21);
            this.dropCheckedTwoControlled.TabIndex = 10;
            // 
            // dropDrawnByControlled
            // 
            this.dropDrawnByControlled.FormattingEnabled = true;
            this.dropDrawnByControlled.Location = new System.Drawing.Point(127, 53);
            this.dropDrawnByControlled.Name = "dropDrawnByControlled";
            this.dropDrawnByControlled.Size = new System.Drawing.Size(105, 21);
            this.dropDrawnByControlled.TabIndex = 9;
            // 
            // dropCheckedOneControlled
            // 
            this.dropCheckedOneControlled.FormattingEnabled = true;
            this.dropCheckedOneControlled.Location = new System.Drawing.Point(127, 77);
            this.dropCheckedOneControlled.Name = "dropCheckedOneControlled";
            this.dropCheckedOneControlled.Size = new System.Drawing.Size(105, 21);
            this.dropCheckedOneControlled.TabIndex = 8;
            // 
            // dropControlledRevision
            // 
            this.dropControlledRevision.FormattingEnabled = true;
            this.dropControlledRevision.Location = new System.Drawing.Point(127, 27);
            this.dropControlledRevision.Name = "dropControlledRevision";
            this.dropControlledRevision.Size = new System.Drawing.Size(105, 21);
            this.dropControlledRevision.TabIndex = 7;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(24, 184);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(37, 13);
            this.label11.TabIndex = 5;
            this.label11.Text = "Status";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(24, 132);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(96, 13);
            this.label12.TabIndex = 4;
            this.label12.Text = "Checked By Three";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(24, 106);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(89, 13);
            this.label13.TabIndex = 3;
            this.label13.Text = "Checked By Two";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(24, 80);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(88, 13);
            this.label14.TabIndex = 2;
            this.label14.Text = "Checked By One";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(24, 56);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(53, 13);
            this.label15.TabIndex = 1;
            this.label15.Text = "Drawn By";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(24, 30);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(48, 13);
            this.label16.TabIndex = 0;
            this.label16.Text = "Revision";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(21, 721);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(89, 13);
            this.label18.TabIndex = 18;
            this.label18.Text = "Comments/Notes";
            // 
            // textNotesComments
            // 
            this.textNotesComments.AcceptsReturn = true;
            this.textNotesComments.AcceptsTab = true;
            this.textNotesComments.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textNotesComments.Location = new System.Drawing.Point(24, 737);
            this.textNotesComments.Multiline = true;
            this.textNotesComments.Name = "textNotesComments";
            this.textNotesComments.Size = new System.Drawing.Size(377, 89);
            this.textNotesComments.TabIndex = 19;
            // 
            // lblSacoComptHeader
            // 
            this.lblSacoComptHeader.AutoSize = true;
            this.lblSacoComptHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSacoComptHeader.Location = new System.Drawing.Point(18, 17);
            this.lblSacoComptHeader.Name = "lblSacoComptHeader";
            this.lblSacoComptHeader.Size = new System.Drawing.Size(206, 16);
            this.lblSacoComptHeader.TabIndex = 20;
            this.lblSacoComptHeader.Text = "SACO Component Properties";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(326, 3);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 44);
            this.btnSave.TabIndex = 21;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // TaskpaneHost
            // 
            this.BackColor = System.Drawing.Color.Silver;
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.lblSacoComptHeader);
            this.Controls.Add(this.textNotesComments);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.groupControlled);
            this.Controls.Add(this.groupUncontrolled);
            this.Controls.Add(this.grpGeneral);
            this.Controls.Add(this.btnRefreshProps);
            this.Name = "TaskpaneHost";
            this.Size = new System.Drawing.Size(1051, 831);
            this.grpGeneral.ResumeLayout(false);
            this.grpGeneral.PerformLayout();
            this.groupUncontrolled.ResumeLayout(false);
            this.groupUncontrolled.PerformLayout();
            this.groupControlled.ResumeLayout(false);
            this.groupControlled.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }


        #region Control Actions

        private void button2_Click(object sender, EventArgs e)
        {

        }


        private void DateTimePicker_ValueChanged(object sender, EventArgs e, string propertyName)
        {
            DateTimePicker dateTimePicker = sender as DateTimePicker;
            dateTimePicker.Format = DateTimePickerFormat.Custom;
            dateTimePicker.CustomFormat = "dd MMMM yyyy";

            //Save Property
            customProps.Set2(propertyName, dateTimePicker.Value.ToString("dd/MM/yyyy"));

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //LoadTaskPaneData();
            //PopulateModel();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Get the lists for the drop down boxes such as the thickness etc.
        /// </summary>
        private void GetLists()
        {
            DesignersInitals = File.ReadLines(@"C:\Users\cabuk\Documents\Development\SACO Sw Properties\Drawn By.txt").ToList();
            CheckedByInitials = File.ReadLines(@"C:\Users\cabuk\Documents\Development\SACO Sw Properties\Checked By.txt").ToList();
            SignedBy = File.ReadLines(@"C:\Users\cabuk\Documents\Development\SACO Sw Properties\Signed By.txt").ToList();
            MaterialThickness = File.ReadLines(@"C:\Users\cabuk\Documents\Development\SACO Sw Properties\Material Thickness.txt").ToList();
            ProductGroup = File.ReadLines(@"C:\Users\cabuk\Documents\Development\SACO Sw Properties\Product Group.txt").ToList();
            MaterialInfo = File.ReadLines(@"C:\Users\cabuk\Documents\Development\SACO Sw Properties\Material Info.txt").ToList();
            FinishIngo = File.ReadLines(@"C:\Users\cabuk\Documents\Development\SACO Sw Properties\Finish Info.txt").ToList();

            AssignListsToDropDowns();

            //check if open file is a part file and that is is the correct template
            //if(swApp != null) 
            //{
            //    swModel = swApp.ActiveDoc;

            //    modelExt = swModel.Extension;
            //    customProps = modelExt.CustomPropertyManager["Default"];

            //    var val = "";
            //    var valOut = "";

            //    customProps.Get4("Template_Type", false, out val, out valOut);

            //    if (val == "SACO Component V1")
            //    {
            //        LoadTaskPaneData();
            //    }
            //}
        }


        /// <summary>
        /// Assign the items from the lists into the relevant drop downs
        /// </summary>
        private void AssignListsToDropDowns()
        {
            foreach (var item in DesignersInitals)
            {
                dropDrawnByControlled.Items.Add(item);
            }

            foreach (var item in CheckedByInitials)
            {
                dropCheckedOneUncontrolled.Items.Add(item);
                dropCheckedTwoUncontrolled.Items.Add(item);
                dropCheckedThreeUncontrolled.Items.Add(item);

                dropCheckedOneControlled.Items.Add(item);
                dropCheckedTwoControlled.Items.Add(item);
                dropCheckedThreeControlled.Items.Add(item);
            }

            foreach (var item in SignedBy)
            {
                dropSignedBy.Items.Add(item);
            }

            foreach (var item in MaterialThickness)
            {
                dropThickness.Items.Add(item);
            }

            foreach (var item in ProductGroup)
            {
                dropProductGroup.Items.Add(item);
            }

            foreach (var item in MaterialThickness)
            {
                dropThickness.Items.Add(item);
            }

            foreach (var item in Revision)
            {
                dropControlledRevision.Items.Add(item);
                dropUncontrolledRevision.Items.Add(item);
            }

            foreach (var item in MaterialInfo)
            {
                dropMaterial.Items.Add(item);
            }

            foreach (var item in FinishIngo)
            {
                dropFinish.Items.Add(item);
            }
        }

        private void SaveProps()
        {
            swModel = swApp.ActiveDoc;

            modelExt = swModel.Extension;
            customProps = modelExt.CustomPropertyManager[""];

            #region General

            customProps.Set2("Title_Info", textTitle.Text);
            customProps.Set2("Notes_Comments", textNotesComments.Text);
            customProps.Set2("Project", textProject.Text);
            customProps.Set2("Material_Info", dropMaterial.SelectedItem.ToString());
            customProps.Set2("Product_Group", dropProductGroup.SelectedItem.ToString());
            customProps.Set2("Thickness", dropThickness.SelectedItem.ToString());
            //customProps.Set2("Finish_Info", dropFinish.SelectedItem.ToString());       

            #endregion

            //#region Uncontrolled

            //customProps.Get4("Drawn_Controlled", false, out val, out valOut);
            //dropDrawnByControlled.SelectedItem = val;

            //customProps.Get4("Drawn_UnControlled", false, out val, out valOut);
            //dropUncontrolledDrawnBy.SelectedItem = val;

            //customProps.Get4("UnControlled_Checked_By_One", false, out val, out valOut);
            //dropCheckedOneUncontrolled.SelectedItem = val;

            //customProps.Get4("UnControlled_Checked_By_Two", false, out val, out valOut);
            //dropCheckedTwoUncontrolled.SelectedItem = val;

            //customProps.Get4("UnControlled_Checked_By_Three", false, out val, out valOut);
            //dropCheckedThreeUncontrolled.SelectedItem = val;

            customProps.Set2("Drawn_Date_Uncontrolled", dateUncontrolledDrawn.Value.ToString("dd/MM/yyyy"));
            customProps.Set2("Uncontrolled_Date_Checked_One", dateUncontolledCheckedOne.Value.ToString("dd/MM/yyyy"));
            customProps.Set2("Uncontrolled_Date_Checked_Two", dateUncontrolledDateCheckedTwo.Value.ToString("dd/MM/yyyy"));

            //customProps.Get4("Uncontrolled_Date_Checked_Two", false, out val, out valOut);
            //dateUncontrolledDateCheckedTwo.Value = ParseDate(val);

            //customProps.Get4("Uncontrolled_Date_Checked_Three", false, out val, out valOut);
            //dateUncontrolledDateCheckedThree.Value = ParseDate(val);

            //#endregion

            //#region Controlled

            //customProps.Get4("Controlled_Checked_By_One", false, out val, out valOut);
            //dropCheckedOneControlled.SelectedItem = val;

            //customProps.Get4("Controlled_Checked_By_Two", false, out val, out valOut);
            //dropCheckedTwoControlled.SelectedItem = val;

            //customProps.Get4("Controlled_Checked_By_Three", false, out val, out valOut);
            //dropCheckedThreeControlled.SelectedItem = val;

            //customProps.Get4("Controlled_Date_Checked_One", false, out val, out valOut);
            //dateContolledCheckedOne.Value = ParseDate(val);

            //customProps.Get4("Controlled_Date_Checked_Two", false, out val, out valOut);
            //dateControlledDateCheckedTwo.Value = ParseDate(val);

            //customProps.Get4("Controlled_Date_Checked_Three", false, out val, out valOut);
            //dateControlledDateCheckedThree.Value = ParseDate(val);

            //customProps.Get4("Signed_By", false, out val, out valOut);
            //dropSignedBy.SelectedItem = val;

            //customProps.Get4("Date_Released", false, out val, out valOut);
            //dateSignedBy.Value = ParseDate(val);

            //#endregion   

        }

        /// <summary>
        /// This method is called each time a file is opened or swapped in the window to a different file
        /// </summary>
        private void LoadTaskPaneData()
        {
            var val = "";
            var valOut = "";
            DateTime dateVal;

            swModel = swApp.ActiveDoc;

            modelExt = swModel.Extension;


            //Ge the extension and make sure it is a part file otherwise do not load

            customProps = modelExt.CustomPropertyManager[""]; //modelExt.CustomPropertyManager["Default"];

            #region General

            customProps.Get4("Title_Info", false, out val, out valOut);
            textTitle.Text = val;

            customProps.Get4("Project", false, out val, out valOut);
            textProject.Text = val;

            customProps.Get4("Material_Info", false, out val, out valOut);
            dropMaterial.SelectedItem = val;

            customProps.Get4("Product_Group", false, out val, out valOut);
            dropProductGroup.SelectedItem = val;

            customProps.Get4("Notes_Comments", false, out val, out valOut);
            textNotesComments.Text = val;

            customProps.Get4("Thickness", false, out val, out valOut);
            dropThickness.SelectedItem = val;

            customProps.Get4("Finish_Info", false, out val, out valOut);
            dropFinish.SelectedItem = val;

            #endregion

            #region Uncontrolled

            

            customProps.Get4("Drawn_Controlled", false, out val, out valOut);
            dropDrawnByControlled.SelectedItem = val;

            customProps.Get4("Drawn_UnControlled", false, out val, out valOut);
            dropUncontrolledDrawnBy.SelectedItem = val;

            customProps.Get4("UnControlled_Checked_By_One", false, out val, out valOut);
            dropCheckedOneUncontrolled.SelectedItem = val;

            customProps.Get4("UnControlled_Checked_By_Two", false, out val, out valOut);
            dropCheckedTwoUncontrolled.SelectedItem = val;

            customProps.Get4("UnControlled_Checked_By_Three", false, out val, out valOut);
            dropCheckedThreeUncontrolled.SelectedItem = val;

            customProps.Get4("Uncontrolled_Date_Checked_One", false, out val, out valOut);
            SetDateFormats(dateUncontolledCheckedOne, val);

            customProps.Get4("Uncontrolled_Date_Checked_Two", false, out val, out valOut);
            SetDateFormats(dateUncontrolledDateCheckedTwo, val);

            customProps.Get4("Uncontrolled_Date_Checked_Three", false, out val, out valOut);
            SetDateFormats(dateUncontrolledDateCheckedThree, val);

            #endregion

            #region Controlled

            customProps.Get4("Controlled_Checked_By_One", false, out val, out valOut);
            dropCheckedOneControlled.SelectedItem = val;

            customProps.Get4("Controlled_Checked_By_Two", false, out val, out valOut);
            dropCheckedTwoControlled.SelectedItem = val;

            customProps.Get4("Controlled_Checked_By_Three", false, out val, out valOut);
            dropCheckedThreeControlled.SelectedItem = val;

            customProps.Get4("Controlled_Date_Checked_One", false, out val, out valOut);
            SetDateFormats(dateContolledCheckedOne, val);

            customProps.Get4("Controlled_Date_Checked_Two", false, out val, out valOut);
            SetDateFormats(dateControlledDateCheckedTwo, val);

            customProps.Get4("Controlled_Date_Checked_Three", false, out val, out valOut);
            SetDateFormats(dateControlledDateCheckedThree, val);  

            customProps.Get4("Signed_By", false, out val, out valOut);
            dropSignedBy.SelectedItem = val;

            customProps.Get4("Date_Released", false, out val, out valOut);
            SetDateFormats(dateSignedBy, val);

            #endregion   
        }

        /// <summary>
        /// Checks if there is a value in the properties for date and if so it parses this otherwise it sets the date time picker to be blank
        /// </summary>
        /// <param name="dtpControl">Name of the DateTimePicker control</param>
        /// <param name="val">The value from the files summary proprties</param>
        private void SetDateFormats(DateTimePicker dtpControl, string val)
        {
            if(val == "")
            {
                dtpControl.Format = DateTimePickerFormat.Custom;
                dtpControl.CustomFormat = " ";
            }
            else
            {
                dtpControl.Value = ParseDate(val);
            }
        }

        /// <summary>
        /// Clear the date from the date control picker
        /// </summary>
        /// <param name="dtpControl">The DateTimePicker control to clear the date</param>
        private void ClearDate(DateTimePicker dtpControl, string propertyName)
        {

            //Clear the control and then clear the property in the file properties
            dtpControl.Format = DateTimePickerFormat.Custom;
            dtpControl.CustomFormat = " ";

            //Clear the property in the file properties
            customProps.Set2(propertyName, "");
        }

        private void PopulateModel()
        {
            var val = "";
            var valOut = "";
            DateTime dateVal;

            swModel = swApp.ActiveDoc;

            modelExt = swModel.Extension;
            customProps = modelExt.CustomPropertyManager["Default"];

            customProps.Get4("Title_Info", false, out val, out valOut);
            PropertiesModel.TitleInfo = val;

            customProps.Get4("Project", false, out val, out valOut);
            PropertiesModel.Project = val;

            customProps.Get4("Material_Info", false, out val, out valOut);
            PropertiesModel.MaterialInfo = val;

            customProps.Get4("Product_Group", false, out val, out valOut);
            PropertiesModel.ProductGroup = val;

        }

        /// <summary>
        /// Parse any string value into a date format for the date time picker
        /// </summary>
        /// <param name="value">Pass the date as a string</param>
        /// <returns>Returns the string entered into a valid DateTime formatted as dd/MM/yyyy</returns>
        private DateTime ParseDate(string value)
        {
            if (value != "")
            {
                return DateTime.ParseExact(value, "dd/MM/yyyy",
                                System.Globalization.CultureInfo.InvariantCulture);
            }

            return DateTime.Now;
        }

        #endregion


        public void getSwApp(SldWorks SwApp)
        {
            swApp = SwApp;

            AttachEvenHandlers();
            //swApp.ActiveDocChangeNotify += SwApp_ActiveDocChangeNotify;

            //swApp.FileCloseNotify += new DSldWorksEvents_FileCloseNotifyEventHandler(SwApp_DocumentClosed);
            
            
        }

        public void AttachEvenHandlers()
        {
            swApp.ActiveDocChangeNotify += this.SwApp_ActiveDocChangeNotify;
            //swApp.CommandCloseNotify += this.SwApp_DocumentClosed; // new DSldWorksEvents_FileCloseNotifyEventHandler(SwApp_DocumentClosed);
            //swApp.ActiveModelDocChangeNotify += this.SwApp_DocumentClosed;
            //swApp.DestroyNotify += this.SwApp_DocumentClosed;
            //swApp.FileCloseNotify += new DSldWorksEvents_FileCloseNotifyEventHandler(SwApp_DocumentClosed);
            // return true;
            //swApp.
            swApp.FileCloseNotify += this.SwApp_DocumentClosed;
        }

        //public delegate void DSldWorksEvents_FileCloseNotifyEventHandler(string FileName, int reason);

       

        public int SwApp_DocumentClosed(string fielname, int number)
        {
            //Console.WriteLine(comm);
            //check if there is an active doucment open and if not hide the task pane
            if(swApp.ActiveDoc == null) 
            {
                SetPropComponentsToHide();
            }

            return 0;
        }   


        private int SwApp_ActiveDocChangeNotify()
        {
            OnActiveDocChangeNotify();
            return 1;
        }

        /// <summary>
        /// Runs when a file is opened or when switching to another
        /// </summary>
        private void OnActiveDocChangeNotify()
        {
            ModelDoc2 activeDoc = swApp.ActiveDoc as ModelDoc2;
            if (activeDoc != null)
            {

               
                // Check if the active document is a part, assembly, or drawing document
                switch (activeDoc.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        // Perform actions for part document
                        //check it is the correct part
                        var val = "";
                        var valOut = "";
                        modelExt = activeDoc.Extension;
                        customProps = modelExt.CustomPropertyManager[""];

                        customProps.Get4("Template_Type", false, out val, out valOut);
                        
                        if(val == "SACO Component V1")
                        {
                            SetPropComponentsToShow();
                            LoadTaskPaneData();
                        }
                        else
                        {
                            SetPropComponentsToHide();                           
                        }
                        break;
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        // Perform actions for assembly document
                        SetPropComponentsToHide();
                        break;
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        // Perform actions for drawing document            
                        SetPropComponentsToHide();
                        break;
                    default:
                        // Do nothing for other document types
                        break;
                }
            }
            else
            {
                SetPropComponentsToHide();
            }
        }

        private void SetPropComponentsToShow()
        {
            grpGeneral.Visible = true;
            groupUncontrolled.Visible = true;
            groupControlled.Visible = true;
            textNotesComments.Visible = true;
            btnSave.Visible = true;
            btnRefreshProps.Visible = true;
        }

        private void SetPropComponentsToHide()
        {
            grpGeneral.Visible = false;
            groupUncontrolled.Visible = false;
            groupControlled.Visible = false;
            textNotesComments.Visible = false;
            btnSave.Visible = false;
            btnRefreshProps.Visible = false;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveProps();
        }

        private void btnUncontrolledCheckedDateRemove_Click(object sender, EventArgs e)
        {
            ClearDate(dateUncontrolledDrawn, "Drawn_Date_Uncontrolled");
        }

 







        //private void dateUncontolledCheckedOne_ValueChanged(object sender, EventArgs e)
        //{
        //    dateUncontolledCheckedOne.Format = DateTimePickerFormat.Custom;
        //    dateUncontolledCheckedOne.CustomFormat = "dd MMMM yyyy";
        //}
    }







    public class FileOpenNotifyHandler : IDisposable
    {
        SldWorks swApp;
        int fileOpenNotifyCookie;

        public void Init(SldWorks swApp)
        {
            this.swApp = swApp;
            fileOpenNotifyCookie = 1; //swApp.SetAddinCallbackInfo((int)swAddinCallback_e.swAddinCallback_FileOpenPostNotify, 
                                      //      null, 
                                      //      new FileOpenNotifyDelegate(FileOpenNotify));

            //fileOpenNotifyCookie = swApp.SetAddinCallbackInfo2((int)swAddinCallback_e.swAddinCallback_FileOpenPostNotify, null, new FileOpenNotifyDelegate(FileOpenNotify));

         
        }

        

        public int FileOpenNotify(object FileName)
        {
            string fileNameString = (string)FileName;

            if (System.IO.Path.GetExtension(fileNameString) == ".sldprt")
            {
                System.Diagnostics.Debug.Print("Part file opened: " + fileNameString);
            }
            else
            {
                System.Diagnostics.Debug.Print("Non-part file opened: " + fileNameString);
            }

            // Return 0 to indicate success
            return 0;
        }

        

        public void Dispose()
        {
            if (swApp != null && fileOpenNotifyCookie != 0)
            {
                swApp.RemoveCallback(fileOpenNotifyCookie);
                //swApp.RemoveAddinCallbackInfo(fileOpenNotifyCookie);
                fileOpenNotifyCookie = 0;
            }
        }


        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        public delegate int FileOpenNotifyDelegate(string fileName);

        //[ComVisible(false)]
        //public delegate int FileOpenNotifyDelegate(string FileName);

    }
}
