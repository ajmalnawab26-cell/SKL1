using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Printing;
using Microsoft.Data.Sqlite;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

// Ambiguity fixes
using Color = System.Drawing.Color;
using SDImage = System.Drawing.Image;

namespace RishtaManagerPro
{
    public class MainForm : Form
    {
        // ---------- DB Paths ----------
        string DbFile  => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "records.db");
        string ConnStr => $"Data Source={DbFile}";
        string imagePath = "";

        // ---------- Top UI ----------
        TextBox? txtSearch;
        DataGridView? grid;
        PictureBox? pic;

        // ---------- Shared helpers ----------
        Color? userAccent;

        // ---------- Personal tab controls ----------
        TextBox? txtName, txtFatherName, txtPhone, txtEducation, txtEducationExtra,
                 txtAge, txtHeight, txtWeight, txtAddress, txtBodyType, txtComplexion,
                 txtWorkDetails, txtHouseType, txtHouseSize, txtOtherProperty,
                 txtParentsAlive, txtFatherJob, txtMotherJob, txtMonthlyIncome,
                 txtFirstWife, txtChildren, txtRelationWithSeeker, txtExtraInfo, txtSmokingDrugs;
        ComboBox? cmbGender, cmbMarital, cmbReligion, cmbMaslak, cmbCaste, cmbCity;

        // ---------- Family tab (counters) ----------
        TextBox? txtSisters, txtMarriedSisters, txtBrothers, txtMarriedBrothers, txtSiblingNumber, txtDisability;

        // ---------- Desired tab ----------
        TextBox? txtD_Age, txtD_Education, txtD_Height, txtD_House, txtD_CityProvince,
                 txtD_Mobile, txtD_WhatsApp, txtD_Reference, txtD_Religion;
        TextBox? txtD_Caste, txtD_Maslak, txtD_MaritalStatus;

        // ---------- Footer tab ----------
        TextBox? txtReceivedBy, txtCheckedBy, txtCommissionNote, txtMainOffice;

        // ---------- Buttons ----------
        Button? btnAdd, btnUpdate, btnDelete, btnNew, btnPrint, btnPdf, btnCsv, btnBackup, btnRestore, btnTheme, btnOptions, btnBrowse;

        public MainForm()
        {
            Text = "Rishta Manager Pro (اردو + English)";
            Width = 1280; Height = 820;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 10);

            CreateUi();
            EnsureDb();
            LoadData();

            using var con = new SqliteConnection(ConnStr);
            con.Open();
            txtMainOffice!.Text = GetSetting(con, "MAIN_OFFICE");
        }

        // =========================================================
        //                          UI
        // =========================================================
        void CreateUi()
        {
            // Search
            var lblSearch = L("Search / تلاش:");
            lblSearch.Left = 20; lblSearch.Top = 15;
            Controls.Add(lblSearch);
            txtSearch = new TextBox { Left = 120, Top = 12, Width = 620 };
            txtSearch.TextChanged += (s, e) => LoadData(txtSearch!.Text);
            Controls.Add(txtSearch);

            // Grid
            grid = new DataGridView
            {
                Left = 20, Top = 45, Width = 1220, Height = 260,
                ReadOnly = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            grid.SelectionChanged += Grid_SelectionChanged;
            Controls.Add(grid);

            // Photo
            pic = new PictureBox { Left = 1060, Top = 320, Width = 180, Height = 220, BorderStyle = BorderStyle.FixedSingle, SizeMode = PictureBoxSizeMode.Zoom };
            btnBrowse = B("تصویر منتخب کریں / Browse", 1060, 545, 180, 34, (s,e)=>BrowseImage());
            Controls.Add(pic); Controls.Add(btnBrowse);

            // Tabs
            var tabs = new TabControl { Left = 20, Top = 315, Width = 1020, Height = 320 };
            var tpPersonal = new TabPage("Personal / ذاتی");
            var tpFamily   = new TabPage("Family & Home / گھرانہ");
            var tpDesired  = new TabPage("Desired / رشتہ مطلوب");
            var tpFooter   = new TabPage("Footer / دستخط");
            tabs.TabPages.AddRange(new[] { tpPersonal, tpFamily, tpDesired, tpFooter });
            Controls.Add(tabs);

            // Personal Tab Layout
            int x = 10, y = 15, w = 320, gap = 34;
            AddText("Name / نام:",     tpPersonal, ref txtName,          x,  ref y, w, gap);
            AddText("Father / والدیت:",tpPersonal, ref txtFatherName,    x,  ref y, w, gap);

            AddCombo("Gender / جنس:",   tpPersonal, ref cmbGender,   new[] { "Male", "Female" }, x, ref y, w, gap);
            AddCombo("Marital / شادی شدہ:", tpPersonal, ref cmbMarital, new[] { "Single", "Married", "Divorced", "Widowed" }, x, ref y, w, gap);

            AddText("Age / عمر:",       tpPersonal, ref txtAge,          x,  ref y, w, gap);
            AddText("Height(cm) / قد:", tpPersonal, ref txtHeight,       x,  ref y, w, gap);
            AddText("Weight(kg) / وزن:",tpPersonal, ref txtWeight,       x,  ref y, w, gap);

            AddText("Phone / فون:",     tpPersonal, ref txtPhone,        x,  ref y, w, gap);

            // from Settings
            using var conOpt = new SqliteConnection(ConnStr); conOpt.Open();
            var casteList    = GetSetting(conOpt,"CASTE_OPTIONS").Split(';', StringSplitOptions.RemoveEmptyEntries);
            var cityList     = GetSetting(conOpt,"CITY_OPTIONS").Split(';', StringSplitOptions.RemoveEmptyEntries);
            var religionList = GetSetting(conOpt,"RELIGION_OPTIONS").Split(';', StringSplitOptions.RemoveEmptyEntries);
            var maslakList   = GetSetting(conOpt,"MASLAK_OPTIONS").Split(';', StringSplitOptions.RemoveEmptyEntries);

            AddCombo("City / شہر:",     tpPersonal, ref cmbCity,     cityList,    x, ref y, w, gap);
            AddCombo("Religion / مذہب:",tpPersonal, ref cmbReligion, religionList,x, ref y, w, gap);
            AddCombo("Maslak / مسلک:",  tpPersonal, ref cmbMaslak,   maslakList,  x, ref y, w, gap);
            AddCombo("Caste / ذات:",    tpPersonal, ref cmbCaste,    casteList,   x, ref y, w, gap);

            int x2 = 360; y = 15;
            AddText("Education / تعلیم:",      tpPersonal, ref txtEducation,       x2, ref y, w, gap);
            AddText("Extra Education / اضافی:",tpPersonal, ref txtEducationExtra,  x2, ref y, w, gap);
            AddText("Body Type / جسامت:",      tpPersonal, ref txtBodyType,        x2, ref y, w, gap);
            AddText("Complexion / رنگت:",      tpPersonal, ref txtComplexion,      x2, ref y, w, gap);
            AddMulti("Work Details / کاروبار:",tpPersonal, ref txtWorkDetails,     x2, ref y, w, 60, gap);
            AddText("Father Job / والد کا پیشہ:",tpPersonal, ref txtFatherJob,     x2, ref y, w, gap);
            AddText("Mother Job / والدہ کا پیشہ:",tpPersonal, ref txtMotherJob,    x2, ref y, w, gap);
            AddText("Monthly Income / آمدنی:", tpPersonal, ref txtMonthlyIncome,   x2, ref y, w, gap);
            AddMulti("Address / پتہ:",         tpPersonal, ref txtAddress,         x2, ref y, w, 60, gap);

            // Family Tab
            x = 10; y = 15;
            AddText("House Type / گھر(ذاتی/کرایہ):", tpFamily, ref txtHouseType,  x, ref y, w, gap);
            AddText("House Size / گھر کا سائز:",      tpFamily, ref txtHouseSize,  x, ref y, w, gap);
            AddText("Other Property / دیگر جائیداد:", tpFamily, ref txtOtherProperty, x, ref y, w, gap);
            AddText("Parents Alive / والدین حیات:",   tpFamily, ref txtParentsAlive,  x, ref y, w, gap);

            AddText("Sisters / بہنیں:",               tpFamily, ref txtSisters,        x, ref y, w, gap);
            AddText("Married Sisters / شادی شدہ بہنیں:",tpFamily, ref txtMarriedSisters, x, ref y, w, gap);
            AddText("Brothers / بھائی:",              tpFamily, ref txtBrothers,       x, ref y, w, gap);
            AddText("Married Brothers / شادی شدہ بھائی:",tpFamily, ref txtMarriedBrothers, x, ref y, w, gap);
            AddText("Sibling Number / نمبر:",         tpFamily, ref txtSiblingNumber,  x, ref y, w, gap);
            AddText("Disability / بیماری/معذوری:",    tpFamily, ref txtDisability,     x, ref y, w, gap);
            AddText("First Wife / پہلی بیوی:",         tpFamily, ref txtFirstWife,      x, ref y, w, gap);
            AddText("Children / بچے:",                tpFamily, ref txtChildren,        x, ref y, w, gap);
            AddText("RelationWithSeeker / تعلق:",     tpFamily, ref txtRelationWithSeeker, x, ref y, w, gap);
            AddMulti("Extra Info / اضافی:",           tpFamily, ref txtExtraInfo,      x, ref y, w, 60, gap);
            AddText("Smoking/Drugs / نشہ:",           tpFamily, ref txtSmokingDrugs,   x, ref y, w, gap);

            // Desired Tab
            x = 10; y = 15;
            AddText("Wanted Marital / شادی شدہ حیثیت:", tpDesired, ref txtD_MaritalStatus, x, ref y, w, gap);
            AddText("Wanted Age / عمر:",                tpDesired, ref txtD_Age,           x, ref y, w, gap);
            AddText("Wanted Education / تعلیم:",        tpDesired, ref txtD_Education,     x, ref y, w, gap);
            AddText("Wanted Religion / مذہب:",          tpDesired, ref txtD_Religion,      x, ref y, w, gap);
            AddText("Wanted Maslak / مسلک:",            tpDesired, ref txtD_Maslak,        x, ref y, w, gap);
            AddText("Wanted Caste / ذات:",              tpDesired, ref txtD_Caste,         x, ref y, w, gap);
            AddText("Wanted Height / قد:",              tpDesired, ref txtD_Height,        x, ref y, w, gap);
            AddText("Wanted House / گھر:",              tpDesired, ref txtD_House,         x, ref y, w, gap);
            AddText("City/Province قید:",               tpDesired, ref txtD_CityProvince,  x, ref y, w, gap);
            AddText("Mobile / موبائل:",                 tpDesired, ref txtD_Mobile,        x, ref y, w, gap);
            AddText("WhatsApp:",                        tpDesired, ref txtD_WhatsApp,      x, ref y, w, gap);
            AddMulti("Reference / ریفرینس:",            tpDesired, ref txtD_Reference,     x, ref y, w, 60, gap);

            // Footer Tab
            x = 10; y = 15;
            AddText("Received By / وصول کنندہ:", tpFooter, ref txtReceivedBy,  x, ref y, 380, gap);
            AddText("Checked By / چیک:",         tpFooter, ref txtCheckedBy,   x, ref y, 380, gap);
            AddMulti("Commission/Note / کمیشن:", tpFooter, ref txtCommissionNote, x, ref y, 380, 60, gap);
            AddMulti("Main Office Line:",        tpFooter, ref txtMainOffice,  x, ref y, 950, 60, gap);

            // Buttons row
            int bx = 20, by = 650, bw = 120, bh = 36, s = 10;
            btnAdd     = B("Add / شامل کریں",  bx,               by, bw, bh, (s1,e)=>AddRecord());
            btnUpdate  = B("Edit / تبدیلی",    bx+(bw+s),        by, bw, bh, (s1,e)=>UpdateRecord());
            btnDelete  = B("Delete / حذف",     bx+2*(bw+s),      by, bw, bh, (s1,e)=>DeleteRecord());
            btnNew     = B("New / نیا",        bx+3*(bw+s),      by, bw, bh, (s1,e)=>ClearForm());
            btnPrint   = B("Print",            bx+4*(bw+s),      by, bw, bh, (s1,e)=>PrintRecord());
            btnPdf     = B("Export PDF",       bx+5*(bw+s),      by, bw, bh, (s1,e)=>ExportPdf());
            btnCsv     = B("Export CSV",       bx+6*(bw+s),      by, bw, bh, (s1,e)=>ExportCsv());
            btnBackup  = B("Backup DB",        bx+7*(bw+s),      by, bw, bh, (s1,e)=>BackupDb());
            btnRestore = B("Restore DB",       bx+8*(bw+s),      by, bw, bh, (s1,e)=>RestoreDb());
            btnTheme   = B("Theme / رنگ",      bx+9*(bw+s),      by, bw, bh, (s1,e)=>PickTheme());
            btnOptions = B("⚙ Options",        bx+10*(bw+s),     by, 110, bh, (s1,e)=>OpenOptionsDialog());
            Controls.AddRange(new Control[]{ btnAdd,btnUpdate,btnDelete,btnNew,btnPrint,btnPdf,btnCsv,btnBackup,btnRestore,btnTheme,btnOptions });
        }

        // ---------- UI helpers ----------
        Label L(string t) => new Label { Text = t, AutoSize = true };
        Button B(string t, int l, int tp, int w, int h, EventHandler onClick)
        { var b=new Button{Text=t,Left=l,Top=tp,Width=w,Height=h}; b.Click+=onClick; return b; }

        void AddText(string caption, Control parent, ref TextBox? box, int x, ref int y, int w, int gap)
        {
            var lbl = new Label{ Text=caption, Left=x, Top=y, AutoSize=true }; parent.Controls.Add(lbl);
            box = new TextBox { Left = x+220, Top = y-4, Width = w }; parent.Controls.Add(box); y+=gap;
        }
        void AddMulti(string caption, Control parent, ref TextBox? box, int x, ref int y, int w, int h, int gap)
        {
            var lbl = new Label{ Text=caption, Left=x, Top=y, AutoSize=true }; parent.Controls.Add(lbl);
            box = new TextBox { Left = x+220, Top = y-4, Width = w, Height = h, Multiline = true, ScrollBars = ScrollBars.Vertical };
            parent.Controls.Add(box); y += (h + (gap-10));
        }
        void AddCombo(string caption, Control parent, ref ComboBox? cmb, string[] items, int x, ref int y, int w, int gap)
        {
            var lbl = new Label{ Text=caption, Left=x, Top=y, AutoSize=true }; parent.Controls.Add(lbl);
            cmb = new ComboBox{ Left = x+220, Top=y-6, Width=w, DropDownStyle=ComboBoxStyle.DropDown };
            cmb.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmb.AutoCompleteSource = AutoCompleteSource.ListItems;
            cmb.Items.AddRange(items);
            parent.Controls.Add(cmb); y+=gap;
        }

        // =========================================================
        //                        DATABASE
        // =========================================================
        void EnsureDb()
        {
            if (!File.Exists(DbFile)) File.Create(DbFile).Dispose();

            using var con = new SqliteConnection(ConnStr);
            con.Open();

            // Users table (superset of all fields)
            new SqliteCommand(@"
CREATE TABLE IF NOT EXISTS Users(
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Name TEXT, FatherName TEXT, City TEXT, Phone TEXT,
    Education TEXT, Religion TEXT, Caste TEXT, Occupation TEXT,
    MaritalStatus TEXT, Address TEXT, ImagePath TEXT,
    Age INTEGER, Height REAL, Weight REAL, Gender TEXT,
    EducationExtra TEXT, BodyType TEXT, Complexion TEXT, WorkDetails TEXT,
    HouseType TEXT, HouseSize TEXT, OtherProperty TEXT, ParentsAlive TEXT,
    FatherJob TEXT, MotherOccupation TEXT, Sisters INTEGER, MarriedSisters INTEGER,
    Brothers INTEGER, MarriedBrothers INTEGER, SiblingNumber INTEGER, Disability TEXT,
    MonthlyIncome TEXT, FirstWife TEXT, Children TEXT, RelationWithSeeker TEXT,
    ExtraInfo TEXT, SmokingDrugs TEXT, FullAddress TEXT,
    D_MaritalStatus TEXT, D_Age TEXT, D_Education TEXT, D_Religion TEXT, D_Maslak TEXT,
    D_Caste TEXT, D_Height TEXT, D_House TEXT, D_CityProvince TEXT, D_Mobile TEXT, D_WhatsApp TEXT, D_Reference TEXT,
    ReceivedBy TEXT, CheckedBy TEXT, CommissionNote TEXT, MainOffice TEXT,
    CreatedAt TEXT
);", con).ExecuteNonQuery();

            EnsureSettings(con);     // settings table & defaults
        }

        void EnsureSettings(SqliteConnection con)
        {
            new SqliteCommand("CREATE TABLE IF NOT EXISTS Settings(Key TEXT PRIMARY KEY, Val TEXT);", con).ExecuteNonQuery();

            void Seed(string k, string v)
            {
                var cmd = new SqliteCommand("INSERT OR IGNORE INTO Settings(Key,Val) VALUES(@k,@v);", con);
                cmd.Parameters.AddWithValue("@k", k);
                cmd.Parameters.AddWithValue("@v", v);
                cmd.ExecuteNonQuery();
            }

            Seed("MAIN_OFFICE", "Main Office: Gulzar Colony, Sialkot Road, Gujranwala");
            Seed("CASTE_OPTIONS", "Rajput;Jutt;Arain;Mughal;Syed;Sheikh;Pathan");
            Seed("CITY_OPTIONS", "Gujranwala;Sialkot;Lahore;Faisalabad;Multan;Karachi;Islamabad");
            Seed("RELIGION_OPTIONS", "Islam;Christian;Hindu");
            Seed("MASLAK_OPTIONS", "Barelvi;Deobandi;Ahl-e-Hadith;Shia");
        }

        string GetSetting(SqliteConnection con, string key)
        {
            var c = new SqliteCommand("SELECT Val FROM Settings WHERE Key=@k;", con);
            c.Parameters.AddWithValue("@k", key);
            return c.ExecuteScalar()?.ToString() ?? "";
        }
        void SetSetting(SqliteConnection con, string key, string val)
        {
            var c = new SqliteCommand("INSERT INTO Settings(Key,Val) VALUES(@k,@v) ON CONFLICT(Key) DO UPDATE SET Val=excluded.Val;", con);
            c.Parameters.AddWithValue("@k", key);
            c.Parameters.AddWithValue("@v", val);
            c.ExecuteNonQuery();
        }

        // =========================================================
        //                         DATA LOAD
        // =========================================================
        void LoadData(string search = "")
        {
            using var con = new SqliteConnection(ConnStr);
            con.Open();
            using var cmd = con.CreateCommand();
            cmd.CommandText =
@"SELECT Id,Name,FatherName,City,Phone,Education,Religion,Caste,Occupation,MaritalStatus,Address,ImagePath,
  Age,Height,Weight,Gender,EducationExtra,BodyType,Complexion,WorkDetails,HouseType,HouseSize,OtherProperty,
  ParentsAlive,FatherJob,MotherOccupation,Sisters,MarriedSisters,Brothers,MarriedBrothers,SiblingNumber,Disability,
  MonthlyIncome,FirstWife,Children,RelationWithSeeker,FullAddress,ExtraInfo,SmokingDrugs,
  D_MaritalStatus,D_Age,D_Education,D_Religion,D_Maslak,D_Caste,D_Height,D_House,D_CityProvince,D_Mobile,D_WhatsApp,D_Reference,
  ReceivedBy,CheckedBy,CommissionNote,MainOffice
  FROM Users
  WHERE @q='' OR (Name LIKE @like OR Phone LIKE @like OR City LIKE @like OR Caste LIKE @like OR Religion LIKE @like)
  ORDER BY Id DESC;";
            cmd.Parameters.AddWithValue("@q", search ?? "");
            cmd.Parameters.AddWithValue("@like", $"%{search}%");

            using var r = cmd.ExecuteReader();
            var dt = new DataTable();
            dt.Load(r);
            grid!.DataSource = dt;

            foreach (var col in new[]{
                "ImagePath","WorkDetails","OtherProperty","ExtraInfo","CommissionNote",
                "D_Reference","FullAddress","Address","MainOffice"})
                if (grid.Columns.Contains(col)) grid.Columns[col].Visible = false;
        }

        // =========================================================
        //                        CRUD
        // =========================================================
        void AddRecord()
        {
            if (!ValidateForm()) return;

            using var con = new SqliteConnection(ConnStr);
            con.Open();
            using var cmd = con.CreateCommand();
            cmd.CommandText = @"INSERT INTO Users
(Name,FatherName,City,Phone,Education,Religion,Caste,Occupation,MaritalStatus,Address,ImagePath,CreatedAt,
 Age,Height,Weight,Gender,EducationExtra,BodyType,Complexion,WorkDetails,HouseType,HouseSize,OtherProperty,
 ParentsAlive,FatherJob,MotherOccupation,Sisters,MarriedSisters,Brothers,MarriedBrothers,SiblingNumber,Disability,
 MonthlyIncome,FirstWife,Children,RelationWithSeeker,FullAddress,ExtraInfo,SmokingDrugs,
 D_MaritalStatus,D_Age,D_Education,D_Religion,D_Maslak,D_Caste,D_Height,D_House,D_CityProvince,D_Mobile,D_WhatsApp,D_Reference,
 ReceivedBy,CheckedBy,CommissionNote,MainOffice)
VALUES(@n,@f,@city,@p,@edu,@rel,@caste,@occ,@mar,@addr,@img,@t,
       @age,@h,@w,@gender,@eduPlus,@body,@comp,@work,@houseType,@houseSize,@otherProp,
       @parentsAlive,@fatherJob,@momJob,@sis,@sisM,@bro,@broM,@sibNo,@dis,
       @income,@firstWife,@children,@relSeeker,@fullAddr,@extra,@smoke,
       @d_ms,@d_age,@d_edu,@d_rel,@d_mas,@d_caste,@d_hgt,@d_house,@d_city,@d_mob,@d_wa,@d_ref,
       @recBy,@chkBy,@comm,@mainOffice);";

            AddCommonParams(cmd);
            cmd.Parameters.AddWithValue("@t", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.ExecuteNonQuery();

            LoadData(txtSearch!.Text);
            ClearForm();
            MessageBox.Show("ریکارڈ شامل ہو گیا");
        }

        void UpdateRecord()
        {
            if (grid!.CurrentRow == null) return;
            if (!ValidateForm()) return;

            var id = Convert.ToInt32(grid.CurrentRow.Cells["Id"].Value);
            using var con = new SqliteConnection(ConnStr);
            con.Open();
            using var cmd = con.CreateCommand();
            cmd.CommandText = @"UPDATE Users SET
Name=@n,FatherName=@f,City=@city,Phone=@p,Education=@edu,Religion=@rel,Caste=@caste,Occupation=@occ,MaritalStatus=@mar,Address=@addr,ImagePath=@img,
Age=@age,Height=@h,Weight=@w,Gender=@gender,EducationExtra=@eduPlus,BodyType=@body,Complexion=@comp,WorkDetails=@work,HouseType=@houseType,HouseSize=@houseSize,OtherProperty=@otherProp,
ParentsAlive=@parentsAlive,FatherJob=@fatherJob,MotherOccupation=@momJob,Sisters=@sis,MarriedSisters=@sisM,Brothers=@bro,MarriedBrothers=@broM,SiblingNumber=@sibNo,Disability=@dis,
MonthlyIncome=@income,FirstWife=@firstWife,Children=@children,RelationWithSeeker=@relSeeker,FullAddress=@fullAddr,ExtraInfo=@extra,SmokingDrugs=@smoke,
D_MaritalStatus=@d_ms,D_Age=@d_age,D_Education=@d_edu,D_Religion=@d_rel,D_Maslak=@d_mas,D_Caste=@d_caste,D_Height=@d_hgt,D_House=@d_house,D_CityProvince=@d_city,D_Mobile=@d_mob,D_WhatsApp=@d_wa,D_Reference=@d_ref,
ReceivedBy=@recBy,CheckedBy=@chkBy,CommissionNote=@comm,MainOffice=@mainOffice
WHERE Id=@id;";
            AddCommonParams(cmd);
            cmd.Parameters.AddWithValue("@id", id);
            cmd.ExecuteNonQuery();

            LoadData(txtSearch!.Text);
            MessageBox.Show("ریکارڈ اپڈیٹ ہو گیا");
        }

        void DeleteRecord()
        {
            if (grid!.CurrentRow == null) return;
            var id = Convert.ToInt32(grid.CurrentRow.Cells["Id"].Value);
            if (MessageBox.Show("حذف کریں؟", "Confirm", MessageBoxButtons.YesNo) == DialogResult.No) return;

            using var con = new SqliteConnection(ConnStr);
            con.Open();
            var cmd = con.CreateCommand();
            cmd.CommandText = "DELETE FROM Users WHERE Id=@id;";
            cmd.Parameters.AddWithValue("@id", id);
            cmd.ExecuteNonQuery();

            LoadData(txtSearch!.Text);
            ClearForm();
            MessageBox.Show("ریکارڈ حذف ہو گیا");
        }

        void AddCommonParams(SqliteCommand cmd)
        {
            cmd.Parameters.AddWithValue("@n", txtName!.Text);
            cmd.Parameters.AddWithValue("@f", txtFatherName!.Text);
            cmd.Parameters.AddWithValue("@city", cmbCity!.Text);
            cmd.Parameters.AddWithValue("@p", txtPhone!.Text);
            cmd.Parameters.AddWithValue("@edu", txtEducation!.Text);
            cmd.Parameters.AddWithValue("@rel", cmbReligion!.Text);
            cmd.Parameters.AddWithValue("@caste", cmbCaste!.Text);
            cmd.Parameters.AddWithValue("@occ", txtFatherJob!.Text); // Occupation = Father Job (label شدہ)
            cmd.Parameters.AddWithValue("@mar", cmbMarital!.Text);
            cmd.Parameters.AddWithValue("@addr", txtAddress!.Text);
            cmd.Parameters.AddWithValue("@img", SavePhotoIfAny());

            cmd.Parameters.AddWithValue("@age",   int.TryParse(txtAge!.Text, out var age) ? age : 0);
            cmd.Parameters.AddWithValue("@h",     double.TryParse(txtHeight!.Text, out var h) ? h : 0);
            cmd.Parameters.AddWithValue("@w",     double.TryParse(txtWeight!.Text, out var w) ? w : 0);
            cmd.Parameters.AddWithValue("@gender", cmbGender!.Text);
            cmd.Parameters.AddWithValue("@eduPlus", txtEducationExtra!.Text ?? "");
            cmd.Parameters.AddWithValue("@body",     txtBodyType!.Text ?? "");
            cmd.Parameters.AddWithValue("@comp",     txtComplexion!.Text ?? "");
            cmd.Parameters.AddWithValue("@work",     txtWorkDetails!.Text ?? "");
            cmd.Parameters.AddWithValue("@houseType",txtHouseType!.Text ?? "");
            cmd.Parameters.AddWithValue("@houseSize",txtHouseSize!.Text ?? "");
            cmd.Parameters.AddWithValue("@otherProp",txtOtherProperty!.Text ?? "");
            cmd.Parameters.AddWithValue("@parentsAlive", txtParentsAlive!.Text ?? "");
            cmd.Parameters.AddWithValue("@fatherJob", txtFatherJob!.Text ?? "");
            cmd.Parameters.AddWithValue("@momJob",    txtMotherJob!.Text ?? "");
            cmd.Parameters.AddWithValue("@sis",       int.TryParse(txtSisters!.Text, out var s1)? s1:0);
            cmd.Parameters.AddWithValue("@sisM",      int.TryParse(txtMarriedSisters!.Text, out var s2)? s2:0);
            cmd.Parameters.AddWithValue("@bro",       int.TryParse(txtBrothers!.Text, out var b1)? b1:0);
            cmd.Parameters.AddWithValue("@broM",      int.TryParse(txtMarriedBrothers!.Text, out var b2)? b2:0);
            cmd.Parameters.AddWithValue("@sibNo",     int.TryParse(txtSiblingNumber!.Text, out var sb)? sb:0);
            cmd.Parameters.AddWithValue("@dis",       txtDisability!.Text ?? "");
            cmd.Parameters.AddWithValue("@income",    txtMonthlyIncome!.Text ?? "");
            cmd.Parameters.AddWithValue("@firstWife", txtFirstWife!.Text ?? "");
            cmd.Parameters.AddWithValue("@children",  txtChildren!.Text ?? "");
            cmd.Parameters.AddWithValue("@relSeeker", txtRelationWithSeeker!.Text ?? "");
            cmd.Parameters.AddWithValue("@fullAddr",  txtAddress!.Text); // FullAddress = Address (same UI)
            cmd.Parameters.AddWithValue("@extra",     txtExtraInfo!.Text ?? "");
            cmd.Parameters.AddWithValue("@smoke",     txtSmokingDrugs!.Text ?? "");

            // Desired
            cmd.Parameters.AddWithValue("@d_ms",  txtD_MaritalStatus!.Text ?? "");
            cmd.Parameters.AddWithValue("@d_age", txtD_Age!.Text ?? "");
            cmd.Parameters.AddWithValue("@d_edu", txtD_Education!.Text ?? "");
            cmd.Parameters.AddWithValue("@d_rel", txtD_Religion!.Text ?? "");
            cmd.Parameters.AddWithValue("@d_mas", txtD_Maslak!.Text ?? "");
            cmd.Parameters.AddWithValue("@d_caste", txtD_Caste!.Text ?? "");
            cmd.Parameters.AddWithValue("@d_hgt", txtD_Height!.Text ?? "");
            cmd.Parameters.AddWithValue("@d_house", txtD_House!.Text ?? "");
            cmd.Parameters.AddWithValue("@d_city", txtD_CityProvince!.Text ?? "");
            cmd.Parameters.AddWithValue("@d_mob",  txtD_Mobile!.Text ?? "");
            cmd.Parameters.AddWithValue("@d_wa",   txtD_WhatsApp!.Text ?? "");
            cmd.Parameters.AddWithValue("@d_ref",  txtD_Reference!.Text ?? "");

            // Footer / Required
            cmd.Parameters.AddWithValue("@recBy",      txtReceivedBy!.Text);
            cmd.Parameters.AddWithValue("@chkBy",      txtCheckedBy!.Text);
            cmd.Parameters.AddWithValue("@comm",       txtCommissionNote!.Text ?? "");
            cmd.Parameters.AddWithValue("@mainOffice", txtMainOffice!.Text);
        }

        // =========================================================
        //                      Event Wiring
        // =========================================================
        void Grid_SelectionChanged(object? sender, EventArgs e)
        {
            if (grid!.CurrentRow == null) return;
            string V(string col) => grid.CurrentRow.Cells[col]?.Value?.ToString() ?? "";

            txtName!.Text          = V("Name");
            txtFatherName!.Text    = V("FatherName");
            cmbCity!.Text          = V("City");
            txtPhone!.Text         = V("Phone");
            txtEducation!.Text     = V("Education");
            cmbReligion!.Text      = V("Religion");
            cmbCaste!.Text         = V("Caste");
            txtFatherJob!.Text     = V("Occupation");
            cmbMarital!.Text       = V("MaritalStatus");
            txtAddress!.Text       = V("Address");

            txtAge!.Text           = V("Age");
            txtHeight!.Text        = V("Height");
            txtWeight!.Text        = V("Weight");
            cmbGender!.Text        = V("Gender");
            txtEducationExtra!.Text= V("EducationExtra");
            txtBodyType!.Text      = V("BodyType");
            txtComplexion!.Text    = V("Complexion");
            txtWorkDetails!.Text   = V("WorkDetails");
            txtHouseType!.Text     = V("HouseType");
            txtHouseSize!.Text     = V("HouseSize");
            txtOtherProperty!.Text = V("OtherProperty");
            txtParentsAlive!.Text  = V("ParentsAlive");
            txtMotherJob!.Text     = V("MotherOccupation");
            txtSisters!.Text       = V("Sisters");
            txtMarriedSisters!.Text= V("MarriedSisters");
            txtBrothers!.Text      = V("Brothers");
            txtMarriedBrothers!.Text=V("MarriedBrothers");
            txtSiblingNumber!.Text = V("SiblingNumber");
            txtDisability!.Text    = V("Disability");
            txtMonthlyIncome!.Text = V("MonthlyIncome");
            txtFirstWife!.Text     = V("FirstWife");
            txtChildren!.Text      = V("Children");
            txtRelationWithSeeker!.Text = V("RelationWithSeeker");
            txtExtraInfo!.Text     = V("ExtraInfo");
            txtSmokingDrugs!.Text  = V("SmokingDrugs");

            // Desired
            txtD_MaritalStatus!.Text = V("D_MaritalStatus");
            txtD_Age!.Text           = V("D_Age");
            txtD_Education!.Text     = V("D_Education");
            txtD_Religion!.Text      = V("D_Religion");
            txtD_Maslak!.Text        = V("D_Maslak");
            txtD_Caste!.Text         = V("D_Caste");
            txtD_Height!.Text        = V("D_Height");
            txtD_House!.Text         = V("D_House");
            txtD_CityProvince!.Text  = V("D_CityProvince");
            txtD_Mobile!.Text        = V("D_Mobile");
            txtD_WhatsApp!.Text      = V("D_WhatsApp");
            txtD_Reference!.Text     = V("D_Reference");

            // Footer
            txtReceivedBy!.Text      = V("ReceivedBy");
            txtCheckedBy!.Text       = V("CheckedBy");
            txtCommissionNote!.Text  = V("CommissionNote");
            txtMainOffice!.Text      = V("MainOffice");

            var img = V("ImagePath");
            if (!string.IsNullOrWhiteSpace(img) && File.Exists(img))
            { pic!.Image = SDImage.FromFile(img); imagePath = img; }
            else { pic!.Image = null; imagePath = ""; }
        }

        // =========================================================
        //                      Utilities
        // =========================================================
        bool ValidateForm()
        {
            if (string.IsNullOrWhiteSpace(txtName!.Text)) { MessageBox.Show("نام لازمی ہے"); return false; }
            if (string.IsNullOrWhiteSpace(txtPhone!.Text) || txtPhone!.Text.Length < 10 || txtPhone.Text.Length > 15)
            { MessageBox.Show("فون نمبر 10 تا 15 ہندسوں کا درج کریں"); return false; }
            if (string.IsNullOrWhiteSpace(txtReceivedBy!.Text)) { MessageBox.Show("وصول کنندہ لازمی ہے"); return false; }
            if (string.IsNullOrWhiteSpace(txtCheckedBy!.Text))  { MessageBox.Show("چیک کرنے والا لازمی ہے"); return false; }
            if (string.IsNullOrWhiteSpace(txtMainOffice!.Text)) { MessageBox.Show("Main Office لائن لازمی ہے"); return false; }
            return true;
        }

        void ClearForm()
        {
            foreach (Control c in Controls)
                ClearRecursive(c);
            imagePath = "";
            pic!.Image = null;
        }
        void ClearRecursive(Control c)
        {
            if (c is TextBox t && t != txtSearch) t.Text = "";
            if (c is ComboBox cb) cb.Text = "";
            foreach (Control k in c.Controls) ClearRecursive(k);
        }

        string SavePhotoIfAny()
        {
            try
            {
                if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath)) return "";
                var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "photos");
                Directory.CreateDirectory(dir);
                var dest = Path.Combine(dir, $"{Guid.NewGuid()}{Path.GetExtension(imagePath)}");
                File.Copy(imagePath, dest, true);
                return dest;
            }
            catch { return ""; }
        }

        void BrowseImage()
        {
            using var ofd = new OpenFileDialog { Filter = "Images|*.jpg;*.jpeg;*.png" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                imagePath = ofd.FileName;
                pic!.Image = SDImage.FromFile(imagePath);
            }
        }

        // ---------- CSV / Backup / Restore ----------
        void ExportCsv()
        {
            var sfd = new SaveFileDialog { Filter = "CSV|*.csv", FileName = "records.csv" };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            using var con = new SqliteConnection(ConnStr);
            con.Open();
            using var cmd = con.CreateCommand();
            cmd.CommandText = "SELECT * FROM Users ORDER BY Id DESC";
            using var r = cmd.ExecuteReader();

            using var sw = new StreamWriter(sfd.FileName, false, System.Text.Encoding.UTF8);
            for (int i = 0; i < r.FieldCount; i++){ if (i>0) sw.Write(","); sw.Write(r.GetName(i)); }
            sw.WriteLine();
            while (r.Read())
            {
                for (int i = 0; i < r.FieldCount; i++)
                {
                    if (i>0) sw.Write(",");
                    var val = r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString()?.Replace("\"","\"\"");
                    sw.Write($"\"{val}\"");
                }
                sw.WriteLine();
            }
            MessageBox.Show("CSV تیار ہو گیا");
        }

        void BackupDb()
        {
            var sfd = new SaveFileDialog { Filter = "SQLite DB|*.db", FileName = $"backup_{DateTime.Now:yyyyMMdd_HHmm}.db" };
            if (sfd.ShowDialog() != DialogResult.OK) return;
            File.Copy(DbFile, sfd.FileName, true);
            MessageBox.Show("Backup مکمل");
        }

        void RestoreDb()
        {
            var ofd = new OpenFileDialog { Filter = "SQLite DB|*.db" };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            var bak = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"before_restore_{DateTime.Now:yyyyMMdd_HHmm}.db");
            File.Copy(DbFile, bak, true);

            File.Copy(ofd.FileName, DbFile, true);
            EnsureDb();
            LoadData();
            MessageBox.Show("Restore مکمل؛ ایپ ریفریش ہو گئی");
        }

        // ---------- Print / PDF ----------
        void PrintRecord()
        {
            if (grid!.CurrentRow == null) { MessageBox.Show("کوئی ریکارڈ منتخب کریں"); return; }

            var doc = new PrintDocument();
            doc.PrintPage += (s, e) =>
            {
                float x = 50, y = 60, lh = 26;
                using var title = new Font("Segoe UI", 16, FontStyle.Bold);
                using var f = new Font("Segoe UI", 10);

                e.Graphics.DrawString("Matrimonial Record / ریکارڈ", title, Brushes.Black, x, y); y += 40;

                void Line(string label, string val){ e.Graphics.DrawString($"{label}: {val}", f, Brushes.Black, x, y); y += lh; }

                Line("Name / نام", txtName!.Text);
                Line("Father / والدیت", txtFatherName!.Text);
                Line("Phone", txtPhone!.Text);
                Line("City / شہر", cmbCity!.Text);
                Line("Caste / ذات", cmbCaste!.Text);
                Line("Religion / مذہب", cmbReligion!.Text);
                Line("Maslak / مسلک", cmbMaslak!.Text);
                Line("Age/Height/Weight", $"{txtAge!.Text} / {txtHeight!.Text} / {txtWeight!.Text}");
                Line("Education", txtEducation!.Text);
                Line("Work", txtWorkDetails!.Text);

                y += 10;
                if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                    e.Graphics.DrawImage(SDImage.FromFile(imagePath), new Rectangle(620, 80, 180, 220));

                // Footer bar
                y += 20;
                e.Graphics.FillRectangle(Brushes.Black, new RectangleF(40, y, e.PageBounds.Width-80, 24));
                e.Graphics.DrawString(txtMainOffice!.Text, new Font("Segoe UI", 10, FontStyle.Bold), Brushes.White, 50, y+4);
            };
            using var preview = new PrintPreviewDialog { Document = doc, Width = 1000, Height = 700 };
            preview.ShowDialog();
        }

        void ExportPdf()
        {
            if (grid!.CurrentRow == null) { MessageBox.Show("کوئی ریکارڈ منتخب کریں"); return; }
            var sfd = new SaveFileDialog { Filter = "PDF|*.pdf", FileName = "Record.pdf" };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            QuestPDF.Settings.License = LicenseType.Community;

            Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Margin(30);
                    page.Header().Text("Matrimonial Record / ریکارڈ")
                        .SemiBold().FontSize(20).FontColor(Colors.Blue.Medium);

                    page.Content().Column(col =>
                    {
                        col.Item().Table(t =>
                        {
                            t.ColumnsDefinition(c => { c.ConstantColumn(180); c.RelativeColumn(); });

                            void Row(string k, string v){ t.Cell().Element(Key).Text(k); t.Cell().Element(Val).Text(v); }
                            IContainer Key(IContainer x)=>x.Background(Colors.Grey.Lighten3).Padding(6);
                            IContainer Val(IContainer x)=>x.BorderBottom(1).Padding(6);

                            Row("Name / نام", txtName!.Text);
                            Row("Father / والدیت", txtFatherName!.Text);
                            Row("Phone", txtPhone!.Text);
                            Row("City / شہر", cmbCity!.Text);
                            Row("Caste / ذات", cmbCaste!.Text);
                            Row("Religion / مذہب", cmbReligion!.Text);
                            Row("Maslak / مسلک", cmbMaslak!.Text);
                            Row("Age / عمر", txtAge!.Text);
                            Row("Height / قد", txtHeight!.Text);
                            Row("Weight / وزن", txtWeight!.Text);
                            Row("Education", txtEducation!.Text);
                            Row("Extra Education", txtEducationExtra!.Text);
                            Row("Body Type", txtBodyType!.Text);
                            Row("Complexion", txtComplexion!.Text);
                            Row("Work", txtWorkDetails!.Text);
                            Row("Parents Alive", txtParentsAlive!.Text);
                            Row("Father Job", txtFatherJob!.Text);
                            Row("Mother Job", txtMotherJob!.Text);
                            Row("Income", txtMonthlyIncome!.Text);
                            Row("Address / پتہ", txtAddress!.Text);
                        });

                        if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                            col.Item().PaddingTop(8).AlignRight().Width(180).Height(220)
                               .Image(imagePath).FitArea();

                        // Footer black bar
                        col.Item().PaddingTop(8).Background(Colors.Black).Padding(6)
                           .Text(txtMainOffice!.Text).FontColor(Colors.White).SemiBold().AlignCenter();
                    });

                });
            }).GeneratePdf(sfd.FileName);

            MessageBox.Show("PDF بن گیا!");
        }

        // ---------- Theme & Options ----------
        void PickTheme()
        {
            using var cd = new ColorDialog();
            if (cd.ShowDialog() == DialogResult.OK)
            {
                userAccent = cd.Color;
                ApplyTheme();
            }
        }
        void ApplyTheme()
        {
            var baseColor = userAccent ?? Color.FromArgb(245, 248, 255);
            BackColor = baseColor;
            foreach (Control c in Controls)
                if (c is Button b) b.BackColor = ControlPaint.Light(baseColor);
        }

        void OpenOptionsDialog()
        {
            using var con = new SqliteConnection(ConnStr); con.Open();
            string caste    = GetSetting(con,"CASTE_OPTIONS");
            string city     = GetSetting(con,"CITY_OPTIONS");
            string rel      = GetSetting(con,"RELIGION_OPTIONS");
            string maslak   = GetSetting(con,"MASLAK_OPTIONS");
            string office   = GetSetting(con,"MAIN_OFFICE");

            var f = new Form { Text="Options", Width=740, Height=540, StartPosition=FormStartPosition.CenterParent };
            var tCaste  = new TextBox{ Left=20, Top=40, Width=680, Text=caste };
            var tCity   = new TextBox{ Left=20, Top=100, Width=680, Text=city };
            var tRel    = new TextBox{ Left=20, Top=160, Width=680, Text=rel };
            var tMas    = new TextBox{ Left=20, Top=220, Width=680, Text=maslak };
            var tOffice = new TextBox{ Left=20, Top=300, Width=680, Text=office };
            f.Controls.AddRange(new Control[]{
                L("CASTE_OPTIONS (semicolon ;)"), tCaste,
                L("CITY_OPTIONS (;)"), tCity,
                L("RELIGION_OPTIONS (;)"), tRel,
                L("MASLAK
