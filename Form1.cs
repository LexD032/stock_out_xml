using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Net;
//using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace WindowsFormsApp1
{
    public partial class SetActiv : Form
    {
            public int CarrentRowDoc = 0;
            public int CarrentRowItem = 0;
            DataSet dsItems = new DataSet();
        

        public SetActiv()
        {

            InitializeComponent();
            textBox1.Text = "";
        }

        public static DataSet GetSelect(string strSQL)
        {
            string strCon = GetConnectionString();

            FbConnectionStringBuilder fb_con = new FbConnectionStringBuilder();
            fb_con.Charset = "WIN1251";
            fb_con.UserID = "SYSDBA";
            fb_con.Password = "masterkey";
            fb_con.Database = @strCon;
            fb_con.ServerType = 0;
            //создаем подключение
            var fb = new FbConnection(fb_con.ToString());
            DataSet ds = new DataSet();
            try
            {
                fb.Open();
                FbDataAdapter myAdapter = new FbDataAdapter(strSQL, fb);
                myAdapter.Fill(ds);
            }
            catch 
            {
               MessageBox.Show("Ошибка подключения к БД. ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            fb.Close(); // по правилам хорошего тона ....
            return (ds);

            //    fb.Open();
            //    var com = fb.CreateCommand();
            //    com.CommandText = strSQL;
            //    com.ExecuteNonQuery();
            //    fb.Close();
            //    }



        }


        public static string GetConnectionString()
        {
            string path1 = "       ";
            string strConnect = "   ";
            if (System.IO.File.Exists(@"c:\IApteka\IApteka.ini"))
            { path1 = @"c:\IApteka\IApteka.ini"; }
            if (System.IO.File.Exists(@"D:\IApteka\IApteka.ini"))
            { path1 = @"D:\IApteka\IApteka.ini"; }
            if (System.IO.File.Exists(@"E:\IApteka\IApteka.ini"))
            { path1 = @"E:\IApteka\IApteka.ini"; }
            foreach (string line in System.IO.File.ReadLines(path1))
            {
                if (line.IndexOf("Path") == 0)
                {
                    strConnect = line.Substring(5, line.Length - 5);
                }

            }
            return (strConnect);
        }


        private void button1_Click(object sender, EventArgs e)
        {
            string dt_1 = dateTimePicker1.ToString();
            int k_1 = dt_1.IndexOf("202") - 6;
            dt_1 = dt_1.Substring(k_1, 10);

            string SqlString = @"
           select d.docno, d.docdate,dp.dep_name,  d.doc_id
           from docs d, department dp
           where
                d.docdate ='" + dt_1 + @"' 
                and d.doctype =19    and
                dp.dep_id = d.dep_id";

            DataSet dsDocs = GetSelect(SqlString);
            int count_str = dsDocs.Tables[0].Rows.Count;
            if (count_str == 0)
            {
                return;
            }

            BindingSource bsDocs = new BindingSource();
            bsDocs.DataSource = dsDocs.Tables[0];
            dgvDocs.ReadOnly = true;
            dgvDocs.DataSource = bsDocs;
            dgvDocs.Columns[0].Width = 70;
            dgvDocs.Columns[1].Width = 90;
            dgvDocs.Columns[2].Width = 450;
            dgvDocs.Columns[0].HeaderText = "Номер";
            dgvDocs.Columns[1].HeaderText = "Дата";
            dgvDocs.Columns[2].HeaderText = "Наименование склада";
        }


        private void button2_Click(object sender, EventArgs e)
        {
            Int32 selectedCellCount = dgvDocs.GetCellCount(DataGridViewElementStates.Selected);

            if ( selectedCellCount == 0)
            {
                MessageBox.Show("Выберите строку в списке документов.");
                return;
            }

            Int32 name_len = textBox1.Text.Trim().Length;
            if (name_len == 0)
            {
                MessageBox.Show("Выберите папку для выгрузки.");
                return;
            }


            XmlDocument xmlDoc = new XmlDocument();

            XmlProcessingInstruction Header = xmlDoc.CreateProcessingInstruction("xml", "version=\"1.0\" encoding=\"windows-1251\" ");
            xmlDoc.AppendChild(Header);

            XmlElement TageMassage = xmlDoc.CreateElement("PACKET");
            TageMassage.SetAttribute("TYPE", "12");
            TageMassage.SetAttribute("ID","1");                                  //  Должен быть номер накладной
            TageMassage.SetAttribute("NAME", "Электронная накладная");
            TageMassage.SetAttribute("FROM", "ГУПБрянскфармация");
            xmlDoc.AppendChild(TageMassage);

            XmlElement TageHead = xmlDoc.CreateElement("SUPPLY");
            TageMassage.AppendChild(TageHead);

            XmlElement INVOICE_NUM = xmlDoc.CreateElement("INVOICE_NUM");
            TageHead.AppendChild(INVOICE_NUM);
            XmlText INVOICE_NUM_ = xmlDoc.CreateTextNode("1");                     //  Должен быть номер накладной
            INVOICE_NUM.AppendChild(INVOICE_NUM_);

            XmlElement INVOICE_DATE = xmlDoc.CreateElement("INVOICE_DATE");
            TageHead.AppendChild(INVOICE_DATE);
            DateTime dt = DateTime.Today;
            XmlText INVOICE_DATE_ = xmlDoc.CreateTextNode(dt.ToShortDateString());   // Здесь должна быть текущая дата
            INVOICE_DATE.AppendChild(INVOICE_DATE_);

            XmlElement DEP_ID_Elem = xmlDoc.CreateElement("DEP_ID");
            TageHead.AppendChild(DEP_ID_Elem);
            XmlText DEP_ID_ = xmlDoc.CreateTextNode("123");   // тут код отдела (склада)
            DEP_ID_Elem.AppendChild(DEP_ID_);

            XmlElement ORDER_ID = xmlDoc.CreateElement("ORDER_ID");
            TageHead.AppendChild(ORDER_ID);
            XmlText ORDER_ID_ = xmlDoc.CreateTextNode("0");
            ORDER_ID.AppendChild(ORDER_ID_);

            XmlElement ABSACCEPT = xmlDoc.CreateElement("ABSACCEPT");
            TageHead.AppendChild(ABSACCEPT);
            XmlText ABSACCEPT_ = xmlDoc.CreateTextNode("1");
            ABSACCEPT.AppendChild(ABSACCEPT_);

            XmlElement ITEM__List = xmlDoc.CreateElement("ITEMS");
            TageHead.AppendChild(ITEM__List);



            // string id_doc = dgvDocs.CurrentRow.Cells[3].Value;
            DataGridViewRow selectedRow = dgvDocs.SelectedRows[0];
            string id_doc = selectedRow.Cells[3].Value.ToString();

//            string id_doc = dgvDocs.CurrentRow.Cells[3].ToString();
            string SqlString = @"
            select  di.num,m.med_id, m.med_name,m.vendor_id as vendor_code,v.vendor_name as vendor_name, m.country_id as country_code,
                    c.country_name as cantry_name, cast(di.qtty as float)/cast(di.divisor as float) as klv, cast(i.vprice as float), cast(i.rprice as float), cast(i.sprice as float), i.nds,
                    cast(i.sndssume as float), cast(di.qtty as float)/cast(di.divisor as float) * cast(i.rprice as float) as summprod ,i.seria, cast(i.valid_date as date) as valid,  i.gtd, i.reg_sert_num, cast(i.reg_price as float),ean13(i.vbarcode)
            from    docs d, docitem di, items i, parties p, partners pt, medicine m, vendor v, country c
            where   d.doc_id      = " + id_doc + @"      and
                    d.doc_id      = di.doc_id    and
                    i.iid         = di.iid       and
                    p.part_id     = i.part_id    and
                    pt.partner_id = p.partner_id and
                    m.med_id      = i.med_id     and
                    v.vendor_id   = m.vendor_id  and
                    c.country_id  = m.country_id";
            
            DataSet dsDI = GetSelect(SqlString);

            int count_str = dsDI.Tables[0].Rows.Count;
            if (count_str == 0) { return; }

            progressBar1.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            progressBar1.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            progressBar1.Maximum = dsDI.Tables[0].Rows.Count;
            // Set the initial value of the ProgressBar.
            progressBar1.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            progressBar1.Step = 1;


            // Тут начался Цикл по номенклатуре
            foreach (DataRow row in dsDI.Tables[0].Rows)
            {
                string nom = row.ItemArray[0].ToString();

                XmlElement TageITEM = xmlDoc.CreateElement("ITEM");
                ITEM__List.AppendChild(TageITEM);

                XmlElement CODE = xmlDoc.CreateElement("CODE");
                TageITEM.AppendChild(CODE);
                XmlText CODE_ = xmlDoc.CreateTextNode(row.ItemArray[1].ToString().Trim());
                CODE.AppendChild(CODE_);

                XmlElement NAME = xmlDoc.CreateElement("NAME");
                TageITEM.AppendChild(NAME);
                XmlText NAME_ = xmlDoc.CreateTextNode(row.ItemArray[2].ToString().Trim());
                NAME.AppendChild(NAME_);

                XmlElement VENDOR_CODE = xmlDoc.CreateElement("VENDOR_CODE");
                TageITEM.AppendChild(VENDOR_CODE);
                XmlText VENDOR_CODE_ = xmlDoc.CreateTextNode(row.ItemArray[3].ToString().Trim());
                VENDOR_CODE.AppendChild(VENDOR_CODE_);

                XmlElement VENDOR = xmlDoc.CreateElement("VENDOR");
                TageITEM.AppendChild(VENDOR);
                XmlText VENDOR_NAME_ = xmlDoc.CreateTextNode(row.ItemArray[4].ToString().Trim());
                VENDOR.AppendChild(VENDOR_NAME_);

                XmlElement COUNTRY_CODE = xmlDoc.CreateElement("COUNTRY_CODE");
                TageITEM.AppendChild(COUNTRY_CODE);
                XmlText COUNTRY_CODE_ = xmlDoc.CreateTextNode(row.ItemArray[5].ToString().Trim());  
                COUNTRY_CODE.AppendChild(COUNTRY_CODE_);

                XmlElement COUNTRY = xmlDoc.CreateElement("COUNTRY");
                TageITEM.AppendChild(COUNTRY);
                XmlText COUNTRY_ = xmlDoc.CreateTextNode(row.ItemArray[6].ToString().Trim());
                COUNTRY.AppendChild(COUNTRY_);

                XmlElement QTTY = xmlDoc.CreateElement("QTTY");
                TageITEM.AppendChild(QTTY);
                XmlText QTTY_ = xmlDoc.CreateTextNode(row.ItemArray[7].ToString().Trim());
                QTTY.AppendChild(QTTY_);

                // Цена производителя
                XmlElement VPRICE = xmlDoc.CreateElement("VPRICE");
                TageITEM.AppendChild(VPRICE);
                XmlText VPRICE_ = xmlDoc.CreateTextNode(row.ItemArray[8].ToString().Trim());
                VPRICE.AppendChild(VPRICE_);

                // Цена продажи
                XmlElement PRPRICE = xmlDoc.CreateElement("RPRICE");
                TageITEM.AppendChild(PRPRICE);
                XmlText PRPRICE_ = xmlDoc.CreateTextNode(row.ItemArray[9].ToString().Trim());
                PRPRICE.AppendChild(PRPRICE_);

                //Цена поставщика
                XmlElement SPRICE = xmlDoc.CreateElement("SPRICE");
                TageITEM.AppendChild(SPRICE);
                XmlText SPRICE_ = xmlDoc.CreateTextNode(row.ItemArray[10].ToString().Trim());
                SPRICE.AppendChild(SPRICE_);

                // Ставка НДС
                XmlElement NDS = xmlDoc.CreateElement("NDS");
                TageITEM.AppendChild(NDS);
                XmlText NDS_ = xmlDoc.CreateTextNode(row.ItemArray[11].ToString().Trim());
                NDS.AppendChild(NDS_);

                // Сумма НДС поставщика по строке
                XmlElement SNDSSUM = xmlDoc.CreateElement("SNDSSUM");
                TageITEM.AppendChild(SNDSSUM);
                XmlText SNDSSUM_ = xmlDoc.CreateTextNode(row.ItemArray[12].ToString().Trim());
                SNDSSUM.AppendChild(SNDSSUM_);

                // Сумма продажи с НДС по строке                          
                XmlElement SALLSUM = xmlDoc.CreateElement("SALLSUM");
                TageITEM.AppendChild(SALLSUM);
                XmlText summprod = xmlDoc.CreateTextNode(row.ItemArray[13].ToString());
                SALLSUM.AppendChild(summprod);

                XmlElement SERIA = xmlDoc.CreateElement("SERIA");
                TageITEM.AppendChild(SERIA);
                XmlText SERIA_ = xmlDoc.CreateTextNode(row.ItemArray[14].ToString());
                SERIA.AppendChild(SERIA_);

                string valid_date = row.ItemArray[15].ToString();
                if (valid_date.Trim().Length > 0) { valid_date = valid_date.Substring(0, 10); }
                XmlElement VALID_DATE = xmlDoc.CreateElement("VALID_DATE");
                TageITEM.AppendChild(VALID_DATE);
                XmlText VALID_DATE_ = xmlDoc.CreateTextNode(valid_date);
                VALID_DATE.AppendChild(VALID_DATE_);

                XmlElement GTD = xmlDoc.CreateElement("GTD");
                TageITEM.AppendChild(GTD);
                XmlText GTD_ = xmlDoc.CreateTextNode(row.ItemArray[16].ToString());
                GTD.AppendChild(GTD_);

                XmlElement SERT_NUM = xmlDoc.CreateElement("SERT_NUM");
                SERT_NUM = xmlDoc.CreateElement("SERT_NUM");
                TageITEM.AppendChild(SERT_NUM);
                XmlText SERT_NUM_ = xmlDoc.CreateTextNode(row.ItemArray[17].ToString());
                SERT_NUM.AppendChild(SERT_NUM_);


                XmlElement VENDORBARCODE = xmlDoc.CreateElement("VENDORBARCODE");
                TageITEM.AppendChild(VENDORBARCODE);
                XmlText VENDORBARCODE_ = xmlDoc.CreateTextNode(row.ItemArray[19].ToString().Trim());
                VENDORBARCODE.AppendChild(VENDORBARCODE_);

                XmlElement REG_PRICE = xmlDoc.CreateElement("REG_PRICE");
                TageITEM.AppendChild(REG_PRICE);
                XmlText REG_PRICE_ = xmlDoc.CreateTextNode(row.ItemArray[19].ToString());
                REG_PRICE.AppendChild(REG_PRICE_);

                //XmlElement ISGV = xmlDoc.CreateElement("ISGV");
                //TageITEM.AppendChild(ISGV);
                //if (row.ItemArray[18].ToString() )
                //{
                //    ISGV_ = XmlText xmlDoc.CreateTextNode("1");
                //}   else
                //{
                //    ISGV_ = XmlText xmlDoc.CreateTextNode("0");
                //}
                //ISGV.AppendChild(ISGV_);

                //Åñëè ÏóñòîåÇíà÷åíèå(Ò.GTIN) = 0 Òîãäà
                //XmlElement GTIN = xmlDoc.CreateElement("GTIN");
                //TageITEM.AppendChild(GTIN);
                //XmlElement GTIN_ = xmlDoc.CreateTextNode(row.ItemArray[6].ToString());
                //GTIN.AppendChild(GTIN_);
                //ÊîíåöÅñëè;

                DataGridViewRow selectedRow1 = dgvDocs.SelectedRows[0];
                string id_doc1 = selectedRow1.Cells[3].Value.ToString();

                string qlString1 = @"
                select ds.sgtin
                from DOCS D, DOCITEM DI, docitem_sgtin ds
                where D.DOC_ID = "+ id_doc1 + @" and
                di.num = "+ nom +@" and
                DI.DOC_ID = D.DOC_ID and
                ds.doc_id = d.doc_id and
                ds.iid = di.iid
                ";

                DataSet dssgtin = GetSelect(qlString1);
                if (dssgtin.Tables[0].Rows.Count > 0 )
                {

                    XmlElement SGTINS = xmlDoc.CreateElement("SGTINS");
                    TageITEM.AppendChild(SGTINS);

                    // Начало цикл вставить SGTINы
                    foreach (DataRow row_sgt in dssgtin.Tables[0].Rows)
                    {
                        XmlElement SGTIN = xmlDoc.CreateElement("SGTIN");
                        SGTINS.AppendChild(SGTIN);
                        XmlText SGTIN_ = xmlDoc.CreateTextNode(row_sgt.ItemArray[0].ToString());
                        SGTIN.AppendChild(SGTIN_);

                    }
                    // конец цикла вставки sgtin-ов

                }

                progressBar1.PerformStep();
                // Тут должен заканчиваться конец цмкла по номенклатуре
            }
           // xmlDoc.Save(OutxmlDoc);
            xmlDoc.Save(textBox1.Text+@"\"+ textBox2.Text.Trim());
            MessageBox.Show(this, "Вывела. ", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }




        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = FBD.SelectedPath;
            }
        }
    }
}
