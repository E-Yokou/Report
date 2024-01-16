using System;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace CreateReport
{
    public partial class Report : Form
    {
        public Report()
        {
            InitializeComponent();
            textBox_charge.Text = "3.1. Поставщик обязан:\r\n3.1.1. отгружать товар в адрес Получателя указанным транспортом в согласованные сроки;\r\n3.1.2. извещать надлежащим образом Получателя об отправке товара, а также направлять ему другие извещения, требующиеся ему для осуществления обычно необходимых мер для принятия поставки товара;\r\n3.1.3. предоставлять Покупателю транспортные и сопроводительные документы;\r\n3.1.4. за свой счет обеспечить упаковку и тару, необходимую для поставки товара;\r\n3.1.5. в случае недопоставки товаров в отдельном периоде поставки, восполнить недопоставленное количество товаров в следующем периоде (периодах) в пределах срока действия настоящего договора.\r\n3.2. Покупатель (Получатель) обязан\r\n3.2.1. оплатить поставляемые товары с соблюдением порядка и формы расчетов, предусмотренных настоящим договором;\r\n3.2.2. совершить все необходимые действия, обеспечивающие принятие товаров, поставляемых в соответствии с настоящим договором;\r\n3.2.3. в разумный срок проверить количество и качество принятых товаров и о выявленных несоответствиях или недостатках незамедлительно письменно уведомить Поставщика;\r\n3.2.4. возвратить Поставщику многооборотную тару и средства пакетирования, в которых поступил товар, в месте отгрузки во время следующей поставки товаров или в любое другое время по требованию Поставщика.\r\n3.3. Покупатель (Получатель) вправе отказаться от оплаты товаров ненадлежащего качества и некомплектных товаров, а если такие товары оплачены, потребовать возврата уплаченных сумм впредь до устранения недостатков и доукомплектования товаров либо их замены.\r\n";
            textBox_risks.Text = "4.1. Поставщик несет все риски, потери или повреждения товара до момента его поставки Получателю.\r\n4.2. Получатель несет все риски, потери или повреждения товара с момента его получения.\r\n";
            textBox_responsibility.Text = "6.1. В случае существенного нарушения требований к качеству товара Поставщик обязан по выбору Покупателя вернуть ему уплаченную за товар сумму или заменить товар ненадлежащего качества товаром, соответствующим договору.\r\n6.2. За недопоставку или просрочку поставки товаров Поставщик уплачивает Покупателю неустойку в размере [значение] % от стоимости всей партии товаров за каждый день просрочки до фактического исполнения обязательства.\r\n6.3. За несвоевременную оплату переданного в соответствии с настоящим договором товара Покупатель уплачивает Поставщику неустойку в размере [значение] % от суммы задолженности за каждый день просрочки.\r\n";
            textBox_contract_term.Text = "7.1. Настоящий договор составлен в двух аутентичных экземплярах по одному для каждой из Сторон.\r\n7.2. Настоящий договор вступает в силу с момента его подписания и действует до [число, месяц, год].\r\n7.3. В случае, если ни одна из Сторон после истечения срока действия Договора не заявит о его расторжении, то договор пролонгируется на тех же условиях на [срок].\r\n";

            dataGridView_product.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            label_current_date.Text = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToLongTimeString();
        }

        private void BUT_add_product_Click(object sender, EventArgs e)
        {
            string value1 = textBox_product.Text;
            string value2 = textBox_count.Text;

            dataGridView_product.Rows.Add(value1, value2);
        }

        private void BUT_create_report_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Word документ|*.doc";
            saveFileDialog1.Title = "Save to Word";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (saveFileDialog1.FileName != "")
                {
                    Word.Application wordApp = new Word.Application();

                    wordApp.Visible = false;

                    Word.Document doc = wordApp.Documents.Add();
                    Word.Range range = doc.Range(0, 0);

                    range.Text = $"Адрес заключения договора: {textBox1.Text}\n";
                    range.Text += $"Поставщик: {textBox2.Text}\n";
                    range.Text += $"Получатель: {textBox3.Text}\n";
                    range.Text += $"Периодичность поставки: {textBox4.Text}\n";
                    range.Text += $"Дата поставки товара: {dateTimePicker1.Value.Day}.{dateTimePicker1.Value.Month}.{dateTimePicker1.Value.Year}\n";

                    range.Text += $"Вид транспорта: {comboBox1.SelectedText}\n";

                    range.Text += $"Права и обязанности сторон:\n {textBox_charge.Text}";
                    range.Text += $"Переход рисков, связанных с товаром:\n {textBox_risks.Text}";
                    range.Text += $"Ответственность сторон:\n {textBox_responsibility.Text}";
                    range.Text += $"Срок и прорядок выполнения договора:\n {textBox_contract_term.Text}";

                    range.Text += $"------------------------\nТовары:\n";

                    Word.Paragraph para;
                    foreach (DataGridViewRow row in dataGridView_product.Rows)
                    {
                        para = doc.Content.Paragraphs.Add();
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            // Товар
                            string product = row.Cells[0].Value == null ? string.Empty : row.Cells[0].Value.ToString();
                            para.Range.Text += "Товар: " + product + " ";

                            // Количество
                            string quantity = row.Cells[1].Value == null ? string.Empty : row.Cells[1].Value.ToString();
                            para.Range.Text += "Количество: " + quantity + " ";
                        }
                        para.Range.InsertParagraphAfter();
                    }

                    doc.SaveAs2(saveFileDialog1.FileName);
                    doc.Close();
                    wordApp.Quit();
                }
            }

        }
    }
}
