using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;


namespace OLX_parser
{
    public partial class Form1 : Form
    {


        CancellationTokenSource s_cts;
        string path;
        public Form1()
        {
            InitializeComponent();
            toolTip1.SetToolTip(this.searchButton, "Парсер не активен");
            toolTip1.SetToolTip(this.stopButton, "Парсер не активен");
        }

        private async void searchButton_Click(object sender, EventArgs e)
        {

            if(s_cts != null)
            {
                MessageBox.Show("Парсер уже активен!");
                return;
            }


            string searchQuery = requestBox.Text;

            string region = regionsComboBox.Text;

            string rubric = RubricComboBox.Text;

            string subrubric = SubrubricComboBox.Text;

            Parse parseObj = new Parse(searchQuery, region, rubric, subrubric, dataGridView1);
          

            if (s_cts == null)
            {
                s_cts = new CancellationTokenSource();
                
                toolTip1.SetToolTip(this.searchButton, "Парсер работает");
                toolTip1.SetToolTip(this.stopButton, "Парсер работает");
                await Task.Run(() => parseObj.parseAsync(s_cts.Token));                   
                
            }
            else
            {
                MessageBox.Show("Парсер уже работает, если вы хотите остановить операцию нажмите кнопку стоп");
            }
        }

        private void stopButton_Click(object sender, EventArgs e)
        {
            if (s_cts != null)
            {
                s_cts.Cancel();
                s_cts = null;
                toolTip1.SetToolTip(this.searchButton, "Парсер не активен");
                toolTip1.SetToolTip(this.stopButton, "Парсер не активен");
            }
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
        }

        private async void saveButton_Click(object sender, EventArgs e)
        {
            path = "c:\\parse1.xlsx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path = saveFileDialog1.FileName;
                await Task.Run(() => saveDataToExcel());
            }   
        }

        private void saveDataToExcel()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            

            if (xlApp == null)
            {
                MessageBox.Show("Excel не установен! Установите сначала Excel");
                return;
            }


            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);



            xlWorkSheet.Cells[1, 1] = "Название";
            xlWorkSheet.Cells[1, 2] = "Описание";
            xlWorkSheet.Cells[1, 3] = "Ссылка";
            xlWorkSheet.Cells[1, 4] = "Цена";
            xlWorkSheet.Cells[1, 5] = "Id";
            xlWorkSheet.Cells[1, 6] = "Опубликовано";
            xlWorkSheet.Cells[1, 7] = "Картинка";



            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    xlWorkSheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }

            }

            //Here saving the file in xlsx
            xlWorkBook.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
            misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            MessageBox.Show("Excel файл создан " + path);
        }

        private void changedRubric(object sender, EventArgs e)
        {
            

            switch(RubricComboBox.Text)
            {
                case "Все рубрики":
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Text = "";
                    break;
                case "Детский мир":
                    SubrubricComboBox.Text = "Всё в рубрике Детский мир";
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Items.AddRange(new object[] {
                    "Всё в рубрике Детский мир",
                    "Детская одежда",
                    "Детская обувь",
                    "Детские коляски",
                    "Детские автокресла",
                    "Детская мебель",
                    "Игрушки",
                    "Детский транспорт",
                    "Товары для кормления",
                    "Товары для школьников",
                    "Прочие детские товары"});
                    break;
                case "Недвижимость":
                    SubrubricComboBox.Text = "Всё в рубрике Недвижимость";
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Items.AddRange(new object[] {
                    "Всё в рубрике Недвижимость",
                    "Квартиры",
                    "Комнаты",
                    "Дома",
                    "Земля",
                    "Коммерческая недвижимость ",
                    "Гаражи, парковки",
                    "Посуточная аренда жилья",
                    "Недвижимость за рубежом"});
                    break;
                case "Авто":                    
                    SubrubricComboBox.Text = "Всё в рубрике Авто";
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Items.AddRange(new object[] {
                    "Всё в рубрике Авто",
                    "Легковые автомобили",
                    "Автомобили из Польши",
                    "Грузовые автомобили",
                    "Грузовики и спецтехника из Польши",
                    "Автобусы",
                    "Мото",
                    "Спецтехника",
                    "Сельхозтехника",
                    "Водный транспорт",
                    "Воздушный транспорт",
                    "Прицепы / дома на колесах",
                    "Другой транспорт"});
                    break;
                case "Запчасти для транспорта":
                    //"Всё в рубрике Запчасти для транспорта"
                    SubrubricComboBox.Text = "Всё в рубрике Запчасти для транспорта";
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Items.AddRange(new object[] {
                    "Всё в рубрике Запчасти для транспорта",
                    "Автозапчасти и аксессуары",
                    "Шины, диски и колёса",
                    "Запчасти для спец / с.х. техники",
                    "Мотозапчасти и аксессуары",
                    "Прочие запчасти"});
                    break;
                case "Работа":
                    //"Всё в рубрике Работа"
                    SubrubricComboBox.Text = "Всё в рубрике Работа";
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Items.AddRange(new object[] {
                    "Всё в рубрике Работа",
                    "Розничная торговля / продажи / закупки ",
                    "Транспорт / логистика",
                    "Строительство",
                    "Телекоммуникации / связь",
                    "Бары / рестораны",
                    "Юриспруденция и бухгалтерия",
                    "Управление персоналом / HR",
                    "Охрана / безопасность",
                    "Домашний персонал",
                    "Красота / фитнес / спорт",
                    "Туризм / отдых / развлечения",
                    "Образование",
                    "Культура / искусство",
                    "Медицина / фармация",
                    "ИТ / телеком / компьютеры",
                    "Банки / финансы / страхование",
                    "Недвижимость",
                    "Маркетинг / реклама / дизайн",
                    "Производство / энергетика",
                    "Сельское хозяйство / агробизнес / лесное хозяйство",
                    "Cекретариат / АХО",
                    "Частичная занятость",
                    "Начало карьеры / Студенты",
                    "Сервис и быт"});
                    break;
                case "Животные":
                    //"Всё в рубрике Животные"
                    SubrubricComboBox.Text = "Всё в рубрике Животные";
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Items.AddRange(new object[] {
                    "Всё в рубрике Животные",
                    "Бесплатно (животные и вязка)",
                    "Собаки",
                    "Кошки",
                    "Аквариумистика",
                    "Птицы",
                    "Грызуны",
                    "Рептилии",
                    "Сельхоз животные",
                    "Другие животные",
                    "Зоотовары",
                    "Вязка",
                    "Бюро находок"});
                    break;
                case "Дом и сад":
                    //"Всё в рубрике Дом и сад"
                    SubrubricComboBox.Text = "Всё в рубрике Дом и сад";
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Items.AddRange(new object[] {
                    "Всё в рубрике Дом и сад",
                    "Канцтовары / расходные материалы",
                    "Мебель",
                    "Продукты питания / напитки",
                    "Сад / огород",
                    "Предметы интерьера",
                    "Строительство / ремонт",
                    "Инструменты",
                    "Комнатные растения",
                    "Посуда / кухонная утварь",
                    "Садовый инвентарь",
                    "Хозяйственный инвентарь / бытовая химия",
                    "Прочие товары для дома" });
                    break;
                case "Электроника":
                    //"Всё в рубрике Электроника"
                    SubrubricComboBox.Text = "Всё в рубрике Электроника";
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Items.AddRange(new object[] {
                    "Всё в рубрике Электроника",
                    "Телефоны и аксессуары",
                    "Компьютеры и комплектующие",
                    "Фото / видео",
                    "Тв / видеотехника",
                    "Аудиотехника",
                    "Игры и игровые приставки",
                    "Планшеты / эл. книги и аксессуары",
                    "Ноутбуки и аксессуары",
                    "Техника для дома",
                    "Техника для кухни",
                    "Климатическое оборудование",
                    "Индивидуальный уход",
                    "Аксессуары и комплектующие",
                    "Прочая электроника"});
                    break;
                case "Бизнес и услуги":
                    //"Всё в рубрике Бизнес и услуги"
                    SubrubricComboBox.Text = "Всё в рубрике Бизнес и услуги";
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Items.AddRange(new object[] {
                    "Всё в рубрике Бизнес и услуги",
                    "Строительство / ремонт / уборка",
                    "Финансовые услуги / партнерство",
                    "Перевозки / аренда транспорта",
                    "Реклама / полиграфия / маркетинг / интернет",
                    "Няни / сиделки",
                    "Сырьё / материалы",
                    "Красота / здоровье",
                    "Оборудование",
                    "Услуги для животных",
                    "Продажа бизнеса",
                    "Развлечения / Искусство / Фото / Видео",
                    "Туризм / иммиграция",
                    "Услуги переводчиков / набор текста",
                    "Авто / мото услуги",
                    "Ремонт и обслуживание техники",
                    "Сетевой маркетинг",
                    "Юридические услуги",
                    "Прокат товаров",
                    "Прочие услуги"});
                    break;
                case "Мода и стиль":
                    //"Всё в рубрике Мода и стиль"
                    SubrubricComboBox.Text = "Всё в рубрике Мода и стиль";
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Items.AddRange(new object[] {
                    "Всё в рубрике Мода и стиль",
                    "Одежда/обувь",
                    "Для свадьбы",
                    "Наручные часы",
                    "Аксессуары",
                    "Подарки",
                    "Красота / здоровье в моде",
                    "Мода разное"});
                    break;
                case "Хобби, отдых и спорт":
                    //"Всё в рубрике Хобби, отдых и спорт"
                    SubrubricComboBox.Text = "Всё в рубрике Хобби, отдых и спорт";
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Items.AddRange(new object[] {
                    "Всё в рубрике Хобби, отдых и спорт",
                    "Антиквариат / коллекции",
                    "Музыкальные инструменты",
                    "Спорт / отдых",
                    "Книги / журналы",
                    "CD / DVD / пластинки / кассеты",
                    "Билеты",
                    "Поиск попутчиков",
                    "Поиск групп",
                    "Другое"});
                    break;
                case "Отдам даром":
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Text = "";
                    break;
                case "Обмен":
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Text = "";
                    break;
                default:
                    SubrubricComboBox.Items.Clear();
                    SubrubricComboBox.Text = "";
                    break;
            }     
        }
    }
}
