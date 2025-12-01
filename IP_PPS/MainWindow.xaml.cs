using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;


using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel;
using System.Collections.ObjectModel;

namespace IP_PPS
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.data = new PlansData();
            dic = this.data.Plans;
            this.DataContext = this.data;
        }
        PlansData data;

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            dic.Clear();
            {
                string data = File.ReadAllText("data.csv").Replace(";;;;;;;;;;;;;\r\n", "").Replace("\r\n", ";");
                var cells = data.Split(';');



                for (int i = 0; i < cells.Length - 14; i += 14)
                {
                    var name = cells[i];
                    var oplata = cells[i + 1];
                    var pred = cells[i + 6];
                    var prtype = cells[i + 7];
                    var sem = cells[i + 8];
                    var prgroups = cells[i + 9];
                    var pr_osen = cells[i + 11];
                    var pr_vesna = cells[i + 12];


                    if (!dic.ContainsKey(name))
                        dic.Add(name, new Plan());

                    dic[name].Name = name;
                    //dic[name].Trudoustr = trud;
                    dic[name].Predmets.Add(new Predmet());
                    dic[name].Predmets.Last().Name = pred;
                    dic[name].Predmets.Last().Type = prtype;
                    dic[name].Predmets.Last().Sem = ParseInt(sem);
                    dic[name].Predmets.Last().Groups = prgroups;
                    dic[name].Predmets.Last().Osen = ParseDec(pr_osen);
                    dic[name].Predmets.Last().Vesna = ParseDec(pr_vesna);
                    dic[name].Predmets.Last().Oplata = oplata;
                }

                List<string> to_remove = new List<string>();
                foreach (var p in dic.Values)
                    if (p.Stavka == 0m)
                        to_remove.Add(p.Name);
                foreach (var prepod in to_remove)
                    dic.Remove(prepod);

                string prdata = File.ReadAllText("prepods.csv").Replace("\r\n", ";");
                cells = prdata.Split(';');

                for (int i = 0; i < cells.Length - 1; i += 12)
                {
                    var surname = cells[i];
                    var name = cells[i + 1];
                    var stepen = cells[i + 2];
                    var dolzh = cells[i + 3];
                    var zvanie = cells[i + 4];
                    var year = ParseInt(cells[i + 5]);
                    var trud = cells[i + 6];
                    var name_rp = cells[i + 11];



                    var prepod = dic.Where(p => p.Key.Contains(surname)).Select(kv => kv.Value).FirstOrDefault();
                    if (prepod == null)
                        prepod = dic.Where(p => p.Key == name).Select(kv => kv.Value).FirstOrDefault();

                    if (prepod != null)
                    {
                        prepod.Dolzhnost = dolzh;
                        prepod.NameRP = name_rp;
                        prepod.Stepen = stepen;
                        prepod.Zvanie = zvanie;
                        if (year != 0)
                            prepod.Stazh = int.Parse(tboxYear.Text) - year;
                        else
                            prepod.Stazh = -1;
                        prepod.Trudoustr = trud;
                        prepod.DolzhnostRP = GetDolzhnostRP(dolzh);
                    }
                }
            }
            
            this.data.OnPropertyChanged(nameof(this.data.Prepods));
            foreach (var p in this.data.Prepods)
            { 
                p.OnDistribute();
            }
        }


        

        

        Dictionary<string, Plan> dic;




        int ParseInt(string s)
        {
            int val = 0;
            int.TryParse(s, out val);
            return val;
        }
        public static decimal ParseDec(string s)
        {
            decimal val = 0;
            decimal.TryParse(s, out val);
            return val;
        }
        string GetDolzhnostRP(string dolzh)
        {
            dolzh = dolzh.Replace("ст.", "старшего");
            if (dolzh.Contains("доцент"))
                return "доцента";
            if (dolzh.Contains("преподаватель"))
                return dolzh.Replace('ь', 'я');
            if (dolzh.Contains("ассистент"))
                return "ассистента";
            if (dolzh.Contains("зав"))
                return "заведующего кафедрой";
            return dolzh + 'а';
        }


        
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.Foses.Add(new Plan.Rabota(data.SelectedPlan.UpdateFoses));
        }
        private void RButton_Click_1(object sender, RoutedEventArgs e)
        {
            if(data.SelectedPlan.SelectedFos != null)
                data.SelectedPlan.Foses.Remove(data.SelectedPlan.SelectedFos);
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.NauchOrg.Add(new Plan.Rabota(data.SelectedPlan.UpdateNauchOrg));
        }
        private void RButton_Click_2(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedNauchOrg != null)
                data.SelectedPlan.NauchOrg.Remove(data.SelectedPlan.SelectedNauchOrg);
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.NauchIssl.Add(new Plan.Rabota(data.SelectedPlan.UpdateNauchIssl));
        }
        private void RButton_Click_3(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedNauchIssl != null)
                data.SelectedPlan.NauchIssl.Remove(data.SelectedPlan.SelectedNauchIssl);
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.NauchMetod.Add(new Plan.Rabota(data.SelectedPlan.UpdateNauchMetod));
        }
        private void RButton_Click_4(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedNauchMetod != null)
                data.SelectedPlan.NauchMetod.Remove(data.SelectedPlan.SelectedNauchMetod);
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            if(comboBox.SelectedIndex > 0)
                comboBox.SelectedIndex--;
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            if (comboBox.SelectedIndex < comboBox.Items.Count - 1)
                comboBox.SelectedIndex++;
        }




        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.Ispob.Add(new Plan.Rabota(data.SelectedPlan.UpdateIspob));
        }
        private void RButton_Click_7(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedIspob != null)
                data.SelectedPlan.Ispob.Remove(data.SelectedPlan.SelectedIspob);
        }
        private void Button_Click_8(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.Kval.Add(new Plan.Rabota(data.SelectedPlan.UpdateKval));
        }
        private void RButton_Click_8(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedKval != null)
                data.SelectedPlan.Kval.Remove(data.SelectedPlan.SelectedKval);
        }



        private void Button_Click_10(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.UchMetodOrg.Add(new Plan.Rabota(data.SelectedPlan.UpdateUchMetodOrg));
        }
        private void RButton_Click_10(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedUchMetodOrg != null)
                data.SelectedPlan.UchMetodOrg.Remove(data.SelectedPlan.SelectedUchMetodOrg);
        }

        private void Button_Click_11(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.MetodOb.Add(new Plan.Rabota(data.SelectedPlan.UpdateMetodOb));
        }
        private void RButton_Click_11(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedMetodOb != null)
                data.SelectedPlan.MetodOb.Remove(data.SelectedPlan.SelectedMetodOb);
        }


        private void Button_Click_12(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.CRC.Add(new Plan.Rabota(data.SelectedPlan.UpdateCRC));
        }
        private void RButton_Click_12(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedCRC != null)
                data.SelectedPlan.CRC.Remove(data.SelectedPlan.SelectedCRC);
        }

        private string GetLet(int let)
        {
            if (let == 0)
                return "лет";
            else if (let >= 11 && let <= 14)
                return "лет";
            else if (let % 10 == 0)
                return "лет";
            else if (let % 10 == 1)
                return "год";
            else if (let % 10 == 2)
                return "года";
            else if (let % 10 == 3)
                return "года";
            else if (let % 10 == 4)
                return "года";

            return "лет";
        }

        private void Button_Click_9(object sender, RoutedEventArgs e)
        {
            var plan = this.data.SelectedPlan;

            if (plan.IsValid() == false)
                if(MyOkCancelForm.Show("Ошибка", "Часы не сходятся, придется редачить, все равно сгенерировать?", "Да", "Нееет, спасите") == MyOkCancelForm.Result.Cancel)
                    return;


            Word.Application app = new Microsoft.Office.Interop.Word.Application { Visible = true };
            Word.Document template = app.Documents.Open(
                System.AppDomain.CurrentDomain.BaseDirectory + @"\Document.docx",
                ReadOnly: false, Visible: true);
            template.Activate();


            


            template.SaveAs2(System.AppDomain.CurrentDomain.BaseDirectory + $@"\Планы\ИТИВС 2022 Индивидуальный план {plan.NameRP}.docx");



            



            Word.Document doc = app.Documents.Open(
                System.AppDomain.CurrentDomain.BaseDirectory + $@"\Планы\ИТИВС 2022 Индивидуальный план {plan.NameRP}.docx",
                ReadOnly: false, Visible: true);
            doc.Activate();
            doc.SelectAllEditableRanges();

            int tableOffset = 0;
            Dictionary<int, decimal> osenHours = new Dictionary<int, decimal>();
            Dictionary<int, decimal> vesnaHours = new Dictionary<int, decimal>();

            app.Replace("dolzhnost", plan.Dolzhnost);
            app.Replace("dolzhrp", plan.DolzhnostRP);
            app.Replace("fiorp", plan.NameRP);
            app.Replace("zvanie", plan.Zvanie);
            app.Replace("stepen", plan.Stepen);
            app.Replace("stazh", plan.Stazh + " " + GetLet(plan.Stazh));
            app.Replace("stavka", plan.Stavka);
            app.Replace("trudoustr", plan.Trudoustr);



            if (plan.GroupedPredmets.Count > 20)
                System.Windows.MessageBox.Show("НУЖНО БОЛЬШЕ ЧЕМ 20 ЗАГЛУШЕК В ШАБЛОНЕ ДЛЯ ТАБЛИЦЫ 1.2");
            for (int pr = 1, i = 0; pr <= 20; pr++, i++)
            {
                var gps = plan.GroupedPredmets;
                if (i < gps.Count)
                {
                    var gp = gps[i];
                    app.Replace($"{pr}pr_sem", gp.Sem);
                    app.Replace($"{pr}pr_nazv",
                        $"{gp.Name} {string.Join(", ", gp.Groups.Select(group => $"({group})").ToArray())}");
                    for (int k = 0; k <= 12; k++)
                    {
                        string str;
                        if (gp.Hours[k] == 0)
                            str = "-";
                        else
                            str = gp.Hours[k].ToString();
                        app.Replace($"{pr}pr_{k + 1}", str);
                    }
                }
                else
                    break;
                //else
                //{
                //    app.Replace($"{pr}pr_sem", "");
                //    app.Replace($"{pr}pr_nazv", "");
                //    for (int k = 0; k <= 12; k++)
                //    {
                //        app.Replace($"{pr}pr_{k + 1}", "");
                //    }
                //}
            }
            {
                var tables = doc.Tables;
                var tb = tables[3];
                var rowstodelete = 20 - plan.GroupedPredmets.Count;
                for (int i = 0; i < rowstodelete; i++)
                {
                    tb.Cell(plan.GroupedPredmets.Count + 3, 1).Select();
                    app.Selection.SelectRow();
                    app.Selection.Cells.Delete();
                }
            }

            {
                app.Replace("lek_osen", plan.GroupedPredmets
                .Where(p => p.Period == "осень")
                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("lab_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("sem_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("kurs_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("ind_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("prakt_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("pred_spec_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("vkr_spec_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("pred_bak_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("vkr_bak_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("mag_osen", plan.GroupedPredmets
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("mag_kons_osen", plan.GroupedPredmets
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("vkr_mag_osen", plan.GroupedPredmets
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("uch_vsego_osen", plan.GroupedPredmets
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
                osenHours.Add(1, plan.GroupedPredmets
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours.Sum())
                   .Sum());
            }

            {
                app.Replace("lek_vesna", plan.GroupedPredmets
                .Where(p => p.Period == "весна")
                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("lab_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("sem_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("kurs_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("ind_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("prakt_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("pred_spec_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("vkr_spec_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("pred_bak_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("vkr_bak_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("mag_vesna", plan.GroupedPredmets
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("mag_kons_vesna", plan.GroupedPredmets
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("vkr_mag_vesna", plan.GroupedPredmets
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("uch_vsego_vesna", plan.GroupedPredmets
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
                vesnaHours.Add(1, plan.GroupedPredmets
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours.Sum())
                   .Sum());
            }

            {
                app.Replace("lek_vsego", plan.GroupedPredmets

                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("lab_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("sem_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("kurs_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("ind_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("prakt_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("pred_spec_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("vkr_spec_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("pred_bak_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("vkr_bak_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("mag_vsego", plan.GroupedPredmets

                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("mag_kons_vsego", plan.GroupedPredmets

                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("vkr_mag_vsego", plan.GroupedPredmets

                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("uch_vsego", plan.GroupedPredmets

                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
            }

            if (plan.GroupedPredmetsDop.Count > 0)
            {

                if (plan.GroupedPredmetsDop.Count > 20)
                    System.Windows.MessageBox.Show("НУЖНО БОЛЬШЕ ЧЕМ 20 ЗАГЛУШЕК В ШАБЛОНЕ ДЛЯ ТАБЛИЦЫ 2.2");
                for (int pr = 1, i = 0; pr <= 20; pr++, i++)
                {
                    var gps = plan.GroupedPredmetsDop;
                    if (i < gps.Count)
                    {
                        var gp = gps[i];
                        app.Replace($"{pr}pd_sem", gp.Sem);
                        app.Replace($"{pr}pd_nazv",
                            $"{gp.Name} {string.Join(", ", gp.Groups.Select(group => $"({group})").ToArray())}");
                        for (int k = 0; k <= 12; k++)
                        {
                            string str;
                            if (gp.Hours[k] == 0)
                                str = "-";
                            else
                                str = gp.Hours[k].ToString();
                            app.Replace($"{pr}pd_{k + 1}", str);
                        }
                    }
                    else
                        break;
                    //else
                    //{
                    //    app.Replace($"{pr}pd_sem", "");
                    //    app.Replace($"{pr}pd_nazv", "");
                    //    for (int k = 0; k <= 12; k++)
                    //    {
                    //        app.Replace($"{pr}pd_{k + 1}", "");
                    //    }
                    //}
                }

                var tables = doc.Tables;
                var tb = tables[5];
                var rowstodelete = 20 - plan.GroupedPredmetsDop.Count;
                for (int i = 0; i < rowstodelete; i++)
                {
                    tb.Cell(plan.GroupedPredmetsDop.Count + 3, 1).Select();
                    app.Selection.SelectRow();
                    app.Selection.Cells.Delete();
                }

                //app.Replace("22nepredusmotr", "");
                app.Replace("21zagolovok", "\t\t\t\t\t2.1. Сводные данные");
                //2.2. Занятия по учебным дисциплинам
                app.Replace("22zagolovok", "2.2. Занятия по учебным дисциплинам");
            }
            else
            {
                app.Replace("21zagolovok", "Не предусмотрено.");
                app.Replace("22zagolovok", "");

                var tables = doc.Tables;
                var tb = tables[4];
                tb.Delete();
                tableOffset++;


                tables = doc.Tables;
                tb = tables[5 - tableOffset];
                tb.Delete();
                tableOffset++;
            }


            {
                app.Replace("dlek_osen", plan.GroupedPredmetsDop
                .Where(p => p.Period == "осень")
                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("dlab_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("dsem_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("dkurs_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("dind_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("dprakt_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("dpred_spec_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("dvkr_spec_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("dpred_bak_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("dvkr_bak_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("dmag_osen", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "осень")
                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("dmag_kons_osen", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("dvkr_mag_osen", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("duch_vsego_osen", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
                osenHours.Add(2, plan.GroupedPredmetsDop
                   .Where(p => p.Period == "осень")
                   .Select(p => p.Hours.Sum())
                   .Sum());
            }

            {
                app.Replace("dlek_vesna", plan.GroupedPredmetsDop
                .Where(p => p.Period == "весна")
                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("dlab_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("dsem_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("dkurs_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("dind_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("dprakt_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("dpred_spec_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("dvkr_spec_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("dpred_bak_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("dvkr_bak_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("dmag_vesna", plan.GroupedPredmetsDop
                    .Where(p => p.Period == "весна")
                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("dmag_kons_vesna", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("dvkr_mag_vesna", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("duch_vsego_vesna", plan.GroupedPredmetsDop
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
                vesnaHours.Add(2, plan.GroupedPredmetsDop
                   .Where(p => p.Period == "весна")
                   .Select(p => p.Hours.Sum())
                   .Sum());
            }

            {
                app.Replace("dlek_vsego", plan.GroupedPredmetsDop

                .Select(p => p.Hours[0])
                .Sum().ToHours());
                app.Replace("dlab_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[1])
                    .Sum().ToHours());
                app.Replace("dsem_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[2])
                    .Sum().ToHours());
                app.Replace("dkurs_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[3])
                    .Sum().ToHours());
                app.Replace("dind_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[4])
                    .Sum().ToHours());
                app.Replace("dprakt_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[5])
                    .Sum().ToHours());
                app.Replace("dpred_spec_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[6])
                    .Sum().ToHours());
                app.Replace("dvkr_spec_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[7])
                    .Sum().ToHours());
                app.Replace("dpred_bak_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[8])
                    .Sum().ToHours());
                app.Replace("dvkr_bak_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[9])
                    .Sum().ToHours());
                app.Replace("dmag_vsego", plan.GroupedPredmetsDop

                    .Select(p => p.Hours[10])
                    .Sum().ToHours());
                app.Replace("dmag_kons_vsego", plan.GroupedPredmetsDop

                   .Select(p => p.Hours[11])
                   .Sum().ToHours());
                app.Replace("dvkr_mag_vsego", plan.GroupedPredmetsDop

                   .Select(p => p.Hours[12])
                   .Sum().ToHours());

                app.Replace("duch_vsego", plan.GroupedPredmetsDop

                   .Select(p => p.Hours.Sum())
                   .Sum().ToHours());
            }


            if(plan.asppredmets.Count == 0)
            {
                var tables = doc.Tables;
                var tb = tables[6 - tableOffset];
                tableOffset++;
                tb.Delete();


                tables = doc.Tables;
                tb = tables[7 - tableOffset];
                tableOffset++;
                tb.Delete();

                //3.1. Сводные данные
                //3.2. Занятия по учебным дисциплинам
                app.Replace("31zagolovok", "Не предусмотрено.");
                app.Replace("32zagolovok", "");
                app.Replace("dop_asp_osen", "-");
                app.Replace("dop_asp_vesna", "-");
                app.Replace("dop_asp_vsego", "-");

                osenHours.Add(3, 0m);
                vesnaHours.Add(3, 0m);
            }
            else
            {
                app.Replace("31zagolovok", "\t\t\t\t\t3.1. Сводные данные");
                app.Replace("32zagolovok", "3.2. Занятия по учебным дисциплинам");
                for (int pr = 1, i = 0; pr <= 24; pr++, i++)
                {
                    var ps = plan.asppredmets.ToList();
                    if (i < ps.Count)
                    {
                        var p = ps[i];
                        app.Replace($"{pr}ap_sem", p.Sem);
                        app.Replace($"{pr}ap_nazv", p.Name + $" ({string.Join(", ", p.Groups)})");
                        for(int k = 1; k <=6; k++)
                        {
                            app.Replace($"{pr}ap_{k}", p.Hours[k - 1].ToHours());
                        }
                    }
                    else
                        break;
                }
                var tables = doc.Tables;
                var tb = tables[7 - tableOffset];
                var count = plan.asppredmets.Count;
                var rowstodelete = 24 - count;
                for (int i = 0; i < rowstodelete; i++)
                {
                    tb.Cell(count + 3, 1).Select();
                    app.Selection.SelectRow();
                    app.Selection.Cells.Delete();
                }

                {
                    app.Replace("alek_osen", plan.asppredmets
                        .Where(p => p.Period == "осень")
                        .Select(p => p.Hours[0]).Sum().ToHours());
                    app.Replace("asem_osen", plan.asppredmets
                        .Where(p => p.Period == "осень")
                        .Select(p => p.Hours[1]).Sum().ToHours());
                    app.Replace("aind_osen", plan.asppredmets
                        .Where(p => p.Period == "осень")
                        .Select(p => p.Hours[2]).Sum().ToHours());
                    app.Replace("aprakt_osen", plan.asppredmets
                        .Where(p => p.Period == "осень")
                        .Select(p => p.Hours[3]).Sum().ToHours());
                    app.Replace("anauch_osen", plan.asppredmets
                        .Where(p => p.Period == "осень")
                        .Select(p => p.Hours[4]).Sum().ToHours());
                    app.Replace("ankr_osen", plan.asppredmets
                        .Where(p => p.Period == "осень")
                        .Select(p => p.Hours[5]).Sum().ToHours());

                    osenHours.Add(3, plan.asppredmets
                        .Where(p => p.Period == "осень")
                        .Select(p => p.Hours.Sum()).Sum());
                    app.Replace("a_vsego_osen", osenHours[3]);
                }
                {
                    app.Replace("alek_vesna", plan.asppredmets
                        .Where(p => p.Period == "весна")
                        .Select(p => p.Hours[0]).Sum().ToHours());
                    app.Replace("asem_vesna", plan.asppredmets
                        .Where(p => p.Period == "весна")
                        .Select(p => p.Hours[1]).Sum().ToHours());
                    app.Replace("aind_vesna", plan.asppredmets
                        .Where(p => p.Period == "весна")
                        .Select(p => p.Hours[2]).Sum().ToHours());
                    app.Replace("aprakt_vesna", plan.asppredmets
                        .Where(p => p.Period == "весна")
                        .Select(p => p.Hours[3]).Sum().ToHours());
                    app.Replace("anauch_vesna", plan.asppredmets
                        .Where(p => p.Period == "весна")
                        .Select(p => p.Hours[4]).Sum().ToHours());
                    app.Replace("ankr_vesna", plan.asppredmets
                        .Where(p => p.Period == "весна")
                        .Select(p => p.Hours[5]).Sum().ToHours());

                    vesnaHours.Add(3, plan.asppredmets
                        .Where(p => p.Period == "весна")
                        .Select(p => p.Hours.Sum()).Sum());
                    app.Replace("a_vsego_vesna", vesnaHours[3]);
                }

                {
                    app.Replace("alek_vsego", plan.asppredmets
                        
                        .Select(p => p.Hours[0]).Sum().ToHours());
                    app.Replace("asem_vsego", plan.asppredmets
                        
                        .Select(p => p.Hours[1]).Sum().ToHours());
                    app.Replace("aind_vsego", plan.asppredmets
                        
                        .Select(p => p.Hours[2]).Sum().ToHours());
                    app.Replace("aprakt_vsego", plan.asppredmets
                        
                        .Select(p => p.Hours[3]).Sum().ToHours());
                    app.Replace("anauch_vsego", plan.asppredmets
                        
                        .Select(p => p.Hours[4]).Sum().ToHours());
                    app.Replace("ankr_vsego", plan.asppredmets
                        
                        .Select(p => p.Hours[5]).Sum().ToHours());


                    app.Replace("a_vsego", vesnaHours[3] + osenHours[3]);
                }

                app.Replace("dop_asp_osen", osenHours[3].ToHours());
                app.Replace("dop_asp_vesna", vesnaHours[3].ToHours());
                app.Replace("dop_asp_vsego", (osenHours[3] + vesnaHours[3]).ToHours());
            }

            //4 tablica
            for (int pr = 1, i = 0; pr <= 11; pr++, i++)
            {
                var mps = plan.MetodPredmets;
                if (i < mps.Count)
                {
                    var mp = mps[i];
                    app.Replace($"metod_pr{pr}_nazv", mp.Name);
                    app.Replace($"metod_pr{pr}_osen", mp.Osen.ToHours());
                    app.Replace($"metod_pr{pr}_vesna", mp.Vesna.ToHours());

                    app.Replace($"metod_pr{pr}_osen_per", "осень");
                    app.Replace($"metod_pr{pr}_vesna_per", "весна");
                }
                else
                {
                    app.Replace($"metod_pr{pr}_nazv", "");
                    app.Replace($"metod_pr{pr}_osen", "");
                    app.Replace($"metod_pr{pr}_vesna", "");

                    app.Replace($"metod_pr{pr}_osen_per", "");
                    app.Replace($"metod_pr{pr}_vesna_per", "");
                }
            }


            //uchmetodorg
            for (int pr = 1, i = 0; pr <= 5; pr++, i++)
            {
                var mps = plan.uchmetodorg.GroupBy(f => f.UchMetodOrg).ToList();
                if (i < mps.Count)
                {
                    var mp = mps[i];

                    var m = mp.Select(f => f).ToList();
                    var mosen = m.FirstOrDefault(f => f.Period == "осень");
                    var mvesna = m.FirstOrDefault(f => f.Period == "весна");

                    if (mosen.UchMetodOrg != null)
                    {
                        app.Replace($"uchmetodorg{pr}_nazv", mosen.UchMetodOrg);
                        app.Replace($"uchmetodorg{pr}_osen", mosen.Hours);
                        app.Replace($"uchmetodorg{pr}_osen_per", mosen.Period);
                    }
                    else
                    {
                        app.Replace($"uchmetodorg{pr}_osen", "-");
                        app.Replace($"uchmetodorg{pr}_osen_per", "осень");
                    }
                    if (mvesna.UchMetodOrg != null)
                    {
                        app.Replace($"uchmetodorg{pr}_nazv", mvesna.UchMetodOrg);
                        app.Replace($"uchmetodorg{pr}_vesna", mvesna.Hours);
                        app.Replace($"uchmetodorg{pr}_vesna_per", mvesna.Period);
                    }
                    else
                    {
                        app.Replace($"uchmetodorg{pr}_vesna", "-");
                        app.Replace($"uchmetodorg{pr}_vesna_per", "весна");
                    }

                }
                else
                {
                    break;
                }
            }
            //metodob
            for (int pr = 1, i = 0; pr <= 5; pr++, i++)
            {
                var mps = plan.metodob.GroupBy(f => f.MetodOb).ToList();
                if (i < mps.Count)
                {
                    var mp = mps[i];

                    var m = mp.Select(f => f).ToList();
                    var mosen = m.FirstOrDefault(f => f.Period == "осень");
                    var mvesna = m.FirstOrDefault(f => f.Period == "весна");

                    if (mosen.MetodOb != null)
                    {
                        app.Replace($"metodob{pr}_nazv", mosen.MetodOb);
                        app.Replace($"metodob{pr}_osen", mosen.Hours);
                        app.Replace($"metodob{pr}_osen_per", mosen.Period);
                    }
                    else
                    {
                        app.Replace($"metodob{pr}_osen", "-");
                        app.Replace($"metodob{pr}_osen_per", "осень");
                    }
                    if (mvesna.MetodOb != null)
                    {
                        app.Replace($"metodob{pr}_nazv", mvesna.MetodOb);
                        app.Replace($"metodob{pr}_vesna", mvesna.Hours);
                        app.Replace($"metodob{pr}_vesna_per", mvesna.Period);
                    }
                    else
                    {
                        app.Replace($"metodob{pr}_vesna", "-");
                        app.Replace($"metodob{pr}_vesna_per", "весна");
                    }

                }
                else
                {
                    break;
                }
            }

            //crc
            for (int pr = 1, i = 0; pr <= 5; pr++, i++)
            {
                var mps = plan.crc.GroupBy(f => f.Crc).ToList();
                if (i < mps.Count)
                {
                    var mp = mps[i];

                    var m = mp.Select(f => f).ToList();
                    var mosen = m.FirstOrDefault(f => f.Period == "осень");
                    var mvesna = m.FirstOrDefault(f => f.Period == "весна");

                    if (mosen.Crc != null)
                    {
                        app.Replace($"metodCRC{pr}_nazv", mosen.Crc);
                        app.Replace($"metodCRC{pr}_osen", mosen.Hours);
                        app.Replace($"metodCRC{pr}_osen_per", mosen.Period);
                    }
                    else
                    {
                        app.Replace($"metodCRC{pr}_osen", "-");
                        app.Replace($"metodCRC{pr}_osen_per", "осень");
                    }
                    if (mvesna.Crc != null)
                    {
                        app.Replace($"metodCRC{pr}_nazv", mvesna.Crc);
                        app.Replace($"metodCRC{pr}_vesna", mvesna.Hours);
                        app.Replace($"metodCRC{pr}_vesna_per", mvesna.Period);
                    }
                    else
                    {
                        app.Replace($"metodCRC{pr}_vesna", "-");
                        app.Replace($"metodCRC{pr}_vesna_per", "весна");
                    }

                }
                else
                {
                    break;
                }
            }

            //foses
            for (int pr = 1, i = 0; pr <= 10; pr++, i++)
            {
                var mps = plan.foses.GroupBy(f => f.Fos).ToList();
                if (i < mps.Count)
                {
                    var mp = mps[i];

                    var m = mp.Select(f => f).ToList();
                    var mosen = m.FirstOrDefault(f => f.Period == "осень");
                    var mvesna = m.FirstOrDefault(f => f.Period == "весна");

                    if (mosen.Fos != null)
                    {
                        app.Replace($"fosrpd{pr}_nazv", mosen.Fos);
                        app.Replace($"fosrpd{pr}_osen", mosen.Hours);
                        app.Replace($"fosrpd{pr}osenper", mosen.Period);
                    }
                    else
                    {
                        app.Replace($"fosrpd{pr}_osen", "-");
                        app.Replace($"fosrpd{pr}osenper", "осень");
                    }
                    if (mvesna.Fos != null)
                    {
                        app.Replace($"fosrpd{pr}_nazv", mvesna.Fos);
                        app.Replace($"fosrpd{pr}_vesna", mvesna.Hours);
                        app.Replace($"fosrpd{pr}vesnaper", mvesna.Period);
                    }
                    else
                    {
                        app.Replace($"fosrpd{pr}_vesna", "-");
                        app.Replace($"fosrpd{pr}vesnaper", "весна");
                    }

                }
                else
                {
                    break;
                }
            }
            
            osenHours.Add(4, plan.foses
                .Where(f => f.Period == "осень")
                .Select(f => f.Hours).Sum()
                +
                plan.MetodPredmets
                .Select(p => p.Osen).Sum()
                +
                plan.uchmetodorg
                .Where(f => f.Period == "осень")
                .Select(f => f.Hours).Sum()
                +
                plan.metodob
                .Where(f => f.Period == "осень")
                .Select(f => f.Hours).Sum()
                +
                plan.crc
                .Where(f => f.Period == "осень")
                .Select(f => f.Hours).Sum()
                );
            app.Replace("4vsego_osen", osenHours[4].ToHours());
            vesnaHours.Add(4, plan.foses
                .Where(f => f.Period == "весна")
                .Select(f => f.Hours).Sum()
                +
                plan.MetodPredmets
                .Select(p => p.Vesna).Sum()
                +
                plan.uchmetodorg
                .Where(f => f.Period == "весна")
                .Select(f => f.Hours).Sum()
                +
                plan.metodob
                .Where(f => f.Period == "весна")
                .Select(f => f.Hours).Sum()
                +
                plan.crc
                .Where(f => f.Period == "весна")
                .Select(f => f.Hours).Sum());
            app.Replace("4vsego_vesna", vesnaHours[4].ToHours());
            app.Replace("4vsego", (plan.foses
                .Select(f => f.Hours).Sum()
                +
                plan.MetodPredmets
                .Select(p => p.Osen + p.Vesna).Sum()
                +
                plan.uchmetodorg
                .Select(f => f.Hours).Sum()
                +
                plan.metodob
                .Select(f => f.Hours).Sum()
                +
                plan.crc
                .Select(f => f.Hours).Sum()).ToHours()
                );


            //delete empty foses rows
            {
                var tables = doc.Tables;
                var tb = tables[8 - tableOffset];
                var count = plan.foses.GroupBy(f => f.Fos).ToList().Count;
                var rowstodelete = 10 - count;

                for (int i = 0; i < rowstodelete; i++)
                {
                    tb.Cell(11 * 2 + 41 + count * 2, 1).Select();
                    app.Selection.SelectRow();
                    app.Selection.Cells.Delete();
                }
            }
            //delete empty CRC rows
            {
                var tables = doc.Tables;
                var tb = tables[8 - tableOffset];
                var count = plan.crc.GroupBy(f => f.Crc).ToList().Count;
                var rowstodelete = 5 - count;

                for (int i = 0; i < rowstodelete; i++)
                {
                    tb.Cell(11 * 2 + 29 + count * 2, 1).Select();
                    app.Selection.SelectRow();
                    app.Selection.Cells.Delete();
                }
            }
            //delete empty metod predmets rows
            {
                var tables = doc.Tables;
                var tb = tables[8 - tableOffset];
                var rowstodelete = 11 - plan.MetodPredmets.Count;
                if (rowstodelete == 11)
                    rowstodelete = 10;
                for (int i = 0; i < rowstodelete; i++)
                {
                    tb.Cell(plan.MetodPredmets.Count * 2 + 28, 1).Select();
                    app.Selection.SelectRow();
                    app.Selection.Cells.Delete();
                }
            }
            //delete empty metod ob rows
            {
                var tables = doc.Tables;
                var tb = tables[8 - tableOffset];
                var count = plan.metodob.GroupBy(f => f.MetodOb).ToList().Count;
                var rowstodelete = 5 - count;

                for (int i = 0; i < rowstodelete; i++)
                {
                    tb.Cell(15 + count * 2, 1).Select();
                    app.Selection.SelectRow();
                    app.Selection.Cells.Delete();
                }
            }
            //delete empty uchmetodorg rows
            {
                var tables = doc.Tables;
                var tb = tables[8 - tableOffset];
                var count = plan.uchmetodorg.GroupBy(f => f.UchMetodOrg).ToList().Count;
                var rowstodelete = 5 - count;

                for (int i = 0; i < rowstodelete; i++)
                {
                    tb.Cell(3 + count * 2, 1).Select();
                    app.Selection.SelectRow();
                    app.Selection.Cells.Delete();
                }
            }


            //5 tablica

            if (plan.nauchorg.Count > 0 ||
                plan.nauchissl.Count > 0 ||
                plan.nauchmetod.Count > 0)
            {
                //nauchorg
                for (int pr = 1, i = 0; pr <= 3; pr++, i++)
                {
                    var mps = plan.nauchorg.GroupBy(f => f.Nauchorg).ToList();
                    if (i < mps.Count)
                    {
                        var mp = mps[i];

                        var m = mp.Select(f => f).ToList();
                        var mosen = m.FirstOrDefault(f => f.Period == "осень");
                        var mvesna = m.FirstOrDefault(f => f.Period == "весна");

                        if (mosen.Nauchorg != null)
                        {
                            app.Replace($"nauchorg{pr}_nazv", mosen.Nauchorg);
                            app.Replace($"nauchorg{pr}_osen", mosen.Hours);
                            app.Replace($"nauchorg{pr}_osen_per", mosen.Period);
                        }
                        else
                        {
                            app.Replace($"nauchorg{pr}_osen", "-");
                            app.Replace($"nauchorg{pr}_osen_per", "осень");
                        }
                        if (mvesna.Nauchorg != null)
                        {
                            app.Replace($"nauchorg{pr}_nazv", mvesna.Nauchorg);
                            app.Replace($"nauchorg{pr}_vesna", mvesna.Hours);
                            app.Replace($"nauchorg{pr}_vesna_per", mvesna.Period);
                        }
                        else
                        {
                            app.Replace($"nauchorg{pr}_vesna", "-");
                            app.Replace($"nauchorg{pr}_vesna_per", "весна");
                        }

                    }
                    else
                    {
                        break;
                    }
                }

                //nauchissl
                for (int pr = 1, i = 0; pr <= 3; pr++, i++)
                {
                    var mps = plan.nauchissl.GroupBy(f => f.Nauchissl).ToList();
                    if (i < mps.Count)
                    {
                        var mp = mps[i];

                        var m = mp.Select(f => f).ToList();
                        var mosen = m.FirstOrDefault(f => f.Period == "осень");
                        var mvesna = m.FirstOrDefault(f => f.Period == "весна");

                        if (mosen.Nauchissl != null)
                        {
                            app.Replace($"nauchissl{pr}_nazv", mosen.Nauchissl);
                            app.Replace($"nauchissl{pr}_osen", mosen.Hours);
                            app.Replace($"nauchissl{pr}_osen_per", mosen.Period);
                        }
                        else
                        {
                            app.Replace($"nauchissl{pr}_osen", "-");
                            app.Replace($"nauchissl{pr}_osen_per", "осень");
                        }
                        if (mvesna.Nauchissl != null)
                        {
                            app.Replace($"nauchissl{pr}_nazv", mvesna.Nauchissl);
                            app.Replace($"nauchissl{pr}_vesna", mvesna.Hours);
                            app.Replace($"nauchissl{pr}_vesna_per", mvesna.Period);
                        }
                        else
                        {
                            app.Replace($"nauchissl{pr}_vesna", "-");
                            app.Replace($"nauchissl{pr}_vesna_per", "весна");
                        }

                    }
                    else
                    {
                        break;
                    }
                }

                //nauchmetod
                for (int pr = 1, i = 0; pr <= 3; pr++, i++)
                {
                    var mps = plan.nauchmetod.GroupBy(f => f.Nauchmetod).ToList();
                    if (i < mps.Count)
                    {
                        var mp = mps[i];

                        var m = mp.Select(f => f).ToList();
                        var mosen = m.FirstOrDefault(f => f.Period == "осень");
                        var mvesna = m.FirstOrDefault(f => f.Period == "весна");

                        if (mosen.Nauchmetod != null)
                        {
                            app.Replace($"nauchmetod{pr}_nazv", mosen.Nauchmetod);
                            app.Replace($"nauchmetod{pr}_osen", mosen.Hours);
                            app.Replace($"nauchmetod{pr}_osen_per", mosen.Period);
                        }
                        else
                        {
                            app.Replace($"nauchmetod{pr}_osen", "-");
                            app.Replace($"nauchmetod{pr}_osen_per", "осень");
                        }
                        if (mvesna.Nauchmetod != null)
                        {
                            app.Replace($"nauchmetod{pr}_nazv", mvesna.Nauchmetod);
                            app.Replace($"nauchmetod{pr}_vesna", mvesna.Hours);
                            app.Replace($"nauchmetod{pr}_vesna_per", mvesna.Period);
                        }
                        else
                        {
                            app.Replace($"nauchmetod{pr}_vesna", "-");
                            app.Replace($"nauchmetod{pr}_vesna_per", "весна");
                        }

                    }
                    else
                    {
                        break;
                    }
                }

                //delete empty nauchmetod rows
                {
                    var tables = doc.Tables;
                    var tb = tables[9 - tableOffset];
                    var count = plan.nauchmetod.GroupBy(f => f.Nauchmetod).ToList().Count;
                    var rowstodelete = 3 - count;

                    for (int i = 0; i < rowstodelete; i++)
                    {
                        tb.Cell(19 + count * 2, 1).Select();
                        app.Selection.SelectRow();
                        app.Selection.Cells.Delete();
                    }
                }
                //delete empty nauchissl rows
                {
                    var tables = doc.Tables;
                    var tb = tables[9 - tableOffset];
                    var count = plan.nauchissl.GroupBy(f => f.Nauchissl).ToList().Count;
                    var rowstodelete = 3 - count;

                    for (int i = 0; i < rowstodelete; i++)
                    {
                        tb.Cell(11 + count * 2, 1).Select();
                        app.Selection.SelectRow();
                        app.Selection.Cells.Delete();
                    }
                }
                //delete empty nauchorg rows
                {
                    var tables = doc.Tables;
                    var tb = tables[9 - tableOffset];
                    var count = plan.nauchorg.GroupBy(f => f.Nauchorg).ToList().Count;
                    var rowstodelete = 3 - count;

                    for (int i = 0; i < rowstodelete; i++)
                    {
                        tb.Cell(3 + count * 2, 1).Select();
                        app.Selection.SelectRow();
                        app.Selection.Cells.Delete();
                    }
                }

                app.Replace("5nepredusmotr", "");
            }
            else
            {
                var tables = doc.Tables;
                var tb = tables[9 - tableOffset];
                tb.Delete();
                tableOffset++;
                app.Replace("5nepredusmotr", "Не предусмотрено.");
            }

            osenHours.Add(5, plan.nauchorg
                .Where(n => n.Period == "осень")
                .Select(n => n.Hours).Sum()
                +
                plan.nauchissl
                .Where(n => n.Period == "осень")
                .Select(n => n.Hours).Sum()
                +
                plan.nauchmetod
                .Where(n => n.Period == "осень")
                .Select(n => n.Hours).Sum());
            app.Replace("5osen", osenHours[5].ToHours());
            vesnaHours.Add(5, plan.nauchorg
                .Where(n => n.Period == "весна")
                .Select(n => n.Hours).Sum()
                +
                plan.nauchissl
                .Where(n => n.Period == "весна")
                .Select(n => n.Hours).Sum()
                +
                plan.nauchmetod
                .Where(n => n.Period == "весна")
                .Select(n => n.Hours).Sum());
            app.Replace("5vesna", vesnaHours[5].ToHours());
            app.Replace("5vsego", (plan.nauchorg
                .Select(n => n.Hours).Sum()
                +
                plan.nauchissl
                .Select(n => n.Hours).Sum()
                +
                plan.nauchmetod
                .Select(n => n.Hours).Sum()).ToHours()
                );

            


            //6 tablica 
            //ispob
            for (int pr = 1, i = 0; pr <= 3; pr++, i++)
            {
                var mps = plan.ispob.GroupBy(f => f.Ispob).ToList();
                if (i < mps.Count)
                {
                    var mp = mps[i];

                    var m = mp.Select(f => f).ToList();
                    var mosen = m.FirstOrDefault(f => f.Period == "осень");
                    var mvesna = m.FirstOrDefault(f => f.Period == "весна");

                    if (mosen.Ispob != null)
                    {
                        app.Replace($"ispob{pr}_nazv", mosen.Ispob);
                        app.Replace($"ispob{pr}_osen", mosen.Hours);
                        app.Replace($"ispob{pr}_osen_per", mosen.Period);
                    }
                    else
                    {
                        app.Replace($"ispob{pr}_osen", "-");
                        app.Replace($"ispob{pr}_osen_per", "осень");
                    }
                    if (mvesna.Ispob != null)
                    {
                        app.Replace($"ispob{pr}_nazv", mvesna.Ispob);
                        app.Replace($"ispob{pr}_vesna", mvesna.Hours);
                        app.Replace($"ispob{pr}_vesna_per", mvesna.Period);
                    }
                    else
                    {
                        app.Replace($"ispob{pr}_vesna", "-");
                        app.Replace($"ispob{pr}_vesna_per", "весна");
                    }

                }
                else
                {
                    break;
                }
            }
            //delete empty ispob rows
            {
                var tables = doc.Tables;
                var tb = tables[10 - tableOffset];
                var count = plan.ispob.GroupBy(f => f.Ispob).ToList().Count;
                var rowstodelete = 3 - count;

                for (int i = 0; i < rowstodelete; i++)
                {
                    tb.Cell(3 + count * 2, 1).Select();
                    app.Selection.SelectRow();
                    app.Selection.Cells.Delete();
                }
            }
            if(plan.KafUch == true)
            {
                app.Replace("kafuch_nazv", "Участие в работе учебно-методической группы кафедры");
                app.Replace("kafuch", "10");
                app.Replace("kafuch_osen_per", "осень");
                app.Replace("kafuch_vesna_per", "весна");
            }
            else
            {
                app.Replace("kafuch_nazv", "");
                app.Replace("kafuch", "");
                app.Replace("kafuch_osen_per", "");
                app.Replace("kafuch_vesna_per", "");
            }

            osenHours.Add(6, plan.ispob
                .Where(o => o.Period == "осень")
                .Select(o => o.Hours).Sum()
                +
                (plan.KafUch ? 10m : 0m)
                +
                10m);
            app.Replace("6osen", osenHours[6]);
            vesnaHours.Add(6, plan.ispob
                .Where(o => o.Period == "весна")
                .Select(o => o.Hours).Sum()
                +
                (plan.KafUch ? 10m : 0m)
                +
                10m);
            app.Replace("6vesna", vesnaHours[6]);
            app.Replace("6vsego", plan.ispob
                .Select(o => o.Hours).Sum()
                +
                (plan.KafUch ? 20m : 0m)
                +
                20m
                );


            //7 tablica
            if(plan.kval.Count > 0)
            {
                for (int pr = 1, i = 0; pr <= 4; pr++, i++)
                {
                    var kps = plan.kval;
                    if (i < kps.Count)
                    {
                        var kv = kps[i];
                        app.Replace($"kval{pr}_nazv", kv.Kval);
                        app.Replace($"kval{pr}", kv.Hours.ToHours());
                        app.Replace($"kval{pr}_per", kv.Period);
                    }
                    else
                    {
                        break;
                        //app.Replace($"kval{pr}_nazv", "");
                        //app.Replace($"kval{pr}", "");
                        //app.Replace($"kval{pr}_per", "");
                    }
                }

                //delete empty kval rows
                {
                    var tables = doc.Tables;
                    var tb = tables[11 - tableOffset];
                    var count = plan.kval.Count;
                    var rowstodelete = 4 - count;

                    for (int i = 0; i < rowstodelete; i++)
                    {
                        tb.Cell(2 + count, 1).Select();
                        app.Selection.SelectRow();
                        app.Selection.Cells.Delete();
                    }
                }

                app.Replace("7nepredusmotr", "");
            }
            else
            {
                var tables = doc.Tables;
                var tb = tables[11 - tableOffset];
                tb.Delete();
                tableOffset++;
                app.Replace("7nepredusmotr", "Не предусмотрено.");
            }
            osenHours.Add(7, plan.kval
                .Where(k => k.Period == "осень")
                .Select(k => k.Hours).Sum());
            app.Replace("7osen", osenHours[7].ToHours());
            vesnaHours.Add(7, plan.kval
                .Where(k => k.Period == "весна")
                .Select(k => k.Hours).Sum());
            app.Replace("7vesna", vesnaHours[7].ToHours());
            app.Replace("7vsego", plan.kval
                .Select(k => k.Hours).Sum().ToHours());


            //end
            app.Replace("vsego_osen", osenHours.Values.Sum());
            app.Replace("vsego_vesna", vesnaHours.Values.Sum());
            app.Replace("vsego", osenHours.Values.Sum() + vesnaHours.Values.Sum());

            app.Replace("otchet_osen1", osenHours[1] > 0 ?
                $"Запланированный на семестр объём учебной работы по штатному расписанию (Раздел 1 плана)" +
                $" выполнен полностью - {osenHours[1]} час."
                :
                $"Работы по разделу 1 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_osen2", osenHours[2] > 0 ?
                $"Запланированный на семестр объём дополнительной учебной работы со студентами (Раздел 2 плана)" +
                $" выполнен полностью - {osenHours[2]} час."
                :
                $"Работы по разделу 2 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_osen3", osenHours[3] > 0 ?
                $"Запланированный на семестр объём дополнительной учебной работы с аспирантами (Раздел 3 плана)" +
                $" выполнен полностью - {osenHours[3]} час."
                :
                $"Работы по разделу 3 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_osen4", osenHours[4] > 0 ?
                $"Запланированный на семестр объём учебно-методической работы (Раздел 4 плана)" +
                $" выполнен полностью - {osenHours[4]} час."
                :
                $"Работы по разделу 4 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_osen5", osenHours[5] > 0 ?
                $"Запланированный на семестр объём научной работы (Раздел 5 плана)" +
                $" выполнен полностью - {osenHours[5]} час."
                :
                $"Работы по разделу 5 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_osen6", osenHours[6] > 0 ?
                $"Запланированный на семестр объём организационно-воспитательной работы (Раздел 6 плана)" +
                $" выполнен полностью - {osenHours[6]} час."
                :
                $"Работы по разделу 6 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_osen7", osenHours[7] > 0 ?
                $"Запланированный на семестр объём работы по повышению квалификации (Раздел 7 плана)" +
                $" выполнен полностью - {osenHours[7]} час."
                :
                $"Работы по разделу 7 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_osen8", 
                $"Запланированный на семестр объём работы выполнен полностью – {osenHours.Values.Sum()} час.");






            app.Replace("otchet_vesna1", vesnaHours[1] > 0 ?
                $"Запланированный на семестр объём учебной работы по штатному расписанию (Раздел 1 плана)" +
                $" выполнен полностью - {vesnaHours[1]} час."
                :
                $"Работы по разделу 1 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_vesna2", vesnaHours[2] > 0 ?
                $"Запланированный на семестр объём дополнительной учебной работы со студентами (Раздел 2 плана)" +
                $" выполнен полностью - {vesnaHours[2]} час."
                :
                $"Работы по разделу 2 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_vesna3", vesnaHours[3] > 0 ?
                $"Запланированный на семестр объём дополнительной учебной работы с аспирантами (Раздел 3 плана)" +
                $" выполнен полностью - {vesnaHours[3]} час."
                :
                $"Работы по разделу 3 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_vesna4", vesnaHours[4] > 0 ?
                $"Запланированный на семестр объём учебно-методической работы (Раздел 4 плана)" +
                $" выполнен полностью - {vesnaHours[4]} час."
                :
                $"Работы по разделу 4 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_vesna5", vesnaHours[5] > 0 ?
                $"Запланированный на семестр объём научной работы (Раздел 5 плана)" +
                $" выполнен полностью - {vesnaHours[5]} час."
                :
                $"Работы по разделу 5 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_vesna6", vesnaHours[6] > 0 ?
                $"Запланированный на семестр объём организационно-воспитательной работы (Раздел 6 плана)" +
                $" выполнен полностью - {vesnaHours[6]} час."
                :
                $"Работы по разделу 6 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_vesna7", vesnaHours[7] > 0 ?
                $"Запланированный на семестр объём работы по повышению квалификации (Раздел 7 плана)" +
                $" выполнен полностью - {vesnaHours[7]} час."
                :
                $"Работы по разделу 7 индивидуального плана в отчётном семестре не планировались.");
            app.Replace("otchet_vesna8",
                $"Запланированный на семестр объём работы выполнен полностью – {vesnaHours.Values.Sum()} час.");

            doc.Save();
            System.Windows.MessageBox.Show($"План {plan.NameRP} сгенерирован!");
        }

        private void Button_Click_13(object sender, RoutedEventArgs e)
        {
            data.SelectedPlan.AspPredmets.Add(new Plan.GroupedAspPredmetStr(
                data.SelectedPlan.UpdateAspPredmets
                ));
        }

        private void Button_Click_14(object sender, RoutedEventArgs e)
        {
            if (data.SelectedPlan.SelectedAspPredmet != null)
                data.SelectedPlan.AspPredmets.Remove(data.SelectedPlan.SelectedAspPredmet);
        }
    }


    class Base : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public static class DocExtensions
    {
        private static void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = true;
            object matchWholeWord = false;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = Word.WdReplace.wdReplaceAll;
            object wrap = Word.WdFindWrap.wdFindContinue;
            //execute find and replace
            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }



        public static void Replace(this Word.Application doc, string key, object value)
        {
            FindAndReplace(doc, $"%{key}%", value.ToString());
        }
        public static void Replace(this Word.Application doc, string key, string value)
        {
            FindAndReplace(doc, $"%{key}%", value);
        }


        public static void ReplaceStr(this Word.Application doc, string str, string value)
        {
            FindAndReplace(doc, str, value);
        }


        public static string ToHours(this decimal hours)
        {
            if (hours == 0)
                return "-";
            else
                return hours.ToString();
        }
        public static decimal Normalize(this decimal value)
        {
            return value / 1.000000000000000000000000000000000m;
        }
    }
}
