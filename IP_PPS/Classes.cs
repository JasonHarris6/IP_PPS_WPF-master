using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Collections.ObjectModel;

namespace IP_PPS
{
    class Predmet
    {
        public string Name;
        public string Groups;
        public string Type;
        public int Sem;
        public decimal Osen;
        public decimal Vesna;

        public string Oplata;

        public decimal Vsego => Osen + Vesna;
    }

    class GroupedPredmet
    {
        public string Name { get; set; }
        public HashSet<string> Groups { get; } = new HashSet<string>();
        public int Sem { get; set; }
        public string Period => Sem % 2 == 1 ? "осень" : "весна";
        public decimal[] Hours { get; } = new decimal[13];

        public decimal AuditorHours => Hours[0] + Hours[1] + Hours[2];
    }
    class Plan : Base
    {
        public string Name { get; set; } = string.Empty;

        public List<Predmet> Predmets { get; } = new List<Predmet>();

        public decimal Stavka => Hours / 900m;
        public decimal Hours => PredmetsShtat.Select(p => p.Vsego).Sum();
        public decimal HoursToCount => Stavka * 600m;

        public List<Predmet> PredmetsShtat => Predmets.Where(p => !p.Oplata.Contains("стимулир")).ToList();
        public List<Predmet> PredmetsDop => Predmets.Where(p => p.Oplata.Contains("стимулир")).ToList();


        List<GroupedPredmet> GroupPredmets(List<Predmet> predmets)
        {
            var gr = predmets.GroupBy(p => new { p.Name, p.Sem, p.Groups });

            List<GroupedPredmet> list = new List<GroupedPredmet>();

            foreach (var p in gr)
            {
                var g = new GroupedPredmet();
                g.Name = p.Key.Name;
                g.Sem = p.Key.Sem;
                foreach (var el in p)
                {
                    g.Groups.Add(el.Groups);
                    switch (el.Type.ToLower())
                    {
                        case "лекции":
                            g.Hours[0] = el.Vsego;
                            break;
                        case "лабораторные работы":
                            g.Hours[1] = el.Vsego;
                            break;
                        case "практические занятия":
                            g.Hours[2] = el.Vsego;
                            break;
                        case "курсовые работы/проекты":
                            g.Hours[3] = el.Vsego;
                            break;
                        case "индивидуальная работа":
                            g.Hours[4] = el.Vsego;
                            break;

                        default:
                            if (el.Name.ToLower() == "Учебная практика".ToLower())
                                g.Hours[5] = el.Vsego;
                            else if (el.Name.ToLower() == "Производственная практика (научно-исследовательская работа)".ToLower())
                                g.Hours[5] = el.Vsego;
                            else if (el.Name.ToLower() == "Производственная практика".ToLower())
                                g.Hours[5] = el.Vsego;

                            else if (el.Name.ToLower() == "Дипломное проектирование (бакалавры)".ToLower())
                                g.Hours[9] = el.Vsego;
                            else if (el.Name.ToLower() == "Преддипломная практика".ToLower())
                                g.Hours[8] = el.Vsego;

                            else if (el.Name.ToLower() == "Защита дипломного проекта (бакалавры)".ToLower())
                                g.Hours[9] = el.Vsego;

                            else if (el.Name.ToLower() == "Учебная практика (ознакомительная)".ToLower())
                                g.Hours[10] = el.Vsego;
                            else if (el.Name.ToLower() == "Учебная практика (педагогическая)".ToLower())
                                g.Hours[10] = el.Vsego;
                            else if (el.Name.ToLower() == "Производственная практика (НИР)".ToLower())
                                g.Hours[10] = el.Vsego;
                            else if (el.Name.ToLower() == "Учебная практика (педагогическая)".ToLower())
                                g.Hours[10] = el.Vsego;
                            else if (el.Name.ToLower() == "Производственная практика (проектно-технологическая)".ToLower())
                                g.Hours[10] = el.Vsego;
                            else if (el.Name.ToLower() == "Производственная практика (НИР)".ToLower())
                                g.Hours[10] = el.Vsego;
                            else if (el.Name.ToLower() == "Производственная практика (проектно-технологическая)".ToLower())
                                g.Hours[10] = el.Vsego;
                            else if (el.Name.ToLower() == "Подготовка магистерской ВКР".ToLower())
                                g.Hours[12] = el.Vsego;

                            else if (el.Name.ToLower() == "Производственная практика (преддипломная)".ToLower())
                                g.Hours[10] = el.Vsego;
                            else if (el.Name.ToLower() == "Защита магистерской ВКР".ToLower())
                                g.Hours[12] = el.Vsego;
                            break;
                    }
                }
                list.Add(g);
            }

            return list;
        }
        public List<GroupedPredmet> GroupedPredmets
        {
            get
            {
                return GroupPredmets(PredmetsShtat);
            }
        }
        public List<GroupedPredmet> GroupedPredmetsDop
        {
            get
            {
                return GroupPredmets(PredmetsDop);
            }
        }

        public class GroupedPredmetStr
        {
            public string Name { get; set; }
            public String Groups { get; set; }
            public int Sem { get; set; }
            public string Period => Sem % 2 == 1 ? "осень" : "весна";
            public string H1 { get; set; }
            public string H2 { get; set; }
            public string H3 { get; set; }
            public string H4 { get; set; }
            public string H5 { get; set; }
            public string H6 { get; set; }
            public string H7 { get; set; }
            public string H8 { get; set; }
            public string H9 { get; set; }
            public string H10 { get; set; }
            public string H11 { get; set; }
            public string H12 { get; set; }
            public string H13 { get; set; }
        }
        public List<GroupedPredmetStr> GroupedPredmetsStrings => GroupedPredmets.Select(g => new GroupedPredmetStr {
            Name = g.Name,
            Groups = string.Join(", ", g.Groups.ToArray()),
            Sem = g.Sem,
            H1 = g.Hours[0].ToHours(),
            H2 = g.Hours[1].ToHours(),
            H3 = g.Hours[2].ToHours(),
            H4 = g.Hours[3].ToHours(),
            H5 = g.Hours[4].ToHours(),
            H6 = g.Hours[5].ToHours(),
            H7 = g.Hours[6].ToHours(),
            H8 = g.Hours[7].ToHours(),
            H9 = g.Hours[8].ToHours(),
            H10 = g.Hours[9].ToHours(),
            H11 = g.Hours[10].ToHours(),
            H12 = g.Hours[11].ToHours(),
            H13 = g.Hours[12].ToHours()
        }).ToList();
        public List<GroupedPredmetStr> GroupedPredmetsDopStrings => GroupedPredmetsDop.Select(g => new GroupedPredmetStr
        {
            Name = g.Name,
            Groups = string.Join(", ", g.Groups.ToArray()),
            Sem = g.Sem,
            H1 = g.Hours[0].ToHours(),
            H2 = g.Hours[1].ToHours(),
            H3 = g.Hours[2].ToHours(),
            H4 = g.Hours[3].ToHours(),
            H5 = g.Hours[4].ToHours(),
            H6 = g.Hours[5].ToHours(),
            H7 = g.Hours[6].ToHours(),
            H8 = g.Hours[7].ToHours(),
            H9 = g.Hours[8].ToHours(),
            H10 = g.Hours[9].ToHours(),
            H11 = g.Hours[10].ToHours(),
            H12 = g.Hours[11].ToHours(),
            H13 = g.Hours[12].ToHours()
        }).ToList();

        public class MetodPredmet
        {
            public string Name { get; set; }
            public decimal Osen { get; set; }
            public decimal Vesna { get; set; }
        }
        public List<MetodPredmet> MetodPredmets
        {
            get
            {
                decimal leftHours = HoursToCount - 20m - (kafuch ? 20m : 0m);
                leftHours -= foses.Select(f => f.Hours).Sum();

                leftHours -= uchmetodorg.Select(f => f.Hours).Sum();
                leftHours -= metodob.Select(f => f.Hours).Sum();
                leftHours -= crc.Select(f => f.Hours).Sum();

                leftHours -= nauchorg.Select(f => f.Hours).Sum();
                leftHours -= nauchissl.Select(f => f.Hours).Sum();
                leftHours -= nauchmetod.Select(f => f.Hours).Sum();
                leftHours -= ispob.Select(f => f.Hours).Sum();
                leftHours -= kval.Select(f => f.Hours).Sum();

                

                var g = GroupedPredmets.Where(p => p.AuditorHours > 0).GroupBy(p => p.Name);
                List<MetodPredmet> mp = new List<MetodPredmet>();

                foreach (var el in g)
                {
                    MetodPredmet pr = new MetodPredmet();
                    pr.Name = el.Key;
                    pr.Osen = el.Where(p => p.Period == "осень").Select(p => p.AuditorHours).Sum();
                    pr.Vesna = el.Where(p => p.Period == "весна").Select(p => p.AuditorHours).Sum();
                    mp.Add(pr);
                }

                if (mp.Count == 0)
                    return mp;

                decimal coef;
                decimal sum = mp.Select(p => p.Osen + p.Vesna).Sum();



                if (leftHours < sum)
                    coef = leftHours / sum;
                else
                    coef = 1m;

                foreach(var p in mp)
                {
                    p.Osen *= coef;
                    p.Vesna *= coef;

                    p.Osen = decimal.Round(p.Osen, 1).Normalize();
                    p.Vesna = decimal.Round(p.Vesna, 1).Normalize();
                }


                if(leftHours < sum)
                if(mp.Select(p => p.Osen + p.Vesna).Sum() != leftHours)
                {
                    var dif = mp.Select(p => p.Osen + p.Vesna).Sum() - leftHours;
                    if (mp[0].Osen != 0)
                        mp[0].Osen -= dif;
                    else
                        mp[0].Vesna -= dif;
                }

                return mp;
            }
        }

        public decimal MetodPrHours => MetodPredmets.Select(p => p.Osen + p.Vesna).Sum().Normalize();
        public decimal NotDistributedHours
        {
            get
            {
                if (HoursToCount == 0m)
                    return 0m;
                decimal leftHours = HoursToCount - 20m - (kafuch ? 20m : 0m);
                leftHours -= foses.Select(f => f.Hours).Sum();

                leftHours -= uchmetodorg.Select(f => f.Hours).Sum();
                leftHours -= metodob.Select(f => f.Hours).Sum();
                leftHours -= crc.Select(f => f.Hours).Sum();

                leftHours -= nauchorg.Select(f => f.Hours).Sum();
                leftHours -= nauchissl.Select(f => f.Hours).Sum();
                leftHours -= nauchmetod.Select(f => f.Hours).Sum();
                leftHours -= ispob.Select(f => f.Hours).Sum();
                leftHours -= kval.Select(f => f.Hours).Sum();
                leftHours -= MetodPredmets.Select(p => p.Osen + p.Vesna).Sum();
                return leftHours.Normalize();
            }
        }

        public string Trudoustr { get; set; } = string.Empty;
        public string Dolzhnost { get; set; } = string.Empty;
        public string DolzhnostRP { get; set; } = string.Empty;
        public string NameRP { get; set; } = string.Empty;
        public string Zvanie { get; set; } = string.Empty;
        public string Stepen { get; set; } = string.Empty;
        public int Stazh { get; set; } = -1;

        public string StazhStr
        {
            get => Stazh.ToString(); set
            {
                int val;
                if (int.TryParse(value, out val))
                    Stazh = val;

                OnPropertyChanged(nameof(Stazh));
                OnPropertyChanged(nameof(StazhStr));
            }
        }

        public override string ToString()
        {
            return Name;
        }
        //----- not parsed


        public void OnDistribute()
        {
            OnPropertyChanged(nameof(HoursEntered));
            OnPropertyChanged(nameof(MetodPredmets));
            OnPropertyChanged(nameof(NotDistributedHours));
            OnPropertyChanged(nameof(MetodPrHours));
            OnPropertyChanged(nameof(HoursMetodPlusHoursEntered));
        }

        public Plan()
        {
            Foses.CollectionChanged += (sender, e) =>
            {
                UpdateFoses();
            };
            UchMetodOrg.CollectionChanged += (sender, e) =>
            {
                UpdateUchMetodOrg();
            };
            MetodOb.CollectionChanged += (sender, e) =>
            {
                UpdateMetodOb();
            };
            CRC.CollectionChanged += (sender, e) =>
            {
                UpdateCRC();
            };
            NauchOrg.CollectionChanged += (sender, e) =>
            {
                UpdateNauchOrg();
            };
            NauchIssl.CollectionChanged += (sender, e) =>
            {
                UpdateNauchIssl();
            };
            NauchMetod.CollectionChanged += (sender, e) =>
            {
                UpdateNauchMetod();
            };
            Ispob.CollectionChanged += (sender, e) =>
            {
                UpdateIspob();
            };
            Kval.CollectionChanged += (sender, e) =>
            {
                UpdateKval();
            };
            AspPredmets.CollectionChanged += (sender, e) =>
            {
                UpdateAspPredmets();
            };
        }

        public class Rabota : Base
        {
            Action update;
            public Rabota(Action update)
            {
                this.update = update;
            }
            string name;
            public string Name { get => name; set { name = value; update?.Invoke(); } }
            string hours;
            public string Hours { get => hours; set { hours = value; update?.Invoke(); } }
            string period;
            public string Period { get => period; set { period = value; update?.Invoke(); } }
        }

        #region foses

        public List<(string Fos, decimal Hours, string Period)> foses = new List<(string Fos, decimal Hours, string Period)>();


        bool CheckPeriod(string period)
        {
            if (period.ToLower() == "осень" ||
                period.ToLower() == "весна")
                return true;
            return false;
        }
        public void UpdateFoses()
        {
            foses.Clear();
            foreach (var fos in Foses)
            {
                if (fos.Period != null && CheckPeriod(fos.Period))
                    foses.Add((fos.Name, Convert.ToDecimal(fos.Hours), fos.Period));
            }
            OnDistribute();
        }
        public ObservableCollection<Rabota> Foses
        {
            get;
        } = new ObservableCollection<Rabota>();
        Rabota selectedFos;
        public Rabota SelectedFos
        {
            get => selectedFos;
            set
            {
                selectedFos = value;
                OnPropertyChanged(nameof(SelectedFos));
            }
        }
        #endregion


        #region uchmetodorg

        public List<(string UchMetodOrg, decimal Hours, string Period)> uchmetodorg = new List<(string UchMetodOrg, decimal Hours, string Period)>();


        public void UpdateUchMetodOrg()
        {
            uchmetodorg.Clear();
            foreach (var r in UchMetodOrg)
            {
                if (r.Period != null && CheckPeriod(r.Period))
                    uchmetodorg.Add((r.Name, Convert.ToDecimal(r.Hours), r.Period));
            }
            OnDistribute();
        }
        public ObservableCollection<Rabota> UchMetodOrg
        {
            get;
        } = new ObservableCollection<Rabota>();
        Rabota selectedUchMetodOrg;
        public Rabota SelectedUchMetodOrg
        {
            get => selectedUchMetodOrg;
            set
            {
                selectedUchMetodOrg = value;
                OnPropertyChanged(nameof(SelectedUchMetodOrg));
            }
        }
        #endregion


        #region metodob

        public List<(string MetodOb, decimal Hours, string Period)> metodob = new List<(string MetodOb, decimal Hours, string Period)>();


        public void UpdateMetodOb()
        {
            metodob.Clear();
            foreach (var r in MetodOb)
            {
                if (r.Period != null && CheckPeriod(r.Period))
                    metodob.Add((r.Name, Convert.ToDecimal(r.Hours), r.Period));
            }
            OnDistribute();
        }
        public ObservableCollection<Rabota> MetodOb
        {
            get;
        } = new ObservableCollection<Rabota>();
        Rabota selectedMetodOb;
        public Rabota SelectedMetodOb
        {
            get => selectedMetodOb;
            set
            {
                selectedMetodOb = value;
                OnPropertyChanged(nameof(SelectedMetodOb));
            }
        }
        #endregion


        #region CRC

        public List<(string Crc, decimal Hours, string Period)> crc = new List<(string Crc, decimal Hours, string Period)>();


        public void UpdateCRC()
        {
            crc.Clear();
            foreach (var r in CRC)
            {
                if (r.Period != null && CheckPeriod(r.Period))
                    crc.Add((r.Name, Convert.ToDecimal(r.Hours), r.Period));
            }
            OnDistribute();
        }
        public ObservableCollection<Rabota> CRC
        {
            get;
        } = new ObservableCollection<Rabota>();
        Rabota selectedCRC;
        public Rabota SelectedCRC
        {
            get => selectedCRC;
            set
            {
                selectedCRC = value;
                OnPropertyChanged(nameof(SelectedCRC));
            }
        }
        #endregion


        #region nauchorg

        public List<(string Nauchorg, decimal Hours, string Period)> nauchorg = new List<(string Nauchorg, decimal Hours, string Period)>();

        public void UpdateNauchOrg()
        {
            nauchorg.Clear();
            foreach (var r in NauchOrg)
            {
                if (r.Period != null && CheckPeriod(r.Period))
                    nauchorg.Add((r.Name, Convert.ToDecimal(r.Hours), r.Period));
            }
            OnDistribute();
        }
        public ObservableCollection<Rabota> NauchOrg
        {
            get;
        } = new ObservableCollection<Rabota>();
        Rabota selectedNauchOrg;
        public Rabota SelectedNauchOrg
        {
            get => selectedNauchOrg;
            set
            {
                selectedNauchOrg = value;
                OnPropertyChanged(nameof(SelectedNauchOrg));
            }
        }
        #endregion


        #region nauchissl

        public List<(string Nauchissl, decimal Hours, string Period)> nauchissl = new List<(string Nauchorg, decimal Hours, string Period)>();

        public void UpdateNauchIssl()
        {
            nauchissl.Clear();
            foreach (var r in NauchIssl)
            {
                if (r.Period != null && CheckPeriod(r.Period))
                    nauchissl.Add((r.Name, Convert.ToDecimal(r.Hours), r.Period));
            }
            OnDistribute();
        }
        public ObservableCollection<Rabota> NauchIssl
        {
            get;
        } = new ObservableCollection<Rabota>();
        Rabota selectedNauchIssl;
        public Rabota SelectedNauchIssl
        {
            get => selectedNauchIssl;
            set
            {
                selectedNauchIssl = value;
                OnPropertyChanged(nameof(SelectedNauchIssl));
            }
        }
        #endregion


        #region nauchmetod

        public List<(string Nauchmetod, decimal Hours, string Period)> nauchmetod = new List<(string Nauchorg, decimal Hours, string Period)>();

        public void UpdateNauchMetod()
        {
            nauchmetod.Clear();
            foreach (var r in NauchMetod)
            {
                if (r.Period != null && CheckPeriod(r.Period))
                    nauchmetod.Add((r.Name, Convert.ToDecimal(r.Hours), r.Period));
            }
            OnDistribute();
        }
        public ObservableCollection<Rabota> NauchMetod
        {
            get;
        } = new ObservableCollection<Rabota>();
        Rabota selectedNauchMetod;
        public Rabota SelectedNauchMetod
        {
            get => selectedNauchMetod;
            set
            {
                selectedNauchMetod = value;
                OnPropertyChanged(nameof(SelectedNauchMetod));
            }
        }
        #endregion



        #region ispob

        public List<(string Ispob, decimal Hours, string Period)> ispob = new List<(string Nauchorg, decimal Hours, string Period)>();

        public void UpdateIspob()
        {
            ispob.Clear();
            foreach (var r in Ispob)
            {
                if (r.Period != null && CheckPeriod(r.Period))
                    ispob.Add((r.Name, Convert.ToDecimal(r.Hours), r.Period));
            }
            OnDistribute();
        }
        public ObservableCollection<Rabota> Ispob
        {
            get;
        } = new ObservableCollection<Rabota>();
        Rabota selectedIspob;
        public Rabota SelectedIspob
        {
            get => selectedIspob;
            set
            {
                selectedIspob = value;
                OnPropertyChanged(nameof(SelectedNauchMetod));
            }
        }
        #endregion



        #region kval

        public List<(string Kval, decimal Hours, string Period)> kval = new List<(string Kval, decimal Hours, string Period)>();

        public void UpdateKval()
        {
            kval.Clear();
            foreach (var r in Kval)
            {
                if (r.Period != null && CheckPeriod(r.Period))
                    kval.Add((r.Name, Convert.ToDecimal(r.Hours), r.Period));
            }
            OnDistribute();
        }
        public ObservableCollection<Rabota> Kval
        {
            get;
        } = new ObservableCollection<Rabota>();
        Rabota selectedKval;
        public Rabota SelectedKval
        {
            get => selectedKval;
            set
            {
                selectedKval = value;
                OnPropertyChanged(nameof(SelectedKval));
            }
        }
        #endregion


        public List<GroupedPredmet> asppredmets = new List<GroupedPredmet>();
        public class GroupedAspPredmetStr
        {
            Action update;
            public GroupedAspPredmetStr(Action update)
            {
                this.update = update;
            }
            string name;
            public string Name { get => name; set { name = value; update?.Invoke(); } }
            string groups;
            public string Groups { get => groups; set { groups = value; update?.Invoke(); } }
            int sem;
            public int Sem { get => sem; set { sem = value; update?.Invoke(); } }
            string h1, h2, h3, h4, h5, h6;
            public string H1 { get => h1; set { h1 = value; update?.Invoke(); } }
            public string H2 { get => h2; set { h2 = value; update?.Invoke(); } }
            public string H3 { get => h3; set { h3 = value; update?.Invoke(); } }
            public string H4 { get => h4; set { h4 = value; update?.Invoke(); } }
            public string H5 { get => h5; set { h5 = value; update?.Invoke(); } }
            public string H6 { get => h6; set { h6 = value; update?.Invoke(); } }
        }


        public ObservableCollection<GroupedAspPredmetStr> AspPredmets
        {
            get;
        } = new ObservableCollection<GroupedAspPredmetStr>();
        public void UpdateAspPredmets()
        {
            asppredmets.Clear();
            foreach(var pr in AspPredmets)
            {
                GroupedPredmet g = new GroupedPredmet();
                g.Name = pr.Name;
                g.Groups.Add(pr.Groups);
                g.Sem = pr.Sem;
                g.Hours[0] = MainWindow.ParseDec(pr.H1);
                g.Hours[1] = MainWindow.ParseDec(pr.H2);
                g.Hours[2] = MainWindow.ParseDec(pr.H3);
                g.Hours[3] = MainWindow.ParseDec(pr.H4);
                g.Hours[4] = MainWindow.ParseDec(pr.H5);
                g.Hours[5] = MainWindow.ParseDec(pr.H6);
                asppredmets.Add(g);
            }
            OnDistribute();
        }

        public GroupedAspPredmetStr selectedAspPredmet;
        public GroupedAspPredmetStr SelectedAspPredmet { get => selectedAspPredmet; set
            {
                selectedAspPredmet = value;
                OnPropertyChanged(nameof(SelectedAspPredmet));
            } 
        }


        bool kafuch = false;
        public bool KafUch { get => kafuch;
        set
            {
                kafuch = value;
                OnPropertyChanged(nameof(KafUch));
                OnDistribute();
            }
        }


        public decimal HoursMetodPlusHoursEntered => (HoursEntered + MetodPrHours).Normalize();
        public bool IsValid()
        {
            if (HoursEntered + MetodPrHours != HoursToCount)
                return false;
            if (MetodPrHours < 0)
                return false;


            return true;
        }

        public decimal HoursEntered =>
            foses.Select(r => r.Hours).Sum() +

            uchmetodorg.Select(r => r.Hours).Sum() +
            metodob.Select(r => r.Hours).Sum() +
            crc.Select(r => r.Hours).Sum() +

            nauchorg.Select(r => r.Hours).Sum() +
            nauchissl.Select(r => r.Hours).Sum() +
            nauchmetod.Select(r => r.Hours).Sum() +
            ispob.Select(r => r.Hours).Sum() +
            kval.Select(r => r.Hours).Sum() + 20m + (kafuch ? 20m : 0m);
    }



    class PlansData : Base
    {
        public Dictionary<string, Plan> Plans { get; } = new Dictionary<string, Plan>();

        public IEnumerable<Plan> Prepods => Plans.Values.ToList();


        Plan selectedPlan;
        public Plan SelectedPlan
        {
            get => selectedPlan; set
            {
                selectedPlan = value;
                OnPropertyChanged(nameof(SelectedPlan));
            }
        }
    }
}
