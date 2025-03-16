using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using WorkWithWord.HelperClasses;
using WorkWithWord.ModelClasses;


namespace WorkWithWord
{
    public partial class MainWindow : Window
    {
        private ModelEF model;

        private List<Users> users;

        private List<Auto> autos;

        public MainWindow()
        {
            InitializeComponent();

            model = new ModelEF();

            users = new List<Users>();
            autos = new List<Auto>();
        }

        private void ComboLoadData()
        {
            comboBoxUsers.Items.Clear();

            users = model.Users.ToList();

            foreach (var item in users)
                comboBoxUsers.Items.Add($"{item.FullName} {item.PSeria} {item.PNumber}");

            comboBoxUsers.SelectedIndex = 0;

            autos = users[comboBoxUsers.SelectedIndex].Auto.ToList();
            comboBoxAutos.Items.Clear();

            foreach (var item in autos)
                comboBoxAutos.Items.Add($"{item.Model} {item.YearOfRelease.Value.Year} {item.VIN}");

            comboBoxAutos.SelectedIndex = 0;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ComboLoadData();
        }

        private void comboBoxUsers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            autos = users[comboBoxUsers.SelectedIndex].Auto.ToList();
            comboBoxAutos.Items.Clear();

            foreach (var item in autos)
                comboBoxAutos.Items.Add($"{item.Model} {item.YearOfRelease.Value.Year} {item.VIN}");

            comboBoxAutos.SelectedIndex = 0;
        }

        private void SaveDocument_Click(object sender, RoutedEventArgs e)
        {
             FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Выберите место сохранения";

            //folderBrowserDialog.Description = description;

            if (System.Windows.Forms.DialogResult.OK == fbd.ShowDialog())
            {
                //var directoryPath = folderBrowserDialog.SelectedPath;
                Users activeuser = users[comboBoxUsers.SelectedIndex];

                Auto activeAuto = activeuser.Auto.ToList()[comboBoxAutos.SelectedIndex];

                CreateDocument(
                    $@"{fbd.SelectedPath}\Купля-Продажа-Автомобиля-{activeuser.FullName}.docx",
                    activeuser,
                    activeAuto);

                System.Windows.MessageBox.Show("Документ успешно сохранен!");
            }
        }

        private void CreateDocument(string directoryPath, Users user, Auto auto)
        {
            var today = DateTime.Now.ToShortDateString();

            WordHelper word = new WordHelper("ContractSale.docx");

            var items = new Dictionary<string, string>
            {
                {"<Today>", today },

                {"<FullName>", user.FullName },
                {"<DateOfBirth>", user.DateOfBirth.Value.ToShortDateString() },
                {"<Adress>", user.Adress },
                {"<PSeria>", user.PSeria.ToString() },
                {"<PNumber>", user.PNumber.ToString() },
                {"<PVidan>", user.PVidan },

                {"<ModelV>", auto.Model },
                {"<CategoryV>", auto.Category },
                {"<TypeV>", auto.TypeV },
                {"<VIN>", auto.VIN },
                {"<RegistrationMark>", auto.RegistrationMark },
                {"<YearV>", auto.YearOfRelease.Value.Year.ToString() },
                {"<EngineV>", auto.EngineNumber },
                {"<ChassisV>", auto.Chassis },
                {"<BodyworkV>", auto.Bodywork },
                {"<ColorV>", auto.Color },
                {"<SeriaPV>", auto.SeriaPasport },
                {"<NumberPV>", auto.NumbePasport },
                {"<VidanPV>", auto.VidanPasport }
            };

            word.Process(items, directoryPath);
        }
    }
}