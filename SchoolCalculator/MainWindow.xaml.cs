using System;
using System.Windows;
using System.Windows.Controls;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace WpfApplication1
{
    public partial class MainWindow : Window
    {
        private string pathIn;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void СreateButton(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
            ex.Workbooks.Open(@Path.Text);
            Microsoft.Office.Interop.Excel.Range forYach1;
            Microsoft.Office.Interop.Excel.Range forYach2;
            Button buttons;
            for (int i = 0; i < Convert.ToInt32(SchoolNum.Text); i++)
            {
                forYach1 = ex.Cells[i + Convert.ToInt32(SchoolY.Text), Convert.ToInt32(School.Text)] as Microsoft.Office.Interop.Excel.Range;
                forYach2 = ex.Cells[i + Convert.ToInt32(SchoolY.Text), Convert.ToInt32(Thing.Text)] as Microsoft.Office.Interop.Excel.Range;
                buttons = new Button
                {
                    Background = System.Windows.Media.Brushes.Green,
                    Width = 100d,
                    Height = 50d,
                    Margin = new Thickness(1d),
                    Content = forYach1.Value2.ToString() + '\n' + " (" + forYach2.Value2.ToString() + ")",
                };
                buttons.Click += ClickAdd;
                buttons.MouseRightButtonDown += ClickDel;
                wrapPanel.Children.Add(buttons);
            }
            ex.Quit();
        }
        private void ClickAdd(object sender, RoutedEventArgs e)
        {
            if ((sender as Button).Background == System.Windows.Media.Brushes.Green)
            {
                L1.Text += (" + " + Convert.ToString((sender as Button).Content).Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)[0]);
                L2.Content = Convert.ToInt32(L2.Content) + Convert.ToInt16(Convert.ToString((sender as Button).Content).Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries)[1]);
                (sender as Button).Background = System.Windows.Media.Brushes.Red;
                if (Convert.ToInt32(L2.Content) < Convert.ToInt32(Child.Text))
                    L2.Foreground = System.Windows.Media.Brushes.Green;
                else
                    L2.Foreground = System.Windows.Media.Brushes.Red;
            }
        }
        private void ClickDel(object sender, RoutedEventArgs e)
        {
            string buf = L1.Text.Replace(" + " + Convert.ToString((sender as Button).Content).Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)[0], "");
            if (L1.Text != buf)
            {
                L1.Text = buf;
                L2.Content = Convert.ToInt32(L2.Content) - Convert.ToInt16(Convert.ToString((sender as Button).Content).Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries)[1]);
                (sender as Button).Background = System.Windows.Media.Brushes.Green;
            }
            if (Convert.ToInt32(L2.Content) < Convert.ToInt32(Child.Text))
                L2.Foreground = System.Windows.Media.Brushes.Green;
            else
                L2.Foreground = System.Windows.Media.Brushes.Red;
        }
        private void DeleteClick(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog fbd = new CommonOpenFileDialog();
            fbd.ShowDialog();
            fbd.Title = "Выберете файл со школыми";
            if (fbd.IsCollectionChangeAllowed())
            {
                pathIn = fbd.FileName;
                L1.Text = "";
                Path.Text = pathIn;
                L2.Content = 0;
                wrapPanel.Children.Clear();
            }
        }
        private void NextClick(object sender, RoutedEventArgs e)
        {
            L1.Text = "";
            L2.Content = 0;
        }
    }
}