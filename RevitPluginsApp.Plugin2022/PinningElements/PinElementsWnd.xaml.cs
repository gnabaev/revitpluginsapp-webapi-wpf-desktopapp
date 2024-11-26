using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
using System.Windows;
using Grid = Autodesk.Revit.DB.Grid;

namespace RevitPluginsApp.Plugin.PinningElements
{
    /// <summary>
    /// Interaction logic for PinElementsWnd.xaml
    /// </summary>
    public partial class PinElementsWnd : Window
    {
        private Document doc;

        private bool gridsAreChecked = false;

        private bool levelsAreChecked = false;

        private bool rvtLinksAreChecked = false;

        public PinElementsWnd(Document doc)
        {
            InitializeComponent();
            this.doc = doc;
        }

        private void gridsCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            gridsAreChecked = true;
        }

        private void gridsCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            gridsAreChecked = false;
        }

        private void levelsCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            levelsAreChecked = true;
        }

        private void levelsCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            levelsAreChecked = false;
        }

        private void rvtLinksCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            rvtLinksAreChecked = true;
        }

        private void rvtLinksCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            rvtLinksAreChecked = false;
        }

        private void pinElements_Click(object sender, RoutedEventArgs e)
        {
            if (gridsAreChecked || levelsAreChecked || rvtLinksAreChecked)
            {
                if (gridsAreChecked)
                {
                    // Фильтрация и закрепление осей
                    var grids = new FilteredElementCollector(doc).OfClass(typeof(Grid)).Cast<Grid>().ToList();

                    if (grids.Count == 0)
                    {
                        TaskDialog.Show("Предупреждение", "В документе отсутствуют оси.");
                        return;
                    }

                    foreach (var grid in grids)
                    {
                        if (!grid.Pinned)
                        {
                            using (Transaction transaction = new Transaction(doc))
                            {
                                transaction.Start("Закрепление осей");

                                grid.Pinned = true;

                                transaction.Commit();
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                }

                if (levelsAreChecked)
                {
                    // Фильтрация и закрепление уровней
                    var levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();

                    if (levels.Count == 0)
                    {
                        TaskDialog.Show("Предупреждение", "В документе отсутствуют уровни.");
                        return;
                    }

                    foreach (var level in levels)
                    {
                        if (!level.Pinned)
                        {
                            using (Transaction transaction = new Transaction(doc))
                            {
                                transaction.Start("Закрепление уровней");

                                level.Pinned = true;

                                transaction.Commit();
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                }

                if (rvtLinksAreChecked)
                {
                    // Фильтрация и закрепление RVT-связей
                    var rvtLinks = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();

                    if (rvtLinks.Count == 0)
                    {
                        TaskDialog.Show("Предупреждение", "В документе отсутствуют RVT-связи.");
                        return;
                    }

                    foreach (var rvtLink in rvtLinks)
                    {
                        if (!rvtLink.Pinned)
                        {
                            using (Transaction transaction = new Transaction(doc))
                            {
                                transaction.Start("Закрепление осей");

                                rvtLink.Pinned = true;

                                transaction.Commit();
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                }

                TaskDialog.Show("Уведомление", "Элементы закреплены успешно.");
                Close();
            }
            else
            {
                TaskDialog.Show("Ошибка", "Не выбрана ни одна категория.");
            }

        }

        private void cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
