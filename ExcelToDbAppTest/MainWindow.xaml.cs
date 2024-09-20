using Ganss.Excel;
using Microsoft.EntityFrameworkCore;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using TextBox = System.Windows.Controls.TextBox;

namespace ExcelToDbAppTest
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		public class User
		{
			public int Id { get; set; }
			public int CardCode { get; set; }
			public string? LastName { get; set; }
			public string? FirstName { get; set; }
			public string? SurName { get; set; }
			[FormulaResult]
			public string? PhoneMobile { get; set; }
			public string? Email { get; set; }
			public string? GenderId { get; set; }
			[FormulaResult]
			public string? Birthday { get; set; }
			public string? City { get; set; }
			public int Pincode { get; set; }
			public int Bonus { get; set; }
			public int Turnover { get; set; }
		}
		public class ApplicationContext : DbContext
		{
			public DbSet<User> Users { get; set; } = null!;
			public ApplicationContext()
			{
			}
			protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
			{
				optionsBuilder.UseSqlServer(@"Data Source=(localdb)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\ExcelDatabase.mdf;Trusted_Connection=True;User Instance=false;Persist Security Info=True");
			}

		}
		public async Task GetExcelFile()
		{
			OpenFileDialog ofd = new OpenFileDialog();
			string excelFile = string.Empty;
			ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
			if(ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				excelFile = ofd.FileName;
			}
			else
			{
				System.Windows.MessageBox.Show("Выберите корректный файл");
			}
			ExcelMapper excel = new ExcelMapper();
			var users = (await excel.FetchAsync<User>(excelFile)).ToList();
			await using var db = new ApplicationContext();
			db.Database.ExecuteSqlRaw("DELETE FROM dbo.Users");
			db.Database.ExecuteSqlRaw("DBCC CHECKIDENT ('dbo.Users', RESEED, -1)");
			db.Users.AddRange(users);
			var affected = await db.SaveChangesAsync();
			System.Windows.MessageBox.Show(affected > 0 ? $"Сохранено {affected} записей" : "В файле нет данных");
			if(ReadDbCB.IsChecked == true)
			{
				ShowDbInDataGrid();
			}
		}
		private async void SaveDataToDbButton_Click(object sender, RoutedEventArgs e)
		{
			await GetExcelFile();
		}

		private void ReadDataFromDbButton_Click(object sender, RoutedEventArgs e)
		{
			ShowDbInDataGrid();
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				AppDomain.CurrentDomain.SetData("DataDirectory", System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database"));
				string path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database", "ExcelDatabase.mdf");
				string sourceFileName = System.IO.Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory())!.Parent!.Parent!.FullName, "Database", "ExcelDatabase.mdf");
				if(!File.Exists(path))
				{
					string dir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database");
					if(!Directory.Exists(dir))
					{
						Directory.CreateDirectory(dir);
					} 
					File.Copy(sourceFileName, System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database", "ExcelDatabase.mdf"));
				}
			}
			catch { System.Windows.MessageBox.Show("Стоковый файл базы данных не найден"); }
		}
		public void ShowDbInDataGrid()
		{
			try
			{
				using(ApplicationContext db = new ApplicationContext())
				{
					var users = db.Users.ToList();
					DG.ItemsSource = users;
				}

			}
			catch { System.Windows.MessageBox.Show("База данных не найдена"); }
		}

		private void DG_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
		{
			string header = e.Column.Header.ToString()!;
			int row = ((DataGrid)sender).ItemContainerGenerator.IndexFromContainer(e.Row);
			using(ApplicationContext db = new ApplicationContext())
			{
				var user = db.Users.Where(u => u.Id == row).FirstOrDefault();
				if(user == null)
				{
					db.Users.Add(new User());
				}
				else
				{
					try
					{
						switch (header)
						{
							case "Birthday": user.Birthday = (e.EditingElement as TextBox)!.Text; break;
							case "Bonus": user.Bonus = int.Parse((e.EditingElement as TextBox)!.Text); break;
							case "CardCode": user.CardCode = int.Parse((e.EditingElement as TextBox)!.Text); break;
							case "City": user.City = (e.EditingElement as TextBox)!.Text; break;
							case "Email": user.Email = (e.EditingElement as TextBox)!.Text; break;
							case "FirstName": user.FirstName = (e.EditingElement as TextBox)!.Text; break;
							case "GenderId": user.GenderId = (e.EditingElement as TextBox)!.Text; break;
							case "LastName": user.LastName = (e.EditingElement as TextBox)!.Text; break;
							case "PhoneMobile": user.PhoneMobile = (e.EditingElement as TextBox)!.Text; break;
							case "Pincode": user.Pincode = int.Parse((e.EditingElement as TextBox)!.Text); break;
							case "SurName": user.SurName = (e.EditingElement as TextBox)!.Text; break;
							case "Turnover": user.Turnover = int.Parse((e.EditingElement as TextBox)!.Text); break;
							default: break;
						}
					}
					catch(FormatException fe)
					{
						System.Windows.MessageBox.Show($"Введите корректный тип для {header} ");
						ShowDbInDataGrid();
					}
					catch (Exception ex)
					{
						
						System.Windows.MessageBox.Show(ex.ToString());
					}
				}
				db.SaveChanges();
			}
		}

	}
}