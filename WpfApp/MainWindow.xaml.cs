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
using System.Net.Http;
using System.Text.Json;
using static Persona.Persona;

using Word = Microsoft.Office.Interop.Word;

namespace WpfApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Variables Globales
        object oMissing = System.Reflection.Missing.Value;
        object oEndOfDoc = "\\endofdoc";

        Word._Application oWord;
        Word._Document oDoc;

        string name, email, picture;

        private void CloseDoc_Click(object sender, RoutedEventArgs e)
        {
            //Se cierra documento word
            oWord.Quit();
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void CreateDoc_Click(object sender, RoutedEventArgs e)
        {
            //Se crea documento word
            oWord = new Word.Application();
            oWord.Visible = true;
        }

        private async void CreateContentDoc_Click(object sender, RoutedEventArgs e)
        {

            Persona.Persona personas = new Persona.Persona();
            
            var baseUrl = "https://dummyapi.io/data/api/user?limit=10";

            var jsonSerializerOptions = new JsonSerializerOptions() { PropertyNameCaseInsensitive = true };

            using (var httpClient = new HttpClient())
            {
                //Se hace la peticion al servicio

                using (var requestMessage = new HttpRequestMessage(HttpMethod.Get, baseUrl))
                {
                    requestMessage.Headers.Add("app-id", "610c71d338b20c4f58d9d7db");
                    var Respuesta = await httpClient.SendAsync(requestMessage);

                    var contenido = await Respuesta.Content.ReadAsStringAsync();

                    personas = JsonSerializer.Deserialize<Persona.Persona>(contenido, jsonSerializerOptions);
                }
                
                //Documento Word
                    oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing);

                    Word.Table oTable;
                    Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    oTable = oDoc.Tables.Add(wrdRng, 4, 3, ref oMissing, ref oMissing);
                    oTable.Range.ParagraphFormat.SpaceAfter = 6;
                    int r, c;

                    //Se crea la tabla en word y se insertan los datos de las personas en las celdas
                    int index = 0;                    

                    for(r = 1; r <= 4; r++)
                    {
                        for(c = 1; c <= 3; c++)
                        {
                            var rng = oWord.Selection.Range;

                            name = $"{personas.Data[index].FirstName.ToString()}" + " " + $"{personas.Data[index].LastName.ToString()}";
                            email = personas.Data[index].Email.ToString();
                            picture = personas.Data[index].Picture.ToString();

                            oTable.Cell(r, c).Range.InlineShapes.AddPictureBullet(picture);

                            oTable.Cell(r, c).Range.InsertParagraphAfter();
                            oTable.Cell(r, c).Range.InsertAfter(name);

                            while (rng.Find.Execute(name))
                            {
                                rng.Font.Bold = 1;
                            }
                                                     
                            oTable.Cell(r, c).Range.InsertParagraphAfter();
                            oTable.Cell(r, c).Range.InsertAfter(email);

                            while (rng.Find.Execute(email))
                            {
                                rng.Font.Italic = 1;

                            }

                            oTable.Cell(r, c).Width = 148;
                            oTable.Cell(r, c).Height = 80;                                                                                 

                            if (index == 9)
                            {
                                break;
                            }

                            index++;                               
                        }
                    }
                //Se crean bordes de tabla
                oTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                oTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                oTable.AllowAutoFit = true;
            }
        }
    }
}
