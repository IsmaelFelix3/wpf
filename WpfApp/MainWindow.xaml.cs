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
        object oMissing = System.Reflection.Missing.Value;
        object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

        Word._Application oWord;
        Word._Document oDoc;

        string name,lastname, email, picture;

        private void CloseDoc_Click(object sender, RoutedEventArgs e)
        {
            oWord.Quit();
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void CreateDoc_Click(object sender, RoutedEventArgs e)
        {
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

                using (var requestMessage = new HttpRequestMessage(HttpMethod.Get, baseUrl))
                {
                    requestMessage.Headers.Add("app-id", "610c71d338b20c4f58d9d7db");
                    var Respuesta = await httpClient.SendAsync(requestMessage);

                    var contenido = await Respuesta.Content.ReadAsStringAsync();

                    personas = JsonSerializer.Deserialize<Persona.Persona>(contenido, jsonSerializerOptions);
                }

                
                    oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing);

                    //Insert a 4 x 3 table, fill it with data, and make the first row
                    //bold and italic.
                    Word.Table oTable;
                    Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    oTable = oDoc.Tables.Add(wrdRng, 4, 3, ref oMissing, ref oMissing);
                    oTable.Range.ParagraphFormat.SpaceAfter = 6;
                    int r, c;

                    int index = 0;
                    
                    for(r = 1; r <= 4; r++)
                    {
                        for(c = 1; c <= 3; c++)
                        {

                        Word.Cell cell = this.oWord.ActiveDocument.Tables[1].Cell(r, c);

                        cell.Range.Text = "Name";
                        cell.Range.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphRight;
                        cell.Range.Text = "Name2";
                        cell.Range.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphRight;

                        //oTable.Cell(r,c).Range.InlineShapes.AddPicture(personas.Data[index].Picture.ToString());
                        // Word.Paragraph oPara;
                        //oTable.Spacing = 10;
                        //oTable.Range.InsertParagraphAfter();
                        //oTable.Cell(r,c).Range.Text = $"{personas.Data[index].FirstName.ToString()}" +" "+
                        // $"{personas.Data[index].LastName.ToString()} {personas.Data[index].Email.ToString()}";

                        //object o_CollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                        //Word.Range imgrng = oDoc.Content;
                        //imgrng.Collapse(ref o_CollapseEnd);
                        //imgrng.InlineShapes.AddPicture(personas.Data[index].Picture.ToString(), oMissing, oMissing, imgrng);

                        //var pic = oDoc.Shapes.AddPicture(personas.Data[index].Picture.ToString());                

                        //pic.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapBehind;

                        //oTable.Cell(r,c).Range.InlineShapes.AddPicture(personas.Data[index].Picture.ToString());                   
                        //oTable.Cell(r,c).Range.Text = $" {personas.Data[index].FirstName.ToString()}" +" "+
                        //$"{personas.Data[index].LastName.ToString()} {personas.Data[index].Email.ToString()}";


                        oTable.Cell(r, c).Width = 100;
                        oTable.Cell(r, c).Height = 100;

                        if (index == 9)
                            {
                                break;
                            }
                            index++;

                      
                        }
                    }
                oTable.Rows[1].Range.Font.Bold = 1;
                oTable.Rows[1].Range.Font.Italic = 1;

                oTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                oTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                oTable.AllowAutoFit = true;
            }
        }
    }
}
