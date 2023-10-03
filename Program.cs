using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

class Program
{
    static async Task Main()
    {
       
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


        Console.Write("Tipo de Elección: ");
        string tipoEleccion = Console.ReadLine();

        Console.Write("Distrito Id: ");
        string distritoId = Console.ReadLine();

        Console.Write("Sección Provincial Id: ");
        string seccionProvincialId = Console.ReadLine();

        Console.Write("Sección Id: ");
        string seccionId = Console.ReadLine();

        using (var httpClient = new HttpClient())
        {

            List<string> nombresCargos = new List<string>
                {
                    "PRESIDENTE",
                    "PARLAMENTARIO MERCOSUR NACIONAL",
                    "SENADOR NACIONAL",
                    "DIPUTADO NACIONAL",
                    "PARLAMENTARIO MERCOSUR PROVINCIAL",
                    "GOBERNADOR",
                    "SENADOR PROVINCIAL",
                    "DIPUTADO PROVINCIAL",
                    "INTENDENTE",
                    "CONCEJAL"
                };



            int rowIndex = 2; 
            List<string> mesaIds = new List<string> { "1244" }; 

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Resultados");

                //foreach (var mesaId in mesaIds)
                // {

                var row = 1;

                for (int categoriaIndex = 0; categoriaIndex < nombresCargos.Count; categoriaIndex++)
                {
                    int categoriaId = categoriaIndex + 1;
                    string nombreCargo = nombresCargos[categoriaIndex];

                    string apiPrueba = $"https://resultados.mininterior.gob.ar/api/resultados/getResultados?anioEleccion=2019&tipoRecuento=1&tipoEleccion=2&distritoId=1&seccionProvincialId=0&seccionId=3&circuitoId=000039&mesaId=1244&categoriaId={categoriaId}";


                    HttpResponseMessage response = await httpClient.GetAsync(apiPrueba);

                    string jsonResult = await response.Content.ReadAsStringAsync();


                    JObject jsonObject = JObject.Parse(jsonResult);


                    worksheet.Cells[row, 1].Value = "Circuito";
                    worksheet.Cells[row, 2].Value = "Mesa";
                    worksheet.Cells[row + 1, 1].Value = "00039";
                    worksheet.Cells[row + 1, 2].Value = "1244";

                    // Obtener la cantidad de votantes de estadoRecuento
                    worksheet.Cells[row, 3].Value = "Cantidad de Votantes";
                    int cantidadVotantes = (int)jsonObject["estadoRecuento"]["cantidadVotantes"];
                    worksheet.Cells[row + 1, 3].Value = cantidadVotantes;

                    var valoresTotalizadosPositivos = jsonObject["valoresTotalizadosPositivos"];



                    var valoresTotalizadosPositivosSorted = valoresTotalizadosPositivos
                            .OrderBy(agrupacion => (string)agrupacion["nombreAgrupacion"])
                            .ToList();



                    int columnNumber = 4;

                    var valoresAgrupaciones = new Dictionary<string, int>();


                    foreach (var agrupacion in valoresTotalizadosPositivosSorted)
                    {
                        string nombreAgrupacion = (string)agrupacion["nombreAgrupacion"];
                        int votos = (int)agrupacion["votos"];

                        // Agregar el valor al diccionario
                        valoresAgrupaciones[nombreAgrupacion] = votos;

                        // Agregar el encabezado al archivo Excel
                        worksheet.Cells[row, columnNumber].Value = nombreAgrupacion;
                        worksheet.Cells[row + 1, columnNumber].Value = votos;
                        columnNumber++;
                    }


                    foreach (var agrupacionNombre in valoresAgrupaciones.Keys)
                    {
                        worksheet.Cells[row, columnNumber].Value = agrupacionNombre;

                        // Obtener el valor del diccionario o establecerlo en 0 si falta
                        if (valoresAgrupaciones.TryGetValue(agrupacionNombre, out int votos))
                        {
                            worksheet.Cells[row + 1, columnNumber].Value = votos;
                        }
                        else
                        {
                            worksheet.Cells[row + 1, columnNumber].Value = 0;
                        }

                        columnNumber++;
                    }



                    var valoresTotalizadosOtros = jsonObject["valoresTotalizadosOtros"];
                    worksheet.Cells[row, columnNumber].Value = "Votos Nulos";
                    worksheet.Cells[row + 1, columnNumber].Value = ((int?)valoresTotalizadosOtros["votosNulos"]) ?? 0;
                    columnNumber++;
                    worksheet.Cells[row, columnNumber].Value = "Votos Recurridos";
                    worksheet.Cells[row + 1, columnNumber].Value = ((int?)valoresTotalizadosOtros["votosRecurridosComandoImpugnados"]) ?? 0;
                    columnNumber++;
                    worksheet.Cells[row, columnNumber].Value = "Votos impugnados";
                    worksheet.Cells[row + 1, columnNumber].Value = ((int?)valoresTotalizadosOtros["votosRecurridosComandoImpugnadosPorcentaje"]) ?? 0;
                    columnNumber++;
                    worksheet.Cells[row, columnNumber].Value = "Votos en blanco";
                    worksheet.Cells[row + 1, columnNumber].Value = ((int?)valoresTotalizadosOtros["votosEnBlanco"]) ?? 0;

                }

                //}
                // Guardar el archivo Excel en disco
                FileInfo excelFile = new FileInfo("Resultados.xlsx");
                package.SaveAs(excelFile);
            

            }


        }

        Console.WriteLine("Archivo Excel generado con éxito: Resultados.xlsx");

    }

    
}

