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
        // System.AppDomain.CurrentDomain.SetData("EPPlusLicenseContext", LicenseContext.NonCommercial);
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


        Console.Write("Tipo de Elección: ");
        string tipoEleccion = Console.ReadLine();

        Console.Write("Distrito Id: ");
        string distritoId = Console.ReadLine();

        Console.Write("Sección Provincial Id: ");
        string seccionProvincialId = Console.ReadLine();

        Console.Write("Sección Id: ");
        string seccionId = Console.ReadLine();

        // Crear un cliente HTTP para hacer solicitudes
        using (var httpClient = new HttpClient())
        {
            // Definir la URL base de la API o servicio que proporciona los datos
            string baseUrl = "https://resultados.mininterior.gob.ar/api"; // Reemplaza con la URL correcta


            int rowIndex = 2; // Empezar desde la segunda fila
            List<string> mesaIds = new List<string> { "1244" }; // Reemplaza con tus mesaIds

            using (var package = new ExcelPackage())
            {
                // Agregar una hoja de trabajo al archivo
                var worksheet = package.Workbook.Worksheets.Add("Resultados");

                foreach (var mesaId in mesaIds)
                {
                    // Construir la URL de la solicitud
                    string apiUrl = $"{baseUrl}?tipoEleccion={tipoEleccion}&distritoId={distritoId}&seccionProvincialId={seccionProvincialId}&seccionId={seccionId}&mesaId={mesaId}";

                    string apiPrueba = "https://resultados.mininterior.gob.ar/api/resultados/getResultados?anioEleccion=2019&tipoRecuento=1&tipoEleccion=2&categoriaId=2&distritoId=1&seccionProvincialId=0&seccionId=3&circuitoId=000039&mesaId=1244&categoriaId=10";
                    // Hacer la solicitud GET

                    HttpResponseMessage response = await httpClient.GetAsync(apiPrueba);

                    string jsonResult = await response.Content.ReadAsStringAsync();

                    // Analizar el JSON en un objeto JObject
                    JObject jsonObject = JObject.Parse(jsonResult);



                    // Establecer valores fijos para Circuito y Mesa (puedes cambiarlos según tus necesidades)
                    worksheet.Cells[1, 1].Value = "Mesa";
                    worksheet.Cells[1, 2].Value = "Circuito";
                    worksheet.Cells[2, 1].Value = "1244";
                    worksheet.Cells[2, 2].Value = "00039";

                    // Obtener la cantidad de votantes de estadoRecuento
                    worksheet.Cells[1, 3].Value = "Cantidad de Votantes";
                    int cantidadVotantes = (int)jsonObject["estadoRecuento"]["cantidadVotantes"];
                    worksheet.Cells[2, 3].Value = cantidadVotantes;

                    // Obtener los valores de VOTOS NULOS, VOTOS RECURRIDOS, VOTOS IMPUGNADOS, VOTOS DEL COMANDO ELECTORAL y VOTOS EN BLANCO
                   

                    // Obtener los valores de votos y nombres de agrupaciones de valoresTotalizadosPositivos
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
                        worksheet.Cells[1, columnNumber].Value = nombreAgrupacion;
                        worksheet.Cells[2, columnNumber].Value = votos;
                        columnNumber++;
                    }


                    foreach (var agrupacionNombre in valoresAgrupaciones.Keys)
                    {
                        worksheet.Cells[1, columnNumber].Value = agrupacionNombre;

                        // Obtener el valor del diccionario o establecerlo en 0 si falta
                        if (valoresAgrupaciones.TryGetValue(agrupacionNombre, out int votos))
                        {
                            worksheet.Cells[2, columnNumber].Value = votos;
                        }
                        else
                        {
                            worksheet.Cells[2, columnNumber].Value = 0;
                        }

                        columnNumber++;
                    }



                    var valoresTotalizadosOtros = jsonObject["valoresTotalizadosOtros"];
                    worksheet.Cells[1, columnNumber].Value = "Votos Nulos";
                    worksheet.Cells[2, columnNumber].Value = (int)valoresTotalizadosOtros["votosNulos"];
                    columnNumber++;
                    worksheet.Cells[1, columnNumber].Value = "Votos Recurridos";
                    worksheet.Cells[2, columnNumber].Value = (int)valoresTotalizadosOtros["votosRecurridosComandoImpugnados"];
                    columnNumber++;
                    worksheet.Cells[1, columnNumber].Value = "Votos impugnados";
                    worksheet.Cells[2, columnNumber].Value = (int)valoresTotalizadosOtros["votosRecurridosComandoImpugnadosPorcentaje"];
                    columnNumber++;
                    worksheet.Cells[1, columnNumber].Value = "Votos en blanco";
                    worksheet.Cells[2, columnNumber].Value = (int)valoresTotalizadosOtros["votosEnBlanco"];


                }
                // Guardar el archivo Excel en disco
                FileInfo excelFile = new FileInfo("Resultados.xlsx");
                package.SaveAs(excelFile);
            

            }


        }

        Console.WriteLine("Archivo Excel generado con éxito: Resultados.xlsx");

    }

    
}

