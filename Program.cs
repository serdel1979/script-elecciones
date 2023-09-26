using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
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

            // Crear un archivo Excel
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Resultados");

                // Encabezados de las columnas
                worksheet.Cells[1, 1].Value = "IdMesa";
                worksheet.Cells[1, 2].Value = "NombreAgrupacion";
                worksheet.Cells[1, 3].Value = "VotosNulos";
                worksheet.Cells[1, 4].Value = "VotosNulosPorcentaje";
                worksheet.Cells[1, 5].Value = "VotosEnBlanco";
                worksheet.Cells[1, 6].Value = "VotosEnBlancoPorcentaje";
                worksheet.Cells[1, 7].Value = "VotosRecurridosComandoImpugnados";
                worksheet.Cells[1, 8].Value = "VotosRecurridosComandoImpugnadosPorcentaje";
                worksheet.Cells[1, 9].Value = "Total";

                // Hacer solicitudes para obtener datos y llenar el archivo Excel
                int rowIndex = 2; // Empezar desde la segunda fila
                List<string> mesaIds = new List<string> { "1244" }; // Reemplaza con tus mesaIds

                foreach (var mesaId in mesaIds)
                {
                    // Construir la URL de la solicitud
                    string apiUrl = $"{baseUrl}?tipoEleccion={tipoEleccion}&distritoId={distritoId}&seccionProvincialId={seccionProvincialId}&seccionId={seccionId}&mesaId={mesaId}";

                    string apiPrueba = "https://resultados.mininterior.gob.ar/api/resultados/getResultados?anioEleccion=2019&tipoRecuento=1&tipoEleccion=2&categoriaId=2&distritoId=1&seccionProvincialId=0&seccionId=3&circuitoId=000039&mesaId=1244";
                    // Hacer la solicitud GET
                    HttpResponseMessage response = await httpClient.GetAsync(apiPrueba);

                    if (response.IsSuccessStatusCode)
                    {
                        // Leer la respuesta JSON
                        string jsonResult = await response.Content.ReadAsStringAsync();


                        worksheet.Cells[rowIndex, 1].Value = mesaId;
                        Console.WriteLine(jsonResult);

                        rowIndex++;
                    }
                }

                // Guardar el archivo Excel en disco
                FileInfo excelFile = new FileInfo("Resultados.xlsx");
                package.SaveAs(excelFile);

                Console.WriteLine("Archivo Excel generado con éxito: Resultados.xlsx");
            }
        }
    }
}
