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

        var filePath = "ResultadosEleccion.xlsx";

        // Lista de cargos
        var cargos = new Dictionary<string, string>
            {
                { "PRESIDENTE", "1" },
                { "PARLAMENTARIO MERCOSUR NACIONAL", "2" },
                { "SENADOR NACIONAL", "3" },
                { "DIPUTADO NACIONAL", "4" },
                { "PARLAMENTARIO MERCOSUR PROVINCIAL", "5" },
                { "GOBERNADOR", "6" },
                { "SENADOR PROVINCIAL", "7" },
                { "DIPUTADO PROVINCIAL", "8" },
                { "INTENDENTE", "9" },
                { "CONCEJAL", "10" }
            };


        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {

            var worksheet = package.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "ResultadosEleccion");

            if (worksheet != null)
            {
                package.Workbook.Worksheets.Delete(worksheet);
            }

            // Crear una hoja en el archivo Excel
            worksheet = package.Workbook.Worksheets.Add("ResultadosEleccion");

            // Agregar encabezados de columna
            var columnHeaders = new List<string>
                {
                    "Circuito",
                    "Mesa",
                    "Cargo",
                    "Cantidad de Votantes",
                    "UNITE POR LA LIBERTAD Y LA DIGNIDAD",
                    "JUNTOS POR EL CAMBIO",
                    "FRENTE NOS",
                    "FRENTE DE IZQUIERDA Y DE TRABAJADORES - UNIDAD",
                    "FRENTE DE TODOS",
                    "CONSENSO FEDERAL",
                    "Votos Nulos",
                    "Votos Impugnados",
                    "Votos en Blanco"
                };

            for (int i = 0; i < columnHeaders.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = columnHeaders[i];
            }

            // Llenar las filas con datos
            int rowIndex = 2; // Comenzar en la segunda fila

            foreach (var cargo in cargos)
            {
                var categoriaId = cargo.Value;
                var response = await ObtenerResultadoEleccionAsync(categoriaId);

                if (response != null && response.valoresTotalizadosPositivos.Count > 0)
                {
                    worksheet.Cells[rowIndex, 1].Value = "000039";
                    worksheet.Cells[rowIndex, 2].Value = "1244"; // Cantidad de Votantes
                    worksheet.Cells[rowIndex, 3].Value = cargo.Key; // Cargo
                    worksheet.Cells[rowIndex, 4].Value = response.estadoRecuento?.cantidadVotantes ?? 0; // Cantidad de Votantes

                    // Llenar otras columnas de votos según la respuesta JSON
                    // Ejemplo:
                    worksheet.Cells[rowIndex, 5].Value = ObtenerVotosAgrupacion(response, "UNITE POR LA LIBERTAD Y LA DIGNIDAD");
                    worksheet.Cells[rowIndex, 6].Value = ObtenerVotosAgrupacion(response, "JUNTOS POR EL CAMBIO");
                    worksheet.Cells[rowIndex, 7].Value = ObtenerVotosAgrupacion(response, "FRENTE NOS");
                    worksheet.Cells[rowIndex, 8].Value = ObtenerVotosAgrupacion(response, "FRENTE DE IZQUIERDA Y DE TRABAJADORES - UNIDAD");
                    worksheet.Cells[rowIndex, 9].Value = ObtenerVotosAgrupacion(response, "FRENTE DE TODOS");
                    worksheet.Cells[rowIndex, 10].Value = ObtenerVotosAgrupacion(response, "CONSENSO FEDERAL");


                    worksheet.Cells[rowIndex, 11].Value = response.valoresTotalizadosOtros?.votosNulos ?? 0; // Votos Nulos
                    worksheet.Cells[rowIndex, 12].Value = response.valoresTotalizadosOtros?.votosRecurridosComandoImpugnados ?? 0; // Votos Impugnados
                    worksheet.Cells[rowIndex, 13].Value = response.valoresTotalizadosOtros?.votosEnBlanco ?? 0; // Votos en Blanco

                    rowIndex++;
                }
            }

            // Guardar el archivo Excel
            package.Save();
        }

        Console.WriteLine($"Archivo Excel generado en: {filePath}");


    }



    static async Task<RespuestaEleccion> ObtenerResultadoEleccionAsync(string categoriaId)
    {
        using (var httpClient = new HttpClient())
        {
            var url = $"https://resultados.mininterior.gob.ar/api/resultados/getResultados?anioEleccion=2019&tipoRecuento=1&tipoEleccion=2&distritoId=1&seccionProvincialId=0&seccionId=3&circuitoId=000039&mesaId=1244&categoriaId={categoriaId}";

            var response = await httpClient.GetAsync(url);

            if (response.IsSuccessStatusCode)
            {
                var jsonResponse = await response.Content.ReadAsStringAsync();
                var resultado = Newtonsoft.Json.JsonConvert.DeserializeObject<RespuestaEleccion>(jsonResponse);
                return resultado;
            }
        }

        return null;
    }

    static int ObtenerVotosAgrupacion(RespuestaEleccion respuesta, string nombreAgrupacion)
    {
        foreach (var agrupacion in respuesta.valoresTotalizadosPositivos)
        {
            if (agrupacion.nombreAgrupacion == nombreAgrupacion)
            {
                return agrupacion.votos;
            }
        }

        return 0;
    }

    public class RespuestaEleccion
    {
        public DateTime fechaTotalizacion { get; set; }
        public EstadoRecuento estadoRecuento { get; set; }
        public List<ValoresTotalizadosPositivos> valoresTotalizadosPositivos { get; set; }
        public ValoresTotalizadosOtros valoresTotalizadosOtros { get; set; }
    }

    public class EstadoRecuento
    {
        public int cantidadVotantes { get; set; }
    }

    public class ValoresTotalizadosPositivos
    {
        public string nombreAgrupacion { get; set; }
        public int votos { get; set; }
    }

    public class ValoresTotalizadosOtros
    {
        public int votosNulos { get; set; }
        public double? votosNulosPorcentaje { get; set; }
        public int votosEnBlanco { get; set; }
        public double? votosEnBlancoPorcentaje { get; set; }
        public int votosRecurridosComandoImpugnados { get; set; }
        public double? votosRecurridosComandoImpugnadosPorcentaje { get; set; }
    }


}

