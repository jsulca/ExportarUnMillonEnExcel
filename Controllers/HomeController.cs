using ClosedXML.Excel;
using ExportarUnMillonEnExcel.Data.Contextos;
using ExportarUnMillonEnExcel.Data.Entidades;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.Diagnostics;

namespace ExportarUnMillonEnExcel.Controllers;

public class HomeController : Controller
{
    private readonly PruebaContexto _contexto;
    private const string URL_EXCEL = "D:\\Documentos\\Personales\\ExportarUnMillonEnExcel\\wwwroot\\reportes\\Reporte_Registro.xlsx";
    private const string MIMETYPE_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public HomeController(PruebaContexto contexto)
    {
        _contexto = contexto;
    }

    public IActionResult Index()
    {
        return View();
    }

    public IActionResult Privacy()
    {
        return View();
    }

    public async Task<IActionResult> DescargarExcelEPPLUS()
    {
        byte[] data;
        long tiempoLectura, tiempoEscritura;
        Stopwatch stopwatch = Stopwatch.StartNew();
        List<Registro> lista = await _contexto.Registro.ToListAsync();
        stopwatch.Stop();
        tiempoLectura = stopwatch.ElapsedMilliseconds;

        stopwatch.Restart();

        using var paquete = new ExcelPackage(URL_EXCEL);
        using var hoja = paquete.Workbook.Worksheets[0];

        int i = 0, fila = 4;

        foreach(Registro item in lista)
        {
            hoja.Cells[fila, 2].Value = ++i;
            hoja.Cells[fila, 3].Value = item.Nombre;
            hoja.Cells[fila, 4].Value = item.Apellido;
            hoja.Cells[fila, 5].Value = item.Nacimiento.ToShortDateString();
            hoja.Cells[fila, 6].Value = item.Sueldo;

            fila++;
        }

        stopwatch.Stop();
        tiempoEscritura = stopwatch.ElapsedMilliseconds;

        hoja.Cells[1, 3].Value = tiempoLectura;
        hoja.Cells[2, 3].Value = tiempoEscritura;

        data = paquete.GetAsByteArray();

        return File(data, MIMETYPE_XLSX, "Prueba_EPPLUS.xlsx");
    }

    public async Task<IActionResult> DescargarExcelNPOI()
    {
        byte[] data;
        long tiempoLectura, tiempoEscritura;
        Stopwatch stopwatch = Stopwatch.StartNew();
        List<Registro> lista = await _contexto.Registro.ToListAsync();
        stopwatch.Stop();
        tiempoLectura = stopwatch.ElapsedMilliseconds;

        stopwatch.Restart();

        using FileStream fileStream = new(URL_EXCEL, FileMode.Open, FileAccess.Read);
        XSSFWorkbook workbook = new(fileStream);
        ISheet hoja = workbook[0];
        IRow row;

        int i = 0, fila = 3;

        foreach(Registro item in lista)
        {
            row = hoja.CreateRow(fila);
            row.CreateCell(1).SetCellValue(++i);
            row.CreateCell(2).SetCellValue(item.Nombre);
            row.CreateCell(3).SetCellValue(item.Apellido);
            row.CreateCell(4).SetCellValue(item.Nacimiento.ToShortDateString());
            row.CreateCell(5).SetCellValue(item.Sueldo.ToString());

            fila++;
        }

        stopwatch.Stop();
        tiempoEscritura = stopwatch.ElapsedMilliseconds;

        row = hoja.GetRow(0);
        row.CreateCell(2).SetCellValue(tiempoLectura);
        row = hoja.GetRow(1);
        row.CreateCell(2).SetCellValue(tiempoEscritura);

        using MemoryStream ms = new();
        workbook.Write(ms);
        data = ms.ToArray();

        return File(data, MIMETYPE_XLSX, "Prueba_NPOI.xlsx");
    }


    public async Task<IActionResult> DescargarExcelClosedXML()
    {
        byte[] data;
        long tiempoLectura, tiempoEscritura;
        Stopwatch stopwatch = Stopwatch.StartNew();
        List<Registro> lista = await _contexto.Registro.ToListAsync();
        stopwatch.Stop();
        tiempoLectura = stopwatch.ElapsedMilliseconds;

        stopwatch.Restart();

        using XLWorkbook workbook = new(URL_EXCEL);
        IXLWorksheet hoja = workbook.Worksheet(1);

        int i = 0, fila = 4;
       
        foreach(Registro item in lista)
        {
            hoja.Cell(fila, 2).Value = ++i;
            hoja.Cell(fila, 3).Value = item.Nombre;
            hoja.Cell(fila, 4).Value = item.Apellido;
            hoja.Cell(fila, 5).Value = item.Nacimiento.ToShortDateString();
            hoja.Cell(fila, 6).Value = item.Sueldo;

            fila++;
        }

        stopwatch.Stop();
        tiempoEscritura = stopwatch.ElapsedMilliseconds;

        hoja.Cell(1, 3).Value = tiempoLectura;
        hoja.Cell(1, 4).Value = tiempoEscritura;

        using MemoryStream ms = new();
        workbook.SaveAs(ms);
        data = ms.ToArray();

        return File(data, MIMETYPE_XLSX, "Prueba_ClosedXML.xlsx");
    }
}