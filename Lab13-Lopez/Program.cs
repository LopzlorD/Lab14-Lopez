using ClosedXML.Excel;
using System;

class Program
{
    static void Main(string[] args)
    {
        //FirstExample();
        //SecondExample();
        //ThirdExample();
        FourthExample();
    }

    public static void FirstExample()
    {
        // Crear un nuevo libro de trabajo de Excel
        using (var workbook = new XLWorkbook())
        {
            // Crear una hoja de trabajo
            var worksheet = workbook.Worksheets.Add(sheetName: "Hoja1");

            // Agregar algunos datos
            worksheet.Cell(row: 1, column: 1).Value = "Nombre";
            worksheet.Cell(row: 1, column: 2).Value = "Edad";
            worksheet.Cell(row: 2, column: 1).Value = "Juan";
            worksheet.Cell(row: 2, column: 2).Value = 28;

            // Guardar el archivo en la ubicación deseada
            workbook.SaveAs(file: @"C:\Users\USER\Downloads\archivo.xlsx");
        }

        Console.WriteLine("Archivo Excel creado con éxito.");
    }
    
    public static void SecondExample()
    {
        // Abrir el archivo de Excel existente
        using (var workbook = new XLWorkbook(file: @"C:\Users\USER\Downloads\archivo.xlsx"))
        {
            // Obtener la primera hoja de trabajo
            var worksheet = workbook.Worksheet(position: 1);

            // Modificar una celda existente
            worksheet.Cell(row: 2, column: 2).Value = 30;

            // Guardar el archivo con los cambios
            workbook.Save();
        }

        Console.WriteLine("Archivo Excel modificado con éxito.");
    }
    
    public static void ThirdExample()
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add(sheetName: "Datos");

            // Agregar los datos de la tabla
            worksheet.Cell(row: 1, column: 1).Value = "ID";
            worksheet.Cell(row: 1, column: 2).Value = "Nombre";
            worksheet.Cell(row: 1, column: 3).Value = "Edad";

            worksheet.Cell(row: 2, column: 1).Value = 1;
            worksheet.Cell(row: 2, column: 2).Value = "Juan";
            worksheet.Cell(row: 2, column: 3).Value = 28;

            worksheet.Cell(row: 3, column: 1).Value = 2;
            worksheet.Cell(row: 3, column: 2).Value = "María";
            worksheet.Cell(row: 3, column: 3).Value = 34;

            // Crear una tabla con los datos
            var range = worksheet.Range("A1:C3");
            range.CreateTable();

            // Guardar el archivo con la tabla
            workbook.SaveAs(file: @"C:\Users\USER\Downloads\archivo.xlsx");
        }

        Console.WriteLine("Archivo Excel con tabla creado con éxito.");
    }
    
    public static void FourthExample()
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add(sheetName: "Estilos");

            // Aplicar un formato a la primera fila
            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Fill.BackgroundColor = XLColor.Cyan;
            headerRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Agregar datos
            worksheet.Cell(row: 1, column: 1).Value = "ID";
            worksheet.Cell(row: 1, column: 2).Value = "Nombre";
            worksheet.Cell(row: 1, column: 3).Value = "Edad";

            worksheet.Cell(row: 2, column: 1).Value = 1;
            worksheet.Cell(row: 2, column: 2).Value = "Juan";
            worksheet.Cell(row: 2, column: 3).Value = 28;

            // Guardar el archivo con formato
            workbook.SaveAs(file: @"C:\Users\USER\Downloads\archivo.xlsx");
        }

        Console.WriteLine("Archivo Excel con formato creado con éxito.");
    }
}