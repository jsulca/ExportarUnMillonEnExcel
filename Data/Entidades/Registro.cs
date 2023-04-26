namespace ExportarUnMillonEnExcel.Data.Entidades;

public class Registro
{
    public int Id { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public string Apellido { get; set; } = string.Empty;
    public DateTime Nacimiento { get; set; }
    public decimal Sueldo { get; set; }
}
