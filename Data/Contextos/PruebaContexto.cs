using ExportarUnMillonEnExcel.Data.Entidades;
using Microsoft.EntityFrameworkCore;

namespace ExportarUnMillonEnExcel.Data.Contextos;

public class PruebaContexto : DbContext
{
    public DbSet<Registro> Registro { get; set; }

    public PruebaContexto(DbContextOptions<PruebaContexto> options) : base(options) { }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Registro>(x =>
        {
            x.ToTable("Registro").HasKey(y => y.Id);

            x.Property(y => y.Id).HasColumnName("id");
            x.Property(y => y.Nombre).HasColumnName("nombre");
            x.Property(y => y.Apellido).HasColumnName("apellido");
            x.Property(y => y.Nacimiento).HasColumnName("nacimiento");
            x.Property(y => y.Sueldo).HasColumnName("sueldo").HasPrecision(20, 2);
        });

        base.OnModelCreating(modelBuilder);
    }

}
