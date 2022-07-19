using System;
using System.Linq;
using ClosedXML.Excel;

namespace LerPlanilhaExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var ListaUsuarios = new List<Usuario>();

            var xls = new XLWorkbook(@"C:\Users\Luiz Pachioni\repositorioDODEV\LerPlanilhaExcel\Usuarios.xlsx");
            var planilha = xls.Worksheets.First(w => w.Name == "Página1");
            var totalLinhas = planilha.Rows().Count();
            // primeira linha é o cabecalho

            for (int l = 2; l <= totalLinhas; l++)
            {
                var usuario = new Usuario()
                {
                    Id = int.Parse(planilha.Cell($"A{l}").Value.ToString()),
                    Email = planilha.Cell($"B{l}").Value.ToString(),
                    Senha = planilha.Cell($"C{l}").Value.ToString(),
                    Nome = planilha.Cell($"D{l}").Value.ToString(),
                    Cpf = planilha.Cell($"E{l}").Value.ToString(),
                    DataNascimento = DateTime.Parse(planilha.Cell($"F{l}").Value.ToString()),
                    UrlImagemCadastro = planilha.Cell($"G{l}").Value.ToString(),
                    DataCriacao = DateTime.Parse(planilha.Cell($"H{l}").Value.ToString())
                };
                ListaUsuarios.Add(usuario);
            }
            foreach (var item in ListaUsuarios)
            {
                Console.WriteLine($"{item.Id} - {item.Email} - {item.Senha} - {item.Nome} - {item.Cpf} - {item.DataNascimento} - {item.UrlImagemCadastro} - {item.DataCriacao}");
            }
        }
    }
}
