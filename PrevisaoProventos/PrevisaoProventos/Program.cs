using OfficeOpenXml;
using System.IO;
using System;
using System.Globalization;

namespace PrevisaoProventos{
    internal class Program
    {
        static void Main(string[] args)
        {
            double valorFi = 0;
            int meses = 0;
            double valorAporte = 0;
            double valorProventos = 0;
            double valorInicial = 0;
            int quantidade = 0;
            double valorSobra = 0;
            double valorUltimoDiv = 0;
            int quantidadeComprada = 0;
            int quantidadeCompSobra = 0;
            int quantidadeCompProv = 0;
            double valorTotalAportes = 0;
            double valorTotalInvestido = 0;
            double valorTotalSobra = 0;
            double valorTotalProventos = 0;

            string directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Simulador FIs");
            string filePath = Path.Combine(directoryPath, "SimuladorProventos_" + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".xlsx");
            
            // Verificar se o diretório existe e criar se necessário
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
                Console.WriteLine($"Diretório criado: {directoryPath}");
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage()) {
            
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Simulação");

                Console.WriteLine("Bem-vindo\nSimulador de dividendos (FIIS)");
                Console.WriteLine("Vamos aos valores!");

                Console.WriteLine("Insira o valor do FII");
                valorFi = double.Parse(Console.ReadLine(), CultureInfo.InvariantCulture);
                Console.WriteLine("Quantos meses você irá investir?");
                meses = int.Parse(Console.ReadLine());
                Console.WriteLine("Valor de aporte mensal?");
                valorAporte = double.Parse(Console.ReadLine(), CultureInfo.InvariantCulture);
                Console.WriteLine("Investimento inicial?");
                valorInicial = double.Parse(Console.ReadLine(), CultureInfo.InvariantCulture);
                Console.WriteLine("Valor do ultimo dividendo");
                valorUltimoDiv = double.Parse(Console.ReadLine(), CultureInfo.InvariantCulture);

                //Cabeçalho
                worksheet.Cells[1,1].Value = "MM/YYYY";
                worksheet.Cells[1, 2].Value = "Quantidade";
                worksheet.Cells[1, 3].Value = "Valor Sobra";
                worksheet.Cells[1, 4].Value = "Valor Proventos";
                worksheet.Cells[1, 5].Value = "Valor total Investido(S/Reinvestimento)";
                worksheet.Cells[1, 6].Value = "Valor total Investido(C/Reinvestimento)";
                worksheet.Cells[1, 7].Value = "Valor Total Proventos Recebidos";
                worksheet.Cells[1, 8].Value = "Num do Mês";

                int row = 2;

                for (int i = 0; i <= meses; i++)
                {
                    if (i == 0)
                    {
                       
                        quantidade = (int)(valorInicial / valorFi);
                        valorSobra = valorInicial - (quantidade * valorFi);

                       

                    }
                    if (i >= 1 && i <= meses)
                    {


                        valorTotalAportes = valorTotalAportes + valorAporte;
                        valorTotalInvestido = valorInicial + valorTotalAportes;
                        valorTotalSobra = valorTotalProventos + valorTotalInvestido;
                        

                        //Quantidade comprada com valor aportado
                        quantidadeComprada = (int)(valorAporte / valorFi);

                        // Calcula valor de Dividendo recebido no mes
                        valorProventos = quantidade * (valorUltimoDiv);
                        valorTotalProventos = valorTotalProventos + valorProventos;
                        //Inicializa quantidade somando o Saldo inicial 
                        // e os valores comprados com a sobra de cada mes
                        // e a quantidade comprada com dividendos
                        quantidade = quantidade + quantidadeComprada + quantidadeCompSobra + quantidadeCompProv;

                        



                        //Valor que sobra do aporte, somada a sobra do mes anterior 
                        valorSobra = valorSobra + valorAporte - (quantidadeComprada * valorFi);


                        if (valorProventos >= valorFi)
                        {
                            // se o valor da sobra for maior que o do Fundo comprar + cotas com o valor q sobrou
                            quantidadeCompProv = (int)(valorProventos / valorFi);

                            //acrescenta o valor da sobra do provento
                            valorSobra = valorSobra + (valorProventos - (quantidadeCompProv * valorFi));

                        }
                        if (valorSobra >= valorFi)
                        {
                            // se o valor da sobra for maior que o do Fundo comprar + cotas com o valor q sobrou
                            quantidadeCompSobra = (int)(valorSobra / valorFi);

                            // subtrai o valor do FI na sobra
                            valorSobra = valorSobra - (valorFi * quantidadeCompSobra);
                        }
                    }

                    worksheet.Cells[row, 1].Value = DateTime.Now.AddMonths(i).ToString("MM/yyyy");
                    worksheet.Cells[row, 2].Value = quantidade;
                    worksheet.Cells[row, 3].Value = valorSobra.ToString("F2");
                    worksheet.Cells[row, 4].Value = valorProventos.ToString("F2");
                    worksheet.Cells[row, 5].Value = valorTotalAportes.ToString("F2");
                    worksheet.Cells[row, 6].Value = valorTotalSobra.ToString("F2");
                    worksheet.Cells[row, 7].Value = valorTotalProventos.ToString("F2");
                    worksheet.Cells[row, 8].Value = i;
                    


                    /*Console.WriteLine($"Valor proventos:{i} | {valorProventos:F2} | {valorSobra:F2} | {quantidade} |" +
                        $" Valor Total Aportes: {valorTotalAportes:F2} |" +
                        $" Valor Total Investido(S/Reinvestimento): {valorTotalInvestido:F2} |" +
                        $" Valor total Proventos: {valorTotalProventos:F2} |" +
                        $" Valor total Investido(C/Reinvestimento): {valorTotalSobra:F2}");*/

                    row++;

                }


                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);

            }

            Console.WriteLine("Arquivo Excel criado!");
        }

                   


            


            
    }
}