using Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace CoversorXLSXtoXLS
{
    class Program
    {
        public static void ConvertXLSX_XLS(string pathArquivo, int opcaoConversao)
        {
            var app = new Application();
            var wb = app.Workbooks.Open(pathArquivo);

            string nomeArquivo = Path.GetFileNameWithoutExtension(pathArquivo);
            string diretorio = Path.GetDirectoryName(pathArquivo);
            string extencao = Path.GetExtension(pathArquivo);

            string novoPath = Path.Combine(diretorio, nomeArquivo);
            var xlsxFile = opcaoConversao == 1 ? novoPath + ".xls" : novoPath + ".xlsx";
            wb.SaveAs(Filename: xlsxFile, FileFormat: XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();

            Console.WriteLine("Arquivo " +  xlsxFile + " salvo com sucesso!!!");
        }

        public static void ConvertAllFiles(string pathFolder, int opcaoConversao)
        {
            DirectoryInfo directory = new DirectoryInfo(pathFolder);
            FileInfo[] Files = opcaoConversao == 1 ? directory.GetFiles("*.xlsx") : directory.GetFiles("*.xls");
            int count = 0;

            foreach (FileInfo file in Files)
            {
                ConvertXLSX_XLS(file.FullName, opcaoConversao);
                count++;
                Console.WriteLine("Convertido " + count + " arquivos de um total de " + Files.Length +"\r");
            }
        }

        static void Main(string[] args)
        {
            Console.Title = "Conversor .XLSX para .XLS";
            bool continuarAplicacao = true;

            while (continuarAplicacao)
            {
                Console.WriteLine("Conversor de arquivo XLSX para XLS\r");
                Console.WriteLine("----------------------------------\n");

                Console.WriteLine("Digite o número da opção escolhida:\r");
                Console.WriteLine("\t1 - Converter de .xlsx para .xls");
                Console.WriteLine("\t2 - Converter de .xls para .xlsx");
                int opcaoConversao = Convert.ToInt32(Console.ReadLine());

                string extensao = opcaoConversao == 1 ? ".xlsx" : ".xls";

                Console.WriteLine("Digite o número da opção escolhida:\r");
                Console.WriteLine("\t1 - arquivo");
                Console.WriteLine("\t2 - pasta(irá converter todas as planilhas com extensão " + extensao + " da pasta)\r");
                int opcao = Convert.ToInt32(Console.ReadLine());

                if (opcao == 1)
                {
                    Console.WriteLine("Informe o local do arquivo:\r");
                }
                else if (opcao == 2)
                {
                    Console.WriteLine("Informe o local da pasta:\r");
                }
                else
                {
                    Console.WriteLine("Opção Inválida");
                }

                string localArquivo = Console.ReadLine().ToString().Replace("\"", "");

                try
                {
                    if (opcao == 1)
                    {
                        ConvertXLSX_XLS(localArquivo, opcaoConversao);
                    }
                    else if (opcao == 2)
                    {
                        ConvertAllFiles(localArquivo, opcaoConversao);
                    }

                    Console.WriteLine("Processo concluído!!!\r");
                }
                catch (Exception e)
                {
                    Console.WriteLine("Local Inválido\n" + e.Message);
                }

                Console.WriteLine("Deseja continuar? 's/n'");
                if (Console.ReadLine() == "n") continuarAplicacao = false;
            }
        }
    }
}
