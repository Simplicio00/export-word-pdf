using System;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Formatting;
using System.Drawing;

namespace exercicio1
{
    class Program
    {
        static void Main(string[] args)
        {
            string nome;
            int valorcompra;
            DateTime data;
            string endereco;
            
            Console.WriteLine("Digite o nome de uma pessoa");
            nome = Console.ReadLine();
            Console.WriteLine(" Digite o valor da compra");
            valorcompra = int.Parse(Console.ReadLine());
            Console.WriteLine(" Digite o dia da compra");
            data = DateTime.Parse(Console.ReadLine());
            Console.WriteLine(" Digite o endereço");
            endereco = Console.ReadLine();


            Document documento = new Document();

            Section sessaoCapa = documento.AddSection();

            Paragraph nomedigitado = sessaoCapa.AddParagraph();
            Paragraph compradigitado = sessaoCapa.AddParagraph();
            Paragraph datadigitado = sessaoCapa.AddParagraph();
            Paragraph enderecodigitado = sessaoCapa.AddParagraph();

            

            CharacterFormat format = new CharacterFormat(documento);
            
            format.Bold=true;

            nomedigitado.AppendText("Nome: ").ApplyCharacterFormat(format);
            nomedigitado.AppendText(nome);
            nomedigitado.Format.HorizontalAlignment = HorizontalAlignment.Left;

            compradigitado.AppendText("Preco da compra: ").ApplyCharacterFormat(format);
            compradigitado.AppendText($"{valorcompra}");
            compradigitado.Format.HorizontalAlignment = HorizontalAlignment.Left;

            
            datadigitado.AppendText("Digite a data: ").ApplyCharacterFormat(format);
            datadigitado.AppendText($"{data}");
            datadigitado.Format.HorizontalAlignment = HorizontalAlignment.Left;

            enderecodigitado.AppendText("Endereço: ").ApplyCharacterFormat(format);
            enderecodigitado.AppendText(endereco);
            enderecodigitado.Format.HorizontalAlignment = HorizontalAlignment.Left;

            documento.SaveToFile(@"novoarquivodeteste.Docx", FileFormat.Docx);
            

        }
    }
}
