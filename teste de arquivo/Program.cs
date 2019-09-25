using System;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

namespace teste_de_arquivo
{
    class Program
    {
        static void Main(string[] args)
        {


            // criando um novo documento com um novo documento.
            Document documento = new Document();

            // Criando uma sessão dentro do documento.
            // a cada sessão criada uma nova página é adicionada.
            Section sessaoCapa = documento.AddSection();

            // insere um título na primeira página
            Paragraph titulo = sessaoCapa.AddParagraph();

            // insere um paragrafo na primeira pagina
            Paragraph paragrafo = sessaoCapa.AddParagraph();
            
            // insiro na minha variável o título, o valor da string "título muito bonito". 
            // Ou seja, no meu documento aparecerá "título muito bonito"....
            titulo.AppendText("Título muito bonito\n\n");

            paragrafo.AppendText("\n\n Este é apenas um parágrafo aleatório");


            // alinha horizontalmente o título
            titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;
            paragrafo.Format.HorizontalAlignment = HorizontalAlignment.Left;

            // instanciando a classe dentro do documento
            ParagraphStyle estilo01 = new ParagraphStyle(documento);
            ParagraphStyle estilo02 = new ParagraphStyle(documento);

            // Define um nome da classe estilo 01.
            estilo01.Name = "Cor do título";
            estilo02.Name = "Cor do parágrafo";

            // Colore a propriedade text-color para azul escuro.
            estilo01.CharacterFormat.TextColor = Color.DarkBlue;
            estilo02.CharacterFormat.TextColor = Color.Red;

            // Transforma a propriedade bold em verdadeiro (true).
            estilo01.CharacterFormat.Bold = true;

            // Adicionar e colocar como usável no documento.
            documento.Styles.Add(estilo01);
            documento.Styles.Add(estilo02);

            // Aplicação das propriedades no título.
            titulo.ApplyStyle(estilo01.Name);  
            paragrafo.ApplyStyle(estilo02.Name); 


            documento.SaveToFile(@"Saida\arquivo_novo.Docx", FileFormat.Docx);


        }
    }
}
