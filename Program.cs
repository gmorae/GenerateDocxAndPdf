using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace GenerateDocxAndPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Programa de Gerar nota fiscal \n");
            Console.WriteLine("Digite o nome do usuário: ");
            string nome = Console.ReadLine();
            Console.WriteLine("Digite o endereço: ");
            string endereco = Console.ReadLine();
            Console.WriteLine("Digite o valor da compra: ");
            double valor = double.Parse(Console.ReadLine());

            Document doc = new Document();
            Section secao = doc.AddSection();
            ParagraphStyle estilo = new ParagraphStyle(doc);


            Paragraph titulo = secao.AddParagraph();
            titulo.AppendText("Gerador de Nota Fiscal\n");

            Paragraph VarNome = secao.AddParagraph();
            VarNome.AppendText("Nome: ").CharacterFormat.Bold = true;
            TextRange boldNome = VarNome.AppendText(nome);
            boldNome.CharacterFormat.Bold = false;
        
            Paragraph VarEndereco = secao.AddParagraph();
            VarEndereco.AppendText("Endereço: ").CharacterFormat.Bold = true;
            TextRange boldEnd = VarEndereco.AppendText(endereco);
            boldEnd.CharacterFormat.Bold = false;

            Paragraph VarValor = secao.AddParagraph();
            VarValor.AppendText("Valor: ").CharacterFormat.Bold = true;
            TextRange boldVal = VarValor.AppendText($"{valor} reais");
            boldVal.CharacterFormat.Bold = false;

            DateTime data = DateTime.Now;
            Paragraph VarData = secao.AddParagraph();
            VarData.AppendText("Data: ").CharacterFormat.Bold = true;
            TextRange boldData = VarData.AppendText(data.ToString("dd/MM/yyyy"));
            boldData.CharacterFormat.Bold = false;


            
            doc.SaveToFile(@"saida\exemplo.docx", FileFormat.Docx);

        }
    }
}
