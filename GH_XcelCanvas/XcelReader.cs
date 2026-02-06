using System;
using System.Collections.Generic;
using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using Excel = Microsoft.Office.Interop.Excel;

namespace GH_XcelCanvas
{
    public class XcelReader : GH_Component
    {
        /// <summary>
        /// Construtor: Define o nome, apelido e descrição do componente.
        /// </summary>
        public XcelReader()
          : base("Xcel Reader", "XRead",
              "Lê um arquivo Excel e mostra os dados no Canvas.",
              "ScaleThinker", "Excel")
        {
        }

        /// <summary>
        /// Define os Inputs do componente.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            // Índice 0: Caminho do arquivo
            pManager.AddTextParameter("File Path", "Path", "Localização do arquivo .xlsx", GH_ParamAccess.item);
            // Índice 1: Botão de leitura
            pManager.AddBooleanParameter("Read", "Read", "Defina como True para ler o arquivo", GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Define os Output do componente.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            // Índice 0: Os dados lidos
            pManager.AddTextParameter("Data", "Data", "Dados lidos do Excel", GH_ParamAccess.list);
            // Índice 1: Mensagens de status (para debug)
            pManager.AddTextParameter("Status", "Msg", "Status da operação", GH_ParamAccess.item);
        }

        /// <summary>
        /// O CEREBRO
        /// </summary>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // 1. Variáveis para receber os inputs
            string filePath = "";
            bool read = false;

            // 2. Coletando os dados do usuário
            if (!DA.GetData(0, ref filePath)) return; // Se não tiver caminho, para aqui.
            if (!DA.GetData(1, ref read)) return;     // Se não tiver booleano, para aqui.

            // Se o botão for False, a gente avisa e para (economiza processamento)
            if (!read)
            {
                DA.SetData(1, "Aguardando sinal de leitura (True)...");
                return;
            }

            // 3. Preparando as variáveis do Excel
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;

            List<string> readData = new List<string>();

            try
            {
                // Inicia o Excel em "segundo plano" (invisível)
                xlApp = new Excel.Application();
                xlApp.Visible = false;

                // Abre o arquivo
                xlWorkBook = xlApp.Workbooks.Open(filePath);
                // Pega a primeira aba (Planilha1)
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                // --- LEITURA DE TESTE (Células A1 a A5) ---
                // Depois vamos mudar isso para ler o que o usuário quiser
                for (int i = 1; i <= 5; i++)
                {
                    // No Excel, [linha, coluna]. Coluna 1 = A.
                    var cell = (xlWorkSheet.Cells[i, 1] as Excel.Range);

                    if (cell.Value != null)
                        readData.Add(cell.Value.ToString());
                    else
                        readData.Add("<Vazio>");
                }

                // 4. Entregando o resultado para o Grasshopper
                DA.SetDataList(0, readData);
                DA.SetData(1, "Sucesso! Arquivo lido.");
            }
            catch (Exception ex)
            {
                // Se der erro (ex: arquivo não existe), avisa o usuário
                DA.SetData(1, "Erro: " + ex.Message);
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ex.Message);
            }
            finally
            {
                // 5. LIMPEZA DE MEMÓRIA (Muito Importante!)
                // Fecha o Excel para não ficar travando o PC
                if (xlWorkBook != null)
                {
                    xlWorkBook.Close(false); // Fecha sem salvar
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
                }
                if (xlApp != null)
                {
                    xlApp.Quit(); // Mata o processo do Excel
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }

                // Zera as variáveis
                xlWorkSheet = null;
                xlWorkBook = null;
                xlApp = null;

                // Força o Windows a limpar a memória RAM agora
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Ícone do componente. Por enquanto nulo.
        /// </summary>
        protected override System.Drawing.Bitmap Icon => null;

        /// <summary>
        /// ID unica
        /// </summary>
        public override Guid ComponentGuid => new Guid("21b69fd1-db72-47f3-8e0f-1dd083240e2a");
    }
}