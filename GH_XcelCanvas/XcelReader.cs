using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;

// Referência ao Excel
using Excel = Microsoft.Office.Interop.Excel;

namespace GH_XcelCanvas
{
    public class XcelReader : GH_Component
    {
        // --- ESTRUTURA DE DADOS INTELIGENTE ---
        public struct CellData
        {
            public string Value;    // O resultado (ex: "10")
            public string Formula;  // A lógica (ex: "=A1+B1")
            public string Address;  // Endereço (ex: "C1")
        }

        // Trocamos a matriz simples de string pela nossa struct
        public CellData[,] CachedData = null;

        public int RowCount = 0;
        public int ColCount = 0;
        // ----------------------------------------------------

        public XcelReader()
          : base("Xcel Reader", "XRead",
              "Visualizador de Excel com suporte a Fórmulas e Valores.",
              "ScaleThinker", "Excel")
        {
        }

        public override void CreateAttributes()
        {
            m_attributes = new XcelAttributes(this);
        }

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("File Path", "Path", "Arquivo .xlsx", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Read", "Read", "Ler arquivo", GH_ParamAccess.item, false);
            // Futuramente adicionaremos aqui a seleção de Aba (Sheet)
        }

        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            // SAÍDA 0: Valores Resultantes
            pManager.AddTextParameter("Values", "Val", "Valores resultantes das células", GH_ParamAccess.list);
            // SAÍDA 1: Fórmulas
            pManager.AddTextParameter("Formulas", "Fm", "Fórmulas originais das células", GH_ParamAccess.list);
            // SAÍDA 2: Status
            pManager.AddTextParameter("Status", "Msg", "Status", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string filePath = "";
            bool read = false;

            if (!DA.GetData(0, ref filePath)) return;
            if (!DA.GetData(1, ref read)) return;

            if (!read)
            {
                DA.SetData(2, "Aguardando leitura...");
                return;
            }

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;

            // Listas para as saídas do Grasshopper
            List<string> outValues = new List<string>();
            List<string> outFormulas = new List<string>();

            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;

                xlWorkBook = xlApp.Workbooks.Open(filePath);

                // POR ENQUANTO: Pega a aba ativa (ActiveSheet) para evitar erros de índice
                // Isso resolve parcialmente a questão de referenciar outras abas
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;

                // Define área de leitura (ex: 10x5). Depois faremos dinâmico.
                int numRows = 15;
                int numCols = 5;

                CachedData = new CellData[numRows, numCols];
                RowCount = numRows;
                ColCount = numCols;

                // Loop de Leitura Otimizado
                for (int i = 1; i <= numRows; i++)
                {
                    for (int j = 1; j <= numCols; j++)
                    {
                        var cell = (xlWorkSheet.Cells[i, j] as Excel.Range);

                        string val = "";
                        string form = "";
                        string addr = cell.Address[false, false]; // Pega endereço tipo "A1"

                        // Tenta ler o Valor
                        if (cell.Value2 != null) val = cell.Value2.ToString();

                        // Tenta ler a Fórmula (Se não tiver fórmula, o Excel retorna o valor mesmo)
                        if (cell.Formula != null) form = cell.Formula.ToString();

                        // Preenche nossa memória para o desenho
                        CachedData[i - 1, j - 1] = new CellData
                        {
                            Value = val,
                            Formula = form,
                            Address = addr
                        };

                        // Preenche as listas de saída
                        outValues.Add(val);
                        outFormulas.Add(form);
                    }
                }

                DA.SetDataList(0, outValues);
                DA.SetDataList(1, outFormulas);
                DA.SetData(2, "Leitura Concluída: " + xlWorkSheet.Name);
            }
            catch (Exception ex)
            {
                DA.SetData(2, "Erro: " + ex.Message);
            }
            finally
            {
                // Limpeza completa
                if (xlWorkBook != null)
                {
                    xlWorkBook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }

                xlWorkSheet = null;
                xlWorkBook = null;
                xlApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        protected override System.Drawing.Bitmap Icon => null;
        public override Guid ComponentGuid => new Guid("21b69fd1-db72-47f3-8e0f-1dd083240e2a");
    }
}