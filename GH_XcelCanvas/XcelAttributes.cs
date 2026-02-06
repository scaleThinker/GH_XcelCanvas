using System;
using System.Drawing; // Importante para Color, Rectangle, Graphics
using System.Collections.Generic;
using Grasshopper.GUI;
using Grasshopper.GUI.Canvas;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Attributes;

namespace GH_XcelCanvas
{
    public class XcelAttributes : GH_ComponentAttributes
    {
        public XcelAttributes(IGH_Component owner) : base(owner) { }

        protected override void Layout()
        {
            // Calcula o tamanho padrão (inputs/outputs)
            base.Layout();

            RectangleF rec = Bounds;

            // Força largura mínima
            if (rec.Width < 300) rec.Width = 300;

            // Adiciona altura para a tabela
            rec.Height += 200;

            Bounds = rec;
        }

        protected override void Render(GH_Canvas canvas, Graphics graphics, GH_CanvasChannel channel)
        {
            if (channel != GH_CanvasChannel.Objects)
            {
                base.Render(canvas, graphics, channel);
                return;
            }

            // 1. CRIA A CÁPSULA (O formato do componente)
            GH_Capsule capsule = GH_Capsule.CreateCapsule(Bounds, GH_Palette.Hidden);

            // --- CORREÇÃO 1: Adicionar Grips (bolinhas) um por um ---
            // O método no plural foi removido, então fazemos um loop
            foreach (var param in Owner.Params.Input)
            {
                // Adiciona o grip na altura Y correta
                capsule.AddInputGrip(param.Attributes.InputGrip.Y);
            }
            foreach (var param in Owner.Params.Output)
            {
                capsule.AddOutputGrip(param.Attributes.OutputGrip.Y);
            }

            // 2. DESENHA A BASE DO COMPONENTE
            // Renderiza a forma preta arredondada e as conexões
            capsule.Render(graphics, Selected, Owner.Locked, false);

            // Libera a memória da cápsula (importante fazer logo após usar)
            capsule.Dispose();

            // 3. DESENHA O PAPEL BRANCO (A "Tela" do Excel)
            // Criamos um retângulo um pouco menor que a borda para não cobrir o arredondado
            RectangleF panelRect = Bounds;
            panelRect.Inflate(-2, -2);

            // --- CORREÇÃO 2: Cores Manuais ---
            // Em vez de chamar GH_Skin (que dá erro), definimos a cor na mão.
            // Se Selecionado = Verde Claro, Se Normal = Branco
            Color corFundo = Selected ? Color.FromArgb(200, 255, 200) : Color.White;

            // Pinta o fundo
            graphics.FillRectangle(new SolidBrush(corFundo), panelRect);

            // 4. DESENHA A GRADE (Lógica da Tabela)
            var myComponent = Owner as XcelReader;

            if (myComponent != null && myComponent.CachedData != null)
            {
                // Margens e Tamanhos
                float startX = Bounds.X + 5;
                float startY = Bounds.Y + 25; // Pula a área dos inputs
                float cellHeight = 20;

                // Proteção contra divisão por zero
                int cols = myComponent.ColCount > 0 ? myComponent.ColCount : 1;
                float cellWidth = (Bounds.Width - 10) / cols;

                // Loop de desenho das células
                for (int i = 0; i < myComponent.RowCount; i++)
                {
                    for (int j = 0; j < myComponent.ColCount; j++)
                    {
                        RectangleF cellRect = new RectangleF(
                            startX + (j * cellWidth),
                            startY + (i * cellHeight),
                            cellWidth,
                            cellHeight
                        );

                        // Desenha borda da célula
                        graphics.DrawRectangle(Pens.LightGray, Rectangle.Round(cellRect));

                        // Pega o valor
                        string texto = myComponent.CachedData[i, j].Value;

                        // Escreve o texto
                        if (!string.IsNullOrEmpty(texto))
                        {
                            StringFormat format = new StringFormat();
                            format.Alignment = StringAlignment.Center;
                            format.LineAlignment = StringAlignment.Center;
                            format.Trimming = StringTrimming.EllipsisCharacter;

                            graphics.DrawString(texto, GH_FontServer.Small, Brushes.Black, cellRect, format);
                        }
                    }
                }
            }
            else
            {
                // Mensagem de espera
                StringFormat msgFormat = new StringFormat();
                msgFormat.Alignment = StringAlignment.Center;
                msgFormat.LineAlignment = StringAlignment.Center;

                graphics.DrawString("Aguardando Arquivo...",
                    GH_FontServer.Standard, Brushes.Gray, Bounds, msgFormat);
            }
        }
    }
}