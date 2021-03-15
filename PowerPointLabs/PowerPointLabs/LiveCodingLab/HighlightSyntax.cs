using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.LiveCodingLab.Model;
using PowerPointLabs.LiveCodingLab.Service;
using PowerPointLabs.LiveCodingLab.Utility;
using PowerPointLabs.LiveCodingLab.Views;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.LiveCodingLab
{
    public partial class LiveCodingLabMain
    {
#pragma warning disable 0618
        public void HighlightSyntax(PowerPoint.ShapeRange shapes, PowerPointSlide currentSlide)
        {
            foreach (PowerPoint.Shape shape in shapes)
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
                    List<PowerPoint.Effect> effects = AsList(sequence, 1, sequence.Count + 1);
                    List<Tuple<int, PowerPoint.Shape>> effectToIndexList = new List<Tuple<int, PowerPoint.Shape>>();
                    List<int> effectOrderToRestore = new List<int>();

                    for (int i = 0; i < effects.Count; i++)
                    {
                        effectToIndexList.Add(Tuple.Create(i + 1, effects[i].Shape));
                    }

                    ShapeUtility.HighlightSyntax(shape, PowerPointCurrentPresentationInfo.CurrentSlide);

                    foreach (Tuple<int, PowerPoint.Shape> eff in effectToIndexList)
                    {
                        if (eff.Item2.Equals(shape))
                        {
                            effectOrderToRestore.Add(eff.Item1);
                        }
                    }

                    sequence = currentSlide.TimeLine.MainSequence;
                    effects = AsList(sequence, 1, sequence.Count + 1);
                    int effectOrderCounter = effectOrderToRestore.Count - 1;

                    for (int j = effects.Count - 1; j > 0; j--)
                    {
                        if (effects[j].Shape.Equals(shape))
                        {
                            effects[j].MoveTo(effectOrderToRestore[effectOrderCounter]);
                            effectOrderCounter--;
                        }
                    }
                }
            }
        }
    }
}
