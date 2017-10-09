using System;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace Semagsoft
{
    public class ImageHelper
    {
        AdornerLayer al;

        public void AddImageResizers(RichTextBox editor)
        {
            var images = GetVisuals(editor).OfType<Image>();
            al = AdornerLayer.GetAdornerLayer(editor);
            foreach (var image in images)
            {
                //ResizingAdorner ral = new ResizingAdorner(image);
                al.Add(new ResizingAdorner(image));
                al.Update(image);
                //LIBTODO:
            }
        }

        static IEnumerable<DependencyObject> GetVisuals(DependencyObject root)
        {
            foreach (var child in LogicalTreeHelper.GetChildren(root).OfType<DependencyObject>())
            {
                yield return child;
                foreach (var descendants in GetVisuals(child))
                    yield return descendants;
            }
        }
    }
}