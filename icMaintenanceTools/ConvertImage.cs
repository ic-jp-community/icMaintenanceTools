using System;
using System.Collections.Generic;
using System.Text;

using System.Drawing;

namespace ICApiAddin.icMaintenanceTools
{
    class ConvertImage : System.Windows.Forms.AxHost
    {
        public ConvertImage() : base("")
        {
        }

        public static stdole.IPictureDisp ImageToPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }

    }
}
