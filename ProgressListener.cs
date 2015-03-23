using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace cbr2pdf
{
    public interface ProgessListener
    {
        void progressUpdate(ProcessFile pcsF, int percentageCompleted);
    }
}
