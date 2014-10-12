using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.Excel.ExcelObjects
{
   public  abstract class ExcelBaseObject : IDisposable
    {

        public ExcelBaseObject()
        {
            ReleaseComObjectOnDispose = false;
        }

        // deze implementatie is alleen handig als je werkt met een static implementatie van 
        // het object. ReleaseComObject zorgt ervoor dat de referentie naar het comobject wordt verbroken.
        // echter omdat het com obuject al in een RCW zit kan dit tot ongewenst gedrag leiden.
        // http://msdn.microsoft.com/en-us/library/8bwh56xe(v=vs.110).aspx
        // door deze setting op false te zetten kun je de objecten gewoon gebruiken en het opschonen van de 
        // manages objects (inclusief de RCW) overlaten aan de garbage collector.
        // Deze implementatie is dus nog niet final en kan verbeterd worden. 
        // Echter is er tot op heden nog geen reden gevonden om hem op true te zetten.
        // Wat meer info over ReleaseComObjecten is hier te vinden.
        // http://blogs.msdn.com/b/visualstudio/archive/2010/03/01/marshal-releasecomobject-considered-dangerous.aspx

        public static bool ReleaseComObjectOnDispose { get; set; }


        ~ExcelBaseObject()
        {
            Dispose(false);
            
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        internal virtual void Dispose(bool disposing)
        {
           
        }

        internal void ReleaseComObject(object comObject)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject);
                comObject = null;
            }
            catch (Exception )
            {
                comObject = null;
            }
            finally
            {
                GC.Collect();
            }

            
        }

    }
}
