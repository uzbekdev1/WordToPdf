using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace WordToPdf
{
   public static class ReflectionProperty
    {
        public static readonly string AssemblyFolder =   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
    }
}
