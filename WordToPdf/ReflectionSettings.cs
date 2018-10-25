using System.IO;
using System.Reflection;

namespace WordToPdf
{
   public static class ReflectionProperty
    {
        public static readonly string AssemblyFolder =   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
    }
}
