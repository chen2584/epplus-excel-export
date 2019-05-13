using System;
using System.IO;
using System.Reflection;

namespace EpPlusExample
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            var rootPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var fullPath = Path.Combine(rootPath, "test.xlsx");
            using(var fileStream = File.Create(fullPath))
            {
                EpPlusService.Generate(fileStream);
            }
        }
    }
}
