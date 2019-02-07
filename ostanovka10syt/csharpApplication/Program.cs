using System;

namespace csharpApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Старт приложения!");
            var main = new Main();
            main.StartProcess();
            Console.WriteLine("Завершено!");
        }
    }
}
