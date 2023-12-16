using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using archlab9;

namespace archlabab9
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var graphDrawer = new FuncGraphDrawer();


            var formatter = new TitlePageFormatter();
            var commands = Enum.GetValues(typeof(Command)).Cast<Command>().ToList();
            var command = Command.Exit;
            do
            {
                commands.ForEach(com => Console.WriteLine($"{(int)com}. {com.ToString()}"));

                var ok = true;
                command = readCommand(out ok);
                if (!ok)
                {
                    Console.WriteLine("Wrong command");
                    continue;
                }

                switch (command)
                {
                    case Command.CreateTitlePage:
                        {
                            var pageData = new TitlePageData()
                            {
                                WorkType = readString("Тип работы"),
                                WorkNumber = readString("Номер работы"),
                                Title = readString("Название работы"),
                                Discilpline = readString("Название дисциплины"),
                                Teacher = readString("ФИО преподавателя"),
                            };
                            formatter.Format("C:\\Users\\Андрей Лузгин\\OneDrive\\Desktop\\учеба\\архитектура ИС\\9\\result.doc", pageData);
                        }
                        break;
                    case Command.CreateGraph:
                        {
                            graphDrawer.DrawGraph(-10, 10);
                        }
                        break;

                }




            } while (command != Command.Exit);
        }

        private static Command readCommand(out bool succeed)
        {
            Console.WriteLine("Enter command: ");
            succeed = false;

            var commandStr = Console.ReadLine();
            Command command;
            succeed = Enum.TryParse<Command>(commandStr, out command);
            return command;
        }

        private int readInt(string name, out bool succeed)
        {
            Console.WriteLine($"Enter {name}: ");
            succeed = false;

            var commandStr = Console.ReadLine();
            int res;
            succeed = int.TryParse(commandStr, out res);
            return res;
        }

        private static string readString(string name)
        {
            Console.WriteLine($"Enter {name}: ");
            return Console.ReadLine();
        }

        private enum Command
        {
            CreateTitlePage = 1,
            CreateGraph,
            Exit,
        }
    }
}