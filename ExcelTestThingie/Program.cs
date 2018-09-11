using ExcelTestThingie.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using ExcelTestThingie.Service;

namespace ExcelTestThingie
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var duelists = GenerateDuelists();
            var titles = new List<string>
            {
                "Epic duelista's",
                "The duelists",
                "Doem dood en verderf Duelists",
                "Duelists thingie"
            };
            var excelService = new ExcelService();
            var random = new Random();

            excelService.AppendDuelistsToExcel(duelists, titles.ElementAt(random.Next(titles.Count)));
            Console.WriteLine("Hello World!");
        }


        private static List<Duelist> GenerateDuelists()
        {
            var duelists = new List<Duelist>();

            var countries = new List<string>
            {
                "NL",
                "JP",
                "US",
                "MA"
            };

            var names = new List<string>
            {
                "Farid",
                "Amali",
                "Sid",
                "Phillips",
                "Seto",
                "Kaiba",
                "Peter",
                "Richter",
                "Belmont"
            };

            var ranks = Enum.GetValues(typeof(Rank));
            var random = new Random();

            for (var i = 0; i < 20; i++)
            {
                duelists.Add(new Duelist()
                {
                    Country = countries.ElementAt(random.Next(countries.Count)),
                    Firstname = names.ElementAt(random.Next(names.Count)),
                    Lastname = names.ElementAt(random.Next(names.Count)),
                    Rank = (Rank)ranks.GetValue(random.Next(ranks.Length))
                });
            }


            return duelists;
        }
    }
}
