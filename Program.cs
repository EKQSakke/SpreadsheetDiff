var input = "C:/temp/Base.xlsx";
var output = "C:/temp/Output.csv";

Console.WriteLine("Args:");

if (args.Length == 0)
{
    Console.WriteLine("No arguments given");
}
else if (args.Length == 1 && args[1] == "c")
{
    Console.WriteLine("Converting with placeholder files");
    CsvConverter.Convert(input, output);
}
else if (args.Length == 3 && args[1] == "c")
{
    Console.WriteLine($"Converting {args[2]} to csv: {args[3]}");
    CsvConverter.Convert(args[2], args[3]);
}