


namespace System
{
    public static class ConsoleHelper
    {
        public static void Console(this string str, ConsoleColor fontColor=ConsoleColor.White)
        {
            System.Console.ForegroundColor = fontColor;
            System.Console.WriteLine(str);
            System.Console.WriteLine();
            System.Console.ResetColor();
        }
    }
}
