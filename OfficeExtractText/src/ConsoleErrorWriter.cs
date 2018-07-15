using System;
using System.IO;
using System.Text;

namespace OfficeExtractText
{
    class ConsoleErrorWriter : TextWriter
    {
        private ConsoleColor ForegroundColor { get; }
        private ConsoleColor BackgroudnColor { get; }
        private TextWriter originalConsoleStream;

        public ConsoleErrorWriter(TextWriter consoleTextWriter, ConsoleColor foregroundColor, ConsoleColor backgroudnColor)
        {
            originalConsoleStream = consoleTextWriter;
            this.ForegroundColor = foregroundColor;
            this.BackgroudnColor = backgroudnColor;
        }

        public override void WriteLine(string value)
        {
            ConsoleColor originalForegroundColor = Console.ForegroundColor;
            ConsoleColor originalBackgroundColor = Console.BackgroundColor;
            Console.ForegroundColor = ForegroundColor;
            Console.BackgroundColor = BackgroudnColor;

            originalConsoleStream.WriteLine(value);

            Console.ForegroundColor = originalForegroundColor;
            Console.BackgroundColor = originalBackgroundColor;
        }

        public override Encoding Encoding
        {
            get { return Encoding.Default; }
        }

        public static void SetToConsoleError(ConsoleColor foregroundColor = ConsoleColor.Red, ConsoleColor backgroundColor = ConsoleColor.Black)
        {
            Console.SetError(new ConsoleErrorWriter(Console.Error, foregroundColor, backgroundColor));
        }
    }
}
