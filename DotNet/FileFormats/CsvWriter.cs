/*
Copyright 2016 Roman Polunin

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
using System;
using System.Collections.Generic;
using System.IO;

namespace FileFormats
{
    /// <summary>
    /// Class for writing to comma-separated-value (CSV) files.
    /// </summary>
    public class CsvWriter
    {
        /// <summary>
        /// Stream writer instance.
        /// </summary>
        protected readonly TextWriter Writer;
        /// <summary>
        /// Quoting character.
        /// </summary>
        protected readonly char Quote;
        /// <summary>
        /// Delimiter character.
        /// </summary>
        protected readonly char Delimiter;
        /// <summary>
        /// Special characters that have to enclosed in quotes and/or escaped by duplication.
        /// </summary>
        protected readonly char[] SpecialChars;

        /// <summary>
        /// Ctr.
        /// </summary>
        public CsvWriter(TextWriter writer, char delimiter, char quoteCharacter)
        {
            if (delimiter == quoteCharacter)
            {
                throw new ArgumentException("Delimiter and quote characters cannot be the same");
            }

            if (delimiter == '\0' || delimiter == '\r' || delimiter == '\n')
            {
                throw new ArgumentOutOfRangeException(nameof(delimiter), "Delimiter cannot be a null character, CR or LF");
            }

            if (quoteCharacter == '\0' || quoteCharacter == '\r' || quoteCharacter == '\n')
            {
                throw new ArgumentOutOfRangeException(nameof(quoteCharacter), "Quote cannot be a null character, CR or LF");
            }

            Writer = writer;
            Quote = quoteCharacter;
            Delimiter = delimiter;
            SpecialChars = new[] { '\0', '\r', '\n', Delimiter, Quote };
        }

        /// <summary>
        /// Ctr.
        /// </summary>
        public CsvWriter(TextWriter writer) : this(writer, ',', '"')
        {
        }

        /// <summary>
        /// Writes a row of column values to the current CSV stream.
        /// </summary>
        /// <param name="values">The list of values to write</param>
        public void WriteRow(List<string> values)
        {
            // Write each column
            for (var i = 0; i < values.Count; i++)
            {
                // Add delimiter if this isn't the first column
                if (i > 0)
                {
                    Writer.Write(Delimiter);
                }

                // Write this column
                if (string.IsNullOrEmpty(values[i]))
                {
                    continue;
                }

                // quote anything with CR, LF, quote or delimiter
                if (values[i].IndexOfAny(SpecialChars) == -1)
                {
                    // non-quoted string is written as-is
                    Writer.Write(values[i]);
                }
                else
                {
                    Writer.Write(Quote);
                    foreach (var c in values[i])
                    {
                        Writer.Write(c);
                        if (c == Quote)
                        {
                            // in quoted string, each single quote character is escaped with additional one
                            Writer.Write(c);
                        }
                    }
                    Writer.Write(Quote);
                }
            }
            Writer.WriteLine();
        }
    }
}
