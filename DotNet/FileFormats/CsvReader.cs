/*
Copyright 2016 Roman Polunin

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace FileFormats
{
    /// <summary>
    /// Utility to parse input CSV text stream with configurable comma and quote characters.
    /// Provides line extractor (<see cref="ConsumeLines"/>) and field extractor (<see cref="ReadField"/>).
    /// Line breaks (ASCII 0x0A) and carriage returns (ASCII 0x0D) are allowed as part of quoted values.
    /// Nulls (ASCII 0x00) are silently eliminated.
    /// Supports optional header. Invoke <see cref="ParseHeader"/> exactly once
    /// prior to using <see cref="ReadField"/>, or <see cref="get_Item(int)"/>, or <see cref="get_Item(string)"/>.
    /// </summary>
    public class CsvReader
    {
        private StringBuilder[] m_columnValues;
        private Dictionary<string, int> m_columnNames;
        private bool m_haveHeader;
        private readonly char m_delimiter;
        private readonly char m_quote;
        private readonly int m_maxLineLength;
        private readonly long m_maxCharCount;

        /// <summary>
        /// Ctr.
        /// </summary>
        public CsvReader()
            : this(',', '"', Int32.MaxValue, Int64.MaxValue)
        {
        }

        /// <summary>
        /// Ctr.
        /// </summary>
        /// <param name="delimiter">Delimiter of individual values, such as single or double quote</param>
        /// <param name="quoteCharacter">Quote character, used to protect delimiters that are part of value</param>
        /// <param name="maxLineLength">Max number of characters in a single line of values</param>
        /// <param name="maxCharCount">Max number of characters to be read from input</param>
        /// <exception cref="ArgumentException">Delimiter and quote characters cannot be the same</exception>
        /// <exception cref="ArgumentOutOfRangeException">Delimiter cannot be a null character, CR or LF</exception>
        /// <exception cref="ArgumentOutOfRangeException">Quote cannot be a null character, CR or LF</exception>
        /// <exception cref="ArgumentOutOfRangeException">Max char count must be at least 1</exception>
        public CsvReader(char delimiter, char quoteCharacter, int maxLineLength, long maxCharCount)
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

            if (maxLineLength < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(maxLineLength), "Max line length must be at least 1");
            }

            if (maxCharCount < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(maxCharCount), "Max char count must be at least 1");
            }

            m_delimiter = delimiter;
            m_quote = quoteCharacter;
            m_maxLineLength = maxLineLength;
            m_maxCharCount = maxCharCount;
        }

        /// <summary>
        /// Number of fields in the header.
        /// </summary>
        /// <exception cref="InvalidOperationException">You must call <see cref="ParseHeader"/> prior to this operation</exception>
        public int ColumnCount { get { RequireHeader(); return m_columnValues.Length; } }

        /// <summary>
        /// Column name by its index in the header.
        /// </summary>
        /// <param name="index">Index of the column in the header</param>
        /// <exception cref="InvalidOperationException">You must call <see cref="ParseHeader"/> prior to this operation</exception>
        /// <exception cref="ArgumentException">Index is not found in the parsed header</exception>
        public string RequireColumnName(int index)
        {
            RequireHeader();

            foreach (var pair in m_columnNames)
            {
                if (pair.Value == index)
                {
                    return pair.Key;
                }
            }

            throw new ArgumentException("Index is not found in the parsed header: " + index);
        }

        /// <summary>
        /// Column index by its name in the header.
        /// </summary>
        /// <param name="name">Name of the column in the header</param>
        /// <exception cref="ArgumentNullException"><paramref name="name"/> is null</exception>
        /// <exception cref="InvalidOperationException">You must call <see cref="ParseHeader"/> prior to this operation</exception>
        /// <exception cref="Exception">Internal assertion failed</exception>
        /// <returns>Column index, or -1 if not found</returns>
        public int FindColumnIndex(string name)
        {
            RequireHeader();

            if (!m_columnNames.TryGetValue(name, out var index))
            {
                return -1;
            }

            if (index < 0 || index >= m_columnValues.Length)
            {
                throw new Exception("Internal error");
            }

            return index;
        }

        /// <summary>
        /// Column value by its index. Allocates a new string from current column value buffer.
        /// </summary>
        /// <param name="index">Index of the column</param>
        /// <exception cref="InvalidOperationException">You must call <see cref="ParseHeader"/> prior to this operation</exception>
        /// <exception cref="IndexOutOfRangeException">Index is out of range of column values</exception>
        
        public string this[int index] { get { RequireHeader(); return m_columnValues[index].ToString(); } }

        /// <summary>
        /// Column value by its name. Allocates a new string from current column value buffer.
        /// </summary>
        /// <param name="name">Name of the column</param>
        /// <exception cref="ArgumentNullException"><paramref name="name"/> is null</exception>
        /// <exception cref="InvalidOperationException">You must call <see cref="ParseHeader"/> prior to this operation</exception>
        /// <exception cref="ArgumentException">Column name is not found in the parsed header</exception>
        
        public string this[string name] { get { RequireHeader(); return m_columnValues[FindColumnIndex(name)].ToString(); } }

        /// <summary>
        /// Initializes CSV header information from a given line of text.
        /// Header line may be empty, but cannot be null.
        /// </summary>
        /// <param name="line">Header line, may be empty, but cannot be null</param>
        /// <exception cref="ArgumentNullException"><paramref name="line"/> is null</exception>
        /// <exception cref="CsvReaderException">Some failure during extraction of column headers, see <see cref="ReadField"/> for more info</exception>
        public void ParseHeader(StringBuilder line)
        {
            // erase current state
            m_columnValues = null;
            m_haveHeader = false;
            m_columnNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            var value = new StringBuilder();
            var colIndex = 0;
            var offset = 0;
            while (ReadField(line, ref offset, value))
            {
                var colName = value.ToString();
                m_columnNames.Add(colName, colIndex);
                colIndex++;
                value.Clear();
            }

            m_columnValues = new StringBuilder[m_columnNames.Count];
            for (colIndex = 0; colIndex < m_columnValues.Length; colIndex++)
            {
                m_columnValues[colIndex] = new StringBuilder(20);
            }

            m_haveHeader = true;
        }

        /// <summary>
        /// Reads column values from given line, so you can access field values by indexer property.
        /// Uses <see cref="ReadField"/> to extract values.
        /// Requires header information to be present, e.g. you must call <see cref="ParseHeader"/> first.
        /// </summary>
        /// <param name="line">Line with column values, may be empty, but cannot be null</param>
        /// <exception cref="InvalidOperationException">You must call <see cref="ParseHeader"/> prior to this operation</exception>
        /// <exception cref="ArgumentNullException"><paramref name="line"/> is null</exception>
        /// <exception cref="CsvReaderException">Line has fewer or more values than declared in the header</exception>
        public void ParseLine(StringBuilder line)
        {
            RequireHeader();

            // to make sure that if we fail in this line, all unread values are empty
            foreach (var valueBuffer in m_columnValues)
            {
                valueBuffer.Clear();
            }

            var offset = 0;
            for (var colIndex = 0; colIndex < m_columnValues.Length; colIndex++)
            {
                var valueBuffer = m_columnValues[colIndex];
                if (!ReadField(line, ref offset, valueBuffer))
                {
                    throw new CsvReaderException("Expected value but reached end of line, column index " + colIndex);
                }
            }

            if (line.Length != offset - 1)
            {
                throw new CsvReaderException("Line has more values than declared in the header");
            }
        }

        /// <summary>
        /// Reads incoming data, splits it into lines, yields every successive line to caller.
        /// Line buffer is supplied by caller.
        /// Does not require header to be present, e.g. you don't have to call <see cref="ParseHeader"/>.
        /// Line breaks (ASCII 0x0A) and carriage returns (ASCII 0x0D) are allowed as part of quoted values.
        /// Nulls (ASCII 0x00) are silently eliminated.
        /// </summary>
        /// <param name="input">Input stream</param>
        /// <param name="lineBuffer">Buffer that receives each successfully read line</param>
        /// <returns>Enumerator of incrementally increasing line numbers</returns>
        
        public IEnumerable<long> ConsumeLines(TextReader input, StringBuilder lineBuffer)
        {
            lineBuffer.Clear();

            var buffer = new char[10000];
            long lineCount = m_haveHeader ? 1 : 0;
            long charCount = 0;

            // parser state variables
            var isQuotedValue = false;
            var isNewValue = true;
            var lastCharWasClosingQuote = false;

            // data reader loop, iterates through all characters
            var read = input.ReadBlock(buffer, 0, buffer.Length);
            while (read > 0)
            {
                if (lineCount == long.MaxValue)
                {
                    throw new Exception("Cannot read more than " + long.MaxValue + " lines");
                }

                charCount += read;
                if (charCount > m_maxCharCount)
                {
                    throw new CsvReaderException("Input stream has more characters than configured max number: " + m_maxCharCount);
                }

                for (var i = 0; i < read; i++)
                {
                    var c = buffer[i];

                    if (c == m_quote)
                    {
                        lineBuffer.Append(m_quote);

                        if (isNewValue)
                        {
                            // if the first character of the value is quote,
                            // we assume that value is quoted and thus has to escape quotes inside
                            isQuotedValue = true;
                            isNewValue = false;
                        }
                        else if (isQuotedValue)
                        {
                            // if we found a quote inside quoted value, have to check for escaping
                            if (lastCharWasClosingQuote)
                            {
                                // previous char was a quote too, so we have a pair now,
                                // which means this is an escaped quote
                                lastCharWasClosingQuote = false;
                            }
                            else
                            {
                                // remember this first quote
                                lastCharWasClosingQuote = true;
                            }
                        }
                    }
                    else if (c == m_delimiter)
                    {
                        if (!isQuotedValue || lastCharWasClosingQuote)
                        {
                            isQuotedValue = false;
                            isNewValue = true;
                            lastCharWasClosingQuote = false;
                        }

                        lineBuffer.Append(m_delimiter);
                    }
                    else if (c == '\n' || c == '\r')
                    {
                        if (!isQuotedValue || lastCharWasClosingQuote)
                        {
                            isQuotedValue = false;
                            isNewValue = true;
                            lastCharWasClosingQuote = false;
                        }

                        if (isQuotedValue)
                        {
                            // inside quoted value, we consume CR/LF chars as part of value
                            lineBuffer.Append(c);
                        }
                        else if (c == '\n')
                        {
                            if (m_haveHeader && lineBuffer.Length == 0)
                            {
                                throw new CsvReaderException(
                                    "Empty content lines are not allowed when header is present, " + lineCount);
                            }

                            if (!m_haveHeader || lineBuffer.Length > 0)
                            {
                                // let consumer know that buffer is ready for processing
                                yield return lineCount;
                            }

                            lineBuffer.Clear();
                            lineCount++;
                        }
                    }
                    else
                    {
                        if (c > '\0')
                        {
                            lineBuffer.Append(c);
                        }

                        isNewValue = false;
                        lastCharWasClosingQuote = false;
                    }

                    if (lineBuffer.Length > m_maxLineLength)
                    {
                        throw new CsvReaderException("Line " + lineCount + " is longer than configured max length: " + m_maxLineLength);
                    }
                }

                read = input.ReadBlock(buffer, 0, buffer.Length);
            }

            if (isQuotedValue && !lastCharWasClosingQuote)
            {
                throw new CsvReaderException("Last line is not terminated properly");
            }

            if (lineBuffer.Length > 0)
            {
                yield return lineCount;
                lineBuffer.Clear();
            }
        }

        /// <summary>
        /// Attempts to read next field value into <paramref name="valueBuffer"/>.
        /// Does not clear previous value of <paramref name="valueBuffer"/>, caller is responsible for this.
        /// Does not require header to be present, e.g. you don't have to call <see cref="ParseHeader"/>.
        /// Line breaks (ASCII 0x0A) and carriage returns (ASCII 0x0D) are allowed as part of quoted values.
        /// Nulls (ASCII 0x00) are silently eliminated.
        /// </summary>
        /// <param name="line">Current line</param>
        /// <param name="offset">Offset in the current line</param>
        /// <param name="valueBuffer">Accepts field value. This buffer must be emptied by caller</param>
        /// <returns>True if any value (including empty) was successfully added into <paramref name="valueBuffer"/></returns>
        /// <exception cref="ArgumentNullException"><paramref name="line"/> is null</exception>
        /// <exception cref="ArgumentNullException"><paramref name="valueBuffer"/> is null</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="offset"/> cannot be negative</exception>
        /// <exception cref="CsvReaderException">Invalid quoting of value</exception>
        public bool ReadField(StringBuilder line, ref int offset, StringBuilder valueBuffer)
        {
            if (offset < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(offset), offset, "Offset cannot be negative");
            }

            if (offset > line.Length)
            {
                return false;
            }

            var end = offset;
            var valueIsQuoted = false;
            var lastCharWasQuote = false;
            var foundTrailingQuote = false;

            var lastChar = '\0';
            for (; end < line.Length; end++)
            {
                lastChar = line[end];

                // If value starts with a double-quote, it will have to end with double-quote.
                // Otherwise, treat double-quote as regular character that does not affect delimiters.

                if (end == offset && lastChar == m_quote)
                {
                    valueIsQuoted = true;
                    continue;
                }

                // is current char an escaped quote? try to consume all those escaped quotes
                if (valueIsQuoted && lastChar == m_quote)
                {
                    if (foundTrailingQuote)
                    {
                        throw new CsvReaderException("Invalid quoting of value at offset " + offset);
                    }

                    var count = 1;
                    for (var i = end + 1; i < line.Length - 1; i++)
                    {
                        if (line[i] != m_quote)
                        {
                            break;
                        }

                        count++;
                    }

                    if (count > 1)
                    {
                        valueBuffer.Append(m_quote, count / 2);
                        end += ((count >> 1) << 1) - 1;
                        lastCharWasQuote = false;
                    }
                    else
                    {
                        lastCharWasQuote = true;
                        foundTrailingQuote = true;
                    }
                }
                else
                {
                    if (lastChar == m_delimiter && (foundTrailingQuote || !valueIsQuoted))
                    {
                        break;
                    }

                    lastCharWasQuote = lastChar == m_quote;

                    if (lastChar != '\0')
                    {
                        valueBuffer.Append(lastChar);
                    }
                }
            }

            // if value starts with a double-quote, it must end with double-quote
            if (valueIsQuoted && !lastCharWasQuote)
            {
                throw new CsvReaderException("Invalid quoting of value at offset " + offset);
            }

            // count current value in even if it is empty - as long as there was a delimiter in front of it
            var result = end > offset || lastChar == m_delimiter || (offset > 0 && line[offset - 1] == m_delimiter);
            offset = end + 1;
            return result;
        }

        private void RequireHeader()
        {
            if (!m_haveHeader)
            {
                throw new InvalidOperationException("To use this operation, ParseHeader must be called at least once");
            }
        }
    }
}
