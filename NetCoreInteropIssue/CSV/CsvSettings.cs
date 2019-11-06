﻿using System.Text;

namespace Microsoft.Csv
{
    public struct CsvSettings
    {
        public static CsvSettings Default = new CsvSettings(
            encoding: Encoding.UTF8,
            delimiter: ',',
            textQualifier: '"'
        );

        public CsvSettings(Encoding encoding, char delimiter, char textQualifier)
            : this()
        {
            Encoding = encoding;
            Delimiter = delimiter;
            TextQualifier = textQualifier;
        }

        public Encoding Encoding { get; private set; }
        public char Delimiter { get; private set; }
        public char TextQualifier { get; private set; }

        public bool IsValid => Encoding != null;
    }
}