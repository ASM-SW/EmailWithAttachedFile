using System.Data;
// Copyright © 2016-2019  ASM-SW
//asmeyers@outlook.com  https://github.com/asm-sw

using System.Text;

public static class Extensions
{
    public static string ToCSV(this DataTable table)
    {
        // reference: adapted from  http://stackoverflow.com/questions/888181/convert-datatable-to-csv-stream
        var result = new StringBuilder();
        for (int i = 0; i < table.Columns.Count; i++)
        {
            result.AppendFormat("\"{0}\"",table.Columns[i].ColumnName);
            result.Append(i == table.Columns.Count - 1 ? "\n" : ",");
        }

        foreach (DataRow row in table.Rows)
        {
            for (int i = 0; i < table.Columns.Count; i++)
            {
                result.AppendFormat("\"{0}\"", row[i].ToString());
                result.Append(i == table.Columns.Count - 1 ? "\n" : ",");
            }
        }

        return result.ToString();
    }
}
