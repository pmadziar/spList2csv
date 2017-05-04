using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using SP=Microsoft.SharePoint.Client;

namespace spList2csv
{
    class Program
    {
        private const int BatchSize = 1000;
        private const string separator = ";";
        private static readonly Regex rxEoLT = new Regex(@"\t|\n|\r", RegexOptions.Multiline | RegexOptions.Compiled);
        private static readonly Regex rxSpaces = new Regex(@"\s{2,}", RegexOptions.Multiline | RegexOptions.Compiled);
        static void Main(string[] args)
        {
            SP.ClientContext context = null;
            SP.ListItemCollectionPosition position = null;

            if (args.Length != 3)
            {
                Console.BackgroundColor = ConsoleColor.DarkRed;
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Error: wrong number of parameters");
                Console.WriteLine(@"Usage:  spList2csv [SiteUrl] [List Name] [Output File Path]
Example: spList2csv http://team/workgroups/blabla ""My Data List Name"" ""c:\test\data.csv""");
                Console.ResetColor();
                return;
            }

            var path = Path.GetFullPath(args[2]);
            var dir = (new FileInfo(path)).Directory.FullName;

            if (!Directory.Exists(dir))
            {
                Console.BackgroundColor = ConsoleColor.DarkRed;
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(string.Format("Error: folder {0} does not exist.", dir));
                Console.ResetColor();
                return;
            }

            if (File.Exists(path))
            {
                Console.BackgroundColor = ConsoleColor.DarkRed;
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(string.Format("Error: file {0} already exists.", path));
                Console.ResetColor();
                return;
            }

            var weburl = args[0];
            try
            {
                context = new SP.ClientContext(weburl);
                var w = context.Web;
                context.Load(w);
                context.ExecuteQuery();
                var listTitle = args[1];
                var l = w.Lists.GetByTitle(listTitle);
                context.Load(l);
                context.ExecuteQuery();

                List<SP.Field> fields = getFields(context, l);
                var camlXml = getCamlViewXml(fields);

                SP.CamlQuery q = SP.CamlQuery.CreateAllItemsQuery();
                q.ViewXml = camlXml;

                var allItms = new List<SP.ListItem>();
                int cc = 0;
                do
                {
                    var itms = l.GetItems(q);
                    context.Load(itms);
                    context.ExecuteQuery();
                    var count = itms.Count;
                    if (count > 0) allItms.AddRange(itms);
                    position = itms.ListItemCollectionPosition;
                    q.ListItemCollectionPosition = position;
                    cc+=count;
                    Console.WriteLine("Read {0} items from Sharepoint", cc);
                } while (position != null);

                //save data
                saveHeader(path, fields);
                //var line = getLineFromItm(fields, allItms.First());
                var lines = allItms.Select(x => getLineFromItm(fields, x)).ToList();
                File.AppendAllLines(path, lines);

            }
            catch (Exception ex)
            {
                Console.BackgroundColor = ConsoleColor.DarkRed;
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Error: unhandled exeption");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.ResetColor();
                return;
            }
            finally
            {
                if (context != null) context.Dispose();
            }

        }

        private static string getLineFromItm(List<SP.Field> fields, SP.ListItem itm)
        {
            var allValues = new List<string>();
            foreach (var fld in fields)
            {
                allValues.Add(getFieldValueString(itm, fld));
            }
            string line = string.Join(separator, allValues);
            return line;
        }

        private static string getFieldValueString(SP.ListItem itm, SP.Field fld)
        {
            string ret = string.Empty;
            object o = itm[fld.InternalName];
            if (o != null)
            {
                if(o is SP.FieldLookupValue)
                {
                    var fl = (SP.FieldLookupValue)o;
                    ret = fl.LookupValue;
                }
                else if (o is SP.FieldUserValue)
                {
                    var fu = (SP.FieldUserValue)o;
                    ret = fu.LookupValue;
                }
                else if (o is DateTime)
                {
                    var dat = ((DateTime)o);
                    ret = dat.ToLocalTime().ToString("yyyy/MM/dd HH:mm");
                }
                else if (o is string)
                {
                    ret = (string)o;
                }
                else
                {
                    ret = o.ToString();
                }

                if (ret.Contains("\""))
                {
                    ret = ret.Replace("\"", string.Empty);
                }

                if (rxEoLT.IsMatch(ret)) ret = rxEoLT.Replace(ret, " ");
                if (rxSpaces.IsMatch(ret)) ret = rxSpaces.Replace(ret, " ");

                if (ret.Contains(separator))
                {
                    ret = string.Format(@"""{0}""", ret);
                }
            }
            return ret;
        }

        private static void saveHeader(string path, List<SP.Field> fields)
        {
            var fieldNames = fields.Select(x => x.Title).ToArray();
            var str = string.Join(separator, fieldNames);
            var allLines = new string[]{str};
            File.WriteAllLines(path, allLines);
        }



        private static List<SP.Field> getFields(SP.ClientContext context, SP.List l)
        {
            var ret = new List<SP.Field>();

            var olFld1 = getFieldByName(context, l, "Title"); if (olFld1 != null) ret.Add(olFld1);
            var olFld2 = getFieldByName(context, l, "Created By"); if (olFld2 != null) ret.Add(olFld2);
            var olFld3 = getFieldByName(context, l, "Created"); if (olFld3 != null) ret.Add(olFld3);
            var olFld4 = getFieldByName(context, l, "Modified"); if (olFld4 != null) ret.Add(olFld4);
            var olFld5 = getFieldByName(context, l, "Modified By"); if (olFld5 != null) ret.Add(olFld5);

            var fldsq = context.LoadQuery(l.Fields.Where(x => !x.FromBaseType));
            context.ExecuteQuery();
            if (fldsq.Count() > 0) ret.AddRange(fldsq);

            return ret;
        }

        private static SP.Field getFieldByName(SP.ClientContext context, SP.List l, string fldName)
        {
            SP.Field fld = null;
            try
            {
                fld = l.Fields.GetByInternalNameOrTitle(fldName);
                context.Load(fld);
                context.ExecuteQuery();
            }
            catch 
            {
                fld = null;
            }
            return fld;
        }

        private static string getCamlViewXml(List<SP.Field> flds)
        {

            StringBuilder sb = new StringBuilder();
            foreach (var fld in flds)
            {
                sb.AppendLine(string.Format(@"<FieldRef Name=""{0}"" />", fld.InternalName));
            }

            string viewXml = string.Format(@"<View>
    <ViewFields>
{0}
    </ViewFields>
    <RowLimit>{1}</RowLimit>
</View>",
            sb.ToString(), BatchSize);
            return viewXml;
        }
    }
}
