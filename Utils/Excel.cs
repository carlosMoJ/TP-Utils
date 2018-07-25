using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Reflection;
using System.Data;
using System.Xml;
using System.Text.RegularExpressions;
using System.IO;

namespace Utils.Excel
{
    /// <summary>
    /// Used to generate an XML excel file, populate and return in various formats.
    /// <para>    See definition or object browser for details on usage</para>
    /// </summary>
    /// <remarks> 
    /// <para>To send to the user as a download use:-</para>
    /// <para>    MemoryStream ms = ExcelObject.getMemoryStream();</para>
    /// <para>    return File(ms, "application/msexcel", "YourFilename.xls");</para>
    /// <para>Or to save to server use:-</para>
    /// <para>    SaveFile("files/00001","Output.xls");</para>
    /// <para>Or just return the xdoc to your calling code for further work with :-</para>
    /// <para>    getXDocument()</para>
    /// <para>This can be used with a Using block:-</para>
    /// <para>    Using(Excel xl = new Excel()) {</para>
    /// <para>        xl.SaveFile("files/00001","Output.xls");</para>
    /// <para>    }</para>
    /// </remarks> 
    public class Excel : IDisposable
    {
        XDocument xdoc;         //Document element
        XElement workbook;      //workbook element
        XElement Styles;        //Styles element
        XElement Names;         //Named Range container
        XNamespace mainNamespace = "urn:schemas-microsoft-com:office:spreadsheet";
        XNamespace o = "urn:schemas-microsoft-com:office:office";
        XNamespace x = "urn:schemas-microsoft-com:office:excel";
        XNamespace ss = "urn:schemas-microsoft-com:office:spreadsheet";
        XNamespace html = "http://www.w3.org/TR/REC-html40";

        public Excel()
        {
            //Create XML from scratch
            xdoc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), new XProcessingInstruction("mso-application", "progid=\"Excel.Sheet\""));
            workbook = new XElement(mainNamespace + "Workbook",
                new XAttribute(XNamespace.Xmlns + "html", html),
                new XAttribute(XName.Get("ss", "http://www.w3.org/2000/xmlns/"), ss),
                new XAttribute(XName.Get("o", "http://www.w3.org/2000/xmlns/"), o),
                new XAttribute(XName.Get("x", "http://www.w3.org/2000/xmlns/"), x),
                new XAttribute(XName.Get("xmlns", ""), mainNamespace)
            );
            //create and add Styles
            Names = new XElement(mainNamespace + "Names");
            Styles = new XElement(mainNamespace + "Styles");
            XElement StyleNorm = new XElement(mainNamespace + "Style", new XAttribute(ss + "ID", "Default"), new XAttribute(ss + "Name", "Normal"));
            XElement StyleDate = new XElement(mainNamespace + "Style", new XAttribute(ss + "ID", "s21"), new XElement(mainNamespace + "NumberFormat", new XAttribute(ss + "Format", "General date")));
            XElement WholeNumFormat = new XElement(mainNamespace + "Style", new XAttribute(ss + "ID", "IntOnly"), new XElement(mainNamespace + "NumberFormat", new XAttribute(ss + "Format", "0")));
            XElement ThreeDecPlace = new XElement(mainNamespace + "Style", new XAttribute(ss + "ID", "ThreeDecPlace"), new XElement(mainNamespace + "NumberFormat", new XAttribute(ss + "Format", "0.000")));
            XElement TwoDecPlace = new XElement(mainNamespace + "Style", new XAttribute(ss + "ID", "TwoDecPlace"), new XElement(mainNamespace + "NumberFormat", new XAttribute(ss + "Format", "0.00")));
            XElement StyleSubheader = new XElement(mainNamespace + "Style", new XAttribute(ss + "ID", "s25"), new XElement(mainNamespace + "Alignment", new XAttribute(ss + "Horizontal", "Center"), new XAttribute(ss + "Vertical", "Bottom")));

            Styles.Add(StyleNorm);
            Styles.Add(StyleDate);
            Styles.Add(WholeNumFormat);
            Styles.Add(ThreeDecPlace);
            Styles.Add(TwoDecPlace);
            Styles.Add(StyleSubheader);

            workbook.Add(Styles);
            workbook.Add(Names);

            xdoc.Add(workbook);
        }

        /// <summary>
        /// Use to add empty sheet 
        /// </summary>
        /// <param name="WorksheetName">Name for the new worksheet, avoid using the [ ] * / \ ? : characters</param>
        public void EmptyDataSheet(string WorksheetName)
        {
            string wrkSht = prepTitle(WorksheetName);
            XElement worksheet = new XElement(mainNamespace + "Worksheet", new XAttribute(ss + "Name", wrkSht));
            workbook.Add(worksheet);
        }
        /// <summary>
        /// Use to add empty sheet with text placeholder
        /// </summary>
        /// <param name="WorksheetName">Name for the new worksheet, avoid using the [ ] * / \ ? : characters</param>
        /// <param name="dataPlaceholder"></param>
        public void EmptyDataSheet(string WorksheetName, string dataPlaceholder)
        {
            string wrkSht = prepTitle(WorksheetName);
            XElement worksheet = new XElement(mainNamespace + "Worksheet", new XAttribute(ss + "Name", wrkSht));
            XElement table = new XElement(mainNamespace + "Table",
                        new XAttribute(ss + "ExpandedColumnCount", 1),
                        new XAttribute(ss + "ExpandedRowCount", 1)
                    ); //close table


            table.Add(new XElement(mainNamespace + "Row", new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", 1), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataPlaceholder))));
            worksheet.Add(table);
            workbook.Add(worksheet);
        }

        /// <summary>
        /// Create a new named worksheet from the collection of objects
        /// </summary>
        /// <typeparam name="T">Enter model/object name</typeparam>
        /// <param name="WorksheetName">Name for the new worksheet, avoid using the [ ] * / \ ? : characters</param>
        /// <param name="listToExport">Collection of generic items</param>
        public void AddDataWorksheet<T>(string WorksheetName, List<T> listToExport, bool withSummary = false)
        {
            string wrkSht = prepTitle(WorksheetName);
            XElement newWorksheet = CreateWorksheet<T>(wrkSht, listToExport, withSummary);
            //add worksheet
            workbook.Add(newWorksheet);
        }

        /// <summary>
        /// Create a new named worksheet from the collection of objects
        /// </summary>
        /// <typeparam name="T">Enter model/object name</typeparam>
        /// <param name="WorksheetName">Name for the new worksheet, avoid using the [ ] * / \ ? : characters</param>
        /// <param name="listToExport">Collection of generic items</param>
        public void AddDataWorksheetSubHeader<T>(string WorksheetName, List<T> listToExport, string[] subheaders, string[] headers)
        {
            string wrkSht = prepTitle(WorksheetName);
            XElement newWorksheet = CreateWorksheet<T>(wrkSht, listToExport, subheaders, headers);
            //add worksheet
            workbook.Add(newWorksheet);
        }

        /// <summary>
        /// Create a new named worksheet from the collection of objects
        /// </summary>
        /// <typeparam name="T">Enter model/object name</typeparam>
        /// <param name="WorksheetName">Name for the new worksheet, avoid using the [ ] * / \ ? : characters</param>
        /// <param name="listToExport">Collection of generic items</param>
        /// <param name="DataRangeName">Name for named range</param>
        public void AddDataWorksheet<T>(string WorksheetName, List<T> listToExport, string DataRangeName)
        {
            AddDataWorksheet<T>(WorksheetName, listToExport, false);

            string nmRng = prepTitle(DataRangeName);
            string wrkSht = prepTitle(WorksheetName);
            //Build named range elements
            XElement range = new XElement(mainNamespace + "NamedRange", new XAttribute(ss + "Name", nmRng), new XAttribute(ss + "RefersTo",
                                                string.Format("='{0}'!R1C1:R{1}C{2}", wrkSht, listToExport.Count() + 1, listToExport[0].GetType().GetProperties().Count())));
            //Add named range
            Names.Add(range);
        }
        /// <summary>
        /// Create a new named worksheet from the collection of objects
        /// </summary>
        /// <typeparam name="T">Enter model/object name</typeparam>
        /// <param name="WorksheetName">Name for the new worksheet, avoid using the [ ] * / \ ? : characters</param>
        /// <param name="listToExport">Collection of generic items</param>
        /// <param name="DataRangeName">Name for named range</param>
        public void AddDataWorksheet<T>(string WorksheetName, List<T> listToExport, string DataRangeName, bool withSummary)
        {
            AddDataWorksheet<T>(WorksheetName, listToExport, withSummary);

            string nmRng = prepTitle(DataRangeName);
            string wrkSht = prepTitle(WorksheetName);
            //Build named range elements
            XElement range = new XElement(mainNamespace + "NamedRange", new XAttribute(ss + "Name", nmRng), new XAttribute(ss + "RefersTo",
                                                string.Format("='{0}'!R1C1:R{1}C{2}", wrkSht, listToExport.Count() + 1, listToExport[0].GetType().GetProperties().Count())));
            //Add named range
            Names.Add(range);
        }
        ///// <summary>
        ///// Create a new named worksheet from the collection of Correspondence items
        ///// </summary>
        ///// <param name="WorksheetName">Name for the new worksheet, avoid using the [ ] * / \ ? : characters</param>
        ///// <param name="listToExport">Collection of Correspondence items</param>
        //public void AddDataWorksheet(string WorksheetName, List<Correspondence> listToExport)
        //{
        //    string wrkSht = prepTitle(WorksheetName);
        //    XElement newWorksheet = CreateWorksheet(wrkSht, listToExport);
        //    //add worksheet
        //    workbook.Add(newWorksheet);
        //}
        ///// <summary>
        ///// Create a new named worksheet from the collection of Correspondence items with a named range that encompasses all data
        ///// </summary>
        ///// <param name="WorksheetName">Name for the new worksheet, avoid using the [ ] * / \ ? : characters</param>
        ///// <param name="listToExport">Collection of Correspondence items</param>
        ///// <param name="DataRangeName">Name for named range</param>
        //public void AddDataWorksheet(string WorksheetName, List<Correspondence> listToExport, string DataRangeName)
        //{
        //    AddDataWorksheet(WorksheetName, listToExport);
        //    string wrkSht = prepTitle(WorksheetName);
        //    string nmRng = prepTitle(DataRangeName);
        //    //Build named range elements
        //    XElement range = new XElement(mainNamespace + "NamedRange", new XAttribute(ss + "Name", nmRng), new XAttribute(ss + "RefersTo",
        //                                        string.Format("='{0}'!R1C1:R{1}C{2}", wrkSht, listToExport.Count() + 1, listToExport[0].GetType().GetProperties().Count())));
        //    //Add named range
        //    Names.Add(range);
        //}
        ///// <summary>
        ///// Create the Workbook XElement
        ///// </summary>
        ///// <param name="WorksheetName">Name for the new worksheet, avoid using the [ ] * / \ ? : characters</param>
        ///// <param name="listToExport">Collection of Correspondence items</param>
        ///// <returns></returns>
        //private XElement CreateWorksheet(string WorksheetName, List<Correspondence> listToExport)
        //{
        //    List<XElement> rows = new List<XElement>();
        //    int maxCol = 19;
        //    //loop here for cols
        //    var HdrRow = new XElement(mainNamespace + "Row");
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Reference")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Their Ref")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Date Entered")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Item dated")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Correspondent")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "MP")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Subject")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Addressee")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Signatory")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Type")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Status")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Stage")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Rag Status")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Allocated to user")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Allocated to team")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Stage Target")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Item Target")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Closed Within Time")));
        //    HdrRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Days taken")));
        //    rows.Add(HdrRow);

        //    //and rows
        //    foreach (var dataItem in listToExport)
        //    {

        //        var dataRow = new XElement(mainNamespace + "Row");
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.Reference)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.TheirRef)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "StyleID", "s21"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Index", "3"), new XAttribute(ss + "Type", "DateTime"), ((DateTime)dataItem.DateRecorded).ToUniversalTime())));
        //        if (dataItem.ItemDated != null)
        //        {
        //            dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "StyleID", "s21"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "DateTime"), ((DateTime)dataItem.ItemDated).ToUniversalTime())));
        //        }
        //        int col = 5;
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", col), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.CorrespondenceLink == null ? "" : dataItem.CorrespondenceLink.Correspondent == null ? "" : dataItem.CorrespondenceLink.Correspondent.DisplayName)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.CorrespondenceLink == null ? "" : dataItem.CorrespondenceLink.MP == null ? "" : dataItem.CorrespondenceLink.MP.DisplayName)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.Subject)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.MoJMinisterID == null ? "" : dataItem.MoJMinister.DisplayName)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.SignatoryID == null ? "" : dataItem.Signatory.DisplayName)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.CorrespondenceType.Name)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.CurrentDisplayStatus)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.CurrentAllocation == null ? "" : dataItem.CurrentAllocation.Stage.Name)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.CurrentAllocation == null ? "" : dataItem.CurrentAllocation.CurrentState)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.CurrentAllocation == null || dataItem.CurrentAllocation.AllocatedToUserID == null ? "" : dataItem.CurrentAllocation.AllocatedToUser.DisplayName)));
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), dataItem.CurrentAllocation == null ? "" : dataItem.CurrentAllocation.AllocatedToTeam.Name)));
        //        if (dataItem.CurrentAllocation.TargetDate != null)
        //        {
        //            dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XAttribute(ss + "StyleID", "s21"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "DateTime"), dataItem.CurrentAllocation.TargetDate.ToUniversalTime())));
        //        }
        //        if (dataItem.TargetDateAtIssue != null)
        //        {
        //            dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", ++col), new XAttribute(ss + "StyleID", "s21"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "DateTime"), dataItem.TargetDateAtIssue)));
        //        }
        //        string ClosedWithinTime = string.Empty;
        //        if (dataItem.CorrespondenceType.TargetDays != null && dataItem.CorrespondenceType.TargetDays > 0)
        //        {
        //            if (dataItem.CurrentStatus.Name == "Closed")
        //            {
        //                if (dataItem.DaysToComplete <= dataItem.CorrespondenceType.TargetDays)
        //                {
        //                    ClosedWithinTime = "Yes";
        //                }
        //                else
        //                {
        //                    ClosedWithinTime = "No";
        //                }
        //            }

        //        }
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", maxCol - 1), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), ClosedWithinTime))); //Closed within time
        //        dataRow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", maxCol), new XAttribute(ss + "StyleID", "IntOnly"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "Number"), dataItem.DaysSinceIssue)));
        //        rows.Add(dataRow);
        //    }

        //    XElement table = new XElement(mainNamespace + "Table",
        //                new XAttribute(ss + "ExpandedColumnCount", maxCol),
        //                new XAttribute(ss + "ExpandedRowCount", rows.Count)
        //            ); //close table
        //    foreach (var _row in rows)
        //    {
        //        var cols = _row.Descendants(mainNamespace + "Data").Count();
        //        var title = prepTitle(_row.Descendants(mainNamespace + "Data").First().Value);
        //        table.Add(_row);
        //    }
        //    // create Worksheet
        //    XElement worksheet = new XElement(mainNamespace + "Worksheet", new XAttribute(ss + "Name", WorksheetName));
        //    worksheet.Add(table);
        //    return worksheet;
        //}
        /// <summary>
        /// Create a Workbook Element from a generic list of objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="WorksheetName">Name for the new worksheet, avoid using the [ ] * / \ ? : characters</param>
        /// <param name="listToExport">Generic Collection</param>
        /// <returns></returns>
        private XElement CreateWorksheet<T>(string WorksheetName, List<T> listToExport)
        {
            List<XElement> rows = new List<XElement>();

            //loop here for cols
            PropertyInfo[] fieldInfo = listToExport[0].GetType().GetProperties();
            var row = new XElement(mainNamespace + "Row");
            foreach (PropertyInfo col in fieldInfo)
            {
                if (col.PropertyType != typeof(EntityKey) && col.PropertyType != typeof(EntityState))
                {
                    var cell = new XElement(mainNamespace + "Cell",
                         new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), col.Name));
                    row.Add(cell);
                }
            }
            rows.Add(row);
            //and rows
            foreach (T dataItem in listToExport)
            {
                var dataRow = new XElement(mainNamespace + "Row");
                PropertyInfo[] allProperties = dataItem.GetType().GetProperties();
                int column = 1;
                foreach (PropertyInfo thisProperty in allProperties)
                {
                    if (thisProperty.PropertyType != typeof(EntityKey) && thisProperty.PropertyType != typeof(EntityKey))
                    {
                        var cell = new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", column++));
                        var type = Nullable.GetUnderlyingType(thisProperty.PropertyType) ?? thisProperty.PropertyType;
                        object value = thisProperty.GetValue(dataItem, null);

                        switch (Type.GetTypeCode(type))
                        {
                            case TypeCode.Boolean:
                                cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), value == null ? String.Empty : value.ToString()));
                                break;
                            case TypeCode.DateTime:
                                if (value != null)
                                {
                                    cell.Add(new XAttribute(ss + "StyleID", "s21"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "DateTime"), value));
                                }
                                else
                                {
                                    cell = null;
                                }
                                break;
                            case TypeCode.Single:
                            case TypeCode.Decimal:
                            case TypeCode.Double:
                                cell.Add(new XAttribute(ss + "StyleID", "ThreeDecPlace"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "Number"), value));
                                break;
                            case TypeCode.DBNull:
                            case TypeCode.Empty:
                                cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), String.Empty));
                                break;
                            case TypeCode.Int16:
                            case TypeCode.Int32:
                            case TypeCode.Int64:
                            case TypeCode.UInt16:
                            case TypeCode.UInt32:
                            case TypeCode.UInt64:
                                if (value != null)
                                {
                                    cell.Add(new XAttribute(ss + "StyleID", "IntOnly"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "Number"), value));
                                }
                                else
                                {
                                    cell = null;
                                }
                                break;
                            case TypeCode.Object:
                                string subItemName = thisProperty.Name;
                                string result = GetDataFromObject(dataItem, subItemName, "Acronym", "Detail", "Name");

                                if (result != null)
                                {
                                    cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), value == null ? String.Empty : result.ToString()));
                                }
                                else
                                {
                                    cell = null;
                                }
                                break;
                            case TypeCode.String:
                            default:
                                cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), value == null ? String.Empty : value.ToString()));
                                break;
                        }
                        if (cell != null)
                        {
                            dataRow.Add(cell);
                        }
                    }
                }
                rows.Add(dataRow);
            }
            XElement table = new XElement(mainNamespace + "Table",
                        new XAttribute(ss + "ExpandedColumnCount", fieldInfo.Count()),
                        new XAttribute(ss + "ExpandedRowCount", rows.Count)
                    ); //close table
            foreach (var _row in rows)
            {
                var cols = _row.Descendants(mainNamespace + "Data").Count();
                var title = prepTitle(_row.Descendants(mainNamespace + "Data").First().Value);
                table.Add(_row);
            }
            // create Worksheet
            XElement worksheet = new XElement(mainNamespace + "Worksheet", new XAttribute(ss + "Name", WorksheetName));
            worksheet.Add(table);
            return worksheet;
        }

        /// <summary>
        /// Create a Workbook Element from a generic list of objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="WorksheetName">Name for the new worksheet, avoid using the [ ] * / \ ? : characters</param>
        /// <param name="listToExport">Generic Collection</param>
        /// <param name="withSummary">Boolean to add summary formulae</param>
        /// 
        /// <returns></returns>
        private XElement CreateWorksheet<T>(string WorksheetName, List<T> listToExport, bool withSummary)
        {
            List<XElement> rows = new List<XElement>();

            //loop here for cols
            PropertyInfo[] fieldInfo = listToExport[0].GetType().GetProperties();
            var row = new XElement(mainNamespace + "Row");
            foreach (PropertyInfo col in fieldInfo)
            {
                if (col.PropertyType != typeof(EntityKey) && col.PropertyType != typeof(EntityState))
                {
                    var cell = new XElement(mainNamespace + "Cell",
                         new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), col.Name));
                    row.Add(cell);
                }
            }
            rows.Add(row);
            //and rows
            foreach (T dataItem in listToExport)
            {
                var dataRow = new XElement(mainNamespace + "Row");
                PropertyInfo[] allProperties = dataItem.GetType().GetProperties();
                int column = 1;
                foreach (PropertyInfo thisProperty in allProperties)
                {
                    if (thisProperty.PropertyType != typeof(EntityKey) && thisProperty.PropertyType != typeof(EntityKey))
                    {
                        var cell = new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", column++));
                        var type = Nullable.GetUnderlyingType(thisProperty.PropertyType) ?? thisProperty.PropertyType;
                        object value = thisProperty.GetValue(dataItem, null);

                        switch (Type.GetTypeCode(type))
                        {
                            case TypeCode.Boolean:
                                cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), value == null ? String.Empty : value.ToString()));
                                break;
                            case TypeCode.DateTime:
                                if (value != null)
                                {
                                    cell.Add(new XAttribute(ss + "StyleID", "s21"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "DateTime"), value));
                                }
                                else
                                {
                                    cell = null;
                                }
                                break;
                            case TypeCode.Single:
                            case TypeCode.Decimal:
                            case TypeCode.Double:
                                cell.Add(new XAttribute(ss + "StyleID", "ThreeDecPlace"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "Number"), value));
                                break;
                            case TypeCode.DBNull:
                            case TypeCode.Empty:
                                cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), String.Empty));
                                break;
                            case TypeCode.Int16:
                            case TypeCode.Int32:
                            case TypeCode.Int64:
                            case TypeCode.UInt16:
                            case TypeCode.UInt32:
                            case TypeCode.UInt64:
                                if (value != null)
                                {
                                    cell.Add(new XAttribute(ss + "StyleID", "IntOnly"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "Number"), value));
                                }
                                else
                                {
                                    cell = null;
                                }
                                break;
                            case TypeCode.Object:
                                string subItemName = thisProperty.Name;
                                string result = GetDataFromObject(dataItem, subItemName, "Acronym", "Detail", "Name");

                                if (result != null)
                                {
                                    cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), value == null ? String.Empty : result.ToString()));
                                }
                                else
                                {
                                    cell = null;
                                }
                                break;
                            case TypeCode.String:
                            default:
                                cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), value == null ? String.Empty : value.ToString()));
                                break;
                        }
                        if (cell != null)
                        {
                            dataRow.Add(cell);
                        }
                    }
                }
                rows.Add(dataRow);
            }
            if (withSummary)
            {
                int rCount = listToExport.Count();
                var summrow = new XElement(mainNamespace + "Row");
                summrow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", 1), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Records")));
                summrow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", 2), new XAttribute(ss + "Formula", string.Format("=COUNTIF(R[-{0}]C[-1]:R[-1]C[-1],\"<>\"\"\")", rCount)), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "Number"), 0)));
                summrow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", 20), new XAttribute(ss + "Formula", string.Format("=COUNTIF(R[-{0}]C:R[-1]C,\"<=10\")", rCount)), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "Number"), 0)));
                summrow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", 23), new XAttribute(ss + "Formula", string.Format("=COUNTIF(R[-{0}]C:R[-1]C,\"<=15\")", rCount)), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "Number"), 0)));
                rows.Add(summrow);
                summrow = new XElement(mainNamespace + "Row");
                summrow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", 1), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), "Unique")));
                summrow.Add(new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", 2), new XAttribute(ss + "Formula", string.Format("=SUM(IF(FREQUENCY(MATCH(R[-{0}]C[-1]:R[-2]C[-1],R[-{0}]C[-1]:R[-2]C[-1],0),MATCH(R[-{0}]C[-1]:R[-2]C[-1],R[-{0}]C[-1]:R[-2]C[-1],0))>0,1))", rCount + 1)), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "Number"), 0)));
                rows.Add(summrow);
            }
            XElement table = new XElement(mainNamespace + "Table",
                        new XAttribute(ss + "ExpandedColumnCount", fieldInfo.Count()),
                        new XAttribute(ss + "ExpandedRowCount", rows.Count)
                    ); //close table
            foreach (var _row in rows)
            {
                var cols = _row.Descendants(mainNamespace + "Data").Count();
                var title = prepTitle(_row.Descendants(mainNamespace + "Data").First().Value);
                table.Add(_row);
            }
            // create Worksheet
            XElement worksheet = new XElement(mainNamespace + "Worksheet", new XAttribute(ss + "Name", WorksheetName));
            worksheet.Add(table);
            return worksheet;
        }

        private XElement CreateWorksheet<T>(string WorksheetName, List<T> listToExport, string[] subheaders, string[] headers)
        {
            List<XElement> rows = new List<XElement>();

            ////loop here for cols
            //PropertyInfo[] fieldInfo = listToExport[0].GetType().GetProperties();
            //var row = new XElement(mainNamespace + "Row");
            //foreach (PropertyInfo col in fieldInfo)
            //{
            //    if (col.PropertyType != typeof(EntityKey) && col.PropertyType != typeof(EntityState))
            //    {
            //        var cell = new XElement(mainNamespace + "Cell",
            //             new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), col.Name));
            //        row.Add(cell);
            //    }
            //}
            //rows.Add(row);
            //Add headers
            //loop here for cols
            var row = new XElement(mainNamespace + "Row");
            foreach (string header in headers)
            {
                var cell = new XElement(mainNamespace + "Cell",
                    new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), header));
                row.Add(cell);
            }
            rows.Add(row);
            //Add subheader
            row = new XElement(mainNamespace + "Row");
            foreach (string subheader in subheaders)
            {
                var cell = new XElement(mainNamespace + "Cell",
                    new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), subheader));
                row.Add(cell);
            }
            rows.Add(row);
            //END Add Subheader
            //and rows
            foreach (T dataItem in listToExport)
            {
                var dataRow = new XElement(mainNamespace + "Row");
                PropertyInfo[] allProperties = dataItem.GetType().GetProperties();
                int column = 1;
                foreach (PropertyInfo thisProperty in allProperties)
                {
                    if (thisProperty.PropertyType != typeof(EntityKey) && thisProperty.PropertyType != typeof(EntityKey))
                    {
                        var cell = new XElement(mainNamespace + "Cell", new XAttribute(ss + "Index", column++));
                        var type = Nullable.GetUnderlyingType(thisProperty.PropertyType) ?? thisProperty.PropertyType;
                        object value = thisProperty.GetValue(dataItem, null);

                        switch (Type.GetTypeCode(type))
                        {
                            case TypeCode.Boolean:
                                cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), value == null ? String.Empty : value.ToString()));
                                break;
                            case TypeCode.DateTime:
                                if (value != null)
                                {
                                    cell.Add(new XAttribute(ss + "StyleID", "s21"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "DateTime"), value));
                                }
                                else
                                {
                                    cell = null;
                                }
                                break;
                            case TypeCode.Single:
                            case TypeCode.Decimal:
                            case TypeCode.Double:
                                cell.Add(new XAttribute(ss + "StyleID", "TwoDecPlace"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "Number"), value));
                                break;
                            case TypeCode.DBNull:
                            case TypeCode.Empty:
                                cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), String.Empty));
                                break;
                            case TypeCode.Int16:
                            case TypeCode.Int32:
                            case TypeCode.Int64:
                            case TypeCode.UInt16:
                            case TypeCode.UInt32:
                            case TypeCode.UInt64:
                                if (value != null)
                                {
                                    cell.Add(new XAttribute(ss + "StyleID", "IntOnly"), new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "Number"), value));
                                }
                                else
                                {
                                    cell = null;
                                }
                                break;
                            case TypeCode.Object:
                                string subItemName = thisProperty.Name;
                                string result = GetDataFromObject(dataItem, subItemName, "Acronym", "Detail", "Name");

                                if (result != null)
                                {
                                    cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), value == null ? String.Empty : result.ToString()));
                                }
                                else
                                {
                                    cell = null;
                                }
                                break;
                            case TypeCode.String:
                            default:
                                cell.Add(new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), value == null ? String.Empty : value.ToString()));
                                break;
                        }
                        if (cell != null)
                        {
                            dataRow.Add(cell);
                        }
                    }
                }
                rows.Add(dataRow);
            }
            XElement table = new XElement(mainNamespace + "Table",
                        new XAttribute(ss + "ExpandedColumnCount", headers.Length),
                        new XAttribute(ss + "ExpandedRowCount", rows.Count)
                    ); //close table
            foreach (var _row in rows)
            {
                var cols = _row.Descendants(mainNamespace + "Data").Count();
                var title = prepTitle(_row.Descendants(mainNamespace + "Data").First().Value);
                table.Add(_row);
            }
            // create Worksheet
            XElement worksheet = new XElement(mainNamespace + "Worksheet", new XAttribute(ss + "Name", WorksheetName));
            worksheet.Add(table);
            return worksheet;
        }


        public void HideWorksheet(string WorksheetName)
        {

            IEnumerable<XElement> users = xdoc.Root
                                          .Elements("Worksheet")
                                          .Where(el => (string)el.Attribute("Name") == WorksheetName);

            XElement WrkSht = xdoc.Descendants(mainNamespace + "Worksheet").Single(x => x.Name == WorksheetName);
            if (WrkSht != null)
            {
                WrkSht.Add(new XElement(mainNamespace + "WorksheetOptions", new XElement(ss + "Visible", "SheetHidden")));
            }
        }


        /// <summary>
        /// Returns a MemoryStream of the XDocument for sending to user as download
        /// </summary>
        /// <returns></returns>
        public MemoryStream getMemoryStream()
        {
            //Compress the XML DATA
            MemoryStream memoryStream = new MemoryStream();
            XmlWriter xmlWriter = XmlWriter.Create(memoryStream);

            //Save data to memoryStream
            xdoc.Save(xmlWriter);

            //writer Close
            xmlWriter.Close();

            //Reset Memorystream postion to 0
            memoryStream.Position = 0;
            return memoryStream;
        }
        /// <summary>
        /// Returns the XDocument
        /// </summary>
        /// <returns></returns>
        public XDocument getXDocument()
        {
            return xdoc;
        }
        /// <summary>
        /// Save to the server
        /// </summary>
        /// <param name="folder">Folder(s) to save, do not include preceeding or trailing slashes (e.g. files/newfiles)</param>
        /// <param name="SaveAsName">The name of the file, you should include the extension (.xls or .xml)</param>
        public void SaveFile(string folder, string SaveAsName)
        {
            string fullname = Path.Combine(string.Format("~/{0}", folder), SaveAsName);

            xdoc.Save(fullname);
        }
        void IDisposable.Dispose()
        {
            //do nothing - does this aid garbage collection?
        }
        private string prepTitle(string value)
        {
            string res = value;
            string pattern = @"[\'\<\>\*\\\/\?|]";
            Match m = Regex.Match(res, pattern);
            bool nameIsValid = (m.Success || (string.IsNullOrEmpty(value)) || (value.Length > 31)) ? false : true;
            if (!nameIsValid)
            {
                res = Regex.Replace(res, pattern, "");
            }
            return res;
        }
        private string GetRoleName(object myObject, string Detail)
        {
            Type objectType = myObject.GetType();
            object internalObject = objectType.GetProperty("Role").GetValue(myObject, null);

            Type internalType = internalObject.GetType();
            PropertyInfo singleProperty = internalType.GetProperty("Detail");

            return singleProperty.GetValue(internalObject, null).ToString();
        }
        private string GetDataFromObject(object myObject, string subItemName, params string[] fieldName)
        {
            Type objectType = myObject.GetType();
            object internalObject = objectType.GetProperty(subItemName).GetValue(myObject, null);

            if (internalObject != null)
            {
                Type internalType = internalObject.GetType();

                foreach (var item in fieldName)
                {
                    try
                    {
                        PropertyInfo singleProperty = internalType.GetProperty(item);
                        return singleProperty.GetValue(internalObject, null).ToString();
                    }
                    catch
                    { //do nothing - continue the loop 
                    }
                }
            }
            return null;
        }
    }
}
