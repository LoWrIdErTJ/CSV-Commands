using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;
using UBotPlugin;
using System.Linq;
using System.Windows;
using System.ServiceModel.Syndication;
using System.Xml;
using System.Xml.Linq;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading;
using System.Security.Cryptography;
using System.Configuration;
using System.Media;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Net;
using System.Management;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Collections;
using System.Collections.Specialized;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.Reflection;
using System.Data.OleDb;

namespace CSVtoHTML
{

    // API KEY HERE
    public class PluginInfo
    {
        public static string HashCode { get { return "dcebc97c4a793e93644e122046e64f9a56599dec"; } }
    }

    // ---------------------------------------------------------------------------------------------------------- //
    //
    // ---------------------------------               COMMANDS               ----------------------------------- //
    //
    // ---------------------------------------------------------------------------------------------------------- //

        
    //
    //
    // CSV FILE TO HTML TABLE
    //
    //
    public class CSVtoHTML : IUBotCommand
    {

        private List<UBotParameterDefinition> _parameters = new List<UBotParameterDefinition>();
        
        public CSVtoHTML()
        {
            _parameters.Add(new UBotParameterDefinition("Path to CSV file", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("CSS style", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("Table style", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("Tr style", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("Td style", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("Border size", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("Return Result", UBotType.UBotVariable));
           
        }

        public string Category
        {
            get { return "Data Commands"; }
        }

        public string CommandName
        {
            get { return "csv to html"; }
        }

        
        public void Execute(IUBotStudio ubotStudio, Dictionary<string,string> parameters)
        {

            string path = parameters["Path to CSV file"];
            string csstyle = parameters["CSS style"];
            string tblstyle = parameters["Table style"];
            string trstyle = parameters["Tr style"];
            string tdstyle = parameters["Td style"];
            string bdrsize = parameters["Border size"];
            string list = parameters["Return Result"];

            const char token = '\t';

            string[] lines = File.ReadAllLines(path);
            StringBuilder result = new StringBuilder();
            result.Append("<html><head>" + csstyle + "<table border ='" + bdrsize + "' " + tblstyle + ">");
            foreach (string line in lines)
            {
                string[] parts = line.Split(token);
                string row = "<tr " + trstyle + "><td " + tdstyle + ">" + string.Join("</td></tr>", parts) + "</td></tr>";
                string v = row.Replace(",", "</td><td " + tdstyle + ">");
                result.Append(v);
            }
            result.Append("</table></body></html>");
            
            ubotStudio.SetVariable(list, result.ToString());
        }

        public bool IsContainer
        {
            get { return false; }
        }

        public IEnumerable<UBotParameterDefinition> ParameterDefinitions
        {
            get { return _parameters; }
        }

        public UBotVersion UBotVersion
        {
            get { return UBotVersion.Standard; }
        }
    }
  
    
    //
    //
    // CSV TEXT TO HTML TABLE
    //
    //
    public class CSVTEXTtoHTML : IUBotCommand
    {

        private List<UBotParameterDefinition> _parameters = new List<UBotParameterDefinition>();

        public CSVTEXTtoHTML()
        {
            _parameters.Add(new UBotParameterDefinition("Comma Delimited Text", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("CSS style", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("Table style", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("Tr style", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("Td style", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("Border size", UBotType.String));
            _parameters.Add(new UBotParameterDefinition("Return Result", UBotType.UBotVariable));

        }

        public string Category
        {
            get { return "Data Commands"; }
        }

        public string CommandName
        {
            get { return "csv text to html"; }
        }


        public void Execute(IUBotStudio ubotStudio, Dictionary<string, string> parameters)
        {

            string textdata = parameters["Comma Delimited Text"];
            string csstyle = parameters["CSS style"];
            string tblstyle = parameters["Table style"];
            string trstyle = parameters["Tr style"];
            string tdstyle = parameters["Td style"];
            string bdrsize = parameters["Border size"];
            string list = parameters["Return Result"];

            const char token = '\n';

            string[] lines =  {textdata};
            StringBuilder result = new StringBuilder();
            result.Append("<html><head>" + csstyle + "<table border ='" + bdrsize + "' " + tblstyle + ">");
            foreach (string line in lines)
            {
                string[] parts = line.Split(token);
                string row = "<tr " + trstyle + "><td " + tdstyle + ">" + string.Join("</td></tr>", parts) + "</td></tr>";
                string u = row.Replace("</tr>", "</tr><td " + tdstyle + ">");
                string v = u.Replace(",", "</td><td " + tdstyle + ">");
                string w = v.Replace("</tr><td " + tdstyle + ">", "</tr><tr " + trstyle + "><td " + tdstyle + ">");
                
                result.Append(w);
            }
            result.Append("</table></body></html>");
            string newlist = result.ToString();
            string x = newlist.Replace("<tr " + trstyle + "><td " + tdstyle + "></table>", "</table>");

            ubotStudio.SetVariable(list, x.ToString());
        }

        public bool IsContainer
        {
            get { return false; }
        }

        public IEnumerable<UBotParameterDefinition> ParameterDefinitions
        {
            get { return _parameters; }
        }

        public UBotVersion UBotVersion
        {
            get { return UBotVersion.Standard; }
        }
    }
  

}
