using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Text.RegularExpressions;

namespace ExcelUplodaer
{
    public partial class uploadExcel : System.Web.UI.Page
    {

        protected static string strFileName = string.Empty;
        System.IO.Compression.
        ZipArchive zip;
        protected void Page_Load(object sender, EventArgs e)
        {
           
        }

        protected void btnOpenFile_Click(object sender, EventArgs e)
        {
            openExcelFile();
        }
        private void openExcelFile()
        {
            ZipArchiveEntry workbook = null;

            openZipFile();

            if (locateFileInZip("xl/workbook.xml", ref workbook))
            {
                getSheets(workbook);
            }
            else
            {

            }
        }

        private void openZipFile()
        {
            strFileName = fileUpload.PostedFile.FileName;



            if (strFileName.Trim().Length == 0)
            {
                showMessage(1, "file is not Available");
            }
            else
            {
                string tmpstr = strFileName.Substring(strFileName.Length - 4);
                tmpstr = tmpstr.ToUpper();
                if ((tmpstr == ".XLS") || (tmpstr == "XLSX"))
                {
                    zip = ZipFile.OpenRead(strFileName);
                    Session["zipFile"] = zip;
                }
                else
                {
                    showMessage(1, "the selected file is not in the required format.");
                }
            }
        }

        private bool locateFileInZip(string filename, ref ZipArchiveEntry fileInZip)
        {
            bool retVal = false;

            if (zip == null)
            {
                //openZipFile();
                zip = (ZipArchive)Session["zipFile"];
            }

            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                if (entry.FullName == filename)
                {
                    fileInZip = entry;
                    retVal = true;
                    break;
                }
            }


            return retVal;
        }

        private void getSheets(ZipArchiveEntry entry)
        {
            StreamReader reader = new StreamReader(entry.Open());
            string sName;
            
            HiddenField2.Value = reader.ReadToEnd();


            XmlDocument xmldoc = new XmlDocument();
            XmlNode node;

            xmldoc.LoadXml(HiddenField2.Value);
            node = xmldoc.DocumentElement.GetElementsByTagName("sheets")[0];

            if (node.ChildNodes.Count > 1)
            {
                showSheets(node);
            }
            else
            {
                sName = "sheet" + node.ChildNodes[0].Attributes["sheetId"].Value;
                showData(sName);
            }
            
        }

        private void showSheets(XmlNode sheetsNode)
        {
            int i;

            ddlSheetList.Items.Clear();

            ddlSheetList.Items.Add(new ListItem("Choose sheet", "null"));

            for (i = 0; i < sheetsNode.ChildNodes.Count; i++)
            {
                ddlSheetList.Items.Add(new ListItem(sheetsNode.ChildNodes[i].Attributes["name"].Value,
                                                                "sheet" + sheetsNode.ChildNodes[i].Attributes["sheetId"].Value));
            }
        }


        private XmlDocument openXmlFileInZip(string fileName)
        {
            XmlDocument retXml = new XmlDocument();
            ZipArchiveEntry locatedFile = null;

            if (locateFileInZip(fileName, ref locatedFile))
            {
                StreamReader reader = new StreamReader(locatedFile.Open());
                retXml.LoadXml(reader.ReadToEnd());
            }
            else
            {
                retXml.LoadXml("<xml>null</xml>");
            }

           return retXml;
        }


        private void showData(string sheetName)
        {
            //XmlDocument xmldata = new XmlDocument();
            XmlDocument xmlSheetContent = new XmlDocument();
            XmlNode node;
            string strDimension;
            string[] sDim;
            string sUppA, sUppD, sLowA, sLowD;
            string tmpStr;
            StringBuilder sb = new StringBuilder();

            sUppA = string.Empty;
            sUppD = string.Empty;
            sLowA = string.Empty;
            sLowD = string.Empty;


            xmlSheetContent = openXmlFileInZip("xl/worksheets/" + sheetName + ".xml");
            node = xmlSheetContent.DocumentElement.GetElementsByTagName("dimension")[0];
            strDimension = node.Attributes["ref"].Value;
            if (strDimension.IndexOf(':') > 0)
            {
                sDim = strDimension.Split(':');
                tmpStr = separateCharsaAndDigits(sDim[0], ref sUppA, ref sUppD);
                tmpStr = separateCharsaAndDigits(sDim[1], ref sLowA, ref sLowD);

                hfRange.Value = sUppA + "," + sUppD + "," + sLowA + "," + sLowD;

            }
            else // Just one cell has data.
            {
                tmpStr = separateCharsaAndDigits(strDimension, ref sLowA, ref sLowD);

                hfRange.Value = "0,0," + sLowA + "," + sLowD;
            }


            hfContent.Value = xmlSheetContent.InnerXml;
            hfCommonValues.Value = openXmlFileInZip("xl/sharedStrings.xml").InnerXml;

            ClientScript.RegisterStartupScript(GetType(), "", "showExcel();", true);
        }


        private string separateCharsaAndDigits(string sSrc, ref string sDigits, ref string sAlphas)
        {
            string retStr = string.Empty;

            Regex re = new Regex(@"([a-zA-Z]+)(\d+)");
            Match result = re.Match(sSrc);

            sAlphas = alpha2Number(result.Groups[1].Value);
            sDigits = result.Groups[2].Value;

            return retStr;
        }



        private string alpha2Number(string sSrc)
        {
            string retStr = string.Empty;
            int iSum = 0;
            string baseStr = "_ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            if (sSrc.Length > 2)
            {
                throw new Exception("Chart in to big!");
            }
            else
            {
                if (sSrc.Length == 1)
                {
                    retStr = baseStr.IndexOf(sSrc).ToString();
                }
                else
                {
                    iSum = baseStr.IndexOf(sSrc[0]) + baseStr.IndexOf(sSrc[1]) + 26; //  Column CF -> C(3) + F(6). But CF comes after the all alphabeit
                                                                                     // so we must add 26 = 3 + 6 + 26 =  Col. 35. 
                    retStr = iSum.ToString();
                }
            }
            return retStr;
        }



        private void showMessage(int iType, string sMessage)
        {

        }

        protected void ddlSheetList_SelectedIndexChanged(object sender, EventArgs e)
        {
            showData(ddlSheetList.SelectedValue);
        }

    }
}