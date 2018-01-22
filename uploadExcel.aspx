<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="uploadExcel.aspx.cs" Inherits="ExcelUplodaer.uploadExcel" ValidateRequest="false"%>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="w3.css" rel="stylesheet" />
</head>

    <script>
        function readExcel() {
            try {
                excel = new ActiveXObject("Excel.Application");
            }
            catch (e) {
                alert("ActiveX object error. Please check your browser security settings.");
                return;
            }

            var fileName = getFileName();
            if(fileName == 0)
            {
                return;
            }
            else {
                
            }

        }



        function getFileName() {
            var obj = document.getElementById("File1");
            var retVal;

            retVal = obj.value;

            if (retVal.length <= 0) {
                alert("Please select file.");
                return 0;
            }
            else {
                return retVal;
            }
        }


        function copyPaste() {
            document.getElementById("div1").innerText = document.getElementById("text1").value;
        }



        var xmlCommon;

        function showExcel() {
            var xmlDoc;
            var xmlNoe;
            var i, j;


            parser = new DOMParser();
            xmlDoc = parser.parseFromString(document.getElementById("hfContent").value, "text/xml");

            parser = new DOMParser();
            xmlCommon = parser.parseFromString(document.getElementById("hfCommonValues").value, "text/xml");

            xmlElement = xmlDoc.documentElement.getElementsByTagName("sheetData")[0];

            var dim = document.getElementById("hfRange").value.split(",");

            var rows = dim[2];
            var cols = dim[3];

            createTable(rows, cols);

            var rowsElement = xmlElement.childNodes;
            var cells;
            var cellId;
            var cellType;

            for (i = 0; i < rowsElement.length; i++) {
                cells = rowsElement[i].getElementsByTagName("c");
                for (j = 0; j < cells.length; j++) {

                    cellId = cells[j].attributes["r"].value;

                    if (!cells[j].hasAttribute("t")) {
                        if (!cells[j].hasAttribute("s")) { // That includes the <f> tag (formula).
                            if (cells[j].getElementsByTagName("v").length > 0) { //TEXT
                                document.getElementById(cellId).value = cells[j].getElementsByTagName("v")[0].textContent;
                            }

                        }
                        else {
                            if (cells[j].attributes["s"].value == "1") // TIME
                            {
                                document.getElementById(cellId).value = getRealdate(cells[j].getElementsByTagName("v")[0].textContent);
                            }

                            if (cells[j].attributes["s"].value == "2") // currency
                            {
                                document.getElementById(cellId).value = cells[j].getElementsByTagName("v")[0].textContent;
                            }

                            if (cells[j].attributes["s"].value == "4") // number
                            {
                                document.getElementById(cellId).value = cells[j].getElementsByTagName("v")[0].textContent;
                            }
                        }
                    }
                    else {

                        switch (cells[j].attributes["t"].value) {
                            case "str":
                                document.getElementById(cellId).value = cells[j].getElementsByTagName("v")[0].textContent;
                                break;
                            case "s":
                                document.getElementById(cellId).value = getCommonValue(cells[j].getElementsByTagName("v")[0].textContent);
                                break;
                            default:
                                document.getElementById(cellId).value = "";
                                break;
                        }

                    }
                }
            }
        }



        function getRealdate(timeNumber) {
            var retVal;
            var retDate;
            var oddDays;
            var date;
            var excelDate;
            var jstime;


            excelDate = timeNumber.split(".")[0] - 25555; //Days between 1/1/1900 - 1/1/1970. So new excel time= js time;

            jstime = excelDate * 24 * 60 * 1000; //change excel number of year to JS number of miliseconeds.



            date = new Date((timeNumber - (25567 + 1)) * 86400 * 1000);

            retVal = date.getDate() + "-" + date.getMonth() + "-" + date.getFullYear();


            return retVal;
        }



        function getCommonValue(enrtyNum) {
            var retVal;
            var entries = xmlCommon.getElementsByTagName("si");

            retVal = entries[enrtyNum].getElementsByTagName("t")[0].textContent;

            return retVal;
        }

        function createTable(rows, cols) {
            var i, j;
            var htmlStr;


            htmlStr = "<table  class=\"w3-container w3-card-4 w3-light-grey\">";

            for (i = 1; i <= rows; i++) {
                htmlStr = htmlStr + "<tr>";
                for (j = 1; j <= cols; j++) {
                    htmlStr = htmlStr + '<td><input type="text" id="' + colsToAlpha(j) + i + '"  class=\"w3-input w3-border\"/></td>';

                }
                htmlStr = htmlStr + "</tr>";
            }

            htmlStr = htmlStr + "</table>";

            document.getElementById("excelTable").innerHTML = htmlStr;
        }



        function colsToAlpha(cols) {
            var i;
            var reminder;
            var retVal = "";
            var colms = "_ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            if (cols > 26) {
                reminder = cols % 26;
                i = Math.floor(cols / 26);
                retVal = colms[i] + colms[reminder];
            }
            else {
                retVal = colms[cols];
            }

            return retVal

        }




        // Check for the various File API support.
        if (window.File && window.FileReader && window.FileList && window.Blob) {
            // Great success! All the File APIs are supported.
        } else {
            alert('The File APIs are not fully supported in this browser.');
        }



        function dndListenerOver(evt) {
            evt.stopPropagation();
            evt.preventDefault();
            evt.dataTransfer.dropEffect = 'copy'; // Explicitly show this is a copy.
        }


        function dndListenerHandler(evt) {
            evt.stopPropagation();
            evt.preventDefault();
            var files = evt.dataTransfer.files;

            if (files.length > 1) {
                alert("One file at a time!");
                return;
            }
            else {
                getFileDetails(evt.dataTransfer.files[0]);
            }

        }


        function getFileDetails(file) {


            var htmlBlock = [];


            htmlBlock.push('<li><strong>', escape(file.name), '</strong> (', file.type || 'n/a', ') - ',
                  file.size, ' bytes, last modified: ',
                  file.lastModifiedDate ? file.lastModifiedDate.toLocaleDateString() : 'n/a',
                  '</li>');

            document.getElementById('table1').innerHTML = '<ul>' + htmlBlock.join('') + '</ul>';
            // reader.onload = (function())
        }

        function showZipContent() {
            document.getElementById("table2").innerHTML = document.getElementById("HiddenField1").value;
            //alert("show");
        }

        function showSheets() {
            document.getElementById("table3").innerText = document.getElementById("HiddenField2").value;
            //alert("show");
        }



    </script>


<body>

    <form id="form1" runat="server">
    <table width="100%">
        <tr>
            <td>
               <h2 style="text-shadow:1px 1px 0 #444"> Upload Excel without ActiveX or server side components.</h2>
            </td>
        </tr>
        <tr>
            <td>

                <br />

            </td>
        </tr>
        <tr>
            <td>
                <table>
                    <tr>
                        <td>
                            <asp:FileUpload ID="fileUpload" runat="server" Width="536px" />
                        </td>
                        <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;</td>
                        <td>
                            <asp:Button ID="btnOpenFile" runat="server" Text="Upload File" OnClick="btnOpenFile_Click" />
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <asp:DropDownList ID="ddlSheetList" runat="server" Height="16px" Width="293px" AutoPostBack="True" OnSelectedIndexChanged="ddlSheetList_SelectedIndexChanged"></asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td>
                <br />
            </td>
        </tr>
        <tr>
            <td>
                <div id="excelTable"></div>
            </td>
        </tr>
    </table>
        <asp:HiddenField ID="HiddenField1" runat="server" />
        <asp:HiddenField ID="HiddenField2" runat="server" />
        <asp:HiddenField ID="hfRange" runat="server" />
        <asp:HiddenField ID="hfCommonValues" runat="server" />
        <asp:HiddenField ID="hfContent" runat="server" />
    </form>
</body>
</html>
