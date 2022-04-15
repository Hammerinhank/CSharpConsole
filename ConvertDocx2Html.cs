using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;


namespace CSharpConsole
{
    class ConvertDocx2Html
    {
        static string GLOBAL_fileNameDstWithoutExtension;
        static XmlNode rel;
        static XmlNode num; static int firstImageId = 1;
        // THIS POC IS ONLY ADRESSING BASIC IMPLEMENTATION OF NUMBERED ITEMS
        static int numberedItem;   // This is a counter for numbered items
        static int bulletNumId; // ID of the current list of items
        static int previousListId; // ID of the previous list of items
        static string folderNameDst;
        // THIS POC IS ONLY ADRESSING BASIC IMPLEMENTATION OF NUMBERED ITEMS

        //DOC https://stackoverflow.com/questions/55828/how-does-one-parse-xml-files
        static public void henri(string folderNameSrc, string folderNameDst0, string fileNameDstWithoutExtension)
        {
            folderNameDst = folderNameDst0;
            GLOBAL_fileNameDstWithoutExtension = fileNameDstWithoutExtension;
            var fileNameDst = fileNameDstWithoutExtension + ".html";
            var filePath_Document_xml = folderNameSrc + @"\word\document.xml";
            var filePath_Document_rel = folderNameSrc + @"\word\_rels\document.xml.rels";
            var filePath_Document_num = folderNameSrc + @"\word\_rels\numbering.xml.rels";
            Console.Write("\n File CHECK starts here");
            if (!File.Exists(filePath_Document_xml)) {Console.Write("\n File \"" + filePath_Document_xml + "\" is missing => ABORT"); return; }
            if (!File.Exists(filePath_Document_rel)) { Console.Write("\n File \"" + filePath_Document_rel + "\" is missing => ABORT"); return; }
            if (!File.Exists(filePath_Document_num)) { Console.Write("\n File \"" + filePath_Document_num + "\" is missing => ABORT");}
            else
            {
              XmlDocument document_num = new XmlDocument(); 
              document_num.Load(filePath_Document_num);
              num = document_num.LastChild;
              checkTheNumberingFileContent();
            }

            XmlDocument document_xml = new XmlDocument();
            XmlDocument document_rel = new XmlDocument();
            

            document_xml.Load(filePath_Document_xml);
            document_rel.Load(filePath_Document_rel);
            
            XmlNode document = document_xml.LastChild;
            rel = document_rel.LastChild;
            

            XmlNode body = document.LastChild;
            Console.Write("\nPress  to exit... " + document.Name);
            Int32 count = body.ChildNodes.Count;
            Console.Write("\nHyperlink => " + getHyperlink("rId5") + "\n");
            
            transferImages(folderNameSrc, folderNameDst , fileNameDst);
            Console.WriteLine("\n");
            var html = scanDocx(body);
            html = "<html version='0.1'>" + html + "</html>";
            string fullFileNameDst = folderNameDst + @"\" + fileNameDst;
            bool exists = File.Exists(fullFileNameDst);
            if (!exists)
            {
                System.IO.File.WriteAllText(fullFileNameDst, html);
                Console.Write("\n\nFile \"" + fullFileNameDst + "\" has been generated\n");
            }
            else
            { Console.Write("\n\nALREADY HERE!!!!\n" + "File \"" + fullFileNameDst + "\" already there\n"); }
        }

        /*
         *  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	            <Relationship Id="rId26"
            		            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            		            Target="media/image18.jpg"/>
	            <Relationship Id="rId117"
            		            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            		            Target="media/image79.jpg"/>
	            <Relationship Id="rId21"
            		            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
            		            Target="footer1.xml"/>
	            <Relationship Id="rId42"
            		            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
            		            Target="header9.xml"/>
	            <Relationship Id="rId47"
            		            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            		            Target="media/image27.jpg"/ .....
        */

        static private string getHyperlink(string id)
        {
            for (var i0 = 0; i0 < rel.ChildNodes.Count; i0++)
            {
                var elt0 = rel.ChildNodes[i0];
                if (elt0.Attributes["Id"].Value == id) return elt0.Attributes["Target"].Value;
            }
            return "";
        }

        static private string scanDocx(XmlNode elt)
        {
            numberedItem = 0;
            previousListId = -1;
            string html = "<style>table, th, td {border:1px solid black;border-collapse:collapse;}</style>";
            html += "<style>p {font-size:11.0pt;font-family:'Calibri',sans-serif;}</style>";
            for (var i0 = 0; i0 < elt.ChildNodes.Count; i0++)
            {
                var elt0 = elt.ChildNodes[i0];
                switch (elt0.Name.ToUpper())
                {
                    case "W:TBL":
                        html += scanTable(elt0);
                        break;
                    case "W:P":
                        html += scanParagraph(elt0);
                        break;
                }
            }
            return html;
        }



        static private string scanParagraph(XmlNode elt0)
        {
            string html = "";
            int posi = 0;
            string commentList = "";
            string commentSepar = "";
            bool bold = false;
            bool italic = false;
            bool underline = false;
            int bulletlevel = -1;
            var paraId = elt0.Attributes["w14:paraId"].Value;
            // Console.Write("\nNode " + i0 + " => " + elt0.Name.ToUpper() + " => " + paraId);
            if (paraId == "7EDFEB2D")
                Console.Write("ParaId => " + paraId);
            // html += "<p>";
            for (var i1 = 0; i1 < elt0.ChildNodes.Count; i1++)
            {
                var elt1 = elt0.ChildNodes[i1];
                switch (elt1.Name.ToUpper())
                {
                    case "W:P":
                        html += scanParagraph(elt1);
                        break;
                    case "W:HYPERLINK":
                        try
                        {
                            var hyperlinkId = elt1.Attributes["r:id"].Value;
                            Console.Write("\nHyperlink ##################### => " + getHyperlink(hyperlinkId));
                            html += "<a href='" + getHyperlink(hyperlinkId) + "'>HYPERLINKS</a>";
                        }
                        catch (Exception e) { html += "MISSING HYPERLINK!!!";  }
                        break;
                    case "W:PPR":
                        for (var i2 = 0; i2 < elt1.ChildNodes.Count; i2++)
                        {
                            XmlNode elt2 = elt1.ChildNodes[i2];
                            switch (elt2.Name.ToUpper())
                            {
                                case "W:RPR":
                                    scanRPR(elt2);
                                    break;
                                case "W:NUMPR":
                                    scanNUMPR(elt2);
                                    break;
                            }
                        }
                        break;
                    case "W:R":
                        bold = false;
                        italic = false;
                        underline = false;
                        for (var i2 = 0; i2 < elt1.ChildNodes.Count; i2++)
                        {
                            XmlNode elt2 = elt1.ChildNodes[i2];
                            switch (elt2.Name.ToUpper())
                            {
                                case "MC:ALTERNATECONTENT":
                                        for (var i3 = 0; i3 < elt2.ChildNodes.Count; i3++)
                                        {
                                            XmlNode elt3 = elt2.ChildNodes[i3];
                                            switch (elt3.Name.ToUpper())
                                            {
                                                case "MC:CHOICE":
                                                    for (var i4 = 0; i4 < elt3.ChildNodes.Count; i4++)
                                                    {
                                                        XmlNode elt4 = elt3.ChildNodes[i4];
                                                        switch (elt4.Name.ToUpper())
                                                        {
                                                            case "W:DRAWING":
                                                            html += scanDrawing(elt4);
                                                            break;
                                                        }
                                                    }
                                                break;
                                            }
                                        }
                                        break;
                                case "W:RPR":
                                    scanRPR(elt2);
                                    break;
                                case "W:DRAWING":
                                    html += scanDrawing(elt2);
                                    break;
                                case "W:T":
                                    if (bold) html += "<b>";
                                    if (italic) html += "<i>";
                                    if (underline) html += "<u>";
                                    html += elt2.InnerText;
                                    posi += elt2.InnerText.Length;
                                    if (italic) html += "</i>";
                                    if (bold) html += "</b>";
                                    if (underline) html += "</u>";
                                    bold = false;
                                    italic = false;
                                    underline = false;
                                    // Console.WriteLine("TAG W:T text => " + elt2.InnerText);
                                    break;
                            }
                        }
                        break;
                    case "W:COMMENTRANGESTART":
                        var commentId_start = elt1.Attributes["w:id"].Value;
                        Console.WriteLine("commentId_start => " + commentId_start);
                        // html += "<comment id='" + commentId_start + "' type='start' posi='" + posi + "'/>";
                        commentList += commentSepar + "{type:0,id:" + commentId_start + ",posi:" + posi + "}";
                        commentSepar = ",";
                        break;
                    case "W:COMMENTRANGEEND":
                        var commentId_end = elt1.Attributes["w:id"].Value;
                        Console.WriteLine("commentId_end => " + commentId_end);
                        commentList += commentSepar + "{type:1,id:" + commentId_end + ",posi:" + posi + "}";
                        commentSepar = ",";
                        // html += "<comment id='" + commentId_end + "' type='end' posi='" + posi + "'/>";
                        break;
                }
            }
            string bullet = "";
            int indentation = 30;
            string space = "&nbsp;&nbsp;&nbsp;";
            string paragraphPrefix = "<p";
            if (commentList != "") paragraphPrefix += " comments=\"[" + commentList + "]\" ";
            switch (bulletlevel)
            {
                case -1:
                    html = paragraphPrefix + ">" + html + "</p>";  // it should not be +=
                    break;
                case 0:
                    //if (bulletNumId == 1)
                    //    bullet = "<span style='font-family:Symbol'>·</span>" + space;
                    //else
                        bullet = "<span style='style='font-family:Courier New>" + numberedItem + "</span>&nbsp;";
                    html = paragraphPrefix + " style='margin-left:" + 20 + ".0pt;'>" + bullet + html + "</p>"; // it should not be +=
                    break;
                default:
                    if (bulletNumId == 1)
                    {
                        switch (bulletlevel)
                        {
                            case 1: bullet = "<span style='font-family:Courier New'>o</span>" + space; break;
                            case 2: bullet = "<span style='font-family:Wingdings'>§</span>" + space; break;
                        }
                    }
                    else
                        bullet = "<span style='style='font-family:Courier New>" + numberedItem + "</span>&nbsp;";
                    html = paragraphPrefix + " style='margin-left:" + (30 + (indentation * bulletlevel)) + ".0pt;text-indent:-18.0pt'>" + bullet + html + "</p>"; // it should not be +=
                    break;
            }
            if (html == "<p></p>") html = "<p>&nbsp;</p>";
            return html;

            string scanDrawing(XmlNode elt2) {
                html = "<td>";
                for (var i3 = 0; i3 < elt2.ChildNodes.Count; i3++)
                {

                    XmlNode elt3 = elt2.ChildNodes[i3];
                    string tmp = elt3.Name.ToUpper();
                    switch (tmp)
                    {
                        case "WP:INLINE":
                        case "WP:ANCHOR":
                            var cx = 0;
                            var cy = 0;
                            for (var i4 = 0; i4 < elt3.ChildNodes.Count; i4++)
                            {
                                XmlNode elt4 = elt3.ChildNodes[i4];
                                string tmp1 = elt4.Name.ToUpper();
                                switch (tmp1)
                                {
                                    case "WP:EXTENT":
                                        cx = Int32.Parse(elt4.Attributes["cx"].Value);
                                        cy = Int32.Parse(elt4.Attributes["cy"].Value);
                                        break;
                                    
                                    case "WP:DOCPR":
                                        var imageId = elt4.Attributes["id"].Value;
                                        var imageName = elt4.Attributes["name"].Value;
                                        // alert("HERE imageId = " + imageId);
                                        // html += htmlImage(imageId, imageName,cx,cy);
                                        break;
                                    case "A:GRAPHIC":
                                        for (var i5 = 0; i5 < elt4.ChildNodes.Count; i5++)
                                        {
                                            XmlNode elt5 = elt4.ChildNodes[i5];
                                            switch (elt5.Name.ToUpper())
                                            {
                                                case "A:GRAPHICDATA":
                                                    for (var i6 = 0; i6 < elt5.ChildNodes.Count; i6++)
                                                    {
                                                        XmlNode elt6 = elt5.ChildNodes[i6];
                                                        string tmp6 = elt6.Name.ToUpper();
                                                        switch (tmp6)
                                                        {
                                                            case "PIC:PIC":
                                                                html += scanPic(elt6,cx,cy);
                                                                break;
                                                            case "WPG:WGP":
                                                                for (var i7 = 0; i7 < elt6.ChildNodes.Count; i7++)
                                                                {
                                                                    XmlNode elt7 = elt6.ChildNodes[i7];
                                                                    string tmp7 = elt7.Name.ToUpper();
                                                                    switch (tmp7)
                                                                    {
                                                                        
                                                                        case "PIC:PIC":
                                                                            html += scanPic(elt7, cx, cy);
                                                                            break;
                                                                    }
                                                                }
                                                                break;
                                                        }
                                                    }
                                                    break;
                                            }
                                        }
                                        break;
                                }
                            }
                            break;
                    }
                }
                return html;
            }

            static string scanPic(XmlNode elt7, int cx, int cy)
            {
                string html = "";
                for (var i8 = 0; i8 < elt7.ChildNodes.Count; i8++)
                {
                    XmlNode elt8 = elt7.ChildNodes[i8];
                    string tmp8 = elt8.Name.ToUpper();
                    switch (tmp8)
                    {
                        case "PIC:BLIPFILL":
                            //ICI
                            for (var i9 = 0; i9 < elt8.ChildNodes.Count; i9++)
                            {
                                XmlNode elt9 = elt8.ChildNodes[i9];
                                string tmp9 = elt9.Name.ToUpper();
                                switch (tmp9)
                                {
                                    case "A:BLIP":
                                        // var imageId = elt6.Attributes["id"].Value;
                                        // alert("HERE imageId = " + imageId);
                                        string tmp9a = elt9.Attributes["r:embed"].Value;
                                        string imageName9 = getHyperlink(tmp9a);
                                        if (imageName9.Substring(0, 6) == "media/") imageName9 = imageName9.Substring(6);
                                        html += htmlImage("imageId", imageName9, cx, cy);
                                        break;
                                }
                            }
                            break;
                    }
                }
                return html;
            }


            void scanRPR(XmlNode elt2)
            {
                for (var i3 = 0; i3 < elt2.ChildNodes.Count; i3++)
                {
                    var elt3 = elt2.ChildNodes[i3];
                    switch (elt3.Name.ToUpper())
                    {
                        case "W:B":
                            bold = true;
                            break;
                        case "W:I":
                            italic = true;
                            break;
                        case "W:U":
                            underline = true;
                            break;
                    }
                }
            }

            void scanNUMPR(XmlNode elt2)
            {
                for (var i3 = 0; i3 < elt2.ChildNodes.Count; i3++)
                {
                    XmlNode elt3 = elt2.ChildNodes[i3];
                    switch (elt3.Name.ToUpper())
                    {
                        case "W:ILVL":
                            bulletlevel = Int32.Parse(elt3.Attributes["w:val"].Value);
                            break;
                        case "W:NUMID":
                            bulletNumId = Int32.Parse(elt3.Attributes["w:val"].Value);
                            if (previousListId != bulletNumId) numberedItem = 1; else numberedItem++;
                            previousListId = bulletNumId;
                            break;
                    }
                }
            }

        }

        static private string scanTable(XmlNode elt0)
        {
            var html = "<br/><div style='width:100%;padding:15px;'><table style='width:100%;'>";
            for (var i1 = 0; i1 < elt0.ChildNodes.Count; i1++)
            {
                var elt1 = elt0.ChildNodes[i1];
                switch (elt1.Name.ToUpper())
                {
                    case "W:TR":
                        html += "<tr>";
                        for (var i2 = 0; i2 < elt1.ChildNodes.Count; i2++)
                        {
                            var elt2 = elt1.ChildNodes[i2];
                            switch (elt2.Name.ToUpper())
                            {
                                case "W:TC":
                                    html += "<td>";
                                    for (var i3 = 0; i3 < elt2.ChildNodes.Count; i3++)
                                    {
                                        var elt3 = elt2.ChildNodes[i3];
                                        switch (elt3.Name.ToUpper())
                                        {
                                            case "W:P":
                                                html += scanParagraph(elt3);
                                                //html += '<p>' + elt3.innerText + '</p>'; 
                                                break;
                                            case "W:TBL":
                                                html += scanTable(elt3);
                                                break;
                                        }
                                    }
                                    html += "</td>";
                                    break;
                            }
                        }
                        html += "</tr>";
                        break;
                }
            }
            html += "</table></div>";
            return html;
        }
        static private string htmlImage(string id, string imageName,int cx, int cy)
        {
            string html = "";
            string tmp = folderNameDst + @"\" + GLOBAL_fileNameDstWithoutExtension;
            FileInfo fi = new FileInfo(imageName);
            // NO MORE... string imageName = checkImageExtension();
            string imageFullFileNameDst = tmp + @"\" + imageName;
            Console.WriteLine(">> " + firstImageId + " => " + GLOBAL_fileNameDstWithoutExtension + " => " + imageName); 
            bool exists = File.Exists(imageFullFileNameDst);
            if (exists)
            {
                string size = getSize(imageFullFileNameDst, cx, cy);
                // Console.WriteLine(" , size => " + size);
                // html += "<p class=MsoNormal>";
                html += "<img id='Picture 1'" + size + "src = './" + GLOBAL_fileNameDstWithoutExtension + "/" + imageName + "' alt = 'A picture...'>"; 
                // html += "</p>";
                // html += "<br/><p style='background-color:lime'>Image file " + imageName + " - " + imageName + " size => " + size + " cx = " + cx + " cy = " + cy + "</p><br/>";
            }
            else
            {
                html = "<br/><p style='background-color:lightpink'>Image file " + firstImageId + "\"" + imageName + "\" not found in folder \"" + tmp + "\"</p><br/>";
            }
            firstImageId++; 
            return html;
        }



        static string checkImageExtension() {
            string shortLeftPart = "image" + firstImageId;
            string longLeftPart = folderNameDst + @"\" + GLOBAL_fileNameDstWithoutExtension + @"\" + shortLeftPart;
            string extension = ".jpeg";
            if (File.Exists(longLeftPart + extension)) return shortLeftPart + extension;
            extension = ".jpg";
            if (File.Exists(longLeftPart + extension)) return shortLeftPart + extension;
            extension = ".png";
            if (File.Exists(longLeftPart + extension)) return shortLeftPart + extension;
            return "";
        }

        static private void transferImages(string folderNameSrc, string folderNameDst , string fileNameDst) {
            string imageFolderNameSrc = folderNameSrc + @"\word\media";
            string imageFolderNameDst = folderNameDst + @"\" + Path.GetFileNameWithoutExtension(fileNameDst);
            Console.Write("\nTRANSFER IMAGES:\nFROM: " + imageFolderNameSrc + "\nTO: " + imageFolderNameDst);
            // Your code goes here
            bool existsDst = System.IO.Directory.Exists(imageFolderNameDst);
            if (!existsDst) System.IO.Directory.CreateDirectory(imageFolderNameDst);
            bool existsSrc = System.IO.Directory.Exists(imageFolderNameSrc);
            if (existsSrc)
            {
                DirectoryInfo d = new DirectoryInfo(imageFolderNameSrc); //Assuming Test is your Folder
                FileInfo[] Files = d.GetFiles("*.*"); //Getting Text files
                string str = ""; int imageCounter = 0;
                foreach (FileInfo file in Files)
                {
                    str += ", " + file.Name;
                    string imageFullFileNameDst = imageFolderNameDst + @"\" + file.Name;
                    bool exists1 = File.Exists(imageFullFileNameDst);
                    if (!exists1) File.Copy(imageFolderNameSrc + @"\" + file.Name, imageFullFileNameDst);
                    // else str +=  " DUPLICATED " ;
                    imageCounter++;
                }
                Console.Write("\n" + imageCounter + " Images: " + str);
            }
        }

        static string getSize(string imageFullFileNameDst, float cx, float cy)
        {
            if (cx == 0.0 || cy == 0.0) Console.Write("\nWARNING on cx, cy");
            System.Drawing.Image image = System.Drawing.Image.FromFile(imageFullFileNameDst);
            float dpiX = image.HorizontalResolution;
            float dpiY = image.VerticalResolution;
            //Console.Write("\n dpiX = " + dpiX + " dpiY = " + dpiY);
            float widthInInch = cx / 914400;
            float heigthInInch = cy / 914400;
            //Console.Write("\n widthInInch = " + widthInInch + " heigthInInch = " + heigthInInch);
            float widthInPixel = widthInInch * dpiX;
            float heigthInPixel = heigthInInch * dpiY;
            widthInPixel = (int)Math.Round(widthInPixel, 0);
            heigthInPixel = (int)Math.Round(heigthInPixel, 0);
            // Console.Write("\n widthInPixel = " + widthInPixel + " heigthInPixel = " + heigthInPixel);
            string tmp = " width=" + widthInPixel + " height= " + heigthInPixel + " ";
            // Console.Write("\n " + tmp);
            return tmp; 
        }
        static void checkTheNumberingFileContent() {
            for (var i0 = 0; i0 < num.ChildNodes.Count; i0++)
            {
                var elt0 = num.ChildNodes[i0];
                Console.WriteLine("Image found in file numbering.xml.rels " + firstImageId + " => " + elt0.Attributes["Target"].Value);
                firstImageId++;
                // if (elt0.Attributes["Id"].Value == id) return elt0.Attributes["Target"].Value;
            }
            return;
        }
    }
}
