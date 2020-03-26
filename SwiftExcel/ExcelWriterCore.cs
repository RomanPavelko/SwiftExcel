using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace SwiftExcel
{
    public abstract class ExcelWriterCore : Disposable
    {
        protected internal bool Finalized;
        protected internal string FilePath;
        protected internal Sheet Sheet;

        protected internal string OutputPath;
        protected internal string TempOutputPath;

        protected ExcelWriterCore(string filePath, Sheet sheet)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new Exception("FilePath must not be empty.");
            }

            FilePath = filePath;
            Sheet = sheet ?? new Sheet { Name = "sheet 1" };

            Init();
        }

        protected internal void Init()
        {
            var fileInfo = new FileInfo(FilePath);
            OutputPath = fileInfo.Directory?.FullName;
            TempOutputPath = $"{OutputPath}/{Guid.NewGuid()}";

            DirectoryHelper.CheckCreatePath(TempOutputPath);

            CreateFolders();

            CreateRels();
            CreateDocProps();
            CreateContentTypes();
            CreateTheme();
            CreateExcelStyles();
            CreateWorkbook();

            StartSheets();
        }

        public void Save()
        {
            FinishSheets();

            try
            {
                DirectoryHelper.DeleteFile(FilePath);
                ZipFile.CreateFromDirectory(TempOutputPath, FilePath);
            }
            finally
            {
                DirectoryHelper.DeleteDirectory(TempOutputPath);
            }

            Finalized = true;
        }

        #region static content

        protected internal void CreateFolders()
        {
            Directory.CreateDirectory($"{TempOutputPath}/_rels");
            Directory.CreateDirectory($"{TempOutputPath}/docProps");
            Directory.CreateDirectory($"{TempOutputPath}/xl");
            Directory.CreateDirectory($"{TempOutputPath}/xl/_rels");
            Directory.CreateDirectory($"{TempOutputPath}/xl/theme");
            Directory.CreateDirectory($"{TempOutputPath}/xl/worksheets");
        }

        protected internal void CreateRels()
        {
            using (TextWriter tw = new StreamWriter($"{TempOutputPath}/_rels/.rels", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                         "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                         "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
                         "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
                         "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                         "</Relationships>");
            }
            using (TextWriter tw = new StreamWriter($"{TempOutputPath}/xl/_rels/workbook.xml.rels", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                         "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                         "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                         "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>" +
                         "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
                         "</Relationships>");
            }
        }

        protected internal void CreateDocProps()
        {
            using (TextWriter tw = new StreamWriter($"{TempOutputPath}/docProps/app.xml", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                         "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                         "<Application>Microsoft Excel</Application>" +
                         "<DocSecurity>0</DocSecurity>" +
                         "<ScaleCrop>false</ScaleCrop>" +
                         "<HeadingPairs>" +
                         "<vt:vector size=\"2\" baseType=\"variant\">" +
                         "<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>" +
                         $"<vt:variant><vt:i4>1</vt:i4></vt:variant>" +
                         "</vt:vector>" +
                         "</HeadingPairs>" +
                         "<TitlesOfParts>" +
                         $"<vt:vector size=\"1\" baseType=\"lpstr\"><vt:lpstr>{Sheet.Name}</vt:lpstr></vt:vector>" +
                         "</TitlesOfParts>" +
                         "<Company></Company>" +
                         "<LinksUpToDate>false</LinksUpToDate>" +
                         "<SharedDoc>false</SharedDoc>" +
                         "<HyperlinksChanged>false</HyperlinksChanged>" +
                         "<AppVersion>16.0300</AppVersion></Properties>");
            }
            using (TextWriter tw = new StreamWriter($"{TempOutputPath}/docProps/core.xml", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                         "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                         "<dc:creator></dc:creator>" +
                         "<cp:lastModifiedBy></cp:lastModifiedBy>" +
                         "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2015-06-05T18:17:20Z</dcterms:created>" +
                         "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2019-11-07T23:00:46Z</dcterms:modified>" +
                         "</cp:coreProperties>");
            }
        }

        protected internal void CreateContentTypes()
        {
            using (TextWriter tw = new StreamWriter($"{TempOutputPath}/[Content_Types].xml", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                         "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                         "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                         "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                         "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                         "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                         "<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>" +
                         "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
                         "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
                         "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>" +
                         "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>");
            }
        }

        protected internal void CreateTheme()
        {
            using (TextWriter tw = new StreamWriter($"{TempOutputPath}/xl/theme/theme1.xml", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                         "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">" +
                         "<a:themeElements>" +
                         "<a:clrScheme name=\"Office\">" +
                         "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>" +
                         "<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>" +
                         "<a:dk2><a:srgbClr val=\"44546A\"/></a:dk2>" +
                         "<a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>" +
                         "<a:accent1><a:srgbClr val=\"5B9BD5\"/></a:accent1>" +
                         "<a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>" +
                         "<a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3>" +
                         "<a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>" +
                         "<a:accent5><a:srgbClr val=\"4472C4\"/></a:accent5>" +
                         "<a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>" +
                         "<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink>" +
                         "<a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>" +
                         "</a:clrScheme>" +
                         "<a:fontScheme name=\"Office\">" +
                         "<a:majorFont>" +
                         "<a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/>" +
                         "<a:ea typeface=\"\"/>" +
                         "<a:cs typeface=\"\"/>" +
                         "<a:font script=\"Jpan\" typeface=\"游ゴシック Light\"/>" +
                         "<a:font script=\"Hang\" typeface=\"맑은 고딕\"/>" +
                         "<a:font script=\"Hans\" typeface=\"等线 Light\"/>" +
                         "<a:font script=\"Hant\" typeface=\"新細明體\"/>" +
                         "<a:font script=\"Arab\" typeface=\"Times New Roman\"/>" +
                         "<a:font script=\"Hebr\" typeface=\"Times New Roman\"/>" +
                         "<a:font script=\"Thai\" typeface=\"Tahoma\"/>" +
                         "<a:font script=\"Ethi\" typeface=\"Nyala\"/>" +
                         "<a:font script=\"Beng\" typeface=\"Vrinda\"/>" +
                         "<a:font script=\"Gujr\" typeface=\"Shruti\"/>" +
                         "<a:font script=\"Khmr\" typeface=\"MoolBoran\"/>" +
                         "<a:font script=\"Knda\" typeface=\"Tunga\"/>" +
                         "<a:font script=\"Guru\" typeface=\"Raavi\"/>" +
                         "<a:font script=\"Cans\" typeface=\"Euphemia\"/>" +
                         "<a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/>" +
                         "<a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/>" +
                         "<a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/>" +
                         "<a:font script=\"Thaa\" typeface=\"MV Boli\"/>" +
                         "<a:font script=\"Deva\" typeface=\"Mangal\"/>" +
                         "<a:font script=\"Telu\" typeface=\"Gautami\"/>" +
                         "<a:font script=\"Taml\" typeface=\"Latha\"/>" +
                         "<a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/>" +
                         "<a:font script=\"Orya\" typeface=\"Kalinga\"/>" +
                         "<a:font script=\"Mlym\" typeface=\"Kartika\"/>" +
                         "<a:font script=\"Laoo\" typeface=\"DokChampa\"/>" +
                         "<a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/>" +
                         "<a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/>" +
                         "<a:font script=\"Viet\" typeface=\"Times New Roman\"/>" +
                         "<a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/>" +
                         "<a:font script=\"Geor\" typeface=\"Sylfaen\"/>" +
                         "</a:majorFont>" +
                         "<a:minorFont>" +
                         "<a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/>" +
                         "<a:ea typeface=\"\"/>" +
                         "<a:cs typeface=\"\"/>" +
                         "<a:font script=\"Jpan\" typeface=\"游ゴシック\"/>" +
                         "<a:font script=\"Hang\" typeface=\"맑은 고딕\"/>" +
                         "<a:font script=\"Hans\" typeface=\"等线\"/>" +
                         "<a:font script=\"Hant\" typeface=\"新細明體\"/>" +
                         "<a:font script=\"Arab\" typeface=\"Arial\"/>" +
                         "<a:font script=\"Hebr\" typeface=\"Arial\"/>" +
                         "<a:font script=\"Thai\" typeface=\"Tahoma\"/>" +
                         "<a:font script=\"Ethi\" typeface=\"Nyala\"/>" +
                         "<a:font script=\"Beng\" typeface=\"Vrinda\"/>" +
                         "<a:font script=\"Gujr\" typeface=\"Shruti\"/>" +
                         "<a:font script=\"Khmr\" typeface=\"DaunPenh\"/>" +
                         "<a:font script=\"Knda\" typeface=\"Tunga\"/>" +
                         "<a:font script=\"Guru\" typeface=\"Raavi\"/>" +
                         "<a:font script=\"Cans\" typeface=\"Euphemia\"/>" +
                         "<a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/>" +
                         "<a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/>" +
                         "<a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/>" +
                         "<a:font script=\"Thaa\" typeface=\"MV Boli\"/>" +
                         "<a:font script=\"Deva\" typeface=\"Mangal\"/>" +
                         "<a:font script=\"Telu\" typeface=\"Gautami\"/>" +
                         "<a:font script=\"Taml\" typeface=\"Latha\"/>" +
                         "<a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/>" +
                         "<a:font script=\"Orya\" typeface=\"Kalinga\"/>" +
                         "<a:font script=\"Mlym\" typeface=\"Kartika\"/>" +
                         "<a:font script=\"Laoo\" typeface=\"DokChampa\"/>" +
                         "<a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/>" +
                         "<a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/>" +
                         "<a:font script=\"Viet\" typeface=\"Arial\"/>" +
                         "<a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/>" +
                         "<a:font script=\"Geor\" typeface=\"Sylfaen\"/>" +
                         "</a:minorFont>" +
                         "</a:fontScheme>" +
                         "<a:fmtScheme name=\"Office\">" +
                         "<a:fillStyleLst>" +
                         "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                         "<a:gradFill rotWithShape=\"1\">" +
                         "<a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/>" +
                         "</a:gradFill>" +
                         "<a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/>" +
                         "</a:gradFill>" +
                         "</a:fillStyleLst>" +
                         "<a:lnStyleLst>" +
                         "<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>" +
                         "</a:lnStyleLst>" +
                         "<a:effectStyleLst>" +
                         "<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>" +
                         "</a:effectStyleLst>" +
                         "<a:bgFillStyleLst>" +
                         "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>" +
                         "</a:bgFillStyleLst>" +
                         "</a:fmtScheme>" +
                         "</a:themeElements>" +
                         "<a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E510DB2}\"><thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\"/></a:ext></a:extLst></a:theme>");
            }
        }

        protected internal void CreateWorkbook()
        {
            using (TextWriter tw = new StreamWriter($"{TempOutputPath}/xl/workbook.xml", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                         "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x15 xr xr6 xr10 xr2\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\">" +
                         "<bookViews><workbookView xWindow=\"3345\" yWindow=\"3675\" windowWidth=\"21600\" windowHeight=\"11385\" xr2:uid=\"{{0B3BF63D-56DA-4710-9C33-DB5A76182BCF}}\"/></bookViews>" +
                         $"<sheets><sheet name=\"{Sheet.Name}\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                         "</workbook>");
            }
        }

        protected internal void CreateExcelStyles()
        {
            using (TextWriter tw = new StreamWriter($"{TempOutputPath}/xl/styles.xml", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                         "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac x16r2 xr\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\">" +
                         "<fonts count=\"1\" x14ac:knownFonts=\"1\">" +
                         "<font><sz val=\"11\"/>" +
                         "<color theme=\"1\"/>" +
                         "<name val=\"Calibri\"/>" +
                         "<family val=\"2\"/>" +
                         "<scheme val=\"minor\"/>" +
                         "</font>" +
                         "</fonts>" +
                         "<fills count=\"2\">" +
                         "<fill><patternFill patternType=\"none\"/></fill>" +
                         "<fill><patternFill patternType=\"gray125\"/></fill>" +
                         "</fills>" +
                         "<borders count=\"1\">" +
                         "<border><left/><right/><top/><bottom/><diagonal/></border>" +
                         "</borders>" +
                         "<cellStyleXfs count=\"1\">" +
                         "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                         "</cellStyleXfs>" +
                         "<cellXfs count=\"1\">" +
                         "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
                         "</cellXfs>" +
                         "<cellStyles count=\"1\">" +
                         "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>" +
                         "</cellStyles>" +
                         "<dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>" +
                         "</styleSheet>");
            }
        }

        protected internal void StartSheets()
        {
            Sheet.TextWriter = new StreamWriter($"{TempOutputPath}/xl/worksheets/sheet1.xml", false);

            Sheet.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" xr:uid=\"{{EA47BE14-E914-42F7-BE8E-AEEE5780E9D7}}\">" +
                        "<dimension ref=\"A1\"/>" +
                        "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>" +
                        "<sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/>");

            //write column definition
            if (Sheet.ColumnsWidth != null && Sheet.ColumnsWidth.Any())
            {
                Sheet.Write("<cols>");
                for (var i = 0; i < Sheet.ColumnsWidth.Count; i++)
                {
                    Sheet.Write(GetExcelColumnDefinition(Sheet.ColumnsWidth[i].ToString(CultureInfo.InvariantCulture), i + 1));
                }
                Sheet.Write("</cols>");
            }

            Sheet.Write("<sheetData>");
        }

        private static string GetExcelColumnDefinition(string width, int col)
        {
            return $"<col width=\"{width}\" min=\"{col}\" max=\"{col}\"/>";
        }

        protected internal void FinishSheets()
        {
            if (Sheet.CurrentRow != 0)
            {
                Sheet.Write("</row>");
            }
            Sheet.Write("</sheetData><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>");
            Sheet.TextWriter.Close();
        }

        #endregion

        protected override void DisposeCore()
        {
            if (!Finalized)
            {
                Save();
            }
        }
    }
}