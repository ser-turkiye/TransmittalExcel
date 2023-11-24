package ser;

import org.json.JSONObject;

import java.util.List;

public class Conf {
    public static class ExcelTransmittalNodes {
        public static final String FromExcelNodeName = "Generate Transmittal From Excel";
    }
    public static class ExcelTransmittalPaths {
        public static final String MainPath = "C:/tmp2/bulk/transmittal";
    }
    public static class ExcelTransmittalSheetIndex {
        public static final Integer FromExcel = 0;
        public static final Integer Cover = 0;
        public static final Integer Mail = 0;
    }
    public static class ExcelTransmittalRowGroups {
        public static final List<Integer> CoverHideCols = List.of();
        public static final String CoverDocs = "DocRows";
        public static final Integer CoverDocColInx = 0;

        public static final List<Integer> MailDocHideCols = List.of(0);
        public static final String MailDocs = "DocRows";
        public static final Integer MailDocColInx = 0;

        public static final List<Integer> MailDistHideCols = List.of(0);
        public static final String MailDists = "DstRows";
        public static final Integer MailDistColInx = 0;
    }
    public static class ExcelTransmittalDocsCellIndex {
        public static final Integer DocNo = 1;
        public static final Integer RevNo = 2;


    }
    public static class ExcelTransmittalDocsRowIndex {
        public static final Integer Begin = 1;
        public static final Integer End = 20;


    }
    public static class ExcelTransmittalDocsCellPos {
        public static final String ProjectNo = "F1";
        public static final String To = "F2";
        public static final String Attention = "F3";
        public static final String CC = "F4";
        public static final String JobNo = "F6";
        public static final String IssueDate = "F7";
        public static final String TransmittalType = "F8";
        public static final String Summary = "F10";
        public static final String Notes = "F16";

    }
    public static class ExcelTransmittalDistCellIndex {
        public static final Integer User = 1;
        public static final Integer Purpose = 2;
        public static final Integer DlvMethod = 4;
        public static final Integer DueDate = 5;

    }
    public static class ExcelTransmittalDistRowIndex {
        public static final Integer Begin = 24;
        public static final Integer End = 29;

    }

    public static class Descriptors{
        public static final String MainDocumentID = "MainDocumentReference";
        public static final String DocTransOutCode = "ccmPrjDocTransOutCode";
        public static final String FileName = "ccmPrjDocFileName";
        public static final String DocNumber = "ccmPrjDocNumber";
        public static final String DocRevision = "ccmPrjDocRevision";
        public static final String DocType = "ccmPrjDocDocType";
        public static final String ObjectName = "ObjectName";
        public static final String Originator = "ccmPrjDocOiginator";
        public static final String Prefix = "ccmPRJCard_prefix";
        public static final String Category = "ccmPrjDocCategory";
        public static final String ParentDocNumber = "ccmPrjDocParentDoc";
        public static final String ParentDocRevision = "ccmPrjDocParentDocRevision";
        public static final String ObjectNumberExternal = "ObjectNumberExternal";
        public static final String TrmtSendType = "ccmTrmtSendType";
        public static final String ProjectNo = "ccmPRJCard_code";
        public static final String ProjectName = "ccmPRJCard_name";
        public static final String DccList = "ccmPrjCard_DccList";
        public static final String ObjectNumber = "ObjectNumber";

    }

    public static class DescriptorLiterals{
        public static final String PrjDocNumber = "CCMPRJDOCNUMBER";
        public static final String PrjCardCode = "CCMPRJCARD_CODE";
        public static final String ObjectNumberExternal = "OBJECTNUMBER2";
        public static final String PrjDocRevision = "CCMPRJDOCREVISION";
        public static final String ReferenceNumber = "CCMREFERENCENUMBER";
        public static final String PrjDocDocType = "CCMPRJDOCDOCTYPE";
        public static final String PrjDocParentDoc = "CCMPRJDOCPARENTDOC";
        public static final String PrjDocParentDocRevision = "CCMPRJDOCPARENTDOCREVISION";

    }
    public static class ClassIDs{
        public static final String Template = "b9cf43d1-a4d3-482f-9806-44ae64c6139d";
        public static final String EngineeringDocument = "3b3078f8-c0d0-4830-b963-88bebe1c1462";
        public static final String ProjectWorkspace = "32e74338-d268-484d-99b0-f90187240549";
        public static final String EngineeringProjectTransmittal = "8bf0a09b-b569-4aef-984b-78cf1644ca19";
        public static final String EngineeringCRS = "3e1fe7b3-3e86-4910-8155-c29b662e71d6";

    }
    public static class Licences{
        public static final String SPIRE_XLS = "RaOICL4BAJZcQ/syuVNDcP9iQcvVrkO1p2mxEcDAH1meuQwbm6Y/MdlNOBpSKLEOvbFRv0OfDcADt4xBm+phBbvNgCo62gounm/x5hUMcaPI+srTqZzJKAcgMj5rCPwQ4IkNGeGjvRGP5m+l2kMpU+bXyhFxQdk8TcCO+nkQdqwNb+Rr469iz6RUcLvEOc2Pf3xl7/rvjFnmBv2xf1dHc5V/CH+Myf8T9sGFMHxvOHgZ9YGjwkHA+NT81cVGYooWQU6DIH/r3qx8KGHm/QUCIKKpzTGvSXqBJ2U/lsv/pnDtg1HsQioFEdLzSaIEnAfzrRzEtfPxq9djB28xJBz51mCx+0wnQpfGhzKak/K6iAgfAmU5xnYjHudHf3yLlLljTbV4pWXXPPknzb5MhEaEqYI5ZXEAvVICEpFNoHJx1Q6f+dDOZJ/pVfNGzyVmHYzwMBq78ZWUNcbCHKegP0U9Rye91KVv/Xr68DajBwnk6cyeYtJz7raPaK0Yooy8jdfqr07IVjrKS6dhRN7Sa+j+foBvoAJ4q8guGOnXFE6Qfem4YO/QeBYxEYKhkSDYpwF/I8L8Znsz7vRbjqh33+P0NWUzk+cMkp4aYsjSQwayXSeeuIS8bx1rgn+gswoPttgv2V3BbbBlNOxctAziCdjAYuzhhoE0yw3ybdOM9uH6KFZNyz/wNE0PndGABtWVACG14mJdvcnjUYKDwsPjJvj19PKd90+x1kctE19hLZVyqk+y723ZGb/J+xPjJPTBS9kbIZTbSAddfntrK2/13xmgnryIsqswgV+RpBx8KBxMUPAngyKC8j27aHj4VH+Qag8qtBvVm1k1E3b36NFJj1yvRHawgzgXP+Io0+9qoVDKpmxRdECGswCFoGeIL/MKQB9FGPR4AvFstTHqLJntr5Qdx7LeBjcPvT80DURNOZ5hUKxUaq18YpJ41Dd8UiIiHCbFHIBOJHeJe/WCW4V93ZyerHUESdcr29tz0NcN3aWwFYRcWSyiT0fPLwz1P0ypq0WezpX7k/zGZOwSRA9X4mxUUdtxHQDGoqIRupOgAuno/LkUTBAaI5EXTNPcjhwOyLa5V2OblesJ4avMXvHr06LyBmpysUBXTP9GT9k+pPK88RiCVthaVZxKK/lxKgKl+JNyE1r0WdtZzmyHKNMEqpazr2w/5/mubU4lZVfPX66cR/16fi0y0z2/DBpv6d80cY3OfXDoRiWGyBuWsHX0MlVPAO38sYPPO70voCoBwDzUleXDkYwPq8wwSuUk/R7A6y2LpGHdrlUbabhTElP1IQkRXri8CHUQmtxFNCv75eJoATk5xtZ0jomIP0PQALYgZL0Q2b3g6JfDwZyllS2ZE5JzOFCcQYLKyDBcmLdXPOYAZeCk8OlIEH84X6YbZyN13oYWPFVtgyj09Bvp50fJUt6BEOd/3e8PcDTqNsH5ppjHuKL2KkfkhHkHCEilOQ9cO0Le09rUBnzZQPdNWAR9jrnf1LqNwt198/961mzFw354ffQEaWTQdtqOLZ1a+pp3bCEkfEa/aFDY+4P3RacCb2SReQaNVh5Rmk7kufz/zRwrUZVUbOL21JzWyjk1FFPBHi/7Au2IrtwByko=";

    }
    public static class Databases{
        public static final String Company = "D_QCON";
        public static final String EngineeringDocument= "PRJ_DOC";
        public static final String ProjectWorkspace = "PRJ_FOLDER";
        public static final String EngineeringCRS = "PRJ_CRS";

    }
    public static class Bookmarks{


        public  static final JSONObject projectWorkspaceTypes() {
            JSONObject rtrn = new JSONObject();
            rtrn.put("DurDay", Integer.class);
            rtrn.put("DurHour", Double.class);
            return rtrn;
        }
        public  static final JSONObject projectWorkspace() {
            JSONObject rtrn = new JSONObject();
            rtrn.put("ProjectNo", "ccmPRJCard_code");
            rtrn.put("ProjectName", "ccmPRJCard_name");
            rtrn.put("ProjectName", "ccmPRJCard_name");
            rtrn.put("To", "To-Receiver");
            rtrn.put("Attention", "ObjectAuthors");
            rtrn.put("CC", "CC-Receiver");
            rtrn.put("JobNo", "JobNo");
            rtrn.put("TransmittalNo", "ObjectNumberExternal");
            rtrn.put("IssueDate", "");
            rtrn.put("Discipline", "");
            rtrn.put("Summary", "ccmTrmtSummary");
            rtrn.put("Notes", "ccmTrmtNotes");
            rtrn.put("DurDay", "ccmPrjProcDurDay");
            rtrn.put("DurHour", "ccmPrjProcDurHour");

            JSONObject ebks = engDocument();
            for (String ekey : ebks.keySet()) {
                for(int p=1;p<=50;p++){
                    String einx = ekey + (p <= 9 ? "0" : "") + p;
                    rtrn.put(einx, "");
                }
            }

            JSONObject dbks = distribution();
            for (String dkey : dbks.keySet()) {
                String dfkl = dbks.getString(dkey);
                for(int p=1;p<=5;p++){

                    String dinx = (p <= 9 ? "0" : "") + p;
                    rtrn.put(dkey + dinx, dfkl.replace("##", dinx));
                }
            }
            rtrn.put("IssueStatus", "");
            rtrn.put("IssueType", "ObjectType");

            rtrn.put("OrigndFullname", "");
            rtrn.put("OrigndJobTitle", "");
            rtrn.put("OrigndDate", "");

            rtrn.put("ApprvdFullname", "");
            rtrn.put("ApprvdJobTitle", "");
            rtrn.put("ApprvdDate", "");

            return rtrn;
        }

        public static final String DistributionMaster = "DistUser";
        public  static final JSONObject distribution() {
            JSONObject rtrn = new JSONObject();
            rtrn.put("DistUser", "ccmDistUser##");
            rtrn.put("DistPurpose", "ccmDistPurpose##");
            rtrn.put("DistDlvMethod", "ccmDistDlvMethod##");
            rtrn.put("DistDueDate", "");
            //rtrn.put("DistDueDate", "ccmDistDueDate##");
            return rtrn;
        }
        public static final String EngDocumentMaster = "DocNo";
        public  static final JSONObject engDocument() {
            JSONObject rtrn = new JSONObject();
            rtrn.put("DocNo", "ccmPrjDocNumber");
            rtrn.put("RevNo", "ccmPrjDocRevision");
            rtrn.put("ParentDoc", "");
            rtrn.put("Desc", "ObjectName");
            rtrn.put("Issue", "ccmPrjDocIssueStatus");
            rtrn.put("FileName", "ccmPrjDocFileName");
            //rtrn.put("FileName", "@EXPORT_FILE_NAME@");
            rtrn.put("Remarks", "");

            return rtrn;
        }

    }


}
