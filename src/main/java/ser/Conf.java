package ser;

import org.json.JSONObject;

import java.util.Date;
import java.util.List;

public class Conf {

    public static class ExcelTransmittalPaths {
        public static final String MainPath = "C:/tmp2/bulk/transmittal";
    }
    public static class ExcelTransmittalSheetIndex {
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

    public static class Descriptors{
        public static final String MainDocumentID = "MainDocumentReference";
        public static final String DocTransOutCode = "ccmPrjDocTransOutCode";
        public static final String FileName = "ccmPrjDocFileName";
        public static final String DocNumber = "ccmPrjDocNumber";
        public static final String DocRevision = "ccmPrjDocRevision";
        public static final String DocType = "ccmPrjDocDocType";
        public static final String ObjectName = "ObjectName";
        public static final String Originator = "ccmPrjDocOiginator";
        public static final String Prefix = "ccmTrmtSenderCode";
        public static final String Category = "ccmPrjDocCategory";
        public static final String ParentDocNumber = "ccmPrjDocParentDoc";
        public static final String ParentDocRevision = "ccmPrjDocParentDocRevision";
        public static final String ObjectNumberExternal = "ObjectNumberExternal";
        public static final String ClientNo = "ccmPrjDocClientPrjNumber";
        public static final String TrmtCounterPattern = "ccmTMCounterPattern";
        public static final String TrmtCounterStart = "ccmTMCounterStart";
        public static final String TrmtSendType = "ccmTrmtSendType";
        public static final String ProjectNo = "ccmPRJCard_code";
        public static final String ProjectName = "ccmPRJCard_name";
        public static final String DccList = "ccmPrjCard_DccList";
        public static final String ObjectNumber = "ObjectNumber";
        public static final String SenderCode = "ccmTrmtSenderCode";
        public static final String ReceiverCode = "ccmTrmtReceiverCode";
        public static final String SenderName = "ccmTrmtSender";
        public static final String ReceiverName = "ccmTrmtReceiver";

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
        public static final String ObjectType = "OBJECTTYPE";

    }
    public static class ClassIDs{
        public static final String Template = "b9cf43d1-a4d3-482f-9806-44ae64c6139d";
        public static final String EngineeringDocument = "3b3078f8-c0d0-4830-b963-88bebe1c1462";
        public static final String ProjectWorkspace = "32e74338-d268-484d-99b0-f90187240549";
        public static final String EngineeringProjectTransmittal = "8bf0a09b-b569-4aef-984b-78cf1644ca19";
        public static final String EngineeringCRS = "3e1fe7b3-3e86-4910-8155-c29b662e71d6";

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

            rtrn.put("IssueDate", Date.class);
            rtrn.put("ApprovedDate", Date.class);
            rtrn.put("OriginatedDate", Date.class);


            JSONObject dbks = distribution();
            JSONObject dbts = distributionTypes();

            for (String dkey : dbks.keySet()) {
                if(!dbts.has(dkey)){continue;}
                for(int p=1;p<=5;p++){
                    String dinx = (p <= 9 ? "0" : "") + p;
                    rtrn.put(dkey + dinx, dbts.get(dkey));
                }
            }

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
            rtrn.put("IssueDate", "DateStart");
            rtrn.put("Discipline", "");
            rtrn.put("Summary", "ccmTrmtSummary");
            rtrn.put("Notes", "ccmTrmtNotes");
            rtrn.put("DurDay", "ccmPrjProcDurDay");
            rtrn.put("DurHour", "ccmPrjProcDurHour");

            rtrn.put("Approved", "ccmApproved");
            rtrn.put("ApprovedDate", "ccmApprovedDate2");
            rtrn.put("Originated", "ccmOriginated");
            rtrn.put("OriginatedDate", "ccmOriginatedDate2");

            rtrn.put("SenderName", "ccmTrmtSender");
            rtrn.put("SenderCode", "ccmTrmtSenderCode");
            rtrn.put("ReceiverName", "ccmTrmtReceiver");
            rtrn.put("ReceiverCode", "ccmTrmtReceiverCode");

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
            rtrn.put("OrigndSignature", "");

            rtrn.put("ApprvdFullname", "");
            rtrn.put("ApprvdJobTitle", "");
            rtrn.put("ApprvdDate", "");
            rtrn.put("ApprvdSignature", "");

            return rtrn;
        }

        public static final String DistributionMaster = "DistUser";
        public  static final JSONObject distribution() {
            JSONObject rtrn = new JSONObject();
            rtrn.put("DistUser", "ccmDistUser##");
            rtrn.put("DistPurpose", "ccmDistPurpose##");
            rtrn.put("DistDlvMethod", "ccmDistDlvMethod##");
            rtrn.put("DistDueDate", "ccmDistDueDate##");
            return rtrn;
        }
        public  static final JSONObject distributionTypes() {
            JSONObject rtrn = new JSONObject();
            rtrn.put("DistDueDate", Date.class);
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
            rtrn.put("Remarks", "");

            return rtrn;
        }

    }


}
