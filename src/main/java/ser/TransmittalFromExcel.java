package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.*;
import java.util.List;
import java.util.UUID;


public class TransmittalFromExcel extends UnifiedAgent {

    ISession ses;
    IDocumentServer srv;
    private ProcessHelper helper;
    @Override
    protected Object execute() {
        if (getEventDocument() == null)
            return resultError("Null Document object");

        ses = getSes();
        srv = ses.getDocumentServer();
        IDocument document = getEventDocument();

        try {

            this.helper = new ProcessHelper(getSes());
            (new File(Conf.ExcelTransmittalPaths.MainPath)).mkdir();

            String uniqueId = UUID.randomUUID().toString();
            String excelPath = FileEvents.fileExport(document, Conf.ExcelTransmittalPaths.MainPath, uniqueId);

            FileInputStream fist = new FileInputStream(excelPath);
            XSSFWorkbook fwrb = new XSSFWorkbook(fist);
            JSONObject data = Utils.getDataOfTransmittal(fwrb, Conf.ExcelTransmittalSheetIndex.FromExcel);
            if(data.get("ProjectNo") == null || data.get("ProjectNo") == ""){
                throw new Exception("Project no not found.");
            }

            IInformationObject prjt = Utils.getProjectWorkspace(data.get("ProjectNo").toString(), helper);
            if(prjt == null){
                throw new Exception("Project not found.");
            }

            IDocument tdoc = Utils.createTransmittalDocument(ses, srv, prjt);
            JSONObject xbks = Conf.Bookmarks.ProjectWorkspace();
            JSONObject pbks = new JSONObject();
            IProcessInstance proi = Utils.createEngineeringProjectTransmittal(tdoc, helper);

            List<JSONObject> dist = Utils.listOfDistributions(fwrb, Conf.ExcelTransmittalSheetIndex.FromExcel);
            int scnt = 0;
            for (JSONObject ldst : dist) {
                scnt++;
                if(ldst.get("user") == null){continue;}
                String slfx = ((scnt <= 9 ? "0" : "") + scnt);

                String distUser = xbks.getString("DistUser" + slfx);
                String distPurpose = xbks.getString("DistPurpose" + slfx);
                String distDlvMethod = xbks.getString("DistDlvMethod" + slfx);
                String distDueDate = xbks.getString("DistDueDate" + slfx);

                if(distUser.isEmpty() == false){
                    proi.setDescriptorValue(distUser, ldst.getString("user"));
                    //tdoc.setDescriptorValue(distUser, ldst.getString("user"));
                }
                if(distPurpose.isEmpty() == false){
                    proi.setDescriptorValue(distPurpose, ldst.getString("purpose"));
                    //tdoc.setDescriptorValue(distPurpose, ldst.getString("purpose"));
                }
                if(distDlvMethod.isEmpty() == false){
                    proi.setDescriptorValue(distDlvMethod, ldst.getString("dlvMethod"));
                    //tdoc.setDescriptorValue(distDlvMethod, ldst.getString("dlvMethod"));
                }
                if(distDueDate.isEmpty() == false){
                    //proi.setDescriptorValue(distDueDate, ldst.getString("dueDate"));
                    //tdoc.setDescriptorValue(distDueDate, ldst.getString("dueDate"));
                }
                pbks.put("DistUser" + slfx, ldst.getString("user"));
                pbks.put("DistPurpose" + slfx, ldst.getString("purpose"));
                pbks.put("DistDlvMethod" + slfx, ldst.getString("dlvMethod"));
                pbks.put("DistDueDate" + slfx, ldst.getString("dueDate"));


            }

            for (String pkey : xbks.keySet()) {
                String pfld = xbks.getString(pkey);
                if(pfld.isEmpty()){continue;}

                String dval = "";
                if(dval == "" && data.has(pkey)) {
                    dval = data.getString(pkey);
                }
                if(dval == "" && pbks.has(pkey)) {
                    dval = pbks.getString(pkey);
                }
                pbks.put(pkey, dval);
                if(dval.isEmpty()){continue;}

                proi.setDescriptorValue(pfld, dval);
                //tdoc.setDescriptorValue(pfld, dval);

            }
            String tmnr = proi.getDescriptorValue(Conf.Descriptors.ObjectNumberExternal, String.class);
            if(tmnr == null || tmnr == "") {
                tmnr = (new CounterHelper(ses, proi.getClassID())).getCounterStr();
            }
            pbks.put("TransmittalNo", tmnr);

            String tuss = Utils.getWorkbasketDisplayNames(ses, srv, proi.getDescriptorValue("To-Receiver", String.class));
            proi.setDescriptorValue("To-Receiver", tuss);
            //tdoc.setDescriptorValue("To-Receiver", tuss);

            String auss = Utils.getWorkbasketDisplayNames(ses, srv, proi.getDescriptorValue("ObjectAuthors", String.class));
            proi.setDescriptorValue("ObjectAuthors", auss);
            //tdoc.setDescriptorValue("ObjectAuthors", auss);

            String cuss = Utils.getWorkbasketDisplayNames(ses, srv, proi.getDescriptorValue("CC-Receiver", String.class));
            proi.setDescriptorValue("CC-Receiver", cuss);
            //tdoc.setDescriptorValue("CC-Receiver", cuss);

            proi.setDescriptorValue("ObjectNumberExternal", tmnr);
            //tdoc.setDescriptorValue("ObjectNumberExternal", tmnr);


            tdoc.setDescriptorValue(Conf.Descriptors.ObjectNumberExternal,
                    tmnr);
            tdoc.setDescriptorValue(Conf.Descriptors.ProjectNo,
                    prjt.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class));
            tdoc.setDescriptorValue(Conf.Descriptors.ProjectName,
                    prjt.getDescriptorValue(Conf.Descriptors.ProjectName, String.class));
            tdoc.setDescriptorValue(Conf.Descriptors.DccList,
                    prjt.getDescriptorValue(Conf.Descriptors.DccList, String.class));

            tdoc.setDescriptorValue(Conf.Descriptors.DocNumber, tmnr);
            tdoc.setDescriptorValue(Conf.Descriptors.DocRevision, "");
            tdoc.setDescriptorValue(Conf.Descriptors.DocType, "Transmittal-Outgoing");
            tdoc.setDescriptorValue(Conf.Descriptors.FileName, "" + tmnr + ".pdf");
            tdoc.setDescriptorValue(Conf.Descriptors.ObjectName, "Transmittal Cover Page");
            tdoc.setDescriptorValue(Conf.Descriptors.Category, "Transmittal");
            tdoc.setDescriptorValue(Conf.Descriptors.Originator,
                    prjt.getDescriptorValue(Conf.Descriptors.Prefix, String.class)
            );


            String tdId = tdoc.getID();
            tdoc.commit();
            Thread.sleep(2000);
            if(!tdId.equals("<new>")) {
                tdoc = srv.getDocument4ID(tdId, ses);
            }

            proi.setDescriptorValue(Conf.Descriptors.ProjectNo,
                    prjt.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class));

            proi.setDescriptorValue(Conf.Descriptors.ProjectName,
                    prjt.getDescriptorValue(Conf.Descriptors.ProjectName, String.class));

            proi.setDescriptorValue(Conf.Descriptors.DccList,
                    prjt.getDescriptorValue(Conf.Descriptors.DccList, String.class));

            String poId = proi.getID();
            Thread.sleep(2000);
            proi.commit();
            if(!poId.equals("<new>")) {
                proi = (IProcessInstance) srv.getInformationObjectByID(poId, ses);
            }



            document.setDescriptorValue("ccmPrjDocNumber", "Transmittal Excel [" + tmnr + "]");
            document.commit();

            /*
            ITask pitk = proi.getLoadedRootTask();

            INode tnod = Utils.getNode((IFolder) prjt, Conf.ExcelTransmittalNodes.FromExcelNodeName);
            IElements tels = tnod.getElements();
            if (tels.getItemByLink2(FMLinkType.TASK, pitk.getID()) == null){
                tels.addNew(FMLinkType.TASK).setLink(pitk.getID());
            }
            */
            prjt.commit();

            pbks.put("ProjectNo", prjt.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class));
            pbks.put("ProjectName", prjt.getDescriptorValue(Conf.Descriptors.ProjectName, String.class));

            IInformationObjectLinks links = proi.getLoadedInformationObjectLinks();

            List<JSONObject> docs = Utils.getListOfDocuments(fwrb);
            //int lcnt = 0;
            for (JSONObject ldoc : docs) {
                //lcnt++;
                if(ldoc.get("docNo") == null){continue;}
                if(ldoc.get("revNo") == null){continue;}

                IDocument edoc = Utils.getEngineeringDocument(ldoc.getString("docNo"), ldoc.getString("revNo"), helper);
                if(edoc == null){continue;}

                links.addInformationObject(edoc.getID());

            }

            links.addInformationObject(tdoc.getID());
            links.addInformationObject(document.getID());
            proi.commit();
            tdoc.commit();


        } catch (Exception e) {
            //throw new RuntimeException(e);
            System.out.println("Exception       : " + e.getMessage());
            System.out.println("    Class       : " + e.getClass());
            System.out.println("    Stack-Trace : " + e.getStackTrace() );
            return resultError("Exception : " + e.getMessage());
        }

        System.out.println("Finished");

        return resultSuccess("Ended successfully");
    }
}