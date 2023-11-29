package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.*;
import java.util.List;
import java.util.UUID;


public class TransmittalFromExcel extends UnifiedAgent {

    ISession session;
    IDocumentServer server;
    ProcessHelper helper;
    IInformationObject projectInfObj;
    IProcessInstance processInstance;
    IDocument transmittalDoc;
    String transmittalNr;
    @Override
    protected Object execute() {
        if (getEventDocument() == null)
            return resultError("Null Document object");

        session = getSes();
        server = session.getDocumentServer();
        IDocument document = getEventDocument();

        try {

            helper = new ProcessHelper(getSes());
            (new File(Conf.ExcelTransmittalPaths.MainPath)).mkdir();

            String uniqueId = UUID.randomUUID().toString();
            String excelPath = FileEvents.fileExport(document, Conf.ExcelTransmittalPaths.MainPath, uniqueId);

            FileInputStream fist = new FileInputStream(excelPath);
            XSSFWorkbook fwrb = new XSSFWorkbook(fist);

            JSONObject data = Utils.getDataOfTransmittal(fwrb, Conf.ExcelTransmittalSheetIndex.FromExcel);
            if(data.get("ProjectNo") == null || data.get("ProjectNo") == ""){
                throw new Exception("Project no not found.");
            }

            IInformationObject projectInfObj = Utils.getProjectWorkspace(data.get("ProjectNo").toString(), helper);
            if(projectInfObj == null){
                throw new Exception("Project not found.");
            }

            JSONObject xbks = Conf.Bookmarks.projectWorkspace();
            processInstance = Utils.createEngineeringProjectTransmittal(helper);

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
                    processInstance.setDescriptorValue(distUser, ldst.getString("user"));
                    //transmittalDoc.setDescriptorValue(distUser, ldst.getString("user"));
                }
                if(distPurpose.isEmpty() == false){
                    processInstance.setDescriptorValue(distPurpose, ldst.getString("purpose"));
                }
                if(distDlvMethod.isEmpty() == false){
                    processInstance.setDescriptorValue(distDlvMethod, ldst.getString("dlvMethod"));
                }
                if(distDueDate.isEmpty() == false){
                    //processInstance.setDescriptorValue(distDueDate, ldst.getString("dueDate"));
                }
            }

            for (String pkey : xbks.keySet()) {
                String pfld = xbks.getString(pkey);
                if(pfld.isEmpty()){continue;}

                String dval = "";
                if(dval == "" && data.has(pkey)) {
                    dval = data.getString(pkey);
                }
                if(dval.isEmpty()){continue;}

                processInstance.setDescriptorValue(pfld, dval);
            }
            transmittalNr = processInstance.getDescriptorValue(Conf.Descriptors.ObjectNumberExternal, String.class);
            if(transmittalNr == null || transmittalNr == "") {
                transmittalNr = (new CounterHelper(session, processInstance.getClassID())).getCounterStr();
            }

            processInstance.setDescriptorValue("ObjectNumberExternal",
                    transmittalNr);

            processInstance.setDescriptorValue(Conf.Descriptors.ProjectNo,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class));
            processInstance.setDescriptorValue(Conf.Descriptors.ProjectName,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectName, String.class));
            processInstance.setDescriptorValue(Conf.Descriptors.DccList,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.DccList, String.class));

            processInstance.setDescriptorValue("To-Receiver",
                    Utils.getWorkbasketDisplayNames(session, server, processInstance.getDescriptorValue("To-Receiver", String.class)));
            processInstance.setDescriptorValue("ObjectAuthors",
                    Utils.getWorkbasketDisplayNames(session, server, processInstance.getDescriptorValue("ObjectAuthors", String.class)));
            processInstance.setDescriptorValue("CC-Receiver",
                    Utils.getWorkbasketDisplayNames(session, server, processInstance.getDescriptorValue("CC-Receiver", String.class)));


            document.setDescriptorValue("ccmPrjDocNumber", transmittalNr + "/Import-Excel");
            document.commit();

            IInformationObjectLinks links = processInstance.getLoadedInformationObjectLinks();

            List<JSONObject> docs = Utils.getListOfDocuments(fwrb);
            for (JSONObject ldoc : docs) {
                if(ldoc.get("docNo") == null){continue;}
                if(ldoc.get("revNo") == null){continue;}

                IDocument edoc = Utils.getEngineeringDocument(ldoc.getString("docNo"), ldoc.getString("revNo"), helper);
                if(edoc == null){continue;}

                links.addInformationObject(edoc.getID());
            }

            processInstance = Utils.updateProcessInstance(processInstance);

            ILink lnk1 = server.createLink(session, processInstance.getID(), null, document.getID());
            lnk1.commit();

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