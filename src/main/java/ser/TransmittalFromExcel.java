package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.*;
import java.util.List;
import java.util.UUID;


public class TransmittalFromExcel extends UnifiedAgent {
    Logger log = LogManager.getLogger();
    ProcessHelper helper;
    IInformationObject projectInfObj;
    IProcessInstance processInstance;
    IDocument transmittalDoc;
    String transmittalNr;
    @Override
    protected Object execute() {
        if (getEventDocument() == null)
            return resultError("Null Document object");


        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();
        Utils.loadDirectory(Conf.Paths.MainPath);
        
        IDocument document = getEventDocument();

        try {

            helper = new ProcessHelper(getSes());
            XTRObjects.setSession(Utils.session);

            String uniqueId = UUID.randomUUID().toString();
            String excelPath = FileEvents.fileExport(document, Conf.Paths.MainPath, uniqueId);

            FileInputStream fist = new FileInputStream(excelPath);
            XSSFWorkbook fwrb = new XSSFWorkbook(fist);

            JSONObject ecfg = Utils.getExcelConfig(fwrb);

            JSONObject data = Utils.getDataOfTransmittal(fwrb, ecfg);
            if(data.get("ProjectNo") == null || data.get("ProjectNo") == ""){
                throw new Exception("Project no not found.");
            }
            projectInfObj = Utils.getProjectWorkspace(data.get("ProjectNo").toString(), helper);
            if(projectInfObj == null){
                throw new Exception("Project not found.");
            }

            JSONObject xbks = Conf.Bookmarks.projectWorkspace();
            processInstance = Utils.createEngineeringProjectTransmittal(helper);

            List<JSONObject> dist = Utils.getListOfDistributions(fwrb, ecfg);
            int scnt = 0;
            for (JSONObject ldst : dist) {
                scnt++;
                if(ldst.get("User") == null){continue;}
                String slfx = ((scnt <= 9 ? "0" : "") + scnt);

                String distUser = xbks.getString("DistUser" + slfx);
                String distPurpose = xbks.getString("DistPurpose" + slfx);
                String distDlvMethod = xbks.getString("DistDlvMethod" + slfx);
                String distDueDate = xbks.getString("DistDueDate" + slfx);

                if(distUser.isEmpty() == false){
                    processInstance.setDescriptorValue(distUser, ldst.getString("User"));
                    //transmittalDoc.setDescriptorValue(distUser, ldst.getString("User"));
                }
                if(distPurpose.isEmpty() == false){
                    processInstance.setDescriptorValue(distPurpose, ldst.getString("Purpose"));
                }
                if(distDlvMethod.isEmpty() == false){
                    processInstance.setDescriptorValue(distDlvMethod, ldst.getString("DlvMethod"));
                }
                if(distDueDate.isEmpty() == false){
                    //processInstance.setDescriptorValue(distDueDate, ldst.getString("DueDate"));
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

            processInstance.setDescriptorValue(Conf.Descriptors.ProjectNo,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class));
            processInstance.setDescriptorValue(Conf.Descriptors.ProjectName,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectName, String.class));

            transmittalNr = Utils.getTransmittalNr(projectInfObj, processInstance);

            if(transmittalNr.isEmpty()){
                throw new Exception("Transmittal number not found.");
            }

            processInstance.setDescriptorValue(Conf.Descriptors.DccList,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.DccList, String.class));
            processInstance.setDescriptorValue("To-Receiver",
                    Utils.getWorkbasketDisplayNames(processInstance.getDescriptorValue("To-Receiver", String.class)));
            processInstance.setDescriptorValue("ObjectAuthors",
                    Utils.getWorkbasketDisplayNames(processInstance.getDescriptorValue("ObjectAuthors", String.class)));
            processInstance.setDescriptorValue("CC-Receiver",
                    Utils.getWorkbasketDisplayNames(processInstance.getDescriptorValue("CC-Receiver", String.class)));


            document.setDescriptorValue("ccmPrjDocNumber", transmittalNr + "/Import-Excel");
            document.commit();

            IInformationObjectLinks links = processInstance.getLoadedInformationObjectLinks();

            List<JSONObject> docs = Utils.getListOfDocuments(fwrb, ecfg);
            for (JSONObject ldoc : docs) {
                if(!ldoc.has("DocNo")
                || ldoc.getString("DocNo") == null
                || ldoc.getString("DocNo").isEmpty()){continue;}

                if(!ldoc.has("RevNo")
                || ldoc.getString("RevNo") == null
                || ldoc.getString("RevNo").isEmpty()){continue;}

                IDocument edoc = Utils.getEngineeringDocument(ldoc.getString("DocNo"), ldoc.getString("RevNo"), helper);
                if(edoc == null){continue;}

                links.addInformationObject(edoc.getID());
            }

            //processInstance = Utils.updateProcessInstance(processInstance);

            processInstance.commit();
            ILink lnk1 = Utils.server.createLink(Utils.session, processInstance.getID(), null, document.getID());
            lnk1.commit();

            Utils.addToNode(projectInfObj, "Transmittal From Excel", document);


        } catch (Exception e) {
            //throw new RuntimeException(e);
            log.error("Exception       : " + e.getMessage());
            log.error("    Class       : " + e.getClass());
            log.error("    Stack-Trace : " + e.getStackTrace() );
            return resultError("Exception : " + e.getMessage());
        }

        log.info("Finished");

        return resultSuccess("Ended successfully");
    }
}