package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONObject;

import java.io.*;
import java.util.*;


public class TransmittalLoad extends UnifiedAgent {
    Logger log = LogManager.getLogger();
    IInformationObjectLinks transmittalLinks;
    IProcessInstance processInstance;
    IInformationObject projectInfObj;
    IInformationObject contractorInfObj;
    ITask task;
    ProcessHelper helper;
    List<String> documentIds = new ArrayList<>();
    List<String> linkedDocIds = new ArrayList<>();
    IDocument transmittalDoc;

    String transmittalNr;
    String projectNo;
    JSONObject bookmarks;
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");

        if(getEventTask().getProcessInstance().findLockInfo().getOwnerID() != null){
            return resultRestart("Restarting Agent");
        }

        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();
        Utils.loadDirectory(Conf.Paths.MainPath);

        task = getEventTask();

        try {

            helper = new ProcessHelper(Utils.session);
            XTRObjects.setSession(Utils.session);

            String uniqueId = UUID.randomUUID().toString();
            String exportPath = Conf.Paths.MainPath + "/Transmittal[" + uniqueId + "]";
            (new File(exportPath)).mkdirs();

            processInstance = task.getProcessInstance();
            projectNo = (processInstance != null ? Utils.projectNr((IInformationObject) processInstance) : "");
            if(projectNo.isEmpty()){
                throw new Exception("Project no is empty.");
            }
            projectInfObj = Utils.getProjectWorkspace(projectNo, helper);
            if(projectInfObj == null){
                throw new Exception("Project not found [" + projectNo + "].");
            }

            String ivpNo = processInstance.getDescriptorValue(Conf.Descriptors.SenderCode, String.class);
            if(ivpNo == null || ivpNo.isEmpty()){
                throw new Exception("Involve Party code is empty.");
            }

            contractorInfObj = Utils.getContractorFolder(projectNo, ivpNo, helper);
            if(contractorInfObj == null){
                throw new Exception("Involve Party [" + projectNo + "/" + ivpNo + "].");
            }

            transmittalNr = Utils.getTransmittalNr(projectInfObj, processInstance);
            if(transmittalNr.isEmpty()){
                throw new Exception("Transmittal number not found.");
            }

            transmittalLinks = processInstance.getLoadedInformationObjectLinks();
            Utils.saveDuration(processInstance);

            String ctpn = "TRANSMITTAL_COVER";

            IDocument ctpl = null;
            ctpl = ctpl != null ? ctpl : Utils.getTemplateDocument(contractorInfObj, ctpn);
            ctpl = ctpl != null ? ctpl : Utils.getTemplateDocument(projectInfObj, ctpn);

            if(ctpl == null){
                throw new Exception("Template-Document [ " + ctpn + " ] not found.");
            }
            String tplCoverPath = Utils.exportDocument(ctpl, exportPath, ctpn);

            transmittalDoc = (IDocument) processInstance.getMainInformationObject();
            if(transmittalDoc != null && !transmittalDoc.getDescriptorValue(Conf.Descriptors.Category, String.class).equals("Correspondence")){
                transmittalDoc = null;
            }

            String docType = processInstance.getDescriptorValue(Conf.Descriptors.DocType, String.class);
            docType = (docType == null ? "" : docType);

            documentIds = Utils.getLinkedDocIds(transmittalLinks, docType);

            processInstance.setDescriptorValue(Conf.Descriptors.ProjectNo,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class));
            processInstance.setDescriptorValue(Conf.Descriptors.ProjectName,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectName, String.class));
            processInstance.setDescriptorValue(Conf.Descriptors.DccList,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.DccList, String.class));


            //processInstance = Utils.updateProcessInstance(processInstance);
            processInstance.commit();

            boolean isTDocLinked = (transmittalDoc == null ? false : true);
            if(transmittalDoc == null) {
                transmittalDoc = Utils.createTransmittalDocument(projectInfObj);
            }

            bookmarks = Utils.loadBookmarks(transmittalNr, transmittalLinks,
                    projectInfObj, contractorInfObj,
                    linkedDocIds, documentIds, processInstance, transmittalDoc, exportPath);

            transmittalDoc.setDescriptorValue(Conf.Descriptors.TransmittalNr,
                    transmittalNr);
            transmittalDoc.setDescriptorValue(Conf.Descriptors.ProjectNo,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class));
            transmittalDoc.setDescriptorValue(Conf.Descriptors.ProjectName,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectName, String.class));
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DccList,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.DccList, String.class));

            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocOriginator,
                    processInstance.getDescriptorValue(Conf.Descriptors.SenderCode, String.class));
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocSenderCode,
                    processInstance.getDescriptorValue(Conf.Descriptors.SenderCode, String.class));
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocReceiverCode,
                    processInstance.getDescriptorValue(Conf.Descriptors.ReceiverCode, String.class));
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocStatus,
                    "40");

            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocNumber,
                    transmittalNr);
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocRevision,
                    "");
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocType,
                    docType);

            String fileName = transmittalNr.replaceAll("[\\\\/:*?\"<>|]", "_") + ".pdf";
            transmittalDoc.setDescriptorValue(Conf.Descriptors.FileName,
                    fileName);
            transmittalDoc.setDescriptorValue(Conf.Descriptors.ObjectName,
                    "Transmittal Cover Page");
            transmittalDoc.setDescriptorValue(Conf.Descriptors.Category,
                    "Correspondence");

            //transmittalDoc = Utils.updateDocument(transmittalDoc);
            //transmittalDoc.commit();

            List dstLines = Utils.excelDstTblLines(bookmarks);
            List docLines = Utils.excelDocTblLines(bookmarks);

            String coverExcelPath = Utils.saveTransmittalExcel(tplCoverPath, Conf.ExcelTransmittalSheetIndex.Cover,
                    exportPath + "/" + ctpn + ".xlsx", bookmarks, docLines, dstLines);

            Utils.addTransmittalRepresentations(transmittalDoc, exportPath, coverExcelPath, "", "");

            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocNumber,
                    transmittalNr);

            //transmittalDoc = Utils.updateDocument(transmittalDoc);
            transmittalDoc.commit();


            ILink[] tlnks = Utils.server.getReferencedRelationships(Utils.session, transmittalDoc, true);
            JSONObject tlnkds = new JSONObject();
            for (ILink tlnk : tlnks) {
                IInformationObject ttgt = tlnk.getTargetInformationObject();
                tlnkds.put(ttgt.getID(), tlnk);
            }
            for(String lnkd : linkedDocIds){
                if(tlnkds.has(lnkd)){
                    tlnkds.remove(lnkd);
                    continue;
                }
                ILink lnk1 = Utils.server.createLink(Utils.session, transmittalDoc.getID(), null, lnkd);
                lnk1.commit();
            }

            if(!isTDocLinked) {
                processInstance.setMainInformationObjectID(transmittalDoc.getID());
            }
            //processInstance = Utils.updateProcessInstance(processInstance);
            processInstance.commit();

        } catch (Exception e) {
            //throw new RuntimeException(e);
            log.error("Exception       : " + e.getMessage());
            log.error("    Class       : " + e.getClass());
            log.error("    Stack-Trace : " + e.getStackTrace() );
            return resultError("Exception : " + e.getMessage());
        }

        return resultSuccess("Ended successfully");
    }
}