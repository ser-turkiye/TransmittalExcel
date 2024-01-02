package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.json.JSONObject;

import java.io.*;
import java.util.*;


public class TransmittalLoad extends UnifiedAgent {

    ISession session;
    IDocumentServer server;
    IBpmService bpm;
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

        session = getSes();
        bpm = getBpm();
        server = session.getDocumentServer();
        task = getEventTask();

        try {

            helper = new ProcessHelper(session);
            (new File(Conf.ExcelTransmittalPaths.MainPath)).mkdirs();

            XTRObjects.setSession(session);

            String uniqueId = UUID.randomUUID().toString();
            String exportPath = Conf.ExcelTransmittalPaths.MainPath + "/Transmittal[" + uniqueId + "]";
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

            transmittalNr = Utils.getTransmittalNr(session, projectInfObj, processInstance);
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
                transmittalDoc = Utils.createTransmittalDocument(session, server, projectInfObj);
            }

            bookmarks = Utils.loadBookmarks(session, server, transmittalNr, transmittalLinks,
                    projectInfObj, contractorInfObj,
                    linkedDocIds, documentIds, processInstance, transmittalDoc, exportPath, helper);

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
                    "50");

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
            transmittalDoc.setDescriptorValue(Conf.Descriptors.ApprCode,
                    "N/A");
            transmittalDoc.setDescriptorValue(Conf.Descriptors.Originator,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.Prefix, String.class));

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


            ILink[] tlnks = server.getReferencedRelationships(session, transmittalDoc, true);
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
                ILink lnk1 = server.createLink(session, transmittalDoc.getID(), null, lnkd);
                lnk1.commit();
            }

            if(!isTDocLinked) {
                processInstance.setMainInformationObjectID(transmittalDoc.getID());
            }
            //processInstance = Utils.updateProcessInstance(processInstance);
            processInstance.commit();

        } catch (Exception e) {
            //throw new RuntimeException(e);
            System.out.println("Exception       : " + e.getMessage());
            System.out.println("    Class       : " + e.getClass());
            System.out.println("    Stack-Trace : " + e.getStackTrace() );
            System.out.println("    Cause is : " + e.getCause() );

            return resultError("Exception : " + e.getMessage());
        }

        return resultSuccess("Ended successfully");
    }
}