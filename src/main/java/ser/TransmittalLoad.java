package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.commons.io.FilenameUtils;
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
            projectNo = processInstance.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            projectNo = (projectNo == null ? "" : projectNo.trim());
            if(projectNo.isEmpty()){
                throw new Exception("Project no is empty.");
            }
            projectInfObj = Utils.getProjectWorkspace(projectNo, helper);
            if(projectInfObj == null){
                throw new Exception("Project not found [" + projectNo + "].");
            }
            transmittalNr = Utils.getTransmittalNr(session, projectInfObj, processInstance);
            if(transmittalNr.isEmpty()){
                throw new Exception("Transmittal number not found.");
            }

            transmittalLinks = processInstance.getLoadedInformationObjectLinks();
            Utils.saveDuration(processInstance);

            String ctpn = "TRANSMITTAL_COVER";
            IDocument ctpl = Utils.getTemplateDocument(projectNo, ctpn, helper);
            if(ctpl == null){
                throw new Exception("Template-Document [ " + ctpn + " ] not found.");
            }
            String tplCoverPath = Utils.exportDocument(ctpl, exportPath, ctpn);

            transmittalDoc = (IDocument) processInstance.getMainInformationObject();
            if(transmittalDoc != null && !transmittalDoc.getDescriptorValue(Conf.Descriptors.Category, String.class).equals("Transmittal")){
                transmittalDoc = null;
            }
            boolean isTDocLinked = (transmittalDoc == null ? false : true);

            documentIds = Utils.getLinkedDocIds(transmittalLinks);

            processInstance.setDescriptorValue(Conf.Descriptors.ProjectNo,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class));
            processInstance.setDescriptorValue(Conf.Descriptors.ProjectName,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectName, String.class));
            processInstance.setDescriptorValue(Conf.Descriptors.DccList,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.DccList, String.class));

            processInstance = Utils.updateProcessInstance(processInstance);


            if(transmittalDoc == null) {
                transmittalDoc = Utils.createTransmittalDocument(session, server, projectInfObj);
            }

            bookmarks = Utils.loadBookmarks(session, server, transmittalNr, transmittalLinks,
                    linkedDocIds, documentIds, processInstance, transmittalDoc, exportPath, helper);
            transmittalDoc.setDescriptorValue(Conf.Descriptors.ObjectNumberExternal,
                    transmittalNr);
            transmittalDoc.setDescriptorValue(Conf.Descriptors.ProjectNo,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class));
            transmittalDoc.setDescriptorValue(Conf.Descriptors.ProjectName,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectName, String.class));
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DccList,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.DccList, String.class));
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocNumber,
                    transmittalNr);
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocRevision,
                    "");
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocType,
                    "Transmittal-Outgoing");
            transmittalDoc.setDescriptorValue(Conf.Descriptors.FileName,
                    "" + transmittalNr + ".pdf");
            transmittalDoc.setDescriptorValue(Conf.Descriptors.ObjectName,
                    "Transmittal Cover Page");
            transmittalDoc.setDescriptorValue(Conf.Descriptors.Category,
                    "Transmittal");
            transmittalDoc.setDescriptorValue(Conf.Descriptors.Originator,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.Prefix, String.class));

            transmittalDoc = Utils.updateDocument(transmittalDoc);

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

            List dstLines = Utils.excelDstTblLines(bookmarks);
            List docLines = Utils.excelDocTblLines(bookmarks);

            String coverExcelPath = Utils.saveTransmittalExcel(tplCoverPath, Conf.ExcelTransmittalSheetIndex.Cover,
                    exportPath + "/" + ctpn + ".xlsx", bookmarks, docLines, dstLines);

            Utils.addTransmittalRepresentations(transmittalDoc, exportPath, coverExcelPath, "", "");

            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocNumber,
                    transmittalNr);

            transmittalDoc = Utils.updateDocument(transmittalDoc);

            if(!isTDocLinked) {
                processInstance.setMainInformationObjectID(transmittalDoc.getID());
            }
            processInstance = Utils.updateProcessInstance(processInstance);


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