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

        com.spire.license.LicenseProvider.setLicenseKey(Conf.Licences.SPIRE_XLS);

        session = getSes();
        bpm = getBpm();
        server = session.getDocumentServer();
        task = getEventTask();

        try {

            helper = new ProcessHelper(session);
            (new File(Conf.ExcelTransmittalPaths.MainPath)).mkdir();

            String uniqueId = UUID.randomUUID().toString();
            String exportPath = Conf.ExcelTransmittalPaths.MainPath + "/Transmittal[" + uniqueId + "]";
            (new File(exportPath)).mkdir();


            processInstance = task.getProcessInstance();
            transmittalLinks = processInstance.getLoadedInformationObjectLinks();

            transmittalNr = processInstance.getDescriptorValue(Conf.Descriptors.ObjectNumberExternal, String.class);

            if(transmittalNr == null || transmittalNr == "") {
                transmittalNr = (new CounterHelper(session, processInstance.getClassID())).getCounterStr();
            }

            projectNo = processInstance.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            if(projectNo.isEmpty()){
                throw new Exception("Project no is empty.");
            }
            projectInfObj = Utils.getProjectWorkspace(projectNo, helper);
            if(projectInfObj == null){
                throw new Exception("Project not found [" + projectNo + "].");
            }

            Utils.saveDuration(processInstance);

            String ctpn = "TRANSMITTAL_COVER";
            IDocument ctpl = Utils.getTemplateDocument(projectNo, ctpn, helper);
            if(ctpl == null){
                throw new Exception("Template-Document [ " + ctpn + " ] not found.");
            }
            String tplCoverPath = Utils.exportDocument(ctpl, exportPath, ctpn);

            transmittalDoc = null;
            boolean isTDocLinked = false;

            for (ILink link : transmittalLinks.getLinks()) {
                IDocument xdoc = (IDocument) link.getTargetInformationObject();
                if(!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}

                String dtyp = xdoc.getDescriptorValue(Conf.Descriptors.DocType, String.class);
                dtyp = (dtyp == null ? "" : dtyp);

                if(!dtyp.equals("Transmittal-Outgoing")) {
                    if(!documentIds.contains(xdoc.getID())) {
                        documentIds.add(xdoc.getID());
                    }
                    continue;
                }

                String docn = xdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                String docr = xdoc.getDescriptorValue(Conf.Descriptors.DocRevision, String.class);
                docn = (docn == null ? "" : docn);
                docr = (docr == null ? "" : docr);

                if(!docn.equals(transmittalNr) || !docr.equals("")) {
                    continue;
                }
                if(transmittalDoc == null){
                    transmittalDoc = xdoc;
                    isTDocLinked = true;
                    continue;
                }
                server.deleteDocument(session, xdoc);

            }


            processInstance.setDescriptorValue(Conf.Descriptors.ObjectNumberExternal,
                    transmittalNr);
            processInstance.setDescriptorValue(Conf.Descriptors.ProjectNo,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class));
            processInstance.setDescriptorValue(Conf.Descriptors.ProjectName,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectName, String.class));
            processInstance.setDescriptorValue(Conf.Descriptors.DccList,
                    projectInfObj.getDescriptorValue(Conf.Descriptors.DccList, String.class));

            String poId = processInstance.getID();
            Thread.sleep(2000);
            processInstance.commit();
            processInstance = (IProcessInstance) server.getInformationObjectByID(poId, session);


            Integer lcnt = 0;

            List<String> expFilePaths = new ArrayList<>();

            if(transmittalDoc == null) {
                transmittalDoc = Utils.createTransmittalDocument(session, server, projectInfObj);
            }

            for (ILink link : transmittalLinks.getLinks()) {
                IDocument edoc = (IDocument) link.getTargetInformationObject();
                if(!edoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
                if(linkedDocIds.contains(edoc.getID())){continue;}
                if(!documentIds.contains(edoc.getID())){continue;}


                String docNo = edoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                docNo = docNo == null ? "" : docNo;

                String fileName = edoc.getDescriptorValue(Conf.Descriptors.FileName, String.class);
                fileName = (fileName == null ? "" : fileName);
                if(fileName.isEmpty()){continue;}

                String expPath = Utils.exportDocument(edoc, exportPath, FilenameUtils.removeExtension(fileName));
                if(expFilePaths.contains(expPath)){continue;}

                lcnt++;
                System.out.println("IDOC [" + lcnt + "] *** " + edoc.getID());
                //String llfx = (lcnt <= 9 ? "0" : "") + lcnt;

                if(docNo.isEmpty()){continue;}

                expFilePaths.add(expPath);

                edoc.setDescriptorValue(Conf.Descriptors.DocTransOutCode, transmittalNr);
                edoc.commit();
                linkedDocIds.add(edoc.getID());

                IDocument cdoc = (IDocument) Utils.getEngineeringCRS(edoc.getID(), helper);
                if(cdoc != null){
                    String crsNo = cdoc.getDescriptorValue(Conf.Descriptors.ObjectNumber, String.class);
                    if(!crsNo.isEmpty()){
                        String crsPath = Utils.exportDocument(cdoc, exportPath,
                                FilenameUtils.removeExtension(fileName) + "_" + crsNo);
                        expFilePaths.add(crsPath);
                    }
                }
            }

            bookmarks = Utils.loadBookmarks(session, server, transmittalNr, transmittalLinks,
                    linkedDocIds, documentIds, processInstance, transmittalDoc);
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

            String transmittalDocId = transmittalDoc.getID();
            transmittalDoc.commit();
            Thread.sleep(2000);
            if(!transmittalDocId.equals("<new>")) {
                transmittalDoc = server.getDocument4ID(transmittalDocId, session);
            }

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

            String zipPath = Utils.zipFiles(exportPath + "/Blobs.zip", "", expFilePaths);
            Utils.addTransmittalRepresentations(transmittalDoc, exportPath, coverExcelPath, "", zipPath);

            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocNumber,
                    transmittalNr);

            if(!isTDocLinked) {
                transmittalLinks.addInformationObject(transmittalDoc.getID());
            }
            processInstance.commit();
            transmittalDoc.commit();


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