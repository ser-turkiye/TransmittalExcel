package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import com.ser.blueline.lock.ILockInfo;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONObject;

import java.io.File;
import java.util.*;


public class TransmittalSend extends UnifiedAgent {
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
            return resultRestart("Restarting Agent - Task.ProcessInstance");
        }

        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();
        Utils.loadDirectory(Conf.Paths.MainPath);

        task = getEventTask();

        try {

            helper = new ProcessHelper(Utils.session);
            XTRObjects.setSession(Utils.session);
            
            JSONObject scfg = Utils.getSystemConfig();
            if(scfg.has("LICS.SPIRE_XLS")){
                com.spire.license.LicenseProvider.setLicenseKey(scfg.getString("LICS.SPIRE_XLS"));
            }

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
                throw new Exception("Involve Party code is empty..");
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
            IDocument tmExcelDoc = (IDocument) processInstance.getMainInformationObject();
            if(tmExcelDoc == null) {
                throw new Exception("Transmittal-Excel-Document not found.");
            }

            ILockInfo cout = tmExcelDoc.getCheckOutInfo();
            if(cout != null && cout.getOwnerID() != null){
                return resultRestart("Restarting Agent - Locked TMExcel Document");
            }


            String mtpn = "TRANSMITTAL_MAIL";
            IDocument mtpl = null;
            mtpl = mtpl != null ? mtpl : Utils.getTemplateDocument(contractorInfObj, mtpn);
            mtpl = mtpl != null ? mtpl : Utils.getTemplateDocument(projectInfObj, mtpn);
            if(mtpl == null){
                throw new Exception("Template-Document [ " + mtpn + " ] not found.");
            }

            String tplMailPath = Utils.exportDocument(mtpl, exportPath, mtpn);

            String docType = processInstance.getDescriptorValue(Conf.Descriptors.DocType, String.class);
            docType = (docType == null ? "" : docType);

            documentIds = Utils.getLinkedDocIds(transmittalLinks, docType);
            Utils.saveDuration(processInstance);

            String ctpn = "TRANSMITTAL_COVER";
            String coverExcelPath = "";
            if(coverExcelPath.isEmpty()){
                coverExcelPath = Utils.getTransmittalReprExport(tmExcelDoc, ".xlsx", "Cover_Excel",
                        exportPath , ctpn);
            }
            if(coverExcelPath.isEmpty()){
                coverExcelPath = Utils.getTransmittalReprExport(tmExcelDoc, ".xlsx", "",
                        exportPath , ctpn);
            }
            if(coverExcelPath.isEmpty()){
                throw new Exception("Transmittal-Cover Excel not found.");
            }

            transmittalDoc = Utils.createTransmittalDocument(null);
            transmittalDoc = Utils.server.copyDocument2(Utils.session, tmExcelDoc, transmittalDoc, CopyScope.COPY_DESCRIPTORS);

            String zipPath = "";
            if(zipPath.isEmpty()) {
                zipPath = Utils.getTransmittalReprExport(transmittalDoc, ".zip", "Eng_Documents",
                        exportPath, "Blobs");
            }
            if(zipPath.isEmpty()) {
                zipPath = Utils.getTransmittalReprExport(transmittalDoc, ".zip", "",
                        exportPath, "Blobs");
            }
            if(zipPath.isEmpty()) {
                zipPath = Utils.getZipFile(transmittalLinks, exportPath, transmittalNr,
                        documentIds, helper);
            }

            String pdfPath = Utils.convertExcelToPdf(coverExcelPath, exportPath + "/" + ctpn + ".pdf");

            Utils.addTransmittalRepresentations(transmittalDoc, exportPath, "", pdfPath, zipPath);

            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocOriginator,
                    processInstance.getDescriptorValue(Conf.Descriptors.SenderCode, String.class));
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocSenderCode,
                    processInstance.getDescriptorValue(Conf.Descriptors.SenderCode, String.class));
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocReceiverCode,
                    processInstance.getDescriptorValue(Conf.Descriptors.ReceiverCode, String.class));
            transmittalDoc.setDescriptorValue(Conf.Descriptors.DocStatus,"50");
            transmittalDoc.setDescriptorValue(Conf.Descriptors.ApprCode, "N/A");

            transmittalDoc.commit();

            try {
                String sendType = processInstance.getDescriptorValue(Conf.Descriptors.TrmtSendType, String.class);
                sendType = (sendType == null ? "" : sendType);

                bookmarks = Utils.loadBookmarks(transmittalNr, transmittalLinks,
                        projectInfObj, contractorInfObj,
                        linkedDocIds, documentIds, processInstance, transmittalDoc, exportPath);

                bookmarks.put("DoxisLink", "");
                if(sendType.contains("URL")) {
                    JSONObject mcfg = Utils.getMailConfig();
                    bookmarks.put("DoxisLink", mcfg.getString("webBase") + helper.getDocumentURL(transmittalDoc.getID()));
                }

                List dstLines = Utils.excelDstTblLines(bookmarks);
                List docLines = Utils.excelDocTblLines(bookmarks);

                String mailExcelPath = Utils.saveTransmittalExcel(tplMailPath, Conf.ExcelTransmittalSheetIndex.Mail,
                        exportPath + "/" + mtpn + ".xlsx", bookmarks, docLines, dstLines);
                String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath, exportPath + "/" + mtpn + ".html");

                JSONObject mail = new JSONObject();

                List<String> stos = processInstance.getDescriptorValues("To-Receiver", String.class);
                List<String> sc1s = processInstance.getDescriptorValues("ObjectAuthors", String.class);
                List<String> sc2s = processInstance.getDescriptorValues("CC-Receiver", String.class);

                String mtos = Utils.getWorkbasketEMails(
                        stos == null || stos.size() == 0 ? "" : String.join(";", stos));
                String mc1s = Utils.getWorkbasketEMails(
                        sc1s == null || sc1s.size() == 0 ? "" : String.join(";", sc1s));
                String mc2s = Utils.getWorkbasketEMails(
                        sc2s == null || sc2s.size() == 0 ? "" : String.join(";", sc2s));

                mail.put("To", mtos);
                mail.put("CC", mc1s + (mc1s != "" && mc2s != "" ? ";" : "") + mc2s);
                mail.put("Subject", "Transmittal - " + transmittalNr);

                List<String> aths = new ArrayList<>();
                if (!pdfPath.isEmpty() && sendType.contains("COVER")) {
                    aths.add(pdfPath);
                }
                if (!zipPath.isEmpty() && sendType.contains("ZIP")) {
                    aths.add(zipPath);
                }

                mail.put("AttachmentPaths", String.join(";", aths));
                if (sendType.contains("COVER")) {
                    mail.put("AttachmentName." + java.nio.file.Paths.get(pdfPath).getFileName().toString(), "Cover_Preview[" + transmittalNr + "].pdf");
                }
                if (sendType.contains("ZIP")) {
                    mail.put("AttachmentName." + java.nio.file.Paths.get(zipPath).getFileName().toString(), "Eng_Documents[" + transmittalNr + "].zip");
                }

                mail.put("BodyHTMLFile", mailHtmlPath);
                Utils.sendHTMLMail(mail);
            }
            catch(Exception ex){
                log.error("EXCP [Send-Mail] : " + ex.getMessage());
            }

            processInstance.setMainInformationObjectID(transmittalDoc.getID());
            processInstance.commit();
            Utils.server.deleteDocument(Utils.session, tmExcelDoc);

            log.info("Finished");

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