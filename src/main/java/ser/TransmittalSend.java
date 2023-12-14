package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.commons.io.FilenameUtils;
import org.json.JSONObject;

import java.io.File;
import java.nio.file.Paths;
import java.util.*;


public class TransmittalSend extends UnifiedAgent {

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
            return resultRestart("Restarting Agent - Task.ProcessInstance");
        }

        session = getSes();
        bpm = getBpm();
        server = session.getDocumentServer();
        task = getEventTask();

        try {
            JSONObject scfg = Utils.getSystemConfig(session);
            if(scfg.has("LICS.SPIRE_XLS")){
                com.spire.license.LicenseProvider.setLicenseKey(scfg.getString("LICS.SPIRE_XLS"));
            }

            helper = new ProcessHelper(session);
            (new File(Conf.ExcelTransmittalPaths.MainPath)).mkdir();

            String uniqueId = UUID.randomUUID().toString();
            String exportPath = Conf.ExcelTransmittalPaths.MainPath + "/Transmittal[" + uniqueId + "]";
            (new File(exportPath)).mkdir();


            processInstance = task.getProcessInstance();
            transmittalLinks = processInstance.getLoadedInformationObjectLinks();

            transmittalNr = processInstance.getDescriptorValue(Conf.Descriptors.ObjectNumberExternal, String.class);
            if(transmittalNr == null || transmittalNr.isEmpty()) {
                throw new Exception("Transmittal no is empty.");
            }

            projectNo = processInstance.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            if(projectNo.isEmpty()){
                throw new Exception("Project no is empty.");
            }
            projectInfObj = Utils.getProjectWorkspace(projectNo, helper);
            if(projectInfObj == null){
                throw new Exception("Project not found [" + projectNo + "].");
            }

            IDocument tmExcelDoc = (IDocument) processInstance.getMainInformationObject();
            if(tmExcelDoc == null) {
                throw new Exception("Transmittal-Excel-Document not found.");
            }

            if(tmExcelDoc.getCheckOutInfo().getOwnerID() != null){
                return resultRestart("Restarting Agent - Locked TMExcel Document");
            }


            String mtpn = "TRANSMITTAL_MAIL";
            IDocument mtpl = Utils.getTemplateDocument(projectNo, mtpn, helper);
            if(mtpl == null){
                throw new Exception("Template-Document [ " + mtpn + " ] not found.");
            }
            String tplMailPath = Utils.exportDocument(mtpl, exportPath, mtpn);


            documentIds = Utils.getLinkedDocIds(transmittalLinks);
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

            transmittalDoc = Utils.createTransmittalDocument(session, server, null);
            transmittalDoc = server.copyDocument2(session, tmExcelDoc, transmittalDoc, CopyScope.COPY_DESCRIPTORS);

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
            //Utils.removeTransmittalRepresentations(transmittalDoc, ".xlsx");
            Utils.addTransmittalRepresentations(transmittalDoc, exportPath, "", pdfPath, zipPath);

            transmittalDoc = Utils.updateDocument(transmittalDoc);

            processInstance = Utils.updateProcessInstance(processInstance);

            String sendType = processInstance.getDescriptorValue(Conf.Descriptors.TrmtSendType, String.class);
            sendType = (sendType == null ? "" : sendType);


            bookmarks = Utils.loadBookmarks(session, server, transmittalNr, transmittalLinks,
                    linkedDocIds, documentIds, processInstance, transmittalDoc, exportPath, helper);

            bookmarks.put("DoxisLink", "");
            if(sendType.contains("URL")) {
                JSONObject mcfg = Utils.getMailConfig(session);
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

            String mtos = Utils.getWorkbasketEMails(session, server, bpm, String.join(";", stos));
            String mc1s = Utils.getWorkbasketEMails(session, server, bpm, String.join(";", sc1s));
            String mc2s = Utils.getWorkbasketEMails(session, server, bpm, String.join(";", sc2s));

            mail.put("To", mtos);
            mail.put("CC", mc1s + (mc1s != "" && mc2s != "" ? ";" : "") + mc2s);
            mail.put("Subject", "Transmittal - " + transmittalNr);

            List<String> aths = new ArrayList<>();
            if(!pdfPath.isEmpty() && sendType.contains("COVER")){
                aths.add(pdfPath);
            }
            if(!zipPath.isEmpty() && sendType.contains("ZIP")){
                aths.add(zipPath);
            }

            mail.put("AttachmentPaths", String.join(";", aths));
            if(sendType.contains("COVER")) {
                mail.put("AttachmentName." + Paths.get(pdfPath).getFileName().toString(), "Cover_Preview[" + transmittalNr + "].pdf");
            }
            if(sendType.contains("ZIP")) {
                mail.put("AttachmentName." + Paths.get(zipPath).getFileName().toString(), "Eng_Documents[" + transmittalNr + "].zip");
            }

            mail.put("BodyHTMLFile", mailHtmlPath);
            try{
                Utils.sendHTMLMail(session, mail);
            } catch (Exception ex){
                System.out.println("EXCP [Send-Mail] : " + ex.getMessage());
            }

            processInstance.setMainInformationObjectID(transmittalDoc.getID());
            server.deleteDocument(session, tmExcelDoc);

            System.out.println("Finished");

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