package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.commons.io.FilenameUtils;
import org.json.JSONObject;

import java.io.*;
import java.nio.file.Paths;
import java.util.*;
import java.util.concurrent.TimeUnit;


public class TransmittalLoad extends UnifiedAgent {

    ISession ses;
    IDocumentServer srv;
    IBpmService bpm;
    private ProcessHelper helper;
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");

        if(getEventTask().getProcessInstance().findLockInfo().getOwnerID() != null){
            return resultRestart("Restarting Agent");
        }

        com.spire.license.LicenseProvider.setLicenseKey(Conf.Licences.SPIRE_XLS);

        ses = getSes();
        srv = ses.getDocumentServer();
        bpm = getBpm();
        ITask task = getEventTask();

        try {

            this.helper = new ProcessHelper(ses);
            (new File(Conf.ExcelTransmittalPaths.MainPath)).mkdir();

            String uniqueId = UUID.randomUUID().toString();
            String exportPath = Conf.ExcelTransmittalPaths.MainPath + "/Transmittal[" + uniqueId + "]";
            (new File(exportPath)).mkdir();


            IProcessInstance proi = task.getProcessInstance();
            String tmnr = proi.getDescriptorValue(Conf.Descriptors.ObjectNumberExternal, String.class);

            if(tmnr == null || tmnr == "") {
                tmnr = (new CounterHelper(ses, proi.getClassID())).getCounterStr();
            }

            String prjn = proi.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            if(prjn.isEmpty()){
                throw new Exception("Project no is empty.");
            }
            IInformationObject prjt = Utils.getProjectWorkspace(prjn, helper);
            if(prjt == null){
                throw new Exception("Project not found [" + prjn + "].");
            }

            Collection<ITask> tsks = proi.findTasks();

            Date tbgn = null, tend = null;
            for(ITask ttsk : tsks){
                if(ttsk.getCreationDate() != null
                        && (tbgn == null  || tbgn.after(ttsk.getCreationDate()))){
                    tbgn = ttsk.getCreationDate();
                }
                if(ttsk.getFinishedDate() != null
                        && (tend == null  || tend.before(ttsk.getFinishedDate()))){
                    tend = ttsk.getFinishedDate();
                }
            }

            long durd  = 0L;
            double durh  = 0.0;
            if(tend != null && tbgn != null) {
                proi.setDescriptorValueTyped("ccmPrjProcStart", tbgn);
                proi.setDescriptorValueTyped("ccmPrjProcFinish", tend);

                long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
            }

            proi.setDescriptorValueTyped("ccmPrjProcDurDay", Integer.valueOf(durd + ""));
            proi.setDescriptorValueTyped("ccmPrjProcDurHour", durh );


            String ctpn = "TRANSMITTAL_COVER";
            IDocument ctpl = Utils.getTemplateDocument(prjn, ctpn, helper);
            if(ctpl == null){
                throw new Exception("Template-Document [ " + ctpn + " ] not found.");
            }
            String tplCoverPath = Utils.exportDocument(ctpl, exportPath, ctpn);

            String mtpn = "TRANSMITTAL_MAIL";
            IDocument mtpl = Utils.getTemplateDocument(prjn, mtpn, helper);
            if(mtpl == null){
                throw new Exception("Template-Document [ " + mtpn + " ] not found.");
            }
            String tplMailPath = Utils.exportDocument(mtpl, exportPath, mtpn);

            IDocument tdoc = null;
            boolean isTDocLinked = true;
            List<String> docIds = new ArrayList<>();
            IInformationObjectLinks links = proi.getLoadedInformationObjectLinks();
            for (ILink link : links.getLinks()) {
                IDocument xdoc = (IDocument) link.getTargetInformationObject();
                if(!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}

                String dtyp = xdoc.getDescriptorValue(Conf.Descriptors.DocType, String.class);
                dtyp = (dtyp == null ? "" : dtyp);

                if(dtyp.equals("Transmittal-Outgoing")){

                    String docn = xdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                    String docr = xdoc.getDescriptorValue(Conf.Descriptors.DocRevision, String.class);
                    docn = (docn == null ? "" : docn);
                    docr = (docr == null ? "" : docr);

                    if(docn.equals(tmnr) && docr.equals("")){
                        if(tdoc == null){
                            tdoc = xdoc;
                        }
                        else{
                            srv.deleteDocument(ses, tdoc);
                        }
                    }
                    continue;
                }

                if(!docIds.contains(xdoc.getID())){
                    docIds.add(xdoc.getID());
                }

            }
            if(tdoc == null) {
                tdoc = Utils.createTransmittalDocument(ses, srv, prjt);
                isTDocLinked = false;
            }

            JSONObject pbks = Conf.Bookmarks.ProjectWorkspace();
            JSONObject pbts = Conf.Bookmarks.ProjectWorkspaceTypes();
            for (String pkey : pbks.keySet()) {
                String pfld = pbks.getString(pkey);
                //System.out.println("&&& PFLD [" + pkey + "] *** " + pfld);
                if(pfld.isEmpty()){continue;}

                if(!pbts.has(pkey)) {
                    String pvalString = proi.getDescriptorValue(pfld, String.class);
                    if(pvalString != null && Utils.hasDescriptor((IInformationObject) tdoc, pfld)) {
                        tdoc.setDescriptorValueTyped(pfld, pvalString);
                    }
                    //System.out.println(">>> PFLD.String [" + pkey + "] *** " + pvalString);
                    pbks.put(pkey, (pvalString == null ? "" : pvalString));
                } else if(pbts.get(pkey) == Integer.class){
                    Integer pvalInteger = proi.getDescriptorValue(pfld, Integer.class);
                    if(pvalInteger != null && Utils.hasDescriptor((IInformationObject) tdoc, pfld)) {
                        tdoc.setDescriptorValueTyped(pfld, pvalInteger);
                    }
                    //System.out.println(">>> PFLD.Integer [" + pkey + "] *** " + pvalInteger);
                    pbks.put(pkey, (pvalInteger == null ? "" : pvalInteger.toString()));

                } else if(pbts.get(pkey) == Double.class){
                    Double pvalDouble = proi.getDescriptorValue(pfld, Double.class);
                    if(pvalDouble != null && Utils.hasDescriptor((IInformationObject) tdoc, pfld)) {
                        tdoc.setDescriptorValueTyped(pfld, pvalDouble);
                    }
                    //System.out.println(">>> PFLD.Double [" + pkey + "] *** " + pvalDouble);
                    pbks.put(pkey, (pvalDouble == null ? "" : pvalDouble.toString()));
                }
            }
            pbks.put("TransmittalNo", tmnr);

            String tuss = Utils.getWorkbasketDisplayNames(ses, srv, pbks.getString("To"));
            pbks.put("To", tuss);

            String auss = Utils.getWorkbasketDisplayNames(ses, srv, pbks.getString("Attention"));
            pbks.put("Attention", auss);

            String cuss = Utils.getWorkbasketDisplayNames(ses, srv, pbks.getString("CC"));
            pbks.put("CC", cuss);


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

            proi.setDescriptorValue(Conf.Descriptors.ObjectNumberExternal,
                    tmnr);
            proi.setDescriptorValue(Conf.Descriptors.ProjectNo,
                    prjt.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class));
            proi.setDescriptorValue(Conf.Descriptors.ProjectName,
                    prjt.getDescriptorValue(Conf.Descriptors.ProjectName, String.class));
            proi.setDescriptorValue(Conf.Descriptors.DccList,
                    prjt.getDescriptorValue(Conf.Descriptors.DccList, String.class));

            String poId = proi.getID();
            Thread.sleep(2000);
            proi.commit();
            proi = (IProcessInstance) srv.getInformationObjectByID(poId, ses);


            Integer lcnt = 0;
            JSONObject ebks = Conf.Bookmarks.EngDocument();

            List<String> expFilePaths = new ArrayList<>();
            List<String> newLinks = new ArrayList<>();
            List<String> linkeds = new ArrayList<>();


            for (ILink link : links.getLinks()) {
                IDocument edoc = (IDocument) link.getTargetInformationObject();
                if(!edoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
                if(linkeds.contains(edoc.getID())){continue;}
                if(!docIds.contains(edoc.getID())){continue;}


                String docNo = edoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                docNo = docNo == null ? "" : docNo;

                String revNo = edoc.getDescriptorValue(Conf.Descriptors.DocRevision, String.class);
                revNo = revNo == null ? "" : revNo;

                String parentDocNo = edoc.getDescriptorValue(Conf.Descriptors.ParentDocNumber, String.class);
                parentDocNo = parentDocNo == null ? "" : parentDocNo;

                String parentRevNo = edoc.getDescriptorValue(Conf.Descriptors.ParentDocRevision, String.class);
                parentRevNo = parentRevNo == null ? "" : parentRevNo;


                String parentDoc = parentDocNo + (!parentDocNo.isEmpty() && !parentRevNo.isEmpty() ? "/" : "") + parentRevNo;

                String fileName = edoc.getDescriptorValue(Conf.Descriptors.FileName, String.class);
                fileName = (fileName == null ? "" : fileName);
                if(fileName.isEmpty()){continue;}

                //String expPath = Utils.exportDocument(edoc, exportPath, docNo + "_" + revNo);
                String expPath = Utils.exportDocument(edoc, exportPath, FilenameUtils.removeExtension(fileName));
                if(expFilePaths.contains(expPath)){continue;}

                lcnt++;
                System.out.println("IDOC [" + lcnt + "] *** " + edoc.getID());
                String llfx = (lcnt <= 9 ? "0" : "") + lcnt;

                for (String ekey : ebks.keySet()) {
                    String einx = ekey + llfx;
                    String efld = ebks.getString(ekey);
                    if(efld.isEmpty()){continue;}
                    String eval = "";
                    if(efld.equals("@EXPORT_FILE_NAME@")){
                        eval = Paths.get(expPath).getFileName().toString();
                    }
                    if(eval.isEmpty()){
                        eval = edoc.getDescriptorValue(efld, String.class);
                    }
                    pbks.put(einx, eval);
                }

                pbks.put("ParentDoc" + llfx, parentDoc);
                if(docNo.isEmpty()){continue;}

                expFilePaths.add(expPath);


                edoc.setDescriptorValue(Conf.Descriptors.DocTransOutCode, tmnr);
                edoc.commit();
                linkeds.add(edoc.getID());

                IDocument cdoc = (IDocument) Utils.getEngineeringCRS(edoc.getID(), helper);
                if(cdoc != null){
                    String crsNo = cdoc.getDescriptorValue(Conf.Descriptors.ObjectNumber, String.class);
                    if(!crsNo.isEmpty()){
                        String crsPath = Utils.exportDocument(cdoc, exportPath, FilenameUtils.removeExtension(fileName) + "_" + crsNo);
                        expFilePaths.add(crsPath);
                    }
                }
            }

            ILink[] tlnks = srv.getReferencedRelationships(ses, tdoc, true);
            JSONObject tlnkds = new JSONObject();
            for (ILink tlnk : tlnks) {
                IInformationObject ttgt = tlnk.getTargetInformationObject();
                tlnkds.put(ttgt.getID(), tlnk);
            }
            for(String lnkd : linkeds){
                if(tlnkds.has(lnkd)){
                    tlnkds.remove(lnkd);
                   continue;
                }
                ILink lnk1 = srv.createLink(ses, tdoc.getID(), null, lnkd);
                lnk1.commit();
            }

            List dstLines = new ArrayList<>();

            for(int p=1;p<=5;p++){
                String plfx = (p <= 9 ? "0" : "") + p;
                if(pbks.getString(Conf.Bookmarks.DistributionMaster + plfx) == ""){continue;}

                if(!dstLines.contains("CMT01")){dstLines.add("CMT01");}
                if(!dstLines.contains("HDR01")){dstLines.add("HDR01");}
                if(!dstLines.contains(plfx)){dstLines.add(plfx);}
            }

            List docLines = new ArrayList<>();
            for(int l=1;l<=50;l++){
                String llfx = (l <= 9 ? "0" : "") + l;
                if(pbks.getString(Conf.Bookmarks.EngDocumentMaster + llfx) == ""){continue;}

                if(!docLines.contains("CMT01")){docLines.add("CMT01");}
                if(!docLines.contains("HDR01")){docLines.add("HDR01");}
                if(!docLines.contains(llfx)){docLines.add(llfx);}

                /*
                if(pbks.getString("ChDocNo" + llfx) != ""){
                    pbks.put("DocNo" + llfx, "");
                }
                */
            }


            String sdty = proi.getDescriptorValue(Conf.Descriptors.TrmtSendType, String.class);
            sdty = (sdty == null ? "" : sdty);
            pbks.put("DoxisLink", "");
            if(sdty.contains("URL")) {
                JSONObject mcfg = Utils.getMailConfig(ses, srv, mtpn);
                pbks.put("DoxisLink", mcfg.getString("webBase") + helper.getDocumentURL(tdoc.getID()));
            }
            String coverExcelPath = Utils.saveTransmittalExcel(tplCoverPath, Conf.ExcelTransmittalSheetIndex.Cover,
                    exportPath + "/" + ctpn + ".xlsx", pbks);

            Utils.removeRows(coverExcelPath, coverExcelPath,
                    Conf.ExcelTransmittalSheetIndex.Cover,
                    Conf.ExcelTransmittalRowGroups.CoverDocs,
                    Conf.ExcelTransmittalRowGroups.CoverDocColInx,
                    Conf.ExcelTransmittalRowGroups.CoverHideCols,
                    docLines);

            String pdfPath = Utils.convertExcelToPdf(coverExcelPath, exportPath + "/" + ctpn + ".pdf");
            String zipPath = Utils.zipFiles(exportPath + "/Blobs.zip", pdfPath, expFilePaths);


            Utils.updateTransmittalDocument(tdoc, exportPath, pdfPath, zipPath);
            tdoc.commit();

            if(!isTDocLinked) {
                links.addInformationObject(tdoc.getID());
            }
            proi.commit();

            String mailExcelPath = Utils.saveTransmittalExcel(tplMailPath, Conf.ExcelTransmittalSheetIndex.Mail,
                    exportPath + "/" + mtpn + ".xlsx", pbks);

            Utils.removeRows(mailExcelPath, mailExcelPath,
                    Conf.ExcelTransmittalSheetIndex.Mail,
                    Conf.ExcelTransmittalRowGroups.MailDocs,
                    Conf.ExcelTransmittalRowGroups.MailDocColInx,
                    Conf.ExcelTransmittalRowGroups.MailDocHideCols,
                    docLines);

            Utils.removeRows(mailExcelPath, mailExcelPath,
                    Conf.ExcelTransmittalSheetIndex.Mail,
                    Conf.ExcelTransmittalRowGroups.MailDists,
                    Conf.ExcelTransmittalRowGroups.MailDistColInx,
                    Conf.ExcelTransmittalRowGroups.MailDistHideCols,
                    dstLines);

            String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath, exportPath + "/" + mtpn + ".html");

            JSONObject mail = new JSONObject();

            List<String> stos = proi.getDescriptorValues("To-Receiver", String.class);
            List<String> sc1s = proi.getDescriptorValues("ObjectAuthors", String.class);
            List<String> sc2s = proi.getDescriptorValues("CC-Receiver", String.class);

            String mtos = Utils.getWorkbasketEMails(ses, srv, bpm, String.join(";", stos));
            String mc1s = Utils.getWorkbasketEMails(ses, srv, bpm, String.join(";", sc1s));
            String mc2s = Utils.getWorkbasketEMails(ses, srv, bpm, String.join(";", sc2s));

            mail.put("To", mtos);
            mail.put("CC", mc1s + (mc1s != "" && mc2s != "" ? ";" : "") + mc2s);
            mail.put("Subject", "Transmittal - " + tmnr);

            List<String> aths = new ArrayList<>();
            if(!pdfPath.isEmpty() && sdty.contains("COVER")){
                aths.add(pdfPath);
            }
            if(!zipPath.isEmpty() && sdty.contains("ZIP")){
                aths.add(zipPath);
            }

            mail.put("AttachmentPaths", String.join(";", aths));
            if(sdty.contains("COVER")) {
                mail.put("AttachmentName." + Paths.get(pdfPath).getFileName().toString(), "Cover_Preview[" + tmnr + "].pdf");
            }
            if(sdty.contains("ZIP")) {
                mail.put("AttachmentName." + Paths.get(zipPath).getFileName().toString(), "Eng_Documents[" + tmnr + "].zip");
            }

            mail.put("BodyHTMLFile", mailHtmlPath);

            Utils.sendHTMLMail(ses, srv, mtpn, mail);

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