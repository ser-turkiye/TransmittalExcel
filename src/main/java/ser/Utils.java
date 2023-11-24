package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import com.ser.blueline.bpm.IWorkbasket;
import com.ser.blueline.metaDataComponents.IArchiveClass;
import com.ser.blueline.metaDataComponents.IStringMatrix;
import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.core.spreadsheet.HTMLOptions;
import jakarta.activation.DataHandler;
import jakarta.activation.DataSource;
import jakarta.activation.FileDataSource;
import jakarta.mail.*;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeBodyPart;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.internet.MimeMultipart;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class Utils {

    static List<JSONObject>
        getWorkbaskets(ISession ses, IDocumentServer srv, String users) throws Exception {
        List<JSONObject> rtrn = new ArrayList<>();
        IStringMatrix mtrx = getWorkbasketMatrix(ses, srv);
        String[] usrs = users.split("\\;");

        for (String usr : usrs) {
            JSONObject wusr = getWorkbasket(ses, srv, usr.trim(), mtrx);
            if(wusr == null){continue;}
            rtrn.add(wusr);
        }
        return rtrn;
    }
    static String
        getWorkbasketEMails(ISession ses, IDocumentServer srv, IBpmService bpm, String users) throws Exception {
        List<JSONObject> wrbs = getWorkbaskets(ses, srv, users);
        List<String> rtrn = new ArrayList<>();
        for (JSONObject wrba : wrbs) {
            if(wrba.get("ID") == null){continue;}
            IWorkbasket wb = bpm.getWorkbasket(wrba.getString("ID"));
            if(wb == null){continue;}
            String mail = wb.getNotifyEMail();
            if(mail == null){continue;}
            rtrn.add(mail);
        }
        return String.join(";", rtrn);
    }
    static String
        getWorkbasketDisplayNames(ISession ses, IDocumentServer srv, String users) throws Exception {
        List<JSONObject> wrbs = getWorkbaskets(ses, srv, users);
        List<String> rtrn = new ArrayList<>();
        for (JSONObject wrba : wrbs) {
            if(wrba.get("DisplayName") == null){continue;}
            rtrn.add(wrba.getString("DisplayName"));
        }
        return String.join(";", rtrn);
    }
    static void
        sendHTMLMail(ISession ses, IDocumentServer srv, String mtpn, JSONObject pars) throws Exception {
        JSONObject mcfg = Utils.getMailConfig(ses, srv, mtpn);

        String host = mcfg.getString("host");
        String port = mcfg.getString("port");
        String protocol = mcfg.getString("protocol");
        String sender = mcfg.getString("sender");
        String subject = "";
        String mailTo = "";
        String mailCC = "";
        String attachments = "";

        if(pars.has("From")){
            sender = pars.getString("From");
        }
        if(pars.has("To")){
            mailTo = pars.getString("To");
        }
        if(pars.has("CC")){
            mailCC = pars.getString("CC");
        }
        if(pars.has("Subject")){
            subject = pars.getString("Subject");
        }
        if(pars.has("AttachmentPaths")){
            attachments = pars.getString("AttachmentPaths");
        }


        Properties props = new Properties();

        props.put("mail.debug","true");
        props.put("mail.smtp.debug", "true");

        props.put("mail.smtp.host", host);
        props.put("mail.smtp.port", port);

        String start_tls = (mcfg.has("start_tls") ? mcfg.getString("start_tls") : "");
        if(start_tls.equals("true")) {
            props.put("mail.smtp.starttls.enable", start_tls);
        }

        String auth = mcfg.getString("auth");
        props.put("mail.smtp.auth", auth);
        jakarta.mail.Authenticator authenticator = null;
        if(!auth.equals("false")) {
            String auth_username = mcfg.getString("auth.username");
            String auth_password = mcfg.getString("auth.password");

            if (host.contains("gmail")) {
                props.put("mail.smtp.socketFactory.port", port);
                props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
                props.put("mail.smtp.socketFactory.fallback", "false");
            }
            if (protocol != null && protocol.contains("TLSv1.2"))  {
                props.put("mail.smtp.ssl.protocols", protocol);
                props.put("mail.smtp.ssl.trust", "*");
                props.put("mail.smtp.socketFactory.port", port);
                props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
                props.put("mail.smtp.socketFactory.fallback", "false");
            }
            authenticator = new jakarta.mail.Authenticator(){
                @Override
                protected jakarta.mail.PasswordAuthentication getPasswordAuthentication(){
                    return new jakarta.mail.PasswordAuthentication(auth_username, auth_password);
                }
            };
        }
        props.put("mail.mime.charset","UTF-8");
        Session session = (authenticator == null ? Session.getDefaultInstance(props) : Session.getDefaultInstance(props, authenticator));

        MimeMessage message = new MimeMessage(session);
        message.setFrom(new InternetAddress(sender.replace(";", ",")));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(mailTo.replace(";", ",")));
        message.setRecipients(Message.RecipientType.CC, InternetAddress.parse(mailCC.replace(";", ",")));
        message.setSubject(subject);

        Multipart multipart = new MimeMultipart("mixed");

        BodyPart htmlBodyPart = new MimeBodyPart();
        htmlBodyPart.setContent(getFileContent(pars.getString("BodyHTMLFile")) , "text/html; charset=UTF-8"); //5
        multipart.addBodyPart(htmlBodyPart);

        String[] atchs = attachments.split("\\;");
        for (String atch : atchs){
            if(atch.isEmpty()){continue;}
            BodyPart attachmentBodyPart = new MimeBodyPart();
            attachmentBodyPart.setDataHandler(new DataHandler((DataSource) new FileDataSource(atch)));

            String fnam = Paths.get(atch).getFileName().toString();
            if(pars.has("AttachmentName." + fnam)){
                fnam = pars.getString("AttachmentName." + fnam);
            }

            attachmentBodyPart.setFileName(fnam);
            multipart.addBodyPart(attachmentBodyPart);

        }

        message.setContent(multipart);
        Transport.send(message);

    }
    static IStringMatrix
        getMailConfigMatrix(ISession ses, IDocumentServer srv, String mtpn) throws Exception {
        IStringMatrix rtrn = srv.getStringMatrix("CCM_MAIL_CONFIG", ses);
        if (rtrn == null) throw new Exception("MailConfig Global Value List not found");
        return rtrn;
    }
    static String
        getFileContent (String path) throws Exception {
        return new String(Files.readAllBytes(Paths.get(path)));
    }
    static JSONObject
        getMailConfig(ISession ses, IDocumentServer srv, String mtpn) throws Exception {
        return getMailConfig(ses, srv, mtpn, null);
    }
    static JSONObject
        getMailConfig(ISession ses, IDocumentServer srv, String mtpn, IStringMatrix mtrx) throws Exception {
        if(mtrx == null){
            mtrx = getMailConfigMatrix(ses, srv, mtpn);
        }
        if(mtrx == null) throw new Exception("MailConfig Global Value List not found");
        List<List<String>> rawTable = mtrx.getRawRows();

        JSONObject rtrn = new JSONObject();
        for(List<String> line : rawTable) {
            rtrn.put(line.get(0), line.get(1));
        }
        return rtrn;
    }
    static IStringMatrix
        getWorkbasketMatrix(ISession ses, IDocumentServer srv) throws Exception {
        IStringMatrix rtrn = srv.getStringMatrixByID("Workbaskets", ses);
        if (rtrn == null) throw new Exception("Workbaskets Global Value List not found");
        return rtrn;
    }
    static JSONObject
        getWorkbasket(ISession ses, IDocumentServer srv, String userID, IStringMatrix mtrx) throws Exception {
        if(mtrx == null){
            mtrx = getWorkbasketMatrix(ses, srv);
        }
        if(mtrx == null) throw new Exception("Workbaskets Global Value List not found");
        List<List<String>> rawTable = mtrx.getRawRows();

        for(List<String> line : rawTable) {
            if(line.contains(userID)) {
                JSONObject rtrn = new JSONObject();
                rtrn.put("ID", line.get(0));
                rtrn.put("Name", line.get(1));
                rtrn.put("DisplayName", line.get(2));
                rtrn.put("Active", line.get(3));
                rtrn.put("Visible", line.get(4));
                rtrn.put("Type", line.get(5));
                rtrn.put("Organization", line.get(6));
                rtrn.put("Access", line.get(7));
                return rtrn;
            }
        }
        return null;
    }
    static IDocument
        createTransmittalDocument(ISession ses, IDocumentServer srv, IInformationObject infObj)  {

        IArchiveClass ac = srv.getArchiveClass(Conf.ClassIDs.EngineeringDocument, ses);
        IDatabase db = ses.getDatabase(ac.getDefaultDatabaseID());

        IDocument rtrn = srv.getClassFactory().getDocumentInstance(db.getDatabaseName(), ac.getID(), "0000" , ses);
        //rtrn.commit();
        //srv.copyDocument2(ses, (IDocument) infObj, rtrn, CopyScope.COPY_DESCRIPTORS);

        rtrn.setDescriptorValue(Conf.Descriptors.MainDocumentID, ((IDocument) infObj).getID());
        //rtrn.commit();

        return rtrn;
    }
    static void
        copyFile(String spth, String tpth) throws Exception {
        FileUtils.copyFile(new File(spth), new File(tpth));
    }
    static void
        saveDuration(IProcessInstance processInstance) throws Exception {

        Collection<ITask> tsks = processInstance.findTasks();

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
            processInstance.setDescriptorValueTyped("ccmPrjProcStart",
                    tbgn);
            processInstance.setDescriptorValueTyped("ccmPrjProcFinish",
                    tend);

            long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
            durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
            durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
        }

        processInstance.setDescriptorValueTyped("ccmPrjProcDurDay",
                Integer.valueOf(durd + ""));
        processInstance.setDescriptorValueTyped("ccmPrjProcDurHour",
                durh);


    }
    static void
        addTransmittalRepresentations(IDocument tdoc, String mainPath, String xlsxPath, String pdfPath, String zipPath) throws Exception {
        String tmnr = tdoc.getDescriptorValue(Conf.Descriptors.ObjectNumberExternal, String.class);

        String _pdfPath = "";
        boolean frep = (tdoc.getRepresentationList().length > 0);
        if(!pdfPath.isEmpty()) {
            _pdfPath = mainPath + "/" + (frep ? "Cover_Preview[" + tmnr + "]" : tmnr) + "." + FilenameUtils.getExtension(pdfPath);
            Utils.copyFile(pdfPath, _pdfPath);

            IRepresentation pdfr = tdoc.addRepresentation(".pdf", "Cover_Preview");
            IDocumentPart ipdf = pdfr.addPartDocument(_pdfPath);
            tdoc.setDefaultRepresentation(tdoc.getRepresentationList().length - 1);
            frep = (tdoc.getRepresentationList().length > 0);
        }
        String _xlsxPath = "";
        if(!xlsxPath.isEmpty()) {
            _xlsxPath = mainPath + "/" + (frep ? "Cover_Excel[" + tmnr + "]" : tmnr) + "." + FilenameUtils.getExtension(xlsxPath);
            Utils.copyFile(xlsxPath, _xlsxPath);

            IRepresentation xlsxr = tdoc.addRepresentation(".xlsx", "Cover_Excel");
            xlsxr.addPartDocument(_xlsxPath);
            frep = (tdoc.getRepresentationList().length > 0);
        }
        String _zipPath = "";
        if(!zipPath.isEmpty()) {
            _zipPath = mainPath + "/" + (frep ? "Eng_Documents[" + tmnr + "]" : tmnr) + "." + FilenameUtils.getExtension(zipPath);
            Utils.copyFile(zipPath, _zipPath);

            IRepresentation zipr = tdoc.addRepresentation(".zip", "Eng_Documents");
            zipr.addPartDocument(_zipPath);
            frep = (tdoc.getRepresentationList().length > 0);
        }
    }
    static String
        getTransmittalReprExport(IDocument tdoc, String type, String desc, String exportPath, String fileName) throws Exception {
        String rtrn = "";
        IRepresentation[] reps = tdoc.getRepresentationList();
        for(IRepresentation repr : reps){
            if(!repr.getType().equals(type)){continue;}
            if(!repr.getDescription().equals(desc)){continue;}
            rtrn = exportRepresentation(tdoc, repr.getRepresentationNumber(), exportPath, fileName);
        }
        return rtrn;
    }
    static IProcessInstance
        createEngineeringProjectTransmittal(IDocument doc, ProcessHelper helper) throws Exception {
        IProcessInstance rtrn = helper.buildNewProcessInstanceForID(Conf.ClassIDs.EngineeringProjectTransmittal);
        if (rtrn == null) throw new Exception("Engineering Project Transmittal couldn't be created");

        helper.mapDescriptorsFromObjectToObject(doc, rtrn, true);
        //rtrn.setMainInformationObjectID(doc.getID());
        //rtrn.commit();
        return rtrn;
    }
    public static void
        removeRows(String spth, String tpth, Integer shtIx, String prfx, Integer colIx, List<Integer> hlst, List<String> tlst) throws IOException {

        FileInputStream tist = new FileInputStream(spth);
        XSSFWorkbook twrb = new XSSFWorkbook(tist);

        Sheet tsht = twrb.getSheetAt(shtIx);
        JSONObject rows = Utils.getRowGroups(tsht, prfx, colIx);

        for (String pkey : rows.keySet()) {
            Row crow = (Row) rows.get(pkey);
            crow.getCell(colIx).setBlank();

            if(tlst.contains(pkey)){
                continue;
            }

            crow.setZeroHeight(true);
            //deleteRow(tsht, crow.getRowNum());
        }

        for(Integer hcix : hlst){
            tsht.setColumnHidden(hcix, true);
        }

        FileOutputStream tost = new FileOutputStream(tpth);
        twrb.write(tost);
        tost.close();

    }
    public static String
        saveTransmittalExcel(String templatePath, Integer shtIx, String tpltSavePath,
                             JSONObject pbks, List<String> docLines, List<String> dstLines) throws IOException {
        String rtrn = tpltSavePath+"";
        FileInputStream tist = new FileInputStream(templatePath);
        XSSFWorkbook twrb = new XSSFWorkbook(tist);

        Sheet tsht = twrb.getSheetAt(shtIx);
        for (Row trow : tsht){
            for(Cell tcll : trow){
                if(tcll.getCellType() != CellType.STRING){continue;}
                String clvl = tcll.getRichStringCellValue().getString();
                String clvv = updateCell(clvl, pbks);
                if(!clvv.equals(clvl)){
                    tcll.setCellValue(clvv);
                }

                if(clvv.indexOf("[[") != (-1) && clvv.indexOf("]]") != (-1)
                && clvv.indexOf("[[") < clvv.indexOf("]]")){
                    String znam = clvv.substring(clvv.indexOf("[[") + "[[".length(), clvv.indexOf("]]"));
                    if(pbks.has(znam)){
                        tcll.setCellValue(znam);
                        String lurl = pbks.getString(znam);
                        if(!lurl.isEmpty()) {
                            Hyperlink link = twrb.getCreationHelper().createHyperlink(HyperlinkType.URL);
                            link.setAddress(lurl);
                            tcll.setHyperlink(link);
                        }
                    }
                }
            }
        }
        FileOutputStream tost = new FileOutputStream(rtrn);
        twrb.write(tost);
        tost.close();

        Utils.removeRows(rtrn, rtrn,
                Conf.ExcelTransmittalSheetIndex.Cover,
                Conf.ExcelTransmittalRowGroups.CoverDocs,
                Conf.ExcelTransmittalRowGroups.CoverDocColInx,
                Conf.ExcelTransmittalRowGroups.CoverHideCols,
                docLines);

        Utils.removeRows(rtrn, rtrn,
                Conf.ExcelTransmittalSheetIndex.Mail,
                Conf.ExcelTransmittalRowGroups.MailDists,
                Conf.ExcelTransmittalRowGroups.MailDistColInx,
                Conf.ExcelTransmittalRowGroups.MailDistHideCols,
                dstLines);

        return rtrn;
    }
    public static boolean
        hasDescriptor(IInformationObject infObj, String dscn) throws Exception {
        IValueDescriptor[] vds = infObj.getDescriptorList();
        for(IValueDescriptor vd : vds){
            if(vd.getName().equals(dscn)){return true;}
        }
        return false;
    }
    public static List<String>
        excelDstTblLines(JSONObject bookmarks) throws Exception{
        List rtrn = new ArrayList<>();

        for(int p=1;p<=5;p++){
            String plfx = (p <= 9 ? "0" : "") + p;
            if(!bookmarks.has(Conf.Bookmarks.DistributionMaster + plfx)){continue;}
            if(bookmarks.getString(Conf.Bookmarks.DistributionMaster + plfx) == ""){continue;}

            if(!rtrn.contains("CMT01")){rtrn.add("CMT01");}
            if(!rtrn.contains("HDR01")){rtrn.add("HDR01");}
            if(!rtrn.contains(plfx)){rtrn.add(plfx);}
        }
        return rtrn;
    }
    public static List<String>
        excelDocTblLines(JSONObject bookmarks) throws Exception{
        List rtrn = new ArrayList<>();
        for(int l=1;l<=50;l++){
            String llfx = (l <= 9 ? "0" : "") + l;
            if(!bookmarks.has(Conf.Bookmarks.EngDocumentMaster + llfx)){continue;}
            if(bookmarks.getString(Conf.Bookmarks.EngDocumentMaster + llfx) == ""){continue;}

            if(!rtrn.contains("CMT01")){rtrn.add("CMT01");}
            if(!rtrn.contains("HDR01")){rtrn.add("HDR01");}
            if(!rtrn.contains(llfx)){rtrn.add(llfx);}
        }
        return rtrn;
    }
    public static JSONObject
        loadBookmarks(ISession session, IDocumentServer server, String transmittalNr, IInformationObjectLinks transmittalLinks,
                                           List<String> linkedDocIds, List<String> documentIds,
                                           IProcessInstance processInstance, IDocument transmittalDoc) throws Exception{
        JSONObject rtrn = new JSONObject();
        JSONObject pbks = Conf.Bookmarks.projectWorkspace();
        JSONObject pbts = Conf.Bookmarks.projectWorkspaceTypes();
        JSONObject ebks = Conf.Bookmarks.engDocument();

        for (String pkey : pbks.keySet()) {
            String pfld = pbks.getString(pkey);
            if(pfld.isEmpty()){continue;}
            System.out.println("&&& PFLD [" + pkey + "] *** " + pfld);

            rtrn.put(pkey, "");
            if(!Utils.hasDescriptor((IInformationObject) processInstance, pfld)) {continue;}


            if(pbts.has(pkey) && pbts.get(pkey) == Integer.class){
                Integer pvalInteger = processInstance.getDescriptorValue(pfld, Integer.class);
                if(pvalInteger != null && Utils.hasDescriptor((IInformationObject) transmittalDoc, pfld)) {
                    transmittalDoc.setDescriptorValueTyped(pfld, pvalInteger);
                }
                rtrn.put(pkey, (pvalInteger == null ? "" : pvalInteger.toString()));
                continue;

            }
            if(pbts.has(pkey) && pbts.get(pkey) == Double.class){
                Double pvalDouble = processInstance.getDescriptorValue(pfld, Double.class);
                if(pvalDouble != null && Utils.hasDescriptor((IInformationObject) transmittalDoc, pfld)) {
                    transmittalDoc.setDescriptorValueTyped(pfld, pvalDouble);
                }
                rtrn.put(pkey, (pvalDouble == null ? "" : pvalDouble.toString()));
                continue;
            }

            String pvalString = processInstance.getDescriptorValue(pfld, String.class);
            if(pvalString != null && Utils.hasDescriptor((IInformationObject) transmittalDoc, pfld)) {
                transmittalDoc.setDescriptorValueTyped(pfld, pvalString);
            }
            rtrn.put(pkey, (pvalString == null ? "" : pvalString));

        }
        rtrn.put("TransmittalNo", transmittalNr);

        String tuss = Utils.getWorkbasketDisplayNames(session, server, rtrn.getString("To"));
        rtrn.put("To", tuss);

        String auss = Utils.getWorkbasketDisplayNames(session, server, rtrn.getString("Attention"));
        rtrn.put("Attention", auss);

        String cuss = Utils.getWorkbasketDisplayNames(session, server, rtrn.getString("CC"));
        rtrn.put("CC", cuss);

        int lcnt = 0;
        for (ILink link : transmittalLinks.getLinks()) {
            IDocument edoc = (IDocument) link.getTargetInformationObject();
            if(!edoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
            if(linkedDocIds.size() > 0 && !linkedDocIds.contains(edoc.getID())){continue;}
            if(!documentIds.contains(edoc.getID())){continue;}

            String parentDocNo = edoc.getDescriptorValue(Conf.Descriptors.ParentDocNumber, String.class);
            parentDocNo = parentDocNo == null ? "" : parentDocNo;

            String parentRevNo = edoc.getDescriptorValue(Conf.Descriptors.ParentDocRevision, String.class);
            parentRevNo = parentRevNo == null ? "" : parentRevNo;


            String parentDoc = parentDocNo + (!parentDocNo.isEmpty() && !parentRevNo.isEmpty() ? "/" : "") + parentRevNo;

            lcnt++;
            System.out.println("IDOC [" + lcnt + "] *** " + edoc.getID());
            String llfx = (lcnt <= 9 ? "0" : "") + lcnt;

            for (String ekey : ebks.keySet()) {
                String einx = ekey + llfx;
                String efld = ebks.getString(ekey);
                if(efld.isEmpty()){continue;}
                String eval = "";
                if(eval.isEmpty()){
                    eval = edoc.getDescriptorValue(efld, String.class);
                }
                rtrn.put(einx, eval);
            }

            rtrn.put("ParentDoc" + llfx, parentDoc);

        }
        return rtrn;
    }
    public static String
        convertExcelToPdf(String excelPath, String pdfPath)  {
        Workbook workbook = new Workbook();
        workbook.loadFromFile(excelPath);
        workbook.getConverterSetting().setSheetFitToPage(true);
        workbook.saveToFile(pdfPath, FileFormat.PDF);

        return pdfPath;
    }
    public static String
        convertExcelToHtml(String excelPath, String htmlPath)  {
        Workbook workbook = new Workbook();
        workbook.loadFromFile(excelPath);
        Worksheet sheet = workbook.getWorksheets().get(0);
        HTMLOptions options = new HTMLOptions();
        options.setImageEmbedded(true);
        sheet.saveToHtml(htmlPath, options);
        return htmlPath;
    }
    public static String
        zipFiles(String zipPath, String pdfPath, List<String> expFilePaths) throws IOException {
        ZipOutputStream zout = new ZipOutputStream(new FileOutputStream(new File(zipPath)));
        if(!pdfPath.isEmpty()) {
            //ZipEntry zltp = new ZipEntry("00." + Paths.get(tpltSavePath).getFileName().toString());
            ZipEntry zltp = new ZipEntry("_Transmittal." + FilenameUtils.getExtension(pdfPath));
            zout.putNextEntry(zltp);
            byte[] zdtp = Files.readAllBytes(Paths.get(pdfPath));
            zout.write(zdtp, 0, zdtp.length);
            zout.closeEntry();
        }

        for (String expFilePath : expFilePaths) {
            String fileName = Paths.get(expFilePath).getFileName().toString();
            fileName = fileName.replace("[@SLASH]", "/");
            ZipEntry zlin = new ZipEntry(fileName);

            zout.putNextEntry(zlin);
            byte[] zdln = Files.readAllBytes(Paths.get(expFilePath));
            zout.write(zdln, 0, zdln.length);
            zout.closeEntry();
        }
        zout.close();
        return zipPath;
    }
    public static String
        exportDocument(IDocument document, String exportPath, String fileName) throws IOException {
        String rtrn ="";
        IDocumentPart partDocument = document.getPartDocument(document.getDefaultRepresentation() , 0);
        String fName = (!fileName.isEmpty() ? fileName : partDocument.getFilename());
        fName = fName.replaceAll("[\\\\/:*?\"<>|]", "_");
        try (InputStream inputStream = partDocument.getRawDataAsStream()) {
            IFDE fde = partDocument.getFDE();
            if (fde.getFDEType() == IFDE.FILE) {
                rtrn = exportPath + "/" + fName + "." + ((IFileFDE) fde).getShortFormatDescription();

                try (FileOutputStream fileOutputStream = new FileOutputStream(rtrn)){
                    byte[] bytes = new byte[2048];
                    int length;
                    while ((length = inputStream.read(bytes)) > -1) {
                        fileOutputStream.write(bytes, 0, length);
                    }
                }
            }
        }
        return rtrn;
    }
    public static String
        exportRepresentation(IDocument document, int rinx, String exportPath, String fileName) throws IOException {
        String rtrn ="";
        IDocumentPart partDocument = document.getPartDocument(rinx , 0);
        String fName = (!fileName.isEmpty() ? fileName : partDocument.getFilename());
        fName = fName.replaceAll("[\\\\/:*?\"<>|]", "_");
        try (InputStream inputStream = partDocument.getRawDataAsStream()) {
            IFDE fde = partDocument.getFDE();
            if (fde.getFDEType() == IFDE.FILE) {
                rtrn = exportPath + "/" + fName + "." + ((IFileFDE) fde).getShortFormatDescription();

                try (FileOutputStream fileOutputStream = new FileOutputStream(rtrn)){
                    byte[] bytes = new byte[2048];
                    int length;
                    while ((length = inputStream.read(bytes)) > -1) {
                        fileOutputStream.write(bytes, 0, length);
                    }
                }
            }
        }
        return rtrn;
    }
    public static JSONObject
        getDataOfTransmittal(XSSFWorkbook workbook, Integer shtIx) throws IOException {
        JSONObject rtrn = new JSONObject();
        Sheet sheet = workbook.getSheetAt(shtIx);
        rtrn.put("ProjectNo", getCellValue(sheet, Conf.ExcelTransmittalDocsCellPos.ProjectNo));
        rtrn.put("ProjectName", "");
        rtrn.put("To", getCellValue(sheet, Conf.ExcelTransmittalDocsCellPos.To));
        rtrn.put("Attention", getCellValue(sheet, Conf.ExcelTransmittalDocsCellPos.Attention));
        rtrn.put("CC", getCellValue(sheet, Conf.ExcelTransmittalDocsCellPos.CC));
        rtrn.put("JobNo", getCellValue(sheet, Conf.ExcelTransmittalDocsCellPos.JobNo));
        rtrn.put("TransmittalNo", "");
        rtrn.put("IssueDate", getCellValue(sheet, Conf.ExcelTransmittalDocsCellPos.IssueDate));
        rtrn.put("TransmittalType", getCellValue(sheet, Conf.ExcelTransmittalDocsCellPos.TransmittalType));
        rtrn.put("Summary", getCellValue(sheet, Conf.ExcelTransmittalDocsCellPos.Summary));
        rtrn.put("Notes", getCellValue(sheet, Conf.ExcelTransmittalDocsCellPos.Notes));
        return rtrn;
    }
    public static String
        getCellValue(Sheet sheet, String refn){

        CellReference cr = new CellReference(refn);
        Row row = sheet.getRow(cr.getRow());
        Cell rtrn = row.getCell(cr.getCol());
        return rtrn.getRichStringCellValue().getString();
    }
    public static String
        updateCell(String str, JSONObject bookmarks){
        StringBuffer rtr1 = new StringBuffer();
        String tmp = str + "";
        Pattern ptr1 = Pattern.compile( "\\{([\\w\\.]+)\\}" );
        Matcher mtc1 = ptr1.matcher(tmp);
        while(mtc1.find()) {
            String mk = mtc1.group(1);
            String mv = "";
            if(bookmarks.has(mk)){
                mv = bookmarks.getString(mk);
            }
            mtc1.appendReplacement(rtr1,  mv);
        }
        mtc1.appendTail(rtr1);
        tmp = rtr1.toString();

        return tmp;
    }
    static IInformationObject
        getProjectWorkspace(String prjn, ProcessHelper helper) {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.ProjectWorkspace).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjCardCode).append(" = '").append(prjn).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.ProjectWorkspace} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return informationObjects[0];
    }
    static IInformationObject
        getEngineeringCRS(String refn, ProcessHelper helper) {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.EngineeringCRS).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjDocDocType).append(" = '").append("CRS").append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.ReferenceNumber).append(" = '").append(refn).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.EngineeringCRS} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return informationObjects[0];
    }
    static IDocument
        getTemplateDocument(String prjNo, String tpltName, ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.Template).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjCardCode).append(" = '").append(prjNo).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.ObjectNumberExternal).append(" = '").append(tpltName).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.Company} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return (IDocument) informationObjects[0];
    }
    static IInformationObject[]
        getChildEngineeringDocuments(String docNo, String revNo, ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.EngineeringDocument).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjDocParentDoc).append(" = '").append(docNo).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjDocParentDocRevision).append(" = '").append(revNo).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        return helper.createQuery(new String[]{Conf.Databases.EngineeringDocument} , whereClause, 0);
    }
    static IDocument
        getEngineeringDocument(String docNo, String revNo, ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.EngineeringDocument).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjDocNumber).append(" = '").append(docNo).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjDocRevision).append(" = '").append(revNo).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.EngineeringDocument} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return (IDocument) informationObjects[0];
    }
    static IDocument
    getTransmittalOutgoingDocument(String docNo, ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.EngineeringDocument).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjDocNumber).append(" = '").append(docNo).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjDocDocType).append(" = '").append("Transmittal-Outgoing").append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.EngineeringDocument} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return (IDocument) informationObjects[0];
    }
    public static List<JSONObject>
        getListOfDocuments(XSSFWorkbook workbook)  {
        List<JSONObject> rtrn = new ArrayList<>();
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            if(row.getRowNum() < Conf.ExcelTransmittalDocsRowIndex.Begin){continue;}
            if(row.getRowNum() > Conf.ExcelTransmittalDocsRowIndex.End){break;}

            Cell cll1 = row.getCell(Conf.ExcelTransmittalDocsCellIndex.DocNo);
            if(cll1 == null){continue;}

            Cell cll2 = row.getCell(Conf.ExcelTransmittalDocsCellIndex.RevNo);
            if(cll2 == null){continue;}

            String docNo = cll1.getRichStringCellValue().getString();
            if(docNo.isEmpty()){continue;}

            String revNo = cll2.getRichStringCellValue().getString();
            if(revNo.isEmpty()){continue;}

            JSONObject rlin = new JSONObject();
            rlin.put("docNo", docNo);
            rlin.put("revNo", revNo);
            rtrn.add(rlin);
        }
        return rtrn;
    }
    public static JSONObject
        getRowGroups(Sheet sheet, String prfx, Integer colIx)  {
        JSONObject rtrn = new JSONObject();
        for (Row row : sheet) {
            Cell cll1 = row.getCell(colIx);
            if(cll1 == null){continue;}

            String cval = cll1.getRichStringCellValue().getString();
            if(cval.isEmpty()){continue;}

            if(!cval.startsWith("[&" + prfx + ".")
            || !cval.endsWith("&]")){continue;}

            String znam = cval.substring(("[&" + prfx + ".").length(), cval.length() - ("]&").length());
            rtrn.put(znam, row);

        }
        return rtrn;
    }
    public static List<JSONObject>
        listOfDistributions(XSSFWorkbook workbook, Integer shtIx)  {
        List<JSONObject> rtrn = new ArrayList<>();
        Sheet sheet = workbook.getSheetAt(shtIx);

        for (Row row : sheet) {
            if(row.getRowNum() < Conf.ExcelTransmittalDistRowIndex.Begin){continue;}
            if(row.getRowNum() > Conf.ExcelTransmittalDistRowIndex.End){break;}

            Cell cll1 = row.getCell(Conf.ExcelTransmittalDistCellIndex.User);
            if(cll1 == null){continue;}


            String user = cll1.getRichStringCellValue().getString();
            if(user.isEmpty()){continue;}

            Cell cll2 = row.getCell(Conf.ExcelTransmittalDistCellIndex.Purpose);
            String purpose = cll2.getRichStringCellValue().getString();

            Cell cll3 = row.getCell(Conf.ExcelTransmittalDistCellIndex.DlvMethod);
            String dlvMethod = cll3.getRichStringCellValue().getString();

            Cell cll4 = row.getCell(Conf.ExcelTransmittalDistCellIndex.DueDate);
            String dueDate = cll4.getRichStringCellValue().getString();

            JSONObject rlin = new JSONObject();
            rlin.put("user", user);
            rlin.put("purpose", purpose.isEmpty() ? "" : purpose);
            rlin.put("dlvMethod", dlvMethod.isEmpty() ? "" : dlvMethod);
            rlin.put("dueDate", dueDate.isEmpty() ? "" : dueDate);
            rtrn.add(rlin);
        }
        return rtrn;
    }

}
