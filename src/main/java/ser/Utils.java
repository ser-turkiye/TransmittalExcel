package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import com.ser.blueline.bpm.IWorkbasket;
import com.ser.blueline.metaDataComponents.IArchiveClass;
import com.ser.blueline.metaDataComponents.IStringMatrix;
import com.ser.foldermanager.IElement;
import com.ser.foldermanager.IElements;
import com.ser.foldermanager.IFolder;
import com.ser.foldermanager.INode;
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
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class Utils {

    public static boolean addToNode(IInformationObject info, String nodeName, IDocument pdoc) throws Exception {
        IFolder fold = ((IFolder) info);
        fold.refresh(true);

        boolean rtrn = false;
        List<INode> nods = fold.getNodesByName(nodeName);

        for(INode node : nods) {
            node.refresh(true);
            boolean isExistElement = false;
            IElements nelements = node.getElements();
            for (int i = 0; i < nelements.getCount2(); i++) {
                IElement nelement = nelements.getItem2(i);
                String edocID = nelement.getLink();
                String pdocID = pdoc.getID();
                if (Objects.equals(pdocID, edocID)) {
                    isExistElement = true;
                    break;
                }
            }
            if (isExistElement) {continue;}

            if (fold.addInformationObjectToNode(pdoc.getID(), node.getID())) {
                pdoc.commit();
            }

            rtrn = true;
        }
        if(rtrn){fold.commit();}
        return rtrn;
    }
    public static IInformationObject getContractorFolder(String prjCode, String compCode, ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.InvolveParty).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjCardCode).append(" = '").append(prjCode).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.ShortName).append(" = '").append(compCode).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.ProjectWorkspace} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return informationObjects[0];
    }
    static String getTransmittalNr(ISession ses, IInformationObject projectInfObj, IProcessInstance processInstance) throws Exception {
        String rtrn = processInstance.getDescriptorValue(Conf.Descriptors.TransmittalNr, String.class);
        rtrn = (rtrn == null ? "" : rtrn.trim());
        if(rtrn.isEmpty()) {

            String clientNo = "";
            if(Utils.hasDescriptor(projectInfObj, Conf.Descriptors.ClientNo)){
                clientNo = projectInfObj.getDescriptorValue(Conf.Descriptors.ClientNo, String.class);
                clientNo = (clientNo == null ? "" : clientNo).trim();
            }
            String projectNo = projectNr(projectInfObj);
            String cntPattern = "";
            if(Utils.hasDescriptor(projectInfObj, Conf.Descriptors.TrmtCounterPattern)){
                cntPattern = projectInfObj.getDescriptorValue(Conf.Descriptors.TrmtCounterPattern, String.class);
                cntPattern = (cntPattern == null ? "" : cntPattern).trim();
            }
            Integer cntStart = 0;
            if(Utils.hasDescriptor(projectInfObj, Conf.Descriptors.TrmtCounterStart)){
                cntStart = projectInfObj.getDescriptorValue(Conf.Descriptors.TrmtCounterStart, Integer.class);
                cntStart = (cntStart == null || cntStart < 0 ? 0 : cntStart);
            }
            String senderCode = "";
            if(Utils.hasDescriptor(processInstance, Conf.Descriptors.SenderCode)){
                senderCode = processInstance.getDescriptorValue(Conf.Descriptors.SenderCode, String.class);
                senderCode = (senderCode == null ? "" : senderCode).trim();
            }
            String receiverCode = "";
            if(Utils.hasDescriptor(processInstance, Conf.Descriptors.ReceiverCode)){
                receiverCode = processInstance.getDescriptorValue(Conf.Descriptors.ReceiverCode, String.class);
                receiverCode = (receiverCode == null ? "" : receiverCode).trim();
            }

            if(!clientNo.isEmpty() && !projectNo.isEmpty() && !cntPattern.isEmpty()
            && !senderCode.isEmpty() && !receiverCode.isEmpty()){
                String counterName = AutoText.init().with(projectInfObj)
                        .param("ClientNo", clientNo)
                        .param("ProjectNo", projectNo)
                        .param("Sender", senderCode)
                        .param("Receiver", receiverCode)
                        .run("{ProjectNo}.{ClientNo}.{Sender}.{Receiver}");

                NumberRange nr = new NumberRange();
                if(!nr.has(counterName)){
                    nr.append(counterName, cntPattern, Long.parseLong(cntStart.toString()));
                }

                nr.parameter("ClientNo", clientNo);
                nr.parameter("ProjectNo", projectNo);
                nr.parameter("Sender", senderCode);
                nr.parameter("Receiver", receiverCode);
                rtrn = nr.increment(counterName);
            }
            //rtrn = (new CounterHelper(ses, processInstance.getClassID())).getCounterStr();
            processInstance.setDescriptorValue(Conf.Descriptors.TransmittalNr,
                    rtrn);
        }
        return rtrn;
    }
    static String projectNr(IInformationObject projectInfObj) throws Exception {
        String rtrn = "";
        if(Utils.hasDescriptor(projectInfObj, Conf.Descriptors.ProjectNo)){
            rtrn = projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            rtrn = (rtrn == null ? "" : rtrn).trim();
        }
        return rtrn;
    }
    static List<JSONObject> getWorkbaskets(ISession ses, IDocumentServer srv, String users) throws Exception {
        List<JSONObject> rtrn = new ArrayList<>();

        IStringMatrix mtrx = srv.getStringMatrixByID("Workbaskets", ses);
        if (mtrx == null) throw new Exception("Workbaskets Global Value List not found");

        String[] usrs = users.split("\\;");

        for (String usr : usrs) {
            JSONObject wusr = getWorkbasket(ses, srv, usr.trim(), mtrx);
            if(wusr == null){continue;}
            rtrn.add(wusr);
        }
        return rtrn;
    }
    static String getWorkbasketEMails(ISession ses, IDocumentServer srv, IBpmService bpm, String users) throws Exception {
        List<JSONObject> wrbs = getWorkbaskets(ses, srv, users);
        List<String> rtrn = new ArrayList<>();
        for (JSONObject wrba : wrbs) {
            if(wrba.get("ID") == null){continue;}
            IWorkbasket wb = bpm.getWorkbasket(wrba.getString("ID"));
            if(wb == null){continue;}

            String mail1 = wb.getNotifyEMail();
            if(mail1 != null && !rtrn.contains(mail1)){
                rtrn.add(mail1);
                continue;
            }

            IUser us = wb.getOwner();
            if(us == null){continue;}

            String mail2 = us.getEMailAddress();
            if(mail2 != null && !rtrn.contains(mail2)){
                rtrn.add(mail2);
            }

        }
        return String.join(";", rtrn);
    }
    static String getWorkbasketDisplayNames(ISession ses, IDocumentServer srv, String users) throws Exception {
        List<JSONObject> wrbs = getWorkbaskets(ses, srv, users);
        List<String> rtrn = new ArrayList<>();
        for (JSONObject wrba : wrbs) {
            if(wrba.get("DisplayName") == null){continue;}
            rtrn.add(wrba.getString("DisplayName"));
        }
        return String.join(";", rtrn);
    }
    static void sendHTMLMail(ISession ses, JSONObject pars) throws Exception {
        JSONObject mcfg = Utils.getMailConfig(ses);

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
        htmlBodyPart.setContent(getHTMLFileContent(pars.getString("BodyHTMLFile")), "text/html; charset=UTF-8"); //5
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
    static String getFileContent (String path) throws Exception {
        String rtrn = new String(Files.readAllBytes(Paths.get(path)));
        return rtrn;
    }
    static String getHTMLFileContent (String path) throws Exception {
        String rtrn = new String(Files.readAllBytes(Paths.get(path)), "UTF-8");
        rtrn = rtrn.replace("\uFEFF", "");
        rtrn = rtrn.replace("ï»¿", "");
        return rtrn;
    }
    static JSONObject getSystemConfig(ISession ses) throws Exception {
        return getSystemConfig(ses, null);
    }
    static JSONObject getSystemConfig(ISession ses, IStringMatrix mtrx) throws Exception {
        if(mtrx == null){
            mtrx = ses.getDocumentServer().getStringMatrix("CCM_SYSTEM_CONFIG", ses);
        }
        if(mtrx == null) throw new Exception("SystemConfig Global Value List not found");

        List<List<String>> rawTable = mtrx.getRawRows();

        String srvn = ses.getSystem().getName().toUpperCase();
        JSONObject rtrn = new JSONObject();
        for(List<String> line : rawTable) {
            String name = line.get(0);
            if(!name.toUpperCase().startsWith(srvn + ".")){continue;}
            name = name.substring(srvn.length() + ".".length());
            rtrn.put(name, line.get(1));
        }
        return rtrn;
    }
    static JSONObject getMailConfig(ISession ses) throws Exception {
        return getMailConfig(ses, null);
    }
    static JSONObject getMailConfig(ISession ses, IStringMatrix mtrx) throws Exception {
        if(mtrx == null) {
            mtrx = ses.getDocumentServer().getStringMatrix("CCM_MAIL_CONFIG", ses);
        }
        if(mtrx == null) throw new Exception("MailConfig Global Value List not found");
        List<List<String>> rawTable = mtrx.getRawRows();

        JSONObject rtrn = new JSONObject();
        for(List<String> line : rawTable) {
            rtrn.put(line.get(0), line.get(1));
        }
        return rtrn;
    }
    static JSONObject getWorkbasket(ISession ses, IDocumentServer srv, String userID, IStringMatrix mtrx) throws Exception {
        if(mtrx == null){
            mtrx = srv.getStringMatrixByID("Workbaskets", ses);
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
    static IDocument createTransmittalDocument(ISession ses, IDocumentServer srv, IInformationObject infObj)  {

        IArchiveClass ac = srv.getArchiveClass(Conf.ClassIDs.EngineeringDocument, ses);
        IDatabase db = ses.getDatabase(ac.getDefaultDatabaseID());

        IDocument rtrn = srv.getClassFactory().getDocumentInstance(db.getDatabaseName(), ac.getID(), "0000" , ses);

        if(infObj != null) {
            rtrn.setDescriptorValue(Conf.Descriptors.MainDocumentID, ((IDocument) infObj).getID());
        }
        return rtrn;
    }
    static void copyFile(String spth, String tpth) throws Exception {
        FileUtils.copyFile(new File(spth), new File(tpth));
    }
    static void saveDuration(IProcessInstance processInstance) throws Exception {

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
    static String dateToString(Date dval) throws Exception {
        if(dval == null) return "";
        return new SimpleDateFormat("dd/MM/yyyy").format(dval);
    }
    static IProcessInstance updateProcessInstance(IProcessInstance prin) throws Exception {
        String prInId = prin.getID();
        prin.commit();
        Thread.sleep(2000);
        if(prInId.equals("<new>")) {
            return prin;
        }
        return (IProcessInstance) prin.getSession().getDocumentServer().getInformationObjectByID(prInId, prin.getSession());
    }
    static IDocument updateDocument(IDocument docu) throws Exception {
        String docuId = docu.getID();
        docu.commit();
        Thread.sleep(3000);
        if(docuId.equals("<new>")) {
            return docu;
        }
        return docu.getSession().getDocumentServer().getDocument4ID(docuId,  docu.getSession());
    }
    static void addTransmittalRepresentations(IDocument tdoc, String mainPath, String xlsxPath, String pdfPath, String zipPath) throws Exception {
        String tmnr = tdoc.getDescriptorValue(Conf.Descriptors.TransmittalNr, String.class);

        String _pdfPath = "";
        boolean frep = (tdoc.getRepresentationList().length > 0);
        if(!pdfPath.isEmpty()) {
            _pdfPath = mainPath + "/" + (frep ? "Cover_Preview[" + tmnr + "]" : tmnr) + "." + FilenameUtils.getExtension(pdfPath);
            Utils.copyFile(pdfPath, _pdfPath);

            IRepresentation pdfr = tdoc.addRepresentation(".pdf", "Cover_Preview");
            IDocumentPart ipdf = pdfr.addPartDocument(_pdfPath);
            tdoc.setDefaultRepresentation(tdoc.getRepresentationList().length - 1);
            frep = (tdoc.getRepresentationList().length > 0);

            //tdoc = Utils.updateDocument(tdoc);
        }
        String _xlsxPath = "";
        if(!xlsxPath.isEmpty()) {
            _xlsxPath = mainPath + "/" + (frep ? "Cover_Excel[" + tmnr + "]" : tmnr) + "." + FilenameUtils.getExtension(xlsxPath);
            Utils.copyFile(xlsxPath, _xlsxPath);

            IRepresentation xlsxr = tdoc.addRepresentation(".xlsx", "Cover_Excel");
            xlsxr.addPartDocument(_xlsxPath);
            frep = (tdoc.getRepresentationList().length > 0);



            //tdoc = Utils.updateDocument(tdoc);
        }
        String _zipPath = "";
        if(!zipPath.isEmpty()) {
            _zipPath = mainPath + "/" + (frep ? "Eng_Documents[" + tmnr + "]" : tmnr) + "." + FilenameUtils.getExtension(zipPath);
            Utils.copyFile(zipPath, _zipPath);

            IRepresentation zipr = tdoc.addRepresentation(".zip", "Eng_Documents");
            zipr.addPartDocument(_zipPath);
            frep = (tdoc.getRepresentationList().length > 0);

            //tdoc = Utils.updateDocument(tdoc);
        }
    }
    static String getZipFile(IInformationObjectLinks transmittalLinks, String exportPath, String transmittalNr,
                   List<String> documentIds, ProcessHelper helper) throws Exception {

        List<String> expFilePaths = new ArrayList<>();
        Integer lcnt = 0;
        for (ILink link : transmittalLinks.getLinks()) {
            IDocument edoc = (IDocument) link.getTargetInformationObject();
            if(!edoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
            if(!documentIds.contains(edoc.getID())){continue;}


            String docNo = edoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
            docNo = docNo == null ? "" : docNo;

            if(docNo.isEmpty()){continue;}

            String fileName = edoc.getDescriptorValue(Conf.Descriptors.FileName, String.class);
            fileName = (fileName == null ? "" : fileName);
            if(fileName.isEmpty()){continue;}

            String expPath = Utils.exportDocument(edoc, exportPath, FilenameUtils.removeExtension(fileName));
            if(expFilePaths.contains(expPath)){continue;}

            lcnt++;
            System.out.println("IDOC [" + lcnt + "] *** " + edoc.getID());
            //String llfx = (lcnt <= 9 ? "0" : "") + lcnt;

            expFilePaths.add(expPath);

            String tcod = edoc.getDescriptorValue(Conf.Descriptors.DocTransOutCode, String.class);
            tcod = (tcod == null ? "" : tcod);

            if(tcod.isEmpty() || !tcod.equals(transmittalNr)) {
                edoc.setDescriptorValue(Conf.Descriptors.DocTransOutCode, transmittalNr);
                //edoc = Utils.updateDocument(edoc);
                edoc.commit();
            }
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

        return Utils.zipFiles(exportPath + "/Blobs.zip", "", expFilePaths);
    }
    static String getTransmittalReprExport(IDocument tdoc, String type, String desc, String exportPath, String fileName) throws Exception {
        String rtrn = "";
        if(type.isEmpty()){return rtrn;}

        IRepresentation[] reps = tdoc.getRepresentationList();
        for(IRepresentation repr : reps){
            if(!repr.getType().equals(type)){continue;}
            if(!desc.isEmpty() && !repr.getDescription().equals(desc)){continue;}
            rtrn = exportRepresentation(tdoc, repr.getRepresentationNumber(), exportPath, fileName);
        }
        return rtrn;
    }
    static IProcessInstance createEngineeringProjectTransmittal(ProcessHelper helper) throws Exception {
        IProcessInstance rtrn = helper.buildNewProcessInstanceForID(Conf.ClassIDs.EngineeringProjectTransmittal);
        if (rtrn == null) throw new Exception("Engineering Project Transmittal couldn't be created");

        return rtrn;
    }
    public static void removeRows(String spth, String tpth, Integer shtIx, String prfx, Integer colIx, List<Integer> hlst, List<String> tlst) throws Exception {

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
    public static String saveTransmittalExcel(String templatePath, Integer shtIx, String tpltSavePath,
                             JSONObject pbks, List<String> docLines, List<String> dstLines) throws Exception {
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
                    String lnam = clvv.substring(clvv.indexOf("[[") + "[[".length(), clvv.indexOf("]]"));
                    if(pbks.has(lnam)){
                        tcll.setCellValue(lnam);
                        String lurl = pbks.getString(lnam);
                        if(!lurl.isEmpty()) {
                            Hyperlink link = twrb.getCreationHelper().createHyperlink(HyperlinkType.URL);
                            link.setAddress(lurl);
                            tcll.setHyperlink(link);
                        }
                    }
                }

                if(clvv.indexOf("[*") != (-1) && clvv.indexOf("*]") != (-1)
                && clvv.indexOf("[*") < clvv.indexOf("*]")){
                    String inam = clvv.substring(clvv.indexOf("[*") + "[*".length(), clvv.indexOf("*]"));
                    tcll.setCellValue("");
                    if(pbks.has(inam)){
                        String ipth = pbks.getString(inam);
                        File file = new File(ipth);
                        if(file.exists() && !file.isDirectory()) {
                            String iext = FilenameUtils.getExtension(ipth).toUpperCase();
                            int inpx = 0;
                            if(iext.equals("PNG")){
                                inpx = XSSFWorkbook.PICTURE_TYPE_PNG;
                            }
                            if(iext.equals("GIF")){
                                inpx = XSSFWorkbook.PICTURE_TYPE_GIF;
                            }
                            if(iext.equals("JPG") || iext.equals("JPEG")){
                                inpx = XSSFWorkbook.PICTURE_TYPE_JPEG;
                            }
                            if(iext.equals("TIF") || iext.equals("TIFF")){
                                inpx = XSSFWorkbook.PICTURE_TYPE_TIFF;
                            }
                            if(iext.equals("BMP")){
                                inpx = XSSFWorkbook.PICTURE_TYPE_BMP;
                            }
                            if(inpx > 0) {
                                int rlen = 1;
                                if(pbks.has(inam + ".RowLen")){
                                    int rltm = pbks.getInt(inam + ".RowLen");
                                    rlen = (rltm > rlen ? rltm : rlen);
                                }

                                FileInputStream inst = new FileInputStream(ipth);
                                int pinx = twrb.addPicture(IOUtils.toByteArray(inst), inpx);
                                inst.close();

                                Drawing drwg = tsht.createDrawingPatriarch();
                                XSSFClientAnchor ancr = twrb.getCreationHelper().createClientAnchor();
                                ancr.setCol1(tcll.getColumnIndex());
                                ancr.setRow1(tcll.getRowIndex());


                                Picture pict = drwg.createPicture(ancr, pinx);
                                float rowh = tcll.getRow().getHeightInPoints();
                                double pich = pict.getImageDimension().getHeight();
                                double scale = (rowh/pich)*(double)rlen;
                                pict.resize();
                                if(scale < 1){pict.resize(scale);}
                            }
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
                Conf.ExcelTransmittalRowGroups.MailDocs,
                Conf.ExcelTransmittalRowGroups.MailDocColInx,
                Conf.ExcelTransmittalRowGroups.MailDocHideCols,
                docLines);

        Utils.removeRows(rtrn, rtrn,
                Conf.ExcelTransmittalSheetIndex.Mail,
                Conf.ExcelTransmittalRowGroups.MailDists,
                Conf.ExcelTransmittalRowGroups.MailDistColInx,
                Conf.ExcelTransmittalRowGroups.MailDistHideCols,
                dstLines);


        return rtrn;
    }
    public static boolean hasDescriptor(IInformationObject infObj, String dscn) throws Exception {
        IValueDescriptor[] vds = infObj.getDescriptorList();
        for(IValueDescriptor vd : vds){
            if(vd.getName().equals(dscn)){return true;}
        }
        return false;
    }
    public static List<String> excelDstTblLines(JSONObject bookmarks) throws Exception{
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
    public static List<String> excelDocTblLines(JSONObject bookmarks) throws Exception{
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
    public static JSONObject loadBookmarks(ISession session, IDocumentServer server,
                      String transmittalNr, IInformationObjectLinks transmittalLinks,
                      IInformationObject projectInfObj, IInformationObject contractorInfObj,
                      List<String> linkedDocIds, List<String> documentIds,
                      IProcessInstance processInstance, IDocument transmittalDoc,
                      String exportPath, ProcessHelper helper) throws Exception{
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


            if(pbts.has(pkey) && pbts.get(pkey) == Date.class){
                Date pvalDate = processInstance.getDescriptorValue(pfld, Date.class);
                if(pvalDate != null && Utils.hasDescriptor((IInformationObject) transmittalDoc, pfld)) {
                    transmittalDoc.setDescriptorValueTyped(pfld, pvalDate);
                }
                rtrn.put(pkey, (pvalDate == null ? "" : Utils.dateToString(pvalDate)));
                continue;

            }
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

        if(!rtrn.getString("Approved").isEmpty()
        && !rtrn.getString("ProjectNo").isEmpty()){
            IUser asgUser = server.getUser(session, rtrn.getString("Approved"));
            if(asgUser != null){
                rtrn.put("ApprvdJobTitle", asgUser.getDescription());
            }
            rtrn.put("ApprvdFullname", Utils.getWorkbasketDisplayNames(session, server, rtrn.getString("Approved")));
            rtrn.put("ApprvdDate", rtrn.has("ApprovedDate") ? rtrn.getString("ApprovedDate") : "");
            IDocument asgDoc = null;
            asgDoc = asgDoc != null ? asgDoc : getSignatureDocument(contractorInfObj, rtrn.getString("Approved"));
            asgDoc = asgDoc != null ? asgDoc : getSignatureDocument(projectInfObj, rtrn.getString("Approved"));
            if(asgDoc != null){
                rtrn.put("ApprvdSignature", Utils.exportDocument(asgDoc, exportPath, rtrn.getString("Approved")));
                rtrn.put("ApprvdSignature.RowLen", 4);
            }
        }
        if(!rtrn.getString("Originated").isEmpty()
        && !rtrn.getString("ProjectNo").isEmpty()){
            IUser asgUser = server.getUser(session, rtrn.getString("Originated"));
            if(asgUser != null){
                rtrn.put("OrigndJobTitle", asgUser.getDescription());
            }
            rtrn.put("OrigndFullname", Utils.getWorkbasketDisplayNames(session, server, rtrn.getString("Originated")));
            rtrn.put("OrigndDate", rtrn.has("OriginatedDate") ? rtrn.getString("OriginatedDate") : "");
            IDocument osgDoc = null;
            osgDoc = osgDoc != null ? osgDoc : getSignatureDocument(contractorInfObj, rtrn.getString("Originated"));
            osgDoc = osgDoc != null ? osgDoc : getSignatureDocument(projectInfObj, rtrn.getString("Originated"));
            if(osgDoc != null){
                rtrn.put("OrigndSignature", Utils.exportDocument(osgDoc, exportPath, rtrn.getString("Originated")));
                rtrn.put("OrigndSignature.RowLen", 4);
            }
        }

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
    public static List<String> getLinkedDocIds(IInformationObjectLinks transmittalLinks)  {
        List<String> rtrn = new ArrayList<>();
        for (ILink link : transmittalLinks.getLinks()) {
            IDocument xdoc = (IDocument) link.getTargetInformationObject();
            if(!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}

            String dtyp = xdoc.getDescriptorValue(Conf.Descriptors.DocType, String.class);
            dtyp = (dtyp == null ? "" : dtyp);

            if(dtyp.equals("Transmittal-Outgoing")) {
                continue;
            }
            if(rtrn.contains(xdoc.getID())) {
                continue;
            }
            rtrn.add(xdoc.getID());
        }
        return rtrn;
    }
    public static String convertExcelToPdf(String excelPath, String pdfPath)  {
        Workbook workbook = new Workbook();
        workbook.loadFromFile(excelPath);
        workbook.getConverterSetting().setSheetFitToPage(true);
        workbook.saveToFile(pdfPath, FileFormat.PDF);

        return pdfPath;
    }
    public static String convertExcelToHtml(String excelPath, String htmlPath)  {
        Workbook workbook = new Workbook();
        workbook.loadFromFile(excelPath);
        Worksheet sheet = workbook.getWorksheets().get(0);
        HTMLOptions options = new HTMLOptions();
        options.setImageEmbedded(true);
        sheet.saveToHtml(htmlPath, options);
        return htmlPath;
    }
    public static String zipFiles(String zipPath, String pdfPath, List<String> expFilePaths) throws IOException {
        if(expFilePaths.size() == 0){return "";}

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
    public static String exportDocument(IDocument document, String exportPath, String fileName) throws IOException {
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
    public static String exportRepresentation(IDocument document, int rinx, String exportPath, String fileName) throws IOException {
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
    public static JSONObject getExcelConfig(XSSFWorkbook workbook) throws Exception {
        JSONObject rtrn = new JSONObject();
        Sheet sheet = workbook.getSheet("#CONFIG");
        if(sheet == null){throw new Exception("#CONFIG sheet not found.");}

        for(Row row : sheet) {
            Cell cll1 = row.getCell(0);
            if(cll1 == null){continue;}

            Cell cll3 = row.getCell(2);
            if(cll3 == null){continue;}

            if(cll1.getCellType() != CellType.STRING){continue;}
            String cnam = cll1.getStringCellValue().trim();
            if(cnam.isEmpty()){continue;}

            String ctyp = "String";
            Cell cll2 = row.getCell(1);
            if(cll2 != null) {
                CellType ttyp = cll2.getCellType();
                if (ttyp == CellType.STRING) {
                    ctyp = cll2.getStringCellValue().trim();
                }
            }

            CellType tval = cll3.getCellType();
            if(tval == CellType.STRING && ctyp.equals("String")) {
                String cvalString = cll3.getStringCellValue().trim();
                rtrn.put(cnam, cvalString);
            }
            if(tval == CellType.NUMERIC && ctyp.equals("Numeric")) {
                Double cvalNumeric = cll3.getNumericCellValue();
                rtrn.put(cnam, cvalNumeric);
            }
        }
        return rtrn;
    }
    public static JSONObject getDataOfTransmittal(XSSFWorkbook workbook, JSONObject ecfg) throws Exception {
        JSONObject rtrn = new JSONObject();
        if(!ecfg.has("SheetName")){throw new Exception("#CONFIG[SheetName] not found.");}

        Sheet sheet = workbook.getSheet(ecfg.getString("SheetName"));
        String[] keys = {"ProjectNo", "ProjectName", "TransmittalNo",
                "To", "Attention", "CC", "JobNo", "IssueDate", "TransmittalType",
                "SenderCode", "SenderName", "ReceiverCode", "ReceiverName",
                "Summary", "Notes"};
        for(String skey : keys){
            rtrn.put(skey,
                (!ecfg.has("CellPos." + skey) ? "" : getCellValue(sheet, ecfg.getString("CellPos." + skey)))
            );
        }
        return rtrn;
    }
    public static String getCellValue(Sheet sheet, String refn){

        CellReference cr = new CellReference(refn);
        Row row = sheet.getRow(cr.getRow());
        Cell rtrn = row.getCell(cr.getCol());
        return rtrn.getRichStringCellValue().getString();
    }
    public static String updateCell(String str, JSONObject bookmarks){
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
    static IInformationObject getProjectWorkspace(String prjn, ProcessHelper helper) {
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
    static IInformationObject getEngineeringCRS(String refn, ProcessHelper helper) {
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
    static IDocument getTemplateDocument(IInformationObject info, String tpltName) throws Exception {
        List<INode> nods = ((IFolder) info).getNodesByName("Templates");
        for(INode node : nods){
            IElements elms = node.getElements();

            for(int i=0;i<elms.getCount2();i++) {
                IElement nelement = elms.getItem2(i);
                String edocID = nelement.getLink();
                IInformationObject tplt = info.getSession().getDocumentServer().getInformationObjectByID(edocID, info.getSession());
                if(tplt == null){continue;}

                if(!hasDescriptor(tplt, Conf.Descriptors.TemplateName)){continue;}

                String etpn = tplt.getDescriptorValue(Conf.Descriptors.TemplateName, String.class);
                if(etpn == null || !etpn.equals(tpltName)){continue;}

                return (IDocument) tplt;
            }
        }
        return null;
    }
    static IDocument getTemplateDocument_old(String prjNo, String tpltName, ProcessHelper helper)  {
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
    static IDocument getSignatureDocument(IInformationObject info, String tpltName) throws Exception {
        List<INode> nods = ((IFolder) info).getNodesByName("Signatures");
        for(INode node : nods){
            IElements elms = node.getElements();

            for(int i=0;i<elms.getCount2();i++) {
                IElement nelement = elms.getItem2(i);
                String edocID = nelement.getLink();
                IInformationObject tplt = info.getSession().getDocumentServer().getInformationObjectByID(edocID, info.getSession());
                if(tplt == null){continue;}

                if(!hasDescriptor(tplt, Conf.Descriptors.TemplateName)){continue;}

                String etpn = tplt.getDescriptorValue(Conf.Descriptors.TemplateName, String.class);
                if(etpn == null || !etpn.equals(tpltName)){continue;}

                return (IDocument) tplt;
            }
        }
        return null;
    }
    static IDocument getSignatureDocument_old(String prjNo, String sgnrName, ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.Template).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjCardCode).append(" = '").append(prjNo).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.ObjectNumberExternal).append(" = '").append(sgnrName).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.Company} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return (IDocument) informationObjects[0];
    }
    static IInformationObject[] getChildEngineeringDocuments(String docNo, String revNo, ProcessHelper helper)  {
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
    static IDocument getEngineeringDocument(String docNo, String revNo, ProcessHelper helper)  {
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
    static IDocument getTransmittalOutgoingDocument(String docNo, ProcessHelper helper)  {
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
    public static JSONObject getRowGroups(Sheet sheet, String prfx, Integer colIx)  {
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
    public static List<JSONObject> getListOfDocuments(XSSFWorkbook workbook, JSONObject ecfg) throws Exception {
        if(!ecfg.has("SheetName"))
            {throw new Exception("#CONFIG[SheetName] not found.");}

        if(!ecfg.has("Docs.Rows-Begin"))
            {throw new Exception("#CONFIG[Docs.Rows-Begin] not found.");}

        if(!ecfg.has("Docs.Rows-End"))
            {throw new Exception("#CONFIG[Docs.Rows-End] not found.");}

        Sheet sheet = workbook.getSheet(ecfg.getString("SheetName"));
        List<JSONObject> rtrn = new ArrayList<>();

        Integer lbgn = (int) ecfg.getDouble("Docs.Rows-Begin");
        if(lbgn == null || lbgn<0)
            {throw new Exception("#CONFIG[Docs.Rows-Begin] : [ " + lbgn + " ] param check");}
        Integer lend = (int) ecfg.getDouble("Docs.Rows-End");
        if(lend == null || lend<0)
            {throw new Exception("#CONFIG[Docs.Rows-End] : [ " + lend + " ] param check");}
        if(lend<=lbgn)
            {throw new Exception("#CONFIG[Docs.Rows-End] : [ " + lend + " ]  <= #CONFIG[Docs.Rows-Begin] : [ " + lbgn + " ]");}


        String mkey = "DocNo";
        String[] keys = {mkey, "RevNo"};

        for (Row row : sheet) {
            int rnum = row.getRowNum();
            if(rnum < lbgn){continue;}
            if(rnum > lend){break;}

            JSONObject rlin = new JSONObject();
            boolean apnd = true;
            for(String skey : keys){
                if(!ecfg.has("Docs.Rows-Column-" + skey) && skey == mkey){
                    apnd = false;
                    break;}

                String cval = "";
                if(ecfg.has("Docs.Rows-Column-" + skey)){
                    cval = getCellValue(sheet, ecfg.getString("Docs.Rows-Column-" + skey) + (rnum+1));
                }

                rlin.put(skey, cval);

            }
            if(!apnd){continue;}
            rtrn.add(rlin);
        }
        return rtrn;
    }
    public static List<JSONObject> getListOfDistributions(XSSFWorkbook workbook, JSONObject ecfg) throws Exception {
        if(!ecfg.has("SheetName"))
            {throw new Exception("#CONFIG[SheetName] not found.");}

        if(!ecfg.has("Dists.Rows-Begin"))
            {throw new Exception("#CONFIG[Dists.Rows-Begin] not found.");}

        if(!ecfg.has("Docs.Rows-End"))
            {throw new Exception("#CONFIG[Dists.Rows-End] not found.");}

        Sheet sheet = workbook.getSheet(ecfg.getString("SheetName"));
        List<JSONObject> rtrn = new ArrayList<>();

        Integer lbgn = (int) ecfg.getDouble("Dists.Rows-Begin");
        if(lbgn == null || lbgn<0)
            {throw new Exception("#CONFIG[Dists.Rows-Begin] : [ " + lbgn + " ] param check");}
        Integer lend = (int) ecfg.getDouble("Dists.Rows-End");
        if(lend == null || lend<0)
            {throw new Exception("#CONFIG[Dists.Rows-End] : [ " + lend + " ] param check");}
        if(lend<=lbgn)
            {throw new Exception("#CONFIG[Dists.Rows-End] : [ " + lend + " ]  <= #CONFIG[Dists.Rows-Begin] : [ " + lbgn + " ]");}


        String mkey = "User";
        String[] keys = {mkey, "Purpose", "DlvMethod", "DueDate"};

        for (Row row : sheet) {
            int rnum = row.getRowNum();
            if(rnum < lbgn){continue;}
            if(rnum > lend){break;}

            JSONObject rlin = new JSONObject();
            boolean apnd = true;
            for(String skey : keys){
                if(!ecfg.has("Dists.Rows-Column-" + skey) && skey == mkey){
                    apnd = false;
                    break;}

                String cval = "";
                if(ecfg.has("Dists.Rows-Column-" + skey)){
                    cval = getCellValue(sheet, ecfg.getString("Dists.Rows-Column-" + skey) + (rnum+1));
                }

                rlin.put(skey, cval);

            }
            if(!apnd){continue;}
            rtrn.add(rlin);
        }
        return rtrn;
    }

}
