package ser;

import com.ser.blueline.IDocumentServer;
import com.ser.blueline.ISession;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;

import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import static java.lang.System.out;


public class TransmittalTest extends UnifiedAgent {

    ISession ses;
    IDocumentServer srv;
    IBpmService bpm;
    private ProcessHelper helper;
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");

        ses = getSes();
        srv = ses.getDocumentServer();
        bpm = getBpm();
        try {
            this.helper = new ProcessHelper(ses);
            ITask task = getEventTask();

            IProcessInstance proi = task.getProcessInstance();
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

            Calendar cbgn = Calendar.getInstance();
            cbgn.setTime(tbgn);
            cbgn.set(Calendar.HOUR, cbgn.get(Calendar.HOUR) - 135);
            cbgn.set(Calendar.MINUTE, cbgn.get(Calendar.MINUTE) + 34);
            tbgn = cbgn.getTime();

            long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
            long durd  = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
            float durh  = (TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60))/60;

            System.out.println("DURD:" + durd);
            System.out.println("DURH:" + durh);

            /*
            System.out.println("TBGN.01:" + tbgn.getTime());
            System.out.println("TEND.01:" + tend.getTime());


            System.out.println("TBGN.02:" + tbgn.getTime());
            System.out.println("TEND.02:" + tend.getTime());


            Long xdur = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
            System.out.println("XDUR:" + xdur);

            Long qhor = (long) (1000*60*60);
            Long qday = qhor*24;
            Long durd = (long) Math.round(xdur / qday);
            Long durh = (long) Math.round((xdur - (durd * qday)) / qhor);
            System.out.println("DURD:" + durd);
            System.out.println("DURH:" + durh);

             */
            /*
            IInformationObject[] lnks = Utils.getChildEngineeringDocuments(
                document.getDescriptorValue(Conf.Descriptors.CCMPrjDocNumber, String.class),
                document.getDescriptorValue(Conf.Descriptors.CCMPrjDocRevision, String.class),
                helper
            );
            */

            //(new File(Conf.ExcelTransmittalPaths.MainPath)).mkdir();

            //String mails = Utils.getWorkbasketEMails(ses, srv, bpm, "osman.dev;yunus.dev");
            //System.out.println(mails);



            //JSONObject bkms = Conf.Bookmarks.ProjectWorkspace();
            //JSONObject mcfg = Utils.getMailConfig(ses, srv, mtpn);

            /*
            JSONObject mail = new JSONObject();
            mail.put("From", "support@serturkiye.com");
            mail.put("To", "oe@serturkiye.com");
            mail.put("CC", "bb@serturkiye.com");
            mail.put("Subject", "test-html mail");
            mail.put("AttachmentPaths", "c:/tmp2/mail/test01.pdf");
            mail.put("BodyHTMLFile", "c:/tmp2/mail/body.html");

            Utils.sendHTMLMail(ses, srv, mtpn, mail);
            */

            /*
            String uniqueId = "TEST";
            String prjn = "PRJ_002";
            String mtpn = "TRANSMITTAL_MAIL";
            IDocument mtpl = Utils.getTemplateDocument(prjn, mtpn, helper);
            if(mtpl == null){
                throw new Exception("Template-Document [ " + mtpn + " ] not found.");
            }

            String tplMailPath = Utils.exportDocument(mtpl, Conf.ExcelTransmittalPaths.MainPath, mtpn + "[" + uniqueId + "]");


            List tlst = new ArrayList<>();
            tlst.add("CMT01");
            tlst.add("HDR01");
            tlst.add("01");
            Utils.removeRows(tplMailPath,
                    Conf.ExcelTransmittalPaths.MainPath + "/" + mtpn + "@@" + uniqueId + ".xlsx",
                    Conf.ExcelTransmittalSheetIndex.Mail,
                    Conf.ExcelTransmittalRowGroups.MailDocs,
                    Conf.ExcelTransmittalRowGroups.MailDocColInx,
                    Conf.ExcelTransmittalRowGroups.MailDocHideCols,
                    tlst);
            */

            out.println("Tested.");

        } catch (Exception e) {
            //throw new RuntimeException(e);
            out.println("Exception       : " + e.getMessage());
            out.println("    Class       : " + e.getClass());
            out.println("    Stack-Trace : " + e.getStackTrace() );
            return resultError("Exception : " + e.getMessage());
        }

        out.println("Finished");
        return resultSuccess("Ended successfully");
    }
}