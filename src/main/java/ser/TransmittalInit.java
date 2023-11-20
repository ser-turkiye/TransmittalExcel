package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;

import java.util.ArrayList;
import java.util.List;

import static java.lang.System.out;


public class TransmittalInit extends UnifiedAgent {
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

        ses = getSes();
        srv = ses.getDocumentServer();
        bpm = getBpm();
        try {
            this.helper = new ProcessHelper(ses);
            ITask task = getEventTask();

            IProcessInstance proi = task.getProcessInstance();

            IInformationObjectLinks links = proi.getLoadedInformationObjectLinks();
            List<String> docIds = new ArrayList<>();

            for (ILink link : links.getLinks()) {
                IDocument xdoc = (IDocument) link.getTargetInformationObject();
                if (!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
                if (docIds.contains(xdoc.getID())){continue;}
                docIds.add(xdoc.getID());
            }

            List<String> newLinks = new ArrayList<>();
            for (ILink link : links.getLinks()) {
                IDocument edoc = (IDocument) link.getTargetInformationObject();
                if (!edoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}

                String docNo = edoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                String revNo = edoc.getDescriptorValue(Conf.Descriptors.DocRevision, String.class);

                IInformationObject[] lnks = Utils.getChildEngineeringDocuments(docNo, revNo, helper);
                for(IInformationObject llnk : lnks) {
                    IDocument ldoc = (IDocument) llnk;
                    if (!ldoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
                    if (docIds.contains(ldoc.getID())){continue;}
                    newLinks.add(ldoc.getID());
                }
            }
            for(String nlnk : newLinks){
                links.addInformationObject(nlnk);
            }


            String tmnr = proi.getDescriptorValue(Conf.Descriptors.ObjectNumberExternal, String.class);
            if(tmnr == null || tmnr == "") {
                tmnr = (new CounterHelper(ses, proi.getClassID())).getCounterStr();
                proi.setDescriptorValue(Conf.Descriptors.ObjectNumberExternal,
                        tmnr);
            }

            proi.commit();
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