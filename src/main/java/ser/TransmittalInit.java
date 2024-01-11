package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import static java.lang.System.out;


public class TransmittalInit extends UnifiedAgent {
    Logger log = LogManager.getLogger();
    IProcessInstance processInstance;
    IInformationObject projectInfObj;
    IInformationObjectLinks transmittalLinks;
    ProcessHelper helper;
    ITask task;
    List<String> documentIds;
    String transmittalNr;
    String projectNo;
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");

        if(getEventTask().getProcessInstance().findLockInfo().getOwnerID() != null){
            return resultRestart("Restarting Agent");
        }

        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();
        Utils.loadDirectory(Conf.Paths.MainPath);
        
        task = getEventTask();

        try {

            helper = new ProcessHelper(Utils.session);
            XTRObjects.setSession(Utils.session);

            processInstance = task.getProcessInstance();
            projectNo = (processInstance != null ? Utils.projectNr((IInformationObject) processInstance) : "");
            if(projectNo.isEmpty()){
                throw new Exception("Project no is empty.");
            }

            projectInfObj = Utils.getProjectWorkspace(projectNo, helper);
            if(projectInfObj == null){
                throw new Exception("Project not found [" + projectNo + "].");
            }
            transmittalNr = Utils.getTransmittalNr(projectInfObj, processInstance);
            if(transmittalNr.isEmpty()){
                throw new Exception("Transmittal number not found.");
            }

            transmittalLinks = processInstance.getLoadedInformationObjectLinks();
            documentIds = new ArrayList<>();

            for (ILink link : transmittalLinks.getLinks()) {
                IDocument xdoc = (IDocument) link.getTargetInformationObject();
                if (!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
                if (documentIds.contains(xdoc.getID())){continue;}
                documentIds.add(xdoc.getID());
            }

            List<String> newLinks = new ArrayList<>();
            for (ILink link : transmittalLinks.getLinks()) {
                IDocument edoc = (IDocument) link.getTargetInformationObject();
                if (!edoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}

                String docNo = edoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                String revNo = edoc.getDescriptorValue(Conf.Descriptors.DocRevision, String.class);

                IInformationObject[] lnks = Utils.getChildEngineeringDocuments(docNo, revNo, helper);
                for(IInformationObject llnk : lnks) {
                    IDocument ldoc = (IDocument) llnk;
                    if (!ldoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
                    if (documentIds.contains(ldoc.getID())){continue;}
                    newLinks.add(ldoc.getID());
                }
            }
            for(String nlnk : newLinks){
                transmittalLinks.addInformationObject(nlnk);
            }


            //processInstance = Utils.updateProcessInstance(processInstance);
            processInstance.commit();
            log.info("Tested.");

        } catch (Exception e) {
            //throw new RuntimeException(e);
            log.error("Exception       : " + e.getMessage());
            log.error("    Class       : " + e.getClass());
            log.error("    Stack-Trace : " + e.getStackTrace() );
            return resultError("Exception : " + e.getMessage());
        }

        log.info("Finished");
        return resultSuccess("Ended successfully");
    }
}