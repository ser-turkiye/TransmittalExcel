package ser;

import com.ser.blueline.IDocument;
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
    ISession session;
    IDocumentServer server;
    IBpmService bpm;
    private ProcessHelper helper;
    @Override
    protected Object execute() {
        if (getEventDocument() == null)
            return resultError("Null Document object");

        session = getSes();
        server = session.getDocumentServer();
        IDocument document = getEventDocument();

        com.spire.license.LicenseProvider.setLicenseKey(Conf.Licences.SPIRE_XLS);

        try {

            //Utils.removeTransmittalRepresentations(document, "xlsx");

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