package ser;

import com.ser.blueline.IDatabase;
import com.ser.blueline.IDescriptor;
import com.ser.blueline.IDocumentServer;
import com.ser.blueline.ISession;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.UUID;
import java.util.function.Function;


public class XTRObjects {
    public static String uniqueId = UUID.randomUUID().toString();
    public static String classNameIInformationObject = "com.ser.internal.foldermanagerimpl.sedna.Folder";
    public static String classNameIDocument = "com.ser.sedna.client.bluelineimpl.document.Document";
    public static String classNameITask = "com.ser.sedna.client.bluelineimpl.bpm.Task";
    public static String classNameIProcessInstance = "com.ser.sedna.client.bluelineimpl.bpm.ProcessInstance";
    public static String patternDefaultDate = "dd.MM.yyyy";
    public static String patternDefaultTime = "HH:mm:ss";
    public static String patternDefaultDateTime = "dd.MM.yyyy HH:mm:ss";
    public static List<String> patternsDateTime = new ArrayList<>(Arrays.asList(
            "yyyyMMdd", "yyyy/MM/dd", "dd/MM/yyyy", "MM/dd/yyyy",
            "yyyy.MM.dd", "dd.MM.yyyy", "MM.dd.yyyy",
            "yyyy-MM-dd", "dd-MM-yyyy", "MM-dd-yyyy",
            "yyyy.MM.dd HH:mm:ss", "dd.MM.yyyy HH:mm:ss", "MM.dd.yyyy HH:mm:ss",
            "yyyy-MM-dd HH:mm:ss", "dd-MM-yyyy HH:mm:ss", "MM-dd-yyyy HH:mm:ss",
            "yy","yyyy","M","MM","d","dd","H","HH","m","mm","s","ss"
    ));
    static ISession session;
    static IDocumentServer server;
    static String principalName;
    static String exportPath;

    public static void setSession(ISession ses) throws Exception {
        XTRObjects.session = ses;
        XTRObjects.server = XTRObjects.session.getDocumentServer();
        XTRObjects.principalName = XTRObjects.server.getServicePrincipalName();
    }
    public static void setExportPath(String exportPath) throws Exception {
        XTRObjects.exportPath = (exportPath + "/" + XTRObjects.uniqueId).replace("//", "/");
            (new File(XTRObjects.exportPath)).mkdirs();
    }
    public static IDatabase getDatabase(String dbn) throws Exception {
        return XTRObjects.getDatabase((db) -> {
            if(db.getName().equals(dbn)){
                return db;
            }
            return null;
        });
    }
    public static IDatabase getDatabase(Function <IDatabase, IDatabase> func) throws Exception {
        IDatabase[] dbs = session.getDatabases();
        int c = 0;
        for(IDatabase db : dbs){
            if(func.apply(db) == null){continue;}
            return db;
        }
        return null;
    }
    public static IDescriptor getDescriptor(String dn) throws Exception {
        return server.getDescriptorForName(session, dn);
    }
    public static void callDatabases(Function <IDatabase, Boolean> func) throws Exception {
        IDatabase[] dbs = session.getDatabases();
        int c = 0;
        for(IDatabase db : dbs){
            if(!func.apply(db)){break;}
        }
    }
    public static ISession getSession(){
        return XTRObjects.session;
    }
    public static IDocumentServer getServer(){
        return XTRObjects.server;
    }
    public static String getPrincipalName(){
        return XTRObjects.principalName;
    }
    public static String getExportPath(){
        return XTRObjects.exportPath;
    }
    public static NumberRange numberRange() throws Exception{
        return NumberRange.init();
    }
    public static NumberRange numberRange(Object obj) throws Exception{
        return XTRObjects.numberRange().with(obj);
    }
    public static InfoObject infoObject(Object obj) throws Exception{
        return InfoObject.init(obj);
    }
    public static AutoText autoText() throws Exception{
        return AutoText.init();
    }
    public static AutoText autoText(Object obj) throws Exception{
        return XTRObjects.autoText().with(obj);
    }
    public static String autoText(Object obj, String atxt) throws Exception{
        return XTRObjects.autoText().with(obj).run(atxt);
    }

}
