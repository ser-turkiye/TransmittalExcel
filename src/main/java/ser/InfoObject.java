package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import org.apache.commons.io.FilenameUtils;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class InfoObject {
    Object infObj;
    public InfoObject(Object object) {
        infObj = object;
    }
    public static InfoObject init(Object object) {
        return new InfoObject(object);
    }
    public boolean hasDescriptor(String descName){
        if(infObj == null){return false;}
        if(isInformationObject() || isProcessInstance() || isTask() || isDocument()){
            IInformationObject info = (IInformationObject) infObj;
            IValueDescriptor[] vds = info.getDescriptorList();
            for(IValueDescriptor vd : vds){
                String pnam = vd.getName();
                if(pnam.equals(descName)){return true;}
            }
        }
        return false;
    }
    public boolean isLocked() {
        if(isDocument()) {
            if (((IDocument) infObj).getCheckOutInfo().getOwnerID() != null) {
               return true;
            }
        }
        if(isTask()) {
            if (((ITask) infObj).getProcessInstance().findLockInfo().getOwnerID() != null) {
               return true;
            }
        }
        if(isProcessInstance()) {
            if (((IProcessInstance) infObj).findLockInfo().getOwnerID() != null) {
               return true;
            }
        }
        return false;
    }
    public boolean isDocument(){
        List<String> dxcs =  new ArrayList<>(Arrays.asList(
                XTRObjects.classNameIDocument
        ));
        return dxcs.contains(infObj.getClass().getName());
    }
    public boolean isProcessInstance(){
        List<String> dxcs =  new ArrayList<>(Arrays.asList(
                XTRObjects.classNameIProcessInstance
        ));
        return dxcs.contains(infObj.getClass().getName());
    }
    public boolean isTask() {
        List<String> dxcs =  new ArrayList<>(Arrays.asList(
                XTRObjects.classNameITask
        ));
        return dxcs.contains(infObj.getClass().getName());
    }
    public boolean isInformationObject(){
        List<String> dxcs =  new ArrayList<>(Arrays.asList(
                XTRObjects.classNameIInformationObject
        ));
        return dxcs.contains(infObj.getClass().getName());
    }
    public String exportDocument(String fileName) throws Exception {
        if(!isDocument()){throw new Exception("Export-Document only for IDocument object.");}
        IDocument doc = (IDocument) infObj;
        return exportRepresentation(doc.getDefaultRepresentation(), XTRObjects.getExportPath(), fileName);
    }
    public String exportDocument(String exportPath, String fileName) throws Exception {
        if(!isDocument()){throw new Exception("Export-Document only for IDocument object.");}
        IDocument doc = (IDocument) infObj;
        return exportRepresentation(doc.getDefaultRepresentation(), exportPath, fileName);
    }
    public String exportRepresentation(int rinx, String fileName) throws Exception {
        if(!isDocument()){throw new Exception("Export-Representation only for IDocument object.");}
        return exportRepresentation(rinx, XTRObjects.getExportPath(), fileName);
    }
    public String exportRepresentation(int rinx, String exportPath, String fileName) throws Exception {
        if(!isDocument()){throw new Exception("Export-Representation only for IDocument object.");}
        String rtrn = "";
        IDocument doc = (IDocument) infObj;
        IDocumentPart partDocument = doc.getPartDocument(rinx, 0);
        String fNam = (!fileName.isEmpty() ? autoText(fileName) : partDocument.getFilename());
        fNam = fNam.replaceAll("[\\\\/:*?\"<>|]", "_");
        try (InputStream inputStream = partDocument.getRawDataAsStream()) {
            IFDE fde = partDocument.getFDE();
            if (fde.getFDEType() == IFDE.FILE) {
                String full = ((IFileFDE) fde).getLongFileName();
                String eXtn = FilenameUtils.getExtension(full);
                String eNam = FilenameUtils.getBaseName(full);
                String zNam = fNam
                        .replaceAll("%FULL%", full)
                        .replaceAll("%EXTN%", eXtn)
                        .replaceAll("%FNAM%", eNam);
                rtrn = exportPath + "/" + zNam + (!fNam.contains("%EXTN%") ? "." + eXtn : "");
                rtrn = rtrn.replaceAll("//", "/");

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
    public String autoText(String atxt) throws Exception {
        return XTRObjects.autoText(infObj).run(atxt);
    }
}
